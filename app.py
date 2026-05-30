"""
Flask API — مذكرتي Pro v28 — Production Ready
- Input sanitization & size limits
- Rate limiting per IP
- Security headers
- Graceful error handling
- No PPTX sent before payment
"""
import base64
import hashlib
import io
import logging
import os
import re
import sys
import threading
import time
import unicodedata
import uuid
from collections import defaultdict
from functools import wraps
from pathlib import Path

from flask import (Flask, abort, jsonify, make_response, redirect,
                   request, send_file, send_from_directory)

_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _BASE)

from core.models import PresentationRequest
from core.payment_models import (
    admin_check, admin_login, approve_order, attach_receipt,
    create_order, get_audit, get_order, get_stats, init_db,
    list_orders, redeem_code, reject_order, store_pptx,
    new_download_code, get_db,
)
from core.preview import (
    generate_preview_async,
    generate_preview_sync,
    get_preview_session,
    get_preview_slides,
    get_cached_preview,
    set_cached_preview,
)
from engine.pipeline import get_pipeline

# ── App Setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="public", static_url_path="")
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-in-prod-v28")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stdout,
)
log = logging.getLogger(__name__)

STORAGE_DIR = Path(os.environ.get("STORAGE_DIR", os.path.join(_BASE, "storage")))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
(STORAGE_DIR / "receipts").mkdir(exist_ok=True)
(STORAGE_DIR / "pptx").mkdir(exist_ok=True)

ALLOWED_RECEIPT_MIME = {
    "image/jpeg", "image/jpg", "image/png", "image/webp",
    "application/pdf",
}
MAX_RECEIPT_SIZE   = 10 * 1024 * 1024   # 10 MB
MAX_REQUEST_BYTES  = 512 * 1024          # 512 KB — حماية من DoS

init_db()

# ── Rate Limiting ─────────────────────────────────────────────────────────────
_rate_store: dict = defaultdict(list)
_rate_lock = threading.Lock()

def _rate_limit(ip: str, endpoint: str, max_calls: int, window: int) -> bool:
    """True = allowed, False = rate-limited"""
    key = f"{ip}:{endpoint}"
    now = time.time()
    with _rate_lock:
        calls = [t for t in _rate_store[key] if now - t < window]
        if len(calls) >= max_calls:
            return False
        calls.append(now)
        _rate_store[key] = calls
    return True

# ── Helpers ───────────────────────────────────────────────────────────────────
def _safe_filename(name: str) -> str:
    if not name:
        return f"prs_{int(time.time())}"
    normalized = unicodedata.normalize("NFKD", name)
    ascii_str = normalized.encode("ascii", "ignore").decode("ascii")
    safe = "".join(c if c.isalnum() else "_" for c in ascii_str).strip("_")
    if not safe:
        safe = f"student_{int(time.time()) % 100000}"
    return safe[:24]

def _get_ip() -> str:
    return (
        request.headers.get("X-Forwarded-For", "").split(",")[0].strip()
        or request.remote_addr or "unknown"
    )

def _admin_token() -> str:
    return (
        request.headers.get("X-Admin-Token")
        or request.cookies.get("admin_token")
        or request.args.get("token") or ""
    )

def require_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not admin_check(_admin_token()):
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return wrapper

# ── CORS + Security Headers ───────────────────────────────────────────────────
@app.after_request
def _security_headers(r):
    r.headers["Access-Control-Allow-Origin"]  = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
    r.headers["X-Content-Type-Options"]       = "nosniff"
    r.headers["X-Frame-Options"]              = "SAMEORIGIN"
    r.headers["Referrer-Policy"]              = "strict-origin-when-cross-origin"
    if "/preview/" in request.path or "/slide/" in request.path:
        r.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    return r

@app.before_request
def _preflight():
    if request.method == "OPTIONS":
        r = make_response("", 204)
        r.headers["Access-Control-Allow-Origin"]  = "*"
        r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE"
        r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
        return r

# ── Static ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("public", "index.html")

@app.route("/admin")
@app.route("/admin/")
def admin_page():
    return send_from_directory("public", "admin.html")

# ── Health ────────────────────────────────────────────────────────────────────
@app.route("/ping")
def ping():
    return "pong", 200

@app.route("/health")
def health():
    pipeline = get_pipeline()
    return jsonify({"status": "ok", "version": "28.0", "font": pipeline._font})

@app.route("/warmup")
def warmup():
    get_pipeline()
    return jsonify({"status": "ready"})

# ── Generate ──────────────────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    ip = _get_ip()

    # Rate limit: 10 توليدات/دقيقة لكل IP
    if not _rate_limit(ip, "generate", 10, 60):
        return jsonify({"error": "تجاوزت الحد المسموح. حاول بعد دقيقة."}), 429

    # Payload size check
    if request.content_length and request.content_length > MAX_REQUEST_BYTES:
        return jsonify({"error": "الطلب كبير جداً"}), 413

    t0 = time.monotonic()
    raw = request.get_json(force=True, silent=True)
    if not raw or not isinstance(raw, dict):
        return jsonify({"error": "بيانات غير صالحة"}), 400

    req = PresentationRequest.from_dict(raw)
    errors = req.validate()
    if errors:
        return jsonify({"error": " | ".join(errors)}), 400

    pipeline = get_pipeline()
    result = pipeline.build(req)

    if not result.success:
        log.error(f"Build failed [{ip}]: {result.error}")
        return jsonify({"error": result.error, "stages": result.stages}), 500

    presentation_id = str(uuid.uuid4())
    pptx_path = STORAGE_DIR / "pptx" / f"{presentation_id}.pptx"
    pptx_path.write_bytes(result.data)

    preview_token, preview_slides = generate_preview_sync(presentation_id, str(pptx_path))

    elapsed = time.monotonic() - t0
    log.info(f"Generated: id={presentation_id} slides={result.slide_count} "
             f"preview={len(preview_slides)} elapsed={elapsed:.2f}s ip={ip}")

    return jsonify({
        "ok": True,
        "presentation_id": presentation_id,
        "preview_token": preview_token,
        "slides": result.slide_count,
        "font": result.font_used,
        "elapsed": round(elapsed, 2),
        "stages": result.stages,
        "filename": f"mathkarati_{_safe_filename(req.student_name)}.pptx",
        "student_name": req.student_name,
        "title_ar": req.title_ar,
        "degree": str(raw.get("degree", raw.get("level", "licence")))[:20],
        "preview_slides": preview_slides,
        "preview_count": len(preview_slides),
    })

# ── Preview ───────────────────────────────────────────────────────────────────
@app.route("/preview/<presentation_id>", methods=["GET"])
def get_preview(presentation_id):
    if not re.match(r'^[0-9a-f\-]{36}$', presentation_id):
        return jsonify({"error": "معرف غير صالح"}), 400

    session = get_preview_session(presentation_id)
    if not session:
        pptx_path = STORAGE_DIR / "pptx" / f"{presentation_id}.pptx"
        if not pptx_path.exists():
            return jsonify({"error": "العرض غير موجود أو انتهت صلاحيته"}), 404
        token = generate_preview_async(presentation_id, str(pptx_path))
        return jsonify({"ok": True, "status": "pending", "preview_token": token,
                        "slide_count": 0, "processing": True})

    status = session.get("status", "pending")
    resp = {"ok": True, "status": status,
            "slide_count": session.get("slide_count", 0),
            "processing": status == "pending"}
    if status == "ready":
        resp["slides"] = session.get("slides", [])
        resp["preview_token"] = session.get("token", "")
    return jsonify(resp)

@app.route("/preview/<presentation_id>/slides", methods=["GET"])
def get_preview_slides_endpoint(presentation_id):
    if not re.match(r'^[0-9a-f\-]{36}$', presentation_id):
        return jsonify({"error": "معرف غير صالح"}), 400
    token = request.args.get("token", "").strip()
    if not token:
        return jsonify({"error": "توكن مطلوب"}), 401
    slides = get_preview_slides(presentation_id, token)
    if slides is None:
        return jsonify({"error": "توكن غير صالح أو انتهت الصلاحية"}), 403
    return jsonify({"ok": True, "slides": slides, "count": len(slides)})

# ── Preview Slide (PPTX — 1 slide) ───────────────────────────────────────────
@app.route("/preview-slide/<presentation_id>", methods=["GET"])
def download_preview_slide(presentation_id):
    if not re.match(r'^[0-9a-f\-]{36}$', presentation_id):
        return jsonify({"error": "معرف غير صالح"}), 400

    pptx_path = STORAGE_DIR / "pptx" / f"{presentation_id}.pptx"
    if not pptx_path.exists():
        return jsonify({"error": "العرض غير موجود أو انتهت صلاحيته"}), 404

    try:
        import zipfile as _zf
        import io as _io
        from lxml import etree as _et

        src_data = pptx_path.read_bytes()

        with _zf.ZipFile(_io.BytesIO(src_data)) as zin:
            all_names = zin.namelist()
            all_files = {n: zin.read(n) for n in all_names}

        slide_names = sorted(
            [n for n in all_names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group())
        )
        slides_to_remove = set(slide_names[1:])

        def _rels_path(sp):
            parts = sp.rsplit('/', 1)
            return parts[0] + '/_rels/' + parts[1] + '.rels'

        rels_to_remove = {_rels_path(s) for s in slides_to_remove}
        to_remove = slides_to_remove | rels_to_remove

        out_buf = _io.BytesIO()
        with _zf.ZipFile(out_buf, 'w', _zf.ZIP_DEFLATED) as zout:
            for name, data in all_files.items():
                if name in to_remove:
                    continue
                if name == 'ppt/presentation.xml':
                    root = _et.fromstring(data)
                    ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
                    sldIdLst = root.find(f'{{{ns}}}sldIdLst')
                    if sldIdLst is not None:
                        for child in list(sldIdLst)[1:]:
                            sldIdLst.remove(child)
                    data = _et.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif name == '[Content_Types].xml':
                    root = _et.fromstring(data)
                    for child in list(root):
                        if child.get('PartName', '').lstrip('/') in to_remove:
                            root.remove(child)
                    data = _et.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif name == 'ppt/_rels/presentation.xml.rels':
                    root = _et.fromstring(data)
                    slide_ct = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
                    for child in list(root):
                        if child.get('Type') == slide_ct:
                            full = 'ppt/' + child.get('Target', '')
                            if full in slides_to_remove:
                                root.remove(child)
                    data = _et.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(name, data)

        preview_data = out_buf.getvalue()
        b64 = base64.b64encode(preview_data).decode('ascii')
        log.info(f"Preview PPTX: {presentation_id} ({len(preview_data)//1024}KB)")
        return jsonify({"ok": True, "data": b64,
                        "filename": "preview-mathkarati.pptx",
                        "size": len(preview_data)})
    except Exception as e:
        log.error(f"Preview slide failed: {e}", exc_info=True)
        return jsonify({"error": "فشل تصدير المعاينة"}), 500

# ── Orders ────────────────────────────────────────────────────────────────────
@app.route("/orders", methods=["POST"])
def create_order_route():
    ip = _get_ip()
    if not _rate_limit(ip, "orders", 5, 60):
        return jsonify({"error": "حاول بعد دقيقة"}), 429

    d = request.get_json(force=True, silent=True) or {}
    pid            = str(d.get("presentation_id", "")).strip()[:36]
    student_name   = str(d.get("student_name",   "")).strip()[:120]
    email          = str(d.get("email",          "")).strip()[:200]
    phone          = str(d.get("phone",          "")).strip()[:30]
    degree         = str(d.get("degree",         "licence")).strip()[:20]
    title_ar       = str(d.get("title_ar",       "")).strip()[:400]
    payment_method = str(d.get("payment_method", "")).strip()[:20]

    if not re.match(r'^[0-9a-f\-]{36}$', pid) if pid else True:
        return jsonify({"error": "معرف العرض غير صالح"}), 400
    if not pid or not student_name or not payment_method:
        return jsonify({"error": "بيانات ناقصة"}), 400
    if payment_method not in ("ccp", "baridimob"):
        return jsonify({"error": "طريقة دفع غير صالحة"}), 400

    order = create_order(pid, student_name, email, phone, degree, title_ar, payment_method)
    log.info(f"Order: {order['id']} student={student_name} ip={ip}")
    return jsonify({"ok": True, "order_id": order["id"]}), 201

@app.route("/orders/<order_id>", methods=["GET"])
def get_order_route(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    safe = {k: v for k, v in order.items()
            if k not in ("pptx_path", "receipt_path", "download_code", "download_ip")}
    return jsonify(safe)

@app.route("/orders/<order_id>/receipt", methods=["POST"])
def upload_receipt(order_id):
    ip = _get_ip()
    if not _rate_limit(ip, "receipt", 3, 60):
        return jsonify({"error": "حاول بعد دقيقة"}), 429

    order = get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    if order["status"] not in ("pending",):
        return jsonify({"error": "لا يمكن رفع وصل لهذا الطلب"}), 400
    if "receipt" not in request.files:
        return jsonify({"error": "لم يتم إرسال الملف"}), 400

    file = request.files["receipt"]
    mime = file.content_type or "application/octet-stream"
    if mime not in ALLOWED_RECEIPT_MIME:
        return jsonify({"error": "نوع الملف غير مسموح (JPG، PNG، PDF فقط)"}), 400

    data = file.read(MAX_RECEIPT_SIZE + 1)
    if len(data) > MAX_RECEIPT_SIZE:
        return jsonify({"error": "حجم الملف يتجاوز 10MB"}), 400
    if len(data) < 100:
        return jsonify({"error": "الملف يبدو فارغاً أو تالفاً"}), 400

    ext = "pdf" if "pdf" in mime else "jpg"
    fname = f"{order_id}_{int(time.time())}.{ext}"
    path = STORAGE_DIR / "receipts" / fname
    path.write_bytes(data)

    attach_receipt(order_id, str(path), mime)
    log.info(f"Receipt: {order_id} {fname} ip={ip}")
    return jsonify({"ok": True, "message": "تم رفع الوصل بنجاح، سيتم مراجعته خلال 24 ساعة"})

# ── Redeem ────────────────────────────────────────────────────────────────────
@app.route("/redeem", methods=["POST"])
def redeem_by_code():
    ip = _get_ip()
    if not _rate_limit(ip, "redeem", 5, 60):
        return jsonify({"error": "محاولات كثيرة، حاول بعد دقيقة"}), 429

    d = request.get_json(force=True, silent=True) or {}
    code = str(d.get("code", "")).strip().upper()[:20]
    if not code:
        return jsonify({"error": "الرجاء إدخال الكود"}), 400

    order = redeem_code(code, ip)
    if not order:
        return jsonify({"error": "الكود غير صحيح أو منتهي الصلاحية أو تم استخدامه مسبقاً"}), 403

    pptx_path = order.get("pptx_path")
    if not pptx_path or not Path(pptx_path).exists():
        return jsonify({"error": "ملف العرض غير متوفر، تواصل مع الدعم"}), 500

    pptx_bytes = Path(pptx_path).read_bytes()
    b64 = base64.b64encode(pptx_bytes).decode("ascii")
    safe = _safe_filename(order["student_name"])
    filename = f"mathkarati_{safe}.pptx"

    log.info(f"Download: order={order['id']} ip={ip}")
    return jsonify({"ok": True, "data": b64, "filename": filename,
                    "order_id": order["id"], "student_name": order["student_name"]})

# ── Admin ─────────────────────────────────────────────────────────────────────
@app.route("/admin/login", methods=["POST"])
def admin_login_route():
    ip = _get_ip()
    if not _rate_limit(ip, "admin_login", 5, 300):  # 5 محاولات / 5 دقائق
        return jsonify({"error": "محاولات كثيرة"}), 429
    d = request.get_json(force=True, silent=True) or {}
    token = admin_login(str(d.get("password", "")), ip)
    if not token:
        return jsonify({"error": "كلمة المرور غير صحيحة"}), 401
    resp = jsonify({"ok": True, "token": token})
    resp.set_cookie("admin_token", token, httponly=True, samesite="Strict", max_age=8 * 3600)
    return resp

@app.route("/admin/stats", methods=["GET"])
@require_admin
def admin_stats():
    return jsonify(get_stats())

@app.route("/admin/orders", methods=["GET"])
@require_admin
def admin_orders():
    status = request.args.get("status")
    return jsonify(list_orders(status=status or None))

@app.route("/admin/orders/<order_id>", methods=["GET"])
@require_admin
def admin_order_detail(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "غير موجود"}), 404
    return jsonify({"order": order, "audit": get_audit(order_id)})

@app.route("/admin/orders/<order_id>/receipt", methods=["GET"])
@require_admin
def admin_view_receipt(order_id):
    order = get_order(order_id)
    if not order or not order.get("receipt_path"):
        abort(404)
    path = Path(order["receipt_path"])
    if not path.exists():
        abort(404)
    return send_file(path, mimetype=order.get("receipt_mime", "application/octet-stream"),
                     as_attachment=False)

@app.route("/admin/orders/<order_id>/approve", methods=["POST"])
@require_admin
def admin_approve(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "غير موجود"}), 404
    if order["status"] not in ("pending", "rejected"):
        return jsonify({"error": f"الطلب في حالة {order['status']}"}), 400

    d = request.get_json(force=True, silent=True) or {}
    note = str(d.get("note", ""))[:500]
    ttl  = min(int(d.get("ttl_hours", 48)), 168)  # max 7 أيام

    pid = order["presentation_id"]
    pptx_src = STORAGE_DIR / "pptx" / f"{pid}.pptx"
    if not pptx_src.exists():
        return jsonify({"error": "ملف العرض غير موجود، ربما انتهت صلاحيته"}), 500

    store_pptx(order_id, str(pptx_src))
    code = approve_order(order_id, note, ttl)
    log.info(f"Approved: {order_id} code={code}")
    return jsonify({"ok": True, "download_code": code, "ttl_hours": ttl})

@app.route("/admin/orders/<order_id>/reject", methods=["POST"])
@require_admin
def admin_reject(order_id):
    d = request.get_json(force=True, silent=True) or {}
    reject_order(order_id, str(d.get("note", ""))[:500])
    log.info(f"Rejected: {order_id}")
    return jsonify({"ok": True})

@app.route("/admin/orders/<order_id>/resend-code", methods=["POST"])
@require_admin
def admin_resend_code(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "غير موجود"}), 404
    if order["status"] not in ("approved", "downloaded"):
        return jsonify({"error": "يجب أن يكون الطلب معتمداً"}), 400

    d = request.get_json(force=True, silent=True) or {}
    ttl  = min(int(d.get("ttl_hours", 24)), 168)
    code = new_download_code()
    expires = time.time() + ttl * 3600
    conn = get_db()
    try:
        conn.execute(
            """UPDATE orders SET download_code=?, code_expires=?, code_used=0,
               status='approved', updated_at=? WHERE id=?""",
            (code, expires, time.time(), order_id),
        )
        conn.commit()
    finally:
        conn.close()
    log.info(f"Code regenerated: {order_id}")
    return jsonify({"ok": True, "download_code": code, "ttl_hours": ttl})

# ── Error handlers ─────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "غير موجود"}), 404

@app.errorhandler(405)
def method_not_allowed(e):
    return jsonify({"error": "الطريقة غير مسموحة"}), 405

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "الطلب كبير جداً"}), 413

@app.errorhandler(500)
def server_error(e):
    log.error(f"500: {e}")
    return jsonify({"error": "خطأ داخلي في الخادم"}), 500

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    get_pipeline()
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
