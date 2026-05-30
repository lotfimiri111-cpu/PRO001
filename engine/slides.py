"""
Canva Engine v28 — Premium Visual Intelligence
نظام تصميم ذكي يتفاعل مع المحتوى بصرياً
"""
from __future__ import annotations
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, hexagon, decorative_dots,
    card_3d, icon_circle, number_badge, slide_number,
    txt, txt2, blank_slide,
    # v28 premium
    accent_bar_v, kpi_card, result_row_premium, section_header_band,
    highlight_chip, content_card_premium, two_col_layout,
    methodology_card, premium_cover_frame, smart_header, decorative_corner,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(n): global _FONT; _FONT = n

HEADER_H = 2.9

# ── Typography Scale v28 ────────────────────────────────────────────────
SZ_SLIDE_TITLE   = 30
SZ_SLIDE_SUB     = 13
SZ_SECTION_LABEL = 19
SZ_BODY          = 13
SZ_BODY_SM       = 11.5
SZ_LABEL         = 13
SZ_VALUE         = 15
SZ_FINAL_TITLE   = 42
SZ_FINAL_SUB     = 24

# ── Smart content helpers ───────────────────────────────────────────────
def _font_for_len(text, base=14, min_s=10, max_s=16):
    n = len(str(text))
    if n < 40:   return min(max_s, base)
    if n < 80:   return base - 1
    if n < 130:  return base - 2
    return max(min_s, base - 3)

def _items_layout(n, CH, gap=0.18, min_h=1.0):
    """Smart row height for n items."""
    row_h = (CH - gap * max(0, n - 1)) / max(1, n)
    return max(min_h, row_h), gap


# ── Background styles ────────────────────────────────────────────────────
def _bg(slide, T, style='a'):
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2,
                  angle={'a': 135, 'b': 160, 'c': 90, 'd': 45}[style])
    if style == 'a':
        decorative_corner(slide, T, 'tr', 4.0)
        decorative_corner(slide, T, 'bl', 3.5)
        decorative_dots(slide, 1.2, H - 4.2, 5, 3, 0.16, 0.42, T.accent_rgb, alpha=12)
    elif style == 'b':
        diamond(slide, W - 7, -2, 6, 6, T.accent_rgb, alpha=6)
        diamond(slide, -1.5, H - 4.5, 4.5, 4.5, T.accent_rgb, alpha=5)
        hexagon(slide, W - 4.5, H * 0.3, 2.8, 2.8, T.accent_rgb, alpha=7)
        decorative_dots(slide, 1.0, 1.8, 4, 4, 0.15, 0.36, T.accent_rgb, alpha=9)
    elif style == 'c':
        decorative_corner(slide, T, 'tr', 4.5)
        decorative_corner(slide, T, 'bl', 3.8)
        decorative_dots(slide, W - 6.5, 1.5, 4, 5, 0.14, 0.35, T.accent_rgb, alpha=10)
    elif style == 'd':
        for r, a in [(26, 4), (20, 5), (14, 6), (8, 8)]:
            oval(slide, W/2 - r/2, H/2 - r/2, r, r, T.accent_rgb, alpha=a)
        decorative_dots(slide, 1.8, H - 3.8, 5, 2, 0.18, 0.44, T.accent_rgb, alpha=12)


# ── Smart Header ─────────────────────────────────────────────────────────
def _hdr(slide, T, title, sub='', side='right', slide_num=None, total_slides=13, req=None):
    if req is not None and hasattr(req, '_total_slides'):
        total_slides = req._total_slides
    smart_header(slide, T, title, sub, slide_num, total_slides,
                 style='gradient', font=_FONT)


CY0 = HEADER_H + 0.32
def _ch(): return H - CY0 - 0.3


# ══════════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'b')

    # Top + bottom accent bars
    tp = rect(slide, 0, 0, W, 0.44, T.accent_rgb)
    if tp: multi_stop_gradient(tp, [(0, T.bg), (30, T.accent), (70, T.accent2), (100, T.bg)], 0)
    bt = rect(slide, 0, H - 0.32, W, 0.32, T.accent_rgb)
    if bt: gradient_fill(bt, T.accent_grad1, T.accent_grad2, 0)

    # Institution badge
    if req.institution:
        ib = rrect(slide, W/2 - 10, 0.52, 20, 0.7, T.card_rgb, radius_pct=40)
        if ib: set_solid_alpha(ib, 72)
        txt(slide, req.institution, W/2 - 10, 0.52, 20, 0.7,
            font=_FONT, size=11, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    title_y = 1.28; title_h = 7.2
    info_y = title_y + title_h + 0.24
    info_h = max(0.5, H - 0.36 - info_y - 0.08)
    cx = 1.8; cw = W - 3.6

    # Premium cover frame (multi-layer)
    premium_cover_frame(slide, cx, title_y, cw, title_h, T)

    # Extract year
    import re as _re
    _year_pat = _re.compile(r'\b\d{4}\s*[-–—]\s*\d{4}\b')
    _year_match = _year_pat.search(req.title_ar or '')
    if _year_match:
        _year_str = _year_match.group(0).strip()
        _title_clean = _year_pat.sub('', req.title_ar).strip().rstrip('–—-').strip()
    else:
        _year_str = getattr(req, 'year', None) or ''
        _title_clean = req.title_ar or ''

    # Content-aware title sizing
    title_len = len(_title_clean)
    title_fs = 26 if title_len < 60 else 22 if title_len < 100 else 18 if title_len < 150 else 15

    title_zone_h = title_h * 0.56
    txt(slide, _title_clean, cx + 0.7, title_y + 0.48, cw - 1.4, title_zone_h,
        font=_FONT, size=title_fs, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.25)

    # Decorative divider
    d1 = rect(slide, cx + cw * 0.12, title_y + 0.48 + title_zone_h + 0.1,
              cw * 0.76, 0.06, T.accent_rgb)
    if d1: multi_stop_gradient(d1, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)

    # Info zone
    info_zone_y = title_y + 0.48 + title_zone_h + 0.28
    info_zone_h = title_h - 0.48 - title_zone_h - 0.38
    fields = []
    if req.student_name:    fields.append(("الطالب", req.student_name, "👤"))
    if req.supervisor:      fields.append(("المشرف", req.supervisor, "🎓"))
    if req.specialization:  fields.append(("التخصص", req.specialization, "📜"))
    if _year_str:           fields.append(("السنة", _year_str, "📅"))
    if not fields and req.student_name:
        fields = [("الطالب", req.student_name, "👤")]

    n_fields = min(len(fields), 4)
    if n_fields > 0:
        fh = min(1.4, info_zone_h / n_fields - 0.1)
        fy = info_zone_y + (info_zone_h - (fh * n_fields + 0.08 * (n_fields - 1))) / 2
        for k, (lbl, val, ico) in enumerate(fields[:4]):
            ky = fy + k * (fh + 0.08)
            fb = rrect(slide, cx + 0.5, ky, cw - 1.0, fh, T.bg_rgb, radius_pct=10)
            if fb:
                multi_stop_gradient(fb, [(0, T.bg), (50, T.bg2), (100, T.bg)], 0)
                set_solid_alpha(fb, 65)
            # Icon
            txt(slide, ico, cx + cw - 1.35, ky, 0.85, fh,
                font='Calibri', size=14, bold=False, color=T.accent_rgb,
                align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
            # Label chip
            highlight_chip(slide, cx + cw - 1.35 - 2.4, ky + (fh - 0.34) / 2,
                           2.2, 0.34, lbl, T, style='ghost')
            # Value
            val_fs = _font_for_len(val, 13, 10, 15)
            txt(slide, val, cx + 0.62, ky, cw - 4.85, fh,
                font=_FONT, size=val_fs, bold=True, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    # Specialty / discipline tag bottom
    if req.specialization:
        spec_b = rrect(slide, cx + cw * 0.12, title_y + title_h - 0.74,
                       cw * 0.76, 0.52, T.accent_rgb, radius_pct=40)
        if spec_b:
            gradient_fill(spec_b, T.accent_grad1, T.accent_grad2, 0)
            shadow(spec_b, blur=10, dist=3, alpha=0.35)
        txt(slide, req.specialization, cx + cw * 0.12, title_y + title_h - 0.74,
            cw * 0.76, 0.52,
            font=_FONT, size=11, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'c')
    _hdr(slide, T, "مقدمة البحث", "نظرة عامة على الدراسة", 'right', slide_num=1, req=req)
    CY = CY0; CH = _ch()

    items = []
    if req.intro_overview: items.append(("📖", "نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("🔬", "المنهج المتبع", req.intro_approach))
    if not items: return slide

    n = len(items)
    gap = 0.4
    col_w = (W - 2.4 - gap * (n - 1)) / n
    CARD_H = min(CH * 0.88, 11.5)
    card_y = CY + (CH - CARD_H) / 2

    ic_s = 1.85
    lbl_h = 0.72
    padding_top = 0.48
    ic_y_offset = padding_top
    lbl_y_offset = ic_y_offset + ic_s + 0.30
    div_y_offset = lbl_y_offset + lbl_h + 0.1
    txt_y_offset = div_y_offset + 0.08 + 0.16
    txt_h = CARD_H - txt_y_offset - 0.42

    for i, (icon, lbl, val) in enumerate(items[:2]):
        x = 1.2 + i * (col_w + gap)
        cc = content_card_premium(slide, x, card_y, col_w, CARD_H, T,
                                   accent_side='right', radius=12, depth=True)

        # Top accent strip
        tp = rrect(slide, x, card_y, col_w, 0.42, T.accent_rgb, radius_pct=0)
        if tp:
            multi_stop_gradient(tp, [(0, T.accent), (50, T.accent2), (100, T.accent)], 0)
            glow(tp, T.accent.lstrip('#'), radius=12, alpha=0.3)

        # Icon circle
        ic_x = x + col_w / 2 - ic_s / 2
        ic_y = card_y + ic_y_offset
        icon_circle(slide, ic_x, ic_y, ic_s,
                    T.accent_grad1, T.accent_grad2, icon, max(16, int(ic_s * 11)), T)

        # Section label
        txt(slide, lbl, x + 0.22, card_y + lbl_y_offset, col_w - 0.44, lbl_h,
            font=_FONT, size=SZ_SECTION_LABEL, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        # Divider
        hline(slide, x + col_w * 0.14, card_y + div_y_offset,
              col_w * 0.72, T.accent_rgb, thickness=0.05)

        # Content — content-aware font size
        val_fs = _font_for_len(val, 14, 11, 15)
        txt(slide, val, x + 0.3, card_y + txt_y_offset, col_w - 0.6, txt_h,
            font=_FONT, size=val_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.35)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'a')
    chapters = req.chapters[:8]
    n = len(chapters)
    _hdr(slide, T, "خطة البحث", f"يتضمن البحث {n} فصول رئيسية", 'left', slide_num=2, req=req)
    CY = CY0; CH = _ch()
    if not chapters: return slide

    gap = 0.16
    row_h, gap = _items_layout(n, CH, gap, min_h=0.9)

    for i, ch in enumerate(chapters):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.2: break
        even = i % 2 == 0

        # Row card
        rw = rrect(slide, 1.0, y, W - 2.0, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=10)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=7, dist=2, alpha=0.22)

        # Right accent bar
        accent_bar_v(slide, W - 1.28, y, row_h, T, width=0.24)

        # Chapter number badge
        nd = min(0.74, row_h * 0.74)
        nx = W - 2.58; ny = y + (row_h - nd) / 2
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
            shadow(nc, blur=8, dist=2, alpha=0.32)
        txt(slide, str(i + 1), nx, ny, nd, nd,
            font='Calibri', size=max(9, int(nd * 11)), bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # Chapter title
        title_fs = max(11, min(15, int(row_h * 9.5)))
        if len(ch.title) > 60: title_fs = max(10, title_fs - 1)
        txt(slide, ch.title, 1.28, y, W - 4.78, row_h,
            font=_FONT, size=title_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.15)

        # Pages badge
        if ch.pages:
            pb = rrect(slide, 1.14, y + (row_h - 0.36) / 2, 2.1, 0.36, T.bg_rgb, radius_pct=40)
            if pb: set_solid_alpha(pb, 55)
            txt(slide, ch.pages, 1.14, y + (row_h - 0.36) / 2, 2.1, 0.36,
                font='Calibri', size=9, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'b')
    _hdr(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية", 'right', slide_num=3, req=req)
    CY = CY0; CH = _ch()

    secs = []; weights = {}
    if req.main_problem:  secs.append('p'); weights['p'] = 2.6
    if req.main_question: secs.append('q'); weights['q'] = 1.5
    if req.sub_questions: secs.append('s'); weights['s'] = 2.0
    if not secs: return slide

    tw = sum(weights[s] for s in secs)
    gap = 0.22; avail = CH - gap * (len(secs) - 1)
    cy = CY

    if 'p' in secs:
        h = avail * weights['p'] / tw
        # Main problem card
        pc = rrect(slide, 1.0, cy, W - 2.0, h, T.card_rgb, radius_pct=14)
        if pc:
            multi_stop_gradient(pc, [(0, T.card), (100, T.bg2)], 135)
            shadow(pc, blur=22, dist=6, alpha=0.45)
            glow(pc, T.accent.lstrip('#'), radius=26, alpha=0.08)

        # Label band
        lb = rrect(slide, W - 8.0, cy, 6.2, 0.54, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb, [(0, T.accent), (100, T.accent2)], 0)
        txt(slide, "◆  الإشكالية الرئيسية", W - 8.0, cy, 6.2, 0.54,
            font=_FONT, size=SZ_LABEL, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        # Right accent bar
        accent_bar_v(slide, W - 1.28, cy, h, T, width=0.24)

        # Quote mark
        txt(slide, "❝", 1.3, cy + 0.62, 1.5, 1.2,
            font='Calibri', size=36, bold=False, color=T.accent_rgb,
            align=PP_ALIGN.LEFT, rtl=False, vcenter=False)

        # Problem text — content-aware
        prob_fs = _font_for_len(req.main_problem, 13, 10, 14)
        txt(slide, req.main_problem, 3.0, cy + 0.6, W - 4.55, h - 0.82,
            font=_FONT, size=prob_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.38)
        cy += h + gap

    if 'q' in secs:
        h = avail * weights['q'] / tw
        qc = rrect(slide, 1.0, cy, W - 2.0, h, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=8, dist=2, alpha=0.22)
        vline(slide, W - 1.42, cy, h, T.accent_rgb, thickness=0.24)
        dot = oval(slide, W - 3.3, cy + h / 2 - 0.22, 0.44, 0.44, T.accent_rgb)
        if dot: multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)
        q_fs = _font_for_len(req.main_question, 13, 11, 14)
        txt(slide, req.main_question, 1.35, cy, W - 3.85, h,
            font=_FONT, size=q_fs, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.22)
        cy += h + gap

    if 's' in secs and req.sub_questions:
        h = avail * weights['s'] / tw
        subs = req.sub_questions[:4]
        sub_h = h / max(1, len(subs))
        for i, q in enumerate(subs):
            y2 = cy + i * sub_h
            if i % 2 == 0:
                sb = rrect(slide, 1.0, y2, W - 2.0, sub_h - 0.06, T.card_rgb, radius_pct=6)
                if sb: set_solid_alpha(sb, 48)
            nd2 = min(0.46, sub_h * 0.55)
            nc2 = oval(slide, W - 2.85, y2 + (sub_h - nd2) / 2, nd2, nd2, T.accent_rgb)
            if nc2: set_solid_alpha(nc2, 65)
            txt(slide, str(i + 1), W - 2.85, y2 + (sub_h - nd2) / 2, nd2, nd2,
                font='Calibri', size=max(7, int(nd2 * 9)), bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
            q_fs2 = max(10, min(12.5, sub_h * 7.5))
            txt(slide, q, 1.32, y2, W - 3.85, sub_h,
                font=_FONT, size=q_fs2, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT,
                rtl=True, vcenter=True, line_spacing=1.15)

    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'c')
    _hdr(slide, T, "أهداف البحث وفرضياته", "", 'left', slide_num=4, req=req)
    CY = CY0; CH = _ch()

    cols = []
    if req.objectives: cols.append(("🎯  الأهداف", req.objectives))
    if req.hypotheses:  cols.append(("💡  الفرضيات", req.hypotheses))
    if not cols: return slide

    two_col_layout(slide, T,
                   cols[0][1] if len(cols) > 0 else [],
                   cols[1][1] if len(cols) > 1 else [],
                   CY, CH,
                   left_label=cols[0][0] if len(cols) > 0 else '',
                   right_label=cols[1][0] if len(cols) > 1 else '',
                   font=_FONT)
    return slide


# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'b')
    _hdr(slide, T, "أهمية البحث", "الأثر العلمي والعملي للدراسة", 'right', slide_num=5, req=req)
    CY = CY0; CH = _ch()
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐", "🔑", "📌", "🌟", "💎", "🏆"]
    n = len(items)
    cols = 2 if n > 3 else 1
    col_w = (W - 2.0 - 0.3 * (cols - 1)) / cols
    rows_n = (n + cols - 1) // cols
    gv = 0.22
    card_h = (CH - gv * (rows_n - 1)) / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = 1.0 + ci * (col_w + 0.3)
        y = CY + ri * (card_h + gv)

        # Card
        cc = content_card_premium(slide, x, y, col_w, card_h, T,
                                   accent_side='right', radius=12)

        # Icon
        ic_s = min(1.38, card_h * 0.62)
        icon_circle(slide, x + 0.3, y + (card_h - ic_s) / 2,
                    ic_s, T.accent_grad1, T.accent_grad2,
                    icons[i % len(icons)], max(13, int(ic_s * 11)), T)

        # Text — content-aware
        item_fs = _font_for_len(item, 13, 11, 14)
        text_x = x + ic_s + 0.58
        txt(slide, item, text_x, y + 0.12, col_w - ic_s - 1.1, card_h - 0.24,
            font=_FONT, size=item_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.32)

    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'd')
    _hdr(slide, T, "منهجية البحث", "الإجراءات والأدوات المستخدمة", 'left', slide_num=6, req=req)
    CY = CY0; CH = _ch()

    icons_map = {"المنهج": "📊", "العينة": "👥", "حجم العينة": "📏", "الأداة": "🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج",     req.methodology))
    if req.sample_type:  fields.append(("العينة",     req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة",     req.tool))
    if not fields: return slide

    cols = 2 if len(fields) > 2 else len(fields)
    rows_n = (len(fields) + cols - 1) // cols
    gh = 0.3; gv = 0.24
    col_w = (W - 2.0 - gh * (cols - 1)) / cols
    card_h = (CH - gv * (rows_n - 1)) / rows_n

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i % cols; ri = i // cols
        x = 1.0 + ci * (col_w + gh)
        y = CY + ri * (card_h + gv)
        methodology_card(slide, x, y, col_w, card_h,
                         lbl, val, icons_map.get(lbl, "📌"), T)

    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'a')
    _hdr(slide, T, "الأرقام والإحصاءات الرئيسية", "", 'right', slide_num=7, req=req)
    CY = CY0; CH = _ch()
    stats = req.stats[:6]
    if not stats: return slide

    cols = 3 if len(stats) >= 3 else len(stats)
    rows_n = (len(stats) + cols - 1) // cols
    gh = 0.3; gv = 0.24
    col_w = (W - 2.0 - gh * (cols - 1)) / cols
    card_h = (CH - gv * (rows_n - 1)) / rows_n

    for i, stat in enumerate(stats):
        ci = i % cols; ri = i // cols
        x = 1.0 + ci * (col_w + gh)
        y = CY + ri * (card_h + gv)
        if y + card_h > H - 0.2: break
        kpi_card(slide, x, y, col_w, card_h,
                 stat.value, stat.label, T,
                 unit=stat.unit or '', rank=i)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'c')
    _hdr(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة", 'left', slide_num=8, req=req)
    CY = CY0; CH = _ch()
    results = req.main_results[:8]
    n = len(results)
    if not results: return slide

    gap = 0.16
    row_h, gap = _items_layout(n, CH, gap, min_h=1.1)

    for i, result in enumerate(results):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.2: break
        # First result gets highlight treatment
        result_row_premium(slide, 1.0, y, W - 2.0, row_h,
                           result, i + 1, T, highlight=(i == 0))

    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'd')
    _hdr(slide, T, "خاتمة البحث", "الاستنتاج العام للدراسة", 'right', slide_num=9, req=req)
    CY = CY0; CH = _ch(); cw = W - 2.8

    # Main card
    cc = rrect(slide, 1.4, CY, cw, CH, T.card_rgb, radius_pct=16)
    if cc:
        multi_stop_gradient(cc, [(0, T.card), (50, T.bg2), (100, T.card)], 135)
        shadow(cc, blur=28, dist=8, alpha=0.50)
        glow(cc, T.accent.lstrip('#'), radius=32, alpha=0.10)

    # Top accent
    tp = rrect(slide, 1.4, CY, cw, 0.38, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp, [(0, T.accent2), (50, T.accent), (100, T.accent2)], 0)
        glow(tp, T.accent.lstrip('#'), radius=14, alpha=0.34)

    # Right accent bar
    accent_bar_v(slide, 1.4 + cw - 0.26, CY + 0.38, CH - 0.66, T, width=0.26)

    # Decorative diamonds
    diamond(slide, 1.72, CY + 0.52, 1.0, 1.0, T.accent_rgb, alpha=14)
    diamond(slide, W - 2.85, CY + CH - 1.72, 0.9, 0.9, T.accent_rgb, alpha=10)

    # Opening quote
    txt(slide, "❝", 2.05, CY + 0.5, 1.85, 1.65,
        font='Calibri', size=52, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.LEFT, rtl=False, vcenter=False)

    # Conclusion text — content-aware
    conc_fs = _font_for_len(req.general_conclusion, 14, 11, 16)
    txt(slide, req.general_conclusion, 2.05, CY + 0.95, cw - 1.25, CH - 2.1,
        font=_FONT, size=conc_fs, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
        rtl=True, vcenter=True, line_spacing=1.42)

    # Student name signature
    ny = CY + CH - 1.08
    hl = rect(slide, 1.4 + cw * 0.18, ny, cw * 0.64, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)
    txt(slide, req.student_name, 1.4, ny + 0.12, cw, 0.82,
        font=_FONT, size=22, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'b')
    _hdr(slide, T, "توصيات البحث", "", 'left', slide_num=10, req=req)
    CY = CY0; CH = _ch()
    recs = req.recommendations[:8]
    n = len(recs)
    if not recs: return slide

    gap = 0.16
    row_h, gap = _items_layout(n, CH, gap, min_h=1.0)

    for i, rec in enumerate(recs):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.2: break

        rw = rrect(slide, 1.0, y, W - 2.0, row_h, T.card_rgb, radius_pct=10)
        if rw:
            multi_stop_gradient(rw, [(0, T.card), (100, T.bg2)], 0)
            shadow(rw, blur=7, dist=2, alpha=0.24)

        # Dot + accent bar
        dot = oval(slide, W - 2.02, y + (row_h - 0.44) / 2, 0.44, 0.44, T.accent_rgb)
        if dot:
            multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)
            shadow(dot, blur=6, dist=1, alpha=0.3)
        accent_bar_v(slide, W - 1.32, y, row_h, T, width=0.26)

        rec_fs = _font_for_len(rec, 13, 10, 15)
        txt(slide, rec, 1.22, y, W - 3.68, row_h,
            font=_FONT, size=rec_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.28)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'a')
    _hdr(slide, T, "آفاق البحث المستقبلية", "", 'right', slide_num=11, req=req)
    CY = CY0; CH = _ch()
    items = req.future_work[:6]
    if not items: return slide

    cols = 2 if len(items) > 3 else 1
    rows_n = (len(items) + cols - 1) // cols
    gh = 0.3; gv = 0.22
    col_w = (W - 2.0 - gh * (cols - 1)) / cols
    card_h = (CH - gv * (rows_n - 1)) / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = 1.0 + ci * (col_w + gh)
        y = CY + ri * (card_h + gv)

        cc = content_card_premium(slide, x, y, col_w, card_h, T, radius=14)

        # Top accent strip
        tp = rrect(slide, x, y, col_w, 0.28, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp, [(0, T.accent), (100, T.accent2)], 0)

        # Number badge
        nd = min(1.02, card_h * 0.32)
        number_badge(slide, x + col_w / 2 - nd / 2, y + 0.36, nd, i + 1, T)

        # Divider
        hline(slide, x + col_w * 0.18, y + nd + 0.52, col_w * 0.64, T.muted_rgb, thickness=0.04)

        # Content
        item_fs = _font_for_len(item, 13, 11, 14)
        txt(slide, item, x + 0.32, y + nd + 0.68, col_w - 0.64, card_h - nd - 0.88,
            font=_FONT, size=item_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER,
            rtl=True, vcenter=True, line_spacing=1.32)

    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'c')
    _hdr(slide, T, "قائمة المراجع والمصادر", "", 'left', slide_num=12, req=req)
    CY = CY0; CH = _ch()
    refs = req.references[:12]
    n = len(refs)
    if not refs: return slide

    gap = 0.12
    row_h, gap = _items_layout(n, CH, gap, min_h=0.8)

    for i, ref in enumerate(refs):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.2: break
        even = i % 2 == 0

        rw = rrect(slide, 1.0, y, W - 2.0, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=6)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)

        # Right accent
        accent_bar_v(slide, W - 1.3, y, row_h, T, width=0.24, alpha=55)

        # Reference number
        nb = rrect(slide, 1.14, y + (row_h - 0.38) / 2, 0.72, 0.38, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 65)
        txt(slide, f"[{i+1}]", 1.14, y + (row_h - 0.38) / 2, 0.72, 0.38,
            font='Calibri', size=9, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        ref_fs = max(11, min(14, row_h * 8))
        txt(slide, ref, 1.98, y + 0.06, W - 3.5, row_h - 0.12,
            font=_FONT, size=ref_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.15)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    # Layered background orbs
    oval(slide, -5, -5, 15, 15, T.accent_rgb, alpha=5)
    oval(slide, W - 11, H - 11, 17, 17, T.accent_rgb, alpha=4)
    oval(slide, W - 6.5, -2.5, 10, 10, T.bg2_rgb, alpha=32)
    oval(slide, -2.5, H - 6.5, 9, 9, T.bg2_rgb, alpha=28)
    diamond(slide, W * 0.28, H * 0.06, 2.4, 2.4, T.accent_rgb, alpha=8)
    diamond(slide, W * 0.62, H * 0.74, 1.9, 1.9, T.accent_rgb, alpha=6)
    decorative_dots(slide, 1.4, H - 4.8, 7, 3, 0.17, 0.4, T.accent_rgb, alpha=11)
    decorative_dots(slide, W - 5.8, 1.4, 5, 4, 0.15, 0.36, T.accent_rgb, alpha=9)

    cw = 23; ch = 11.8; cx = (W - cw) / 2; cy = (H - ch) / 2

    # Multi-layer card
    premium_cover_frame(slide, cx, cy, cw, ch, T)

    # Star icon
    txt(slide, "✦", cx + cw / 2 - 0.75, cy + 0.54, 1.5, 1.45,
        font='Calibri', size=34, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Main thank you text
    txt(slide, "شكراً وتقديراً", cx + 0.8, cy + 1.22, cw - 1.6, 2.8,
        font=_FONT, size=SZ_FINAL_TITLE, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.1)

    # Double divider
    d1 = rect(slide, cx + cw * 0.14, cy + 4.22, cw * 0.72, 0.06, T.accent_rgb)
    if d1: multi_stop_gradient(d1, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)
    rect(slide, cx + cw * 0.24, cy + 4.34, cw * 0.52, 0.03, T.muted_rgb)

    # Student name
    txt(slide, req.student_name, cx + 0.8, cy + 4.48, cw - 1.6, 1.4,
        font=_FONT, size=SZ_FINAL_SUB, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # Title
    ts = req.title_ar[:72] + ("..." if len(req.title_ar) > 72 else "")
    txt(slide, ts, cx + 1.2, cy + 5.98, cw - 2.4, 2.2,
        font=_FONT, size=12, bold=False, italic=True, color=T.muted_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.32)

    # Footer info
    footer = []
    if req.institution: footer.append(req.institution)
    if req.year:        footer.append(req.year)
    if footer:
        fb = rrect(slide, cx + cw * 0.1, cy + ch - 1.32, cw * 0.8, 0.62, T.bg_rgb, radius_pct=40)
        if fb: set_solid_alpha(fb, 52)
        txt(slide, "  ·  ".join(footer), cx + 0.8, cy + ch - 1.32, cw - 1.6, 0.62,
            font=_FONT, size=11, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # Bottom accent bar
    bbar = rect(slide, 0, H - 0.28, W, 0.28, T.accent_rgb)
    if bbar: multi_stop_gradient(bbar, [(0, T.bg), (30, T.accent), (70, T.accent2), (100, T.bg)], 0)

    return slide
