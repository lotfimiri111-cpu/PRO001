"""
Classic Engine v28 — هيدر أكاديمي راقٍ + محتوى Premium
أكاديمي نظيف حديث — يبدو وكأن مصمماً عالمياً صنعه يدوياً
"""
from __future__ import annotations
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, decorative_dots,
    number_badge, icon_circle, slide_number, txt, blank_slide,
    accent_bar_v, kpi_card, result_row_premium, content_card_premium,
    methodology_card, highlight_chip, decorative_corner,
    premium_cover_frame, two_col_layout,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"

HEADER_H = 2.65
FOOTER_H = 0.30
MX       = 1.1   # horizontal margin

# ── Typography Scale ────────────────────────────────────────────────
SZ_SECTION_LABEL = 18
SZ_BODY          = 13
SZ_LABEL         = 13
SZ_FINAL_TITLE   = 38
SZ_FINAL_SUB     = 22

def set_font(n): global _FONT; _FONT = n

def _font_for_len(text, base=13, mn=10, mx=15):
    n = len(str(text))
    if n < 40:   return min(mx, base)
    if n < 80:   return base - 1
    if n < 130:  return base - 2
    return max(mn, base - 3)

def _items_h(n, CH, gap=0.16, min_h=0.95):
    rh = (CH - gap * max(0, n-1)) / max(1, n)
    return max(min_h, rh), gap

def _content_y(): return HEADER_H + 0.30
def _content_h(): return H - HEADER_H - FOOTER_H - 0.58


# ══════════════════════════════════════════════════════════════════════
# HEADER v28 — Academic Premium
# ══════════════════════════════════════════════════════════════════════
def _header(slide, T: Theme, title: str, page_num: int = 0, req=None):
    bg(slide, T.bg_rgb)

    # Main gradient
    gradient_rect(slide, 0, 0, W, HEADER_H, T.grad2, T.grad1, angle=90)

    # Right accent bar
    vr = rect(slide, W - 0.48, 0, 0.48, HEADER_H, T.accent_rgb)
    if vr: gradient_fill(vr, T.accent_grad1, T.accent_grad2, 90)

    # Bottom accent line
    al = rect(slide, 0, HEADER_H - 0.22, W, 0.22, T.accent_rgb)
    if al:
        multi_stop_gradient(al, [(0, T.bg), (35, T.accent), (65, T.accent2), (100, T.bg)], 0)
        glow(al, T.accent.lstrip('#'), radius=8, alpha=0.22)

    # Second thin line
    rect(slide, 0, HEADER_H - 0.30, W, 0.06, T.muted_rgb)

    # Decorative orb
    oval(slide, W - 6.5, -2.5, 9, 9, T.accent_rgb, alpha=8)
    decorative_dots(slide, MX, HEADER_H * 0.12, 4, 2, 0.13, 0.30, T.accent_rgb, alpha=10)

    # Page number badge
    title_x = MX
    if page_num > 0:
        nb_s = 0.80
        nb_x = 0.30; nb_y = (HEADER_H - nb_s) / 2
        pb = rrect(slide, nb_x, nb_y, nb_s, nb_s, T.accent_rgb, radius_pct=50)
        if pb:
            multi_stop_gradient(pb, [(0, T.accent), (100, T.accent2)], 135)
            shadow(pb, blur=8, dist=2, alpha=0.35)
        txt(slide, str(page_num), nb_x, nb_y, nb_s, nb_s,
            font='Calibri', size=15, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        total = getattr(req, '_total_slides', 13) if req else 13
        txt(slide, f'/{total}', nb_x + nb_s, nb_y + nb_s * 0.34, 0.82, nb_s * 0.38,
            font='Calibri', size=7, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
        title_x = MX + 1.32

    # Content-aware title
    tlen = len(title)
    tfs = 28 if tlen <= 12 else 26 if tlen <= 20 else 23 if tlen <= 32 else 20
    title_w = W - title_x - 0.58
    txt(slide, title, title_x, 0.22, title_w, HEADER_H - 0.48,
        font=_FONT, size=tfs, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.06)

    # Content background
    gradient_rect(slide, 0, HEADER_H, W, H - HEADER_H, T.bg, T.bg2, angle=90)
    decorative_corner(slide, T, 'br', 4.5)

    # Footer
    bt = rect(slide, 0, H - FOOTER_H, W, FOOTER_H, T.bg2_rgb)
    bta = rect(slide, 0, H - FOOTER_H, W, 0.06, T.accent_rgb)
    if bta: gradient_fill(bta, T.accent_grad1, T.accent_grad2, 0)


# ══════════════════════════════════════════════════════════════════════
# COVER — Classic Premium
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)

    # Decorative orbs
    oval(slide, W - 9.5, -2.5, 12.5, 12.5, T.accent_rgb, alpha=5)
    oval(slide, -2.5, H - 8.5, 10.5, 10.5, T.bg2_rgb, alpha=34)
    diamond(slide, W * 0.28, H * 0.06, 2.5, 2.5, T.accent_rgb, alpha=8)
    diamond(slide, W * 0.62, H * 0.76, 2.0, 2.0, T.accent_rgb, alpha=6)
    decorative_dots(slide, 1.6, H - 4.8, 6, 3, 0.16, 0.40, T.accent_rgb, alpha=10)

    # Frame lines
    r_top = rect(slide, 0, 0, W, 0.24, T.accent_rgb)
    if r_top: gradient_fill(r_top, T.accent_grad1, T.accent_grad2, 0)
    r_bot = rect(slide, 0, H - 0.24, W, 0.24, T.accent_rgb)
    if r_bot: gradient_fill(r_bot, T.accent_grad1, T.accent_grad2, 0)
    vline(slide, 0.24, 0.24, H - 0.48, T.accent_rgb, thickness=0.08)
    vline(slide, W - 0.32, 0.24, H - 0.48, T.accent_rgb, thickness=0.08)

    # Institution
    if req.institution:
        txt(slide, req.institution, 2.5, 0.40, W - 5.0, 0.88,
            font=_FONT, size=11.5, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    hline(slide, W * 0.18, 1.36, W * 0.64, T.accent_rgb, thickness=0.05)

    import re as _re
    _yr = _re.compile(r'\b\d{4}\s*[-–—]\s*\d{4}\b')
    _m = _yr.search(req.title_ar or '')
    if _m:
        _year_str    = _m.group(0).strip()
        _title_clean = _yr.sub('', req.title_ar).strip(' —–-،, ')
    elif req.year:
        _year_str    = req.year.strip()
        _title_clean = req.title_ar
    else:
        _year_str    = ''
        _title_clean = req.title_ar

    title_y = 1.52; total_h = H - 0.24 - title_y
    title_h = total_h * 0.42; info_y = title_y + title_h + 0.24
    info_h = H - 0.24 - info_y - 0.12

    ts = 26 if len(_title_clean) < 42 else 21 if len(_title_clean) < 65 else 17
    _tth = title_h * (0.56 if _year_str else 0.70)
    txt(slide, _title_clean, 2.2, title_y, W - 4.4, _tth,
        font=_FONT, size=ts, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.22)

    if req.title_en:
        txt(slide, req.title_en, 2.5, title_y + _tth + 0.06, W - 5.0, 0.76,
            font='Calibri', size=11, bold=False, italic=True, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    if _year_str:
        _yr_y = title_y + title_h * 0.74
        _yr_h = 0.60
        _yr_cx = W * 0.25; _yr_cw = W * 0.50
        yb = rrect(slide, _yr_cx, _yr_y, _yr_cw, _yr_h, T.accent_rgb, radius_pct=50)
        if yb: set_solid_alpha(yb, 22)
        hline(slide, _yr_cx + _yr_cw*0.08, _yr_y,          _yr_cw*0.84, T.accent_rgb, thickness=0.03)
        hline(slide, _yr_cx + _yr_cw*0.08, _yr_y + _yr_h,  _yr_cw*0.84, T.accent_rgb, thickness=0.03)
        txt(slide, f'( {_year_str} )', _yr_cx, _yr_y, _yr_cw, _yr_h,
            font=_FONT, size=13, bold=False, italic=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    hl1 = rect(slide, W*0.12, info_y - 0.18, W*0.76, 0.07, T.accent_rgb)
    if hl1: multi_stop_gradient(hl1, [(0, T.bg), (50, T.accent), (100, T.bg)], 0)
    rect(slide, W*0.20, info_y - 0.08, W*0.60, 0.03, T.muted_rgb)

    rows = [("اسم الطالب", req.student_name)]
    if req.supervisor:     rows.append(("المشرف",         req.supervisor))
    if req.co_supervisor:  rows.append(("المشرف المساعد", req.co_supervisor))
    if req.specialization: rows.append(("التخصص",         req.specialization))

    rh = info_h / max(len(rows), 1)
    row_w = W - MX * 2
    lbl_w = 3.9; val_w = row_w - lbl_w - 0.55
    for i, (lbl, val) in enumerate(rows):
        y = info_y + i * rh
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rb = rrect(slide, MX, y, row_w, rh - 0.06, fill, radius_pct=6)
        if rb: shadow(rb, blur=4, dist=1, alpha=0.15)

        acc = rect(slide, W - MX - 0.22, y, 0.22, rh - 0.06, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        lbl_x = W - MX - lbl_w - 0.24
        val_x = MX + 0.34
        txt(slide, f"{lbl} :", lbl_x, y, lbl_w, rh - 0.06,
            font=_FONT, size=max(13, min(15, rh * 8.5)), bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        vline(slide, lbl_x - 0.16, y + rh*0.1, rh*0.7, T.muted_rgb, thickness=0.04)
        fs = _font_for_len(val, 14, 11, 16)
        txt(slide, val, val_x, y, val_w, rh - 0.06,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "مقدمة البحث", 1, req)
    CY = _content_y(); CH = _content_h()
    items = []
    if req.intro_overview: items.append(("نظرة عامة",    req.intro_overview))
    if req.intro_approach:  items.append(("المنهج المتبع", req.intro_approach))
    if not items: return slide

    gap = 0.26
    card_h = (CH - gap * (len(items) - 1)) / len(items)

    for i, (lbl, val) in enumerate(items[:2]):
        y = CY + i * (card_h + gap)
        even = i % 2 == 0

        # Card
        cb = rrect(slide, MX, y, W - MX*2, card_h,
                   T.bg2_rgb if even else T.card_rgb, radius_pct=10)
        if cb:
            stops = [(0, T.bg2), (100, T.card)] if even else [(0, T.card), (100, T.bg2)]
            multi_stop_gradient(cb, stops, 0)
            shadow(cb, blur=16, dist=4, alpha=0.34)

        # Header band
        hdr = rrect(slide, MX, y, W - MX*2, 0.64, T.accent_rgb, radius_pct=0)
        if hdr: multi_stop_gradient(hdr, [(0, T.accent), (100, T.accent2)], 0)

        # Right accent bar
        accent_bar_v(slide, W - MX - 0.26, y, card_h, T, width=0.24)

        # Label
        txt(slide, lbl, MX + 0.22, y, W - MX*2 - 0.62, 0.64,
            font=_FONT, size=SZ_SECTION_LABEL, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        # Content
        val_fs = _font_for_len(val, 13, 11, 15)
        txt(slide, val, MX + 0.30, y + 0.72, W - MX*2 - 0.82, card_h - 0.84,
            font=_FONT, size=val_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.35)
    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "خطة البحث", 2, req)
    CY = _content_y(); CH = _content_h()
    chapters = req.chapters[:8]; n = len(chapters)
    if not chapters: return slide

    row_h, gap = _items_h(n, CH, 0.14, 0.92)

    for i, ch in enumerate(chapters):
        y = CY + i * (row_h + gap)
        if y + row_h > H - FOOTER_H - 0.1: break
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb

        rw = rrect(slide, MX, y, W - MX*2, row_h, fill, radius_pct=9)
        if rw:
            stops = [(0, T.bg2), (100, T.card)] if i%2==0 else [(0, T.card), (100, T.bg2)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=6, dist=2, alpha=0.18)

        # Right accent
        accent_bar_v(slide, W - MX - 0.24, y, row_h, T, width=0.22)

        # Number badge
        nd = min(0.70, row_h * 0.72)
        nx = W - MX - 1.72; ny = y + (row_h - nd) / 2
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
            shadow(nc, blur=7, dist=2, alpha=0.28)
        txt(slide, str(i + 1), nx, ny, nd, nd,
            font='Calibri', size=max(9, int(nd * 11)), bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # Chapter title
        title_fs = _font_for_len(ch.title, 13, 11, 15)
        title_fs = max(title_fs, min(15, int(row_h * 9.5)))
        txt(slide, ch.title, MX + 0.25, y, W - MX*2 - 2.55, row_h,
            font=_FONT, size=title_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        if ch.pages:
            pb = rrect(slide, MX + 0.20, y + (row_h - 0.36)/2, 2.0, 0.36, T.bg_rgb, radius_pct=40)
            if pb: set_solid_alpha(pb, 52)
            txt(slide, ch.pages, MX + 0.20, y + (row_h - 0.36)/2, 2.0, 0.36,
                font='Calibri', size=9, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "إشكالية البحث", 3, req)
    CY = _content_y(); CH = _content_h()
    cw = W - MX * 2

    secs = []
    if req.main_problem:  secs.append('p')
    if req.main_question: secs.append('q')
    if req.sub_questions: secs.append('s')
    if not secs: return slide

    weights = {'p': 2.8, 'q': 1.5, 's': 2.0}
    tw = sum(weights[s] for s in secs)
    gap = 0.20; avail = CH - gap * (len(secs) - 1)
    cy = CY

    if 'p' in secs:
        h = avail * weights['p'] / tw
        pc = rrect(slide, MX, cy, cw, h, T.card_rgb, radius_pct=12)
        if pc:
            multi_stop_gradient(pc, [(0, T.card), (100, T.bg2)], 135)
            shadow(pc, blur=18, dist=5, alpha=0.38)
            glow(pc, T.accent.lstrip('#'), radius=20, alpha=0.07)

        lb = rrect(slide, W - MX - 6.2, cy, 5.8, 0.52, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb, [(0, T.accent), (100, T.accent2)], 0)
        txt(slide, "◆  الإشكالية الرئيسية", W - MX - 6.2, cy, 5.8, 0.52,
            font=_FONT, size=11, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        accent_bar_v(slide, W - MX - 0.24, cy, h, T, width=0.22)

        txt(slide, "❝", MX + 0.32, cy + 0.58, 1.5, 1.2,
            font='Calibri', size=32, bold=False, color=T.accent_rgb,
            align=PP_ALIGN.LEFT, rtl=False)

        prob_fs = _font_for_len(req.main_problem, 13, 10, 14)
        txt(slide, req.main_problem, MX + 1.95, cy + 0.56, cw - 2.45, h - 0.72,
            font=_FONT, size=prob_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.38)
        cy += h + gap

    if 'q' in secs:
        h = avail * weights['q'] / tw
        qc = rrect(slide, MX, cy, cw, h, T.bg2_rgb, radius_pct=9)
        if qc: shadow(qc, blur=7, dist=2, alpha=0.20)
        vline(slide, W - MX - 0.20, cy, h, T.accent_rgb, thickness=0.20)
        dot = oval(slide, W - MX - 2.85, cy + h/2 - 0.19, 0.40, 0.40, T.accent_rgb)
        if dot: multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)
        q_fs = _font_for_len(req.main_question, 13, 11, 14)
        txt(slide, req.main_question, MX + 0.30, cy, cw - 1.55, h,
            font=_FONT, size=q_fs, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        cy += h + gap

    if 's' in secs and req.sub_questions:
        h = avail * weights['s'] / tw
        subs = req.sub_questions[:5]
        sub_h = h / max(1, len(subs))
        for i, q in enumerate(subs):
            y2 = cy + i * sub_h
            if i % 2 == 0:
                sb = rrect(slide, MX, y2, cw, sub_h - 0.05, T.card_rgb, radius_pct=6)
                if sb: set_solid_alpha(sb, 46)
            nd = min(0.42, sub_h * 0.54)
            nc = oval(slide, W - MX - 2.6, y2 + (sub_h - nd)/2, nd, nd, T.accent_rgb)
            if nc: set_solid_alpha(nc, 62)
            txt(slide, str(i+1), W - MX - 2.6, y2 + (sub_h - nd)/2, nd, nd,
                font='Calibri', size=max(7, int(nd*9)), bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
            q_fs2 = _font_for_len(q, 12, 10, 13)
            txt(slide, q, MX + 0.28, y2, cw - 3.2, sub_h,
                font=_FONT, size=q_fs2, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "أهداف البحث وفرضياته", 4, req)
    CY = _content_y(); CH = _content_h()
    cols = []
    if req.objectives: cols.append(("🎯  الأهداف",   req.objectives))
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
    _header(slide, T, "أهمية البحث", 5, req)
    CY = _content_y(); CH = _content_h()
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐", "🔑", "📌", "🌟", "💎", "🏆"]
    cols = 2 if len(items) > 3 else 1
    rows_n = (len(items) + cols - 1) // cols
    gh = 0.22; gv = 0.18
    cw = (W - MX*2 - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = MX + ci*(cw+gh); y = CY + ri*(ch+gv)

        content_card_premium(slide, x, y, cw, ch, T, accent_side='right', radius=10)

        ic_s = min(1.32, ch * 0.62)
        icon_circle(slide, x + 0.28, y + (ch - ic_s)/2,
                    ic_s, T.accent_grad1, T.accent_grad2,
                    icons[i % len(icons)], max(12, int(ic_s*11)), T)

        item_fs = _font_for_len(item, 13, 11, 14)
        txt(slide, item, x + ic_s + 0.58, y + 0.12, cw - ic_s - 1.12, ch - 0.24,
            font=_FONT, size=item_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.32)
    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "منهجية البحث", 6, req)
    CY = _content_y(); CH = _content_h()
    icons_map = {"المنهج": "📊", "العينة": "👥", "حجم العينة": "📏", "الأداة": "🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج",     req.methodology))
    if req.sample_type:  fields.append(("العينة",     req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة",     req.tool))
    if not fields: return slide

    cols = 2 if len(fields) > 2 else len(fields)
    rows_n = (len(fields) + cols - 1) // cols
    gh = 0.22; gv = 0.20
    cw = (W - MX*2 - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i % cols; ri = i // cols
        x = MX + ci*(cw+gh); y = CY + ri*(ch+gv)
        methodology_card(slide, x, y, cw, ch, lbl, val, icons_map.get(lbl, "📌"), T)
    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "الأرقام والإحصاءات الرئيسية", 7, req)
    CY = _content_y(); CH = _content_h()
    stats = req.stats[:6]
    if not stats: return slide

    cols = 3 if len(stats) >= 3 else len(stats)
    rows_n = (len(stats) + cols - 1) // cols
    gh = 0.26; gv = 0.20
    cw = (W - MX*2 - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, stat in enumerate(stats):
        ci = i % cols; ri = i // cols
        x = MX + ci*(cw+gh); y = CY + ri*(ch+gv)
        if y + ch > H - FOOTER_H - 0.1: break
        kpi_card(slide, x, y, cw, ch, stat.value, stat.label, T,
                 unit=stat.unit or '', rank=i)
    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "نتائج البحث", 8, req)
    CY = _content_y(); CH = _content_h()
    results = req.main_results[:8]
    if not results: return slide

    row_h, gap = _items_h(len(results), CH, 0.14, 1.05)
    for i, result in enumerate(results):
        y = CY + i * (row_h + gap)
        if y + row_h > H - FOOTER_H - 0.1: break
        result_row_premium(slide, MX, y, W - MX*2, row_h, result, i+1, T, highlight=(i==0))
    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "خاتمة البحث", 9, req)
    CY = _content_y(); CH = _content_h()
    cw = W - MX * 2

    cc = rrect(slide, MX, CY, cw, CH, T.card_rgb, radius_pct=12)
    if cc:
        multi_stop_gradient(cc, [(0, T.card), (50, T.bg2), (100, T.card)], 135)
        shadow(cc, blur=22, dist=6, alpha=0.42)
        glow(cc, T.accent.lstrip('#'), radius=26, alpha=0.08)

    tp = rrect(slide, MX, CY, cw, 0.32, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp, [(0, T.accent2), (50, T.accent), (100, T.accent2)], 0)

    accent_bar_v(slide, W - MX - 0.24, CY + 0.32, CH - 0.56, T, width=0.22)
    diamond(slide, MX + 0.32, CY + 0.50, 0.95, 0.95, T.accent_rgb, alpha=13)

    txt(slide, "❝", MX + 0.36, CY + 0.48, 1.65, 1.45,
        font='Calibri', size=48, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.LEFT, rtl=False)

    conc_fs = _font_for_len(req.general_conclusion, 14, 11, 16)
    txt(slide, req.general_conclusion, MX + 0.36, CY + 1.05, cw - 0.88, CH - 2.1,
        font=_FONT, size=conc_fs, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.40)

    ny = CY + CH - 0.98
    hl = rect(slide, MX + cw*0.18, ny, cw*0.64, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)
    txt(slide, req.student_name, MX, ny + 0.12, cw, 0.78,
        font=_FONT, size=22, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "توصيات البحث", 10, req)
    CY = _content_y(); CH = _content_h()
    recs = req.recommendations[:8]
    if not recs: return slide

    row_h, gap = _items_h(len(recs), CH, 0.16, 0.98)
    for i, rec in enumerate(recs):
        y = CY + i * (row_h + gap)
        if y + row_h > H - FOOTER_H - 0.1: break

        rw = rrect(slide, MX, y, W - MX*2, row_h, T.card_rgb, radius_pct=9)
        if rw:
            multi_stop_gradient(rw, [(0, T.card), (100, T.bg2)], 0)
            shadow(rw, blur=6, dist=2, alpha=0.20)

        dot = oval(slide, W - MX - 1.94, y + (row_h - 0.42)/2, 0.42, 0.42, T.accent_rgb)
        if dot:
            multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)
            shadow(dot, blur=5, dist=1, alpha=0.26)
        accent_bar_v(slide, W - MX - 0.24, y, row_h, T, width=0.22)

        rec_fs = _font_for_len(rec, 13, 11, 15)
        txt(slide, rec, MX + 0.24, y, W - MX*2 - 2.72, row_h,
            font=_FONT, size=rec_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.26)
    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "آفاق البحث المستقبلية", 11, req)
    CY = _content_y(); CH = _content_h()
    items = req.future_work[:6]
    if not items: return slide

    cols = 2 if len(items) > 3 else 1
    rows_n = (len(items) + cols - 1) // cols
    gh = 0.24; gv = 0.20
    cw = (W - MX*2 - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = MX + ci*(cw+gh); y = CY + ri*(ch+gv)

        content_card_premium(slide, x, y, cw, ch, T, radius=12)

        tp = rrect(slide, x, y, cw, 0.24, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp, [(0, T.accent), (100, T.accent2)], 0)

        nd = min(0.86, ch * 0.32)
        number_badge(slide, x + cw/2 - nd/2, y + 0.36, nd, i+1, T)

        hline(slide, x + cw*0.18, y + nd + 0.52, cw*0.64, T.muted_rgb, thickness=0.04)

        item_fs = _font_for_len(item, 13, 11, 14)
        txt(slide, item, x + 0.28, y + nd + 0.68, cw - 0.56, ch - nd - 0.88,
            font=_FONT, size=item_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.32)
    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "قائمة المراجع والمصادر", 12, req)
    CY = _content_y(); CH = _content_h()
    refs = req.references[:12]
    if not refs: return slide

    row_h, gap = _items_h(len(refs), CH, 0.12, 0.80)
    for i, ref in enumerate(refs):
        y = CY + i * (row_h + gap)
        if y + row_h > H - FOOTER_H - 0.1: break
        even = i % 2 == 0

        rw = rrect(slide, MX, y, W - MX*2, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=6)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)

        accent_bar_v(slide, W - MX - 0.22, y, row_h, T, width=0.20, alpha=55)

        nb = rrect(slide, MX + 0.12, y + (row_h - 0.36)/2, 0.70, 0.36, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 60)
        txt(slide, f"[{i+1}]", MX + 0.12, y + (row_h - 0.36)/2, 0.70, 0.36,
            font='Calibri', size=9, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        ref_fs = max(10, min(13, row_h * 8))
        txt(slide, ref, MX + 0.94, y + 0.04, W - MX*2 - 1.40, row_h - 0.08,
            font=_FONT, size=ref_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)

    # Decorative
    oval(slide, W - 9.5, -2.5, 12.5, 12.5, T.accent_rgb, alpha=5)
    oval(slide, -2.5, H - 8.5, 10.5, 10.5, T.bg2_rgb, alpha=32)
    diamond(slide, W*0.25, H*0.06, 2.3, 2.3, T.accent_rgb, alpha=7)
    diamond(slide, W*0.65, H*0.76, 1.9, 1.9, T.accent_rgb, alpha=6)
    decorative_dots(slide, 1.6, H - 4.8, 6, 3, 0.16, 0.40, T.accent_rgb, alpha=9)
    decorative_dots(slide, W - 6.2, 1.4, 4, 4, 0.14, 0.34, T.accent_rgb, alpha=8)

    # Frame
    r_top = rect(slide, 0, 0, W, 0.24, T.accent_rgb)
    if r_top: gradient_fill(r_top, T.accent_grad1, T.accent_grad2, 0)
    r_bot = rect(slide, 0, H - 0.24, W, 0.24, T.accent_rgb)
    if r_bot: gradient_fill(r_bot, T.accent_grad1, T.accent_grad2, 0)
    vline(slide, 0.24, 0.24, H - 0.48, T.accent_rgb, thickness=0.08)
    vline(slide, W - 0.32, 0.24, H - 0.48, T.accent_rgb, thickness=0.08)

    cw = W - MX * 2
    cy = H * 0.10; ch = H * 0.80

    cb = rrect(slide, MX, cy, cw, ch, T.bg2_rgb, radius_pct=12)
    if cb:
        multi_stop_gradient(cb, [(0, T.bg2), (50, T.card), (100, T.bg2)], 135)
        shadow(cb, blur=30, dist=9, alpha=0.52)
        glow(cb, T.accent.lstrip('#'), radius=34, alpha=0.10)

    acc_top = rrect(slide, MX, cy, cw, 0.38, T.accent_rgb, radius_pct=0)
    if acc_top:
        multi_stop_gradient(acc_top, [(0, T.accent), (50, T.accent2), (100, T.accent)], 0)
        glow(acc_top, T.accent.lstrip('#'), radius=14, alpha=0.28)
    acc_bot = rect(slide, MX, cy + ch - 0.26, cw, 0.26, T.accent_rgb)
    if acc_bot: set_solid_alpha(acc_bot, 42)
    acc_r = rect(slide, MX + cw - 0.22, cy + 0.38, 0.22, ch - 0.64, T.accent_rgb)
    if acc_r:
        gradient_fill(acc_r, T.accent_grad1, T.accent_grad2, 90)

    txt(slide, "✦", MX + 0.28, cy + 0.52, cw - 0.56, H*0.12,
        font='Calibri', size=28, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    txt(slide, "شكراً وتقديراً", MX + 0.28, cy + 0.60, cw - 0.56, H * 0.24,
        font=_FONT, size=SZ_FINAL_TITLE, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.1)

    hl1 = rect(slide, MX + cw*0.15, cy + H*0.30, cw*0.70, 0.07, T.accent_rgb)
    if hl1: multi_stop_gradient(hl1, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)
    rect(slide, MX + cw*0.25, cy + H*0.30 + 0.16, cw*0.50, 0.03, T.muted_rgb)

    txt(slide, req.student_name, MX + 0.28, cy + H*0.33, cw - 0.56, H * 0.14,
        font=_FONT, size=SZ_FINAL_SUB, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    ts = req.title_ar[:70] + ("..." if len(req.title_ar) > 70 else "")
    txt(slide, ts, MX + 0.55, cy + H*0.49, cw - 1.10, H * 0.17,
        font=_FONT, size=SZ_BODY, bold=False, italic=True, color=T.muted_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.28)

    footer = []
    if req.institution: footer.append(req.institution)
    if req.year:        footer.append(req.year)
    if footer:
        fb = rrect(slide, MX + cw*0.1, cy + H*0.69, cw*0.8, 0.62, T.bg_rgb, radius_pct=40)
        if fb: set_solid_alpha(fb, 50)
        txt(slide, "  ·  ".join(footer), MX + 0.28, cy + H*0.69, cw - 0.56, 0.62,
            font=_FONT, size=SZ_BODY, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)
    return slide
