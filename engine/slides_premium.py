"""
Premium Engine v28 — شريط جانبي فخم + محتوى Premium
نظام تصميم ذكي بمستوى Envato / Pitch Deck Agency
"""
from __future__ import annotations
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, hexagon, decorative_dots,
    number_badge, icon_circle, slide_number, card_3d,
    txt, blank_slide,
    accent_bar_v, kpi_card, result_row_premium, content_card_premium,
    methodology_card, highlight_chip, decorative_corner,
    premium_cover_frame, two_col_layout,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(n): global _FONT; _FONT = n

# ── Typography Scale ─────────────────────────────────────────────────
SZ_SIDEBAR_TITLE = 22
SZ_SECTION_LABEL = 18
SZ_BODY          = 13
SZ_LABEL         = 13
SZ_FINAL_TITLE   = 38
SZ_FINAL_SUB     = 22

SW  = 5.2          # sidebar width
CX  = SW + 0.55    # content start x
CW  = W - CX - 0.45  # content width

# ── Content helpers ──────────────────────────────────────────────────
def _font_for_len(text, base=13, mn=10, mx=15):
    n = len(str(text))
    if n < 40:   return min(mx, base)
    if n < 80:   return base - 1
    if n < 130:  return base - 2
    return max(mn, base - 3)

def _items_h(n, CH, gap=0.16, min_h=0.95):
    rh = (CH - gap * max(0, n-1)) / max(1, n)
    return max(min_h, rh), gap


# ══════════════════════════════════════════════════════════════════════
# SIDEBAR — Premium v28
# ══════════════════════════════════════════════════════════════════════
def _sidebar(slide, T, icon, title1, title2="", slide_num_n=0, total=13):
    # ── Base gradient ──
    gradient_rect(slide, 0, 0, SW, H, T.grad2, T.grad1, angle=185)

    # Separator line
    sep = rect(slide, SW, 0, 0.14, H, T.accent_rgb)
    if sep: gradient_fill(sep, T.accent_grad1, T.accent_grad2, 90)

    # Decorative orbs
    oval(slide, -3.5, -3.5, 10, 10, T.accent_rgb, alpha=7)
    oval(slide, 0.3, H - 6.5, 7.5, 7.5, T.bg2_rgb, alpha=42)
    oval(slide, SW - 2.5, H * 0.38, 3.2, 3.2, T.accent_rgb, alpha=10)
    decorative_dots(slide, 0.45, H * 0.58, 5, 4, 0.13, 0.30, T.accent_rgb, alpha=13)
    diamond(slide, SW * 0.1, H * 0.12, 1.2, 1.2, T.accent_rgb, alpha=9)

    # Top accent bar
    ta = rect(slide, 0, 0, SW, 0.36, T.accent_rgb)
    if ta: multi_stop_gradient(ta, [(0, T.accent), (100, T.accent2)], 0)

    # ── Icon circle ──
    ic_y = H * 0.14
    ic_s = 3.1
    ic_bg = oval(slide, SW/2 - ic_s/2, ic_y, ic_s, ic_s, T.accent_rgb)
    if ic_bg:
        multi_stop_gradient(ic_bg, [(0, T.accent), (100, T.accent2)], 135)
        shadow(ic_bg, blur=18, dist=5, alpha=0.45)
        glow(ic_bg, T.accent.lstrip('#'), radius=24, alpha=0.22)

    # Inner ring
    ring_s = ic_s + 0.28
    ring = oval(slide, SW/2 - ring_s/2, ic_y - 0.14, ring_s, ring_s, T.accent_rgb)
    if ring: set_solid_alpha(ring, 18)

    txt(slide, icon, SW/2 - ic_s/2, ic_y + 0.28, ic_s, ic_s * 0.82,
        font='Calibri', size=32, bold=False,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # ── Section title ──
    label_y = ic_y + ic_s + 0.45
    title_len = len(title1)
    fs = 22 if title_len <= 6 else 19 if title_len <= 10 else 16
    txt(slide, title1, 0.22, label_y, SW - 0.44, 1.2,
        font=_FONT, size=fs, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    if title2:
        txt(slide, title2, 0.22, label_y + 1.24, SW - 0.44, 0.84,
            font=_FONT, size=14, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # Divider
    div_y = label_y + (2.16 if title2 else 1.32)
    d1 = rect(slide, 0.55, div_y, SW - 1.1, 0.05, T.accent_rgb)
    if d1: multi_stop_gradient(d1, [(0, T.bg), (50, T.accent), (100, T.bg)], 0)

    # Slide counter at bottom
    if slide_num_n > 0:
        cy_n = H - 1.55
        nb = rrect(slide, SW/2 - 1.2, cy_n, 2.4, 0.72, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 55)
        txt(slide, f"{slide_num_n} / {total}", SW/2 - 1.2, cy_n, 2.4, 0.72,
            font='Calibri', size=11, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # ── Content area background ──
    rect(slide, SW + 0.14, 0, W - SW - 0.14, H, T.bg2_rgb)
    gradient_rect(slide, SW + 0.14, 0, W - SW - 0.14, H, T.grad1, T.bg2, angle=135)
    decorative_corner(slide, T, 'tr', 5.5)
    diamond(slide, W - 5.2, H - 5.2, 4.2, 4.2, T.accent_rgb, alpha=5)


def _section_slide(prs, T, icon, t1, t2="", sn=0, total=13):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _sidebar(slide, T, icon, t1, t2, sn, total)
    return slide


# ══════════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)

    # Sidebar
    gradient_rect(slide, 0, 0, SW + 0.5, H, T.grad2, T.grad1, angle=185)
    sep = rect(slide, SW + 0.5, 0, 0.14, H, T.accent_rgb)
    if sep: gradient_fill(sep, T.accent_grad1, T.accent_grad2, 90)
    oval(slide, -3.5, -3.5, 10, 10, T.accent_rgb, alpha=7)
    oval(slide, 0.3, H - 6.5, 7.5, 7.5, T.bg2_rgb, alpha=40)
    decorative_dots(slide, 0.45, H * 0.55, 5, 4, 0.13, 0.30, T.accent_rgb, alpha=12)

    # Institution
    if req.institution:
        txt(slide, req.institution, 0.2, H * 0.38, SW + 0.28, 2.4,
            font=_FONT, size=9.5, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.2)

    # Icon
    ic_s = 3.3
    ic_bg = oval(slide, SW / 2 - ic_s / 2 - 0.1, H * 0.06, ic_s, ic_s, T.accent_rgb)
    if ic_bg:
        multi_stop_gradient(ic_bg, [(0, T.accent), (100, T.accent2)], 135)
        shadow(ic_bg, blur=16, dist=5, alpha=0.44)
        glow(ic_bg, T.accent.lstrip('#'), radius=22, alpha=0.22)
    txt(slide, "🎓", SW/2 - ic_s/2 - 0.1, H * 0.06 + 0.32, ic_s, ic_s * 0.82,
        font='Calibri', size=34, bold=False, color=T.text_dark_rgb,
        align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Year badge
    if req.year:
        yb = rrect(slide, 0.44, H - 1.52, SW - 0.32, 0.66, T.accent_rgb, radius_pct=40)
        if yb: multi_stop_gradient(yb, [(0, T.accent), (100, T.accent2)], 0)
        txt(slide, req.year, 0.44, H - 1.52, SW - 0.32, 0.66,
            font='Calibri', size=13, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Content area
    gradient_rect(slide, SW + 0.64, 0, W - SW - 0.64, H, T.grad1, T.bg2, angle=135)
    decorative_corner(slide, T, 'tr', 5.5)
    diamond(slide, W - 5.2, H - 5.2, 4.2, 4.2, T.accent_rgb, alpha=5)

    mcx = SW + 1.0; mcw = W - mcx - 0.75
    tp = rect(slide, mcx - 0.3, 0, mcw + 0.3, 0.36, T.accent_rgb)
    if tp: multi_stop_gradient(tp, [(0, T.bg), (40, T.accent), (60, T.accent2), (100, T.bg)], 0)

    import re as _re
    _yr = _re.compile(r'\b\d{4}\s*[-–—]\s*\d{4}\b')
    _m = _yr.search(req.title_ar or '')
    if _m:
        _year_str   = _m.group(0).strip()
        _title_clean = _yr.sub('', req.title_ar).strip(' —–-،, ')
    elif req.year:
        _year_str   = req.year.strip()
        _title_clean = req.title_ar
    else:
        _year_str   = ''
        _title_clean = req.title_ar

    title_y = 0.50; title_h = 7.4
    info_y = title_y + title_h + 0.22
    info_h = H - info_y - 0.34

    # Cover frame
    premium_cover_frame(slide, mcx, title_y, mcw, title_h, T)

    # Title text
    ts = 24 if len(_title_clean) < 42 else 20 if len(_title_clean) < 65 else 16
    _tth = title_h * (0.54 if _year_str else 0.64)
    txt(slide, _title_clean, mcx + 0.55, title_y + 0.44, mcw - 1.05, _tth,
        font=_FONT, size=ts, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.22)

    if req.title_en:
        txt(slide, req.title_en, mcx + 0.55, title_y + _tth + 0.34, mcw - 1.05, title_h * 0.12,
            font='Calibri', size=10, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    if _year_str:
        yy = title_y + title_h * 0.74
        yb2 = rrect(slide, mcx + mcw * 0.18, yy, mcw * 0.64, title_h * 0.12,
                    T.accent_rgb, radius_pct=50)
        if yb2: set_solid_alpha(yb2, 24)
        hline(slide, mcx + mcw * 0.24, yy, mcw * 0.52, T.accent_rgb, thickness=0.03)
        hline(slide, mcx + mcw * 0.24, yy + title_h * 0.12, mcw * 0.52, T.accent_rgb, thickness=0.03)
        txt(slide, f'( {_year_str} )', mcx + mcw * 0.18, yy,
            mcw * 0.64, title_h * 0.12,
            font=_FONT, size=12, bold=False, italic=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    hl = rect(slide, mcx + mcw * 0.08, title_y + title_h * 0.91, mcw * 0.84, 0.05, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)

    # Info card
    ic2 = rrect(slide, mcx, info_y, mcw, info_h, T.card_rgb, radius_pct=12)
    if ic2:
        multi_stop_gradient(ic2, [(0, T.bg2), (100, T.card)], 135)
        shadow(ic2, blur=14, dist=4, alpha=0.34)
    vline(slide, mcx + mcw - 0.2, info_y, info_h, T.accent_rgb, thickness=0.2)

    rows = [("الطالب", req.student_name)]
    if req.supervisor:     rows.append(("المشرف",         req.supervisor))
    if req.co_supervisor:  rows.append(("المشرف المساعد", req.co_supervisor))
    if req.specialization: rows.append(("التخصص",         req.specialization))

    rh = info_h / max(len(rows), 1)
    for i, (lbl, val) in enumerate(rows):
        y = info_y + i * rh
        rb = rrect(slide, mcx + 0.25, y + 0.04, mcw - 0.62, rh - 0.08, T.bg_rgb, radius_pct=7)
        if rb: set_solid_alpha(rb, 50)
        lbl_w = 4.4; val_w = mcw - lbl_w - 0.95
        lbl_x = mcx + mcw - lbl_w - 0.3
        val_x = mcx + 0.45
        txt(slide, f"{lbl} :", lbl_x, y + 0.04, lbl_w, rh - 0.08,
            font=_FONT, size=max(12, min(14, rh * 8.5)), bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        vline(slide, lbl_x - 0.18, y + rh * 0.12, rh * 0.76, T.muted_rgb, thickness=0.04)
        fs = _font_for_len(val, 14, 11, 16)
        txt(slide, val, val_x, y + 0.04, val_w, rh - 0.08,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.LEFT, rtl=False, vcenter=True)

    bt = rect(slide, 0, H - 0.26, W, 0.26, T.accent_rgb)
    if bt: multi_stop_gradient(bt, [(0, T.bg), (30, T.accent), (70, T.accent2), (100, T.bg)], 0)
    return slide


# ══════════════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "📖", "مقدمة", "البحث", 1, total)
    CY = 0.28; CH = H - CY - 0.22
    items = []
    if req.intro_overview: items.append(("📖", "نظرة عامة",  req.intro_overview))
    if req.intro_approach:  items.append(("🔬", "المنهج المتبع", req.intro_approach))
    if not items: return slide

    n = len(items); gap = 0.28
    cw = (CW - gap * (n - 1)) / n
    CARD_H = min(CH * 0.90, 12.2)
    card_y = CY + (CH - CARD_H) / 2

    ic_s = min(1.78, CARD_H * 0.22)
    ic_y_off = 0.52; lbl_h = 0.72
    lbl_y_off = ic_y_off + ic_s + 0.30
    div_y_off = lbl_y_off + lbl_h + 0.10
    txt_y_off = div_y_off + 0.14
    txt_h = CARD_H - txt_y_off - 0.42

    for i, (icon, lbl, val) in enumerate(items[:2]):
        x = CX + i * (cw + gap)
        content_card_premium(slide, x, card_y, cw, CARD_H, T, radius=14, depth=True)

        # Top accent strip
        tp = rrect(slide, x, card_y, cw, 0.36, T.accent_rgb, radius_pct=0)
        if tp:
            multi_stop_gradient(tp, [(0, T.accent), (50, T.accent2), (100, T.accent)], 0)
            glow(tp, T.accent.lstrip('#'), radius=10, alpha=0.28)

        # Icon circle
        icon_circle(slide, x + cw/2 - ic_s/2, card_y + ic_y_off, ic_s,
                    T.accent_grad1, T.accent_grad2, icon, max(15, int(ic_s * 11)), T)

        # Label
        txt(slide, lbl, x + 0.22, card_y + lbl_y_off, cw - 0.44, lbl_h,
            font=_FONT, size=SZ_SECTION_LABEL, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        hline(slide, x + cw * 0.15, card_y + div_y_off, cw * 0.7, T.accent_rgb, thickness=0.05)

        # Content — content-aware
        val_fs = _font_for_len(val, 13, 11, 15)
        txt(slide, val, x + 0.32, card_y + txt_y_off, cw - 0.64, txt_h,
            font=_FONT, size=val_fs, bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.38)
    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "📋", "خطة", "البحث", 2, total)
    chapters = req.chapters[:8]; n = len(chapters)
    if not chapters: return slide
    CY = 0.28; CH = H - CY - 0.22
    row_h, gap = _items_h(n, CH, 0.16, 0.92)

    for i, ch in enumerate(chapters):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.16: break
        even = i % 2 == 0

        rw = rrect(slide, CX, y, CW, row_h, T.card_rgb if even else T.bg2_rgb, radius_pct=10)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=7, dist=2, alpha=0.22)

        # Right accent
        accent_bar_v(slide, CX + CW - 0.26, y, row_h, T, width=0.24)

        # Number badge
        nd = min(0.72, row_h * 0.74)
        nx = CX + CW - 1.78; ny = y + (row_h - nd) / 2
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
            shadow(nc, blur=8, dist=2, alpha=0.32)
        txt(slide, str(i + 1), nx, ny, nd, nd,
            font='Calibri', size=max(9, int(nd * 11)), bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # Chapter title — content-aware
        title_fs = _font_for_len(ch.title, 13, 10, 15)
        title_fs = max(title_fs, min(15, int(row_h * 9.5)))
        txt(slide, ch.title, CX + 0.28, y, CW - 2.62, row_h,
            font=_FONT, size=title_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        if ch.pages:
            pb = rrect(slide, CX + 0.22, y + (row_h - 0.36) / 2, 2.0, 0.36, T.bg_rgb, radius_pct=40)
            if pb: set_solid_alpha(pb, 55)
            txt(slide, ch.pages, CX + 0.22, y + (row_h - 0.36) / 2, 2.0, 0.36,
                font='Calibri', size=9, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "❓", "إشكالية", "البحث", 3, total)
    CY = 0.28; CH = H - CY - 0.22

    secs = []
    if req.main_problem:  secs.append('p')
    if req.main_question: secs.append('q')
    if req.sub_questions: secs.append('s')
    if not secs: return slide

    weights = {'p': 2.8, 'q': 1.6, 's': 2.0}
    tw = sum(weights[s] for s in secs)
    gap = 0.22; avail = CH - gap * (len(secs) - 1)
    cy = CY

    if 'p' in secs:
        h = avail * weights['p'] / tw
        pc = rrect(slide, CX, cy, CW, h, T.card_rgb, radius_pct=14)
        if pc:
            multi_stop_gradient(pc, [(0, T.card), (100, T.bg2)], 135)
            shadow(pc, blur=18, dist=5, alpha=0.42)
            glow(pc, T.accent.lstrip('#'), radius=22, alpha=0.08)

        lb = rrect(slide, CX + CW - 6.4, cy, 5.9, 0.52, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb, [(0, T.accent), (100, T.accent2)], 0)
        txt(slide, "◆  الإشكالية الرئيسية", CX + CW - 6.4, cy, 5.9, 0.52,
            font=_FONT, size=11, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        accent_bar_v(slide, CX + CW - 0.26, cy, h, T, width=0.24)

        txt(slide, "❝", CX + 0.32, cy + 0.6, 1.4, 1.2,
            font='Calibri', size=34, bold=False, color=T.accent_rgb,
            align=PP_ALIGN.LEFT, rtl=False)

        prob_fs = _font_for_len(req.main_problem, 13, 10, 14)
        txt(slide, req.main_problem, CX + 1.9, cy + 0.58, CW - 2.4, h - 0.74,
            font=_FONT, size=prob_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.38)
        cy += h + gap

    if 'q' in secs:
        h = avail * weights['q'] / tw
        qc = rrect(slide, CX, cy, CW, h, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=8, dist=2, alpha=0.24)
        vline(slide, CX + CW - 0.22, cy, h, T.accent_rgb, thickness=0.22)
        dot = oval(slide, CX + CW - 2.9, cy + h/2 - 0.2, 0.42, 0.42, T.accent_rgb)
        if dot: multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)
        q_fs = _font_for_len(req.main_question, 13, 11, 14)
        txt(slide, req.main_question, CX + 0.32, cy, CW - 1.65, h,
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
                sb = rrect(slide, CX, y2, CW, sub_h - 0.06, T.card_rgb, radius_pct=6)
                if sb: set_solid_alpha(sb, 50)
            nd = min(0.44, sub_h * 0.56)
            nc = oval(slide, CX + CW - 2.6, y2 + (sub_h - nd) / 2, nd, nd, T.accent_rgb)
            if nc: set_solid_alpha(nc, 65)
            txt(slide, str(i + 1), CX + CW - 2.6, y2 + (sub_h - nd) / 2, nd, nd,
                font='Calibri', size=max(7, int(nd * 9)), bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
            q_fs2 = _font_for_len(q, 12, 10, 13)
            txt(slide, q, CX + 0.32, y2, CW - 3.2, sub_h,
                font=_FONT, size=q_fs2, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "🎯", "أهداف", "وفرضيات", 4, total)
    CY = 0.28; CH = H - CY - 0.22
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
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "⭐", "أهمية", "البحث", 5, total)
    CY = 0.28; CH = H - CY - 0.22
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐", "🔑", "📌", "🌟", "💎", "🏆"]
    cols = 2 if len(items) > 3 else 1
    rows_n = (len(items) + cols - 1) // cols
    gh = 0.22; gv = 0.20
    cw = (CW - gh * (cols - 1)) / cols
    ch = (CH - gv * (rows_n - 1)) / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = CX + ci * (cw + gh); y = CY + ri * (ch + gv)

        content_card_premium(slide, x, y, cw, ch, T, accent_side='right', radius=12)

        ic_s = min(1.35, ch * 0.62)
        icon_circle(slide, x + 0.30, y + (ch - ic_s) / 2,
                    ic_s, T.accent_grad1, T.accent_grad2,
                    icons[i % len(icons)], max(13, int(ic_s * 11)), T)

        item_fs = _font_for_len(item, 13, 11, 14)
        txt(slide, item, x + ic_s + 0.62, y + 0.12, cw - ic_s - 1.18, ch - 0.24,
            font=_FONT, size=item_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.32)
    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "🔬", "منهجية", "البحث", 6, total)
    CY = 0.28; CH = H - CY - 0.22
    icons_map = {"المنهج": "📊", "العينة": "👥", "حجم العينة": "📏", "الأداة": "🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج",     req.methodology))
    if req.sample_type:  fields.append(("العينة",     req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة",     req.tool))
    if not fields: return slide

    cols = 2 if len(fields) > 2 else len(fields)
    rows_n = (len(fields) + cols - 1) // cols
    gh = 0.24; gv = 0.22
    cw = (CW - gh * (cols - 1)) / cols
    ch = (CH - gv * (rows_n - 1)) / rows_n

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i % cols; ri = i // cols
        x = CX + ci * (cw + gh); y = CY + ri * (ch + gv)
        methodology_card(slide, x, y, cw, ch, lbl, val, icons_map.get(lbl, "📌"), T)
    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "📈", "الأرقام", "الرئيسية", 7, total)
    CY = 0.28; CH = H - CY - 0.22
    stats = req.stats[:6]
    if not stats: return slide

    cols = 3 if len(stats) >= 3 else len(stats)
    rows_n = (len(stats) + cols - 1) // cols
    gh = 0.24; gv = 0.20
    cw = (CW - gh * (cols - 1)) / cols
    ch = (CH - gv * (rows_n - 1)) / rows_n

    for i, stat in enumerate(stats):
        ci = i % cols; ri = i // cols
        x = CX + ci * (cw + gh); y = CY + ri * (ch + gv)
        if y + ch > H - 0.2: break
        kpi_card(slide, x, y, cw, ch, stat.value, stat.label, T,
                 unit=stat.unit or '', rank=i)
    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "📊", "نتائج", "البحث", 8, total)
    CY = 0.28; CH = H - CY - 0.22
    results = req.main_results[:8]
    if not results: return slide

    row_h, gap = _items_h(len(results), CH, 0.16, 1.05)
    for i, result in enumerate(results):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.16: break
        result_row_premium(slide, CX, y, CW, row_h, result, i + 1, T, highlight=(i == 0))
    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "💡", "خاتمة", "البحث", 9, total)
    CY = 0.28; CH = H - CY - 0.22

    cc = rrect(slide, CX, CY, CW, CH, T.card_rgb, radius_pct=14)
    if cc:
        multi_stop_gradient(cc, [(0, T.card), (50, T.bg2), (100, T.card)], 135)
        shadow(cc, blur=24, dist=7, alpha=0.46)
        glow(cc, T.accent.lstrip('#'), radius=28, alpha=0.09)

    tp = rrect(slide, CX, CY, CW, 0.32, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp, [(0, T.accent2), (50, T.accent), (100, T.accent2)], 0)
        glow(tp, T.accent.lstrip('#'), radius=12, alpha=0.3)

    accent_bar_v(slide, CX + CW - 0.26, CY + 0.32, CH - 0.58, T, width=0.26)

    diamond(slide, CX + 0.32, CY + 0.50, 1.0, 1.0, T.accent_rgb, alpha=14)
    diamond(slide, CX + CW - 1.5, CY + CH - 1.5, 0.92, 0.92, T.accent_rgb, alpha=10)

    txt(slide, "❝", CX + 0.38, CY + 0.46, 1.7, 1.5,
        font='Calibri', size=48, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.LEFT, rtl=False)

    conc_fs = _font_for_len(req.general_conclusion, 14, 11, 16)
    txt(slide, req.general_conclusion, CX + 0.38, CY + 1.02, CW - 0.96, CH - 2.18,
        font=_FONT, size=conc_fs, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.42)

    ny = CY + CH - 1.05
    hl = rect(slide, CX + CW * 0.18, ny, CW * 0.64, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)
    txt(slide, req.student_name, CX, ny + 0.14, CW, 0.82,
        font=_FONT, size=20, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "✅", "توصيات", "البحث", 10, total)
    CY = 0.28; CH = H - CY - 0.22
    recs = req.recommendations[:8]
    if not recs: return slide

    row_h, gap = _items_h(len(recs), CH, 0.16, 1.0)
    for i, rec in enumerate(recs):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.16: break

        rw = rrect(slide, CX, y, CW, row_h, T.card_rgb, radius_pct=10)
        if rw:
            multi_stop_gradient(rw, [(0, T.card), (100, T.bg2)], 0)
            shadow(rw, blur=7, dist=2, alpha=0.22)

        dot = oval(slide, CX + CW - 1.9, y + (row_h - 0.44) / 2, 0.44, 0.44, T.accent_rgb)
        if dot:
            multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)
            shadow(dot, blur=6, dist=1, alpha=0.28)
        accent_bar_v(slide, CX + CW - 0.26, y, row_h, T, width=0.24)

        rec_fs = _font_for_len(rec, 13, 11, 15)
        txt(slide, rec, CX + 0.26, y, CW - 2.8, row_h,
            font=_FONT, size=rec_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.28)
    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "🔭", "آفاق", "مستقبلية", 11, total)
    CY = 0.28; CH = H - CY - 0.22
    items = req.future_work[:6]
    if not items: return slide

    cols = 2 if len(items) > 3 else 1
    rows_n = (len(items) + cols - 1) // cols
    gh = 0.24; gv = 0.22
    cw = (CW - gh * (cols - 1)) / cols
    ch = (CH - gv * (rows_n - 1)) / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = CX + ci * (cw + gh); y = CY + ri * (ch + gv)

        content_card_premium(slide, x, y, cw, ch, T, radius=14)

        tp = rrect(slide, x, y, cw, 0.26, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp, [(0, T.accent), (100, T.accent2)], 0)

        nd = min(0.88, ch * 0.32)
        number_badge(slide, x + cw/2 - nd/2, y + 0.38, nd, i + 1, T)

        hline(slide, x + cw * 0.18, y + nd + 0.54, cw * 0.64, T.muted_rgb, thickness=0.04)

        item_fs = _font_for_len(item, 13, 11, 14)
        txt(slide, item, x + 0.30, y + nd + 0.70, cw - 0.60, ch - nd - 0.92,
            font=_FONT, size=item_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.32)
    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    total = getattr(req, '_total_slides', 13)
    slide = _section_slide(prs, T, "📚", "المراجع", "والمصادر", 12, total)
    CY = 0.28; CH = H - CY - 0.22
    refs = req.references[:12]
    if not refs: return slide

    row_h, gap = _items_h(len(refs), CH, 0.12, 0.82)
    total_h = len(refs) * (row_h + gap) - gap
    sy = CY + (CH - total_h) / 2

    for i, ref in enumerate(refs):
        y = sy + i * (row_h + gap)
        if y + row_h > H - 0.16: break
        even = i % 2 == 0

        rw = rrect(slide, CX, y, CW, row_h, T.card_rgb if even else T.bg2_rgb, radius_pct=6)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)

        accent_bar_v(slide, CX + CW - 0.26, y, row_h, T, width=0.22, alpha=55)

        nb = rrect(slide, CX + 0.12, y + (row_h - 0.38) / 2, 0.72, 0.38, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 60)
        txt(slide, f"[{i+1}]", CX + 0.12, y + (row_h - 0.38) / 2, 0.72, 0.38,
            font='Calibri', size=9, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        ref_fs = max(10, min(13, row_h * 8))
        txt(slide, ref, CX + 0.96, y + 0.04, CW - 1.46, row_h - 0.08,
            font=_FONT, size=ref_fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=135)

    # Sidebar
    gradient_rect(slide, 0, 0, SW + 0.5, H, T.grad2, T.grad1, angle=185)
    sep = rect(slide, SW + 0.5, 0, 0.14, H, T.accent_rgb)
    if sep: gradient_fill(sep, T.accent_grad1, T.accent_grad2, 90)
    oval(slide, -3.5, -3.5, 10, 10, T.accent_rgb, alpha=7)
    oval(slide, 0.3, H - 6.5, 7.5, 7.5, T.bg2_rgb, alpha=40)
    decorative_dots(slide, 0.45, H * 0.55, 5, 4, 0.13, 0.30, T.accent_rgb, alpha=12)

    txt(slide, "✦", 0, H * 0.26, SW + 0.5, 2.0,
        font='Calibri', size=30, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    if req.institution:
        txt(slide, req.institution, 0.2, H * 0.52, SW + 0.28, 1.4,
            font=_FONT, size=9.5, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.2)

    if req.year:
        yb = rrect(slide, 0.44, H - 1.52, SW - 0.32, 0.66, T.accent_rgb, radius_pct=40)
        if yb: multi_stop_gradient(yb, [(0, T.accent), (100, T.accent2)], 0)
        txt(slide, req.year, 0.44, H - 1.52, SW - 0.32, 0.66,
            font='Calibri', size=13, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Content area
    gradient_rect(slide, SW + 0.64, 0, W - SW - 0.64, H, T.grad1, T.bg2, angle=135)
    decorative_corner(slide, T, 'tr', 5.5)
    diamond(slide, W - 5.2, H - 5.2, 4.2, 4.2, T.accent_rgb, alpha=5)
    decorative_dots(slide, W - 6.5, 1.5, 4, 4, 0.14, 0.34, T.accent_rgb, alpha=9)

    mcx = SW + 1.0; mcw = W - mcx - 0.75
    ccy = H * 0.05; cch = H * 0.90

    # Multi-layer card
    premium_cover_frame(slide, mcx, ccy, mcw, cch, T)

    txt(slide, "✦", mcx + mcw/2 - 0.78, ccy + 0.55, 1.56, 1.32,
        font='Calibri', size=26, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    txt(slide, "شكراً وتقديراً", mcx + 0.72, ccy + 1.08, mcw - 1.44, cch * 0.28,
        font=_FONT, size=SZ_FINAL_TITLE, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.1)

    d1 = rect(slide, mcx + mcw * 0.14, ccy + cch * 0.37, mcw * 0.72, 0.06, T.accent_rgb)
    if d1: multi_stop_gradient(d1, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)
    rect(slide, mcx + mcw * 0.24, ccy + cch * 0.37 + 0.14, mcw * 0.52, 0.03, T.muted_rgb)

    txt(slide, req.student_name, mcx + 0.72, ccy + cch * 0.41, mcw - 1.44, cch * 0.15,
        font=_FONT, size=SZ_FINAL_SUB, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    ts = req.title_ar[:68] + ("..." if len(req.title_ar) > 68 else "")
    txt(slide, ts, mcx + 1.05, ccy + cch * 0.57, mcw - 2.1, cch * 0.2,
        font=_FONT, size=11.5, bold=False, italic=True, color=T.muted_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.28)

    footer = []
    if req.institution: footer.append(req.institution)
    if req.year:        footer.append(req.year)
    if footer:
        fb = rrect(slide, mcx + mcw * 0.1, ccy + cch * 0.82, mcw * 0.8, 0.62,
                   T.bg_rgb, radius_pct=40)
        if fb: set_solid_alpha(fb, 52)
        txt(slide, "  ·  ".join(footer), mcx + 0.8, ccy + cch * 0.82, mcw - 1.6, 0.62,
            font=_FONT, size=SZ_BODY, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    bt = rect(slide, 0, H - 0.26, W, 0.26, T.accent_rgb)
    if bt: multi_stop_gradient(bt, [(0, T.bg), (30, T.accent), (70, T.accent2), (100, T.bg)], 0)
    return slide
