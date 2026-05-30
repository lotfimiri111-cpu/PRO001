"""
Canva Engine v28.3 — Target Design Match
نظام تصميم يطابق الشرائح المستهدفة بدقة
"""
from __future__ import annotations
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, hexagon, decorative_dots,
    card_3d, icon_circle, number_badge, slide_number,
    txt, txt2, blank_slide,
    smart_title, accent_pill, premium_card, card_with_accent_top,
    card_with_accent_side, kpi_card, result_row, premium_header,
    section_divider_line, adaptive_body_size, premium_bg,
    _smart_font_size,
    glass_card, mesh_gradient_bg, cinematic_bg,
    hero_stat, timeline_row, icon_text_row, two_panel_layout,
    highlight_keyword, cinematic_cover, cinematic_final,
    layout_pick,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(n): global _FONT; _FONT = n

HDR_H = 3.0
CY0   = HDR_H + 0.32
def _ch(): return H - CY0 - 0.30

SZ_SLIDE_TITLE   = 30
SZ_SLIDE_SUB     = 13.5
SZ_SECTION_LABEL = 19
SZ_BODY          = 13.5
SZ_BODY_SM       = 11.5
SZ_LABEL         = 13
SZ_VALUE         = 15
SZ_FINAL_TITLE   = 44
SZ_FINAL_SUB     = 24


def _hdr(slide, T, title, sub='', slide_num=None, total_slides=13, req=None):
    if req is not None and hasattr(req, '_total_slides'):
        total_slides = req._total_slides
    premium_header(slide, T, title, sub, slide_num, total_slides, font=_FONT)


def _bg(slide, T, style='a'):
    premium_bg(slide, T, style)


# ══════════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide, T, 'b')

    tp = rect(slide, 0, 0, W, 0.46, T.accent_rgb)
    if tp: multi_stop_gradient(tp, [(0,T.bg),(28,T.accent),(50,T.accent2),(72,T.accent),(100,T.bg)], 0)
    bt = rect(slide, 0, H-0.34, W, 0.34, T.accent_rgb)
    if bt: gradient_fill(bt, T.accent_grad1, T.accent_grad2, 0)

    if req.institution:
        ib = rrect(slide, W/2-10.5, 0.55, 21, 0.72, T.card_rgb, radius_pct=40)
        if ib: set_solid_alpha(ib, 72)
        txt(slide, req.institution, W/2-10.5, 0.55, 21, 0.72,
            font=_FONT, size=11, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    title_y = 1.36; title_h = 7.4
    info_y = title_y + title_h + 0.24
    info_h = max(0.5, H - 0.38 - info_y)
    cx = 1.6; cw = W - 3.2

    outer = rrect(slide, cx - 0.12, title_y - 0.12, cw + 0.24, title_h + 0.24, T.accent_rgb, radius_pct=16)
    if outer: set_solid_alpha(outer, 14)

    mc = rrect(slide, cx, title_y, cw, title_h, T.card_rgb, radius_pct=14)
    if mc:
        multi_stop_gradient(mc, [(0,T.card),(60,T.bg2),(100,T.bg)], 135)
        shadow(mc, blur=28, dist=8, alpha=0.52)

    ct = rrect(slide, cx, title_y, cw, 0.38, T.accent_rgb, radius_pct=0)
    if ct: multi_stop_gradient(ct, [(0,T.accent),(50,T.accent2),(100,T.accent)], 0)

    vline(slide, cx + cw - 0.24, title_y + 0.38, title_h - 0.38, T.accent_rgb, thickness=0.24)

    import re as _re
    _yp = _re.compile(r'\b\d{4}\s*[-–—]\s*\d{4}\b')
    _ym = _yp.search(req.title_ar or '')
    if _ym:
        _year_str = _ym.group(0).strip()
        _title_clean = _yp.sub('', req.title_ar).strip(' —–-،, ')
    elif req.year:
        _year_str = req.year.strip()
        _title_clean = req.title_ar
    else:
        _year_str = ''
        _title_clean = req.title_ar

    has_year = bool(_year_str)
    title_text_h = title_h * (0.56 if has_year else 0.66)
    ts = _smart_font_size(_title_clean, 24, 14, 28, cw - 0.9, title_text_h)

    txt(slide, _title_clean, cx + 0.45, title_y + 0.44, cw - 0.9, title_text_h,
        font=_FONT, size=ts, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.22)

    if req.title_en:
        txt(slide, req.title_en, cx + 0.45, title_y + title_text_h + 0.28, cw - 0.9, title_h * 0.14,
            font="Calibri", size=10.5, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    if has_year:
        yr_y = title_y + title_h * 0.76
        yr_h = title_h * 0.13
        yr_cx = cx + cw * 0.24; yr_cw = cw * 0.52
        yb = rrect(slide, yr_cx, yr_y, yr_cw, yr_h, T.accent_rgb, radius_pct=50)
        if yb: set_solid_alpha(yb, 28)
        hline(slide, yr_cx + yr_cw * 0.08, yr_y, yr_cw * 0.84, T.accent_rgb, thickness=0.035)
        hline(slide, yr_cx + yr_cw * 0.08, yr_y + yr_h, yr_cw * 0.84, T.accent_rgb, thickness=0.035)
        txt(slide, f'( {_year_str} )', yr_cx, yr_y, yr_cw, yr_h,
            font=_FONT, size=13, bold=False, italic=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    hl = rect(slide, cx + cw * 0.1, title_y + title_h * 0.92, cw * 0.8, 0.055, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0,T.bg2),(50,T.accent),(100,T.bg2)], 0)

    ic = rrect(slide, cx, info_y, cw, info_h, T.card_rgb, radius_pct=12)
    if ic:
        multi_stop_gradient(ic, [(0,T.bg2),(100,T.card)], 135)
        shadow(ic, blur=16, dist=4, alpha=0.36)
    vline(slide, cx + cw - 0.2, info_y, info_h, T.accent_rgb, thickness=0.2)

    rows = [("الطالب", req.student_name)]
    if req.supervisor:     rows.append(("المشرف", req.supervisor))
    if req.co_supervisor:  rows.append(("المشرف المساعد", req.co_supervisor))
    if req.specialization: rows.append(("التخصص", req.specialization))

    rh = info_h / max(len(rows), 1)
    for i, (lbl, val) in enumerate(rows):
        y = info_y + i * rh
        rb = rrect(slide, cx + 0.22, y + 0.05, cw - 0.58, rh - 0.1, T.bg_rgb, radius_pct=7)
        if rb: set_solid_alpha(rb, 48)
        lbl_w = 4.4; val_w = cw - lbl_w - 1.0
        lbl_x = cx + cw - lbl_w - 0.32
        val_x = cx + 0.44
        txt(slide, f"{lbl} :", lbl_x, y + 0.05, lbl_w, rh - 0.1,
            font=_FONT, size=max(13, min(15, rh * 8.5)), bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        vline(slide, lbl_x - 0.2, y + rh * 0.14, rh * 0.72, T.muted_rgb, thickness=0.045)
        txt(slide, val, val_x, y + 0.05, val_w, rh - 0.1,
            font=_FONT, size=max(14, min(16, rh * 10)), bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'c')
    _hdr(slide, T, "مقدمة البحث", "نظرة عامة على الدراسة", slide_num=1, req=req)
    CY = CY0; CH = _ch()

    items = []
    if req.intro_overview: items.append(("📖", "نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("🔬", "المنهج المتبع", req.intro_approach))
    if not items: return slide

    n = len(items); gap = 0.4
    col_w = (W - 2.4 - gap * (n - 1)) / n
    CARD_H = min(CH * 0.92, 12.5)
    card_y = CY + (CH - CARD_H) / 2
    ic_s = min(1.9, CARD_H * 0.2)
    ic_y_off = 0.52
    lbl_y_off = ic_y_off + ic_s + 0.30
    div_y_off = lbl_y_off + 0.76 + 0.12
    txt_y_off = div_y_off + 0.08 + 0.14
    txt_h = CARD_H - txt_y_off - 0.44

    for i, (icon, lbl, val) in enumerate(items[:2]):
        x = 1.2 + i * (col_w + gap)
        cc = card_with_accent_top(slide, x, card_y, col_w, CARD_H, T, radius=12, bar_h=0.40)
        oval(slide, x + col_w - 2.8, card_y - 1.0, 3.5, 3.5, T.accent_rgb, alpha=5)

        ic_x = x + col_w / 2 - ic_s / 2
        icon_circle(slide, ic_x, card_y + ic_y_off, ic_s,
                    T.accent_grad1, T.accent_grad2, icon, max(16, int(ic_s * 11.5)), T)

        txt(slide, lbl, x + 0.24, card_y + lbl_y_off, col_w - 0.48, 0.76,
            font=_FONT, size=SZ_SECTION_LABEL, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        section_divider_line(slide, x + col_w * 0.12, card_y + div_y_off, col_w * 0.76, T)

        fs = adaptive_body_size(val, txt_h, base=13.5, min_s=11, max_s=15.5)
        txt(slide, val, x + 0.30, card_y + txt_y_off, col_w - 0.60, txt_h,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.35)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'a')
    _hdr(slide, T, "خطة البحث", f"يتضمن البحث {len(req.chapters)} فصول رئيسية", slide_num=2, req=req)
    CY = CY0; CH = _ch()
    chapters = req.chapters[:8]; n = len(chapters)
    if not chapters: return slide

    gap = 0.18
    row_h = (CH - gap * (n - 1)) / n

    for i, ch in enumerate(chapters):
        y = CY + i * (row_h + gap)
        even = i % 2 == 0

        rw = rrect(slide, 0.9, y, W - 1.8, row_h, T.card_rgb if even else T.bg2_rgb, radius_pct=10)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=6, dist=2, alpha=0.18)

        acc = rect(slide, W - 1.2, y, 0.24, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        nd = min(0.74, row_h * 0.74)
        nx = W - 2.55; ny = y + (row_h - nd) / 2
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
            shadow(nc, blur=7, dist=2, alpha=0.30)
        txt(slide, str(i + 1), nx, ny, nd, nd,
            font="Calibri", size=max(9, int(nd * 11.5)), bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        fs = _smart_font_size(ch.title, 14, 11, 16, W - 4.8, row_h)
        txt(slide, ch.title, 1.2, y, W - 4.7, row_h,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.15)

        if ch.pages:
            pb = rrect(slide, 1.06, y + (row_h - 0.36) / 2, 2.1, 0.36, T.bg_rgb, radius_pct=40)
            if pb: set_solid_alpha(pb, 55)
            txt(slide, ch.pages, 1.06, y + (row_h - 0.36) / 2, 2.1, 0.36,
                font="Calibri", size=9, bold=False, color=T.muted_rgb,
                align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# PROBLEM — TARGET: 3 zones stacked (main problem large, main question, sub-questions list)
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'b')
    _hdr(slide, T, "إشكالية البحث", "التساؤلات الرئيسية والفرعية", slide_num=3, req=req)
    CY = CY0; CH = _ch()

    secs = []; weights = {}
    if req.main_problem:  secs.append('p'); weights['p'] = 2.6
    if req.main_question: secs.append('q'); weights['q'] = 1.6
    if req.sub_questions: secs.append('s'); weights['s'] = 2.0
    if not secs: return slide

    tw = sum(weights[s] for s in secs)
    gap = 0.24; avail = CH - gap * (len(secs) - 1)
    cy = CY

    if 'p' in secs:
        h = avail * weights['p'] / tw
        # Large oval background card
        pc = rrect(slide, 0.9, cy, W - 1.8, h, T.card_rgb, radius_pct=13)
        if pc:
            multi_stop_gradient(pc, [(0, T.card), (100, T.bg2)], 135)
            shadow(pc, blur=20, dist=6, alpha=0.44)

        # Decorative large oval
        oval(slide, W * 0.3, cy - h * 0.2, W * 0.5, h * 1.1, T.accent_rgb, alpha=4)

        # Label badge top-right
        lb = rrect(slide, W - 7.8, cy, 6.0, 0.54, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb, [(0, T.accent), (100, T.accent2)], 0)
        txt(slide, "◆  الإشكالية الرئيسية", W - 7.8, cy, 6.0, 0.54,
            font=_FONT, size=SZ_LABEL, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        # Quote mark
        txt(slide, "❝", 1.2, cy + 0.62, 1.5, 1.2,
            font="Calibri", size=34, bold=False, color=T.accent_rgb,
            align=PP_ALIGN.LEFT, rtl=False, vcenter=False)

        fs = adaptive_body_size(req.main_problem, h - 0.85, base=13, min_s=11, max_s=15)
        txt(slide, req.main_problem, 2.9, cy + 0.65, W - 4.4, h - 0.82,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.38)
        cy += h + gap

    if 'q' in secs:
        h = avail * weights['q'] / tw
        # Main question — bold italic, large oval bg
        qc = rrect(slide, 0.9, cy, W - 1.8, h, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=9, dist=2, alpha=0.24)
        vline(slide, W - 1.35, cy, h, T.accent_rgb, thickness=0.24)
        fs = adaptive_body_size(req.main_question, h, base=13.5, min_s=11, max_s=15)
        txt(slide, req.main_question, 1.2, cy, W - 2.8, h,
            font=_FONT, size=fs, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.22)
        cy += h + gap

    if 's' in secs and req.sub_questions:
        h = avail * weights['s'] / tw
        subs = req.sub_questions[:4]
        sub_h = h / len(subs)
        for i, q in enumerate(subs):
            y2 = cy + i * sub_h
            if i % 2 == 0:
                sb = rrect(slide, 0.9, y2, W - 1.8, sub_h - 0.06, T.card_rgb, radius_pct=6)
                if sb: set_solid_alpha(sb, 48)
            nd2 = min(0.46, sub_h * 0.56)
            nc2 = oval(slide, W - 2.9, y2 + (sub_h - nd2) / 2, nd2, nd2, T.accent_rgb)
            if nc2: set_solid_alpha(nc2, 65)
            txt(slide, str(i + 1), W - 2.9, y2 + (sub_h - nd2) / 2, nd2, nd2,
                font="Calibri", size=max(7, int(nd2 * 9)), bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
            fs = _smart_font_size(q, 12, 10, 14, W - 4.2, sub_h)
            txt(slide, q, 1.2, y2, W - 3.8, sub_h,
                font=_FONT, size=fs, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.RIGHT,
                rtl=True, vcenter=True, line_spacing=1.15)

    return slide


# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES — TARGET: 2 columns, rows with number badges, no card headers
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'c')
    _hdr(slide, T, "أهداف البحث وفرضياته", "", slide_num=4, req=req)
    CY = CY0; CH = _ch()
    cols = []
    if req.objectives: cols.append(("🎯  الأهداف", req.objectives))
    if req.hypotheses:  cols.append(("💡  الفرضيات", req.hypotheses))
    if not cols: return slide

    n_c = len(cols); gap = 0.20
    col_w = (W - 2.0 - gap * (n_c - 1)) / n_c

    for i, (lbl, items) in enumerate(cols):
        x = 1.0 + i * (col_w + gap)

        # Column header — full-width gold bar
        hh = 0.72
        hdr = rrect(slide, x, CY, col_w, hh, T.accent_rgb, radius_pct=0)
        if hdr: multi_stop_gradient(hdr, [(0, T.accent2), (100, T.accent)], 0)
        txt(slide, lbl, x + 0.2, CY, col_w - 0.4, hh,
            font=_FONT, size=SZ_SECTION_LABEL, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        # Items as rows — each column sizes rows by its own item count
        ia = CH - hh - 0.10; n_items = min(len(items), 8); ig = 0.16
        ih = (ia - ig * (n_items - 1)) / n_items

        for j, item in enumerate(items[:8]):
            iy = CY + hh + 0.05 + j * (ih + ig)
            # Row background — large oval/rrect
            rb = rrect(slide, x, iy, col_w, ih, T.card_rgb if j % 2 == 0 else T.bg2_rgb, radius_pct=30)
            if rb:
                stops = [(0, T.card), (100, T.bg2)] if j % 2 == 0 else [(0, T.bg2), (100, T.card)]
                multi_stop_gradient(rb, stops, 0)
                shadow(rb, blur=4, dist=1, alpha=0.15)

            # Number badge — right side
            nd = min(0.56, ih * 0.66)
            nx = x + col_w - nd - 0.22
            ny = iy + (ih - nd) / 2
            nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
            if nc: multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
            txt(slide, str(j + 1), nx, ny, nd, nd,
                font="Calibri", size=max(10, int(nd * 11)), bold=True,
                color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

            fs = _smart_font_size(item, 13.5, 10.5, 15, col_w - nd - 0.66, ih)
            txt(slide, item, x + 0.24, iy, col_w - nd - 0.60, ih,
                font=_FONT, size=fs, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
                rtl=True, vcenter=True, line_spacing=1.2)

    return slide


# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE — TARGET: 2×2 large oval cards with icons
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'b')
    _hdr(slide, T, "أهمية البحث", "الأثر العلمي والعملي للدراسة", slide_num=5, req=req)
    CY = CY0; CH = _ch()
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐", "🔑", "📌", "🌟", "💎", "🏆"]
    cols = 2 if len(items) > 3 else 1
    rows_n = (len(items) + cols - 1) // cols

    # Gold cross separator lines — drawn on top of cards
    SEP_X = W / 2
    SEP_THICK = 0.14

    gv = 0.0; gh = 0.0
    card_w = (W - 1.8) / cols
    card_h = CH / rows_n

    for i, item in enumerate(items):
        ci = i % cols; ri = i // cols
        x = 0.9 + ci * card_w
        y = CY + ri * card_h

        # Large oval pill filling cell (with small inset)
        pad_x = 0.22; pad_y = 0.14
        ov = rrect(slide, x + pad_x, y + pad_y,
                   card_w - pad_x * 2, card_h - pad_y * 2,
                   T.bg2_rgb, radius_pct=38)
        if ov:
            multi_stop_gradient(ov, [(0, T.card), (100, T.bg2)], 135)
            shadow(ov, blur=14, dist=4, alpha=0.32)

        # Icon circle — always LEFT side of cell
        ic_s = min(1.6, card_h * 0.60)
        ic_x = x + pad_x + 0.30
        ic_y = y + (card_h - ic_s) / 2
        icon_circle(slide, ic_x, ic_y, ic_s,
                    T.accent_grad1, T.accent_grad2, icons[i % len(icons)],
                    max(14, int(ic_s * 11.5)), T)

        # Text — right of icon, aligned right (RTL)
        txt_x = ic_x + ic_s + 0.26
        txt_w = (x + card_w - pad_x - 0.20) - txt_x
        fs = adaptive_body_size(item, card_h - 0.32, base=13.5, min_s=11, max_s=15)
        txt(slide, item, txt_x, y + 0.18, txt_w, card_h - 0.36,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.35)

    # Draw separators ON TOP of cards
    if cols == 2:
        vline(slide, SEP_X, CY, CH, T.accent_rgb, thickness=SEP_THICK)
    for ri in range(1, rows_n):
        y_sep = CY + ri * card_h
        hline(slide, 0.9, y_sep, W - 1.8, T.accent_rgb, thickness=SEP_THICK)

    return slide


# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY — TARGET: 2×2 large oval cards with centered icons + gold top bar
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'd')
    _hdr(slide, T, "منهجية البحث", "الإجراءات والأدوات المستخدمة", slide_num=6, req=req)
    CY = CY0; CH = _ch()

    icons_map = {"المنهج": "📊", "العينة": "👥", "حجم العينة": "📏", "الأداة": "🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج", req.methodology))
    if req.sample_type:  fields.append(("العينة", req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة", req.tool))
    if not fields: return slide

    cols = 2 if len(fields) > 2 else len(fields)
    rows_n = (len(fields) + cols - 1) // cols

    # Gold separator lines
    if cols == 2:
        vline(slide, W / 2, CY, CH, T.accent_rgb, thickness=0.14)
    for ri in range(1, rows_n):
        y_sep = CY + ri * CH / rows_n
        hline(slide, 0.9, y_sep - 0.07, W - 1.8, T.accent_rgb, thickness=0.14)

    gh = 0.14; gv = 0.14
    col_w = (W - 2.0 - gh * (cols - 1)) / cols
    card_h = (CH - gv * (rows_n - 1)) / rows_n

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i % cols; ri = i // cols
        x = 1.0 + ci * (col_w + gh)
        y = CY + ri * (card_h + gv)

        # Large oval background
        ov = oval(slide, x - 0.2, y + card_h * 0.08, col_w + 0.4, card_h * 0.84, T.bg2_rgb, alpha=75)

        # Icon circle — centered top
        ic_s = min(2.0, card_h * 0.38)
        ic_x = x + col_w / 2 - ic_s / 2
        ic_y = y + 0.38
        ic_bg = oval(slide, ic_x, ic_y, ic_s, ic_s, T.accent_rgb)
        if ic_bg:
            multi_stop_gradient(ic_bg, [(0, T.accent), (100, T.accent2)], 135)
            shadow(ic_bg, blur=10, dist=3, alpha=0.30)
        txt(slide, icons_map.get(lbl, "📌"), ic_x, ic_y + ic_s * 0.07, ic_s, ic_s * 0.86,
            font="Calibri", size=max(14, int(ic_s * 10.5)), bold=False,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # Label underlined
        lbl_y = y + ic_s + 0.56
        txt(slide, lbl, x + 0.2, lbl_y, col_w - 0.4, 0.68,
            font=_FONT, size=SZ_SECTION_LABEL, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        section_divider_line(slide, x + col_w * 0.14, lbl_y + 0.70, col_w * 0.72, T)

        remain = card_h - (lbl_y - y) - 0.94
        fs = adaptive_body_size(val, remain, base=13, min_s=10.5, max_s=14.5)
        txt(slide, val, x + 0.22, lbl_y + 0.82, col_w - 0.44, remain,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER,
            rtl=True, vcenter=True, line_spacing=1.28)

    return slide


# ══════════════════════════════════════════════════════════════════════
# STATS — Hero layout
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    mesh_gradient_bg(slide, T, complexity=3)
    _hdr(slide, T, "الأرقام والإحصاءات الرئيسية", "", slide_num=7, req=req)
    CY = CY0; CH = _ch()
    stats = req.stats[:6]
    if not stats: return slide

    n = len(stats)
    if n == 1:
        hero_stat(slide, 2.0, CY+CH*0.06, W-4.0, CH*0.88, T,
                  value=stats[0].value, label=stats[0].label,
                  unit=getattr(stats[0],'unit',''), font=_FONT, rank=0)
    elif n == 2:
        w2 = (W-3.0)/2; gh = 0.30
        for i,s in enumerate(stats):
            x = 1.2 + i*(w2+gh)
            hero_stat(slide, x, CY+CH*0.06, w2, CH*0.88, T,
                      value=s.value, label=s.label,
                      unit=getattr(s,'unit',''), font=_FONT, rank=i)
    elif n == 3:
        hw = W*0.40; sw2 = (W-hw-2.8)/2
        hx = W-hw-0.8
        hero_stat(slide, hx, CY+CH*0.04, hw, CH*0.92, T,
                  value=stats[0].value, label=stats[0].label,
                  unit=getattr(stats[0],'unit',''), font=_FONT, rank=0)
        for i,s in enumerate(stats[1:3]):
            sx = 0.8+i*(sw2+0.24)
            sy = CY+CH*0.04+i*(CH*0.92/2+0.15)
            sh2 = CH*0.92/2-0.08
            hero_stat(slide, sx, sy, sw2, sh2, T,
                      value=s.value, label=s.label,
                      unit=getattr(s,'unit',''), font=_FONT, rank=i+1)
    else:
        cols = 3 if n >= 3 else n
        rows_n = (n+cols-1)//cols
        gh = 0.28; gv = 0.22
        cw = (W-2.0-gh*(cols-1))/cols
        ch2 = (CH-gv*(rows_n-1))/rows_n
        for i,s in enumerate(stats):
            ci = i%cols; ri = i//cols
            x = 1.0+ci*(cw+gh); y = CY+ri*(ch2+gv)
            hero_stat(slide, x, y, cw, ch2, T,
                      value=s.value, label=s.label,
                      unit=getattr(s,'unit',''), font=_FONT, rank=i)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RESULTS — Icon-text rows with highlight on first 2
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    mesh_gradient_bg(slide, T, complexity=2)
    _hdr(slide, T, "نتائج البحث", "أبرز ما توصلت إليه الدراسة", slide_num=8, req=req)
    CY = CY0; CH = _ch()
    results = req.main_results[:8]; n = len(results)
    if not results: return slide

    icons = ["🏆","⭐","📊","🔬","💡","🎯","📌","✅"]
    gap = 0.16
    row_h = (CH-gap*(n-1))/n

    for i,result in enumerate(results):
        y = CY+i*(row_h+gap)
        highlight = (i < 2 and n >= 3)
        icon_text_row(slide, 0.9, y, W-1.8, row_h, T,
                      result, icons[i%len(icons)], i,
                      font=_FONT, highlight=highlight)

    return slide


# ══════════════════════════════════════════════════════════════════════
# CONCLUSION — TARGET: Large oval with quote mark, name at bottom
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'd')
    _hdr(slide, T, "خاتمة البحث", "الاستنتاج العام للدراسة", slide_num=9, req=req)
    CY = CY0; CH = _ch()

    # Gold top accent bar
    gold_bar = rect(slide, 0.9, CY, W - 1.8, 0.18, T.accent_rgb)
    if gold_bar: multi_stop_gradient(gold_bar, [(0, T.bg),(40, T.accent),(60, T.accent2),(100, T.bg)], 0)

    # Large oval card
    ov_pad = 0.3
    ov = oval(slide, 0.9, CY + 0.22, W - 1.8, CH - 1.2, T.bg2_rgb, alpha=88)

    # Quote mark
    txt(slide, "❝", 1.6, CY + 0.60, 1.5, 1.4,
        font="Calibri", size=42, bold=False, color=T.accent_rgb,
        align=PP_ALIGN.LEFT, rtl=False, vcenter=False)

    fs = adaptive_body_size(req.general_conclusion, CH - 2.5, base=14, min_s=12, max_s=17)
    txt(slide, req.general_conclusion, 1.8, CY + 1.10, W - 3.2, CH - 2.45,
        font=_FONT, size=fs, bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
        rtl=True, vcenter=True, line_spacing=1.42)

    # Name at bottom — inside oval, centered
    ny = CY + CH - 1.05
    hl = rect(slide, W * 0.15, ny, W * 0.70, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl, [(0, T.bg2),(50, T.accent),(100, T.bg2)], 0)
    txt(slide, req.student_name, 1.0, ny + 0.12, W - 2.0, 0.86,
        font=_FONT, size=22, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS — TARGET: Rows with bullet dot right, gold right bar
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'b')
    _hdr(slide, T, "توصيات البحث", "", slide_num=10, req=req)
    CY = CY0; CH = _ch()
    recs = req.recommendations[:8]; n = len(recs)
    if not recs: return slide

    gap = 0.18
    row_h = (CH - gap * (n - 1)) / n

    for i, rec in enumerate(recs):
        y = CY + i * (row_h + gap)

        # Row background — large rounded pill
        rw = rrect(slide, 0.9, y, W - 1.8, row_h, T.card_rgb, radius_pct=30)
        if rw:
            multi_stop_gradient(rw, [(0, T.card), (100, T.bg2)], 0)
            shadow(rw, blur=6, dist=2, alpha=0.22)

        # Gold right bar
        acc = rect(slide, W - 1.2, y, 0.24, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        # Bullet dot — gold circle (no number)
        dot_s = min(0.42, row_h * 0.55)
        dot_x = W - 2.55
        dot_y = y + (row_h - dot_s) / 2
        dot = oval(slide, dot_x, dot_y, dot_s, dot_s, T.accent_rgb)
        if dot: multi_stop_gradient(dot, [(0, T.accent), (100, T.accent2)], 135)

        fs = _smart_font_size(rec, 13.5, 11, 15.5, W - 3.6, row_h)
        txt(slide, rec, 1.2, y, W - 3.55, row_h,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.25)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FUTURE — TARGET: Vertical rows with number badge top-center + gold bar + text
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'a')
    _hdr(slide, T, "آفاق البحث المستقبلية", "", slide_num=11, req=req)
    CY = CY0; CH = _ch()
    items = req.future_work[:6]
    if not items: return slide

    n = len(items)
    gap = 0.16
    row_h = (CH - gap * (n - 1)) / n

    for i, item in enumerate(items):
        y = CY + i * (row_h + gap)

        # Row card — large rrect pill
        cc = rrect(slide, 0.9, y, W - 1.8, row_h, T.card_rgb, radius_pct=28)
        if cc:
            multi_stop_gradient(cc, [(0, T.card), (70, T.bg2), (100, T.bg)], 155)
            shadow(cc, blur=8, dist=2, alpha=0.28)

        # Gold top bar (full width of card)
        tp = rrect(slide, 0.9, y, W - 1.8, 0.18, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp, [(0, T.bg),(35, T.accent),(65, T.accent2),(100, T.bg)], 0)

        # Number badge — centered at top of card
        nd = min(0.88, row_h * 0.56)
        nx = W / 2 - nd / 2
        ny = y + 0.08
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
            shadow(nc, blur=7, dist=2, alpha=0.30)
        txt(slide, str(i + 1), nx, ny, nd, nd,
            font="Calibri", size=max(11, int(nd * 11)), bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # Divider below number
        section_divider_line(slide, W * 0.15, y + nd + 0.18, W * 0.70, T)

        # Text below
        fs = adaptive_body_size(item, row_h - nd - 0.52, base=13.5, min_s=11, max_s=15)
        txt(slide, item, 1.6, y + nd + 0.32, W - 3.2, row_h - nd - 0.50,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.CENTER,
            rtl=True, vcenter=True, line_spacing=1.32)

    return slide


# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide, T, 'c')
    _hdr(slide, T, "قائمة المراجع والمصادر", "", slide_num=12, req=req)
    CY = CY0; CH = _ch()
    refs = req.references[:12]; n = len(refs)
    if not refs: return slide

    gap = 0.13
    row_h = (CH - gap * (n - 1)) / n

    for i, ref in enumerate(refs):
        y = CY + i * (row_h + gap)
        if y + row_h > H - 0.2: break
        even = i % 2 == 0
        rw = rrect(slide, 0.9, y, W - 1.8, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=6)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)

        acc = rect(slide, W - 1.28, y, 0.24, row_h, T.accent_rgb)
        if acc: set_solid_alpha(acc, 55)

        nb = rrect(slide, 1.02, y + (row_h - 0.40) / 2, 0.74, 0.40, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 65)
        txt(slide, f"[{i+1}]", 1.02, y + (row_h - 0.40) / 2, 0.74, 0.40,
            font="Calibri", size=9, bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        fs = _smart_font_size(ref, 13, 10.5, 14.5, W - 3.5, row_h)
        txt(slide, ref, 1.88, y + 0.05, W - 3.38, row_h - 0.1,
            font=_FONT, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.15)

    return slide


# ══════════════════════════════════════════════════════════════════════
# FINAL
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    cinematic_final(slide, T, req, font=_FONT)
    return slide
