"""
Premium Engine v17.3 — شريط جانبي + محتوى يملأ الشريحة 100%
"""
from __future__ import annotations
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, decorative_dots,
    number_badge, icon_circle, slide_number, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest
from engine.ultra_core import (
    ultra_bg, rtl_bullets, hero_stats, cinematic_cover, cinematic_thankyou,
    patch_header_rtl, trunc, smart_size,
)

_FONT = "Cairo"
def set_font(n): global _FONT; _FONT = n

SW = 5.2   # عرض الشريط الجانبي — ثابت لكل الشرائح
CX = SW + 0.55  # بداية منطقة المحتوى
CW = W - CX - 0.55  # عرض منطقة المحتوى

# ─── شريط جانبي ثابت لكل الشرائح ─────────────────────────────────────
def _section_slide(prs, T, icon, label1, label2=""):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _sidebar(slide, T, icon, label1, label2)
    return slide


def _sidebar(slide, T, icon, label1, label2=""):
    # خلفية الشريط بتدرج
    sb = gradient_rect(slide, 0, 0, SW, H, T.grad2, T.grad1, angle=180)
    # خط فاصل
    sep = rect(slide, SW, 0, 0.12, H, T.accent_rgb)
    if sep: gradient_fill(sep, T.accent_grad1, T.accent_grad2, 90)
    # زخارف الشريط
    oval(slide, -3, -3, 9, 9, T.accent_rgb, alpha=6)
    oval(slide, 0.5, H-5.5, 6.5, 6.5, T.bg2_rgb, alpha=45)
    decorative_dots(slide, 0.4, H*0.55, 5, 4, 0.14, 0.32, T.accent_rgb, alpha=12)

    # دائرة الأيقونة في المنتصف العلوي من الشريط
    ic_y = H*0.18
    ic_bg = oval(slide, SW/2-1.5, ic_y, 3.0, 3.0, T.accent_rgb)
    if ic_bg:
        multi_stop_gradient(ic_bg,[(0,T.accent),(100,T.accent2)],135)
        shadow(ic_bg, blur=14, dist=4, alpha=0.4)
        glow(ic_bg, T.accent.lstrip('#'), radius=20, alpha=0.2)
    txt(slide, icon, SW/2-1.5, ic_y+0.3, 3.0, 2.4,
        font="Calibri", size=34, bold=False,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # عنوان القسم في الشريط
    label_y = ic_y + 3.2
    txt(slide, label1, 0.2, label_y, SW-0.4, 1.1,
        font=_FONT, size=16, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)
    if label2:
        txt(slide, label2, 0.2, label_y+1.15, SW-0.4, 0.9,
            font=_FONT, size=12, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # خط فاصل في الشريط
    hl = rect(slide, 0.5, label_y+2.2, SW-1.0, 0.05, T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg),(50,T.accent),(100,T.bg)],0)

    # خلفية منطقة المحتوى
    rect(slide, SW+0.12, 0, W-SW-0.12, H, T.bg2_rgb)
    # تدرج خفيف على منطقة المحتوى
    gr = gradient_rect(slide, SW+0.12, 0, W-SW-0.12, H, T.grad1, T.bg2, angle=135)
    # زخارف المحتوى
    oval(slide, W-8, -3, 11, 11, T.accent_rgb, alpha=4)
    diamond(slide, W-5, H-5, 4, 4, T.accent_rgb, alpha=5)

# ─── COVER ─────────────────────────────────────────────────────────────

def make_cover(prs, req: PresentationRequest, T: Theme):
    return cinematic_cover(prs, req, T, _FONT)

def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📋", "خطة", "البحث")
    chapters = req.chapters[:8]
    n = len(chapters)
    if not chapters: return slide

    # يملأ الشريحة من أعلى إلى أسفل
    CY = 0.28; CH = H - CY - 0.22
    gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, ch in enumerate(chapters):
        y = CY + i*(row_h+gap)
        even = i%2==0

        # خلفية الصف بالكامل
        rw = rect(slide, CX, y, CW, row_h,
                  T.card_rgb if even else T.bg2_rgb)
        if rw:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw, stops, 0)

        # شريط acc يميني
        acc = rect(slide, CX+CW-0.22, y, 0.22, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        # دائرة رقم
        nd = min(0.68, row_h*0.72)
        nx = CX+CW-1.7; ny = y+(row_h-nd)/2
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc,[(0,T.accent),(100,T.accent2)],135)
            shadow(nc, blur=6, dist=2, alpha=0.28)
        txt(slide, str(i+1), nx, ny, nd, nd,
            font="Calibri", size=max(9, int(nd*10)),
            bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

        # عنوان الفصل
        txt(slide, ch.title, CX+0.25, y, CW-2.45, row_h,
            font=_FONT, size=max(11, min(15, int(row_h*9))),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        # الصفحات
        if ch.pages:
            pb = rrect(slide, CX+0.22, y+(row_h-0.34)/2, 1.85, 0.34,
                       T.bg_rgb, radius_pct=40)
            if pb: set_solid_alpha(pb, 52)
            txt(slide, ch.pages, CX+0.22, y+(row_h-0.34)/2, 1.85, 0.34,
                font="Calibri", size=8.5, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    slide_number(slide, 2, getattr(req, "_total_slides", 13), T)
    return slide


# ─── PROBLEM ───────────────────────────────────────────────────────────
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "❓", "إشكالية", "البحث")
    CY = 0.35; CH = H - CY - 0.3

    sections = []
    if req.main_problem: sections.append('p')
    if req.main_question: sections.append('q')
    if req.sub_questions: sections.append('s')

    weights = {'p':2.8,'q':1.6,'s':2.0}
    total_w = sum(weights[s] for s in sections)
    gap = 0.22
    avail = CH - gap*(len(sections)-1)

    cy = CY
    if 'p' in sections:
        h = avail * weights['p']/total_w
        pc = rrect(slide, CX, cy, CW, h, T.card_rgb, radius_pct=12)
        if pc:
            multi_stop_gradient(pc,[(0,T.card),(100,T.bg2)],135)
            shadow(pc, blur=16, dist=5, alpha=0.4)
            glow(pc, T.accent.lstrip('#'), radius=20, alpha=0.09)
        lb = rrect(slide, CX+CW-6.2, cy, 5.8, 0.5, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb,[(0,T.accent),(100,T.accent2)],0)
        txt(slide, "◆  الإشكالية الرئيسية", CX+CW-6.2, cy, 5.8, 0.5,
            font=_FONT, size=11, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, "❝", CX+0.3, cy+0.58, 1.3, 1.0,
            font="Calibri", size=28, bold=False,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)
        txt(slide, req.main_problem, CX+1.75, cy+0.55, CW-2.2, h-0.72,
            font=_FONT, size=12.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        cy += h+gap

    if 'q' in sections:
        h = avail * weights['q']/total_w
        qc = rrect(slide, CX, cy, CW, h, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=8, dist=2, alpha=0.25)
        vline(slide, CX+CW-0.2, cy, h, T.accent_rgb, thickness=0.2)
        dot = oval(slide, CX+CW-2.8, cy+h/2-0.18, 0.38, 0.38, T.accent_rgb)
        if dot: multi_stop_gradient(dot,[(0,T.accent),(100,T.accent2)],135)
        txt(slide, req.main_question, CX+0.3, cy, CW-1.5, h,
            font=_FONT, size=12.5, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        cy += h+gap

    if 's' in sections and req.sub_questions:
        h = avail * weights['s']/total_w
        subs = req.sub_questions[:5]
        sub_h = h / len(subs)
        for i, q in enumerate(subs):
            y2 = cy+i*sub_h
            if i%2==0:
                sb = rrect(slide, CX, y2, CW, sub_h-0.05, T.card_rgb, radius_pct=5)
                if sb: set_solid_alpha(sb, 48)
            nc = oval(slide, CX+CW-2.5, y2+(sub_h-0.33)/2, 0.33, 0.33, T.accent_rgb)
            if nc: set_solid_alpha(nc, 65)
            txt(slide, str(i+1), CX+CW-2.5, y2+(sub_h-0.33)/2, 0.33, 0.33,
                font="Calibri", size=7.5, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
            txt(slide, q, CX+0.3, y2, CW-3.1, sub_h,
                font=_FONT, size=max(10, min(12, sub_h*8)),
                bold=False, color=T.muted_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 3, getattr(req, "_total_slides", 13), T)
    return slide

# ─── OBJECTIVES ────────────────────────────────────────────────────────
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "🎯", "أهداف", "وفرضيات")
    CY = 0.35; CH = H - CY - 0.3
    cols = []
    if req.objectives: cols.append(("🎯  الأهداف", req.objectives))
    if req.hypotheses:  cols.append(("💡  الفرضيات", req.hypotheses))
    if not cols: return slide

    n_c = len(cols)
    gap = 0.25
    col_w = (CW - gap*(n_c-1)) / n_c

    for i, (lbl, items) in enumerate(cols):
        x = CX + i*(col_w+gap)
        cc = rrect(slide, x, CY, col_w, CH, T.card_rgb, radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],150)
            shadow(cc, blur=14, dist=4, alpha=0.35)
        hh = 0.68
        hdr = rrect(slide, x, CY, col_w, hh, T.accent_rgb, radius_pct=0)
        if hdr: multi_stop_gradient(hdr,[(0,T.accent2),(100,T.accent)],0)
        txt(slide, lbl, x+0.18, CY, col_w-0.36, hh,
            font=_FONT, size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        ia = CH - hh - 0.12
        n_items = min(len(items), 8)
        ig = 0.1
        ih = (ia - ig*(n_items-1)) / n_items

        for j, item in enumerate(items[:8]):
            iy = CY + hh + 0.06 + j*(ih+ig)
            rb = rrect(slide, x+0.1, iy, col_w-0.2, ih,
                       T.bg2_rgb if j%2==0 else T.bg_rgb, radius_pct=7)
            if rb: set_solid_alpha(rb, 72)
            number_badge(slide, x+col_w-0.78, iy+(ih-0.5)/2, 0.5, j+1, T)
            txt(slide, item, x+0.22, iy+0.05, col_w-1.2, ih-0.1,
                font=_FONT, size=max(9, min(11.5, ih*7.5)),
                bold=False, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 4, getattr(req, "_total_slides", 13), T)
    return slide

# ─── IMPORTANCE ────────────────────────────────────────────────────────
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "⭐", "أهمية", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐","🔑","📌","🌟","💎","🏆"]
    cols = 2 if len(items)>3 else 1
    rows_n = (len(items)+cols-1)//cols
    gh = 0.18; gv = 0.18
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],145)
            shadow(cc, blur=12, dist=3, alpha=0.32)
        acc = rect(slide, x+cw-0.25, y, 0.25, ch, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        ic_s = min(1.3, ch*0.6)
        icon_circle(slide, x+0.25, y+(ch-ic_s)/2, ic_s,
                    T.accent_grad1, T.accent_grad2,
                    icons[i%len(icons)], max(13, int(ic_s*11)), T)
        txt(slide, item, x+ic_s+0.5, y+0.1, cw-ic_s-1.0, ch-0.2,
            font=_FONT, size=max(9.5, min(12.5, ch*7)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    slide_number(slide, 5, getattr(req, "_total_slides", 13), T)
    return slide

# ─── METHODOLOGY ───────────────────────────────────────────────────────
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "🔬", "منهجية", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    icons_map = {"المنهج":"📊","العينة":"👥","حجم العينة":"📏","الأداة":"🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج", req.methodology))
    if req.sample_type:  fields.append(("العينة", req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة", req.tool))
    if not fields: return slide

    cols = 2 if len(fields)>2 else len(fields)
    rows_n = (len(fields)+cols-1)//cols
    gh = 0.22; gv = 0.2
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],145)
            shadow(cc, blur=13, dist=4, alpha=0.38)

        ic_s = min(1.8, ch*0.36)
        ic_x = x+cw/2-ic_s/2
        ic_bg = oval(slide, ic_x, y+0.28, ic_s, ic_s, T.accent_rgb)
        if ic_bg:
            multi_stop_gradient(ic_bg,[(0,T.accent),(100,T.accent2)],135)
            shadow(ic_bg, blur=9, dist=3, alpha=0.3)
            glow(ic_bg, T.accent.lstrip('#'), radius=14, alpha=0.16)
        txt(slide, icons_map.get(lbl,"📌"), ic_x, y+0.32, ic_s, ic_s*0.86,
            font="Calibri", size=max(14, int(ic_s*10)),
            bold=False, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

        lbl_y = y + ic_s + 0.44
        txt(slide, lbl, x+0.2, lbl_y, cw-0.4, 0.62,
            font=_FONT, size=12.5, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)
        hl = rect(slide, x+cw*0.14, lbl_y+0.65, cw*0.72, 0.04, T.muted_rgb)
        txt(slide, val, x+0.2, lbl_y+0.76, cw-0.4, ch-(lbl_y-y)-0.92,
            font=_FONT, size=max(9, min(11, ch*5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    slide_number(slide, 6, getattr(req, "_total_slides", 13), T)
    return slide

# ─── STATS ─────────────────────────────────────────────────────────────
def make_stats(prs, req: PresentationRequest, T: Theme):
    from engine.primitives import blank_slide as _bs
    slide = _bs(prs)
    ultra_bg(slide, T, 'cover')
    patch_header_rtl(slide, T, _FONT, "الأرقام والإحصاءات الرئيسية", "", 7,
                     getattr(req, '_total_slides', 13), 2.9)
    CY = 2.9 + 0.3
    CH = 19.05 - CY - 0.28
    if not req.stats: return slide
    hero_stats(slide, T, _FONT, req.stats, 1.0, CY, W - 2.0, CH)
    return slide

def make_results(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📊", "نتائج", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    results = req.main_results[:8]
    n = len(results)
    if not results: return slide

    gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, result in enumerate(results):
        y = CY + i*(row_h+gap)
        even = i%2==0
        rw = rrect(slide, CX, y, CW, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=8)
        if rw:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=4, dist=1, alpha=0.16)
        acc = rect(slide, CX+CW-0.28, y, 0.28, row_h, T.accent_rgb)
        if acc:
            gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(acc, max(18, 56-i*7))
        nd = min(0.6, row_h*0.7)
        number_badge(slide, CX+CW-1.6, y+(row_h-nd)/2, nd, i+1, T)
        txt(slide, result, CX+0.25, y+0.07, CW-2.3, row_h-0.14,
            font=_FONT, size=max(10, min(12.5, row_h*8)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 8, getattr(req, "_total_slides", 13), T)
    return slide

# ─── CONCLUSION ────────────────────────────────────────────────────────
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "💡", "خاتمة", "البحث")
    CY = 0.28; CH = H - CY - 0.22; cw = CW

    cc = rrect(slide, CX, CY, cw, CH, T.card_rgb, radius_pct=14)
    if cc:
        multi_stop_gradient(cc,[(0,T.card),(50,T.bg2),(100,T.card)],135)
        shadow(cc, blur=22, dist=6, alpha=0.44)
        glow(cc, T.accent.lstrip('#'), radius=26, alpha=0.08)

    tp = rrect(slide, CX, CY, cw, 0.3, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp,[(0,T.accent2),(50,T.accent),(100,T.accent2)],0)
        glow(tp, T.accent.lstrip('#'), radius=11, alpha=0.28)

    diamond(slide, CX+0.28, CY+0.44, 0.95, 0.95, T.accent_rgb, alpha=13)

    txt(slide, "❝", CX+0.32, CY+0.42, 1.55, 1.4,
        font="Calibri", size=44, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

    txt(slide, req.general_conclusion,
        CX+0.32, CY+1.28, cw-0.9, CH-2.2,
        font=_FONT, size=max(11, min(14, int((CH-2.2)*5))), bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.35)

    ny = CY+CH-0.88
    hl = rect(slide, CX+cw*0.18, ny, cw*0.64, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)
    txt(slide, req.student_name, CX, ny+0.12, cw, 0.68,
        font=_FONT, size=13, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 9, getattr(req, "_total_slides", 13), T)
    return slide


# ─── RECOMMENDATIONS ───────────────────────────────────────────────────
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "✅", "توصيات", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    recs = req.recommendations[:8]
    n = len(recs)
    if not recs: return slide
    gap = 0.14
    row_h = (CH - gap*(n-1)) / n
    for i, rec in enumerate(recs):
        y = CY+i*(row_h+gap)
        rw = rrect(slide, CX, y, CW, row_h, T.card_rgb, radius_pct=9)
        if rw:
            multi_stop_gradient(rw,[(0,T.card),(100,T.bg2)],0)
            shadow(rw, blur=5, dist=2, alpha=0.2)
        dot = oval(slide, CX+CW-1.75, y+(row_h-0.35)/2, 0.35, 0.35, T.accent_rgb)
        if dot:
            multi_stop_gradient(dot,[(0,T.accent),(100,T.accent2)],135)
        acc = rect(slide, CX+CW-0.24, y, 0.24, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, rec, CX+0.25, y+0.07, CW-2.3, row_h-0.14,
            font=_FONT, size=max(10, min(12.5, row_h*8)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    slide_number(slide, 10, getattr(req, "_total_slides", 13), T)
    return slide

# ─── FUTURE ────────────────────────────────────────────────────────────
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "🔭", "آفاق", "مستقبلية")
    CY = 0.35; CH = H - CY - 0.3
    items = req.future_work[:6]
    if not items: return slide
    cols = 2 if len(items)>3 else 1
    rows_n = (len(items)+cols-1)//cols
    gh = 0.22; gv = 0.18
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n
    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=11)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(70,T.bg2),(100,T.bg)],155)
            shadow(cc, blur=12, dist=3, alpha=0.32)
        tp = rrect(slide, x, y, cw, 0.26, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp,[(0,T.accent),(100,T.accent2)],0)
        nd = min(0.82, ch*0.3)
        number_badge(slide, x+cw/2-nd/2, y+0.38, nd, i+1, T)
        hl = rect(slide, x+cw*0.18, y+nd+0.5, cw*0.64, 0.04, T.muted_rgb)
        txt(slide, item, x+0.25, y+nd+0.65, cw-0.5, ch-nd-0.82,
            font=_FONT, size=max(9.5, min(12, ch*5.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)
    slide_number(slide, 11, getattr(req, "_total_slides", 13), T)
    return slide

# ─── REFERENCES ────────────────────────────────────────────────────────
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📚", "المراجع", "والمصادر")
    CY = 0.35; CH = H - CY - 0.3
    refs = req.references[:12]
    n = len(refs)
    if not refs: return slide
    gap = 0.1
    row_h = min((CH - gap*(n-1)) / n, 1.35)
    total_h = n*(row_h+gap)-gap
    sy = CY + (CH-total_h)/2
    for i, ref in enumerate(refs):
        y = sy+i*(row_h+gap)
        if y+row_h > H-0.18: break
        even = i%2==0
        rw = rrect(slide, CX, y, CW, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=5)
        if rw:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw, stops, 0)
        acc = rect(slide, CX+CW-0.22, y, 0.22, row_h, T.accent_rgb)
        if acc: set_solid_alpha(acc, 52)
        nb = rrect(slide, CX+0.1, y+(row_h-0.36)/2, 0.68, 0.36, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 62)
        txt(slide, f"[{i+1}]", CX+0.1, y+(row_h-0.36)/2, 0.68, 0.36,
            font="Calibri", size=8.5, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
        txt(slide, ref, CX+0.9, y+0.04, CW-1.4, row_h-0.08,
            font=_FONT, size=max(9, min(11, row_h*7.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    slide_number(slide, 12, getattr(req, "_total_slides", 13), T)
    return slide

# ─── FINAL ─────────────────────────────────────────────────────────────
def make_final(prs, req: PresentationRequest, T: Theme):
    return cinematic_thankyou(prs, req, T, _FONT)

def make_intro(prs, req, T):
    slide = _section_slide(prs, T, "📖", "مقدمة", "البحث")
    CY=0.28; CH=H-CY-0.22
    items=[]
    if req.intro_overview: items.append(("📖","نظرة عامة",req.intro_overview))
    if req.intro_approach:  items.append(("🔬","المنهج المتبع",req.intro_approach))
    if not items: return slide
    n=len(items); gap=0.25
    cw=(CW-gap*(n-1))/n
    # ارتفاع البطاقة محدود لضمان بقاء النص بداخلها
    CARD_H = min(CH * 0.76, 10.5)
    card_y = CY + (CH - CARD_H) / 2
    ic_s = min(1.7, CARD_H * 0.22)
    ic_y_off = 0.48
    lbl_y_off = ic_y_off + ic_s + 0.26
    div_y_off = lbl_y_off + 0.68 + 0.1
    txt_y_off = div_y_off + 0.08
    txt_h = CARD_H - txt_y_off - 0.38

    for i,(icon,lbl,val) in enumerate(items[:2]):
        x=CX+i*(cw+gap)
        cc=rrect(slide,x,card_y,cw,CARD_H,T.card_rgb,radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.card)],150)
            shadow(cc,blur=16,dist=5,alpha=0.48)
        tp=rrect(slide,x,card_y,cw,0.28,T.accent_rgb,radius_pct=0)
        if tp: multi_stop_gradient(tp,[(0,T.accent),(100,T.accent2)],0)
        icon_circle(slide,x+cw/2-ic_s/2,card_y+ic_y_off,ic_s,
                    T.accent_grad1,T.accent_grad2,icon,max(14,int(ic_s*11)),T)
        txt(slide,lbl,x+0.22,card_y+lbl_y_off,cw-0.44,0.68,
            font=_FONT,size=14,bold=True,color=T.accent_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        rect(slide,x+cw*0.14,card_y+div_y_off,cw*0.72,0.04,T.muted_rgb)
        txt(slide,val,x+0.22,card_y+txt_y_off,cw-0.44,txt_h,
            font=_FONT,size=max(10,min(12,txt_h*2.2)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.3)
    slide_number(slide,1,13,T)
    return slide
