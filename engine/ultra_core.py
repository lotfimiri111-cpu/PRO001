"""
ultra_core.py — مكتبة التحسينات المشتركة لكل المحركات الثلاثة
يُستورد من slides.py و slides_premium.py و slides_classic.py
"""
from __future__ import annotations
import math
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, decorative_dots,
    txt, blank_slide,
)
from core.themes import Theme


# ══════════════════════════════════════════════════════════════════════
# SMART TEXT ENGINE
# ══════════════════════════════════════════════════════════════════════
def trunc(text: str, max_words: int = 20) -> str:
    if not text:
        return ""
    words = text.split()
    return text if len(words) <= max_words else " ".join(words[:max_words]) + "..."


def smart_size(text: str, base: int, threshold: int = 40) -> int:
    n = len(text)
    if n <= threshold:
        return base
    if n <= threshold * 1.6:
        return max(int(base * 0.84), 9)
    return max(int(base * 0.70), 9)


# ══════════════════════════════════════════════════════════════════════
# CINEMATIC BACKGROUNDS — 5 variants
# ══════════════════════════════════════════════════════════════════════
def ultra_bg(slide, T, variant: str = 'a'):
    bg(slide, T.bg_rgb)
    angles = {'a': 135, 'b': 160, 'c': 90, 'd': 45, 'cover': 145}
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=angles.get(variant, 135))

    if variant == 'cover':
        gradient_rect(slide, -1, -1, W * 0.52, H + 2, T.bg2, T.bg, angle=165)
        s = rect(slide, W * 0.60, 0, W * 0.45, H, T.accent_rgb)
        if s:
            multi_stop_gradient(s, [(0, T.accent), (55, T.bg2), (100, T.bg)], angle=180)
            set_solid_alpha(s, 16)
        oval(slide, W * 0.48, -H * 0.28, H * 1.35, H * 1.35, T.bg2_rgb, alpha=50)
        oval(slide, -W * 0.08, H * 0.22, H * 0.75, H * 0.75, T.accent_rgb, alpha=6)
        decorative_dots(slide, 0.7, 1.4, 4, 6, 0.13, 0.36, T.accent_rgb, alpha=13)
        decorative_dots(slide, W - 5.2, H - 4.2, 5, 4, 0.11, 0.30, T.accent_rgb, alpha=9)

    elif variant == 'a':
        oval(slide, -3, -3, 11, 11, T.accent_rgb, alpha=5)
        oval(slide, W - 9, H - 8, 14, 14, T.bg2_rgb, alpha=45)
        decorative_dots(slide, 1.2, H - 4.2, 5, 3, 0.16, 0.42, T.accent_rgb, alpha=12)

    elif variant == 'b':
        s = rect(slide, 0, 0, W * 0.36, H, T.bg2_rgb)
        if s:
            gradient_fill(s, T.bg2, T.bg, angle=180)
        diamond(slide, W - 7, -2, 6, 6, T.accent_rgb, alpha=6)
        diamond(slide, -1.5, H - 4.5, 4.5, 4.5, T.accent_rgb, alpha=5)
        decorative_dots(slide, 1.0, 1.8, 4, 4, 0.15, 0.36, T.accent_rgb, alpha=9)

    elif variant == 'c':
        r2 = rect(slide, 0, H * 0.58, W, H * 0.42, T.bg2_rgb)
        if r2:
            gradient_fill(r2, T.bg2, T.bg, angle=90)
        oval(slide, -4, -3, 12, 12, T.accent_rgb, alpha=4)
        oval(slide, W - 10, H - 9, 15, 15, T.accent_rgb, alpha=4)
        decorative_dots(slide, W - 6.5, 1.5, 4, 5, 0.14, 0.35, T.accent_rgb, alpha=10)

    elif variant == 'd':
        for r, a in [(26, 4), (20, 5), (14, 6), (8, 8)]:
            oval(slide, W / 2 - r / 2, H / 2 - r / 2, r, r, T.accent_rgb, alpha=a)
        s2 = rect(slide, W * 0.68, 0, W * 0.36, H, T.bg2_rgb)
        if s2:
            gradient_fill(s2, T.accent, T.bg2, angle=90)
            set_solid_alpha(s2, 11)
        decorative_dots(slide, 1.8, H - 3.8, 5, 2, 0.18, 0.44, T.accent_rgb, alpha=12)


# ══════════════════════════════════════════════════════════════════════
# RTL BULLET LIST — نظام bullets أصيل للعربية
# ══════════════════════════════════════════════════════════════════════
def rtl_bullets(slide, T, font, items: list, x, y, w, h,
                max_items=7, font_size=13):
    """
    قائمة bullets بتصميم RTL حقيقي:
    - الشريط accent على اليمين
    - الرقم في دائرة على اليمين
    - النص يتدفق من اليمين لليسار
    - كل عنصر له خلفية متبادلة
    """
    if not items:
        return
    items = [str(i) for i in items[:max_items]]
    n = len(items)
    lh = h / n

    for i, item in enumerate(items):
        iy = y + i * lh
        item_text = trunc(item, 22)

        # خلفية متبادلة
        alpha = 38 if i % 2 == 0 else 22
        card = rrect(slide, x, iy + 0.05, w, lh - 0.10, T.card_rgb, radius_pct=6)
        if card:
            if i % 2 == 0:
                multi_stop_gradient(card, [(0, T.card), (100, T.bg2)], 0)
            else:
                multi_stop_gradient(card, [(0, T.bg2), (100, T.card)], 0)
            set_solid_alpha(card, alpha)

        # شريط accent أيمن (RTL indicator)
        ab = rect(slide, x + w - 0.16, iy + 0.12, 0.16, lh - 0.24, T.accent_rgb)
        if ab:
            gradient_fill(ab, T.accent_grad1, T.accent_grad2, 90)

        # دائرة الرقم
        ns = min(0.44, lh * 0.55)
        nx = x + w - ns - 0.28
        ny = iy + (lh - ns) / 2
        nc = oval(slide, nx, ny, ns, ns, T.accent_rgb)
        if nc:
            gradient_fill(nc, T.accent_grad1, T.accent_grad2, 135)
            shadow(nc, blur=6, dist=1, alpha=0.35)
        txt(slide, str(i + 1), nx, ny, ns, ns,
            font="Calibri", size=max(8, int(ns * 14)), bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # النص — RTL native
        tx = x + 0.22
        tw2 = w - ns - 0.22 - 0.52
        fs = smart_size(item_text, font_size, 50)
        txt(slide, item_text, tx, iy + 0.05, tw2, lh - 0.10,
            font=font, size=fs, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.15)


# ══════════════════════════════════════════════════════════════════════
# HERO STATS — شريحة إحصاءات بطولية
# ══════════════════════════════════════════════════════════════════════
def hero_stats(slide, T, font, stats, x, y, w, h):
    """أرقام hero ضخمة مع visual weight حقيقي"""
    n = len(stats[:6])
    if not n:
        return

    if n <= 3:
        # Hero — أرقام ضخمة جداً
        gap = 0.45
        cw = (w - gap * (n - 1)) / n
        CARD_H = min(h * 0.88, 11.0)
        cy = y + (h - CARD_H) / 2

        for i, stat in enumerate(stats[:3]):
            cx = x + i * (cw + gap)

            # ظل عميق
            sh = rrect(slide, cx + 0.18, cy + 0.22, cw, CARD_H, T.bg_rgb, radius_pct=14)
            if sh:
                set_solid_alpha(sh, 50)

            c = rrect(slide, cx, cy, cw, CARD_H, T.card_rgb, radius_pct=14)
            if c:
                multi_stop_gradient(c, [(0, T.card), (50, T.bg2), (100, T.card)], 145)
                shadow(c, blur=28, dist=9, alpha=0.55)

            # شريط accent أعلى
            ct = rrect(slide, cx, cy, cw, 0.3, T.accent_rgb, radius_pct=0)
            if ct:
                gradient_fill(ct, T.accent_grad1, T.accent_grad2, 0)

            # الرقم — hero
            val = str(stat.value)
            vfs = 44 if len(val) <= 4 else 34 if len(val) <= 7 else 26
            txt(slide, val, cx + 0.2, cy + 0.38, cw - 0.4, CARD_H * 0.52,
                font="Calibri", size=vfs, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

            # فاصل
            dv = rect(slide, cx + cw * 0.12, cy + CARD_H * 0.61,
                      cw * 0.76, 0.05, T.accent_rgb)
            if dv:
                multi_stop_gradient(dv, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)

            # التسمية
            lbl = trunc(str(stat.label), 16)
            txt(slide, lbl, cx + 0.2, cy + CARD_H * 0.65, cw - 0.4,
                CARD_H * 0.28,
                font=font, size=smart_size(lbl, 13, 18), bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.CENTER,
                rtl=True, vcenter=True, line_spacing=1.1)

            if getattr(stat, 'unit', ''):
                txt(slide, stat.unit, cx + 0.2, cy + CARD_H * 0.52,
                    cw - 0.4, CARD_H * 0.1,
                    font=font, size=10, bold=False,
                    color=T.muted_rgb, align=PP_ALIGN.CENTER,
                    rtl=True, vcenter=True)
    else:
        # Grid — 3 أعمدة
        cols = 3
        rows = math.ceil(n / cols)
        gx, gy = 0.38, 0.35
        cw = (w - gx * (cols - 1)) / cols
        ch = (h - gy * (rows - 1)) / rows

        for i, stat in enumerate(stats[:6]):
            col = i % cols
            row = i // cols
            cx = x + col * (cw + gx)
            cy = y + row * (ch + gy)

            c = rrect(slide, cx, cy, cw, ch, T.card_rgb, radius_pct=10)
            if c:
                multi_stop_gradient(c, [(0, T.card), (100, T.bg2)], 135)
                shadow(c, blur=16, dist=4, alpha=0.42)

            # شريط accent أيمن
            vline(slide, cx + cw - 0.16, cy, ch, T.accent_rgb, thickness=0.16)

            val = str(stat.value)
            vfs = 28 if len(val) <= 5 else 22
            txt(slide, val, cx + 0.15, cy + 0.1, cw - 0.45, ch * 0.55,
                font="Calibri", size=vfs, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

            lbl = trunc(str(stat.label), 14)
            txt(slide, lbl, cx + 0.12, cy + ch * 0.6, cw - 0.38, ch * 0.32,
                font=font, size=10, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.CENTER,
                rtl=True, vcenter=True, line_spacing=1.1)


# ══════════════════════════════════════════════════════════════════════
# CINEMATIC COVER — غلاف سينمائي مشترك
# ══════════════════════════════════════════════════════════════════════
def cinematic_cover(prs, req, T, font):
    """غلاف سينمائي موحّد لكل المحركات"""
    slide = blank_slide(prs)
    ultra_bg(slide, T, 'cover')

    # شريط المؤسسة
    top = rect(slide, 0, 0, W, 0.50, T.bg2_rgb)
    if top:
        gradient_fill(top, T.bg2, T.bg, angle=0)
        set_solid_alpha(top, 78)
    hl_top = rect(slide, 0, 0.50, W, 0.055, T.accent_rgb)
    if hl_top:
        multi_stop_gradient(hl_top, [(0, T.bg), (30, T.accent), (70, T.accent2), (100, T.bg)], 0)

    if getattr(req, 'institution', ''):
        txt(slide, req.institution, 0.8, 0, W - 1.6, 0.50,
            font=font, size=10.5, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # البطاقة الرئيسية
    title_len = len(req.title_ar)
    title_h = 7.0 if title_len < 55 else 7.7 if title_len < 85 else 8.4
    cx, cw2 = 1.5, W - 3.0
    cy = 0.82

    # ظل
    sh = rrect(slide, cx + 0.22, cy + 0.26, cw2, title_h, T.bg_rgb, radius_pct=12)
    if sh:
        set_solid_alpha(sh, 52)

    mc = rrect(slide, cx, cy, cw2, title_h, T.card_rgb, radius_pct=12)
    if mc:
        multi_stop_gradient(mc, [(0, T.card), (60, T.bg2), (100, T.card)], 145)
        shadow(mc, blur=30, dist=9, alpha=0.56)

    # شريط accent علوي
    ct = rrect(slide, cx, cy, cw2, 0.30, T.accent_rgb, radius_pct=0)
    if ct:
        gradient_fill(ct, T.accent_grad1, T.accent_grad2, 0)

    # شريط accent أيمن (RTL)
    vline(slide, cx + cw2 - 0.20, cy + 0.30, title_h - 0.30, T.accent_rgb, thickness=0.20)

    # أيقونة
    ic_s = 1.55
    ic_x = cx + cw2 - ic_s - 0.55
    ic_y = cy + 0.52
    ic_bg = oval(slide, ic_x, ic_y, ic_s, ic_s, T.accent_rgb)
    if ic_bg:
        gradient_fill(ic_bg, T.accent_grad1, T.accent_grad2, 135)
        shadow(ic_bg, blur=16, dist=5, alpha=0.44)
    txt(slide, "🎓", ic_x, ic_y, ic_s, ic_s,
        font="Calibri", size=25, bold=False,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # العنوان — hero
    tfs = 26 if title_len < 42 else 21 if title_len < 68 else 17
    txt(slide, req.title_ar,
        cx + 0.42, cy + 0.36, cw2 - ic_s - 1.2, title_h * 0.64,
        font=font, size=tfs, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
        rtl=True, vcenter=True, line_spacing=1.25)

    if getattr(req, 'title_en', ''):
        txt(slide, req.title_en,
            cx + 0.42, cy + title_h * 0.64, cw2 - 0.9, title_h * 0.2,
            font="Calibri", size=11, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # خط فاصل
    dv = rect(slide, cx + cw2 * 0.08, cy + title_h * 0.86,
              cw2 * 0.84, 0.048, T.accent_rgb)
    if dv:
        multi_stop_gradient(dv, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)

    # بطاقة المعلومات
    info_y = cy + title_h + 0.20
    info_h = H - 0.36 - info_y
    if info_h > 0.5:
        rows = []
        if req.student_name:
            rows.append(("👤", "الطالب", req.student_name))
        if getattr(req, 'supervisor', ''):
            rows.append(("👨‍🏫", "المشرف", req.supervisor))
        if getattr(req, 'specialization', ''):
            rows.append(("📚", "التخصص", req.specialization))
        if getattr(req, 'year', ''):
            rows.append(("📅", "السنة", req.year))

        ic2 = rrect(slide, cx, info_y, cw2, info_h, T.card_rgb, radius_pct=10)
        if ic2:
            multi_stop_gradient(ic2, [(0, T.bg2), (100, T.card)], 135)
            shadow(ic2, blur=14, dist=4, alpha=0.34)
        vline(slide, cx + cw2 - 0.16, info_y, info_h, T.accent_rgb, thickness=0.16)

        rh = info_h / max(len(rows), 1)
        for i, (icon, lbl, val) in enumerate(rows):
            ry = info_y + i * rh
            if i > 0:
                hline(slide, cx + 0.28, ry, cw2 - 0.44, T.muted_rgb, thickness=0.03)

            # أيقونة
            ic_s2 = min(rh * 0.52, 0.62)
            ic_x2 = cx + cw2 - ic_s2 - 0.36
            ic_y2 = ry + (rh - ic_s2) / 2
            ci = oval(slide, ic_x2, ic_y2, ic_s2, ic_s2, T.accent_rgb)
            if ci:
                gradient_fill(ci, T.accent_grad1, T.accent_grad2, 135)
                set_solid_alpha(ci, 68)
            txt(slide, icon, ic_x2, ic_y2, ic_s2, ic_s2,
                font="Calibri", size=int(ic_s2 * 15), bold=False,
                color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

            # التسمية
            txt(slide, f"{lbl}:", cx + cw2 - ic_s2 - 3.5, ry, 2.9, rh,
                font=font, size=max(9, min(11, int(rh * 7))), bold=True,
                color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

            # فاصل
            vline(slide, cx + cw2 - ic_s2 - 3.62, ry + rh * 0.18,
                  rh * 0.64, T.muted_rgb, thickness=0.04)

            # القيمة
            txt(slide, val, cx + 0.36, ry, cw2 - ic_s2 - 4.05, rh,
                font=font, size=max(10, min(13, int(rh * 8.5))), bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
                rtl=True, vcenter=True)

    # شريط سفلي
    bot = rect(slide, 0, H - 0.28, W, 0.28, T.accent_rgb)
    if bot:
        multi_stop_gradient(bot, [(0, T.bg), (40, T.accent), (60, T.accent2), (100, T.bg)], 0)

    return slide


# ══════════════════════════════════════════════════════════════════════
# CINEMATIC THANK YOU — خاتمة سينمائية مشتركة
# ══════════════════════════════════════════════════════════════════════
def cinematic_thankyou(prs, req, T, font):
    """خاتمة سينمائية موحّدة"""
    slide = blank_slide(prs)
    ultra_bg(slide, T, 'cover')

    # دوائر عمق
    for r, a in [(9, 4), (7, 6), (5.5, 8), (4, 10)]:
        oval(slide, W / 2 - r / 2, H / 2 - r / 2, r, r, T.accent_rgb, alpha=a)

    # شريط علوي
    top = rect(slide, 0, 0, W, 0.42, T.bg_rgb)
    if top:
        gradient_fill(top, T.bg, T.bg2, 0)
        set_solid_alpha(top, 78)
    rect(slide, 0, 0.42, W, 0.055, T.accent_rgb)

    # شكراً — hero
    txt(slide, "شكراً جزيلاً", 0.5, H * 0.16, W - 1.0, H * 0.26,
        font=font, size=52, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER,
        rtl=True, vcenter=True, line_spacing=1.1)

    txt(slide, "وتقديراً", 0.5, H * 0.40, W - 1.0, H * 0.12,
        font=font, size=28, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.CENTER,
        rtl=True, vcenter=True)

    # خط ذهبي
    dv = rect(slide, W * 0.2, H * 0.55, W * 0.6, 0.055, T.accent_rgb)
    if dv:
        multi_stop_gradient(dv, [(0, T.bg2), (50, T.accent), (100, T.bg2)], 0)

    # بطاقة التوقيع
    iy = H * 0.60
    ih = H - iy - 0.52
    ix = W * 0.18
    iw = W * 0.64

    ic = rrect(slide, ix, iy, iw, ih, T.card_rgb, radius_pct=12)
    if ic:
        multi_stop_gradient(ic, [(0, T.bg2), (100, T.card)], 135)
        shadow(ic, blur=24, dist=7, alpha=0.50)
    vline(slide, ix + iw - 0.20, iy, ih, T.accent_rgb, thickness=0.20)

    rows_t = []
    if req.student_name:
        rows_t.append((req.student_name, 15, True, T.text_light_rgb))
    if getattr(req, 'supervisor', ''):
        rows_t.append((f"إشراف: {req.supervisor}", 12, False, T.muted_rgb))
    if getattr(req, 'year', ''):
        rows_t.append((req.year, 11, False, T.muted_rgb))

    rh = ih / max(len(rows_t), 1)
    for i, (text, fs, bold, color) in enumerate(rows_t):
        txt(slide, text, ix + 0.3, iy + i * rh, iw - 0.5, rh,
            font=font, size=fs, bold=bold,
            color=color, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # شريط سفلي
    bot = rect(slide, 0, H - 0.36, W, 0.36, T.bg_rgb)
    if bot:
        gradient_fill(bot, T.bg2, T.bg, 0)
    hl2 = rect(slide, 0, H - 0.36, W, 0.055, T.accent_rgb)
    if hl2:
        multi_stop_gradient(hl2, [(0, T.bg), (40, T.accent), (60, T.accent2), (100, T.bg)], 0)

    txt(slide, "مذكرتي Pro ✦", 0, H - 0.33, W, 0.33,
        font=font, size=9, bold=False,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    return slide


# ══════════════════════════════════════════════════════════════════════
# SMART HEADER PATCH — يُحسّن أي هيدر موجود
# ══════════════════════════════════════════════════════════════════════
def patch_header_rtl(slide, T, font, title, sub='', num=None, total=13, header_h=2.9):
    """
    هيدر RTL محسّن — يُستخدم بدلاً من _hdr القديم
    الفرق: شريط accent على اليمين دائماً، رقم الشريحة على اليسار
    """
    gradient_rect(slide, 0, 0, W, header_h, T.grad2, T.grad1, angle=135)

    # شريط accent أسفل الهيدر
    al = rect(slide, 0, header_h - 0.20, W, 0.20, T.accent_rgb)
    if al:
        multi_stop_gradient(al, [(0, T.bg), (38, T.accent), (62, T.accent2), (100, T.bg)], 0)

    # شريط accent أيمن — RTL دائماً
    av = rect(slide, W - 0.50, 0, 0.50, header_h - 0.20, T.accent_rgb)
    if av:
        gradient_fill(av, T.accent_grad1, T.accent_grad2, 90)

    # رقم الشريحة — أيسر
    title_x = 0.72
    if num is not None:
        nb_s = 0.70
        nb_x = 0.85
        nb_y = (header_h - 0.20 - nb_s) / 2
        nb = oval(slide, nb_x, nb_y, nb_s, nb_s, T.accent_rgb)
        if nb:
            gradient_fill(nb, T.accent_grad1, T.accent_grad2, 135)
            shadow(nb, blur=8, dist=2, alpha=0.38)
        txt(slide, str(num), nb_x, nb_y, nb_s, nb_s,
            font="Calibri", size=13, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        txt(slide, f"/{total}", nb_x + nb_s, nb_y + nb_s * 0.30, 0.75, nb_s * 0.40,
            font="Calibri", size=8, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
        title_x = nb_x + nb_s + 0.80

    # العنوان
    title_w = W - title_x - 0.65
    tfs = smart_size(title, 24, 36)
    txt(slide, title, title_x, 0.18, title_w, header_h * 0.62,
        font=font, size=tfs, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
        rtl=True, vcenter=True, line_spacing=1.1)

    if sub:
        txt(slide, sub, title_x, header_h * 0.60, title_w, header_h * 0.36,
            font=font, size=11, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.0)
