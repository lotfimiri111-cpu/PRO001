"""
Drawing Primitives — مذكرتي Pro v18.0
Professional low-level shape/text builders.

v18 improvements:
- glow() now uses real <a:glow> OOXML element (not shadow hack)
- shadow() and glow() coexist correctly in effectLst
- multi_stop_gradient() supports per-stop alpha
- gradient_fill() supports per-stop alpha
- number_badge() font size fixed
- card_3d() uses real shadow() instead of manual offset rect
- inner_glow() new primitive
- accent_line() new premium line primitive
- glass_card() new frosted-glass effect
- shimmer_rect() new animated-look shimmer
"""
from __future__ import annotations

from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# Slide dimensions (cm) — 16:9
W, H = 33.867, 19.05

# OOXML spPr required child order
_SPPR_ORDER = [
    qn('a:xfrm'), qn('a:prstGeom'), qn('a:custGeom'),
    qn('a:noFill'), qn('a:solidFill'), qn('a:gradFill'),
    qn('a:blipFill'), qn('a:pattFill'), qn('a:grpFill'),
    qn('a:ln'),
    qn('a:effectLst'), qn('a:effectDag'),
    qn('a:scene3d'), qn('a:sp3d'), qn('a:extLst'),
]
_SPPR_RANK = {tag: i for i, tag in enumerate(_SPPR_ORDER)}


def _sort_spPr(spPr) -> None:
    children = list(spPr)
    children.sort(key=lambda el: _SPPR_RANK.get(el.tag, 99))
    for child in children:
        spPr.remove(child)
    for child in children:
        spPr.append(child)


def _get_spPr(shape):
    return shape._element.find(qn('p:spPr'))


def _ensure_effectLst(spPr):
    """Get or create effectLst, then re-sort spPr."""
    eLst = spPr.find(qn('a:effectLst'))
    if eLst is None:
        eLst = etree.SubElement(spPr, qn('a:effectLst'))
    return eLst


# ── Unit helpers ─────────────────────────────────────────────────────
def cm(v: float) -> int:
    return int(Cm(v))

def pt(v: float) -> int:
    return int(Pt(v))


# ── Shape builders ───────────────────────────────────────────────────
def rect(slide, x, y, w, h, fill: RGBColor, line=None, line_w=0.5):
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(1, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = pt(line_w)
    else:
        s.line.fill.background()
    return s


def rrect(slide, x, y, w, h, fill: RGBColor, radius_pct=8, line=None, line_w=0.5):
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(5, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = pt(line_w)
    else:
        s.line.fill.background()
    try:
        adj = s.adjustments
        if adj and len(adj) > 0:
            adj[0] = max(0, min(50, radius_pct)) * 1000
    except Exception:
        pass
    return s


def oval(slide, x, y, w, h, fill: RGBColor, alpha=100):
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(9, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def bg(slide, color: RGBColor):
    rect(slide, 0, 0, W, H, color)

def hline(slide, x, y, w, color: RGBColor, thickness=0.08):
    rect(slide, x, y, w, thickness, color)

def vline(slide, x, y, h2, color: RGBColor, thickness=0.08):
    rect(slide, x, y, thickness, h2, color)


# ── XML fill helpers ─────────────────────────────────────────────────
def set_solid_alpha(shape, alpha_pct: int) -> None:
    try:
        spPr = _get_spPr(shape)
        srgb = spPr.find('.//' + qn('a:srgbClr'))
        if srgb is not None:
            for e in srgb.findall(qn('a:alpha')):
                srgb.remove(e)
            alp = etree.SubElement(srgb, qn('a:alpha'))
            alp.set('val', str(int(alpha_pct * 1000)))
    except Exception:
        pass


def gradient_fill(shape, c1: str, c2: str, angle: float = 90,
                  alpha1: int = 100, alpha2: int = 100) -> None:
    """Linear gradient with optional per-stop alpha."""
    try:
        spPr = _get_spPr(shape)
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'),
                    qn('a:pattFill'), qn('a:blipFill'), qn('a:grpFill')]:
            for el in spPr.findall(tag):
                spPr.remove(el)

        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))

        gs0 = etree.SubElement(gsLst, qn('a:gs'))
        gs0.set('pos', '0')
        clr0 = etree.SubElement(gs0, qn('a:srgbClr'))
        clr0.set('val', c1.lstrip('#'))
        if alpha1 < 100:
            a0 = etree.SubElement(clr0, qn('a:alpha'))
            a0.set('val', str(int(alpha1 * 1000)))

        gs1 = etree.SubElement(gsLst, qn('a:gs'))
        gs1.set('pos', '100000')
        clr1 = etree.SubElement(gs1, qn('a:srgbClr'))
        clr1.set('val', c2.lstrip('#'))
        if alpha2 < 100:
            a1 = etree.SubElement(clr1, qn('a:alpha'))
            a1.set('val', str(int(alpha2 * 1000)))

        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000)))
        lin.set('scaled', '0')

        spPr.append(grad)
        _sort_spPr(spPr)
    except Exception:
        pass


def gradient_rect(slide, x, y, w, h, c1: str, c2: str, angle=0):
    c1h = c1.lstrip('#')
    fill_color = RGBColor(int(c1h[0:2], 16), int(c1h[2:4], 16), int(c1h[4:6], 16))
    s = rect(slide, x, y, w, h, fill_color)
    if s:
        gradient_fill(s, c1, c2, angle)
    return s


def multi_stop_gradient(shape, stops: list[tuple], angle: float = 90) -> None:
    """
    Multi-stop gradient with optional per-stop alpha.
    stops = [(pos_pct, '#RRGGBB'), ...]
         or [(pos_pct, '#RRGGBB', alpha_pct), ...]
    """
    try:
        spPr = _get_spPr(shape)
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'),
                    qn('a:pattFill'), qn('a:blipFill'), qn('a:grpFill')]:
            for el in spPr.findall(tag):
                spPr.remove(el)

        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))

        for stop in stops:
            pos_pct  = stop[0]
            hex_color = stop[1]
            alpha_pct = stop[2] if len(stop) > 2 else 100

            gs = etree.SubElement(gsLst, qn('a:gs'))
            gs.set('pos', str(int(pos_pct * 1000)))
            clr = etree.SubElement(gs, qn('a:srgbClr'))
            clr.set('val', hex_color.lstrip('#'))
            if alpha_pct < 100:
                a = etree.SubElement(clr, qn('a:alpha'))
                a.set('val', str(int(alpha_pct * 1000)))

        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000)))
        lin.set('scaled', '0')
        spPr.append(grad)
        _sort_spPr(spPr)
    except Exception:
        pass


# ── Effects ───────────────────────────────────────────────────────────
def shadow(shape, blur=16, dist=5, angle=135, alpha=0.30,
           color="000000") -> None:
    """Outer drop shadow — preserves existing glow in effectLst."""
    try:
        spPr = _get_spPr(shape)
        eLst = _ensure_effectLst(spPr)

        # Remove existing outerShdw only
        for old in eLst.findall(qn('a:outerShdw')):
            eLst.remove(old)

        shdw = etree.SubElement(eLst, qn('a:outerShdw'))
        shdw.set('blurRad', str(int(blur * 12700)))
        shdw.set('dist',    str(int(dist * 12700)))
        shdw.set('dir',     str(int(angle * 60000)))
        shdw.set('algn', 'tl')
        srgb = etree.SubElement(shdw, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))

        _sort_spPr(spPr)
    except Exception:
        pass


def glow(shape, color: str = "C6A03C", radius: int = 20,
         alpha: float = 0.35) -> None:
    """
    Real OOXML <a:glow> effect — luminous halo around shape.
    Coexists with shadow() correctly.
    """
    try:
        spPr = _get_spPr(shape)
        eLst = _ensure_effectLst(spPr)

        # Remove existing glow only
        for old in eLst.findall(qn('a:glow')):
            eLst.remove(old)

        g = etree.Element(qn('a:glow'))
        g.set('rad', str(int(radius * 12700)))
        srgb = etree.SubElement(g, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))

        # glow must come before outerShdw in effectLst
        first_shadow = eLst.find(qn('a:outerShdw'))
        if first_shadow is not None:
            eLst.insert(list(eLst).index(first_shadow), g)
        else:
            eLst.append(g)

        _sort_spPr(spPr)
    except Exception:
        pass


def inner_glow(shape, color: str = "C6A03C", radius: int = 8,
               alpha: float = 0.25) -> None:
    """Inner glow using <a:innerShdw> with zero dist."""
    try:
        spPr = _get_spPr(shape)
        eLst = _ensure_effectLst(spPr)
        for old in eLst.findall(qn('a:innerShdw')):
            eLst.remove(old)
        ins = etree.SubElement(eLst, qn('a:innerShdw'))
        ins.set('blurRad', str(int(radius * 12700)))
        ins.set('dist', '0')
        ins.set('dir', '0')
        srgb = etree.SubElement(ins, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))
        _sort_spPr(spPr)
    except Exception:
        pass


def reflection(shape, alpha_start=0.5, alpha_end=0.0,
               dist=0, size_pct=50) -> None:
    """Subtle reflection effect below shape."""
    try:
        spPr = _get_spPr(shape)
        eLst = _ensure_effectLst(spPr)
        for old in eLst.findall(qn('a:reflection')):
            eLst.remove(old)
        ref = etree.SubElement(eLst, qn('a:reflection'))
        ref.set('blurRad', str(int(6 * 12700)))
        ref.set('stA',     str(int(alpha_start * 100000)))
        ref.set('endA',    str(int(alpha_end   * 100000)))
        ref.set('endPos',  str(int(size_pct * 1000)))
        ref.set('dist',    str(int(dist * 12700)))
        ref.set('dir',     str(int(90 * 60000)))
        ref.set('fadeDir', str(int(90 * 60000)))
        _sort_spPr(spPr)
    except Exception:
        pass


# ── Text ─────────────────────────────────────────────────────────────
def txt(slide, text: str, x, y, w, h,
        font="Cairo", size=14, bold=False, italic=False,
        color: RGBColor | None = None,
        align=PP_ALIGN.RIGHT,
        margin=0.1, rtl=True, spacing=None,
        vcenter=True, line_spacing=1.15,
        letter_spacing: int = 0):
    """
    Professional text with true vertical centering.
    letter_spacing: spacing in 100ths of a point (0 = normal, 50 = loose)
    """
    if not text or w <= 0 or h <= 0:
        return None
    try:
        from pptx.enum.text import MSO_ANCHOR
        sh = slide.shapes.add_shape(1, cm(x), cm(y), cm(w), cm(h))
        sh.fill.background()
        sh.line.fill.background()
        tf = sh.text_frame
        tf.word_wrap = True
        tf.margin_left   = cm(margin)
        tf.margin_right  = cm(margin)
        tf.margin_top    = cm(0.04)
        tf.margin_bottom = cm(0.04)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if vcenter else MSO_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.alignment = align

        # Line spacing
        try:
            pPr = p._p.get_or_add_pPr()
            for old in pPr.findall(qn('a:lnSpc')):
                pPr.remove(old)
            lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
            spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
            spcPct.set('val', str(int(line_spacing * 100000)))
        except Exception:
            pass

        run = p.add_run()
        run.text = str(text)
        run.font.name   = font
        run.font.size   = Pt(size)
        run.font.bold   = bold
        run.font.italic = italic
        if color:
            run.font.color.rgb = color

        # Letter spacing
        if letter_spacing != 0:
            try:
                rPr = run._r.get_or_add_rPr()
                rPr.set('spc', str(int(letter_spacing * 100)))
            except Exception:
                pass

        return sh
    except Exception:
        tb = slide.shapes.add_textbox(cm(x), cm(y), cm(w), cm(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.name  = font
        run.font.size  = Pt(size)
        run.font.bold  = bold
        run.font.italic = italic
        if color:
            run.font.color.rgb = color
        return tb


def txt2(slide, label: str, value: str, x, y, w, h,
         font="Cairo", label_size=10, value_size=13,
         label_color: RGBColor | None = None,
         value_color: RGBColor | None = None,
         align=PP_ALIGN.RIGHT, margin=0.1):
    if w <= 0 or h <= 0: return None
    try:
        from pptx.enum.text import MSO_ANCHOR
        sh = slide.shapes.add_shape(1, cm(x), cm(y), cm(w), cm(h))
        sh.fill.background()
        sh.line.fill.background()
        tf = sh.text_frame
        tf.word_wrap = True
        tf.margin_left   = cm(margin)
        tf.margin_right  = cm(margin)
        tf.margin_top    = cm(0.04)
        tf.margin_bottom = cm(0.04)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        p1 = tf.paragraphs[0]
        p1.alignment = align
        r1 = p1.add_run()
        r1.text = str(label)
        r1.font.name  = font
        r1.font.size  = Pt(label_size)
        r1.font.bold  = True
        if label_color: r1.font.color.rgb = label_color

        p2 = tf.add_paragraph()
        p2.alignment = align
        r2 = p2.add_run()
        r2.text = str(value)
        r2.font.name  = font
        r2.font.size  = Pt(value_size)
        r2.font.bold  = False
        if value_color: r2.font.color.rgb = value_color
        return sh
    except Exception:
        return None


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ── Advanced Visual Primitives ────────────────────────────────────────

def gradient_oval(slide, x, y, w, h, c1: str, c2: str, angle=90, alpha=100):
    if w <= 0 or h <= 0: return None
    c1h = c1.lstrip('#')
    fill_color = RGBColor(int(c1h[0:2], 16), int(c1h[2:4], 16), int(c1h[4:6], 16))
    s = slide.shapes.add_shape(9, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill_color
    s.line.fill.background()
    gradient_fill(s, c1, c2, angle)
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def triangle(slide, x, y, w, h, fill: RGBColor, pointing='up'):
    if w <= 0 or h <= 0: return None
    s = slide.shapes.add_shape(6, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if pointing == 'down':   s.rotation = 180
    elif pointing == 'left': s.rotation = 90
    elif pointing == 'right':s.rotation = 270
    return s


def diamond(slide, x, y, w, h, fill: RGBColor, alpha=100):
    if w <= 0 or h <= 0: return None
    s = slide.shapes.add_shape(4, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def hexagon(slide, x, y, w, h, fill: RGBColor, alpha=100):
    if w <= 0 or h <= 0: return None
    s = slide.shapes.add_shape(10, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def decorative_dots(slide, x, y, cols, rows, dot_size, gap,
                    color: RGBColor, alpha=20):
    for r in range(rows):
        for c in range(cols):
            dx = x + c * (dot_size + gap)
            dy = y + r * (dot_size + gap)
            o = oval(slide, dx, dy, dot_size, dot_size, color)
            if o and alpha < 100:
                set_solid_alpha(o, alpha)


def arc_progress(slide, x, y, size, fill: RGBColor, bg_color: RGBColor,
                 thickness=0.4) -> None:
    outer = oval(slide, x, y, size, size, bg_color, alpha=30)
    inner_offset = thickness
    inner_s = size - 2 * inner_offset
    oval(slide, x + inner_offset, y + inner_offset, inner_s, inner_s,
         bg_color, alpha=80)


def wave_rect(slide, x, y, w, h, fill: RGBColor, wavy_top=True):
    if wavy_top:
        return rrect(slide, x, y, w, h, fill, radius_pct=12)
    return rect(slide, x, y, w, h, fill)


def badge(slide, x, y, w, h, fill_c1: str, fill_c2: str, label: str,
          font="Cairo", font_size=11, text_color: RGBColor = None, T=None):
    b = rrect(slide, x, y, w, h, RGBColor(0xC6, 0xA0, 0x3C), radius_pct=50)
    if b:
        gradient_fill(b, fill_c1, fill_c2, angle=0)
    if text_color is None and T is not None:
        text_color = T.text_dark_rgb
    txt(slide, label, x, y, w, h,
        font=font, size=font_size, bold=True,
        color=text_color, align=PP_ALIGN.CENTER, rtl=True)
    return b


def icon_circle(slide, x, y, size, bg_c1: str, bg_c2: str,
                icon: str, icon_size=20, T=None):
    """Circle with gradient bg + centered emoji/icon + shadow."""
    c = oval(slide, x, y, size, size,
             RGBColor(int(bg_c1.lstrip('#')[0:2], 16),
                      int(bg_c1.lstrip('#')[2:4], 16),
                      int(bg_c1.lstrip('#')[4:6], 16)))
    if c:
        gradient_fill(c, bg_c1, bg_c2, angle=135)
        shadow(c, blur=12, dist=4, alpha=0.35)
        glow(c, bg_c1.lstrip('#'), radius=10, alpha=0.20)
    txt(slide, icon, x, y, size, size,
        font="Calibri", size=icon_size, bold=False,
        color=T.text_dark_rgb if T else RGBColor(0xFF, 0xFF, 0xFF),
        align=PP_ALIGN.CENTER, rtl=False)
    return c


def number_badge(slide, x, y, size, num: int | str, T):
    """Circular number badge — fixed font size calculation."""
    c = oval(slide, x, y, size, size, T.accent_rgb)
    if c:
        gradient_fill(c, T.accent_grad1, T.accent_grad2, 135)
        shadow(c, blur=10, dist=3, alpha=0.38)
    font_size = max(7, int(size * 10))   # FIX: was *18 (too large)
    txt(slide, str(num), x, y, size, size,
        font="Calibri", size=font_size, bold=True,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)
    return c


def card_3d(slide, x, y, w, h, T, radius=10):
    """Card with real shadow + gradient — no fake offset rect."""
    c = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=radius)
    if c:
        multi_stop_gradient(c, [(0, T.card), (100, T.bg2)], angle=135)
        shadow(c, blur=20, dist=6, alpha=0.42)
    return c


def glass_card(slide, x, y, w, h, T, radius=12):
    """
    Frosted-glass style card: very light fill + border + inner glow.
    Great for overlay / premium feel.
    """
    c = rrect(slide, x, y, w, h, T.bg2_rgb, radius_pct=radius,
              line=RGBColor(*[int(T.accent.lstrip('#')[i:i+2], 16) for i in (0,2,4)]),
              line_w=0.8)
    if c:
        multi_stop_gradient(c, [
            (0,   T.card,  30),   # 30% alpha top
            (50,  T.bg2,   20),
            (100, T.card,  30),
        ], angle=135)
        inner_glow(c, T.accent.lstrip('#'), radius=8, alpha=0.18)
        shadow(c, blur=18, dist=5, alpha=0.35)
    return c


def shimmer_rect(slide, x, y, w, h, T, radius=6):
    """
    Elegant shimmer: dark card with a diagonal light band.
    """
    base = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=radius)
    if base:
        multi_stop_gradient(base, [
            (0,  T.bg2),
            (45, T.card),
            (55, T.bg2),
            (100,T.bg2),
        ], angle=145)
        shadow(base, blur=14, dist=4, alpha=0.32)
    # Highlight band
    band_w = w * 0.35
    band = rrect(slide, x + w*0.15, y, band_w, h, T.bg2_rgb, radius_pct=radius)
    if band:
        multi_stop_gradient(band, [
            (0,   T.bg2, 0),
            (50,  T.bg2, 12),
            (100, T.bg2, 0),
        ], angle=90)
    return base


def accent_line(slide, x, y, w, T, style='gradient', thickness=0.06):
    """
    Premium decorative line with multiple styles.
    style: 'gradient' | 'double' | 'glow' | 'diamond_ends'
    """
    if style == 'gradient':
        d = rect(slide, x, y, w, thickness, T.accent_rgb)
        if d:
            multi_stop_gradient(d, [(0, T.bg2, 0), (30, T.accent),
                                    (70, T.accent), (100, T.bg2, 0)], 0)
    elif style == 'double':
        d1 = rect(slide, x, y, w, thickness, T.accent_rgb)
        if d1: multi_stop_gradient(d1,[(0,T.bg2,0),(50,T.accent),(100,T.bg2,0)],0)
        d2 = rect(slide, x + w*0.08, y + thickness*2.2, w*0.84, thickness*0.5,
                  T.muted_rgb)
    elif style == 'glow':
        d = rect(slide, x, y, w, thickness, T.accent_rgb)
        if d:
            multi_stop_gradient(d, [(0, T.accent2, 0), (50, T.accent),
                                    (100, T.accent2, 0)], 0)
            glow(d, T.accent.lstrip('#'), radius=8, alpha=0.45)
    elif style == 'diamond_ends':
        d = rect(slide, x + 0.8, y, w - 1.6, thickness, T.accent_rgb)
        if d: multi_stop_gradient(d,[(0,T.bg2,0),(50,T.accent),(100,T.bg2,0)],0)
        diamond(slide, x, y - thickness*2, thickness*4, thickness*4, T.accent_rgb)
        diamond(slide, x+w-thickness*4, y - thickness*2, thickness*4, thickness*4,
                T.accent_rgb)


def section_header_bar(slide, x, y, w, h, label: str, T,
                       icon: str = '', font="Cairo"):
    """
    Full-width section header bar with gradient + label + optional icon.
    Professional replacement for simple color bars.
    """
    bar = rrect(slide, x, y, w, h, T.accent_rgb, radius_pct=0)
    if bar:
        multi_stop_gradient(bar, [(0, T.accent2), (40, T.accent),
                                  (60, T.accent), (100, T.accent2)], 0)
        glow(bar, T.accent.lstrip('#'), radius=10, alpha=0.22)
        shadow(bar, blur=8, dist=2, alpha=0.28)
    # optional left icon
    if icon:
        txt(slide, icon, x + 0.18, y, h, h,
            font="Calibri", size=int(h * 10), bold=False,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
    label_x = x + (h * 1.1 if icon else 0.28)
    txt(slide, label, label_x, y, w - label_x + x - 0.3, h,
        font=font, size=max(11, int(h * 9)), bold=True,
        color=T.text_dark_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    return bar


def divider(slide, x, y, w, T, style='gradient'):
    if style == 'gradient':
        d = rect(slide, x, y, w, 0.06, T.accent_rgb)
        if d:
            multi_stop_gradient(d, [(0, T.bg, 0), (50, T.accent),
                                    (100, T.bg, 0)], 0)
    elif style == 'double':
        rect(slide, x, y, w, 0.05, T.accent_rgb)
        rect(slide, x + w*0.05, y + 0.12, w*0.9, 0.03, T.muted_rgb)
    elif style == 'glow':
        d = rect(slide, x, y, w, 0.06, T.accent_rgb)
        if d:
            multi_stop_gradient(d,[(0,T.accent2,0),(50,T.accent),(100,T.accent2,0)],0)
            glow(d, T.accent.lstrip('#'), radius=6, alpha=0.40)
    else:
        rect(slide, x, y, w, 0.06, T.accent_rgb)


def slide_number(slide, num: int, total: int, T):
    label = f"{num} / {total}"
    txt(slide, label, W - 3.5, H - 0.55, 3.2, 0.45,
        font="Calibri", size=9, bold=False,
        color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)


def watermark(slide, text: str, T):
    txt(slide, text, 0.4, H - 0.55, 6.0, 0.45,
        font="Calibri", size=8, bold=False,
        color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=False)


def section_tag(slide, label: str, x, y, T):
    w, h = 3.5, 0.52
    b = rrect(slide, x, y, w, h, T.accent_rgb, radius_pct=50)
    if b:
        gradient_fill(b, T.accent_grad1, T.accent_grad2, 0)
    txt(slide, label, x, y, w, h,
        font="Cairo", size=10, bold=True,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)
