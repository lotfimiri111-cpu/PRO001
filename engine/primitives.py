"""
Drawing Primitives — مذكرتي Pro v17.1
Low-level, deterministic shape/text builders.

FIXED BUGS vs v17.0:
- gradient_fill() now inserts gradFill in CORRECT OOXML position (before <a:ln>)
- shadow() now inserts effectLst in CORRECT position (after <a:ln>)
- _sort_spPr() enforces strict OOXML child order on every shape
- set_solid_alpha() exposed (no duplication with slides.py)
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
    """Reorder <p:spPr> children to comply with OOXML schema."""
    children = list(spPr)
    children.sort(key=lambda el: _SPPR_RANK.get(el.tag, 99))
    for child in children:
        spPr.remove(child)
    for child in children:
        spPr.append(child)


def _get_spPr(shape):
    return shape._element.find(qn('p:spPr'))


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
    """Set transparency on a solidFill shape (0=transparent, 100=opaque)."""
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


def gradient_fill(shape, c1: str, c2: str, angle: float = 90) -> None:
    """
    Apply linear gradient via XML, with correct OOXML element ordering.
    BUG FIX: previously appended gradFill AFTER <a:ln> — now sorted correctly.
    """
    try:
        spPr = _get_spPr(shape)
        # Remove all existing fill variants
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'),
                    qn('a:pattFill'), qn('a:blipFill'), qn('a:grpFill')]:
            for el in spPr.findall(tag):
                spPr.remove(el)

        # Build gradFill
        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))

        gs0 = etree.SubElement(gsLst, qn('a:gs'))
        gs0.set('pos', '0')
        etree.SubElement(gs0, qn('a:srgbClr')).set('val', c1.lstrip('#'))

        gs1 = etree.SubElement(gsLst, qn('a:gs'))
        gs1.set('pos', '100000')
        etree.SubElement(gs1, qn('a:srgbClr')).set('val', c2.lstrip('#'))

        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000)))
        lin.set('scaled', '0')

        spPr.append(grad)
        _sort_spPr(spPr)  # ← enforce correct order
    except Exception:
        pass


def gradient_rect(slide, x, y, w, h, c1: str, c2: str, angle=0):
    c1h = c1.lstrip('#')
    fill_color = RGBColor(int(c1h[0:2], 16), int(c1h[2:4], 16), int(c1h[4:6], 16))
    s = rect(slide, x, y, w, h, fill_color)
    if s:
        gradient_fill(s, c1, c2, angle)
    return s


def shadow(shape, blur=16, dist=5, angle=135, alpha=0.22, color="000000") -> None:
    """
    Add outer drop shadow via XML.
    BUG FIX: effectLst now inserted in correct OOXML position (after <a:ln>).
    """
    try:
        spPr = _get_spPr(shape)
        for old in spPr.findall(qn('a:effectLst')):
            spPr.remove(old)

        eLst = etree.Element(qn('a:effectLst'))
        shdw = etree.SubElement(eLst, qn('a:outerShdw'))
        shdw.set('blurRad', str(int(blur * 12700)))
        shdw.set('dist', str(int(dist * 12700)))
        shdw.set('dir', str(int(angle * 60000)))
        shdw.set('algn', 'tl')
        srgb = etree.SubElement(shdw, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))

        spPr.append(eLst)
        _sort_spPr(spPr)  # ← enforce correct order
    except Exception:
        pass


# ── Text ─────────────────────────────────────────────────────────────
def txt(slide, text: str, x, y, w, h,
        font="Cairo", size=14, bold=False, italic=False,
        color: RGBColor | None = None,
        align=PP_ALIGN.RIGHT,
        margin=0.1, rtl=True, spacing=None,
        vcenter=True, line_spacing=1.15):
    """
    نص احترافي باستخدام Shape مع توسيط عمودي حقيقي.
    vcenter=True → MSO_ANCHOR.MIDDLE (يعمل في PowerPoint وLibreOffice)
    line_spacing → ارتفاع السطر النسبي
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

        if vcenter:
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        else:
            tf.vertical_anchor = MSO_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.alignment = align

        # ارتفاع السطر
        try:
            from pptx.oxml.ns import qn as _qn
            pPr = p._p.get_or_add_pPr()
            # أزل lnSpc القديم إن وُجد
            for old in pPr.findall(_qn('a:lnSpc')):
                pPr.remove(old)
            lnSpc = etree.SubElement(pPr, _qn('a:lnSpc'))
            spcPct = etree.SubElement(lnSpc, _qn('a:spcPct'))
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
        return sh
    except Exception:
        # fallback إلى textbox
        tb = slide.shapes.add_textbox(cm(x), cm(y), cm(w), cm(h))
        tb.word_wrap = True
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.name   = font
        run.font.size   = Pt(size)
        run.font.bold   = bold
        run.font.italic = italic
        if color:
            run.font.color.rgb = color
        return tb


def txt2(slide, label: str, value: str, x, y, w, h,
         font="Cairo", label_size=10, value_size=13,
         label_color: RGBColor | None = None,
         value_color: RGBColor | None = None,
         align=PP_ALIGN.RIGHT, margin=0.1):
    """
    نص بسطرين: تسمية + قيمة مع توسيط عمودي.
    مثالي لبطاقات المعلومات.
    """
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
    """Add a completely blank slide (layout 6 = Blank)."""
    return prs.slides.add_slide(prs.slide_layouts[6])


# ── Advanced Visual Primitives ────────────────────────────────────────

def glow(shape, color: str = "C6A03C", radius: int = 20, alpha: float = 0.4) -> None:
    """Add a glow effect (outerShdw with zero distance = glow)."""
    try:
        spPr = _get_spPr(shape)
        for old in spPr.findall(qn('a:effectLst')):
            spPr.remove(old)
        eLst = etree.Element(qn('a:effectLst'))
        g = etree.SubElement(eLst, qn('a:outerShdw'))
        g.set('blurRad', str(int(radius * 12700)))
        g.set('dist', '0')
        g.set('dir', '0')
        g.set('algn', 'ctr')
        srgb = etree.SubElement(g, qn('a:srgbClr'))
        srgb.set('val', color.lstrip('#'))
        alp = etree.SubElement(srgb, qn('a:alpha'))
        alp.set('val', str(int(alpha * 100000)))
        spPr.append(eLst)
        _sort_spPr(spPr)
    except Exception:
        pass


def multi_stop_gradient(shape, stops: list[tuple[int, str]], angle: float = 90) -> None:
    """
    Apply a multi-stop gradient.
    stops = [(pos_pct, '#RRGGBB'), ...]  e.g. [(0,'#07172F'),(50,'#1A3A6E'),(100,'#07172F')]
    """
    try:
        spPr = _get_spPr(shape)
        for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'),
                    qn('a:pattFill'), qn('a:blipFill'), qn('a:grpFill')]:
            for el in spPr.findall(tag):
                spPr.remove(el)
        grad = etree.Element(qn('a:gradFill'))
        gsLst = etree.SubElement(grad, qn('a:gsLst'))
        for pos_pct, hex_color in stops:
            gs = etree.SubElement(gsLst, qn('a:gs'))
            gs.set('pos', str(int(pos_pct * 1000)))
            etree.SubElement(gs, qn('a:srgbClr')).set('val', hex_color.lstrip('#'))
        lin = etree.SubElement(grad, qn('a:lin'))
        lin.set('ang', str(int(angle * 60000)))
        lin.set('scaled', '0')
        spPr.append(grad)
        _sort_spPr(spPr)
    except Exception:
        pass


def gradient_oval(slide, x, y, w, h, c1: str, c2: str, angle=90, alpha=100):
    """Oval with gradient fill."""
    if w <= 0 or h <= 0:
        return None
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
    """Equilateral-ish triangle shape."""
    if w <= 0 or h <= 0:
        return None
    # Use right-triangle preset (shape 6) rotated for pointing direction
    shape_id = 6  # rtTriangle
    s = slide.shapes.add_shape(shape_id, cm(x), cm(y), cm(w), cm(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if pointing == 'down':
        s.rotation = 180
    elif pointing == 'left':
        s.rotation = 90
    elif pointing == 'right':
        s.rotation = 270
    return s


def diamond(slide, x, y, w, h, fill: RGBColor, alpha=100):
    """Diamond shape."""
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(4, cm(x), cm(y), cm(w), cm(h))  # shape 4 = diamond
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def hexagon(slide, x, y, w, h, fill: RGBColor, alpha=100):
    """Hexagon shape."""
    if w <= 0 or h <= 0:
        return None
    s = slide.shapes.add_shape(10, cm(x), cm(y), cm(w), cm(h))  # shape 10 = hexagon
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    if alpha < 100:
        set_solid_alpha(s, alpha)
    return s


def arc_progress(slide, x, y, size, fill: RGBColor, bg_color: RGBColor,
                 thickness=0.4) -> None:
    """Simulated progress ring using two arcs (outer ring + inner mask)."""
    # Outer ring background
    outer = oval(slide, x, y, size, size, bg_color, alpha=30)
    # Inner circle mask (creates ring effect)
    inner_offset = thickness
    inner_s = size - 2 * inner_offset
    oval(slide, x + inner_offset, y + inner_offset, inner_s, inner_s, bg_color, alpha=80)


def decorative_dots(slide, x, y, cols, rows, dot_size, gap, color: RGBColor, alpha=20):
    """Grid of decorative dots."""
    for r in range(rows):
        for c in range(cols):
            dx = x + c * (dot_size + gap)
            dy = y + r * (dot_size + gap)
            o = oval(slide, dx, dy, dot_size, dot_size, color)
            if o and alpha < 100:
                set_solid_alpha(o, alpha)


def wave_rect(slide, x, y, w, h, fill: RGBColor, wavy_top=True):
    """Rectangle with rounded top (simulates wave). Uses rrect with high radius."""
    if wavy_top:
        return rrect(slide, x, y, w, h, fill, radius_pct=12)
    return rect(slide, x, y, w, h, fill)


def badge(slide, x, y, w, h, fill_c1: str, fill_c2: str, label: str,
          font="Cairo", font_size=11, text_color: RGBColor = None, T=None):
    """Pill-shaped badge with gradient and centered label."""
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
    """Circle with gradient bg + centered emoji/icon."""
    c = oval(slide, x, y, size, size,
             RGBColor(int(bg_c1.lstrip('#')[0:2], 16),
                      int(bg_c1.lstrip('#')[2:4], 16),
                      int(bg_c1.lstrip('#')[4:6], 16)))
    if c:
        gradient_fill(c, bg_c1, bg_c2, angle=135)
    txt(slide, icon, x, y, size, size,
        font="Calibri", size=icon_size, bold=False,
        color=T.text_dark_rgb if T else RGBColor(0xFF, 0xFF, 0xFF),
        align=PP_ALIGN.CENTER, rtl=False)
    return c


def number_badge(slide, x, y, size, num: int | str, T):
    """Circular number badge with accent gradient."""
    c = oval(slide, x, y, size, size, T.accent_rgb)
    if c:
        gradient_fill(c, T.accent_grad1, T.accent_grad2, 135)
        shadow(c, blur=10, dist=3, alpha=0.35)
    txt(slide, str(num), x, y, size, size,
        font="Calibri", size=max(8, int(size * 18)), bold=True,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)
    return c


def divider(slide, x, y, w, T, style='gradient'):
    """Decorative divider line."""
    if style == 'gradient':
        d = rect(slide, x, y, w, 0.06, T.accent_rgb)
        if d:
            multi_stop_gradient(d, [(0, T.bg), (50, T.accent), (100, T.bg)], angle=0)
    elif style == 'double':
        rect(slide, x, y, w, 0.05, T.accent_rgb)
        rect(slide, x + w * 0.05, y + 0.12, w * 0.9, 0.03, T.muted_rgb)
    else:
        rect(slide, x, y, w, 0.06, T.accent_rgb)


def card_3d(slide, x, y, w, h, T, radius=10):
    """Card with shadow + subtle gradient for 3D feel."""
    # Shadow layer (slightly offset)
    shadow_s = rrect(slide, x + 0.15, y + 0.2, w, h, T.bg_rgb, radius_pct=radius)
    if shadow_s:
        set_solid_alpha(shadow_s, 40)

    # Main card
    c = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=radius)
    if c:
        multi_stop_gradient(c, [(0, T.card), (100, T.bg2)], angle=135)
        shadow(c, blur=18, dist=5, alpha=0.4)
    return c


def slide_number(slide, num: int, total: int, T):
    """Slide number indicator bottom-right."""
    label = f"{num} / {total}"
    txt(slide, label, W - 3.5, H - 0.55, 3.2, 0.45,
        font="Calibri", size=9, bold=False,
        color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)


def watermark(slide, text: str, T):
    """Subtle watermark bottom-left."""
    txt(slide, text, 0.4, H - 0.55, 6.0, 0.45,
        font="Calibri", size=8, bold=False,
        color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=False)


def section_tag(slide, label: str, x, y, T):
    """Small colored tag/label."""
    w, h = 3.5, 0.52
    b = rrect(slide, x, y, w, h, T.accent_rgb, radius_pct=50)
    if b:
        gradient_fill(b, T.accent_grad1, T.accent_grad2, 0)
    txt(slide, label, x, y, w, h,
        font="Cairo", size=10, bold=True,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PREMIUM DESIGN SYSTEM v28 — Advanced Visual Intelligence
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def accent_bar_v(slide, x, y, h, T, width=0.18, alpha=100):
    """Vertical accent bar — signature premium left/right edge."""
    b = rect(slide, x, y, width, h, T.accent_rgb)
    if b:
        gradient_fill(b, T.accent_grad1, T.accent_grad2, 90)
        if alpha < 100:
            set_solid_alpha(b, alpha)
    return b


def kpi_card(slide, x, y, w, h, value: str, label: str, T,
             unit: str = '', icon: str = '', rank: int = 0):
    """
    KPI stat card — content-aware sizing.
    value: big number/text  label: description  unit: optional unit
    Automatically scales font based on value length.
    """
    # Outer card
    c = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=14)
    if c:
        multi_stop_gradient(c, [(0, T.bg2), (60, T.card), (100, T.bg2)], 135)
        shadow(c, blur=20, dist=6, alpha=0.45)

    # Top accent strip
    tp = rrect(slide, x, y, w, 0.26, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp, [(0, T.accent2), (50, T.accent), (100, T.accent2)], 0)

    # Bottom accent strip
    bp = rrect(slide, x, y + h - 0.18, w, 0.18, T.accent_rgb, radius_pct=0)
    if bp:
        set_solid_alpha(bp, 35)

    # Value — dynamically sized
    vlen = len(str(value))
    vs = 44 if vlen <= 3 else 34 if vlen <= 6 else 26 if vlen <= 10 else 20
    value_h = h * 0.52
    txt(slide, str(value), x + 0.1, y + 0.26, w - 0.2, value_h,
        font='Calibri', size=vs, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Unit badge (if provided)
    if unit:
        ub = rrect(slide, x + w/2 - 1.3, y + 0.26 + value_h - 0.05, 2.6, 0.38,
                   T.bg_rgb, radius_pct=40)
        if ub:
            set_solid_alpha(ub, 48)
        txt(slide, unit, x + w/2 - 1.3, y + 0.26 + value_h - 0.05, 2.6, 0.38,
            font='Cairo', size=9, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # Divider
    div_y = y + 0.26 + value_h + (0.44 if unit else 0.06)
    hline(slide, x + w*0.12, div_y, w*0.76, T.muted_rgb, thickness=0.04)

    # Label
    label_h = h - (div_y - y) - 0.28
    txt(slide, label, x + 0.15, div_y + 0.08, w - 0.3, max(0.4, label_h),
        font='Cairo', size=max(11, min(13, label_h * 7)), bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER,
        rtl=True, vcenter=True, line_spacing=1.15)
    return c


def result_row_premium(slide, x, y, w, h, text: str, num: int, T,
                        highlight: bool = False):
    """
    Premium result row — content-aware with optional highlight.
    Highlight = first/key result gets special treatment.
    """
    if highlight:
        # Highlighted row: gradient + glow
        rw = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=10)
        if rw:
            multi_stop_gradient(rw, [(0, T.card), (100, T.bg2)], 0)
            shadow(rw, blur=12, dist=3, alpha=0.35)
            glow(rw, T.accent.lstrip('#'), radius=15, alpha=0.08)
        # Left accent line (RTL: right side)
        al = rect(slide, x + w - 0.32, y, 0.32, h, T.accent_rgb)
        if al:
            gradient_fill(al, T.accent_grad1, T.accent_grad2, 90)
    else:
        even = num % 2 == 0
        rw = rrect(slide, x, y, w, h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=8)
        if rw:
            stops = [(0, T.card), (100, T.bg2)] if even else [(0, T.bg2), (100, T.card)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=4, dist=1, alpha=0.15)
        al = rect(slide, x + w - 0.28, y, 0.28, h, T.accent_rgb)
        if al:
            gradient_fill(al, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(al, max(18, 55 - num * 6))

    # Number badge
    nd = min(0.68, h * 0.72)
    nb_x = x + w - 0.32 - 0.22 - nd
    nb_y = y + (h - nd) / 2
    nc = oval(slide, nb_x, nb_y, nd, nd, T.accent_rgb)
    if nc:
        multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
        shadow(nc, blur=6, dist=2, alpha=0.3)
    txt(slide, str(num), nb_x, nb_y, nd, nd,
        font='Calibri', size=max(8, int(nd * 12)), bold=True,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Text — content-aware font size
    text_w = w - nd - 0.65 - (0.35 if highlight else 0.30)
    char_count = len(text)
    fs = 14 if char_count < 60 else 13 if char_count < 100 else 12 if char_count < 150 else 11
    fs = max(fs, min(15, h * 9))
    txt(slide, text, x + 0.28, y, text_w, h,
        font='Cairo', size=fs, bold=highlight, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.25)
    return rw


def section_header_band(slide, x, y, w, h, title: str, subtitle: str, T,
                          icon: str = ''):
    """Premium section header band — replaces simple label."""
    bg_b = rrect(slide, x, y, w, h, T.bg2_rgb, radius_pct=10)
    if bg_b:
        multi_stop_gradient(bg_b, [(0, T.card), (100, T.bg2)], 0)
        shadow(bg_b, blur=10, dist=3, alpha=0.25)

    # Left accent bar (RTL: right)
    accent_bar_v(slide, x + w - 0.22, y, h, T, width=0.22)

    if icon:
        txt(slide, icon, x + 0.3, y, h * 1.2, h,
            font='Calibri', size=int(h * 12), bold=False,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
        text_x = x + h * 1.1 + 0.5
    else:
        text_x = x + 0.4

    title_w = w - (text_x - x) - 0.4
    if subtitle:
        txt(slide, title, text_x, y, title_w, h * 0.58,
            font='Cairo', size=max(13, int(h * 11)), bold=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        txt(slide, subtitle, text_x, y + h * 0.56, title_w, h * 0.44,
            font='Cairo', size=max(10, int(h * 8)), bold=False,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    else:
        txt(slide, title, text_x, y, title_w, h,
            font='Cairo', size=max(13, int(h * 11)), bold=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
    return bg_b


def highlight_chip(slide, x, y, w, h, text: str, T, style='accent'):
    """Pill-shaped highlight chip for key terms."""
    if style == 'accent':
        b = rrect(slide, x, y, w, h, T.accent_rgb, radius_pct=50)
        if b:
            gradient_fill(b, T.accent_grad1, T.accent_grad2, 0)
            shadow(b, blur=8, dist=2, alpha=0.3)
        tc = T.text_dark_rgb
    elif style == 'muted':
        b = rrect(slide, x, y, w, h, T.bg2_rgb, radius_pct=50)
        if b:
            set_solid_alpha(b, 70)
        tc = T.muted_rgb
    else:  # ghost
        b = rrect(slide, x, y, w, h, T.accent_rgb, radius_pct=50)
        if b:
            set_solid_alpha(b, 18)
        tc = T.accent_rgb

    txt(slide, text, x, y, w, h,
        font='Cairo', size=max(9, int(h * 10)), bold=True,
        color=tc, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)
    return b


def content_card_premium(slide, x, y, w, h, T, accent_side='right',
                           radius=12, depth=True):
    """
    Premium content card with optional 3D depth effect.
    More sophisticated than card_3d — uses multi-layer system.
    """
    if depth:
        # Shadow layer
        sh_s = rrect(slide, x + 0.18, y + 0.25, w, h, T.bg_rgb, radius_pct=radius)
        if sh_s:
            set_solid_alpha(sh_s, 35)

    # Main card body
    c = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=radius)
    if c:
        multi_stop_gradient(c, [(0, T.card), (55, T.bg2), (100, T.card)], 135)
        shadow(c, blur=22, dist=7, alpha=0.48)

    # Accent edge
    edge_w = 0.25
    if accent_side == 'right':
        e = rrect(slide, x + w - edge_w, y, edge_w, h, T.accent_rgb, radius_pct=0)
    else:
        e = rrect(slide, x, y, edge_w, h, T.accent_rgb, radius_pct=0)
    if e:
        gradient_fill(e, T.accent_grad1, T.accent_grad2, 90)

    # Subtle inner highlight (top edge)
    hl = rrect(slide, x, y, w, 0.08, T.accent_rgb, radius_pct=0)
    if hl:
        set_solid_alpha(hl, 22)

    return c


def two_col_layout(slide, T, left_items, right_items,
                    CY, CH, left_label='', right_label='',
                    left_icon='', right_icon='', font='Cairo'):
    """
    Smart two-column layout — content-aware distribution.
    Automatically balances items between columns.
    """
    gap = 0.35
    col_w = (W - 2.0 - gap) / 2

    for ci, (label, icon, items) in enumerate([
        (left_label, left_icon, left_items),
        (right_label, right_icon, right_items)
    ]):
        x = 1.0 + ci * (col_w + gap)

        # Column card
        cc = content_card_premium(slide, x, CY, col_w, CH, T,
                                   accent_side='right' if ci == 0 else 'left')

        # Column header
        hh = 0.78
        hdr = rrect(slide, x, CY, col_w, hh, T.accent_rgb, radius_pct=0)
        if hdr:
            multi_stop_gradient(hdr, [(0, T.accent2), (50, T.accent), (100, T.accent2)], 0)

        icon_offset = 0.0
        if icon:
            txt(slide, icon, x + col_w - 1.1, CY, 0.9, hh,
                font='Calibri', size=18, bold=False, color=T.text_dark_rgb,
                align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
            icon_offset = 0.9

        if label:
            txt(slide, label, x + 0.2, CY, col_w - 0.3 - icon_offset, hh,
                font=font, size=17, bold=True, color=T.text_dark_rgb,
                align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        # Items
        ia = CH - hh - 0.1
        n_items = min(len(items), 8)
        ig = 0.1
        ih = (ia - ig * max(0, n_items - 1)) / max(1, n_items)

        for j, item in enumerate(items[:8]):
            iy = CY + hh + 0.05 + j * (ih + ig)
            rb = rrect(slide, x + 0.1, iy, col_w - 0.2, ih,
                       T.bg2_rgb if j % 2 == 0 else T.bg_rgb, radius_pct=6)
            if rb:
                set_solid_alpha(rb, 65)

            # Number badge
            nd = min(0.50, ih * 0.68)
            nc = oval(slide, x + col_w - 0.14 - nd - 0.22, iy + (ih - nd) / 2, nd, nd, T.accent_rgb)
            if nc:
                multi_stop_gradient(nc, [(0, T.accent), (100, T.accent2)], 135)
                shadow(nc, blur=5, dist=1, alpha=0.28)
            txt(slide, str(j + 1),
                x + col_w - 0.14 - nd - 0.22, iy + (ih - nd) / 2, nd, nd,
                font='Calibri', size=max(7, int(nd * 11)), bold=True,
                color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

            # Content-aware font size
            char_c = len(str(item))
            fs = max(11, min(14, ih * 8.5))
            if char_c > 80:
                fs = max(10, fs - 1)

            txt(slide, item, x + 0.22, iy, col_w - nd - 0.62, ih,
                font=font, size=fs, bold=False, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.2)


def methodology_card(slide, x, y, w, h, label: str, value: str, icon: str, T):
    """
    Methodology card — content-aware layout.
    Adjusts icon/text balance based on text length.
    """
    # Card base
    cc = rrect(slide, x, y, w, h, T.card_rgb, radius_pct=14)
    if cc:
        multi_stop_gradient(cc, [(0, T.card), (100, T.bg2)], 145)
        shadow(cc, blur=16, dist=5, alpha=0.42)

    # Icon section height = 38% of card, but less if text is long
    text_len = len(str(value))
    icon_ratio = 0.32 if text_len > 80 else 0.38 if text_len > 40 else 0.44

    ic_s = min(1.85, h * icon_ratio)
    ic_x = x + w / 2 - ic_s / 2
    ic_y = y + max(0.22, (h * icon_ratio - ic_s) / 2 + 0.15)

    # Icon circle
    ic_bg = oval(slide, ic_x, ic_y, ic_s, ic_s, T.accent_rgb)
    if ic_bg:
        multi_stop_gradient(ic_bg, [(0, T.accent), (100, T.accent2)], 135)
        shadow(ic_bg, blur=10, dist=3, alpha=0.32)
        glow(ic_bg, T.accent.lstrip('#'), radius=14, alpha=0.18)
    txt(slide, icon, ic_x, ic_y + 0.04, ic_s, ic_s * 0.9,
        font='Calibri', size=max(14, int(ic_s * 10)), bold=False,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

    # Label
    lbl_y = ic_y + ic_s + 0.22
    txt(slide, label, x + 0.18, lbl_y, w - 0.36, 0.64,
        font='Cairo', size=17, bold=True, color=T.accent_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    # Divider
    hline(slide, x + w * 0.15, lbl_y + 0.66, w * 0.7, T.muted_rgb, thickness=0.04)

    # Value text — content-aware
    val_y = lbl_y + 0.76
    val_h = h - (val_y - y) - 0.3
    fs = max(11, min(14, val_h * 5.5))
    if text_len > 100:
        fs = max(10, fs - 1)
    txt(slide, value, x + 0.2, val_y, w - 0.4, val_h,
        font='Cairo', size=fs, bold=False, color=T.text_light_rgb,
        align=PP_ALIGN.CENTER, rtl=True, vcenter=True, line_spacing=1.28)
    return cc


def premium_cover_frame(slide, cx, cy, cw, ch, T):
    """
    Premium cover card frame — multi-layer with depth.
    Returns the main card shape.
    """
    # Deep shadow
    ds = rrect(slide, cx + 0.28, cy + 0.38, cw, ch, T.bg_rgb, radius_pct=16)
    if ds:
        set_solid_alpha(ds, 28)

    # Medium shadow
    ms = rrect(slide, cx + 0.14, cy + 0.2, cw, ch, T.bg_rgb, radius_pct=16)
    if ms:
        set_solid_alpha(ms, 22)

    # Main card
    mc = rrect(slide, cx, cy, cw, ch, T.card_rgb, radius_pct=16)
    if mc:
        multi_stop_gradient(mc, [(0, T.card), (50, T.bg2), (100, T.card)], 135)
        shadow(mc, blur=32, dist=10, alpha=0.55)
        glow(mc, T.accent.lstrip('#'), radius=42, alpha=0.11)

    # Top accent bar
    ct = rrect(slide, cx, cy, cw, 0.42, T.accent_rgb, radius_pct=0)
    if ct:
        multi_stop_gradient(ct, [(0, T.accent), (50, T.accent2), (100, T.accent)], 0)
        glow(ct, T.accent.lstrip('#'), radius=18, alpha=0.35)

    # Bottom accent bar
    cb = rrect(slide, cx, cy + ch - 0.28, cw, 0.28, T.accent_rgb, radius_pct=0)
    if cb:
        set_solid_alpha(cb, 45)

    # Right accent edge
    accent_bar_v(slide, cx + cw - 0.26, cy + 0.42, ch - 0.7, T, width=0.26)

    return mc


def smart_header(slide, T, title: str, sub: str = '', slide_num=None,
                  total=13, style='gradient', font='Cairo'):
    """
    Smart header v28 — content-aware title sizing.
    Automatically adjusts font size based on title length.
    style: 'gradient' | 'split' | 'minimal'
    """
    HEADER_H = 2.9

    if style == 'gradient':
        gradient_rect(slide, 0, 0, W, HEADER_H, T.grad2, T.grad1, angle=135)
        # Bottom accent line
        al = rect(slide, 0, HEADER_H - 0.22, W, 0.22, T.accent_rgb)
        if al:
            multi_stop_gradient(al, [(0, T.bg), (40, T.accent),
                                     (60, T.accent2), (100, T.bg)], 0)
        # Right accent bar
        av = rect(slide, W - 0.52, 0, 0.52, HEADER_H, T.accent_rgb)
        if av:
            gradient_fill(av, T.accent_grad1, T.accent_grad2, 90)
            glow(av, T.accent.lstrip('#'), radius=12, alpha=0.22)

        # Decorative orb
        oval(slide, W - 6, -2.5, 8, 8, T.accent_rgb, alpha=7)

    elif style == 'split':
        # Upper half dark gradient
        gradient_rect(slide, 0, 0, W, HEADER_H * 0.65, T.grad1, T.grad2, angle=0)
        # Lower transition
        rect(slide, 0, HEADER_H * 0.65, W, HEADER_H * 0.35, T.bg_rgb)
        al = rect(slide, 0, HEADER_H - 0.18, W, 0.18, T.accent_rgb)
        if al:
            multi_stop_gradient(al, [(0, T.bg), (50, T.accent), (100, T.bg)], 0)
        av = rect(slide, W - 0.42, 0, 0.42, HEADER_H * 0.65, T.accent_rgb)
        if av:
            gradient_fill(av, T.accent_grad1, T.accent_grad2, 90)

    else:  # minimal
        gradient_rect(slide, 0, 0, W, HEADER_H, T.grad2, T.grad1, angle=160)
        hline(slide, 0, HEADER_H - 0.16, W, T.accent_rgb, thickness=0.16)

    # Slide number badge
    if slide_num is not None:
        nb_s = 0.74
        nb_x = 1.1
        nb_y = (HEADER_H - nb_s) / 2
        nb = oval(slide, nb_x, nb_y, nb_s, nb_s, T.accent_rgb)
        if nb:
            gradient_fill(nb, T.accent_grad1, T.accent_grad2, 135)
            shadow(nb, blur=8, dist=2, alpha=0.38)
        txt(slide, str(slide_num), nb_x, nb_y, nb_s, nb_s,
            font='Calibri', size=14, bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        txt(slide, f'/{total}', nb_x + nb_s, nb_y + nb_s * 0.32, 0.8, nb_s * 0.36,
            font='Calibri', size=8, bold=False, color=T.muted_rgb,
            align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
        title_x = nb_x + nb_s + 0.95
    else:
        title_x = 0.7

    # Content-aware title sizing
    title_len = len(title)
    if title_len <= 15:
        title_fs = 32
    elif title_len <= 25:
        title_fs = 30
    elif title_len <= 35:
        title_fs = 27
    else:
        title_fs = 24

    title_w = W - title_x - 0.65
    title_area_h = HEADER_H * (0.60 if sub else 0.82)
    title_y = (HEADER_H - title_area_h - (0.0 if not sub else 0.0)) * 0.35

    txt(slide, title, title_x, title_y, title_w, title_area_h,
        font=font, size=title_fs, bold=True, color=T.text_light_rgb,
        align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.08)

    if sub:
        sub_y = title_y + title_area_h
        sub_h = HEADER_H - sub_y - 0.28
        txt(slide, sub, title_x, sub_y, title_w, sub_h,
            font=font, size=13, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.0)

    return HEADER_H


def decorative_corner(slide, T, corner='tr', size=3.5):
    """Decorative geometric corner element."""
    if corner == 'tr':   # top-right
        oval(slide, W - size * 0.7, -size * 0.3, size, size, T.accent_rgb, alpha=6)
        oval(slide, W - size * 0.4, -size * 0.1, size * 0.55, size * 0.55, T.bg2_rgb, alpha=38)
    elif corner == 'bl':  # bottom-left
        oval(slide, -size * 0.3, H - size * 0.7, size, size, T.accent_rgb, alpha=5)
        oval(slide, -size * 0.1, H - size * 0.4, size * 0.5, size * 0.5, T.bg2_rgb, alpha=32)
    elif corner == 'tl':
        oval(slide, -size * 0.4, -size * 0.3, size, size, T.accent_rgb, alpha=5)
    elif corner == 'br':
        oval(slide, W - size * 0.6, H - size * 0.6, size, size, T.accent_rgb, alpha=5)
