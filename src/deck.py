"""
deck.py — a small corporate presentation design system built on python-pptx.

Everything the five unit decks need lives here: palette, type scale, grid and a
set of reusable slide layouts (title, agenda, section, bullets, cards, compare,
process, quadrant, stats, table, steps, closing).

Design intent: clean consulting-style slides, lots of white space, one idea per
slide, short plain-English copy.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from copy import deepcopy
import re

# --------------------------------------------------------------------------
# Palette
# --------------------------------------------------------------------------
NAVY       = RGBColor(0x0E, 0x1E, 0x3C)   # primary dark
NAVY_DEEP  = RGBColor(0x07, 0x12, 0x28)   # gradient end
NAVY_SOFT  = RGBColor(0x1B, 0x33, 0x5E)   # panels on dark
BLUE       = RGBColor(0x3B, 0x7A, 0xF7)   # primary accent
BLUE_LIGHT = RGBColor(0xE8, 0xF0, 0xFE)   # tinted panel
TEAL       = RGBColor(0x00, 0xB3, 0xA4)
TEAL_LIGHT = RGBColor(0xE0, 0xF7, 0xF4)
AMBER      = RGBColor(0xF5, 0xA6, 0x23)
AMBER_LIGHT= RGBColor(0xFE, 0xF4, 0xE2)
RED        = RGBColor(0xE0, 0x51, 0x51)
RED_LIGHT  = RGBColor(0xFD, 0xED, 0xED)
GREEN      = RGBColor(0x27, 0xA0, 0x6A)
GREEN_LIGHT= RGBColor(0xE7, 0xF5, 0xEF)
VIOLET     = RGBColor(0x74, 0x5C, 0xE0)
VIOLET_LIGHT = RGBColor(0xEE, 0xEB, 0xFC)

WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
PAPER      = RGBColor(0xF6, 0xF8, 0xFB)   # light panel
BORDER     = RGBColor(0xDD, 0xE3, 0xEC)
INK        = RGBColor(0x14, 0x1C, 0x2E)   # body text
MUTED      = RGBColor(0x63, 0x6E, 0x85)   # secondary text
MUTED_LT   = RGBColor(0x93, 0x9E, 0xB3)

ACCENTS = [BLUE, TEAL, AMBER, VIOLET, GREEN, RED]
ACCENT_TINTS = {
    str(BLUE): BLUE_LIGHT, str(TEAL): TEAL_LIGHT, str(AMBER): AMBER_LIGHT,
    str(VIOLET): VIOLET_LIGHT, str(GREEN): GREEN_LIGHT, str(RED): RED_LIGHT,
}

FONT = "Segoe UI"
FONT_LIGHT = "Segoe UI Light"

# --------------------------------------------------------------------------
# Grid (16:9 — 13.333in x 7.5in)
# --------------------------------------------------------------------------
SW, SH = 13.333, 7.5
ML, MR = 0.78, 0.78                 # side margins
CW = SW - ML - MR                   # content width  = 11.773
BODY_TOP = 1.78                     # first row of content on a normal slide
BODY_BOTTOM = 6.62                  # content must end above this
FOOTER_Y = 6.86


def _i(v):
    return Inches(v) if not isinstance(v, (Emu,)) else v


# --------------------------------------------------------------------------
# Low-level shape helpers
# --------------------------------------------------------------------------
def rect(slide, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE,
         radius=None, shadow=False):
    s = slide.shapes.add_shape(shape, _i(x), _i(y), _i(w), _i(h))
    s.shadow.inherit = False
    if radius is not None and shape in (MSO_SHAPE.ROUNDED_RECTANGLE,):
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    if shadow:
        soft_shadow(s)
    s.text_frame.word_wrap = True
    return s


def card(slide, x, y, w, h, fill=WHITE, line=BORDER, radius=0.06, shadow=True):
    return rect(slide, x, y, w, h, fill=fill, line=line, lw=1.0,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius, shadow=shadow)


def find_logo(assets_dir=None):
    """First image found in assets/ - the college logo. None if not supplied."""
    import glob
    import os
    if assets_dir is None:
        assets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  "..", "assets")
    for ext in ("png", "PNG", "jpg", "JPG", "jpeg", "JPEG", "webp", "gif"):
        hits = sorted(glob.glob(os.path.join(assets_dir, f"*.{ext}")))
        if hits:
            return hits[0]
    return None


def logo_size(path, h, max_w=3.0):
    """Width and height the logo will occupy at target height `h`."""
    if not path:
        return 0.0, 0.0
    try:
        from PIL import Image
        pw, ph = Image.open(path).size
        w = h * (pw / float(ph))
    except Exception:
        w = h
    if w > max_w:                      # very wide logo → constrain by width
        h = h * (max_w / w)
        w = max_w
    return w, h


def place_logo(slide, path, x, y, h, on_dark=False, pad=0.13, max_w=3.0,
               right_edge=None):
    """Drop the logo at a given height; on dark slides sit it on a white chip."""
    if not path:
        return 0.0
    w, h = logo_size(path, h, max_w)
    if right_edge is not None:
        x = right_edge - w
    if on_dark:
        plate = rect(slide, x - pad, y - pad, w + pad * 2, h + pad * 2, fill=WHITE,
                     line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.14)
        plate.name = "logo-plate"
    pic = slide.shapes.add_picture(path, _i(x), _i(y), height=_i(h))
    pic.name = "college-logo"
    return w


def chip_width(text, size=9.5, bold=True, spacing=1.4, pad=0.36):
    """Pill width that actually fits its label, including letter-spacing."""
    return (text_width_in(text, size, bold)
            + (spacing / 72.0) * max(len(text) - 1, 0) + pad)


def circle(slide, cx, cy, d, fill=BLUE, line=None):
    return rect(slide, cx - d / 2, cy - d / 2, d, d, fill=fill, line=line,
                shape=MSO_SHAPE.OVAL)


def hline(slide, x, y, w, color=BORDER, weight=1.0):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _i(x), _i(y), _i(w), Pt(weight))
    ln.shadow.inherit = False
    ln.fill.solid()
    ln.fill.fore_color.rgb = color
    ln.line.fill.background()
    return ln


def soft_shadow(shape, blur=16, dist=3, alpha=11000):
    """Subtle drop shadow so cards lift off the page."""
    spPr = shape._element.spPr
    for tag in ("a:effectLst",):
        for e in spPr.findall(qn(tag)):
            spPr.remove(e)
    xml = (
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:outerShdw blurRad="{blur * 12700}" dist="{dist * 12700}" dir="5400000" '
        'rotWithShape="0"><a:srgbClr val="0E1E3C">'
        f'<a:alpha val="{alpha}"/></a:srgbClr></a:outerShdw></a:effectLst>'
    )
    from pptx.oxml import parse_xml
    spPr.append(parse_xml(xml))


def gradient_bg(slide, c1=NAVY, c2=NAVY_DEEP, angle=45.0):
    s = rect(slide, 0, 0, SW, SH, fill=c1)
    s.fill.gradient()
    stops = s.fill.gradient_stops
    stops[0].color.rgb = c1
    stops[0].position = 0.0
    stops[1].color.rgb = c2
    stops[1].position = 1.0
    s.fill.gradient_angle = angle
    s.line.fill.background()
    return s


def ghost(slide, x, y, w, h, color=WHITE, alpha=6000, shape=MSO_SHAPE.OVAL):
    """Very faint decorative shape used on dark slides."""
    s = slide.shapes.add_shape(shape, _i(x), _i(y), _i(w), _i(h))
    s.name = "deco-bleed"
    s.shadow.inherit = False
    s.line.fill.background()
    from pptx.oxml import parse_xml
    spPr = s._element.spPr
    for e in spPr.findall(qn('a:solidFill')):
        spPr.remove(e)
    xml = ('<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           f'<a:srgbClr val="{color}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>')
    # insert solidFill right after the geometry element
    geom = spPr.find(qn('a:prstGeom'))
    geom.addnext(parse_xml(xml))
    return s


# --------------------------------------------------------------------------
# Text helpers
# --------------------------------------------------------------------------
def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(_i(x), _i(y), _i(w), _i(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return tf


def style_run(run, size=14, bold=False, color=INK, font=FONT, italic=False,
              spacing=None, caps=False):
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    if spacing:
        rPr.set('spc', str(int(spacing * 100)))
    if caps:
        rPr.set('cap', 'all')
    # make east-asian / complex-script fonts match so nothing falls back to Calibri
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            from pptx.oxml import parse_xml
            el = parse_xml(f'<{tag} xmlns:a="http://schemas.openxmlformats.org/'
                           f'drawingml/2006/main" typeface="{font}"/>')
            rPr.append(el)
        else:
            el.set('typeface', font)
    return run


def para(tf, text, size=14, bold=False, color=INK, font=FONT, align=PP_ALIGN.LEFT,
         space_before=0, space_after=0, line=1.22, italic=False, spacing=None,
         caps=False, first=False, bullet=None, indent=0.0):
    p = tf.paragraphs[0] if (first or (len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs
                                      and not getattr(tf, "_used", False))) else tf.add_paragraph()
    tf._used = True
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line
    if indent:
        pPr = p._p.get_or_add_pPr()
        pPr.set('marL', str(int(indent * 914400)))
        pPr.set('indent', str(-int(0.18 * 914400)))
    if bullet:
        pPr = p._p.get_or_add_pPr()
        from pptx.oxml import parse_xml
        pPr.append(parse_xml(
            '<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' typeface="Arial"/>'))
        pPr.append(parse_xml(
            '<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            f' char="{bullet}"/>'))
    from pptx.oxml import parse_xml
    chunks = text.split("\n")
    for j, chunk in enumerate(chunks):
        if j:
            p._p.append(parse_xml(
                '<a:br xmlns:a="http://schemas.openxmlformats.org/drawingml/'
                '2006/main"/>'))
        r = p.add_run()
        r.text = chunk
        style_run(r, size=size, bold=bold, color=color, font=font, italic=italic,
                  spacing=spacing, caps=caps)
    return p


def rich(tf, parts, size=14, color=INK, align=PP_ALIGN.LEFT, line=1.22,
         space_before=0, space_after=0, first=False):
    """parts = [(text, {overrides}), ...] inside one paragraph."""
    p = tf.paragraphs[0] if (first or not getattr(tf, "_used", False)) else tf.add_paragraph()
    tf._used = True
    p.alignment = align
    p.line_spacing = line
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    for text, ov in parts:
        r = p.add_run()
        r.text = text
        style_run(r, size=ov.get('size', size), bold=ov.get('bold', False),
                  color=ov.get('color', color), font=ov.get('font', FONT),
                  italic=ov.get('italic', False), spacing=ov.get('spacing'),
                  caps=ov.get('caps', False))
    return p


# ---- real text measurement (Selawik is metric-compatible with Segoe UI) ----
_FONT_PATHS = {
    False: "/usr/share/fonts/selawik/selawk.ttf",
    True: "/usr/share/fonts/selawik/selawkb.ttf",
}
_PIL_CACHE = {}
_SCALE = 8  # render font at 8x for sub-point precision


def _pil_font(size, bold):
    key = (round(float(size), 2), bool(bold))
    if key not in _PIL_CACHE:
        from PIL import ImageFont
        _PIL_CACHE[key] = ImageFont.truetype(_FONT_PATHS[bool(bold)],
                                             max(1, int(round(size * _SCALE))))
    return _PIL_CACHE[key]


def text_width_in(text, size, bold=False):
    """Width of `text` in inches at `size` points."""
    if not text:
        return 0.0
    try:
        f = _pil_font(size, bold)
        return (f.getlength(text) / _SCALE) / 72.0
    except Exception:                                    # font missing → estimate
        return len(text) * size * 0.50 / 72.0


def fit(text, size, width_in, _legacy=None, bold=True):
    """Number of wrapped lines `text` needs inside `width_in`."""
    total = 0
    for hard in str(text).replace("\v", "\n").replace("\x0b", "\n").split("\n"):
        if not hard.strip():
            total += 1
            continue
        lines, cur = 1, ""
        for w in hard.split():
            trial = (cur + " " + w).strip()
            if cur and text_width_in(trial, size, bold) > width_in:
                lines += 1
                cur = w
            else:
                cur = trial
        total += lines
    return max(total, 1)


# Segoe UI / Selawik natural line height as a multiple of the em size.
# PowerPoint percentage line spacing scales THIS, not the point size.
LINE_FACTOR = 1.201


def line_height_in(size, line=1.25):
    return size * LINE_FACTOR * line / 72.0


def text_height_in(text, size, width_in, bold=False, line=1.25):
    """Height the wrapped text will occupy, in inches."""
    return fit(text, size, width_in, bold=bold) * line_height_in(size, line)


def autosize(text, width_in, height_in, size, bold=False, line=1.25,
             min_size=8.5, step=0.5):
    """Largest font size <= `size` whose wrapped text fits the given box."""
    s = float(size)
    while s > min_size and text_height_in(text, s, width_in, bold, line) > height_in:
        s -= step
    return s


# ==========================================================================
# Deck
# ==========================================================================
class Deck:
    # presenter details, shared by every deck
    PRESENTER = "Lingesh K"
    REGISTER = "OSI2509030"

    def __init__(self, unit_no, unit_title, course="PGPM · Systems Specialisation",
                 subject="Artificial Intelligence & Generative AI", logo=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(SW)
        self.prs.slide_height = Inches(SH)
        self.unit_no = unit_no
        self.unit_title = unit_title
        self.course = course
        self.subject = subject
        self.logo = logo if logo is not None else find_logo()
        self.blank = self.prs.slide_layouts[6]
        self._sections = []

        core = self.prs.core_properties
        core.title = f"Unit {unit_no} — {unit_title}"
        core.author = f"{self.PRESENTER} ({self.REGISTER})"
        core.subject = subject
        core.comments = f"{course} · Unit {unit_no} of 5"

    # ---------------- infrastructure ----------------
    def _new(self, footer=True, bg=None):
        s = self.prs.slides.add_slide(self.blank)
        if bg is not None:
            rect(s, 0, 0, SW, SH, fill=bg)
        if footer:
            self._footer(s)
        return s

    def _footer(self, slide):
        hline(slide, ML, FOOTER_Y - 0.16, CW, BORDER, 0.75)
        right = ML + CW
        logo_w = 0.0
        if self.logo:
            logo_w = place_logo(slide, self.logo, 0, FOOTER_Y - 0.06, 0.32,
                                on_dark=False, max_w=0.9,
                                right_edge=right) + 0.26

        tf = textbox(slide, ML, FOOTER_Y, CW * 0.52, 0.24)
        rich(tf, [(f"Unit {self.unit_no}", {'bold': True, 'color': NAVY}),
                  ("  ·  ", {'color': MUTED_LT}),
                  (self.unit_title, {'color': MUTED})],
             size=9, first=True)

        n = len(self.prs.slides._sldIdLst) - 1  # index of the slide being built
        credit = f"{self.PRESENTER}  ·  {self.REGISTER}"
        block_w = 2.9
        x = right - logo_w - block_w
        tf2 = textbox(slide, x, FOOTER_Y, block_w, 0.24)
        rich(tf2, [(credit, {'color': MUTED}),
                   ("     ", {}),
                   (f"{n + 1:02d}", {'bold': True, 'color': NAVY})],
             size=9, align=PP_ALIGN.RIGHT, first=True)

    def _head(self, slide, kicker, title, sub=None, accent=BLUE):
        """Standard slide header: small label, big title, accent rule."""
        y = 0.62
        if kicker:
            tf = textbox(slide, ML, y, CW, 0.24)
            para(tf, kicker.upper(), size=9.5, bold=True, color=accent,
                 spacing=1.6, first=True)
            y += 0.33
        tsize = 28.0
        while tsize > 20 and fit(title, tsize, CW, bold=True) > 1:
            tsize -= 1.0
        th = text_height_in(title, tsize, CW, True, 1.06)
        tf = textbox(slide, ML, y, CW, th + 0.12)
        para(tf, title, size=tsize, bold=True, color=NAVY, line=1.06, first=True)
        y += th + 0.1
        rect(slide, ML, y, 0.52, 0.045, fill=accent)
        y += 0.045
        if sub:
            sw = CW * 0.86
            tf = textbox(slide, ML, y + 0.18, sw, 0.5)
            para(tf, sub, size=12.5, color=MUTED, line=1.3, first=True)
            y += 0.18 + text_height_in(sub, 12.5, sw, False, 1.3)
        return y

    def save(self, path):
        self.prs.save(path)
        return path

    # ================= LAYOUTS =================

    def title_slide(self, title, subtitle, meta_lines=None, chips=None):
        s = self._new(footer=False)
        gradient_bg(s, NAVY, NAVY_DEEP, 315.0)
        # decorative geometry
        ghost(s, 9.5, -1.6, 5.6, 5.6, "3B7AF7", 9000)
        ghost(s, 11.0, 3.4, 4.2, 4.2, "00B3A4", 7000)
        ghost(s, 8.0, 1.1, 3.0, 3.0, "FFFFFF", 3500)
        rect(s, 0, 0, 0.16, SH, fill=BLUE)

        if self.logo:
            place_logo(s, self.logo, 0, 0.72, 1.0, on_dark=True, max_w=2.5,
                       right_edge=ML + CW - 0.1)

        tf = textbox(s, ML + 0.24, 1.28, 7.9, 0.3)
        para(tf, self.course.upper(), size=10.5, bold=True, color=TEAL,
             spacing=2.2, first=True)

        # unit badge
        b = rect(s, ML + 0.24, 1.78, 1.62, 0.42, fill=None, line=BLUE, lw=1.25,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        tfb = b.text_frame
        tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tfb, f"UNIT {self.unit_no:02d}", size=10.5, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, spacing=1.6, first=True)

        tw = 8.1
        tsize = 44.0
        while tsize > 32 and fit(title, tsize, tw, bold=True) > 2:
            tsize -= 2.0
        thh = text_height_in(title, tsize, tw, True, 1.04)
        tf = textbox(s, ML + 0.24, 2.52, tw, thh + 0.2)
        para(tf, title, size=tsize, bold=True, color=WHITE, line=1.04, first=True)

        y = 2.52 + thh + 0.22
        rect(s, ML + 0.24, y, 0.72, 0.05, fill=TEAL)
        sw = 7.4
        ssize = autosize(subtitle, sw, 1.0, 15, False, 1.35, min_size=12)
        tf = textbox(s, ML + 0.24, y + 0.34, sw, 1.05)
        para(tf, subtitle, size=ssize, color=RGBColor(0xC5, 0xD2, 0xE8), line=1.35,
             first=True)

        sub_bottom = y + 0.34 + text_height_in(subtitle, ssize, sw, False, 1.35)
        if chips:
            cx = ML + 0.24
            chip_y = min(max(5.34, sub_bottom + 0.34), 5.46)
            for i, c in enumerate(chips):
                w = chip_width(c, 9.5, True, 0.0, 0.42)
                ch = rect(s, cx, chip_y, w, 0.36, fill=None, line=NAVY_SOFT, lw=1.0,
                          shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
                ch.line.color.rgb = RGBColor(0x33, 0x4B, 0x78)
                t = ch.text_frame
                t.vertical_anchor = MSO_ANCHOR.MIDDLE
                para(t, c, size=9.5, bold=True, color=RGBColor(0x9F, 0xB4, 0xD8),
                     align=PP_ALIGN.CENTER, first=True)
                cx += w + 0.14

        # presenter block, always shown
        hline(s, ML + 0.24, 6.14, 4.2, RGBColor(0x2A, 0x3E, 0x66), 1.0)
        tf = textbox(s, ML + 0.24, 6.34, 5.4, 0.7)
        para(tf, "PRESENTED BY", size=8.5, bold=True, color=TEAL, spacing=1.6,
             first=True)
        rich(tf, [(self.PRESENTER, {'bold': True, 'color': WHITE, 'size': 13}),
                  ("      Register No. ", {'color': RGBColor(0x7E, 0x91, 0xB4),
                                           'size': 10.5}),
                  (self.REGISTER, {'bold': True,
                                   'color': RGBColor(0xC5, 0xD2, 0xE8),
                                   'size': 10.5})],
             space_before=4)
        if meta_lines:
            tf2 = textbox(s, ML + CW - 4.7, 6.42, 4.6, 0.7)
            for i, m in enumerate(meta_lines):
                para(tf2, m, size=10, bold=(i == 0),
                     color=RGBColor(0xC5, 0xD2, 0xE8) if i == 0
                     else RGBColor(0x7E, 0x91, 0xB4),
                     line=1.34, align=PP_ALIGN.RIGHT, first=(i == 0))
        return s

    def agenda_slide(self, items, title="What we will cover",
                     kicker="Agenda", note=None):
        s = self._new()
        self._head(s, kicker, title, note)
        n = len(items)
        col = 2 if n > 4 else 1
        rows = (n + col - 1) // col
        gap = 0.26
        cw = (CW - gap) / col if col == 2 else CW * 0.78
        top = 2.18 if not note else 2.42
        ch = min(0.86, (BODY_BOTTOM - top - (rows - 1) * 0.16) / rows)
        for i, it in enumerate(items):
            head, desc = (it if isinstance(it, tuple) else (it, None))
            r, c = i % rows, i // rows
            x = ML + c * (cw + gap)
            y = top + r * (ch + 0.16)
            accent = ACCENTS[i % len(ACCENTS)]
            rect(s, x, y, 0.035, ch, fill=accent)
            tfn = textbox(s, x + 0.22, y + 0.02, 0.62, ch)
            para(tfn, f"{i + 1:02d}", size=21, bold=True, color=accent, first=True)
            iw = cw - 1.0
            hsize = autosize(head, iw, ch * (0.5 if desc else 1.0), 14.5, True,
                             1.14, min_size=12)
            hh = text_height_in(head, hsize, iw, True, 1.14)
            tf = textbox(s, x + 0.86, y + (0.06 if desc else 0.14), iw, ch)
            para(tf, head, size=hsize, bold=True, color=NAVY, line=1.14, first=True)
            if desc:
                dsize = autosize(desc, iw, ch - hh - 0.14, 11, False, 1.22,
                                 min_size=9)
                para(tf, desc, size=dsize, color=MUTED, line=1.22, space_before=3)
        return s

    def section_slide(self, number, title, blurb=None, accent=BLUE):
        s = self._new(footer=False)
        gradient_bg(s, NAVY, NAVY_DEEP, 315.0)
        ghost(s, 10.2, 1.0, 4.6, 4.6, "3B7AF7", 8000)
        ghost(s, 12.0, 4.4, 3.0, 3.0, "00B3A4", 6000)
        rect(s, 0, 0, 0.16, SH, fill=accent)
        tf = textbox(s, ML + 0.3, 2.62, 1.9, 1.0)
        para(tf, f"{number:02d}", size=64, bold=True, color=accent, line=0.92,
             first=True)
        rect(s, ML + 0.3, 3.66, 0.66, 0.05, fill=WHITE)
        tw = 8.4
        tsize = autosize(title, tw, 1.7, 34, True, 1.08, min_size=26)
        th = text_height_in(title, tsize, tw, True, 1.08)
        tf = textbox(s, ML + 2.5, 2.72, tw, th + 0.16)
        para(tf, title, size=tsize, bold=True, color=WHITE, line=1.08, first=True)
        if blurb:
            tf = textbox(s, ML + 2.5, 2.72 + th + 0.24, 7.6, 0.9)
            para(tf, blurb, size=13.5, color=RGBColor(0xAF, 0xC1, 0xDE), line=1.34,
                 first=True)
        tf = textbox(s, ML + 0.3, 6.5, 7.0, 0.3)
        para(tf, f"UNIT {self.unit_no}  ·  {self.unit_title}".upper(), size=9,
             bold=True, color=RGBColor(0x6B, 0x7F, 0xA6), spacing=1.6, first=True)
        if self.logo:
            place_logo(s, self.logo, 0, 0.72, 0.82, on_dark=True, max_w=2.1,
                       right_edge=ML + CW - 0.1)
        return s

    def bullets_slide(self, kicker, title, bullets, sub=None, lead=None,
                      accent=BLUE, two_col=False):
        s = self._new()
        y = self._head(s, kicker, title, sub, accent)
        y += 0.34
        if lead:
            p = card(s, ML, y, CW, 0.72, fill=ACCENT_TINTS[str(accent)], line=None,
                     shadow=False)
            rect(s, ML, y, 0.035, 0.72, fill=accent)
            tf = p.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.28)
            tf.margin_right = Inches(0.24)
            lsize = autosize(lead, CW - 0.6, 0.58, 13, True, 1.25, min_size=10.5)
            para(tf, lead, size=lsize, bold=True, color=NAVY, line=1.25, first=True)
            y += 0.72 + 0.34

        cols = 2 if two_col else 1
        rows = (len(bullets) + cols - 1) // cols
        gap = 0.5
        cw = (CW - gap) / cols if cols == 2 else CW * 0.94
        avail = BODY_BOTTOM - y
        rh = min(0.94, avail / max(rows, 1))
        y += max(0.0, (avail - rh * rows) / 2)
        for i, b in enumerate(bullets):
            head, desc = (b if isinstance(b, tuple) else (b, None))
            r, c = (i % rows, i // rows) if cols == 2 else (i, 0)
            x = ML + c * (cw + gap)
            yy = y + r * rh
            rect(s, x + 0.02, yy + 0.145, 0.13, 0.13, fill=accent,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
            iw = cw - 0.5
            budget = rh - 0.14
            hsize = autosize(head, iw, budget * (0.5 if desc else 1.0), 14,
                             bool(desc), 1.2, min_size=11)
            hh = text_height_in(head, hsize, iw, bool(desc), 1.2)
            tf = textbox(s, x + 0.4, yy + 0.04, iw, budget)
            para(tf, head, size=hsize, bold=bool(desc), color=NAVY if desc else INK,
                 line=1.2, first=True)
            if desc:
                dsize = autosize(desc, iw, budget - hh - 0.06, 11.5, False, 1.26,
                                 min_size=9)
                para(tf, desc, size=dsize, color=MUTED, line=1.26, space_before=3)
        return s

    def cards_slide(self, kicker, title, cards_data, sub=None, cols=None,
                    accent=BLUE, numbered=False, tinted=False):
        """cards_data = [(heading, body), ...] or [(heading, body, accent), ...]"""
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.36
        n = len(cards_data)
        cols = cols or (3 if n % 3 == 0 or n > 4 else (2 if n <= 4 else 3))
        rows = (n + cols - 1) // cols
        gx, gy = 0.28, 0.26
        cw = (CW - gx * (cols - 1)) / cols
        ch = (BODY_BOTTOM - y - gy * (rows - 1)) / rows
        # shrink cards to the height their content actually needs
        pad0 = 0.3
        iw0 = cw - pad0 * 2
        need = max(
            0.36 + (0.56 if numbered else 0.0)
            + text_height_in(c[0], 14.5, iw0, True, 1.14) + 0.16
            + text_height_in(c[1], 11.5, iw0, False, 1.32) + 0.34
            for c in cards_data)
        ch = min(ch, max(need, 2.25 if rows == 1 else 1.5))
        y += max(0.0, (BODY_BOTTOM - y - (ch * rows + gy * (rows - 1))) / 2)
        for i, c in enumerate(cards_data):
            head, body = c[0], c[1]
            acc = c[2] if len(c) > 2 else ACCENTS[i % 3] if tinted else accent
            r, cc = i // cols, i % cols
            x = ML + cc * (cw + gx)
            yy = y + r * (ch + gy)
            bg = ACCENT_TINTS[str(acc)] if tinted else WHITE
            card(s, x, yy, cw, ch, fill=bg, line=None if tinted else BORDER,
                 shadow=not tinted)
            rect(s, x + 0.001, yy + 0.001, cw * 0.999, 0.07, fill=acc)
            pad = 0.3
            contenth = ((0.56 if numbered else 0.0)
                        + text_height_in(head, 14.5, cw - pad * 2, True, 1.14)
                        + 0.16
                        + text_height_in(body, 11.5, cw - pad * 2, False, 1.32))
            ty = yy + max(0.36, (ch - contenth) / 2 + 0.03)
            if numbered:
                circle(s, x + pad + 0.19, ty + 0.19, 0.38, fill=acc)
                tfn = textbox(s, x + pad, ty + 0.045, 0.38, 0.3)
                para(tfn, f"{i + 1}", size=13, bold=True, color=WHITE,
                     align=PP_ALIGN.CENTER, first=True)
                ty += 0.56
            iw = cw - pad * 2
            hsize = autosize(head, iw, 0.78, 14.5, True, 1.14, min_size=12)
            hl = text_height_in(head, hsize, iw, True, 1.14)
            tf = textbox(s, x + pad, ty, iw, hl + 0.1)
            para(tf, head, size=hsize, bold=True, color=NAVY, line=1.14, first=True)
            bh = ch - (ty - yy) - hl - 0.34
            bsize = autosize(body, iw, bh, 11.5, False, 1.32, min_size=8.5)
            tf = textbox(s, x + pad, ty + hl + 0.16, iw, bh)
            para(tf, body, size=bsize, color=MUTED, line=1.32, first=True)
        return s

    def compare_slide(self, kicker, title, left, right, sub=None,
                      lacc=BLUE, racc=AMBER, verdict=None):
        """left/right = dict(label=, headline=, points=[..])"""
        s = self._new()
        y = self._head(s, kicker, title, sub, lacc) + 0.36
        bottom = BODY_BOTTOM - (0.78 if verdict else 0)
        gap = 0.34
        cw = (CW - gap) / 2
        for side, x, acc in ((left, ML, lacc), (right, ML + cw + gap, racc)):
            h = bottom - y
            card(s, x, y, cw, h, fill=WHITE, line=BORDER, shadow=True)
            rect(s, x + 0.001, y + 0.001, cw * 0.999, 0.07, fill=acc)
            pad = 0.34
            chip = rect(s, x + pad, y + 0.34,
                        chip_width(side['label'].upper(), 9, True, 1.2, 0.36), 0.34,
                        fill=ACCENT_TINTS[str(acc)],
                        shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
            t = chip.text_frame
            t.vertical_anchor = MSO_ANCHOR.MIDDLE
            para(t, side['label'].upper(), size=9, bold=True, color=acc,
                 align=PP_ALIGN.CENTER, spacing=1.2, first=True)
            iw = cw - pad * 2
            hdsize = autosize(side['headline'], iw, 0.92, 15, True, 1.14,
                              min_size=12.5)
            hh = text_height_in(side['headline'], hdsize, iw, True, 1.14)
            tf = textbox(s, x + pad, y + 0.86, iw, hh + 0.1)
            para(tf, side['headline'], size=hdsize, bold=True, color=NAVY,
                 line=1.14, first=True)
            yy = y + 0.86 + hh + 0.22
            hline(s, x + pad, yy, iw, BORDER, 0.75)
            yy += 0.24
            npts = max(len(side['points']), 1)
            rowh = min(0.66, (y + h - 0.26 - yy) / npts)
            pw = iw - 0.32
            psize = min([autosize(p, pw, rowh - 0.08, 12, False, 1.25, min_size=9.5)
                         for p in side['points']] or [12])
            for pnt in side['points']:
                rect(s, x + pad, yy + 0.115, 0.11, 0.11, fill=acc,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tfp = textbox(s, x + pad + 0.32, yy, pw, rowh)
                para(tfp, pnt, size=psize, color=INK, line=1.25, first=True)
                yy += rowh
        if verdict:
            vy = BODY_BOTTOM - 0.62
            b = card(s, ML, vy, CW, 0.62, fill=NAVY, line=None, shadow=False)
            tf = b.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.3)
            rich(tf, [("In short   ", {'bold': True, 'color': TEAL, 'size': 10,
                                      'spacing': 1.4, 'caps': True}),
                      (verdict, {'color': WHITE, 'size': 12.5, 'bold': True})],
                 first=True)
        return s

    def process_slide(self, kicker, title, steps, sub=None, accent=BLUE,
                      note=None):
        """steps = [(label, desc), ...]  → horizontal numbered flow."""
        s = self._new()
        base = accent[0] if isinstance(accent, list) else accent
        y = self._head(s, kicker, title, sub, base) + 0.5
        n = len(steps)
        gx = 0.2
        cw = (CW - gx * (n - 1)) / n
        h = min(2.9, BODY_BOTTOM - y - (0.86 if note else 0))
        iw0 = cw - 0.36
        cardh = min(h - 0.78, max(
            0.22 + text_height_in(l, 13, iw0, True, 1.14) + 0.14
            + text_height_in(dsc, 11, iw0, False, 1.28) + 0.26
            for l, dsc in steps))
        cy = y + 0.42 + max(0.0, (h - (0.78 + cardh)) / 2)
        hline(s, ML + cw * 0.5, cy - 0.012, CW - cw, BORDER, 2.0)
        for i, (label, desc) in enumerate(steps):
            x = ML + i * (cw + gx)
            acc = accent if not isinstance(accent, list) else accent[i % len(accent)]
            circle(s, x + cw / 2, cy, 0.52, fill=acc)
            tfn = textbox(s, x, cy - 0.145, cw, 0.3)
            para(tfn, f"{i + 1}", size=15, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, first=True)
            card(s, x, cy + 0.44, cw, cardh, fill=PAPER, line=BORDER, shadow=False)
            iw = cw - 0.36
            lsize = autosize(label, iw, 0.62, 13, True, 1.14, min_size=10.5)
            lh = text_height_in(label, lsize, iw, True, 1.14)
            tf = textbox(s, x + 0.18, cy + 0.66, iw, lh + 0.1)
            para(tf, label, size=lsize, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER, line=1.14, first=True)
            dh = cardh - lh - 0.52
            dsize = autosize(desc, iw, dh, 11, False, 1.28, min_size=8.5)
            tf = textbox(s, x + 0.18, cy + 0.66 + lh + 0.14, iw, dh)
            para(tf, desc, size=dsize, color=MUTED, align=PP_ALIGN.CENTER,
                 line=1.28, first=True)
        if note:
            ny = BODY_BOTTOM - 0.66
            b = card(s, ML, ny, CW, 0.66, fill=ACCENT_TINTS[str(base)],
                     line=None, shadow=False)
            rect(s, ML, ny, 0.035, 0.66, fill=base)
            tf = b.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.3)
            tf.margin_right = Inches(0.24)
            para(tf, note, size=12, bold=True, color=NAVY, first=True)
        return s

    def steps_slide(self, kicker, title, levels, sub=None, accent=BLUE):
        """Vertical staircase — good for maturity levels."""
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.36
        n = len(levels)
        gap = 0.14
        h = (BODY_BOTTOM - y - gap * (n - 1)) / n
        h = min(h, 1.05)
        step = (CW * 0.30) / max(n - 1, 1)
        for i, lv in enumerate(levels):
            label, desc = lv[0], lv[1]
            acc = lv[2] if len(lv) > 2 else ACCENTS[i % len(ACCENTS)]
            x = ML + i * step
            w = CW - i * step
            yy = y + (n - 1 - i) * (h + gap)
            card(s, x, yy, w, h, fill=WHITE, line=BORDER, shadow=True)
            rect(s, x, yy, 0.05, h, fill=acc)
            circle(s, x + 0.48, yy + h / 2, 0.44, fill=ACCENT_TINTS[str(acc)])
            tfn = textbox(s, x + 0.26, yy + h / 2 - 0.115, 0.44, 0.26)
            para(tfn, f"L{i + 1}", size=11.5, bold=True, color=acc,
                 align=PP_ALIGN.CENTER, first=True)
            lw2 = min(3.1, w - 1.1)
            lsize = autosize(label, lw2, h - 0.34, 13.5, True, 1.14, min_size=11)
            lh = text_height_in(label, lsize, lw2, True, 1.14)
            tf = textbox(s, x + 0.9, yy + max(0.14, (h - lh) / 2), lw2, lh + 0.1)
            para(tf, label, size=lsize, bold=True, color=NAVY, line=1.14, first=True)
            dw = w - 1.1 - lw2 - 0.4
            dsize = autosize(desc, dw, h - 0.34, 11.5, False, 1.26, min_size=9.5)
            dh = text_height_in(desc, dsize, dw, False, 1.26)
            tf = textbox(s, x + 0.9 + lw2 + 0.2, yy + max(0.14, (h - dh) / 2),
                         dw, dh + 0.1)
            para(tf, desc, size=dsize, color=MUTED, line=1.26, first=True)
        return s

    def stats_slide(self, kicker, title, stats, sub=None, accent=BLUE, note=None):
        """stats = [(big, label, desc), ...]"""
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.44
        n = len(stats)
        gx = 0.28
        cw = (CW - gx * (n - 1)) / n
        h = min(2.6, BODY_BOTTOM - y - (0.9 if note else 0))
        for i, (big, label, desc) in enumerate(stats):
            x = ML + i * (cw + gx)
            acc = ACCENTS[i % 3]
            card(s, x, y, cw, h, fill=WHITE, line=BORDER, shadow=True)
            bigsize = autosize(big, cw - 0.52, 0.84, 38, True, 1.0, min_size=17)
            bigh = text_height_in(big, bigsize, cw - 0.52, True, 1.0)
            tf = textbox(s, x + 0.26, y + 0.3 + max(0, 0.84 - bigh) / 2,
                         cw - 0.52, bigh + 0.1)
            para(tf, big, size=bigsize, bold=True, color=acc, line=1.0, first=True)
            rect(s, x + 0.26, y + 1.16, 0.42, 0.045, fill=acc)
            iw = cw - 0.52
            lsize = autosize(label, iw, 0.6, 13, True, 1.14, min_size=11)
            lh = text_height_in(label, lsize, iw, True, 1.14)
            tf = textbox(s, x + 0.26, y + 1.36, iw, lh + 0.08)
            para(tf, label, size=lsize, bold=True, color=NAVY, line=1.14, first=True)
            dh = h - 1.36 - lh - 0.3
            dsize = autosize(desc, iw, dh, 11, False, 1.3, min_size=8.5)
            tf = textbox(s, x + 0.26, y + 1.36 + lh + 0.12, iw, dh)
            para(tf, desc, size=dsize, color=MUTED, line=1.3, first=True)
        if note:
            ny = y + h + 0.3
            b = card(s, ML, ny, CW, 0.62, fill=NAVY, line=None, shadow=False)
            tf = b.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.3)
            para(tf, note, size=12.5, bold=True, color=WHITE, first=True)
        return s

    def quadrant_slide(self, kicker, title, quads, sub=None, accent=BLUE):
        """quads = [(heading, [points], accent), ...] — 2x2 grid."""
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.36
        gx, gy = 0.28, 0.24
        cw = (CW - gx) / 2
        ch = (BODY_BOTTOM - y - gy) / 2
        for i, q in enumerate(quads):
            head, pts = q[0], q[1]
            acc = q[2] if len(q) > 2 else ACCENTS[i % len(ACCENTS)]
            x = ML + (i % 2) * (cw + gx)
            yy = y + (i // 2) * (ch + gy)
            card(s, x, yy, cw, ch, fill=WHITE, line=BORDER, shadow=True)
            rect(s, x, yy, 0.05, ch, fill=acc)
            tf = textbox(s, x + 0.34, yy + 0.26, cw - 0.6, 0.36)
            para(tf, head, size=14, bold=True, color=NAVY, first=True)
            yy2 = yy + 0.74
            rowh = min(0.54, (yy + ch - 0.24 - yy2) / max(len(pts), 1))
            pw = cw - 0.96
            psize = min([autosize(p, pw, rowh - 0.06, 11.5, False, 1.24,
                                  min_size=9) for p in pts] or [11.5])
            for p in pts:
                rect(s, x + 0.34, yy2 + 0.1, 0.1, 0.1, fill=acc,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tfp = textbox(s, x + 0.62, yy2 - 0.01, pw, rowh)
                para(tfp, p, size=psize, color=INK, line=1.24, first=True)
                yy2 += rowh
        return s

    def table_slide(self, kicker, title, headers, rows, sub=None, accent=BLUE,
                    widths=None, note=None):
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.4
        ncol, nrow = len(headers), len(rows) + 1
        h = min(BODY_BOTTOM - y - (0.86 if note else 0), 0.52 + 0.62 * len(rows))
        gt = s.shapes.add_table(nrow, ncol, _i(ML), _i(y), _i(CW), _i(h))
        tbl = gt.table
        tbl.first_row = True
        # kill banding
        for tag in ('bandRow', 'bandCol', 'firstCol', 'lastRow', 'lastCol'):
            tbl._tbl.tblPr.set(tag, '0')
        tbl._tbl.tblPr.set('firstRow', '1')
        if widths:
            total = sum(widths)
            for i, w in enumerate(widths):
                tbl.columns[i].width = Emu(int(Inches(CW) * w / total))
        colw = [tbl.columns[i].width / 914400.0 for i in range(ncol)]
        rowh = max(0.42, (h - 0.5) / len(rows))
        tbl.rows[0].height = Inches(0.5)
        for r in range(1, nrow):
            tbl.rows[r].height = Inches(rowh)
        # shrink body text until the tallest cell in every row fits its row
        body_size = 11.0
        while body_size > 8.5:
            worst = 0.0
            for row in rows:
                for c, val in enumerate(row):
                    worst = max(worst, text_height_in(
                        val, body_size, max(colw[c] - 0.4, 0.5), c == 0, 1.24))
            if worst <= rowh - 0.18:
                break
            body_size -= 0.5
        head_size = 11.5
        while head_size > 9.0 and max(
                text_height_in(h_, head_size, max(colw[c] - 0.4, 0.5), True, 1.15)
                for c, h_ in enumerate(headers)) > 0.32:
            head_size -= 0.5
        for c, htxt in enumerate(headers):
            cell = tbl.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            cell.margin_left = cell.margin_right = Inches(0.18)
            cell.margin_top = cell.margin_bottom = Inches(0.1)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            para(tf, htxt, size=head_size, bold=True, color=WHITE, line=1.15,
                 first=True)
        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r % 2 else PAPER
                cell.margin_left = cell.margin_right = Inches(0.18)
                cell.margin_top = cell.margin_bottom = Inches(0.09)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                para(tf, val, size=body_size, bold=(c == 0),
                     color=NAVY if c == 0 else INK, line=1.24, first=True)
        if note:
            ny = y + h + 0.24
            b = card(s, ML, ny, CW, 0.62, fill=ACCENT_TINTS[str(accent)],
                     line=None, shadow=False)
            rect(s, ML, ny, 0.035, 0.62, fill=accent)
            tf = b.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.28)
            para(tf, note, size=12, bold=True, color=NAVY, first=True)
        return s

    def split_slide(self, kicker, title, left_head, left_points, right_cards,
                    sub=None, accent=BLUE):
        """Left = explanation, right = stacked mini cards."""
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.36
        lw = CW * 0.42
        rw = CW - lw - 0.4
        h = BODY_BOTTOM - y
        lhsize = autosize(left_head, lw, 1.4, 17, True, 1.14, min_size=14)
        lhh = text_height_in(left_head, lhsize, lw, True, 1.14)
        tf = textbox(s, ML, y + 0.04, lw, lhh + 0.14)
        para(tf, left_head, size=lhsize, bold=True, color=NAVY, line=1.14,
             first=True)
        yy = y + 0.06 + lhh + 0.22
        rect(s, ML, yy, 0.5, 0.045, fill=accent)
        yy += 0.3
        pw = lw - 0.32
        avail = (y + h) - yy
        gapp = 0.16
        psize = 12.0
        while psize > 9.5 and (sum(text_height_in(p, psize, pw, False, 1.26)
                                  for p in left_points)
                               + gapp * (len(left_points) - 1)) > avail:
            psize -= 0.5
        for p in left_points:
            ph = text_height_in(p, psize, pw, False, 1.26)
            rect(s, ML, yy + 0.12, 0.11, 0.11, fill=accent,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
            tfp = textbox(s, ML + 0.32, yy, pw, ph + 0.08)
            para(tfp, p, size=psize, color=INK, line=1.26, first=True)
            yy += ph + gapp
        x = ML + lw + 0.4
        n = len(right_cards)
        gy = 0.18
        ch = (h - gy * (n - 1)) / n
        for i, (hd, bd) in enumerate(right_cards):
            acc = ACCENTS[i % len(ACCENTS)]
            yy2 = y + i * (ch + gy)
            card(s, x, yy2, rw, ch, fill=WHITE, line=BORDER, shadow=True)
            rect(s, x, yy2, 0.05, ch, fill=acc)
            iw = rw - 0.56
            hsize = autosize(hd, iw, 0.44, 13.5, True, 1.14, min_size=11.5)
            hh = text_height_in(hd, hsize, iw, True, 1.14)
            bsize = autosize(bd, iw, ch - hh - 0.46, 11.5, False, 1.28, min_size=9)
            bh = text_height_in(bd, bsize, iw, False, 1.28)
            top = yy2 + max(0.16, (ch - hh - bh - 0.1) / 2)
            tf = textbox(s, x + 0.3, top, iw, hh + 0.08)
            para(tf, hd, size=hsize, bold=True, color=NAVY, line=1.14, first=True)
            tf = textbox(s, x + 0.3, top + hh + 0.1, iw, bh + 0.1)
            para(tf, bd, size=bsize, color=MUTED, line=1.28, first=True)
        return s

    def quote_slide(self, statement, attribution=None, kicker=None, accent=TEAL):
        s = self._new(footer=False)
        gradient_bg(s, NAVY, NAVY_DEEP, 315.0)
        ghost(s, -1.4, 4.0, 5.0, 5.0, "3B7AF7", 8000)
        ghost(s, 10.6, -1.2, 5.0, 5.0, "00B3A4", 6000)
        if self.logo:
            place_logo(s, self.logo, 0, 0.68, 0.7, on_dark=True, max_w=1.9,
                       right_edge=ML + CW - 0.1)
        if kicker:
            tf = textbox(s, ML + 0.4, 1.9, CW - 0.8, 0.3)
            para(tf, kicker.upper(), size=10, bold=True, color=accent, spacing=2.0,
                 align=PP_ALIGN.CENTER, first=True)
        size = 30 if len(statement) > 90 else 34
        tf = textbox(s, 1.7, 2.5, SW - 3.4, 2.4, anchor=MSO_ANCHOR.MIDDLE)
        para(tf, statement, size=size, bold=True, color=WHITE, line=1.18,
             align=PP_ALIGN.CENTER, first=True)
        rect(s, SW / 2 - 0.35, 4.98, 0.7, 0.05, fill=accent)
        if attribution:
            tf = textbox(s, 2.2, 5.24, SW - 4.4, 0.5)
            para(tf, attribution, size=12.5, color=RGBColor(0xA9, 0xBB, 0xDA),
                 align=PP_ALIGN.CENTER, line=1.3, first=True)
        return s

    def takeaways_slide(self, points, title="Key takeaways", kicker="Recap",
                        sub=None, accent=TEAL):
        s = self._new()
        y = self._head(s, kicker, title, sub, accent) + 0.4
        n = len(points)
        cols = 2 if n > 5 else 1
        rows = (n + cols - 1) // cols
        gx, gy = 0.3, 0.2
        cw = (CW - gx) / 2 if cols == 2 else CW
        h = min(1.55, (BODY_BOTTOM - y - gy * (rows - 1)) / rows)
        for i, p in enumerate(points):
            head, desc = (p if isinstance(p, tuple) else (p, None))
            if cols == 2:
                r, c = i % rows, i // rows
            else:
                r, c = i, 0
            x = ML + c * (cw + gx)
            yy = y + r * (h + gy)
            acc = ACCENTS[i % len(ACCENTS)]
            card(s, x, yy, cw, h, fill=WHITE, line=BORDER, shadow=True)
            rect(s, x, yy, 0.05, h, fill=acc)
            circle(s, x + 0.54, yy + h / 2, 0.42, fill=ACCENT_TINTS[str(acc)])
            tfn = textbox(s, x + 0.33, yy + h / 2 - 0.115, 0.42, 0.26)
            para(tfn, f"{i + 1}", size=12.5, bold=True, color=acc,
                 align=PP_ALIGN.CENTER, first=True)
            iw = cw - 1.24
            budget = h - 0.28
            hsize = autosize(head, iw, budget * (0.56 if desc else 1.0), 13.5,
                             True, 1.16, min_size=11)
            hh = text_height_in(head, hsize, iw, True, 1.16)
            dsize = hh2 = 0
            if desc:
                dsize = autosize(desc, iw, budget - hh - 0.06, 11.5, False, 1.26,
                                 min_size=9)
                hh2 = text_height_in(desc, dsize, iw, False, 1.26) + 0.05
            top = yy + max(0.12, (h - hh - hh2) / 2)
            tf = textbox(s, x + 0.98, top, iw, hh + hh2 + 0.1)
            para(tf, head, size=hsize, bold=True, color=NAVY, line=1.16, first=True)
            if desc:
                para(tf, desc, size=dsize, color=MUTED, line=1.26, space_before=3)
        return s

    def closing_slide(self, title="Thank you", subtitle=None, questions=None):
        s = self._new(footer=False)
        gradient_bg(s, NAVY, NAVY_DEEP, 315.0)
        ghost(s, 9.8, -1.4, 5.4, 5.4, "3B7AF7", 9000)
        ghost(s, 11.4, 3.8, 3.6, 3.6, "00B3A4", 7000)
        rect(s, 0, 0, 0.16, SH, fill=TEAL)
        tf = textbox(s, ML + 0.3, 2.36, 7.4, 0.3)
        para(tf, f"UNIT {self.unit_no:02d}  ·  {self.unit_title}".upper(), size=10,
             bold=True, color=TEAL, spacing=2.0, first=True)
        tf = textbox(s, ML + 0.3, 2.78, 8.0, 1.1)
        para(tf, title, size=46, bold=True, color=WHITE, line=1.04, first=True)
        rect(s, ML + 0.3, 4.02, 0.72, 0.05, fill=BLUE)
        if subtitle:
            tf = textbox(s, ML + 0.3, 4.32, 7.2, 0.8)
            para(tf, subtitle, size=14.5, color=RGBColor(0xB8, 0xC8, 0xE4), line=1.34,
                 first=True)
        if questions:
            qw = 8.0
            qbox = 1.32
            tf = textbox(s, ML + 0.3, 5.16, qw, qbox)
            para(tf, "Questions to think about", size=10, bold=True, color=TEAL,
                 spacing=1.4, caps=True, first=True)
            qs = 11.5
            while qs > 9.0 and sum(
                    text_height_in("— " + q, qs, qw, False, 1.28) + 0.07
                    for q in questions) > qbox - 0.26:
                qs -= 0.5
            for q in questions:
                para(tf, "— " + q, size=qs, color=RGBColor(0x93, 0xA6, 0xC6),
                     line=1.28, space_before=4)
        if self.logo:
            place_logo(s, self.logo, 0, 0.72, 0.95, on_dark=True, max_w=2.4,
                       right_edge=ML + CW - 0.1)
        # presenter sign-off
        hline(s, ML + 0.3, 6.6, 3.6, RGBColor(0x2A, 0x3E, 0x66), 1.0)
        tf = textbox(s, ML + 0.3, 6.74, 6.4, 0.32)
        rich(tf, [(self.PRESENTER, {'bold': True, 'color': WHITE, 'size': 11.5}),
                  ("   ·   Register No. ", {'color': RGBColor(0x7E, 0x91, 0xB4),
                                            'size': 10}),
                  (self.REGISTER, {'bold': True,
                                   'color': RGBColor(0xC5, 0xD2, 0xE8),
                                   'size': 10}),
                  ("   ·   ", {'color': RGBColor(0x4A, 0x5E, 0x86), 'size': 10}),
                  (self.course, {'color': RGBColor(0x7E, 0x91, 0xB4),
                                 'size': 10})],
             first=True)
        return s


# --------------------------------------------------------------------------
# Notes support
# --------------------------------------------------------------------------
def notes(slide, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = text
    return slide
