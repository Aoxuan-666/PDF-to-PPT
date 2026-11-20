"""Reusable OOXML, typography, shape, and layout helpers for vector PDF-to-PPTX reconstruction."""

from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Map PDF font base names (subset prefix stripped) -> (Latin typeface, East Asian typeface, is_bold)
FONT_MAP = {
    "MicrosoftYaHei": ("Microsoft YaHei", "微软雅黑", False),
    "MicrosoftYaHei-Bold": ("Microsoft YaHei", "微软雅黑", True),
    "MicrosoftYaHeiUI": ("Microsoft YaHei UI", "微软雅黑", False),
    "MicrosoftYaHeiUI-Bold": ("Microsoft YaHei UI", "微软雅黑", True),
    "FangSong_GB2312": ("FangSong", "仿宋", False),
    "FangSong": ("FangSong", "仿宋", False),
    "SimSun": ("SimSun", "宋体", False),
    "NSimSun": ("NSimSun", "新宋体", False),
    "SimHei": ("SimHei", "黑体", False),
    "DengXian": ("DengXian", "等线", False),
    "DengXian-Bold": ("DengXian", "等线", True),
    "STKaiti": ("STKaiti", "华文楷体", False),
    "KaiTi": ("KaiTi", "楷体", False),
    "KaiTi_GB2312": ("KaiTi", "楷体", False),
    "MS-UIGothic": ("MS UI Gothic", "MS UI Gothic", False),
}


def norm_size(sz: float) -> float:
    """Round PDF export font size jitter (±0.05pt) to the nearest 0.5pt."""
    return round(sz * 2.0) / 2.0


def int_to_rgb(color_int: int) -> tuple[int, int, int]:
    """Convert PyMuPDF packed 24-bit integer color to (R, G, B)."""
    return ((color_int >> 16) & 255, (color_int >> 8) & 255, color_int & 255)


def float_to_rgb(rgb_tuple) -> tuple[int, int, int]:
    """Convert PyMuPDF normalized 0..1 RGB tuple to 0..255 (R, G, B)."""
    return tuple(max(0, min(255, round(c * 255))) for c in rgb_tuple)


def resolve_font_names(pdf_font: str) -> tuple[str, str, bool]:
    """Resolve PDF embedded font name to (latin, east_asian, map_bold)."""
    base_font = pdf_font.split("+")[-1]
    if base_font in FONT_MAP:
        return FONT_MAP[base_font]
    if base_font.endswith("-Bold"):
        stem = base_font[:-5]
        if stem in FONT_MAP:
            lat, ea, _ = FONT_MAP[stem]
            return (lat, ea, True)
        return (stem, stem, True)
    return (base_font, base_font, False)


def set_run_font(run, pdf_font: str, size_pt: float, rgb: tuple[int, int, int], bold_flag: bool = False):
    """Apply font family (writing <a:latin>, <a:ea>, <a:cs>), size, bold, and RGB color."""
    latin, ea, map_bold = resolve_font_names(pdf_font)
    is_bold = map_bold or bold_flag

    run.font.name = latin
    run.font.size = Pt(norm_size(size_pt))
    run.font.bold = bool(is_bold)
    run.font.color.rgb = RGBColor(*rgb)

    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(f"{{{A_NS}}}{tag}")
        if el is None:
            el = etree.SubElement(rPr, f"{{{A_NS}}}{tag}")
        el.set("typeface", ea if tag in ("ea", "cs") else latin)


def set_native_bullet(paragraph, mar_l_pt: float = 27.0, indent_pt: float = -27.0,
                      bullet_char: str = "•", bullet_font: str = "Microsoft YaHei",
                      bullet_rgb: tuple[int, int, int] | None = None):
    """Configure native OOXML hanging-indent bullet on a paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(round(mar_l_pt * 12700))))
    pPr.set("indent", str(int(round(indent_pt * 12700))))

    buNone = pPr.find(f"{{{A_NS}}}buNone")
    if buNone is not None:
        pPr.remove(buNone)

    if bullet_rgb is not None:
        buClr = pPr.find(f"{{{A_NS}}}buClr")
        if buClr is None:
            buClr = etree.SubElement(pPr, f"{{{A_NS}}}buClr")
        for child in list(buClr):
            buClr.remove(child)
        srgb = etree.SubElement(buClr, f"{{{A_NS}}}srgbClr")
        srgb.set("val", "{:02X}{:02X}{:02X}".format(*bullet_rgb))

    buFont = pPr.find(f"{{{A_NS}}}buFont")
    if buFont is None:
        buFont = etree.SubElement(pPr, f"{{{A_NS}}}buFont")
    buFont.set("typeface", bullet_font)

    buChar = pPr.find(f"{{{A_NS}}}buChar")
    if buChar is None:
        buChar = etree.SubElement(pPr, f"{{{A_NS}}}buChar")
    buChar.set("char", bullet_char)


def add_runs_to_paragraph(paragraph, spans):
    """Append styled PPTX runs from PyMuPDF span dicts while preserving inline colors/fonts."""
    for sp in spans:
        txt = sp["text"]
        if not txt:
            continue
        r = paragraph.add_run()
        r.text = txt
        is_bold = bool(sp["flags"] & 16)
        set_run_font(r, sp["font"], sp["size"], int_to_rgb(sp["color"]), is_bold)


def configure_tf(tf, ml=0.0, mr=0.0, mt=0.0, mb=0.0, word_wrap=True, anchor=MSO_ANCHOR.TOP):
    """Set explicit padding, word wrap, and vertical anchor on a TextFrame."""
    tf.word_wrap = word_wrap
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(ml)
    tf.margin_right = Pt(mr)
    tf.margin_top = Pt(mt)
    tf.margin_bottom = Pt(mb)


def add_textbox(slide, x, y, w, h, ml=0.0, mr=0.0, mt=0.0, mb=0.0,
                word_wrap=True, anchor=MSO_ANCHOR.TOP):
    """Create a textbox positioned in points with zero margins by default."""
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    configure_tf(tb.text_frame, ml=ml, mr=mr, mt=mt, mb=mb, word_wrap=word_wrap, anchor=anchor)
    return tb


def remove_shape_shadow(shape):
    """Strip PowerPoint's default theme drop shadow (<a:effectRef idx='2'>) so shapes render flat."""
    shape.shadow.inherit = False
    style = shape._element.find(f"{{{P_NS}}}style")
    if style is not None:
        effectRef = style.find(f"{{{A_NS}}}effectRef")
        if effectRef is not None:
            effectRef.set("idx", "0")
    spPr = shape._element.find(f"{{{P_NS}}}spPr")
    if spPr is not None:
        eff = spPr.find(f"{{{A_NS}}}effectLst")
        if eff is None:
            etree.SubElement(spPr, f"{{{A_NS}}}effectLst")
        else:
            for ch in list(eff):
                eff.remove(ch)


def fill_shape_solid(shape, rgb: tuple[int, int, int], stroke_rgb: tuple[int, int, int] | None = None,
                     stroke_pt: float | None = None):
    """Apply flat solid fill (and optional stroke) with shadows stripped."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*rgb)
    if stroke_rgb is not None and stroke_pt:
        shape.line.color.rgb = RGBColor(*stroke_rgb)
        shape.line.width = Pt(stroke_pt)
    else:
        shape.line.fill.background()
    remove_shape_shadow(shape)


def measure_rounded_rect_radius(drawing) -> float:
    """Measure exact corner radius (pt) from a {'c':4, 'l':4} rounded-rectangle drawing."""
    r = drawing["rect"]
    radii = []
    for it in drawing["items"]:
        if it[0] == "l":
            p1, p2 = it[1], it[2]
            dx, dy = abs(p2.x - p1.x), abs(p2.y - p1.y)
            radii.append((r.width - dx) / 2.0 if dx > dy else (r.height - dy) / 2.0)
    return radii[0] if radii else 5.76


def measure_right_arrow_adjustments(drawing) -> tuple[float, float]:
    """Compute DrawingML rightArrow (shaft_height_ratio, head_length_ratio) from 7-edge vector polygon."""
    r = drawing["rect"]
    if r.width <= 0 or r.height <= 0:
        return (0.5, 0.5)
    pts = [(it[1].x, it[1].y) for it in drawing["items"] if it[0] == "l"]
    if not pts:
        return (0.5, 0.5)
    ys_norm = sorted(set(round((y - r.y0) / r.height, 4) for _, y in pts))
    adj0 = (ys_norm[-2] - ys_norm[1]) if len(ys_norm) >= 4 else 0.5
    neck_xs = [x for x, y in pts if abs(y - r.y0) < 0.2 or abs(y - r.y1) < 0.2]
    if neck_xs:
        neck_x = sum(neck_xs) / len(neck_xs)
        head_len = max(0.0, r.x1 - neck_x)
        adj1 = head_len / min(r.width, r.height)
    else:
        adj1 = 0.5
    return (round(adj0, 4), round(adj1, 4))


def extract_clean_shapes(page, slide_w: float, slide_h: float):
    """Extract background color and deduplicated native shapes from a PDF page."""
    drawings = page.get_drawings()
    bg_rgb = (255, 255, 255)
    raw_shapes = []

    for idx, d in enumerate(drawings):
        r = d["rect"]
        ops = {}
        for it in d["items"]:
            ops[it[0]] = ops.get(it[0], 0) + 1

        # Skip WPS 0.5pt-expanded tessellated border duplicates
        if ops.get("l", 0) >= 50:
            continue

        fcol = float_to_rgb(d["fill"]) if d.get("fill") else None
        scol = float_to_rgb(d["color"]) if d.get("color") else None
        stroke_pt = d.get("width") if scol is not None else None

        if fcol is None and scol is None:
            continue

        # Full-page solid rect becomes native slide background
        if (
            fcol is not None
            and scol is None
            and ops == {"re": 1}
            and abs(r.x0) < 1.0
            and abs(r.y0) < 1.0
            and abs(r.width - slide_w) < 1.0
            and abs(r.height - slide_h) < 1.0
        ):
            bg_rgb = fcol
            continue

        if ops == {"c": 4, "l": 4}:
            kind = "rounded_rect"
            extra = {"radius": measure_rounded_rect_radius(d)}
        elif ops == {"c": 4}:
            kind = "circle"
            extra = {}
        elif ops == {"l": 7}:
            kind = "right_arrow"
            extra = {"arrow_adj": measure_right_arrow_adjustments(d)}
        else:
            kind = "rect"
            extra = {}

        raw_shapes.append({
            "idx": idx,
            "kind": kind,
            "rect": (r.x0, r.y0, r.x1, r.y1),
            "w": r.width,
            "h": r.height,
            "fill": fcol if fcol is not None else (255, 255, 255),
            "stroke": scol,
            "stroke_pt": stroke_pt,
            **extra,
        })

    # Drop any earlier opaque shape completely occluded by a later identical-bbox opaque shape
    cleaned = []
    for i, s in enumerate(raw_shapes):
        occluded = False
        x0, y0, x1, y1 = s["rect"]
        for later in raw_shapes[i + 1:]:
            lx0, ly0, lx1, ly1 = later["rect"]
            if (
                abs(x0 - lx0) < 0.5
                and abs(y0 - ly0) < 0.5
                and abs(x1 - lx1) < 0.5
                and abs(y1 - ly1) < 0.5
                and later["stroke"] is None
            ):
                occluded = True
                break
        if not occluded:
            cleaned.append(s)

    return bg_rgb, cleaned


def add_vector_shapes(slide, shapes):
    """Create native PowerPoint shapes from deduplicated descriptors and return created shapes."""
    created = []
    for s in shapes:
        x0, y0, x1, y1 = s["rect"]
        w, h = s["w"], s["h"]
        if w <= 0 or h <= 0:
            created.append(None)
            continue
        if s["kind"] == "rounded_rect":
            shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x0), Pt(y0), Pt(w), Pt(h))
            shp.adjustments[0] = min(0.5, s.get("radius", 5.76) / min(w, h))
        elif s["kind"] == "circle":
            shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(x0), Pt(y0), Pt(w), Pt(h))
        elif s["kind"] == "right_arrow":
            shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Pt(x0), Pt(y0), Pt(w), Pt(h))
            adj0, adj1 = s.get("arrow_adj", (0.5, 0.5))
            shp.adjustments[0] = adj0
            shp.adjustments[1] = adj1
        else:
            shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x0), Pt(y0), Pt(w), Pt(h))
        fill_shape_solid(shp, s["fill"], s.get("stroke"), s.get("stroke_pt"))
        created.append(shp)
    return created


def block_all_spans(block):
    return [sp for line in block["lines"] for sp in line["spans"]]


def place_badge_in_oval(oval_shape, line_dict):
    """Place centered badge digit/label directly inside an OVAL shape."""
    tf = oval_shape.text_frame
    configure_tf(tf, ml=0, mr=0, mt=0, mb=0, word_wrap=False, anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs_to_paragraph(p, line_dict["spans"])


def place_text_in_banner(rect_shape, spans, ml_pt: float = 14.4, align=PP_ALIGN.LEFT):
    """Place centered or left-padded single-line/short text directly inside a rounded rect."""
    tf = rect_shape.text_frame
    mr_pt = ml_pt if align == PP_ALIGN.CENTER else 10.0
    configure_tf(tf, ml=ml_pt, mr=mr_pt, mt=0, mb=0, word_wrap=True, anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = align
    add_runs_to_paragraph(p, spans)


def exact_pitch_pt(lines, first_sz: float) -> float:
    """Measure exact baseline pitch from multi-line PDF block, else 1.2 * font_size."""
    if len(lines) >= 2:
        pitches = [lines[i + 1]["bbox"][1] - lines[i]["bbox"][1] for i in range(len(lines) - 1)]
        return round(sum(pitches) / len(pitches), 2)
    return round(first_sz * 1.2, 2)


def default_font_y_offset(pdf_font: str, size_pt: float, is_multiline: bool, pitch_pt: float | None = None) -> float:
    """Calibrated vertical offset (pt) aligning PowerPoint line-box top with WPS PDF bbox."""
    latin, _, _ = resolve_font_names(pdf_font)
    sz = norm_size(size_pt)
    if latin == "FangSong":
        return -3.2 if sz >= 15.0 else -2.6
    if latin == "DengXian":
        return -2.2
    if latin == "STKaiti":
        if is_multiline and pitch_pt is not None and pitch_pt > sz * 1.4:
            return -4.9
        return -1.5
    if latin.startswith("Microsoft YaHei"):
        if sz >= 44.0:
            return 7.2 if is_multiline else 0.0
        if sz >= 30.0:
            return 3.5
        if sz >= 26.0 and not is_multiline:
            return 2.7
        if 13.5 <= sz <= 16.5 and not is_multiline:
            return 1.4
    return 0.0


def place_simple_block(slide, block, w_override=None, pitch_pt=None, align=PP_ALIGN.LEFT,
                       join_lines=True, word_wrap=True, y_offset=None):
    """Place a standalone text block (single line or wrapped paragraph) as an editable textbox."""
    x0, y0, x1, y1 = block["bbox"]
    lines = block["lines"]
    max_lw = max(l["bbox"][2] - l["bbox"][0] for l in lines)
    first_span = lines[0]["spans"][0]
    first_sz = norm_size(first_span["size"])
    if w_override is not None:
        w = w_override
    elif len(lines) == 1:
        w = max_lw + first_sz * 1.5
    else:
        w = max_lw + first_sz * 0.45

    if pitch_pt is None:
        pitch_pt = exact_pitch_pt(lines, first_sz)
    if y_offset is None:
        y_offset = default_font_y_offset(first_span["font"], first_sz, len(lines) > 1, pitch_pt)

    h = max(y1 - y0 + 6.0, first_sz * 1.4)
    tb = add_textbox(slide, x0, y0 + y_offset, w, h, word_wrap=word_wrap)
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.alignment = align
    if len(lines) > 1:
        p.line_spacing = Pt(pitch_pt)
    if join_lines:
        add_runs_to_paragraph(p, block_all_spans(block))
    else:
        for li, line in enumerate(lines):
            if li > 0:
                p = tf.add_paragraph()
                p.alignment = align
                p.line_spacing = Pt(pitch_pt)
            add_runs_to_paragraph(p, line["spans"])
    return tb


def place_bullet_group(slide, bullet_blocks, x_bullet_override=None, w_override=None,
                       space_after_pt=8.0, y_offset=1.5):
    """Group consecutive bullet blocks into one textbox with native OOXML hanging bullets."""
    first = bullet_blocks[0]
    last = bullet_blocks[-1]
    bx0 = x_bullet_override if x_bullet_override is not None else first["lines"][0]["bbox"][0]
    by0 = first["bbox"][1] + y_offset
    tx0 = first["lines"][1]["bbox"][0]
    mar_l = tx0 - bx0

    max_lw = 0.0
    for b in bullet_blocks:
        for l in b["lines"][1:]:
            max_lw = max(max_lw, l["bbox"][2] - l["bbox"][0])

    first_sz = norm_size(first["lines"][1]["spans"][0]["size"])
    pitch_pt = round(first_sz * 1.2, 2)
    if w_override is not None:
        w = w_override
    else:
        has_wrap = any(len(b["lines"]) > 2 for b in bullet_blocks)
        pad = first_sz * (0.45 if has_wrap else 1.5)
        w = mar_l + max_lw + pad

    h = (last["bbox"][3] - first["bbox"][1]) + 10.0
    tb = add_textbox(slide, bx0, by0, w, h, word_wrap=True)
    tf = tb.text_frame

    for bi, b in enumerate(bullet_blocks):
        p = tf.paragraphs[0] if bi == 0 else tf.add_paragraph()
        p.line_spacing = Pt(pitch_pt)
        if bi < len(bullet_blocks) - 1:
            p.space_after = Pt(space_after_pt)
        b_span = b["lines"][0]["spans"][0]
        b_rgb = int_to_rgb(b_span["color"])
        set_native_bullet(p, mar_l_pt=mar_l, indent_pt=-mar_l, bullet_char="•", bullet_rgb=b_rgb)
        body_spans = [sp for l in b["lines"][1:] for sp in l["spans"]]
        add_runs_to_paragraph(p, body_spans)
    return tb


def place_header_block(slide, block):
    """Place a 2-line slide header (kicker + title) inside one editable textbox."""
    lines = block["lines"]
    x0 = block["bbox"][0]
    y0 = block["bbox"][1]
    if len(lines) == 2:
        kicker, title = lines[0], lines[1]
        max_lw = max(kicker["bbox"][2] - kicker["bbox"][0], title["bbox"][2] - title["bbox"][0])
        title_sz = norm_size(title["spans"][0]["size"])
        tb = add_textbox(slide, x0, y0, max_lw + title_sz * 1.5, block["bbox"][3] - y0 + 10.0, word_wrap=False)
        tf = tb.text_frame
        p0 = tf.paragraphs[0]
        p0.line_spacing = Pt(15.8)
        gap = title["bbox"][1] - kicker["bbox"][3]
        p0.space_after = Pt(round(gap + 3.2, 1))
        add_runs_to_paragraph(p0, kicker["spans"])
        p1 = tf.add_paragraph()
        p1.line_spacing = 1.0
        add_runs_to_paragraph(p1, title["spans"])
        return tb
    return place_simple_block(slide, block, word_wrap=False)
