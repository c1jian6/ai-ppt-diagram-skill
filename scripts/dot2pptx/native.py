"""Draw a laid-out Graphviz graph as *native, editable* PowerPoint shapes.

Graphviz does the layout (``dot -Tjson`` returns node centres, sizes, cluster
boxes and edge splines in points, with the origin at the *bottom* left); this
module replays that geometry as real autoshapes, freeform connectors and text,
so the result can be restyled and edited inside PowerPoint instead of being a
flat picture.
"""

from __future__ import annotations

import math

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from lxml import etree
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from . import gv, themes

EMU_PER_PT = 12700

#: spec shape alias -> (PowerPoint autoshape, corner adjustment or None)
_SHAPE_MAP = {
    "box": (MSO_SHAPE.RECTANGLE, None),
    "rect": (MSO_SHAPE.RECTANGLE, None),
    "round": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.14),
    "rounded": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.14),
    "step": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.1),
    "stadium": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.5),
    "terminal": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.5),
    "pill": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.5),
    "start": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.5),
    "end": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.5),
    "diamond": (MSO_SHAPE.DIAMOND, None),
    "decision": (MSO_SHAPE.DIAMOND, None),
    "oval": (MSO_SHAPE.OVAL, None),
    "ellipse": (MSO_SHAPE.OVAL, None),
    "circle": (MSO_SHAPE.OVAL, None),
    "db": (MSO_SHAPE.CAN, None),
    "database": (MSO_SHAPE.CAN, None),
    "cylinder": (MSO_SHAPE.CAN, None),
    "hex": (MSO_SHAPE.HEXAGON, None),
    "hexagon": (MSO_SHAPE.HEXAGON, None),
    "io": (MSO_SHAPE.PARALLELOGRAM, None),
    "input": (MSO_SHAPE.PARALLELOGRAM, None),
    "doc": (MSO_SHAPE.FOLDED_CORNER, None),
    "document": (MSO_SHAPE.FOLDED_CORNER, None),
    "folder": (MSO_SHAPE.ROUNDED_RECTANGLE, 0.08),
    "cloud": (MSO_SHAPE.CLOUD, None),
    "proc": (MSO_SHAPE.CUBE, None),
    "plain": (MSO_SHAPE.RECTANGLE, None),
    "none": (MSO_SHAPE.RECTANGLE, None),
}

_DASH = {
    "dashed": MSO_LINE_DASH_STYLE.DASH,
    "dash": MSO_LINE_DASH_STYLE.DASH,
    "dotted": MSO_LINE_DASH_STYLE.ROUND_DOT,
    "dot": MSO_LINE_DASH_STYLE.ROUND_DOT,
    "solid": MSO_LINE_DASH_STYLE.SOLID,
}


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(str(value).lstrip("#").upper())


def _num(value: str | float) -> float:
    return float(str(value).replace("in", "").strip())


def _bb(raw: str) -> tuple[float, float, float, float]:
    values = [float(part) for part in str(raw).split(",")]
    if len(values) != 4:
        raise ValueError(f"bad bounding box {raw!r}")
    return values[0], values[1], values[2], values[3]


def _points(pos: str) -> list[tuple[float, float]]:
    """Parse a Graphviz spline string into points, ordered tail -> head.

    Graphviz writes the arrow end first (``e,x,y``) followed by the control
    points from tail to head, so the tip has to be moved to the back.
    """
    tokens: list[tuple[str, tuple[float, float]]] = []
    for token in str(pos).split():
        prefix = ""
        if token[:2] in ("e,", "s,"):
            prefix, token = token[:2], token[2:]
        try:
            x, y = token.split(",")
            tokens.append((prefix, (float(x), float(y))))
        except ValueError:
            continue
    if not tokens:
        return []
    prefix, first = tokens[0]
    rest = [point for _, point in tokens[1:]]
    if prefix == "e,":
        return rest + [first]
    return list(reversed(rest)) + [first]


def _label_point(pts: list[tuple[float, float]], width: float, height: float) -> tuple[float, float]:
    """A sensible label anchor: middle of the path, nudged clear of the line."""
    if not pts:
        return (0.0, 0.0)
    if len(pts) == 2:
        a, b = pts
    else:
        middle = len(pts) // 2
        a, b = pts[max(middle - 1, 0)], pts[middle]
    mx, my = (a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0
    dx, dy = b[0] - a[0], b[1] - a[1]
    length = math.hypot(dx, dy) or 1.0
    nx, ny = -dy / length, dx / length
    # Prefer above horizontal lines and to the right of vertical ones.
    if ny > 0 or (abs(ny) < 1e-6 and nx < 0):
        nx, ny = -nx, -ny
    # Push the label clear by half of its own extent along the normal.
    offset = abs(nx) * width / 2.0 + abs(ny) * height / 2.0 + 4.0 * 12700
    return mx + nx * offset, my + ny * offset


def _dedupe(pts: list[tuple[float, float]], eps: float = 0.05) -> list[tuple[float, float]]:
    out: list[tuple[float, float]] = []
    for point in pts:
        if not out or abs(point[0] - out[-1][0]) > eps or abs(point[1] - out[-1][1]) > eps:
            out.append(point)
    return out


def _text_width_pt(text: str, fontsize: float) -> float:
    """Rough advance width: CJK glyphs are ~1 em, Latin ~0.55 em."""
    em = sum(1.0 if ord(ch) >= 0x2E80 else 0.55 for ch in str(text))
    return em * fontsize


def _overlap(a, b) -> bool:
    return not (a[0] + a[2] <= b[0] or b[0] + b[2] <= a[0]
                or a[1] + a[3] <= b[1] or b[1] + b[3] <= a[1])


def _drawing_bbox(data: dict) -> tuple[float, float, float, float]:
    """True drawing extent.

    Graphviz's ``bb`` can be smaller than the nodes when a node has an explicit
    ``width``/``height`` (the box grows, the bbox is not recomputed), so the
    union of the graph bbox, every node box and every cluster box is used.
    """
    x0, y0, x1, y1 = _bb(data.get("bb", "0,0,1,1"))
    for obj in data.get("objects", []):
        if "pos" in obj:
            cx, cy = (float(part) for part in str(obj["pos"]).split(",")[:2])
            half_w = _num(obj.get("width", "1")) * 72 / 2.0
            half_h = _num(obj.get("height", "0.5")) * 72 / 2.0
            x0, x1 = min(x0, cx - half_w), max(x1, cx + half_w)
            y0, y1 = min(y0, cy - half_h), max(y1, cy + half_h)
        elif "bb" in obj:
            cx0, cy0, cx1, cy1 = _bb(obj["bb"])
            x0, y0 = min(x0, cx0), min(y0, cy0)
            x1, y1 = max(x1, cx1), max(y1, cy1)
    return x0, y0, x1, y1


class _Painter:
    """Maps Graphviz points (y up) onto the slide (y down) and adds shapes."""

    def __init__(self, slide, theme: dict, bbox, scale: float, origin: tuple[float, float],
                 graph: dict, font_scale: float):
        self.slide = slide
        self.theme = theme
        self.bbox = bbox
        self.scale = scale
        self.origin = origin
        self.graph = graph
        self.font_scale = font_scale
        self.label_rects: list[tuple[int, int, int, int]] = []

    # -- geometry ---------------------------------------------------------- #
    def px(self, value: float) -> int:
        return int(round((self.origin[0] + (value - self.bbox[0]) * self.scale) * EMU_PER_PT))

    def py(self, value: float) -> int:
        # Graphviz y grows upwards, PowerPoint y grows downwards.
        return int(round((self.origin[1] + (self.bbox[3] - value) * self.scale) * EMU_PER_PT))

    def w(self, value: float) -> int:
        return int(round(value * self.scale * EMU_PER_PT))

    # -- primitives -------------------------------------------------------- #
    def _clean(self, shape) -> None:
        """Drop the theme style reference (it re-adds shadows and fills)."""
        element = shape._element
        style = element.find(qn("p:style"))
        if style is not None:
            element.remove(style)
        try:
            shape.shadow.inherit = False
        except Exception:
            pass

    def _shadow(self, shape, blur=12.0, distance=3.0, alpha=0.16):
        """Attach a soft outer shadow (points), for a little depth on key cards."""
        sp_pr = shape._element.spPr
        effect = sp_pr.find(qn("a:effectLst"))
        if effect is None:
            effect = etree.SubElement(sp_pr, qn("a:effectLst"))
        shadow = etree.SubElement(effect, qn("a:outerShdw"))
        shadow.set("blurRad", str(int(blur * 12700)))
        shadow.set("dist", str(int(distance * 12700)))
        shadow.set("dir", "5400000")
        shadow.set("rotWithShape", "0")
        colour = etree.SubElement(shadow, qn("a:srgbClr"))
        colour.set("val", "1B2A3A")
        alpha_el = etree.SubElement(colour, qn("a:alpha"))
        alpha_el.set("val", str(int(alpha * 100000)))

    def add_box(self, left, top, width, height, mso_shape, radius=None):
        shape = self.slide.shapes.add_shape(mso_shape, Emu(left), Emu(top), Emu(width), Emu(height))
        if radius is not None:
            try:
                shape.adjustments[0] = radius
            except Exception:
                pass
        self._clean(shape)
        return shape

    # -- content ----------------------------------------------------------- #
    def cluster(self, left, top, width, height, label, fontsize):
        shape = self.add_box(left, top, width, height, MSO_SHAPE.ROUNDED_RECTANGLE, 0.05)
        shape.name = f"cluster:{label}"
        shape.fill.background()
        shape.line.color.rgb = _rgb(self.theme["cluster"]["stroke"])
        shape.line.width = Pt(1.0)
        shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if label:
            box = self.slide.shapes.add_textbox(
                Emu(left + self.w(8)),
                Emu(top + self.w(6)),
                Emu(max(width - self.w(16), self.w(24))),
                Emu(self.w(fontsize * 1.9)),
            )
            box.name = f"cluster-label:{label}"
            tf = box.text_frame
            tf.word_wrap = False
            tf.margin_left = tf.margin_right = Emu(0)
            tf.margin_top = tf.margin_bottom = Emu(0)
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT
            run = para.add_run()
            run.text = label
            run.font.size = Pt(max(fontsize * self.scale, 8))
            run.font.name = themes.FONT_DEFAULT
            run.font.color.rgb = _rgb(self.theme["cluster"]["font"])
        return shape

    def node(self, node, tone, left, top, width, height, fontsize):
        mso_shape, radius = _SHAPE_MAP.get(node["shape"], (MSO_SHAPE.ROUNDED_RECTANGLE, 0.14))
        shape = self.add_box(left, top, width, height, mso_shape, radius)
        shape.name = f"node:{node['id']}"

        fill = tone.get("fill", "none")
        if not fill or fill == "none":
            shape.fill.background()
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb(fill)
        shape.line.color.rgb = _rgb(tone.get("stroke", "#999999"))
        shape.line.width = Pt(max(1.0, float(self.graph.get("node_penwidth", 1.2))))

        tf = shape.text_frame
        tf.word_wrap = False
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = Emu(self.w(2))
        tf.margin_top = tf.margin_bottom = Emu(self.w(1))
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        colour = tone.get("font", "#111111")
        body_colour = "#C9D6E4" if themes.is_dark_colour(fill) else self.theme["edge"]["font"]
        bullets = node.get("bullets") or []

        if bullets:
            # "Card" node: badge + bold title on the first line, then bullets.
            tf.margin_left = tf.margin_right = Emu(self.w(6))
            tf.margin_top = tf.margin_bottom = Emu(self.w(3))
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT
            para.space_before = Pt(0)
            para.space_after = Pt(0)
            para.line_spacing = 1.0
            badge = node.get("badge")
            if badge:
                run = para.add_run()
                run.text = f"{badge}　"
                run.font.size = Pt(max(fontsize, 6))
                run.font.bold = True
                run.font.name = themes.FONT_DEFAULT
                run.font.color.rgb = _rgb("#C9D6E4" if themes.is_dark_colour(fill)
                                          else self.theme["on_bg"]["accent"])
            run = para.add_run()
            run.text = str(node["label"]).replace("\n", " ")
            run.font.size = Pt(max(fontsize, 6))
            run.font.bold = True
            run.font.name = themes.FONT_DEFAULT
            run.font.color.rgb = _rgb(colour)

            body_size = max(fontsize - 2.2 * self.scale, 6.0)
            for bullet in bullets:
                para = tf.add_paragraph()
                para.alignment = PP_ALIGN.LEFT
                para.space_before = Pt(max(fontsize * 0.22, 1))
                para.space_after = Pt(0)
                para.line_spacing = 1.0
                run = para.add_run()
                run.text = f"· {bullet}"
                run.font.size = Pt(body_size)
                run.font.name = themes.FONT_DEFAULT
                run.font.color.rgb = _rgb(body_colour)

        else:
            first = True
            for line in str(node["label"]).split("\n"):
                para = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                para.alignment = PP_ALIGN.CENTER
                para.space_before = Pt(0)
                para.space_after = Pt(0)
                para.line_spacing = 0.95
                run = para.add_run()
                run.text = line
                run.font.size = Pt(max(fontsize, 6))
                run.font.name = themes.FONT_DEFAULT
                run.font.color.rgb = _rgb(colour)

            if node.get("shadow"):
                self._shadow(shape)

            if node.get("sub"):
                para = tf.add_paragraph()
                para.alignment = PP_ALIGN.CENTER
                para.space_before = Pt(max(fontsize * 0.12, 1))
                para.space_after = Pt(0)
                para.line_spacing = 0.95
                run = para.add_run()
                run.text = str(node["sub"])
                run.font.size = Pt(max(fontsize * 0.8, 6))
                run.font.name = themes.FONT_DEFAULT
                run.font.color.rgb = _rgb(body_colour)
        return shape

    def edge(self, pts, colour, width, dash, arrow):
        pts = _dedupe(pts)
        if len(pts) < 2:
            return
        builder = self.slide.shapes.build_freeform(Emu(pts[0][0]), Emu(pts[0][1]), scale=1.0)
        builder.add_line_segments(pts[1:], close=False)
        shape = builder.convert_to_shape()
        shape.name = "edge"
        self._clean(shape)
        shape.fill.background()
        shape.line.color.rgb = _rgb(colour)
        shape.line.width = Pt(width)
        if dash:
            shape.line.dash_style = _DASH.get(str(dash).lower(), MSO_LINE_DASH_STYLE.DASH)

        if arrow and arrow not in ("none", "invis"):
            self._arrowhead(pts[-2], pts[-1], colour, width)

    def _arrowhead(self, previous, head, colour, width):
        dx = head[0] - previous[0]
        dy = head[1] - previous[1]
        length = math.hypot(dx, dy)
        if length < 1:
            return
        ux, uy = dx / length, dy / length
        size = max(5.0, min(9.0, 5.0 + width * 2.2))
        height = size * 1.15
        size_emu = self.w(size)
        height_emu = self.w(height)
        cx = head[0] - ux * height_emu / 2.0
        cy = head[1] - uy * height_emu / 2.0
        shape = self.add_box(int(cx - size_emu / 2), int(cy - height_emu / 2),
                             size_emu, height_emu, MSO_SHAPE.ISOSCELES_TRIANGLE)
        shape.rotation = math.degrees(math.atan2(dy, dx)) + 90.0
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(colour)
        shape.line.fill.background()

    def edge_label(self, centre_x, centre_y, width_pt, text, fontsize, colour):
        lines = str(text).split("\n")
        width = width_pt * 1.35 + 8
        height = len(lines) * fontsize * 1.6 + 4
        left = int(centre_x - self.w(width) / 2)
        top = int(centre_y - self.w(height) / 2)
        # Nudge labels apart when two edges run close together.
        for _ in range(6):
            rect = (left, top, self.w(width), self.w(height))
            if not any(_overlap(rect, other) for other in self.label_rects):
                break
            top += self.w(height + 2)
        self.label_rects.append((left, top, self.w(width), self.w(height)))

        box = self.slide.shapes.add_textbox(
            Emu(left), Emu(top), Emu(self.w(width)), Emu(self.w(height)),
        )
        box.name = "edge-label"
        tf = box.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        first = True
        for line in lines:
            para = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            para.alignment = PP_ALIGN.CENTER
            para.space_before = Pt(0)
            para.space_after = Pt(0)
            run = para.add_run()
            run.text = line
            run.font.size = Pt(max(fontsize, 6))
            run.font.name = themes.FONT_DEFAULT
            run.font.color.rgb = _rgb(colour)
        return box


def draw(slide, slide_spec: dict, theme: dict, box: tuple[int, int, int, int],
         dot_text: str, engine: str | None = None, font_scale: float = 1.0) -> dict:
    """Render *dot_text* onto *slide* as native shapes inside *box* (EMU)."""
    data = gv.layout_json(dot_text, engine=engine)
    bbox = _drawing_bbox(data)
    graph_pt_w = bbox[2] - bbox[0]
    graph_pt_h = bbox[3] - bbox[1]
    if graph_pt_w <= 0 or graph_pt_h <= 0:
        raise ValueError("Graphviz returned an empty drawing")

    box_left, box_top, box_width, box_height = box
    box_w_pt = box_width / EMU_PER_PT
    box_h_pt = box_height / EMU_PER_PT
    scale = min(box_w_pt / graph_pt_w, box_h_pt / graph_pt_h)
    origin = (
        box_left / EMU_PER_PT + (box_w_pt - graph_pt_w * scale) / 2.0,
        box_top / EMU_PER_PT + (box_h_pt - graph_pt_h * scale) / 2.0,
    )

    graph = slide_spec["graph"]
    custom = slide_spec.get("tones") or {}
    painter = _Painter(slide, theme, bbox, scale, origin, graph, font_scale)
    by_name = {node["id"]: node for node in slide_spec["nodes"]}
    clusters = {f"cluster_{item['id']}": item for item in slide_spec["clusters"]}

    # 1. cluster frames sit behind everything
    for obj in data.get("objects", []):
        if "bb" not in obj or not str(obj.get("name", "")).startswith("cluster_"):
            continue
        x0, y0, x1, y1 = _bb(obj["bb"])
        info = clusters.get(obj["name"])
        painter.cluster(
            painter.px(x0), painter.py(y1), painter.w(x1 - x0), painter.w(y1 - y0),
            info["label"] if info else "", float(graph["edge_fontsize"]),
        )

    # 2. edges
    edge_font = max(float(graph["edge_fontsize"]) * scale * painter.font_scale, 6.0)
    edge_tones = theme["edge_tones"]
    for index, obj in enumerate(data.get("edges", [])):
        edge_spec = slide_spec["edges"][index] if index < len(slide_spec["edges"]) else {}
        if edge_spec.get("hidden") or edge_spec.get("style") == "invis" or not obj.get("pos"):
            continue
        colour = edge_tones.get(str(edge_spec.get("tone") or "").lower(), theme["edge"]["color"])
        width = max(1.0, float(graph["edge_penwidth"]))
        if str(edge_spec.get("style") or "").lower() in ("bold", "thick"):
            width *= 1.7
        screen_pts = [(painter.px(px), painter.py(py)) for px, py in _points(obj["pos"])]
        painter.edge(screen_pts, colour, width, edge_spec.get("style"),
                     edge_spec.get("arrow") or "normal")

        label = edge_spec.get("label")
        if label:
            lines = str(label).split("\n")
            width_pt = max(_text_width_pt(line, edge_font) for line in lines) + 4
            height_pt = len(lines) * edge_font * 1.4 + 2
            lx, ly = _label_point(screen_pts, painter.w(width_pt), painter.w(height_pt))
            painter.edge_label(lx, ly, width_pt, label, edge_font, theme["edge"]["font"])

    # 3. nodes on top
    count = 0
    for obj in data.get("objects", []):
        if "pos" not in obj:
            continue
        node = by_name.get(obj["name"])
        if node is None:
            continue
        cx, cy = (float(part) for part in obj["pos"].split(",")[:2])
        nw = _num(obj.get("width", "1")) * 72
        nh = _num(obj.get("height", "0.5")) * 72
        tone = themes.tone(theme, node["tone"], custom)
        fs = max(float(graph["node_fontsize"]) * scale * painter.font_scale, 6.0)
        painter.node(
            node, tone,
            painter.px(cx - nw / 2), painter.py(cy + nh / 2), painter.w(nw), painter.w(nh), fs,
        )
        count += 1

    return {"scale": scale, "nodes": count, "shapes": len(slide.shapes),
            "size_pt": (graph_pt_w, graph_pt_h)}
