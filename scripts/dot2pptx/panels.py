"""Four (or more) tall panels side by side, each holding a vertical chain.

Spec shape (JSON)::

    {
      "layout": "panels",
      "panels": [
        {"title": "数据链路", "tone": "data", "flow": true,
         "items": ["多源数据", "信息抽取", "实体对齐", "图谱补全"]},
        {"title": "综合", "tone": "core",
         "items": [{"text": "Unified layer", "span": true},
                   {"text": "时空属性", "chip": true}]},
        {"title": "应用", "tone": "output", "items": ["目标画像", "关联网络"]}
      ],
      "links": [{"from": 0, "to": 2, "route": "channel"}, {"from": 2, "to": 3}]
    }

Panel width comes from the slide width (four equal columns); content inside a
panel is scaled to the available height. Links between adjacent panels are a
block arrow in the gap; links that jump over a panel drop into a channel under
the panels and rise into the target, so nothing crosses a frame.
"""

from __future__ import annotations

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from . import themes

EMU_PER_PT = 12700

HEAD_FONT = 11.0
ITEM_FONT = 12.0
SUB_FONT = 9.5
CHIP_FONT = 10.0

PANEL_PAD = 11.0
HEADER_H = 34.0
ITEM_PAD_Y = 9.0
ITEM_MIN_H = 36.0
CHIP_MIN_H = 26.0
GAP_ITEMS = 10.0
ARROW_H = 14.0
ARROW_GAP = 5.0
PANEL_GAP = 20.0
CHANNEL_H = 44.0
GAP_ARROW_W = 16.0
GAP_ARROW_H = 14.0
ICON_H = 40.0
ICON_GAP = 6.0


def _weighted_len(text: str) -> float:
    """Approximate text width in em: CJK is 1 em, Latin ~0.55 em."""
    return sum(1.0 if ord(ch) >= 0x2E80 else 0.55 for ch in str(text))


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(str(value).lstrip("#").upper())


def _clean(shape) -> None:
    element = shape._element
    style = element.find(qn("p:style"))
    if style is not None:
        element.remove(style)
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def _shadow(shape, blur=10.0, distance=2.5, alpha=0.18) -> None:
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


def _box(slide, kind, left, top, width, height, radius=None):
    shape = slide.shapes.add_shape(kind, Emu(int(left)), Emu(int(top)), Emu(int(width)), Emu(int(height)))
    if radius is not None:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    _clean(shape)
    return shape


def _text(shape, text, size, colour, bold=False, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
          top_margin=1.0):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = Emu(int(3 * EMU_PER_PT))
    tf.margin_top = Emu(int(top_margin * EMU_PER_PT))
    tf.margin_bottom = Emu(int(1 * EMU_PER_PT))
    tf.vertical_anchor = anchor
    for index, line in enumerate(str(text).split("\n")):
        para = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_before = Pt(0)
        para.space_after = Pt(0)
        para.line_spacing = 1.0
        run = para.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = themes.FONT_DEFAULT
        run.font.color.rgb = _rgb(colour)
    return tf


def _add_para(tf, text, size, colour, bold=False, before=0.0):
    para = tf.add_paragraph()
    para.alignment = PP_ALIGN.CENTER
    para.space_before = Pt(before)
    para.space_after = Pt(0)
    para.line_spacing = 1.0
    run = para.add_run()
    run.text = str(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = themes.FONT_DEFAULT
    run.font.color.rgb = _rgb(colour)
    return para


def _line(slide, x1, y1, x2, y2, colour, width=1.5, head=False):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)),
                                      Emu(int(x2)), Emu(int(y2)))
    conn.name = "link"
    _clean(conn)
    conn.line.color.rgb = _rgb(colour)
    conn.line.width = Pt(width)
    if head:
        ln = conn.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    return conn


def normalize(raw: dict) -> dict:
    panels = []
    for index, panel in enumerate(raw.get("panels") or []):
        items = []
        for item in panel.get("items") or []:
            if isinstance(item, str):
                item = {"text": item}
            items.append(
                {
                    "text": str(item.get("text") or ""),
                    "sub": (str(item["sub"]) if item.get("sub") else None),
                    "span": bool(item.get("span")),
                    "chip": bool(item.get("chip")),
                    "height": item.get("height"),
                    "icon": item.get("icon"),
                    "tone": item.get("tone"),
                }
            )
        panels.append(
            {
                "title": str(panel.get("title") or f"栏 {index + 1}"),
                "tone": panel.get("tone"),
                "flow": bool(panel.get("flow")),
                "items": items,
            }
        )
    if not panels:
        raise ValueError("layout 'panels' needs at least one panel")
    links = []
    for link in raw.get("links") or []:
        links.append(
            {
                "from": int(link["from"]),
                "to": int(link["to"]),
                "tone": link.get("tone"),
                "route": str(link.get("route") or "").lower(),
            }
        )
    return {"layout": "panels", "panels": panels, "links": links}


def measure(spec: dict) -> dict:
    """Content heights in points at scale 1.0; panel widths come from the slide."""
    panels = []
    for panel in spec["panels"]:
        height = 0.0
        heights = []
        for item in panel["items"]:
            size = CHIP_FONT if item["chip"] else ITEM_FONT
            lines = str(item["text"]).split("\n")
            h = len(lines) * size * 1.45 + (SUB_FONT * 1.6 if item["sub"] else 0) + 2 * ITEM_PAD_Y
            if item.get("icon"):
                h += ICON_H + ICON_GAP
            h = max(CHIP_MIN_H if item["chip"] else ITEM_MIN_H, h)
            if item.get("height"):
                h = max(h, float(item["height"]) * 72)
            heights.append(h)
            height += h
        gaps = max(len(heights) - 1, 0)
        if gaps:
            height += gaps * GAP_ITEMS
            if panel["flow"]:
                height += gaps * (ARROW_H + ARROW_GAP)
        panels.append({"panel": panel, "item_heights": heights, "content_h": height})
    return {"panels": panels, "max_content": max(p["content_h"] for p in panels)}


def draw(slide, spec: dict, theme: dict, box: tuple[int, int, int, int]) -> dict:
    layout = measure(spec)
    panels = layout["panels"]
    count = len(panels)
    box_left, box_top, box_w, box_h = box
    box_w_pt = box_w / EMU_PER_PT
    box_h_pt = box_h / EMU_PER_PT

    has_channel = any(abs(link["to"] - link["from"]) > 1 or link.get("route") == "channel"
                      for link in spec["links"])
    reserve = CHANNEL_H if has_channel else 0.0

    panel_w = (box_w_pt * 0.97 - (count - 1) * PANEL_GAP) / count
    chrome_h = HEADER_H + 2 * PANEL_PAD
    scale = min((box_h_pt - reserve) / (layout["max_content"] + chrome_h), 1.35)
    panel_h = layout["max_content"] * scale + chrome_h * scale

    gap = PANEL_GAP * scale
    total_w = panel_w * count + gap * (count - 1)
    left0 = box_left / EMU_PER_PT + (box_w_pt - total_w) / 2.0
    top0 = box_top / EMU_PER_PT + (box_h_pt - reserve - panel_h) / 2.0
    channel_y = top0 + panel_h + max(reserve * 0.62, 16.0 * scale)

    def e(value: float) -> int:
        return int(round(value * EMU_PER_PT))

    geometry = []
    for index, info in enumerate(panels):
        panel = info["panel"]
        tone = themes.tone(theme, panel["tone"], spec.get("tones"))
        left = left0 + index * (panel_w + gap)
        top = top0

        frame = _box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, e(left), e(top), e(panel_w), e(panel_h), radius=0.06)
        frame.name = f"panel:{panel['title']}"
        frame.fill.solid()
        frame.fill.fore_color.rgb = _rgb(theme["slide_bg"])
        frame.line.color.rgb = _rgb(tone["stroke"])
        frame.line.width = Pt(1.0)

        band = _box(slide, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, e(left), e(top),
                    e(panel_w), e(HEADER_H * scale), radius=0.16)
        band.name = f"panel-head:{panel['title']}"
        band.fill.solid()
        band.fill.fore_color.rgb = _rgb(tone["stroke"])
        band.line.fill.background()
        band_text = "#FFFFFF" if themes.is_dark_colour(tone["stroke"]) else "#111111"
        _text(band, panel["title"], HEAD_FONT * scale, band_text, bold=True)

        content_h = info["content_h"] * scale
        y = top + HEADER_H * scale + (panel_h - HEADER_H * scale - content_h) / 2.0
        for position, item in enumerate(panel["items"]):
            item_tone = themes.tone(theme, item["tone"] or panel["tone"], spec.get("tones"))
            item_h = info["item_heights"][position] * scale
            item_w = panel_w - 2 * PANEL_PAD * scale
            x = left + (panel_w - item_w) / 2.0

            shape = _box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, e(x), e(y), e(item_w), e(item_h),
                         radius=0.5 if item["chip"] else 0.14)
            shape.name = f"item:{item['text'].splitlines()[0]}"
            fill = item_tone.get("fill", "none")
            if not fill or fill == "none":
                shape.fill.background()
            else:
                shape.fill.solid()
                shape.fill.fore_color.rgb = _rgb(fill)
            shape.line.color.rgb = _rgb(item_tone.get("stroke", "#999999"))
            shape.line.width = Pt(1.0 if not item["chip"] else 0.9)

            base = ITEM_FONT * 1.25 if item["span"] else (CHIP_FONT if item["chip"] else ITEM_FONT)
            size = base * scale
            longest = max((_weighted_len(line) for line in str(item["text"]).split("\n")), default=1.0) or 1.0
            fit = (item_w - 14.0) / (longest * 1.10)
            size = min(size, fit)
            icon = item.get("icon")
            if icon:
                icon_side = ICON_H * scale
                slide.shapes.add_picture(
                    str(icon),
                    e(x + (item_w - icon_side) / 2),
                    e(y + ITEM_PAD_Y * scale),
                    e(icon_side), e(icon_side),
                )
            tf = _text(shape, str(item["text"]), size, item_tone.get("font", "#111111"),
                       bold=bool(item["span"]),
                       top_margin=(ICON_H + ICON_GAP) if icon else 1.0)
            if item["sub"]:
                _add_para(tf, item["sub"], SUB_FONT * scale,
                          "#E3D6F5" if themes.is_dark_colour(fill) else theme["edge"]["font"],
                          before=3)
            if item["span"]:
                _shadow(shape)

            y += item_h
            if panel["flow"] and position < len(panel["items"]) - 1:
                arrow = _box(slide, MSO_SHAPE.DOWN_ARROW,
                             e(left + panel_w / 2 - GAP_ARROW_W * scale / 2),
                             e(y + ARROW_GAP * scale),
                             e(GAP_ARROW_W * scale), e(ARROW_H * scale))
                arrow.name = "panel-arrow"
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = _rgb(tone["stroke"])
                arrow.line.fill.background()
                y += ARROW_H * scale + ARROW_GAP * scale + GAP_ITEMS * scale
            else:
                y += GAP_ITEMS * scale

        geometry.append((left, top, panel_w, panel_h, tone))

    # links
    channel_slot = 0
    for link in spec["links"]:
        source = geometry[link["from"]]
        target = geometry[link["to"]]
        colour = themes.tone(theme, link.get("tone") or spec["panels"][link["from"]]["tone"],
                             spec.get("tones"))["stroke"]
        adjacent = abs(link["to"] - link["from"]) == 1 and link.get("route") != "channel"
        if adjacent:
            sx = source[0] + source[2]
            tx = target[0]
            sy = source[1] + source[3] / 2.0
            arrow = _box(slide, MSO_SHAPE.RIGHT_ARROW, e(sx + 2 * scale),
                         e(sy - GAP_ARROW_H * scale / 2),
                         e(max(tx - sx - 4 * scale, 8 * scale)), e(GAP_ARROW_H * scale))
            arrow.name = "link"
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = _rgb(colour)
            arrow.line.fill.background()
        else:
            sx = source[0] + source[2] / 2.0
            tx = target[0] + target[2] / 2.0
            sy = source[1] + source[3]
            ty = target[1] + target[3]
            row_y = channel_y + channel_slot * 7.0 * scale
            channel_slot += 1
            _line(slide, e(sx), e(sy), e(sx), e(row_y), colour)
            _line(slide, e(sx), e(row_y), e(tx), e(row_y), colour)
            _line(slide, e(tx), e(row_y), e(tx), e(ty), colour, head=True)

    return {"scale": scale, "panels": count, "shapes": len(slide.shapes)}
