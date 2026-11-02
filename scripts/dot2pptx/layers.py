"""Layered architecture diagrams (input -> method -> fusion -> output).

Graphviz is good at graphs, not at rigid grids: columns of equal height, every
item box the same width, arrows centred between frames. This module lays such
a diagram out deterministically with real PowerPoint shapes.

Spec shape (JSON)::

    {
      "layout": "layers",
      "layers": [
        {"label": "输入层", "items": ["遥感图像", "网络开源文本"]},
        {"label": "方法层", "items": [{"text": "图像与视频描述生成"}]},
        {"label": "融合层", "items": [{"text": "特征关联", "tone": "primary"}]},
        {"label": "输出层", "items": [{"text": "目标特征信息", "sub": "供下一专题使用"}]}
      ]
    }
"""

from __future__ import annotations

from pathlib import Path

from PIL import ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from . import fonts, themes

EMU_PER_PT = 12700



# layout constants, in points at scale 1.0
ITEM_FONT = 11.5
SUB_FONT = 9.5
HEAD_FONT = 10.5
ITEM_PAD_X = 13.0
ITEM_PAD_Y = 10.0
ITEM_MIN_H = 40.0
GAP_ITEMS = 10.0
FRAME_PAD = 13.0
FRAME_HEAD = 30.0
ARROW_GAP = 46.0
ARROW_W = 26.0
ARROW_H = 20.0

_font_cache: dict[tuple[str, int], ImageFont.FreeTypeFont] = {}


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(str(value).lstrip("#").upper())


def _font(size: float, bold: bool = False):
    path = fonts.measure_path(bold) or fonts.measure_path()
    key = (path, max(int(round(size * 4)), 8))
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(path, key[1])
    return _font_cache[key]


def _text_width(text: str, size: float, bold: bool = False) -> float:
    font = _font(size, bold)
    return max(font.getlength(line) for line in str(text).split("\n")) / 4.0


def _block_height(text: str, size: float, sub: str | None) -> float:
    lines = str(text).split("\n")
    height = len(lines) * size * 1.42
    if sub:
        height += SUB_FONT * 1.42 + 2.0
    return height


def _clean(shape) -> None:
    element = shape._element
    style = element.find(qn("p:style"))
    if style is not None:
        element.remove(style)
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def _add_shape(slide, kind, left, top, width, height, radius=None):
    shape = slide.shapes.add_shape(kind, Emu(int(left)), Emu(int(top)), Emu(int(width)), Emu(int(height)))
    if radius is not None:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    _clean(shape)
    return shape


def _text_frame(shape, colour: str, size: float, bold: bool = False, anchor=MSO_ANCHOR.MIDDLE,
                align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = Emu(int(4 * EMU_PER_PT))
    tf.margin_top = tf.margin_bottom = Emu(int(2 * EMU_PER_PT))
    tf.vertical_anchor = anchor
    para = tf.paragraphs[0]
    para.alignment = align
    para.space_before = Pt(0)
    para.space_after = Pt(0)
    para.line_spacing = 1.0
    return tf, para


def _write_run(para, text: str, size: float, colour: str, bold: bool = False):
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = themes.FONT_DEFAULT
    run.font.color.rgb = _rgb(colour)
    return run


def normalize(raw: dict) -> dict:
    """Normalise the ``layers`` part of a spec."""
    layers = []
    for index, layer in enumerate(raw.get("layers") or []):
        items = []
        for item in layer.get("items") or []:
            if isinstance(item, str):
                item = {"text": item}
            text = str(item.get("text") or "")
            items.append(
                {
                    "text": text,
                    "sub": (str(item["sub"]) if item.get("sub") else None),
                    "tone": item.get("tone"),
                }
            )
        layers.append(
            {
                "label": str(layer.get("label") or f"层 {index + 1}"),
                "items": items,
                "tone": layer.get("tone"),
            }
        )
    if not layers:
        raise ValueError("layout 'layers' needs at least one layer")
    return {"layout": "layers", "layers": layers}


def measure(spec: dict) -> dict:
    """Compute frame and item geometry in points (unscaled)."""
    columns = []
    for layer in spec["layers"]:
        widths, heights = [], []
        for item in layer["items"]:
            w = _text_width(item["text"], ITEM_FONT)
            if item["sub"]:
                w = max(w, _text_width(item["sub"], SUB_FONT))
            widths.append(w + 2 * ITEM_PAD_X)
            heights.append(max(ITEM_MIN_H, _block_height(item["text"], ITEM_FONT, item["sub"]) + 2 * ITEM_PAD_Y))
        item_w = max(widths) if widths else 120.0
        content_h = sum(heights) + GAP_ITEMS * (len(heights) - 1 if heights else 0)
        columns.append(
            {
                "layer": layer,
                "item_width": item_w,
                "item_heights": heights,
                "content_h": content_h,
                "frame_w": item_w + 2 * FRAME_PAD,
                "frame_h": FRAME_HEAD + content_h + 2 * FRAME_PAD,
            }
        )

    frame_h = max(column["frame_h"] for column in columns)
    total_w = sum(column["frame_w"] for column in columns) + ARROW_GAP * (len(columns) - 1)
    return {"columns": columns, "total_w": total_w, "total_h": frame_h}


def draw(slide, spec: dict, theme: dict, box: tuple[int, int, int, int]) -> dict:
    """Draw a horizontal layer diagram inside *box* (EMU); returns a report."""
    layout = measure(spec)
    columns = layout["columns"]
    box_left, box_top, box_w, box_h = box
    box_w_pt = box_w / EMU_PER_PT
    box_h_pt = box_h / EMU_PER_PT

    scale = min(box_w_pt / layout["total_w"], box_h_pt / layout["total_h"], 1.35)
    total_w = layout["total_w"] * scale
    total_h = layout["total_h"] * scale
    origin_x = box_left / EMU_PER_PT + (box_w_pt - total_w) / 2.0
    origin_y = box_top / EMU_PER_PT + (box_h_pt - total_h) / 2.0
    centre_y = origin_y + total_h / 2.0

    def emu(value: float) -> int:
        return int(round(value * EMU_PER_PT))

    cursor = origin_x
    frames = []
    for column in columns:
        layer = column["layer"]
        frame_w = column["frame_w"] * scale
        frame_h = column["frame_h"] * scale
        left = cursor
        top = centre_y - frame_h / 2.0

        frame = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, emu(left), emu(top),
                           emu(frame_w), emu(frame_h), radius=0.045)
        frame.name = f"layer:{layer['label']}"
        frame.fill.solid()
        frame.fill.fore_color.rgb = _rgb(theme["slide_bg"])
        frame.line.color.rgb = _rgb(theme["node"]["stroke"])
        frame.line.width = Pt(1.0)

        # layer header
        label_box = slide.shapes.add_textbox(emu(left + 12 * scale), emu(top + 8 * scale),
                                             emu(frame_w - 24 * scale), emu(FRAME_HEAD * scale))
        label_box.name = f"layer-label:{layer['label']}"
        tf = label_box.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT
        _write_run(para, str(layer["label"]), HEAD_FONT * scale, theme["cluster"]["font"], bold=True)

        # items, stacked under the layer header
        content_h = column["content_h"] * scale
        y = top + (FRAME_HEAD + FRAME_PAD) * scale
        for index, item in enumerate(layer["items"]):
            item_tone = themes.tone(theme, item["tone"])
            item_h = column["item_heights"][index] * scale
            item_w = column["item_width"] * scale
            sw = _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, emu(left + FRAME_PAD * scale), emu(y),
                            emu(item_w), emu(item_h), radius=0.16)
            sw.name = f"item:{item['text'].splitlines()[0]}"
            fill = item_tone.get("fill", "none")
            if not fill or fill == "none":
                sw.fill.background()
            else:
                sw.fill.solid()
                sw.fill.fore_color.rgb = _rgb(fill)
            sw.line.color.rgb = _rgb(item_tone.get("stroke", "#999999"))
            sw.line.width = Pt(1.0)

            tf, para = _text_frame(sw, item_tone.get("font", "#111111"), ITEM_FONT * scale)
            colour = item_tone.get("font", "#111111")
            text = str(item["text"])
            for line_index, line in enumerate(text.split("\n")):
                if line_index:
                    para = tf.add_paragraph()
                    para.alignment = PP_ALIGN.CENTER
                    para.space_before = Pt(0)
                    para.space_after = Pt(0)
                _write_run(para, line, ITEM_FONT * scale, colour, bold=item["tone"] == "primary")
            if item["sub"]:
                sub = tf.add_paragraph()
                sub.alignment = PP_ALIGN.CENTER
                sub.space_before = Pt(2)
                _write_run(sub, str(item["sub"]), SUB_FONT * scale,
                           item_tone.get("font", "#111111") if item["tone"] == "primary"
                           else theme["edge"]["font"])
            y += item_h + GAP_ITEMS * scale

        frames.append((left, frame_w))
        cursor += frame_w + ARROW_GAP * scale

    # block arrows between frames, all on the centre line
    for index in range(len(frames) - 1):
        left, width = frames[index]
        gap_start = left + width
        arrow_w = min(ARROW_W * scale, ARROW_GAP * scale * 0.75)
        arrow_h = min(ARROW_H * scale, total_h * 0.16)
        ax = gap_start + (ARROW_GAP * scale - arrow_w) / 2.0
        arrow = _add_shape(slide, MSO_SHAPE.RIGHT_ARROW, emu(ax), emu(centre_y - arrow_h / 2.0),
                           emu(arrow_w), emu(arrow_h))
        arrow.name = "layer-arrow"
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = _rgb(theme["on_bg"]["accent"])
        arrow.line.fill.background()

    return {"scale": scale, "layers": len(columns), "shapes": len(slide.shapes)}
