"""Compose PowerPoint slides around a Graphviz drawing.

Two rendering modes:

``image``   the drawing is rendered by Graphviz (EMF vector or high-DPI PNG)
            and placed as one picture — pixel-perfect, safest for complex graphs
``native``  the drawing is rebuilt from Graphviz's layout data as real
            autoshapes / freeforms — fully editable inside PowerPoint
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

from . import gv, layers, native, panels, spec as spec_mod, themes

EMU_PER_INCH = 914400

SIZES = {
    "16:9": (13.333, 7.5),
    "16:10": (13.333, 8.333),
    "4:3": (10.0, 7.5),
    "a4": (11.69, 8.27),
}

LEFT = 0.78
RIGHT = 0.78
TITLE_TOP = 0.58


def parse_size(value) -> tuple[float, float]:
    if isinstance(value, (tuple, list)) and len(value) == 2:
        return float(value[0]), float(value[1])
    key = str(value or "16:9").lower()
    if key in SIZES:
        return SIZES[key]
    if "x" in key:
        width, height = key.split("x")
        return float(width), float(height)
    raise ValueError(f"unknown slide size {value!r}; use 16:9, 4:3, 16:10 or WxH")


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(str(value).lstrip("#").upper())


def _warn_small_text(effective_pt: float, title: str) -> None:
    """Nudge the author when a diagram is squeezed so far that labels shrink."""
    if effective_pt < 9.0:
        print(
            f"[warn] '{title}' labels end up at {effective_pt:.1f} pt on the slide; "
            "consider direction LR, fewer nodes, or splitting the diagram."
        )


def _apply_font(run, size: float, colour: str, bold: bool = False, italic: bool = False,
                name: str = themes.FONT_DEFAULT) -> None:
    font = run.font
    font.size = Pt(size)
    font.bold = bold
    font.italic = italic
    font.name = name
    font.color.rgb = _rgb(colour)
    # PowerPoint picks the East-Asian face from a:ea; python-pptx only writes a:latin.
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rpr, qn("a:ea"))
    ea.set("typeface", name)


class DeckBuilder:
    """Builds (or extends) a deck, one diagram per slide."""

    def __init__(self, theme: str = "ocean", size="16:9", template: str | Path | None = None,
                 chrome: bool = True, footer: str | None = None, assets_dir: str | Path | None = None):
        self.prs = Presentation(str(template)) if template else Presentation()
        if template is None:
            width, height = parse_size(size)
            self.prs.slide_width = Inches(width)
            self.prs.slide_height = Inches(height)
        self.theme_name = theme
        self.chrome = chrome
        self.footer = footer or ""
        self.assets_dir = Path(assets_dir) if assets_dir else Path("out") / "assets"
        self.reports: list[dict] = []

    # -- primitives -------------------------------------------------------- #
    @property
    def width_in(self) -> float:
        return self.prs.slide_width / EMU_PER_INCH

    @property
    def height_in(self) -> float:
        return self.prs.slide_height / EMU_PER_INCH

    def _text(self, slide, left, top, width, height, text, size, colour,
              bold=False, italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
              line_spacing=None):
        box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        tf.vertical_anchor = anchor
        para = tf.paragraphs[0]
        para.alignment = align
        if line_spacing:
            para.line_spacing = line_spacing
        run = para.add_run()
        run.text = text
        _apply_font(run, size, colour, bold=bold, italic=italic)
        return box

    def _rect(self, slide, left, top, width, height, colour):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(colour)
        shape.line.fill.background()
        try:
            shape.shadow.inherit = False
        except Exception:
            pass
        return shape

    # -- slide ------------------------------------------------------------- #
    def add_slide(self, slide_spec: dict, mode: str = "image", vector: str = "emf",
                  dpi: int = 300, page: int | None = None, notes: bool = True) -> dict:
        theme = themes.get(slide_spec.get("theme") or self.theme_name)
        graph = slide_spec.get("graph") or {}
        engine = str(graph.get("engine") or "dot")
        title = slide_spec.get("title") or ""
        subtitle = slide_spec.get("subtitle")
        caption = slide_spec.get("caption")
        kicker = slide_spec.get("kicker")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _rgb(theme["slide_bg"])

        width, height = self.width_in, self.height_in
        inner_w = width - LEFT - RIGHT

        if self.chrome:
            title_top = TITLE_TOP
            if kicker:
                self._text(slide, LEFT, 0.28, inner_w, 0.26, kicker, 10.5, theme["on_bg"]["muted"],
                           bold=False)
                title_top = TITLE_TOP + 0.16
            title_size = 25.0
            if len(title) > 26:
                title_size = 21.5
            if len(title) > 40:
                title_size = 19.0
            self._text(slide, LEFT, title_top, inner_w, 0.62, title, title_size,
                       theme["on_bg"]["title"], bold=True, anchor=MSO_ANCHOR.TOP)
            rule_y = title_top + 0.66
            self._rect(slide, LEFT, rule_y, inner_w, 0.012, theme["on_bg"]["rule"])
            self._rect(slide, LEFT, rule_y - 0.008, 0.92, 0.028, theme["on_bg"]["accent"])
            content_top = rule_y + 0.22
            if subtitle:
                self._text(slide, LEFT, rule_y + 0.14, inner_w, 0.28, subtitle, 12.5,
                           theme["on_bg"]["text"])
                content_top = rule_y + 0.56
        else:
            content_top = 0.35

        footer_h = 0.34
        content_bottom = height - 0.55 - (0.46 if caption else 0.0)

        box = (
            int(Inches(LEFT)),
            int(Inches(content_top)),
            int(Inches(inner_w)),
            int(Inches(max(content_bottom - content_top, 1.0))),
        )

        dot_text = ""
        report = {"index": len(self.prs.slides), "title": title, "mode": mode, "theme": theme["label"]}

        if slide_spec.get("layout") == "layers":
            info = layers.draw(slide, slide_spec, theme, box)
            report.update(info)
            report["mode"] = "layers"
        elif slide_spec.get("layout") == "panels":
            info = panels.draw(slide, slide_spec, theme, box)
            report.update(info)
            report["mode"] = "panels"
        else:
            dot_text = spec_mod.build_dot(slide_spec, theme, mode=mode)
            if mode == "native":
                info = native.draw(slide, slide_spec, theme, box, dot_text, engine=engine)
                report.update(info)
                _warn_small_text(graph["node_fontsize"] * info["scale"], title)
            else:
                fmt = str(vector).lower().lstrip(".")
                if fmt not in ("emf", "png"):
                    raise ValueError("vector must be 'emf' or 'png'")
                self.assets_dir.mkdir(parents=True, exist_ok=True)
                asset = self.assets_dir / f"slide{report['index']:02d}.{fmt}"
                gv.render(dot_text, asset, fmt=fmt, dpi=dpi, engine=engine)
                natural_w, natural_h = spec_mod.graph_size(dot_text, engine=engine)
                pic = self._place_picture(slide, asset, box, title, natural_w / natural_h)
                report.update({"asset": str(asset), "picture": (pic.width, pic.height)})
                if natural_w:
                    _warn_small_text(graph["node_fontsize"] * (pic.width / 12700) / natural_w, title)

        if self.chrome and caption:
            self._text(slide, LEFT, content_bottom + 0.10, inner_w, 0.32, caption, 9.5,
                       theme["on_bg"]["muted"])
        if self.chrome:
            if self.footer:
                self._text(slide, LEFT, height - footer_h, inner_w * 0.6, 0.24, self.footer, 8.5,
                           theme["on_bg"]["muted"])
            number = page if page is not None else len(self.prs.slides)
            self._text(slide, width - RIGHT - 0.6, height - footer_h, 0.6, 0.24, str(number), 8.5,
                       theme["on_bg"]["muted"], align=PP_ALIGN.RIGHT)

        if notes:
            text = slide_spec.get("notes") or (f"Graphviz 源码：\n\n{dot_text}" if dot_text else "")
            if text:
                slide.notes_slide.notes_text_frame.text = text

        self.reports.append(report)
        return report

    def _place_picture(self, slide, path: Path, box, alt: str, aspect: float):
        """Insert *path* fitted into *box*; *aspect* is the drawing's true w/h.

        The picture's own header size is not trusted: Graphviz EMF headers can
        carry different x/y DPIs, which would stretch the image if used.
        """
        pic = slide.shapes.add_picture(str(path), Emu(box[0]), Emu(box[1]), width=Emu(EMU_PER_INCH))
        ratio = 1.0 / aspect                       # height per unit width
        max_w, max_h = box[2], box[3]
        if max_w * ratio <= max_h:                 # width is the limiting dimension
            new_w = max_w
            new_h = int(round(max_w * ratio))
        else:                                      # height is the limiting dimension
            new_h = max_h
            new_w = int(round(max_h / ratio))
        pic.width = Emu(new_w)
        pic.height = Emu(new_h)
        pic.left = Emu(box[0] + (box[2] - new_w) // 2)
        pic.top = Emu(box[1] + (box[3] - new_h) // 2)
        try:
            pic._element.nvPicPr.cNvPr.set("descr", alt or "diagram")
        except Exception:
            pass
        return pic

    def save(self, path: str | Path) -> Path:
        out = Path(path)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out))
        return out
