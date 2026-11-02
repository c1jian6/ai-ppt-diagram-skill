"""Command line interface: ``dot2pptx <command>`` or ``python -m dot2pptx``."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from . import __version__, deck as deck_mod
from . import fonts, gv, spec as spec_mod, themes


def _slug(text: str | None, fallback: str = "diagram") -> str:
    cleaned = re.sub(r"[^0-9A-Za-z_\-\u4e00-\u9fff]+", "-", str(text or "")).strip("-")
    return cleaned[:60] or fallback


def _fmt_names(value: str) -> list[str]:
    return [part.strip().lower().lstrip(".") for part in str(value).split(",") if part.strip()]


# --------------------------------------------------------------------------- #
def cmd_check(args) -> int:
    info = gv.probe()
    print("Graphviz :", info["version"])
    print("dot.exe  :", info["dot"])
    print("formats  :", ", ".join(info["formats"]))

    import pptx
    import PIL

    print("python   :", sys.version.split()[0])
    print("pptx     :", pptx.__version__)
    print("pillow   :", PIL.__version__)

    print("font     :", fonts.describe())
    if not fonts.is_cjk_capable():
        print("         : WARNING no CJK font found - install one or set DOT2PPTX_FONT")
    print("themes   :", ", ".join(themes.names()))

    # smoke test: render + geometry
    dot = 'digraph G { a [label="ok"]; b [label="test"]; a -> b; }'
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        out = gv.render(dot, Path(tmp) / "probe.png", fmt="png", dpi=96)
        data = gv.layout_json(dot)
        print("smoke    :", f"png {out.stat().st_size} bytes, bb={data['bb']}")
    print("status   : ready")
    return 0


def cmd_render(args) -> int:
    slides = spec_mod.load_slides(args.spec)
    formats = _fmt_names(args.formats)
    base = Path(args.out or "out/diagram")
    if base.suffix:
        base = base.with_suffix("")

    for index, slide in enumerate(slides):
        if slide.get("layout") in ("layers", "panels"):
            print(f"[skip] {slide.get('title')}: 'layers' layout is PowerPoint-only; use the deck command")
            continue
        theme = themes.get(args.theme or slide.get("theme"))
        dot_text = spec_mod.build_dot(slide, theme, mode="image")
        stem = base if len(slides) == 1 else base.with_name(f"{base.name}-{index + 1}")
        for fmt in formats:
            path = stem.with_suffix("." + fmt)
            gv.render(dot_text, path, fmt=fmt, dpi=args.dpi,
                      engine=str(slide["graph"].get("engine") or "dot"))
            print(f"[ok] {path}")

    if args.keep_dot:
        dot_dir = Path(args.keep_dot)
        dot_dir.mkdir(parents=True, exist_ok=True)
        for index, slide in enumerate(slides):
            theme = themes.get(args.theme or slide.get("theme"))
            name = _slug(slide.get("title"), f"diagram-{index + 1}")
            path = dot_dir / f"{name}.dot"
            path.write_text(spec_mod.build_dot(slide, theme, mode="image"), encoding="utf-8")
            print(f"[ok] {path}")
    return 0


def _build(args, template: str | Path | None) -> deck_mod.DeckBuilder:
    slides = spec_mod.load_slides(args.spec)
    builder = deck_mod.DeckBuilder(
        theme=args.theme,
        size=args.size,
        template=template,
        chrome=not args.no_chrome,
        footer=args.footer,
        assets_dir=args.assets,
    )
    for slide in slides:
        report = builder.add_slide(slide, mode=args.mode, vector=args.vector, dpi=args.dpi,
                                   notes=not args.no_notes)
        detail = f", {report.get('nodes', 0)} nodes" if args.mode == "native" else ""
        print(f"[ok] slide {report['index']}: {report['title']} ({args.mode}{detail})")
    return builder


def cmd_deck(args) -> int:
    builder = _build(args, None)
    out = builder.save(args.out or "out/deck.pptx")
    print(f"[saved] {out}")
    return 0


def cmd_add(args) -> int:
    target = Path(args.deck)
    if not target.is_file():
        print(f"[error] deck not found: {target}", file=sys.stderr)
        return 2
    builder = _build(args, target)
    out = builder.save(args.out or target)
    print(f"[saved] {out}")
    return 0


def cmd_demo(args) -> int:
    demo_dir = Path(getattr(args, "examples", None) or "examples")
    specs = sorted(demo_dir.glob("*.json"))
    if not specs:
        print(f"[error] no demo specs in {demo_dir}", file=sys.stderr)
        return 2
    builder = deck_mod.DeckBuilder(theme=args.theme, size=args.size, footer=args.footer,
                                   assets_dir=args.assets)
    for path in specs:
        slides = spec_mod.load_slides(path)
        for slide in slides:
            report = builder.add_slide(slide, mode=args.mode, vector=args.vector, dpi=args.dpi)
            print(f"[ok] {path.name} -> slide {report['index']}: {report['title']}")
    out = builder.save(args.out)
    print(f"[saved] {out}")
    return 0


# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="dot2pptx",
        description="Graphviz diagrams -> PowerPoint (image or native editable shapes).",
    )
    parser.add_argument("--version", action="version", version=f"dot2pptx {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    check = sub.add_parser("check", help="verify Graphviz / python-pptx installation")
    check.set_defaults(func=cmd_check)

    common_out = dict(default="out", help="output file or directory (default: out/)")

    render = sub.add_parser("render", help="render a spec to png/svg/pdf/emf")
    render.add_argument("spec", help="JSON spec file")
    render.add_argument("-o", "--out", default="out/diagram.png",
                        help="output base path; one file per format (default: out/diagram.png)")
    render.add_argument("--formats", default="png", help="comma list: png,svg,pdf,emf")
    render.add_argument("--dpi", type=int, default=300)
    render.add_argument("--theme", default=None)
    render.add_argument("--keep-dot", default=None, help="also write the .dot source here")
    render.set_defaults(func=cmd_render)

    def add_common(cmd):
        cmd.add_argument("-o", "--out", default=None, help="output .pptx path")
        cmd.add_argument("--mode", choices=("image", "native"), default="image",
                         help="image = Graphviz picture, native = editable PPT shapes")
        cmd.add_argument("--vector", choices=("emf", "png"), default="emf")
        cmd.add_argument("--dpi", type=int, default=300)
        cmd.add_argument("--theme", default="ocean")
        cmd.add_argument("--size", default="16:9", help="16:9, 4:3, 16:10 or WxH inches")
        cmd.add_argument("--footer", default=None)
        cmd.add_argument("--assets", default="out/assets", help="where rendered images are stored")
        cmd.add_argument("--no-chrome", action="store_true", help="no title / rule / page number")
        cmd.add_argument("--no-notes", action="store_true", help="do not write DOT source to notes")

    deck = sub.add_parser("deck", help="build a new deck from a spec")
    deck.add_argument("spec")
    add_common(deck)
    deck.set_defaults(func=cmd_deck)

    add = sub.add_parser("add", help="append slides to an existing deck")
    add.add_argument("deck")
    add.add_argument("spec")
    add_common(add)
    add.set_defaults(func=cmd_add)

    demo = sub.add_parser("demo", help="build the showcase deck from examples/*.json")
    demo.add_argument("--examples", default="examples")
    demo.add_argument("-o", "--out", default="out/demo.pptx", help="output .pptx path")
    demo.add_argument("--mode", choices=("image", "native"), default="image")
    demo.add_argument("--vector", choices=("emf", "png"), default="emf")
    demo.add_argument("--dpi", type=int, default=300)
    demo.add_argument("--theme", default="ocean")
    demo.add_argument("--size", default="16:9")
    demo.add_argument("--footer", default=None)
    demo.add_argument("--assets", default="out/assets")
    demo.set_defaults(func=cmd_demo)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except (gv.GraphvizError, KeyError, ValueError, FileNotFoundError) as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
