#!/usr/bin/env python
"""One-shot renderer used by the skill: JSON spec in, .pptx out.

    python scripts/draw.py spec.json -o out/diagram.pptx
    python scripts/draw.py spec.json -o out/diagram.pptx --mode image --vector png

Everything is bundled in ``scripts/dot2pptx``; no installation step is needed
beyond having Graphviz on PATH for topology diagrams.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from dot2pptx import deck as deck_mod  # noqa: E402
from dot2pptx import spec as spec_mod  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="draw.py",
        description="Render a JSON diagram spec into an editable PowerPoint file.",
    )
    parser.add_argument("spec", help="path to the JSON spec")
    parser.add_argument("-o", "--out", default="out/diagram.pptx", help="output .pptx path")
    parser.add_argument("--mode", choices=("auto", "native", "image"), default="auto",
                        help="native = editable shapes (default), image = rasterised picture")
    parser.add_argument("--vector", choices=("emf", "png"), default="emf",
                        help="format used when --mode image")
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--theme", default="ocean")
    parser.add_argument("--size", default="16:9")
    parser.add_argument("--footer", default=None)
    parser.add_argument("--no-chrome", action="store_true",
                        help="no title, rule, caption or page number")
    parser.add_argument("--assets", default=None,
                        help="where rendered media goes (default: next to the output file)")
    parser.add_argument("--quiet", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    source = Path(args.spec)
    if not source.is_file():
        print(f"[error] spec not found: {source}", file=sys.stderr)
        return 2

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    assets = Path(args.assets) if args.assets else out.parent / "assets"

    slides = spec_mod.load_slides(source)
    builder = deck_mod.DeckBuilder(
        theme=args.theme,
        size=args.size,
        chrome=not args.no_chrome,
        footer=args.footer,
        assets_dir=assets,
    )

    for slide in slides:
        mode = "native" if args.mode == "auto" else args.mode
        report = builder.add_slide(
            slide, mode=mode, vector=args.vector, dpi=args.dpi, notes=True
        )
        if not args.quiet:
            detail = (f"{report.get('panels')} panels" if report.get("panels")
                      else f"{report.get('nodes', 0)} nodes")
            print(f"[ok] slide {report['index']}: {report.get('title') or '(untitled)'} "
                  f"[{report['mode']}, {detail}]")

    builder.save(out)
    if not args.quiet:
        print(f"[saved] {out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
