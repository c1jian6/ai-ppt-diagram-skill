"""dot2pptx — Graphviz layouts turned into editable PowerPoint shapes.

Quick start (after ``pip install -e .``)::

    dot2pptx check
    dot2pptx deck examples/01-pipeline.json -o out/pipeline.pptx --mode native
    dot2pptx render examples/01-pipeline.json -o out/pipeline.png --formats png,svg,emf
"""

from __future__ import annotations

__version__ = "0.1.0"

from . import fonts, gv, layers, native, panels, spec, themes  # noqa: F401
from .deck import DeckBuilder, parse_size  # noqa: F401

__all__ = [
    "DeckBuilder",
    "parse_size",
    "fonts",
    "gv",
    "layers",
    "native",
    "panels",
    "spec",
    "themes",
    "__version__",
]
