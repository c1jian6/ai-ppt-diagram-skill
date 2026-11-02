"""Cross-platform CJK font discovery.

The renderer needs one thing from the host: a font that can draw Chinese.
Nothing is bundled (fonts are licensed, and the file would dwarf the package),
so we probe the usual system locations and fall back in a predictable order.

Override with either environment variable:

    DOT2PPTX_FONT="Noto Sans CJK SC"     # OOXML family name
    DOT2PPTX_FONT_FILE=/path/to/font.ttf  # exact file, used for measuring
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

#: (family for OOXML, regular file(s), bold file(s))
_CANDIDATES: list[tuple[str, tuple[str, ...], tuple[str, ...]]] = [
    (
        "Microsoft YaHei",
        (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf"),
        (r"C:\Windows\Fonts\msyhbd.ttc", r"C:\Windows\Fonts\msyhbd.ttf"),
    ),
    (
        "PingFang SC",
        ("/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/PingFang SC.ttc"),
        (),
    ),
    (
        "Hiragino Sans GB",
        ("/System/Library/Fonts/Hiragino Sans GB.ttc",),
        (),
    ),
    (
        "Noto Sans CJK SC",
        (
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ),
        (
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Bold.otf",
        ),
    ),
    (
        "Source Han Sans SC",
        (
            "/usr/share/fonts/opentype/source-han-sans/SourceHanSansSC-Regular.otf",
            "/usr/share/fonts/adobe-source-han-sans/SourceHanSansSC-Regular.otf",
        ),
        ("/usr/share/fonts/opentype/source-han-sans/SourceHanSansSC-Bold.otf",),
    ),
    (
        "WenQuanYi Zen Hei",
        ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",),
        (),
    ),
    (
        "Arial Unicode MS",
        ("/Library/Fonts/Arial Unicode.ttf", r"C:\Windows\Fonts\arialuni.ttf"),
        (),
    ),
]


def _first_existing(paths: tuple[str, ...]) -> str | None:
    for path in paths:
        if Path(path).is_file():
            return path
    return None


def _resolve() -> tuple[str, str | None, str | None]:
    override_family = os.environ.get("DOT2PPTX_FONT")
    override_file = os.environ.get("DOT2PPTX_FONT_FILE")
    if override_file and Path(override_file).is_file():
        return override_family or Path(override_file).stem, override_file, override_file
    for family, regular, bold in _CANDIDATES:
        path = _first_existing(regular)
        if path:
            return override_family or family, path, _first_existing(bold) or path
    return override_family or "Arial", None, None


#: OOXML family name to write into the run properties.
FAMILY, _REGULAR, _BOLD = _resolve()


def measure_path(bold: bool = False) -> str | None:
    """Font file used to measure text width (Pillow); None when no font was found."""
    return _BOLD if bold else _REGULAR


def describe() -> str:
    path = _REGULAR or "not found"
    return f"{FAMILY} ({path})"


def is_cjk_capable() -> bool:
    return _REGULAR is not None
