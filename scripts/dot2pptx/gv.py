"""Graphviz bridge: locate the `dot` executable, render images, read geometry.

The toolkit never depends on the `graphviz` PyPI package; it talks to the
Graphviz binaries directly so that every attribute of the DOT language stays
available.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

#: Known install locations (the official Windows installer puts dot.exe here).
_SEARCH_PATHS = (
    r"C:\Program Files\Graphviz\bin\dot.exe",
    r"C:\Program Files (x86)\Graphviz\bin\dot.exe",
    r"C:\ProgramData\chocolatey\bin\dot.exe",
    r"C:\ProgramData\miniconda3\Library\bin\dot.exe",
    r"C:\ProgramData\anaconda3\Library\bin\dot.exe",
)

#: Output formats that `dot` renders and that we expose.
FORMATS = {
    "png": "raster image (transparent background, DPI controlled)",
    "svg": "vector, for web / Word / Illustrator",
    "pdf": "vector, for print",
    "emf": "vector, native insert into PowerPoint (scales cleanly)",
}

_dot_cache: str | None = None


class GraphvizError(RuntimeError):
    """Raised when Graphviz cannot be found or fails to render."""


def find_dot() -> str:
    """Return the absolute path of the Graphviz `dot` executable."""
    global _dot_cache
    if _dot_cache and Path(_dot_cache).is_file():
        return _dot_cache

    for var in ("DOT2PPTX_DOT", "GRAPHVIZ_DOT"):
        candidate = os.environ.get(var)
        if candidate and Path(candidate).is_file():
            _dot_cache = candidate
            return candidate

    on_path = shutil.which("dot")
    if on_path:
        _dot_cache = on_path
        return on_path

    for candidate in _SEARCH_PATHS:
        if Path(candidate).is_file():
            _dot_cache = candidate
            return candidate

    raise GraphvizError(
        "Graphviz `dot` was not found. Install it with\n"
        "    winget install --id Graphviz.Graphviz -e\n"
        "or set DOT2PPTX_DOT to the full path of the dot executable."
    )


def version() -> str:
    """Return the Graphviz version string, e.g. ``16.1.0``."""
    exe = find_dot()
    proc = subprocess.run([exe, "-V"], capture_output=True, text=True)
    text = (proc.stderr or proc.stdout or "").strip()
    # "dot - graphviz version 16.1.0 (20260904.0139)"
    parts = text.split("version")
    return parts[1].strip().split()[0] if len(parts) > 1 else text


def _run_dot(dot_text: str, args: list[str], engine: str | None = None) -> subprocess.CompletedProcess:
    """Run `dot` on *dot_text* with extra *args*; the DOT source travels via file."""
    exe = find_dot()
    with tempfile.TemporaryDirectory(prefix="dot2pptx-") as tmp:
        src = Path(tmp) / "graph.dot"
        src.write_text(dot_text, encoding="utf-8")
        cmd = [exe]
        if engine:
            cmd.append("-K" + engine)
        cmd.extend(args)
        cmd.append(str(src))
        proc = subprocess.run(cmd, capture_output=True)
        if proc.returncode != 0:
            detail = proc.stderr.decode("utf-8", "replace").strip()
            raise GraphvizError(f"Graphviz failed ({' '.join(args)}):\n{detail}")
        return proc


def render(
    dot_text: str,
    out_path: str | Path,
    fmt: str = "png",
    dpi: int | None = None,
    engine: str | None = None,
) -> Path:
    """Render *dot_text* to *out_path* in *fmt* and return the path."""
    fmt = fmt.lower().lstrip(".")
    if fmt not in FORMATS:
        raise ValueError(f"unsupported format {fmt!r}; choose from {sorted(FORMATS)}")
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    args = ["-T" + fmt, "-o", str(out)]
    if dpi and fmt in ("png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"):
        args.insert(0, f"-Gdpi={int(dpi)}")
    _run_dot(dot_text, args, engine=engine)
    if not out.is_file() or out.stat().st_size == 0:
        raise GraphvizError(f"Graphviz produced no output for {out}")
    return out


def layout_json(dot_text: str, engine: str | None = None) -> dict:
    """Run the layout and return Graphviz's JSON description (geometry in points)."""
    proc = _run_dot(dot_text, ["-Tjson"], engine=engine)
    return json.loads(proc.stdout.decode("utf-8"))


def probe() -> dict:
    """Return a small dict describing the Graphviz installation."""
    exe = find_dot()
    return {"dot": exe, "version": version(), "formats": sorted(FORMATS)}
