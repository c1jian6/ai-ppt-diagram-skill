"""Colour themes shared by the Graphviz output and the slide chrome.

Every theme exposes the same shape so the renderer and the deck builder can
stay generic:

``slide_bg`` / ``on_bg``   slide background and text colours for the chrome
``node`` / ``edge`` / ``cluster``   Graphviz defaults
``tones``   semantic fills: primary, accent, success, warn, danger, neutral, ghost
"""

from __future__ import annotations

from . import fonts

#: OOXML font family used for every run (auto-detected CJK font).
FONT_DEFAULT = fonts.FAMILY

TONE_NAMES = ("primary", "accent", "success", "warn", "danger", "neutral", "ghost")


def _tone(fill: str, stroke: str, font: str) -> dict:
    return {"fill": fill, "stroke": stroke, "font": font}


THEMES: dict[str, dict] = {
    "ocean": {
        "label": "海洋蓝 · Ocean",
        "dark": False,
        "slide_bg": "#FFFFFF",
        "on_bg": {
            "title": "#0E2233",
            "text": "#33455A",
            "muted": "#7C8CA0",
            "rule": "#D9E4F0",
            "accent": "#1466A8",
        },
        "node": {"fill": "#F2F7FD", "stroke": "#B7CFE8", "font": "#16324B"},
        "edge": {"color": "#8AA0B8", "font": "#566B80"},
        "cluster": {"stroke": "#CFDEED", "font": "#6A8098", "fill": "none"},
        "tones": {
            "primary": _tone("#1466A8", "#0F5288", "#FFFFFF"),
            "accent": _tone("#E1EFFB", "#7FB0DA", "#0F3557"),
            "success": _tone("#E3F3E7", "#7CBB8C", "#1B4E28"),
            "warn": _tone("#FFF2D9", "#DFB262", "#6B4308"),
            "danger": _tone("#FBE7E6", "#DB8B86", "#7C201C"),
            "neutral": _tone("#EFF3F7", "#BAC6D3", "#293747"),
            "ghost": _tone("none", "#A6B4C4", "#5C6E80"),
        },
        "edge_tones": {
            "normal": "#8AA0B8",
            "muted": "#B9C6D6",
            "strong": "#1466A8",
            "success": "#4E9A63",
            "danger": "#C0504A",
        },
    },
    "forest": {
        "label": "森林绿 · Forest",
        "dark": False,
        "slide_bg": "#FFFFFF",
        "on_bg": {
            "title": "#10281E",
            "text": "#31463A",
            "muted": "#75887C",
            "rule": "#D8E7DD",
            "accent": "#1E7A4F",
        },
        "node": {"fill": "#F1F8F3", "stroke": "#B8D8C2", "font": "#153025"},
        "edge": {"color": "#8AA99B", "font": "#526B5D"},
        "cluster": {"stroke": "#CCE2D3", "font": "#63806F", "fill": "none"},
        "tones": {
            "primary": _tone("#1E7A4F", "#176240", "#FFFFFF"),
            "accent": _tone("#DFF1E6", "#7CBC98", "#10402B"),
            "success": _tone("#E2F4E6", "#7ABF8A", "#194D26"),
            "warn": _tone("#FDF3D8", "#D8B25F", "#65430A"),
            "danger": _tone("#FBE8E5", "#D98C82", "#7A241C"),
            "neutral": _tone("#EEF4F0", "#B6C8BD", "#26362D"),
            "ghost": _tone("none", "#A3B5AA", "#566B5F"),
        },
        "edge_tones": {
            "normal": "#8AA99B",
            "muted": "#BACABF",
            "strong": "#1E7A4F",
            "success": "#3F8F5B",
            "danger": "#B4544B",
        },
    },
    "sunset": {
        "label": "暖阳橙 · Sunset",
        "dark": False,
        "slide_bg": "#FFFFFF",
        "on_bg": {
            "title": "#33200F",
            "text": "#4A3A2A",
            "muted": "#96826E",
            "rule": "#F0DED0",
            "accent": "#C05621",
        },
        "node": {"fill": "#FDF6F1", "stroke": "#EACDB8", "font": "#3A2415"},
        "edge": {"color": "#B99C8A", "font": "#7A6353"},
        "cluster": {"stroke": "#EED9C9", "font": "#8B7461", "fill": "none"},
        "tones": {
            "primary": _tone("#C05621", "#9C4519", "#FFFFFF"),
            "accent": _tone("#FBE8DA", "#E2A177", "#5C2B0E"),
            "success": _tone("#E7F2E3", "#8FBC85", "#284A22"),
            "warn": _tone("#FDF0D5", "#DDB15C", "#67430A"),
            "danger": _tone("#FAE6E2", "#D98A7E", "#7B2418"),
            "neutral": _tone("#F5F1ED", "#CBBBAF", "#3A2A1D"),
            "ghost": _tone("none", "#C0AB9C", "#7A6353"),
        },
        "edge_tones": {
            "normal": "#B99C8A",
            "muted": "#D8C7BB",
            "strong": "#C05621",
            "success": "#5C8F50",
            "danger": "#B4443A",
        },
    },
    "ink": {
        "label": "墨线 · Ink (print)",
        "dark": False,
        "slide_bg": "#FFFFFF",
        "on_bg": {
            "title": "#111827",
            "text": "#374151",
            "muted": "#6B7280",
            "rule": "#E5E7EB",
            "accent": "#111827",
        },
        "node": {"fill": "#F9FAFB", "stroke": "#9CA3AF", "font": "#111827"},
        "edge": {"color": "#6B7280", "font": "#4B5563"},
        "cluster": {"stroke": "#D1D5DB", "font": "#6B7280", "fill": "none"},
        "tones": {
            "primary": _tone("#374151", "#1F2937", "#FFFFFF"),
            "accent": _tone("#E5E7EB", "#9CA3AF", "#111827"),
            "success": _tone("#E9EFE9", "#8FA893", "#1F2A21"),
            "warn": _tone("#F5F1E6", "#B7A98A", "#3B331F"),
            "danger": _tone("#F2E8E7", "#B39390", "#3B2220"),
            "neutral": _tone("#F3F4F6", "#C4C9D0", "#1F2937"),
            "ghost": _tone("none", "#9CA3AF", "#4B5563"),
        },
        "edge_tones": {
            "normal": "#6B7280",
            "muted": "#B4BCC6",
            "strong": "#111827",
            "success": "#4B6B4F",
            "danger": "#7F3B36",
        },
    },
    "tech": {
        "label": "深空 · Tech (dark)",
        "dark": True,
        "slide_bg": "#0B1220",
        "on_bg": {
            "title": "#E6F0FA",
            "text": "#A9BCD2",
            "muted": "#6C819B",
            "rule": "#1E2C42",
            "accent": "#22D3EE",
        },
        "node": {"fill": "#101B2E", "stroke": "#2A4260", "font": "#DCE9F7"},
        "edge": {"color": "#3C5A7D", "font": "#8FA9C6"},
        "cluster": {"stroke": "#1E3350", "font": "#7E97B5", "fill": "none"},
        "tones": {
            "primary": _tone("#0E7C8C", "#22D3EE", "#E8FDFF"),
            "accent": _tone("#13263E", "#2F5677", "#BFE6F5"),
            "success": _tone("#10301F", "#2F7D55", "#C9F2D8"),
            "warn": _tone("#33240D", "#8A6320", "#F5DDB0"),
            "danger": _tone("#34161A", "#8E3B42", "#F7C9CD"),
            "neutral": _tone("#141E2E", "#33465E", "#C6D4E4"),
            "ghost": _tone("none", "#33465E", "#9DB3CC"),
        },
        "edge_tones": {
            "normal": "#3C5A7D",
            "muted": "#2A4260",
            "strong": "#22D3EE",
            "success": "#3F9E6B",
            "danger": "#B4585F",
        },
    },
}

ALIASES = {
    "default": "ocean",
    "blue": "ocean",
    "green": "forest",
    "warm": "sunset",
    "orange": "sunset",
    "mono": "ink",
    "gray": "ink",
    "grey": "ink",
    "print": "ink",
    "dark": "tech",
    "night": "tech",
}


def names() -> list[str]:
    return list(THEMES)


def get(name: str | None) -> dict:
    """Return the theme dict for *name* (defaults to ocean)."""
    key = (name or "ocean").strip().lower()
    key = ALIASES.get(key, key)
    if key not in THEMES:
        raise KeyError(f"unknown theme {name!r}; available: {', '.join(THEMES)}")
    return THEMES[key]


def normalize_custom(raw: dict | None) -> dict:
    """Validate a spec-level ``tones`` palette and fill in sensible defaults."""
    custom: dict[str, dict] = {}
    for key, value in (raw or {}).items():
        if not isinstance(value, dict) or not value.get("fill"):
            raise ValueError(f"custom tone {key!r} needs at least a 'fill' colour")
        fill = str(value["fill"])
        custom[str(key).strip().lower()] = {
            "fill": fill,
            "stroke": str(value.get("stroke") or _shift(fill, -0.22)),
            "font": str(value.get("font") or ("#FFFFFF" if is_dark_colour(fill) else "#111111")),
        }
    return custom


def _shift(hex_colour: str, factor: float) -> str:
    """Lighten (positive) or darken (negative) a colour by *factor*."""
    value = (hex_colour or "").lstrip("#")
    if len(value) != 6:
        return hex_colour
    channels = []
    for index in (0, 2, 4):
        channel = int(value[index : index + 2], 16) / 255.0
        channel = channel * (1.0 + factor) if factor < 0 else channel + (1.0 - channel) * factor
        channels.append(max(0, min(255, int(round(channel * 255)))))
    return "#{:02X}{:02X}{:02X}".format(*channels)


def tone(theme: dict, name: str | None, custom: dict | None = None) -> dict:
    """Resolve a tone name against *theme* (or the spec's custom palette)."""
    if not name or str(name) in ("auto", "default", "base"):
        return theme["node"]
    key = str(name).strip().lower()
    if custom and key in custom:
        return custom[key]
    if key.startswith("#"):
        return {
            "fill": key,
            "stroke": _shift(key, -0.22),
            "font": "#FFFFFF" if is_dark_colour(key) else "#111111",
        }
    tones = theme["tones"]
    if key not in tones:
        raise KeyError(f"unknown tone {name!r}; available: {', '.join(list(tones) + list(custom or {}))}")
    return tones[key]


def edge_colour(theme: dict, name: str | None, custom: dict | None = None) -> str | None:
    """Resolve an edge tone to a stroke colour."""
    if not name:
        return None
    key = str(name).strip().lower()
    if custom and key in custom:
        return custom[key]["stroke"]
    if key.startswith("#"):
        return key
    return theme["edge_tones"].get(key)


def is_dark_colour(hex_colour: str) -> bool:
    """True when *hex_colour* is dark enough to need light text on it."""
    value = (hex_colour or "").lstrip("#")
    if len(value) != 6:
        return False
    r, g, b = (int(value[i : i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) < 140
