"""Theme styling utilities for document tables."""

from gitdiff2docx.utils.dict import deep_merge_dict
from gitdiff2docx.utils.color import normalize_hex_color


def get_theme_font(theme, config):
    """Get font name and size from theme and config."""
    theme_font = theme.get("font", {})

    font_name = theme_font.get("name") or config.get("diff_font", "Courier New")

    theme_size = theme_font.get("size")
    if theme_size is None:
        resolved_size = int(config.get("diff_font_size", 8))
    else:
        resolved_size = int(theme_size)
    resolved_size = max(resolved_size, int(theme_font.get("min_size", 0)))

    return font_name, resolved_size


def build_border_attrs(side_config, default_config):
    """Build border attributes for a table side."""
    cfg = deep_merge_dict(default_config, side_config if isinstance(side_config, dict) else {})
    visible = bool(cfg.get("visible", False))
    if not visible:
        return {"val": "nil"}

    style = str(cfg.get("style", "single")).strip() or "single"
    weight_pt = float(cfg.get("weight_pt", 0.5))
    sz = int(round(weight_pt * 8))
    sz = max(2, min(96, sz))
    color = normalize_hex_color(cfg.get("color", "auto"), fallback="auto")
    space = max(0, int(cfg.get("space", 0)))

    return {
        "val": style,
        "sz": sz,
        "space": space,
        "color": color,
    }


def build_table_border_spec(theme):
    """Build complete table border specification from theme."""
    table_borders_cfg = theme.get("table_borders", {})
    default_side_cfg = table_borders_cfg.get("default", {})
    default_side = default_side_cfg if isinstance(default_side_cfg, dict) else {}

    return {
        "top": build_border_attrs(table_borders_cfg.get("top", {}), default_side),
        "left": build_border_attrs(table_borders_cfg.get("left", {}), default_side),
        "bottom": build_border_attrs(table_borders_cfg.get("bottom", {}), default_side),
        "right": build_border_attrs(table_borders_cfg.get("right", {}), default_side),
        "insideH": build_border_attrs(table_borders_cfg.get("inside_h", {}), default_side),
        "insideV": build_border_attrs(table_borders_cfg.get("inside_v", {}), default_side),
    }
