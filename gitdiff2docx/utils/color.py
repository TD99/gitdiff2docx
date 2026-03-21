"""Color conversion and manipulation utilities."""

from docx.shared import RGBColor


def rgb_from_hex(hex_color, fallback=(0, 0, 0)):
    """Convert hex color string to RGBColor object."""
    try:
        cleaned = hex_color.lstrip("#")
        return RGBColor(
            int(cleaned[0:2], 16),
            int(cleaned[2:4], 16),
            int(cleaned[4:6], 16),
        )
    except Exception:
        return RGBColor(*fallback)


def normalize_hex_color(hex_color, fallback="auto"):
    """Normalize and validate hex color string."""
    if isinstance(hex_color, str):
        cleaned = hex_color.strip().lstrip("#")
        if len(cleaned) == 6:
            try:
                int(cleaned, 16)
                return cleaned.upper()
            except ValueError:
                pass
    return fallback
