"""Default warm presentation palette; independent of customer and assessment data.

Legacy drawing colors are normalized at the native PowerPoint boundary so all
layouts (including optional pages) share the same palette. Images and Oracle
masters are retained as authored. New layouts should use the semantic constants.
"""

INK = "#342D2B"
BODY = "#514C49"
HEADER = "#85463D"
CONTRAST = "#633C34"
CREAM = "#F3DFD0"
ACCENT = "#C74634"
BORDER = "#DEC6B8"
BAND = "#F6E8E2"
LIGHT = "#FDF8F4"
ACCENTS = (ACCENT, "#B46B42", "#927346")
PANELS = ("#FCF0EB", "#FAF1E5", "#F5F0E5")

# Compatibility palette for existing drawing primitives. No content or geometry
# decisions belong here; customer-specific labels and shape IDs are never used.
_ALIASES = {
    "1D252C": INK, "252B2B": INK, "353B3B": BODY,
    "4C5961": BODY, "467653": ACCENT, "365F47": HEADER,
    "367A9B": ACCENTS[1], "B86A00": ACCENTS[2],
    "124883": HEADER, "087A78": ACCENTS[1],
    "6941A5": ACCENTS[2], "D46A00": ACCENT, "6A5B8C": ACCENTS[2],
    "24576C": INK,
    "EEF4F0": BAND, "E5F0EA": BAND, "F2F6F3": BAND,
    "F6F7F8": BAND, "F5F7F8": BAND, "EEF2F5": BAND,
    "F8F8F7": LIGHT, "FAF9F7": LIGHT, "F9FAFB": LIGHT,
    "F6F9FC": LIGHT, "E6EBF3": BAND,
    "FFF4E9": PANELS[1], "EDF4F8": PANELS[2],
    "D3E3D8": PANELS[0], "F3DDD8": PANELS[1], "DCEAF1": PANELS[2],
    "FFF5F1": PANELS[0], "F2F7FA": PANELS[1], "FFF8EC": PANELS[2],
    "FFF7EE": PANELS[1], "C9D8CF": BORDER, "D8DEE2": BORDER,
    "D9E0E3": BORDER, "C9D4DF": BORDER, "9ABBE0": BORDER,
}


def warm_rgb(red: int, green: int, blue: int) -> tuple[int, int, int]:
    """Normalize known legacy colors; preserve unspecified colors and white."""
    value = _ALIASES.get(f"{red:02X}{green:02X}{blue:02X}")
    if value is None:
        return red, green, blue
    value = value.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))
