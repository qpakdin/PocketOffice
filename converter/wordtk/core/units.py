# -*- coding: utf-8 -*-
"""Length / colour parsing and conversion helpers.

Internally every length is a float number of *points*.
OOXML uses twips (1/20 pt), half-points, eighth-points and EMU.
"""
import re

_FA_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
_AR_DIGITS = "٠١٢٣٤٥٦٧٨٩"
_TO_ASCII = {ord(c): str(i) for i, c in enumerate(_FA_DIGITS)}
_TO_ASCII.update({ord(c): str(i) for i, c in enumerate(_AR_DIGITS)})
_TO_ASCII.update({ord("٫"): ".", ord("،"): ",", ord("−"): "-"})


def ascii_digits(text):
    """Convert Persian / Arabic-Indic digits to ASCII."""
    return str(text).translate(_TO_ASCII)


UNIT_PT = {
    "pt": 1.0, "pc": 12.0, "in": 72.0, "inch": 72.0, '"': 72.0,
    "cm": 72.0 / 2.54, "mm": 72.0 / 25.4, "px": 0.75,
    "twip": 0.05, "twips": 0.05, "emu": 1.0 / 12700.0,
}
_RE = re.compile(r'^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*([a-zA-Z"]*)\s*$')


def parse_length(value, default_unit="pt"):
    """Parse 12, '12', '1.5cm', '0.5in', '۲٫۵cm' -> points (float)."""
    if isinstance(value, bool):
        raise ValueError("not a length: %r" % (value,))
    if isinstance(value, (int, float)):
        return float(value) * UNIT_PT[default_unit]
    m = _RE.match(ascii_digits(value))
    if not m:
        raise ValueError("invalid length: %r (examples: 12, 1.5cm, 0.5in, 18pt)" % (value,))
    unit = (m.group(2) or default_unit).lower()
    if unit not in UNIT_PT:
        raise ValueError("unknown unit %r in %r" % (unit, value))
    return float(m.group(1)) * UNIT_PT[unit]


def parse_number(value):
    if isinstance(value, bool):
        raise ValueError("not a number: %r" % (value,))
    if isinstance(value, (int, float)):
        return float(value)
    return float(ascii_digits(value).replace(",", ".").strip().rstrip("%"))


def twips(pt):
    return int(round(pt * 20))


def half_points(pt):
    return int(round(pt * 2))


def emu(pt):
    return int(round(pt * 12700))


def from_twips(v):
    return float(v) / 20.0


def from_emu(v):
    return float(v) / 12700.0


def fmt_length(pt, unit="cm"):
    """Points -> short human string like '2.54cm'."""
    if pt is None:
        return None
    v = pt / UNIT_PT[unit]
    s = ("%.2f" % v).rstrip("0").rstrip(".")
    return s + unit


PAPER_TWIPS = {  # name -> (width, height) portrait
    "A3": (16838, 23811), "A4": (11906, 16838), "A5": (8391, 11906),
    "Letter": (12240, 15840), "Legal": (12240, 20160),
    "Executive": (10440, 15120), "Tabloid": (15840, 24480),
}

NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "cyan": "00FFFF", "magenta": "FF00FF",
    "gray": "808080", "grey": "808080", "lightgray": "D3D3D3", "silver": "C0C0C0",
    "orange": "FFA500", "purple": "800080", "navy": "000080", "darkblue": "00008B",
    "darkred": "8B0000", "darkgreen": "006400", "brown": "A52A2A", "pink": "FFC0CB",
}


def normalize_color(value):
    """'#003399' / '003399' / 'red' / 'auto' -> 'RRGGBB' or 'auto'."""
    s = str(value).strip().lstrip("#")
    if s.lower() == "auto":
        return "auto"
    if s.lower() in NAMED_COLORS:
        return NAMED_COLORS[s.lower()]
    if re.fullmatch(r"[0-9a-fA-F]{6}", s):
        return s.upper()
    if re.fullmatch(r"[0-9a-fA-F]{3}", s):
        return "".join(c * 2 for c in s).upper()
    raise ValueError("invalid colour: %r (use 003399, #003399 or a name like red)" % (value,))
