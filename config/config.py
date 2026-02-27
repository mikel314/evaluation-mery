"""
config.py
Central configuration for project paths.
All paths are resolved relative to the project root so the project
works regardless of where it is cloned.
"""

from pathlib import Path

CM_PER_INCH = 2.54

# ---------------------------------------------------------------------------
# Project root (one level up from this file)
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent

# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------
DATA_DIR       = ROOT / "data"
RAW_DIR        = DATA_DIR / "raw"
PROCESSED_DIR  = DATA_DIR / "processed"


INPUTS_DIR     = Path("/mnt/c/Users/mikel/projects/Evaluacion Mery")
TABLE_PATH     = INPUTS_DIR / "grades_v0.xlsx"
TEMPLATE_PATH  = INPUTS_DIR / "informe 4 anys FEBRER curs 25-26.docx"
PICTURES_DIR   = INPUTS_DIR / "Fotos individuals"
OUTPUT_DIR     = INPUTS_DIR / "output"

# ---------------------------------------------------------------------------
# Picture layout
# All x/y positions are in inches from the top-left corner of the page.
# ---------------------------------------------------------------------------

# --- PORT: first-page portrait, centered, in front of text ---
_PORT_WIDTH_IN      = 4.25          # half page width (8.5 / 2)
_PORT_X_IN          = 2.125         # (8.5 - _PORT_WIDTH_IN) / 2
_PORT_Y_IN          = 2.4           # from top of page
_PORT_BORDER_COLOR  = "000000"      # hex RGB, black
_PORT_BORDER_PT     = 3.0           # line weight in points

# --- Section pictures (PROP … HORT): right side, square wrap ---
_SEC_WIDTH_CM   = 5
_SEC_WIDTH_IN   = round(_SEC_WIDTH_CM  / CM_PER_INCH, 4)   # ≈ 2.19 in
_SEC_X_IN       = round(8.5 - 1.0 - _SEC_WIDTH_IN,  4)    # flush with right margin
_SEC_Y_IN       = round(0.7 / CM_PER_INCH, 4)              # 0.5 cm below anchor paragraph
_SEC_WRAP       = "square"

# --- Distance from text for section pictures (inches) ---
_SEC_DIST_T     = 0.2
_SEC_DIST_B     = 0.2
_SEC_DIST_L     = 0.2
_SEC_DIST_R     = 0.2

# ---------------------------------------------------------------------------
# PICTURE_SPECS — one entry per picture slot, in document order.
#
# Keys:
#   stem       — uppercase filename stem to look for in the student folder
#   section    — exact text of the section title to anchor to
#                (None = first page, absolute positioning)
#   width_in   — image width in inches (height is derived from aspect ratio)
#   x_in       — horizontal position in inches from left page edge
#   y_in       — vertical position in inches (meaning depends on v_relative)
#   wrap_type  — "front" | "behind" | "square" | "tight" | "top_bottom"
#   v_relative — "page" (y from page top) | "paragraph" (y from anchor para)
# ---------------------------------------------------------------------------
PICTURE_SPECS = [
    {
        "stem":         "PORT",
        "section":      None,
        "width_in":     _PORT_WIDTH_IN,
        "x_in":         _PORT_X_IN,
        "y_in":         _PORT_Y_IN,
        "wrap_type":    "front",
        "v_relative":   "page",
        "border_color": _PORT_BORDER_COLOR,
        "border_pt":    _PORT_BORDER_PT,
    },
    {
        "stem":       "PROP",
        "section":    "LES PROPOTES DE TREBALL",
        "width_in":   _SEC_WIDTH_IN,
        "x_in":       _SEC_X_IN,
        "y_in":       _SEC_Y_IN,
        "wrap_type":  _SEC_WRAP,
        "v_relative": "paragraph",
        "dist_t":     _SEC_DIST_T,
        "dist_b":     _SEC_DIST_B,
        "dist_l":     _SEC_DIST_L,
        "dist_r":     _SEC_DIST_R,
    },
    {
        "stem":       "INTER",
        "section":    "ACTIVITATS INTERNIVELL",
        "width_in":   _SEC_WIDTH_IN,
        "x_in":       _SEC_X_IN,
        "y_in":       _SEC_Y_IN,
        "wrap_type":  _SEC_WRAP,
        "v_relative": "paragraph",
        "dist_t":     _SEC_DIST_T,
        "dist_b":     _SEC_DIST_B,
        "dist_l":     _SEC_DIST_L,
        "dist_r":     _SEC_DIST_R,
    },
    {
        "stem":       "ENDE",
        "section":    "ENDEVINALLES DELS ANIMALS",
        "width_in":   _SEC_WIDTH_IN,
        "x_in":       _SEC_X_IN,
        "y_in":       _SEC_Y_IN,
        "wrap_type":  _SEC_WRAP,
        "v_relative": "paragraph",
        "dist_t":     _SEC_DIST_T,
        "dist_b":     _SEC_DIST_B,
        "dist_l":     _SEC_DIST_L,
        "dist_r":     _SEC_DIST_R,
    },
    {
        "stem":       "ENTU",
        "section":    "ENTUSIASMAT",
        "width_in":   _SEC_WIDTH_IN,
        "x_in":       _SEC_X_IN,
        "y_in":       _SEC_Y_IN,
        "wrap_type":  _SEC_WRAP,
        "v_relative": "paragraph",
        "dist_t":     _SEC_DIST_T,
        "dist_b":     _SEC_DIST_B,
        "dist_l":     _SEC_DIST_L,
        "dist_r":     _SEC_DIST_R,
    },
    {
        "stem":       "PSICO",
        "section":    "PSICOMOTRICITAT I EMAT",
        "width_in":   _SEC_WIDTH_IN,
        "x_in":       _SEC_X_IN,
        "y_in":       _SEC_Y_IN,
        "wrap_type":  _SEC_WRAP,
        "v_relative": "paragraph",
        "dist_t":     _SEC_DIST_T,
        "dist_b":     _SEC_DIST_B,
        "dist_l":     _SEC_DIST_L,
        "dist_r":     _SEC_DIST_R,
    },
    {
        "stem":       "HORT",
        "section":    "TREBALL PER PROJECTES: HORT ESCOLAR",
        "width_in":   _SEC_WIDTH_IN,
        "x_in":       _SEC_X_IN,
        "y_in":       _SEC_Y_IN,
        "wrap_type":  _SEC_WRAP,
        "v_relative": "paragraph",
        "dist_t":     _SEC_DIST_T,
        "dist_b":     _SEC_DIST_B,
        "dist_l":     _SEC_DIST_L,
        "dist_r":     _SEC_DIST_R,
    },
]
