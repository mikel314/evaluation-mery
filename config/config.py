"""
config.py
Central configuration for project paths.
All paths are resolved relative to the project root so the project
works regardless of where it is cloned.
"""

from pathlib import Path

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
OUTPUT_DIR     = INPUTS_DIR / "output"


