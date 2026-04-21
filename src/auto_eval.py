"""
auto_eval.py
Reads the grades table from an Excel file, generates one personalised .docx
report per student based on a template, and saves them to the output folder.

Usage:
    python src/auto_eval.py
"""

import re
import sys
from pathlib import Path

# Allow imports from project root regardless of where the script is invoked
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config.config import (
    TABLE_PATH, TEMPLATE_PATH, OUTPUT_DIR, PICTURES_DIR,
    PICTURE_SPECS,
)
from src.xlsx_utils import get_file_info, read_sheet, get_schema
from src.docx_utils import (
    open_doc, save_doc, find_and_replace, find_in_doc,
    find_student_pictures, insert_floating_image,
    apply_section_grades,
)


# Section columns in the grades file (order matches the template)
GRADE_COLUMNS = [
    "LES PROPOTES DE TREBALL",
    "ACTIVITATS INTERNIVELL",
    "ENDEVINALLES DELS ANIMALS",
    "ANGLÈS",
    "ENTUSIASMAT",
    "PSICOMOTRICITAT I EMAT",
    "TREBALL PER PROJECTES: HORT ESCOLAR",
    "ACOMIADAMENT",
]


def _safe_grade(val):
    """Return val as int (1/2/3) or None if missing/invalid."""
    try:
        v = int(float(val))
        return v if v in (1, 2, 3) else None
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_grades():
    """
    Inspect and load the grades spreadsheet.
    Returns a DataFrame with one row per student.
    """
    info = get_file_info(TABLE_PATH)
    print("Workbook info:")
    for sheet_name, meta in info["sheets"].items():
        print(f"  {sheet_name!r} — {meta['rows']} rows x {meta['cols']} columns")

    df = read_sheet(TABLE_PATH, sheet=0)
    print(f"\nLoaded {len(df)} students, {len(df.columns)} columns")
    #print(get_schema(df).to_string(index=False))
    print(df['Estudiant'])
    return df


# ---------------------------------------------------------------------------
# Row parsing
# ---------------------------------------------------------------------------

def parse_student_row(row) -> dict:
    """
    Unpack all columns of a DataFrame row into a named dictionary.
    Returns a dict with one key per grade category.
    """
    section_grades = {
        col: g
        for col in GRADE_COLUMNS
        if (g := _safe_grade(row.get(col))) is not None
    }
    return {
        "student":        row["Estudiant"],
        "cognoms":        row.get("Cognoms", "") or "",
        "section_grades": section_grades,
    }


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def generate_report(grades: dict, template_stem: str) -> Path:
    """
    Open the template, fill in the student name, insert up to 7 student photos
    (whichever files are present in the student's folder) and save to OUTPUT_DIR.
    Picture layout is fully driven by PICTURE_SPECS in config.py.
    """
    doc = open_doc(str(TEMPLATE_PATH))

    full_name = grades["student"]
    if grades["cognoms"].strip():
        full_name = f"{grades['student']} {grades['cognoms'].strip()}"
    find_and_replace(doc, "NOM: ", f"NOM: {full_name}")
    apply_section_grades(doc, grades["section_grades"])

    pictures = find_student_pictures(PICTURES_DIR, grades["student"])
    if not pictures:
        print(f"  [!] No pictures found for {grades['student']}")

    for spec in PICTURE_SPECS:
        pic_path = pictures.get(spec["stem"])
        if not pic_path:
            continue

        dist_kwargs = {
            "dist_t": spec.get("dist_t", 0.0),
            "dist_b": spec.get("dist_b", 0.0),
            "dist_l": spec.get("dist_l", 0.0),
            "dist_r": spec.get("dist_r", 0.0),
        }
        border_kwargs = {
            "border_color": spec.get("border_color"),
            "border_pt":    spec.get("border_pt", 1.0),
        }

        if spec["section"] is None:
            # PORT: absolute position on first page
            insert_floating_image(
                doc, pic_path,
                x_inches=spec["x_in"],
                y_inches=spec["y_in"],
                width_inches=spec["width_in"],
                wrap_type=spec["wrap_type"],
                v_relative=spec["v_relative"],
                **dist_kwargs,
                **border_kwargs,
            )
        else:
            anchors = find_in_doc(doc, re.escape(spec["section"]))
            if not anchors:
                print(f"  [!] Section not found for {spec['stem']!r}: {spec['section']!r}")
                continue
            # Anchor to the paragraph *after* the heading so the image appears
            # below the section title rather than overlapping it.
            all_paras = doc.paragraphs
            heading_para = anchors[0]
            heading_idx = next(
                (i for i, p in enumerate(all_paras) if p._element is heading_para._element),
                None,
            )
            anchor_para = (
                all_paras[heading_idx + 1]
                if heading_idx is not None and heading_idx + 1 < len(all_paras)
                else heading_para
            )
            insert_floating_image(
                doc, pic_path,
                x_inches=spec["x_in"],
                y_inches=spec["y_in"],
                width_inches=spec["width_in"],
                wrap_type=spec["wrap_type"],
                anchor_paragraph=anchor_para,
                v_relative=spec["v_relative"],
                **dist_kwargs,
                **border_kwargs,
            )

    output_path = OUTPUT_DIR / f"{template_stem}_{grades['student']}.docx"
    save_doc(doc, str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _wsl_to_windows_path(path: Path) -> str:
    """Convert a WSL /mnt/c/... path to a Windows C:\\... path."""
    s = str(path)
    if s.startswith("/mnt/") and len(s) > 6:
        drive = s[5].upper()
        rest = s[6:].replace("/", "\\")
        return f"{drive}:{rest}"
    return s


def convert_reports_to_pdf() -> None:
    """Convert every .docx in OUTPUT_DIR to PDF using Microsoft Word via PowerShell."""
    import subprocess
    docx_files = sorted(OUTPUT_DIR.glob("*.docx"))
    if not docx_files:
        print("No .docx files found in output directory.")
        return

    print(f"Output dir: {OUTPUT_DIR}")
    print(f"Converting {len(docx_files)} file(s) to PDF using Microsoft Word...\n")

    win_paths = [_wsl_to_windows_path(p) for p in docx_files]
    files_array = ", ".join(f"'{p}'" for p in win_paths)

    ps_script = f"""
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$files = @({files_array})
foreach ($f in $files) {{
    $pdf = [System.IO.Path]::ChangeExtension($f, '.pdf')
    $doc = $word.Documents.Open($f)
    $doc.SaveAs([ref]$pdf, [ref]17)
    $doc.Close()
    Write-Host "Converted: $(Split-Path $f -Leaf)"
}}
$word.Quit()
"""
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command", ps_script],
        text=True,
    )
    if result.returncode != 0:
        print(f"PowerShell exited with code {result.returncode}")
    else:
        print(f"\nDone — PDFs saved to {OUTPUT_DIR}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Auto-eval report generator")
    parser.add_argument(
        "--step",
        choices=["reports", "pdfs"],
        default="reports",
        help="reports = generate .docx reports (default)   pdfs = convert .docx to PDF",
    )
    args = parser.parse_args()

    if args.step == "pdfs":
        convert_reports_to_pdf()
        return

    # --- Step: reports ---
    print(f"Table path:    {TABLE_PATH}")
    print(f"Template path: {TEMPLATE_PATH}")
    print(f"Output dir:    {OUTPUT_DIR}\n")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    template_stem = TEMPLATE_PATH.stem

    df = load_grades()
    print()

    for _, row in df[:].iterrows():
        grades = parse_student_row(row)
        output_path = generate_report(grades, template_stem)
        print(f"Saved: {output_path.name}")

    print(f"\nDone — {len(df)} reports saved to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
