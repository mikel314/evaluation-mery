"""
auto_eval.py
Reads the grades table from an Excel file, generates one personalised .docx
report per student based on a template, and saves them to the output folder.

Usage:
    python src/auto_eval.py
"""

import sys
from pathlib import Path

# Allow imports from project root regardless of where the script is invoked
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config.config import TABLE_PATH, TEMPLATE_PATH, OUTPUT_DIR, PICTURES_DIR
from src.xlsx_utils import get_file_info, read_sheet, get_schema
from src.docx_utils import (
    open_doc, save_doc, find_and_replace,
    find_student_picture, insert_floating_image,
)


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
    print(get_schema(df).to_string(index=False))
    return df


# ---------------------------------------------------------------------------
# Row parsing
# ---------------------------------------------------------------------------

def parse_student_row(row) -> dict:
    """
    Unpack all columns of a DataFrame row into a named dictionary.
    Returns a dict with one key per grade category.
    """
    return {
        # Identity
        "student":             row["Estudiant"],
        # Language / Writing
        "writes_own_name":     row["Escriu el seu nom sense model"],
        "phonics":             row["(vocals + P L M D G F C)"],
        "copies_from_board":   row["Copia paraules de la pissarra"],
        "dictation":           row["Realitza dictat lletra a lletra"],
        # Reading & Communication
        "interest_in_reading": row["Mostra interès per la lectura de literatura infantil"],
        "oral_communication":  row["Participa en situacions de comunicació oral"],
        # Technology
        "use_of_devices":      row["Fa un bon us dels dispositius tecnologics (pissarra, tablet, blue-bot)"],
        # Arts
        "james_rizzi":         row["James Rizzi"],
        "drawing":             row["Realitza dibuixos adequats a l'edat"],
        "use_of_tools":        row["Utilitza correctament les diferents eines (pinzells, tisores, etc.)"],
    }


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def generate_report(grades: dict, template_stem: str) -> Path:
    """
    Open the template, fill in the student name, insert the student photo
    centred on the first page at 1/3 of the page width (floating, behind text),
    and save the personalised report to OUTPUT_DIR.
    Returns the path of the saved file.
    """
    doc = open_doc(str(TEMPLATE_PATH))

    find_and_replace(doc, "NOM: ", f"NOM: {grades['student']}")

    picture_path = find_student_picture(PICTURES_DIR, grades["student"])
    if picture_path:
        section    = doc.sections[0]
        page_w     = section.page_width.inches
        page_h     = section.page_height.inches
        img_width  = page_w / 2

        from PIL import Image as _Image
        with _Image.open(picture_path) as img:
            img_w, img_h = img.size
        img_height = img_width * img_h / img_w

        x = (page_w - img_width)  / 2
        y = (page_h - img_height) / 2

        insert_floating_image(doc, picture_path, x_inches=x, y_inches=y, width_inches=img_width)
    else:
        print(f"  [!] No picture found for {grades['student']}")

    output_path = OUTPUT_DIR / f"{template_stem}_{grades['student']}.docx"
    save_doc(doc, str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print(f"Table path:    {TABLE_PATH}")
    print(f"Template path: {TEMPLATE_PATH}")
    print(f"Output dir:    {OUTPUT_DIR}\n")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    template_stem = TEMPLATE_PATH.stem

    df = load_grades()
    print()

    for _, row in df.iterrows():
        grades = parse_student_row(row)
        output_path = generate_report(grades, template_stem)
        print(f"Saved: {output_path.name}")

    print(f"\nDone — {len(df)} reports saved to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
