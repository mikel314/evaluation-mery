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
    Open the template, fill in the student name, insert up to 7 student photos
    (whichever files are present in the student's folder) and save to OUTPUT_DIR.
    Picture layout is fully driven by PICTURE_SPECS in config.py.
    """
    doc = open_doc(str(TEMPLATE_PATH))

    find_and_replace(doc, "NOM: ", f"NOM: {grades['student']}")

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

def main():
    print(f"Table path:    {TABLE_PATH}")
    print(f"Template path: {TEMPLATE_PATH}")
    print(f"Output dir:    {OUTPUT_DIR}\n")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    template_stem = TEMPLATE_PATH.stem

    df = load_grades()
    print()

    for _, row in df[0:5].iterrows():
        grades = parse_student_row(row)
        output_path = generate_report(grades, template_stem)
        print(f"Saved: {output_path.name}")

    print(f"\nDone — {len(df)} reports saved to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
