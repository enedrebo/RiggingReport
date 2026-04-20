import sys
import os

from excel_reader import read_excel
from report_template import build_pdf

# ── Input / output files ──────────────────────────────────────────────────────
EXCEL_FILE  = "RiggingCalculations_template.xlsx"
OUTPUT_FILE = "output_report.pdf"
# ─────────────────────────────────────────────────────────────────────────────


def main():
    excel_path = EXCEL_FILE
    if not os.path.isfile(excel_path):
        print(f"Error: file not found: {excel_path}")
        sys.exit(1)

    output_path = OUTPUT_FILE

    print(f"Reading:       {excel_path}")
    try:
        loads_data, items_data = read_excel(excel_path)
    except ValueError as e:
        print(f"Error reading Excel: {e}")
        sys.exit(1)

    print(f"Generating:    {output_path}")
    build_pdf(loads_data, items_data, output_path)
    print(f"Done:          {output_path}")


if __name__ == "__main__":
    main()
