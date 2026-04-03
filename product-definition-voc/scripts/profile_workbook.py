#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from voc_workbook_utils import collect_focus_candidates, load_workbook_safe, sheet_profile


def main():
    parser = argparse.ArgumentParser(description="Profile a VOC workbook for product-definition analysis.")
    parser.add_argument("workbook", help="Path to the xlsx workbook")
    parser.add_argument("--json", action="store_true", help="Print JSON instead of plain text")
    args = parser.parse_args()

    workbook_path = Path(args.workbook).resolve()
    wb = load_workbook_safe(workbook_path, data_only=True)
    profiles = [sheet_profile(ws, workbook_path) for ws in wb.worksheets]
    output = {
        "workbook": str(workbook_path),
        "focus_candidates": collect_focus_candidates(workbook_path),
        "sheets": profiles,
    }

    if args.json:
        print(json.dumps(output, ensure_ascii=False, indent=2))
        return

    print(f"Workbook: {workbook_path}")
    print("Focus candidates:", " | ".join(output["focus_candidates"]) or "None detected")
    print()
    for profile in profiles:
        print(f"Sheet: {profile['sheet']}")
        print(f"  Candidate sheet: {profile['candidate']}")
        print(f"  Header row: {profile['header_row']}")
        print(f"  Source type: {profile['source_type']}")
        print(f"  Image columns: {profile['image_columns']}")
        print(f"  Column map: {json.dumps(profile['column_map'], ensure_ascii=False)}")
        print()


if __name__ == "__main__":
    main()
