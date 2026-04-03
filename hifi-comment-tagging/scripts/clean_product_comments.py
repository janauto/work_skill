#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path

from openpyxl import Workbook

from workbook_utils import collect_rows_for_product, normalize_text


OUTPUT_COLUMNS = [
    "record_id",
    "product_name",
    "source_type",
    "source_sheet",
    "source_row",
    "return_time",
    "store",
    "country",
    "raw_comment",
    "translated_comment",
    "return_reason",
    "cleaned_comment",
    "is_valid_feedback",
    "level_1",
    "level_2",
    "level_3",
    "level_4",
    "severity",
    "is_user_reason",
    "is_quality_risk",
]


def dedupe_rows(rows):
    deduped = []
    seen = set()
    dropped = 0
    for row in rows:
        if not row["is_valid_feedback"]:
            continue
        # Do not dedupe rows whose only usable text is the standardized return reason.
        # Those rows are distinct orders even when the reason code string repeats.
        has_free_text = bool(normalize_text(row.get("translated_comment")) or normalize_text(row.get("raw_comment")))
        if not has_free_text and normalize_text(row.get("return_reason")):
            deduped.append(row)
            continue
        key = (row["product_name"], normalize_text(row["cleaned_comment"]))
        if key in seen:
            dropped += 1
            continue
        seen.add(key)
        deduped.append(row)
    return deduped, dropped


def build_output_workbook(rows, metadata, output_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "CleanedComments"
    ws.append(OUTPUT_COLUMNS)
    for row in rows:
        ws.append([row.get(column, "") for column in OUTPUT_COLUMNS])

    meta = wb.create_sheet("Metadata")
    meta.append(["field", "value"])
    for key, value in metadata.items():
        if isinstance(value, (list, dict)):
            value = json.dumps(value, ensure_ascii=False)
        meta.append([key, value])

    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Clean target-product comments from a workbook.")
    parser.add_argument("workbook", help="Path to the xlsx workbook")
    parser.add_argument("--product", required=True, help="Target product, for example P4 or ZD3")
    parser.add_argument("--sheet", help="Optional explicit sheet name")
    parser.add_argument(
        "--output",
        help="Output xlsx path. Defaults to <product>_cleaned_comments.xlsx next to the workbook.",
    )
    args = parser.parse_args()

    workbook_path = Path(args.workbook).resolve()
    rows, profiles = collect_rows_for_product(workbook_path, args.product, args.sheet)
    if not rows:
        raise SystemExit(
            f"No rows matched product '{args.product}'. Run profile_workbook.py first to inspect candidate sheets."
        )

    raw_count = len(rows)
    invalid_count = sum(1 for row in rows if not row["is_valid_feedback"])
    rows, duplicate_count = dedupe_rows(rows)
    by_sheet = Counter(row["source_sheet"] for row in rows)

    output_path = (
        Path(args.output).resolve()
        if args.output
        else workbook_path.with_name(f"{args.product.upper()}_cleaned_comments.xlsx")
    )
    metadata = {
        "source_workbook": str(workbook_path),
        "target_product": args.product.upper(),
        "valid_feedback_rule": "return_workbooks_require_free_text_comment",
        "matched_rows_before_dedupe": raw_count,
        "invalid_rows_dropped": invalid_count,
        "duplicate_rows_dropped": duplicate_count,
        "clean_rows": len(rows),
        "matched_sheets": dict(by_sheet),
        "profiled_sheets": [profile["sheet"] for profile in profiles],
    }
    build_output_workbook(rows, metadata, output_path)
    print(json.dumps({"output": str(output_path), **metadata}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
