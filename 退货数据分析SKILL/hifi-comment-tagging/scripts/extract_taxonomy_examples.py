#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path

from workbook_utils import load_workbook_safe, sheet_profile, row_value


def extract_from_workbook(path: Path):
    wb = load_workbook_safe(path, data_only=True)
    chains = Counter()
    examples = defaultdict(list)

    for ws in wb.worksheets:
        profile = sheet_profile(ws, path)
        column_map = profile["column_map"]
        if not any(key in column_map for key in ("level_1", "level_2", "level_3", "level_4")):
            continue
        header_row = profile["header_row"]
        for row_idx in range(header_row + 1, ws.max_row + 1):
            chain = tuple(
                row_value(ws, row_idx, column_map, key) for key in ("level_1", "level_2", "level_3", "level_4")
            )
            if not any(chain):
                continue
            chains[chain] += 1
            if len(examples[chain]) < 3:
                sample_text = (
                    row_value(ws, row_idx, column_map, "translated_comment")
                    or row_value(ws, row_idx, column_map, "raw_comment")
                    or row_value(ws, row_idx, column_map, "return_reason")
                )
                examples[chain].append(
                    {
                        "workbook": path.name,
                        "sheet": ws.title,
                        "row": row_idx,
                        "text": sample_text,
                    }
                )
    return chains, examples


def build_markdown(paths):
    merged_counter = Counter()
    merged_examples = defaultdict(list)
    for path in paths:
        chains, examples = extract_from_workbook(path)
        merged_counter.update(chains)
        for chain, items in examples.items():
            existing = merged_examples[chain]
            for item in items:
                if len(existing) < 3:
                    existing.append(item)

    lines = ["# Taxonomy Examples", ""]
    for chain, count in merged_counter.most_common():
        chain_text = " -> ".join(part for part in chain if part) or "未标注"
        lines.append(f"## {chain_text}")
        lines.append(f"- Count: {count}")
        for item in merged_examples[chain]:
            lines.append(
                f"- Example: `{item['workbook']}::{item['sheet']}#{item['row']}` {item['text'] or '[no text]'}"
            )
        lines.append("")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Extract taxonomy chains and examples from labeled workbooks.")
    parser.add_argument("paths", nargs="+", help="Workbook paths")
    parser.add_argument("--output", help="Write markdown or json output to a file")
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of markdown")
    args = parser.parse_args()

    workbook_paths = [Path(path).resolve() for path in args.paths]
    if args.json:
        payload = {}
        for path in workbook_paths:
            chains, examples = extract_from_workbook(path)
            payload[str(path)] = {
                "chains": [
                    {"chain": list(chain), "count": count, "examples": examples[chain]}
                    for chain, count in chains.most_common()
                ]
            }
        output = json.dumps(payload, ensure_ascii=False, indent=2)
    else:
        output = build_markdown(workbook_paths)

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
    else:
        print(output)


if __name__ == "__main__":
    main()
