#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
DEFAULT_DATASET_ROOT = SKILL_ROOT / "runs" / "default"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the full ASR Review scraping pipeline.")
    parser.add_argument("--dataset-root", default=str(DEFAULT_DATASET_ROOT), help="Dataset directory for URLs, raw threads, images, and workbook outputs.")
    parser.add_argument("--urls-file", help="Optional thread URL list file. Defaults to <dataset-root>/curated_threads.txt.")
    parser.add_argument("--skip-fetch", action="store_true", help="Skip thread fetch and only rebuild the workbook.")
    parser.add_argument("--skip-workbook", action="store_true", help="Skip workbook generation and only fetch threads.")
    parser.add_argument("--chrome-cache-dir", help="Optional Chrome cache dir override for image extraction.")
    return parser.parse_args()


def run_command(cmd: list[str]) -> None:
    print(">", " ".join(cmd))
    subprocess.run(cmd, check=True)


def main() -> int:
    args = parse_args()
    dataset_root = Path(args.dataset_root).expanduser().resolve()
    dataset_root.mkdir(parents=True, exist_ok=True)

    urls_file = Path(args.urls_file).expanduser().resolve() if args.urls_file else dataset_root / "curated_threads.txt"
    if not args.skip_fetch and not urls_file.exists():
        raise SystemExit(f"Thread URL list not found: {urls_file}")

    if not args.skip_fetch:
        fetch_cmd = [
            sys.executable,
            str(SCRIPT_DIR / "fetch_asr_threads.py"),
            "--dataset-root",
            str(dataset_root),
            "--urls-file",
            str(urls_file),
        ]
        run_command(fetch_cmd)

    if not args.skip_workbook:
        workbook_cmd = [
            sys.executable,
            str(SCRIPT_DIR / "build_asr_workbook.py"),
            "--dataset-root",
            str(dataset_root),
        ]
        if args.chrome_cache_dir:
            workbook_cmd.extend(["--chrome-cache-dir", str(Path(args.chrome_cache_dir).expanduser().resolve())])
        run_command(workbook_cmd)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
