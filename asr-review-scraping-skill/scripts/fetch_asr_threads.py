#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_DATASET_ROOT = SCRIPT_DIR.parent / "runs" / "default"
ROOT = DEFAULT_DATASET_ROOT
URLS_FILE = ROOT / "curated_threads.txt"
RAW_DIR = ROOT / "raw_threads"
INDEX_FILE = ROOT / "thread_index.json"
SUMMARY_FILE = ROOT / "thread_summary.md"


def configure_paths(dataset_root: Path) -> None:
    global ROOT, URLS_FILE, RAW_DIR, INDEX_FILE, SUMMARY_FILE
    ROOT = dataset_root.resolve()
    URLS_FILE = ROOT / "curated_threads.txt"
    RAW_DIR = ROOT / "raw_threads"
    INDEX_FILE = ROOT / "thread_index.json"
    SUMMARY_FILE = ROOT / "thread_summary.md"


def display_path(path: Path) -> str:
    try:
        return str(path.relative_to(ROOT))
    except ValueError:
        return str(path)


def read_urls(path: Path) -> list[str]:
    urls: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        urls.append(line)
    return urls


def slug_from_url(url: str) -> tuple[str, str]:
    match = re.search(r"threads/([^./]+(?:\.[^/]+)?)\.(\d+)/?$", url)
    if not match:
        return ("unknown", re.sub(r"[^a-z0-9]+", "-", url.lower()).strip("-")[:80])
    slug = match.group(1)
    thread_id = match.group(2)
    safe_slug = re.sub(r"[^a-z0-9]+", "-", slug.lower()).strip("-")
    return thread_id, safe_slug or "thread"


def fetch_markdown(url: str, retries: int = 4) -> str:
    jina_url = f"https://r.jina.ai/http://{url}"
    last_exc: Exception | None = None
    for attempt in range(retries + 1):
        req = Request(
            jina_url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "text/plain, text/markdown;q=0.9, */*;q=0.8",
            },
        )
        try:
            with urlopen(req, timeout=120) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except HTTPError as exc:
            last_exc = exc
            if exc.code != 429 or attempt >= retries:
                raise
            time.sleep(3 * (attempt + 1))
        except (URLError, TimeoutError) as exc:
            last_exc = exc
            if attempt >= retries:
                raise
            time.sleep(2 * (attempt + 1))
    assert last_exc is not None
    raise last_exc


def extract_title(markdown: str, fallback: str) -> str:
    for line in markdown.splitlines():
        if line.startswith("Title: "):
            return line[len("Title: ") :].strip()
    for line in markdown.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return fallback


def classify(url: str, title: str) -> list[str]:
    text = f"{url} {title}".lower()
    labels: list[str] = []

    if any(k in text for k in ("switcher", "a-b-switch", "switching-box", "switch-box", "speaker-switch", "selector")):
        labels.append("analog_switcher")
    if any(k in text for k in ("toslink", "optical", "spdif", "aes-ebu", "hdmi-3x1", "auto-input-switching")):
        labels.append("digital_switcher")
    if any(k in text for k in ("preamp", "2-preamps", "schiit-sys", "passive-preamp")):
        labels.append("switching_preamp")
    if any(k in text for k in ("avr", "av-amplifier-and-active-crossover", "soundcard-avr-as-preamp")):
        labels.append("avr_switching")
    if any(k in text for k in ("audio-hub", "source-control-hub", "all-in-one-device", "multiple-sources", "setup-confusion", "bp50")):
        labels.append("audio_hub")

    seen = set()
    deduped: list[str] = []
    for label in labels:
        if label not in seen:
            deduped.append(label)
            seen.add(label)
    return deduped


def estimate_post_count(markdown: str) -> int:
    return len(re.findall(r"^\*\s+\[#\d+\]", markdown, flags=re.MULTILINE))


def write_summary(items: list[dict], summary_path: Path) -> None:
    topic_map: dict[str, list[dict]] = {}
    for item in items:
        for label in item["labels"] or ["unclassified"]:
            topic_map.setdefault(label, []).append(item)

    lines = [
        "# ASR Review thread archive",
        "",
        f"- Fetched threads: {len(items)}",
        f"- Generated at: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]

    for label in sorted(topic_map):
        lines.append(f"## {label}")
        lines.append("")
        for item in topic_map[label]:
            lines.append(f"- [{item['title']}]({item['url']})")
            lines.append(f"  - local: `{item['file']}`")
            lines.append(f"  - estimated_posts: {item['estimated_posts']}")
        lines.append("")

    summary_path.write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch ASR threads via r.jina.ai mirror into a local dataset folder.")
    parser.add_argument("--dataset-root", default=str(DEFAULT_DATASET_ROOT), help="Dataset directory containing curated_threads.txt and generated outputs.")
    parser.add_argument("--urls-file", help="Override URL list file. Defaults to <dataset-root>/curated_threads.txt.")
    parser.add_argument("--raw-dir", help="Override raw thread markdown output dir. Defaults to <dataset-root>/raw_threads.")
    parser.add_argument("--index-file", help="Override JSON index output path. Defaults to <dataset-root>/thread_index.json.")
    parser.add_argument("--summary-file", help="Override summary markdown output path. Defaults to <dataset-root>/thread_summary.md.")
    parser.add_argument("--sleep-seconds", type=float, default=0.4, help="Delay between successful thread fetches.")
    return parser.parse_args()


def main() -> int:
    global URLS_FILE, RAW_DIR, INDEX_FILE, SUMMARY_FILE

    args = parse_args()
    dataset_root = Path(args.dataset_root).expanduser().resolve()
    configure_paths(dataset_root)

    if args.urls_file:
        URLS_FILE = Path(args.urls_file).expanduser().resolve()
    if args.raw_dir:
        RAW_DIR = Path(args.raw_dir).expanduser().resolve()
    if args.index_file:
        INDEX_FILE = Path(args.index_file).expanduser().resolve()
    if args.summary_file:
        SUMMARY_FILE = Path(args.summary_file).expanduser().resolve()

    RAW_DIR.mkdir(parents=True, exist_ok=True)
    INDEX_FILE.parent.mkdir(parents=True, exist_ok=True)
    SUMMARY_FILE.parent.mkdir(parents=True, exist_ok=True)

    if not URLS_FILE.exists():
        raise SystemExit(f"URL list not found: {URLS_FILE}")

    urls = read_urls(URLS_FILE)
    items: list[dict] = []

    for idx, url in enumerate(urls, start=1):
        thread_id, slug = slug_from_url(url)
        out_file = RAW_DIR / f"{idx:02d}_{thread_id}_{slug}.md"
        print(f"[{idx}/{len(urls)}] {url}")
        try:
            markdown = fetch_markdown(url)
        except (HTTPError, URLError, TimeoutError) as exc:
            items.append(
                {
                    "url": url,
                    "title": slug,
                    "file": display_path(out_file),
                    "status": f"error: {exc}",
                    "labels": classify(url, slug),
                    "estimated_posts": 0,
                }
            )
            continue

        title = extract_title(markdown, slug)
        out_file.write_text(markdown, encoding="utf-8")
        items.append(
            {
                "url": url,
                "title": title,
                "file": display_path(out_file),
                "status": "ok",
                "labels": classify(url, title),
                "estimated_posts": estimate_post_count(markdown),
            }
        )
        time.sleep(args.sleep_seconds)

    INDEX_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    write_summary(items, SUMMARY_FILE)
    print(json.dumps({"threads": len(items), "index": str(INDEX_FILE), "summary": str(SUMMARY_FILE)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
