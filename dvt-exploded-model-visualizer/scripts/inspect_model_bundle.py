#!/usr/bin/env python3
"""Inventory a hardware visualization input folder and report readiness as JSON."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path


SOURCE_CAD = {".step", ".stp", ".iges", ".igs", ".x_t", ".x_b", ".jt"}
WEB_3D = {".glb", ".gltf"}
MESH_3D = {".obj", ".fbx", ".3mf", ".ply", ".stl"}
DRAWINGS = {".dxf", ".dwg", ".pdf"}
BOM = {".xlsx", ".xls", ".csv", ".tsv"}
DOCUMENTS = {".docx", ".doc", ".md", ".txt", ".ppt", ".pptx", ".pdf"}
IMAGES = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
EXISTING_OUTPUT = {".html", ".json"}
SKIP_DIRS = {".git", "node_modules", "__pycache__", ".venv", "venv"}


def classify(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    categories: list[str] = []
    for name, extensions in (
        ("source_cad", SOURCE_CAD),
        ("web_3d", WEB_3D),
        ("mesh_3d", MESH_3D),
        ("drawing", DRAWINGS),
        ("bom", BOM),
        ("document", DOCUMENTS),
        ("image", IMAGES),
        ("existing_output", EXISTING_OUTPUT),
    ):
        if suffix in extensions:
            categories.append(name)
    return categories


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="Project or input-bundle directory")
    parser.add_argument("--max-files", type=int, default=5000)
    args = parser.parse_args()

    root = args.path.expanduser().resolve()
    if not root.is_dir():
        parser.error(f"not a directory: {root}")

    files: list[dict[str, object]] = []
    placeholders: list[str] = []
    category_counts: dict[str, int] = {}

    for current, dirs, names in os.walk(root):
        dirs[:] = [name for name in dirs if name not in SKIP_DIRS]
        for name in names:
            if len(files) >= args.max_files:
                break
            path = Path(current) / name
            categories = classify(path)
            if not categories and not name.endswith(".icloud"):
                continue
            try:
                size = path.stat().st_size
            except OSError:
                size = -1
            relative = str(path.relative_to(root))
            placeholder = name.endswith(".icloud") or size <= 0
            if placeholder:
                placeholders.append(relative)
            for category in categories:
                category_counts[category] = category_counts.get(category, 0) + 1
            files.append(
                {
                    "path": relative,
                    "size_bytes": size,
                    "categories": categories,
                    "available": not placeholder,
                }
            )
        if len(files) >= args.max_files:
            break

    available = [item for item in files if item["available"]]
    has_source = any("source_cad" in item["categories"] for item in available)
    has_web = any("web_3d" in item["categories"] for item in available)
    has_mesh = any("mesh_3d" in item["categories"] for item in available)
    has_change_evidence = any(
        set(item["categories"]) & {"document", "drawing"} for item in available
    )

    if has_source:
        visualization_state = "actual-solid-source-available"
    elif has_web:
        visualization_state = "visualization-only-derived-model"
    elif has_mesh:
        visualization_state = "mesh-only-limited-engineering-evidence"
    else:
        visualization_state = "missing-3d-source"

    report = {
        "root": str(root),
        "visualization_state": visualization_state,
        "ready_for_actual_exploded_view": has_source,
        "ready_for_visualization_only": has_source or has_web or has_mesh,
        "change_evidence_file_present": has_change_evidence,
        "category_counts": category_counts,
        "placeholders_or_empty": placeholders,
        "missing_or_recommended": {
            "source_step_stp": not has_source,
            "controlled_change_requirement_or_user_instruction": not has_change_evidence,
            "bom_or_part_name_mapping": category_counts.get("bom", 0) == 0,
            "local_section_or_drawing": category_counts.get("drawing", 0) == 0,
            "reference_images": category_counts.get("image", 0) == 0,
        },
        "files": files,
        "truncated": len(files) >= args.max_files,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
