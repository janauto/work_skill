#!/usr/bin/env python3
"""Compose internal explanation sheets from CLI-produced CAD; never synthesize CAD.

The manifest supplies content only. Geometry, leaders and original reference numbers
are copied together with one uniform transform. All layout constants live here.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys

import ezdxf
from ezdxf import bbox
from ezdxf.enums import TextEntityAlignment
from ezdxf.fonts import fonts
from ezdxf.math import Matrix44

VERSION = "qwen-patent-review/1"
PAGE = (420.0, 297.0)
SOURCE_LAYERS = {"GEOM", "HIDDEN", "LEADER", "NUM"}
SUPPORTED = {"LINE", "LWPOLYLINE", "POLYLINE", "CIRCLE", "ARC", "ELLIPSE", "SPLINE", "TEXT"}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def claim(value):
    require(isinstance(value, dict), "说明须包含 text / status / evidence")
    text = value.get("text", "").strip()
    require(bool(text), "功能、装配或概述不可空白")
    status = value.get("status")
    require(status in {"observed", "confirmed", "pending"}, "未知的证据状态")
    if status != "pending":
        evidence = value.get("evidence", [])
        require(isinstance(evidence, list) and bool(evidence)
                and all(isinstance(x, str) and x.strip() for x in evidence),
                "已观察/已确认的说明必须有来源")
    return {"observed": "观察：", "confirmed": "已核实：", "pending": "待核实："}[status] + text


def validate(data, base):
    require(data.get("schema") == VERSION, "不支持的说明图版本")
    require(set(data) <= {"schema", "reference_numerals", "parts", "sheets"}, "内容清单含未知字段；不接受自填排版参数")
    reference = load(base / data["reference_numerals"])
    require(reference.get("schema") == "patent-numerals/1", "不支持的源标记表")
    numerals = reference["numerals"]
    require(all(type(r["numeral"]) is int and r["numeral"] > 0 and r["term"].strip() for r in numerals), "源编号或名称无效")
    by_selector = {r["selector"]: r for r in numerals}
    require(len(by_selector) == len(numerals), "源标记表 selector 重复")
    require(len({r["numeral"] for r in numerals}) == len(numerals), "源标记表编号重复")
    parts, instances, selectors = {}, set(), set()
    for part in data["parts"]:
        require(set(part) <= {"id", "selector", "instance_ids", "function", "assembly", "name",
                              "unresolved_reason", "unlabelled_reason"}, "零件含未知字段；数量和编号必须由源数据计算")
        key = part["id"]
        require(key not in parts, "零件 id 重复")
        ids = part.get("instance_ids", [])
        require(isinstance(ids, list) and ids and all(isinstance(i, str) and i for i in ids),
                "每个表格条目须列出源模型实例 id；用量由实例数计算")
        require(len(ids) == len(set(ids)) and not instances.intersection(ids),
                "同一实例被重复计数或同时分配给不同零件")
        instances.update(ids)
        selector = part.get("selector")
        if selector is not None:
            require(selector in by_selector, "selector 不在源标记表中")
            require(selector not in selectors, "同名不同零件不可共用一个 selector；请用未编号条目记录限制")
            selectors.add(selector)
            source = by_selector[selector]
            name, number = source["term"], str(source["numeral"])
        else:
            name, number = part.get("name", "").strip(), "—"
            require(name and part.get("unresolved_reason", "").strip(), "未编号条目须有名称及原因")
        remark = "作用·" + claim(part["function"]) + "\n装配·" + claim(part["assembly"])
        reason = part.get("unresolved_reason") or part.get("unlabelled_reason")
        if reason:
            remark += "\n标注限制：" + reason
        parts[key] = dict(part, name=name, number=number, quantity=len(ids), remark=remark)
    require(parts and data.get("sheets"), "零件与图纸不能为空")
    seen = set()
    for sheet in data["sheets"]:
        require(set(sheet) <= {"id", "title", "source_dxf", "overview", "paragraphs", "rows"}, "图纸含未知字段；排版由程序计算")
        require(re.fullmatch(r"[A-Za-z0-9_-]+", sheet["id"]) is not None, "图纸 id 必须是安全文件名")
        require(sheet["id"] not in seen, "图纸 id 重复")
        seen.add(sheet["id"])
        require(sheet.get("title", "").strip(), "图纸缺少标题")
        rows = sheet["rows"]
        require(rows and len(rows) == len(set(rows)) and all(x in parts for x in rows), "表格行缺失、重复或引用未知零件")
        claim(sheet["overview"])
        require(sheet.get("paragraphs"), "缺少功能/装配说明")
        for paragraph in sheet["paragraphs"]:
            claim(paragraph)
        require((base / sheet["source_dxf"]).is_file(), "源 DXF 不存在")
    return parts


def wrap(text, width, font, height):
    face = fonts.make_font(str(font), height)
    lines, line = [], ""
    for char in text:
        if char == "\n":
            lines.append(line)
            line = ""
            continue
        require(face.text_width(char) <= width, "单字超出列宽")
        if line and face.text_width(line + char) > width:
            lines.append(line)
            line = ""
        line += char
    lines.append(line)
    return lines


def compose(sheet, parts, source_path, font, page_number):
    source = ezdxf.readfile(source_path)
    entities = [e for e in source.modelspace() if e.dxf.layer in SOURCE_LAYERS]
    require(any(e.dxf.layer in {"GEOM", "HIDDEN"} for e in entities), "源 DXF 缺少 CAD 几何层")
    require(all(e.dxftype() in SUPPORTED for e in entities), "源几何含不支持的实体；不可静默丢弃")
    require(all(e.dxftype() != "TEXT" or e.dxf.layer == "NUM" for e in entities), "源几何含文本，请先核对来源")
    numbers = {e.dxf.text for e in entities if e.dxftype() == "TEXT" and e.dxf.layer == "NUM"}
    rows = [parts[k] for k in sheet["rows"]]
    require(numbers.issubset({r["number"] for r in rows}), "图上存在表格未解释的编号")
    require(all(r["number"] == "—" or r["number"] in numbers or r.get("unlabelled_reason") for r in rows),
            "表格引用图上不存在的编号；须给出 unlabelled_reason 并在交付中复核")

    doc = ezdxf.new("R2018")
    doc.header["$INSUNITS"] = 4
    doc.header["$LUNITS"] = 2
    for layer in SOURCE_LAYERS | {"NOTE", "TABLE", "CAPTION"}:
        doc.layers.new(layer, dxfattribs={"linetype": "CONTINUOUS", "lineweight": 18})
    doc.styles.new("REVIEW_CN", dxfattribs={"font": font.name})
    model = doc.modelspace()

    def text(x, y, value, height=3.5, layer="NOTE", center=False):
        entity = model.add_text(value, dxfattribs={"height": height, "layer": layer,
                                                  "style": "REVIEW_CN", "color": 7})
        entity.set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER if center else TextEntityAlignment.LEFT)

    def block(x, top, value, width, height=3.5, layer="NOTE"):
        lines = wrap(value, width, font, height)
        for i, line in enumerate(lines):
            text(x, top - height - i * height * 1.5, line, height, layer)
        return len(lines) * height * 1.5

    # Measure EVERY column, including names, before positioning the diagram.
    xs = [15, 32, 84, 101, 255]
    cells = [[r["number"], r["name"], str(r["quantity"]), r["remark"]] for r in rows]
    line_counts = [max(len(wrap(v, xs[j + 1] - xs[j] - 6, font, 3.2))
                       for j, v in enumerate(row)) for row in cells]
    heights = [max(10, n * 4.8 + 4) for n in line_counts]
    table_top = 18 + 10 + sum(heights)
    summary = claim(sheet["overview"])
    summary_h = len(wrap(summary, 240, font, 3.5)) * 5.25
    diagram_bottom = table_top + summary_h + 16
    diagram_h = 280 - diagram_bottom
    require(diagram_h >= 100, "说明表过长，挤占主图；请按功能拆页或精简文字，不缩小字号")

    y = table_top
    all_heights = [10] + heights
    for i, row in enumerate([["序号", "名称", "用量", "备注（作用与装配）"]] + cells):
        model.add_line((15, y), (255, y), dxfattribs={"layer": "TABLE"})
        for j, value in enumerate(row):
            block(xs[j] + 3, y - 2, value, xs[j + 1] - xs[j] - 6, 3.2, "TABLE")
        y -= all_heights[i]
    model.add_line((15, y), (255, y), dxfattribs={"layer": "TABLE"})
    for x in xs:
        model.add_line((x, y), (x, table_top), dxfattribs={"layer": "TABLE"})
    block(15, table_top + summary_h + 8, summary, 240)

    # One uniform affine transform for geometry, original leaders and numerals.
    bounds = bbox.extents(entities)
    require(bounds.has_data and bounds.size.x > 0 and bounds.size.y > 0, "源图范围无效")
    scale = min(232 / bounds.size.x, (diagram_h - 8) / bounds.size.y)
    transform = Matrix44.chain(Matrix44.scale(scale), Matrix44.translate(
        19 + (232 - bounds.size.x * scale) / 2 - bounds.extmin.x * scale,
        diagram_bottom + 4 + (diagram_h - 8 - bounds.size.y * scale) / 2 - bounds.extmin.y * scale, 0))
    for entity in entities:
        copy = entity.copy()
        copy.transform(transform)
        copy.dxf.linetype = "CONTINUOUS"
        copy.dxf.color = 7
        if copy.dxftype() == "TEXT":
            require(copy.dxf.height >= 2.5, "缩放后编号低于内部说明图 2.5 mm 字高；需拆图或调整源视图")
            copy.dxf.style = "REVIEW_CN"
        model.add_entity(copy)

    top = 279
    top -= block(270, top, "【" + sheet["title"] + "】", 135, 4.5) + 8
    for paragraph in sheet["paragraphs"]:
        height = len(wrap(claim(paragraph), 135, font, 3.7)) * 5.55
        require(top - height > 45, "右侧说明溢出，请精简或拆页")
        top -= block(270, top, claim(paragraph), 135, 3.7) + 8
    text(283, 26, "图 " + str(page_number), 6, "CAPTION")
    text(15, 7, "内部说明图 · 功能 / 装配 / 零件对照", 2.6)
    text(270, 7, "状态以证据记录为准 · " + VERSION, 2.6)
    final_bounds = bbox.extents(model)
    require(final_bounds.extmin.x >= 0 and final_bounds.extmin.y >= 0
            and final_bounds.extmax.x <= PAGE[0] and final_bounds.extmax.y <= PAGE[1], "实体超出纸张")
    return doc, {"source_sha256": sha(source_path), "source_entities": len(entities),
                 "scale": scale, "transform": list(transform), "page_mm": list(PAGE),
                 "font_sha256": sha(font), "visual_review": "pending", "engineering_review": "pending"}


def preview(dxf, font):
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import pyplot as plt, font_manager
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    from ezdxf.addons.drawing.config import Configuration, ColorPolicy, BackgroundPolicy

    font_manager.fontManager.addfont(str(font))
    final = ezdxf.readfile(dxf)  # render the saved CAD, never an independent drawing
    fig = plt.figure(figsize=(PAGE[0] / 25.4, PAGE[1] / 25.4))
    ax = fig.add_axes([0, 0, 1, 1])
    try:
        Frontend(RenderContext(final), MatplotlibBackend(ax), config=Configuration(
            color_policy=ColorPolicy.BLACK, background_policy=BackgroundPolicy.WHITE
        )).draw_layout(final.modelspace(), finalize=True)
        # Backend.finalize() changes figure size; restore physical paper AFTER it.
        fig.set_size_inches(PAGE[0] / 25.4, PAGE[1] / 25.4)
        ax.set_position([0, 0, 1, 1])
        ax.set_xlim(0, PAGE[0]); ax.set_ylim(0, PAGE[1])
        ax.set_aspect("equal"); ax.axis("off")
        fig.savefig(dxf.with_suffix(".png"), dpi=160, facecolor="white")
        fig.savefig(dxf.with_suffix(".pdf"), facecolor="white", metadata={"CreationDate": None})
    finally:
        plt.close(fig)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True, help="new or empty output directory")
    parser.add_argument("--font", type=Path, required=True, help="installed CJK TrueType/OpenType font")
    parser.add_argument("--preview", action="store_true", help="also render PNG and physical A3 PDF from saved DXF")
    args = parser.parse_args(argv)
    try:
        require(args.font.is_file(), "中文字体不存在")
        require(not args.output.exists() or not any(args.output.iterdir()), "输出目录须为空，避免新旧交付混用")
        data = load(args.manifest)
        parts = validate(data, args.manifest.parent)
        built = [compose(s, parts, args.manifest.parent / s["source_dxf"], args.font, i)
                 for i, s in enumerate(data["sheets"], 1)]
        args.output.mkdir(parents=True, exist_ok=True)
        records = []
        for sheet, (doc, report) in zip(data["sheets"], built):
            path = args.output / (sheet["id"] + "_engineering.dxf")
            doc.saveas(path)
            if args.preview:
                preview(path, args.font)
            report.update(id=sheet["id"], files={p.name: sha(p) for p in sorted(args.output.glob(sheet["id"] + "_engineering.*"))})
            records.append(report)
        result = {"schema": VERSION, "manifest_sha256": sha(args.manifest),
                  "reference_numerals_sha256": sha(args.manifest.parent / data["reference_numerals"]),
                  "status": "generated_requires_review", "sheets": records,
                  "unresolved_parts": [p["id"] for p in parts.values() if p["number"] == "—"],
                  "declared_instances": len({i for p in parts.values() for i in p["instance_ids"]}),
                  "coverage_note": "实例清单由输入声明；不是对 STEP 几何覆盖的自动认证"}
        (args.output / "sheet-report.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"status": result["status"], "sheets": len(records)}, ensure_ascii=False))
        return 0
    except (ValueError, KeyError, TypeError, OSError, ezdxf.DXFError) as exc:
        print("说明图生成失败：" + str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
