from __future__ import annotations
import copy
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys

import ezdxf
import numpy as np
from ezdxf.math import Matrix44
import pytest

ROOT = Path(__file__).resolve().parents[1]


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


review = module("review_sheets", ROOT / "scripts/compose_review_sheets.py")
demo = module("review_demo", ROOT / "examples/make_demo.py")


@pytest.fixture
def project(tmp_path):
    path = tmp_path / "input"
    demo.make_demo(path)
    return path, review.load(path / "review-sheets.json")


@pytest.fixture
def font():
    path = os.environ.get("REVIEW_CJK_FONT")
    if not path:
        pytest.skip("Set REVIEW_CJK_FONT to an installed Chinese OTF/TTF for rendering tests")
    assert Path(path).is_file()
    return Path(path)


def test_quantity_and_names_are_derived(project):
    base, data = project
    data["parts"][0]["instance_ids"].append("SYN/root/base/1")
    parts = review.validate(data, base)
    assert parts["base"]["quantity"] == 2
    assert parts["base"]["name"] == "底座"
    assert parts["shaft"]["number"] == "2"


@pytest.mark.parametrize("mode", ["empty_function", "missing_evidence", "repeated_instance", "repeated_selector", "unknown_row", "unsafe_id", "layout_override", "typed_quantity"])
def test_content_failures(project, mode):
    base, data = project
    if mode == "empty_function":
        data["parts"][0]["function"]["text"] = " "
    elif mode == "missing_evidence":
        data["parts"][0]["function"]["evidence"] = []
    elif mode == "repeated_instance":
        data["parts"][1]["instance_ids"] = data["parts"][0]["instance_ids"]
    elif mode == "repeated_selector":
        data["parts"][1]["selector"] = "SYN-A"
    elif mode == "unknown_row":
        data["sheets"][0]["rows"].append("unknown")
    elif mode == "unsafe_id":
        data["sheets"][0]["id"] = "../overwrite"
    elif mode == "layout_override":
        data["sheets"][0]["font_height"] = 1
    else:
        data["parts"][0]["quantity"] = 99
    with pytest.raises(ValueError):
        review.validate(data, base)


def test_same_name_unresolved_instances_not_merged(project):
    base, data = project
    for index in (3, 16, 22):
        data["parts"].append({"id": str(index), "name": "同名候选件", "instance_ids": [f"SYN/reused/{index}"],
                              "unresolved_reason": "同名不同几何，源 CLI 无法分别编号",
                              "function": {"status": "pending", "text": "身份需要核对"},
                              "assembly": {"status": "pending", "text": "连接方式需要核对"}})
    parts = review.validate(data, base)
    assert [parts[str(i)]["quantity"] for i in (3, 16, 22)] == [1, 1, 1]
    assert all(parts[str(i)]["number"] == "—" for i in (3, 16, 22))


def test_transform_preserves_geometry_and_leaders(project, font):
    base, data = project
    parts = review.validate(data, base)
    doc, report = review.compose(data["sheets"][0], parts, base / "synthetic.dxf", font, 1)
    source = ezdxf.readfile(base / "synthetic.dxf")
    original = [e for e in source.modelspace() if e.dxf.layer in {"GEOM", "HIDDEN", "LEADER"}]
    result = [e for e in doc.modelspace() if e.dxf.layer in {"GEOM", "HIDDEN", "LEADER"}]
    assert len(original) == len(result)
    inverse = Matrix44(report["transform"])
    inverse.inverse()
    for before, after in zip(original, result):
        restored = after.copy()
        restored.transform(inverse)
        assert before.dxftype() == restored.dxftype()
        if before.dxftype() == "LWPOLYLINE":
            np.testing.assert_allclose(list(restored.get_points()), list(before.get_points()), atol=1e-9, rtol=1e-10)
        elif before.dxftype() == "CIRCLE":
            assert tuple(restored.dxf.center) == pytest.approx(tuple(before.dxf.center))
            assert restored.dxf.radius == pytest.approx(before.dxf.radius)
    assert {e.dxf.text for e in doc.modelspace().query('TEXT[layer=="NUM"]')} == {"1", "2"}
    assert not doc.audit().errors


def test_missing_number_and_overflow_are_blocked(project, font):
    base, data = project
    parts = review.validate(data, base)
    bad = copy.deepcopy(data["sheets"][0])
    bad["rows"] = ["base"]
    with pytest.raises(ValueError, match="表格未解释"):
        review.compose(bad, parts, base / "synthetic.dxf", font, 1)
    parts["base"]["remark"] = "很长的功能装配说明" * 500
    with pytest.raises(ValueError, match="挤占主图"):
        review.compose(data["sheets"][0], parts, base / "synthetic.dxf", font, 1)


def test_chinese_wrap_measures_actual_font(font):
    value = "带有较长中文名称的安装支承结构（功能与连接关系）" * 4
    lines = review.wrap(value, 48, font, 3.2)
    face = review.fonts.make_font(str(font), 3.2)
    assert "".join(lines) == value
    assert len(lines) > 1
    assert all(face.text_width(line) <= 48 for line in lines)


def test_cli_exports_true_a3_from_saved_cad(project, font, tmp_path):
    base, _ = project
    output = tmp_path / "output"
    command = [sys.executable, str(ROOT / "scripts/compose_review_sheets.py"),
               str(base / "review-sheets.json"), "--font", str(font), "-o", str(output), "--preview"]
    result = subprocess.run(command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    report = review.load(output / "sheet-report.json")
    assert report["status"] == "generated_requires_review"
    assert report["sheets"][0]["visual_review"] == "pending"
    # PDF MediaBox is readable without another Python package.
    import re
    pdf = (output / "demo_engineering.pdf").read_bytes()
    match = re.search(rb"/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]", pdf)
    assert match
    assert float(match[1]) == pytest.approx(420 / 25.4 * 72, abs=.01)
    assert float(match[2]) == pytest.approx(297 / 25.4 * 72, abs=.01)
    texts = [e.dxf.text for e in ezdxf.readfile(output / "demo_engineering.dxf").modelspace().query("TEXT")]
    assert "底座" in texts and "阶梯轴" in texts
    assert any("待核实" in t for t in texts)
    assert (output / "demo_engineering.png").stat().st_size > 1000
    assert subprocess.run(command, capture_output=True).returncode == 1  # no mixed stale output
