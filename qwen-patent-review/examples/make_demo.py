#!/usr/bin/env python3
"""Create synthetic 2D test inputs outside the repository (not customer CAD).

Run: python3 examples/make_demo.py -o /tmp/review-demo
Then compose /tmp/review-demo/review-sheets.json with the normal sheet CLI.
This verifies the compositor; the sibling CAD acceptance suite verifies STEP/HLR.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import shutil
import ezdxf


def make_demo(output):
    output.mkdir(parents=True, exist_ok=False)
    doc = ezdxf.new("R2018")
    doc.header["$INSUNITS"] = 4
    for layer in ("GEOM", "HIDDEN", "LEADER", "NUM"):
        doc.layers.new(layer, dxfattribs={"linetype": "CONTINUOUS"})
    model = doc.modelspace()
    model.add_lwpolyline([(30, 30), (110, 30), (110, 70), (30, 70)], close=True, dxfattribs={"layer": "GEOM"})
    model.add_circle((70, 50), 12, dxfattribs={"layer": "GEOM"})
    model.add_circle((70, 50), 7, dxfattribs={"layer": "GEOM"})
    model.add_lwpolyline([(62, 98), (78, 98), (78, 115), (75, 115), (75, 145),
                         (65, 145), (65, 115), (62, 115)], close=True, dxfattribs={"layer": "GEOM"})
    model.add_lwpolyline([(30, 70), (15, 85), (5, 85)], dxfattribs={"layer": "LEADER"})
    model.add_text("1", dxfattribs={"height": 5, "insert": (5, 88), "layer": "NUM"})
    model.add_lwpolyline([(75, 133), (108, 149), (122, 149)], dxfattribs={"layer": "LEADER"})
    model.add_text("2", dxfattribs={"height": 5, "insert": (117, 152), "layer": "NUM"})
    doc.saveas(output / "synthetic.dxf")
    numerals = {"schema": "patent-numerals/1", "numerals": [
        {"selector": "SYN-A", "numeral": 1, "term": "底座", "figures": ["demo"]},
        {"selector": "SYN-B", "numeral": 2, "term": "阶梯轴", "figures": ["demo"]}]}
    (output / "reference-numerals.json").write_text(json.dumps(numerals, ensure_ascii=False, indent=2), encoding="utf-8")
    shutil.copyfile(Path(__file__).with_name("review-sheets.json"), output / "review-sheets.json")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("-o", "--output", type=Path, required=True)
    make_demo(parser.parse_args().output)
