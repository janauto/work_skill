#!/usr/bin/env python3
"""Plan Studio — a local browser UI for hand-editing figure-plan.json.

    python3 scripts/plan_studio.py ASM.stp

The UI is the plan's third author (LLM draft, human correction, scripts render). Its ONLY
output is figure-plan.json; every render goes through the same CLIs and QA gates the model
route uses, so a hand-made plan must reproduce bit-identically like any other. The page
exposes intent-level controls only — which part belongs to which figure, what it is called,
whether it carries a numeral. No coordinate, text-height or spacing control exists here,
for the same reason those were taken away from the model in v2.

Local-only by design: binds 127.0.0.1, verifies the Host header (DNS-rebinding guard), and
requires a per-session random token on every API call (hostile-webpage guard). The STEP,
the GLB and every artefact stay inside a work directory next to the STEP file.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import re
import secrets
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

try:
    import uvicorn
    from fastapi import FastAPI, HTTPException, Request, Response
    from fastapi.responses import FileResponse, JSONResponse
    from fastapi.staticfiles import StaticFiles
except ImportError:
    print("Plan Studio 需要 fastapi 与 uvicorn（仅本工具需要，渲染链路不依赖）：\n"
          "  python3 -m pip install fastapi uvicorn", file=sys.stderr)
    raise SystemExit(1)

SCRIPTS = Path(__file__).resolve().parent
REPO = SCRIPTS.parent
WEBUI = REPO / "webui"
PY = sys.executable or "python3"

#: SVG stroke widths in millimetres, preview-only (the DXF is the deliverable, this is a picture
#: of it). GEOM heavier than LEADER matches how the sheet is meant to read on paper.
SVG_STROKE = {"GEOM": 0.35, "HIDDEN": 0.2, "LEADER": 0.18, "TABLE": 0.18}
SVG_MARGIN_MM = 6.0
HISTORY_KEEP = 200


# ----------------------------------------------------------------------------- work dir


class Workspace:
    """Everything Plan Studio knows lives in ``<step stem>.plan-studio/`` beside the STEP."""

    def __init__(self, step: Path, workdir: Path | None) -> None:
        self.step = step.resolve()
        self.root = (workdir or self.step.parent / (self.step.stem + ".plan-studio")).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.assembly = self.root / "assembly.json"
        self.plan = self.root / "plan.json"
        self.glb = self.root / "model.glb"
        self.out = self.root / "out"
        self.cache = self.root / "cache"
        self.history = self.root / ".plan-history"
        self.state = self.root / "state.json"

    def stale(self) -> bool:
        if not self.state.is_file():
            return True
        try:
            recorded = json.loads(self.state.read_text(encoding="utf-8"))
        except ValueError:
            return True
        return (recorded.get("step_mtime") != self.step.stat().st_mtime_ns
                or recorded.get("step_size") != self.step.stat().st_size)

    def remember(self) -> None:
        self.state.write_text(json.dumps({
            "step": str(self.step),
            "step_mtime": self.step.stat().st_mtime_ns,
            "step_size": self.step.stat().st_size,
        }, ensure_ascii=False, indent=2), encoding="utf-8")

    def snapshot_plan(self) -> None:
        if not self.plan.is_file():
            return
        self.history.mkdir(exist_ok=True)
        stamp = time.strftime("%Y%m%d-%H%M%S") + ("-%03d" % (time.time_ns() // 1_000_000 % 1000))
        shutil.copy2(self.plan, self.history / ("plan-%s.json" % stamp))
        old = sorted(self.history.glob("plan-*.json"))
        for path in old[:-HISTORY_KEEP]:
            path.unlink()


def run_cli(args: list, timeout: int = 1800) -> subprocess.CompletedProcess:
    """Every render-chain step goes through the same CLIs the model route uses — never an
    import of package internals. That keeps the human path and the model path behind one
    set of gates with one set of error messages."""
    return subprocess.run([PY] + args, capture_output=True, text=True,
                          timeout=timeout, cwd=str(REPO))


def prepare(ws: Workspace, quiet: bool = False) -> None:
    def say(msg: str) -> None:
        if not quiet:
            print(msg, flush=True)

    if ws.stale() or not ws.assembly.is_file():
        say("· analyze：解析装配体（大模型件首次约半分钟）…")
        proc = run_cli([str(SCRIPTS / "analyze_assembly.py"), str(ws.step),
                       "-o", str(ws.assembly)])
        if proc.returncode != 0:
            print(proc.stdout[-2000:] + proc.stderr[-2000:], file=sys.stderr)
            raise SystemExit("analyze 失败，无法启动 Studio")
    if ws.stale() or not ws.glb.is_file():
        say("· glb：导出 3D 预览模型…")
        proc = run_cli([str(SCRIPTS / "export_step_glb.py"), str(ws.step), "-o", str(ws.glb)])
        if proc.returncode != 0:
            print(proc.stdout[-2000:] + proc.stderr[-2000:], file=sys.stderr)
            raise SystemExit("GLB 导出失败，无法启动 Studio")
    if not ws.plan.is_file():
        say("· plan：生成空白计划（零件清单来自 assembly.json）…")
        ws.plan.write_text(json.dumps(blank_plan(ws), ensure_ascii=False, indent=2),
                           encoding="utf-8")
    ws.remember()


def blank_plan(ws: Workspace) -> dict:
    assembly = json.loads(ws.assembly.read_text(encoding="utf-8"))
    return {
        "schema": "patent-figure-plan/1",
        "source": {"step": str(ws.step), "include": ["*"], "exclude": []},
        # One empty term row per part: the validator's W_UNLABELLED_PART / empty-term errors
        # then double as the user's to-do list instead of a blank page.
        "terms": [{"selector": part["name"], "term": "", "label": "once"}
                  for part in assembly.get("parts", [])],
        "figures": [{"id": "fig1", "caption": "整体结构示意图",
                     "kind": "assembly", "members": ["*"]}],
        "layout": {"view": "iso", "explode_axis": "auto", "axis_angle": "auto",
                   "density": "normal", "max_labels_per_figure": 20,
                   "engineering_table": False},
    }


# ----------------------------------------------------------------------------- BOM

#: Header keywords for locating the code / name / qty columns in a BOM sheet. Chinese
#: manufacturing BOMs vary wildly; these cover the common house styles. Matching is
#: substring, first hit wins, scanned over the first BOM_HEADER_SCAN rows.
BOM_CODE_HEADERS = ("物料编码", "物料代码", "物料编号", "图号", "件号", "编码", "代号",
                    "料号", "零件号", "part no", "p/n", "partnumber", "code")
BOM_NAME_HEADERS = ("物料名称", "零件名称", "名称", "品名", "描述", "规格名称",
                    "description", "name")
BOM_QTY_HEADERS = ("数量", "用量", "qty", "quantity")
BOM_HEADER_SCAN = 30


def _bom_rows_from_table(rows: list) -> dict:
    """Locate the header row and pull (code -> {name, qty}) out of a 2-D table."""
    header_idx = code_col = name_col = qty_col = None
    for i, row in enumerate(rows[:BOM_HEADER_SCAN]):
        cells = [str(c or "").strip().lower() for c in row]
        c_col = n_col = q_col = None
        for j, cell in enumerate(cells):
            if c_col is None and any(k in cell for k in BOM_CODE_HEADERS):
                c_col = j
            elif n_col is None and any(k in cell for k in BOM_NAME_HEADERS):
                n_col = j
            elif q_col is None and any(k in cell for k in BOM_QTY_HEADERS):
                q_col = j
        if c_col is not None and n_col is not None:
            header_idx, code_col, name_col, qty_col = i, c_col, n_col, q_col
            break
    if header_idx is None:
        raise ValueError("找不到表头：需要同时含「件号/物料编码」列与「名称/品名」列")
    out = {}
    for row in rows[header_idx + 1:]:
        code = str(row[code_col] or "").strip() if code_col < len(row) else ""
        name = str(row[name_col] or "").strip() if name_col < len(row) else ""
        if not code or not name:
            continue
        qty = None
        if qty_col is not None and qty_col < len(row):
            try:
                qty = int(float(row[qty_col]))
            except (TypeError, ValueError):
                qty = None
        out.setdefault(code, {"name": name, "qty": qty})
    return out


def parse_bom(filename: str, data: bytes) -> dict:
    """xlsx via openpyxl（全部工作表都扫）, csv/tsv via stdlib. Returns code -> {name, qty}."""
    suffix = Path(filename).suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        merged, errors = {}, []
        for sheet in wb.worksheets:
            rows = [[c for c in r] for r in sheet.iter_rows(values_only=True)]
            try:
                found = _bom_rows_from_table(rows)
            except ValueError as exc:
                errors.append("%s: %s" % (sheet.title, exc))
                continue
            for code, info in found.items():
                merged.setdefault(code, info)
        if not merged:
            raise ValueError("；".join(errors) or "工作簿为空")
        return merged
    if suffix in (".csv", ".tsv", ".txt"):
        import csv
        import io
        text = data.decode("utf-8-sig", errors="replace")
        dialect = "excel-tab" if suffix == ".tsv" else "excel"
        rows = list(csv.reader(io.StringIO(text), dialect))
        return _bom_rows_from_table(rows)
    raise ValueError("不支持的格式 %s——请用 .xlsx 或 .csv" % suffix)


_INSTANCE_SUFFIX = re.compile(r"(?:[_-]\d+)+$")


def match_bom_to_parts(bom: dict, part_names: list) -> dict:
    """Match BOM codes to STEP part names — conservatively.

    The first draft also ran a longest-prefix pass. On a real assembly it filled fourteen
    distinct gears, shells and bearings sharing one export-tool stem (``<HASH>_<seq>`` style
    names) with a single bearing's name, and every ``BREP_*`` blob with one cover's name:
    recall bought with wrong names, which on a patent figure is worse than no name. So only
    two passes survive:

    1. raw name == raw code;
    2. instance-suffix-stripped equality, accepted only when the stripped key is at least
       MIN_STEM chars AND unique on both sides (one part, one BOM row). Stripping ``_1_1``
       style suffixes is what lets ``<CODE>_1_1`` meet its drawing code ``<CODE>``; the
       uniqueness demand is what keeps ``BREP_<n>`` -> ``BREP`` from meeting every other
       ``BREP_*``.

    Whatever stays unmatched is reported honestly and left for the human or the model.
    """
    MIN_STEM = 6
    matched = {}
    codes = set(bom)
    for name in part_names:
        if name in codes:
            matched[name] = {"code": name, "name": bom[name]["name"], "via": "exact"}

    def stem_index(values):
        index = {}
        for v in values:
            stem = _INSTANCE_SUFFIX.sub("", v)
            if len(stem) >= MIN_STEM:
                index.setdefault(stem, []).append(v)
        return index

    name_stems = stem_index(n for n in part_names if n not in matched)
    code_stems = stem_index(codes)
    for stem, names in name_stems.items():
        cands = code_stems.get(stem, [])
        if len(names) == 1 and len(cands) == 1:
            matched[names[0]] = {"code": cands[0], "name": bom[cands[0]]["name"],
                                 "via": "stripped"}
    return matched


# ----------------------------------------------------------------------------- DXF -> SVG


def dxf_to_svg(dxf: Path) -> str:
    """Minimal DXF-to-SVG for THIS renderer's known output vocabulary.

    Hand-rolled instead of ezdxf's SVGBackend for one reason: the preview must be
    interactive — every NUM numeral needs ``data-numeral`` so a click can highlight the
    part in 3D — and a generic backend gives no per-entity hooks. The sheet only ever
    contains LINE / LWPOLYLINE / CIRCLE / TEXT on known layers, so sixty lines cover it.

    Y handling per the sheet's +Y-up convention: geometry sits in a ``scale(1,-1)`` group;
    text is emitted OUTSIDE that group at ``y' = -y``, because a flipped group mirrors
    glyphs. All coordinates are millimetres straight from the DXF.
    """
    import ezdxf

    doc = ezdxf.readfile(str(dxf))
    msp = doc.modelspace()
    xs, ys = [], []

    def track(x: float, y: float) -> None:
        xs.append(float(x)); ys.append(float(y))

    shapes, texts = [], []
    for e in msp:
        kind = e.dxftype()
        layer = e.dxf.layer
        if kind == "LINE":
            (x1, y1, _), (x2, y2, _) = e.dxf.start, e.dxf.end
            track(x1, y1); track(x2, y2)
            shapes.append('<line class="ly-%s" x1="%.3f" y1="%.3f" x2="%.3f" y2="%.3f"/>'
                          % (layer, x1, y1, x2, y2))
        elif kind == "LWPOLYLINE":
            pts = [(float(p[0]), float(p[1])) for p in e.get_points()]
            for x, y in pts:
                track(x, y)
            d = " ".join("%.3f,%.3f" % p for p in pts)
            tag = "polygon" if e.closed else "polyline"
            shapes.append('<%s class="ly-%s" points="%s"/>' % (tag, layer, d))
        elif kind == "CIRCLE":
            (cx, cy, _) = e.dxf.center
            r = float(e.dxf.radius)
            track(cx - r, cy - r); track(cx + r, cy + r)
            shapes.append('<circle class="ly-%s dot" cx="%.3f" cy="%.3f" r="%.3f"/>'
                          % (layer, cx, cy, r))
        elif kind == "TEXT":
            align, p1, _ = e.get_placement()   # ezdxf: (alignment, p1, p2) — 对齐枚举在首位
            x, y = float(p1[0]), float(p1[1])
            h = float(e.dxf.height)
            track(x, y - h); track(x, y + h)
            name = align.name if hasattr(align, "name") else str(align)
            anchor = "end" if name.endswith("RIGHT") else (
                "start" if name.endswith("LEFT") else "middle")
            value = (e.dxf.text.replace("&", "&amp;").replace("<", "&lt;"))
            extra = ""
            if layer == "NUM" and value.strip().isdigit():
                extra = ' data-numeral="%s"' % value.strip()
            texts.append('<text class="ly-%s" x="%.3f" y="%.3f" font-size="%.3f" '
                         'text-anchor="%s"%s>%s</text>'
                         % (layer, x, -y, h, anchor, extra, value))
    if not xs:
        return '<svg xmlns="http://www.w3.org/2000/svg"/>'
    m = SVG_MARGIN_MM
    x0, x1 = min(xs) - m, max(xs) + m
    y0, y1 = min(ys) - m, max(ys) + m
    styles = "".join(".ly-%s{stroke-width:%.2f}" % (k, v) for k, v in SVG_STROKE.items())
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="%.3f %.3f %.3f %.3f" '
        'font-family="system-ui, \'PingFang SC\', sans-serif">'
        '<style>line,polyline,polygon,circle{stroke:#1a1a1a;fill:none;stroke-linecap:round}'
        'circle.dot{fill:#1a1a1a;stroke:none}'
        'text{fill:#1a1a1a;dominant-baseline:central}'
        'text[data-numeral]{cursor:pointer}text[data-numeral]:hover{fill:#0a6cbd}%s</style>'
        '<g transform="scale(1,-1)">%s</g>%s</svg>'
        % (x0, -y1, x1 - x0, y1 - y0, styles, "".join(shapes), "".join(texts))
    )


# ----------------------------------------------------------------------------- app


def build_app(ws: Workspace, token: str) -> FastAPI:
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
    render_lock = threading.Lock()
    render_state: dict = {"running": False, "result": None, "started": None}

    @app.middleware("http")
    async def guard(request: Request, call_next):
        host = (request.headers.get("host") or "").split(":")[0]
        if host not in ("127.0.0.1", "localhost"):
            return Response("仅限本机访问", status_code=403)
        if request.url.path.startswith("/api/"):
            supplied = request.headers.get("x-studio-token") or \
                request.query_params.get("token")
            if supplied != token:
                return Response("token 无效——请从终端打印的完整地址进入", status_code=401)
        response = await call_next(request)
        if not request.url.path.startswith("/api/"):
            # 前端是本机文件、无构建版本号：升级 Studio 后浏览器的 ES 模块缓存会继续
            # 跑旧代码（同页导航尤甚），排障时症状诡异。本地服务器带宽为零成本，
            # 直接禁缓存是最省心的正确解。
            response.headers["Cache-Control"] = "no-store"
        return response

    # ---- state -------------------------------------------------------------

    def load_json(path: Path):
        return json.loads(path.read_text(encoding="utf-8"))

    def validate_now() -> dict:
        proc = run_cli([str(SCRIPTS / "validate_figure_plan.py"), str(ws.plan),
                       "--assembly", str(ws.assembly),
                       "--json", str(ws.root / "issues.json")], timeout=120)
        issues = []
        if (ws.root / "issues.json").is_file():
            try:
                issues = load_json(ws.root / "issues.json").get("issues", [])
            except ValueError:
                issues = []
        return {"ok": proc.returncode == 0, "exit": proc.returncode,
                "issues": issues, "stdout": proc.stdout[-4000:]}

    @app.get("/api/state")
    def state():
        return {
            "step": str(ws.step),
            "workdir": str(ws.root),
            "assembly": load_json(ws.assembly),
            "plan": load_json(ws.plan),
            "validate": validate_now(),
            "render": render_state["result"],
            "rendering": render_state["running"],
        }

    @app.put("/api/plan")
    async def put_plan(request: Request):
        body = await request.json()
        plan = body.get("plan")
        if not isinstance(plan, dict):
            raise HTTPException(400, "body.plan 必须是对象")
        ws.snapshot_plan()
        ws.plan.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"validate": validate_now()}

    @app.post("/api/bom")
    async def import_bom(request: Request):
        filename = request.query_params.get("name", "bom.xlsx")
        data = await request.body()
        if not data:
            raise HTTPException(400, "空文件")
        try:
            bom = parse_bom(filename, data)
        except ValueError as exc:
            raise HTTPException(422, "BOM 解析失败：%s" % exc)
        assembly = load_json(ws.assembly)
        names = [p["name"] for p in assembly.get("parts", [])]
        matched = match_bom_to_parts(bom, names)
        plan = load_json(ws.plan)
        ws.snapshot_plan()
        filled, kept = [], []
        by_selector = {t.get("selector"): t for t in plan.get("terms", [])}
        for part, hit in matched.items():
            row = by_selector.get(part)
            if row is None:
                continue
            if (row.get("term") or "").strip():
                kept.append(part)          # 人已填的绝不覆盖
            else:
                row["term"] = hit["name"]
                filled.append(part)
        ws.plan.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        (ws.root / "bom.json").write_text(
            json.dumps({"file": filename, "rows": len(bom), "matched": matched},
                       ensure_ascii=False, indent=2), encoding="utf-8")
        unmatched_parts = sorted(set(names) - set(matched))
        return {
            "plan": plan,
            "validate": validate_now(),
            "report": {
                "bom_rows": len(bom), "matched_parts": len(matched),
                "filled": len(filled), "kept_human": len(kept),
                "unmatched_parts": unmatched_parts[:40],
                "unmatched_count": len(unmatched_parts),
            },
        }

    # ---- render ------------------------------------------------------------

    def resolve_numeral_parts(numerals: dict, part_names: list) -> None:
        for entry in numerals.get("numerals", []):
            selector = entry.get("selector", "")
            entry["parts"] = sorted(n for n in part_names
                                    if fnmatch.fnmatchcase(n, selector))

    def render_job() -> None:
        result: dict = {"figures": [], "ok": False, "log": "", "numerals": None,
                        "finished": None}
        try:
            if ws.out.is_dir():
                shutil.rmtree(ws.out)
            check = validate_now()
            if not check["ok"]:
                result["log"] = "计划未通过校验，未渲染。"
                result["validate"] = check
                return
            proc = run_cli([str(SCRIPTS / "render_patent_figure.py"), str(ws.plan),
                           "--assembly", str(ws.assembly), "-o", str(ws.out),
                           "--cache", str(ws.cache)])
            result["log"] = (proc.stdout[-6000:] + "\n" + proc.stderr[-2000:]).strip()
            result["ok"] = proc.returncode == 0
            part_names = [p["name"] for p in load_json(ws.assembly).get("parts", [])]
            numerals_file = ws.out / "reference-numerals.json"
            if numerals_file.is_file():
                result["numerals"] = load_json(numerals_file)
                resolve_numeral_parts(result["numerals"], part_names)
            for dxf in sorted(ws.out.glob("*.dxf")):
                fig_id = dxf.stem
                (ws.out / (fig_id + ".svg")).write_text(dxf_to_svg(dxf), encoding="utf-8")
                qa_file = ws.out / (fig_id + ".qa.json")
                qa = load_json(qa_file) if qa_file.is_file() else None
                result["figures"].append({
                    "id": fig_id,
                    "pass": bool(qa and qa.get("pass")),
                    "qa": qa,
                    "svg": "/api/preview/%s.svg" % fig_id,
                })
        except Exception as exc:  # surfaced to the UI, never swallowed
            result["log"] += "\n渲染进程异常：%r" % exc
        finally:
            result["finished"] = time.strftime("%H:%M:%S")
            render_state["result"] = result
            render_state["running"] = False
            render_lock.release()

    @app.post("/api/render")
    def render():
        if not render_lock.acquire(blocking=False):
            raise HTTPException(409, "已有一次渲染在进行——几何缓存不支持并发写入")
        render_state.update(running=True, started=time.strftime("%H:%M:%S"))
        threading.Thread(target=render_job, daemon=True).start()
        return {"started": render_state["started"]}

    @app.get("/api/render/status")
    def render_status():
        return {"running": render_state["running"], "result": render_state["result"],
                "started": render_state["started"]}

    @app.get("/api/preview/{name}")
    def preview(name: str):
        if not re.fullmatch(r"[A-Za-z0-9_\-]+\.(svg|png|dxf)", name):
            raise HTTPException(400, "非法文件名")
        path = ws.out / name
        if not path.is_file():
            raise HTTPException(404)
        media = {"svg": "image/svg+xml", "png": "image/png",
                 "dxf": "application/octet-stream"}[path.suffix[1:]]
        return FileResponse(str(path), media_type=media)

    @app.get("/api/model.glb")
    def model():
        return FileResponse(str(ws.glb), media_type="model/gltf-binary")

    # ---- export ------------------------------------------------------------

    @app.post("/api/export")
    async def export(request: Request):
        body = await request.json()
        want_dwg = bool(body.get("dwg", False))
        result = render_state["result"]
        if not (result and result.get("ok") and result["figures"]
                and all(f["pass"] for f in result["figures"])):
            raise HTTPException(409, "存在未通过 QA 的图，不能导出——按提示改计划后重渲")
        stamp = time.strftime("%Y%m%d-%H%M%S")
        dest = ws.root / ("export-%s" % stamp)
        dest.mkdir(parents=True)
        files = []
        for fig in result["figures"]:
            for suffix in (".dxf", ".png", ".svg"):
                src = ws.out / (fig["id"] + suffix)
                if src.is_file():
                    shutil.copy2(src, dest / src.name)
                    files.append(src.name)
        numerals = ws.out / "reference-numerals.json"
        if numerals.is_file():
            shutil.copy2(numerals, dest / numerals.name)
            files.append(numerals.name)
            text = load_json(numerals).get("description_zh", "")
            (dest / "附图标记说明.txt").write_text(text + "\n", encoding="utf-8")
            files.append("附图标记说明.txt")
        dwg_log = []
        if want_dwg:
            for fig in result["figures"]:
                dxf = dest / (fig["id"] + ".dxf")
                proc = run_cli([str(SCRIPTS / "autocad_core_dxf_to_dwg.py"), str(dxf),
                               str(dest / (fig["id"] + ".dwg"))], timeout=600)
                if proc.returncode != 0:
                    proc = run_cli([str(SCRIPTS / "libredwg_dxf_to_dwg.py"), str(dxf),
                                   "-o", str(dest)], timeout=600)
                    dwg_log.append("%s: AutoCAD 失败，改用 LibreDWG（exit=%d）"
                                   % (fig["id"], proc.returncode))
                else:
                    dwg_log.append("%s: AutoCAD 转换成功" % fig["id"])
                if (dest / (fig["id"] + ".dwg")).is_file():
                    files.append(fig["id"] + ".dwg")
        return {"dir": str(dest), "files": sorted(set(files)), "dwg_log": dwg_log}

    app.mount("/", StaticFiles(directory=str(WEBUI), html=True), name="webui")
    return app


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("step", type=Path, help="STEP 装配体路径")
    ap.add_argument("--workdir", type=Path, default=None,
                    help="工作目录（默认：STEP 旁的 <名字>.plan-studio/）")
    ap.add_argument("--port", type=int, default=8425)
    ap.add_argument("--token", default=None,
                    help="固定 API token（默认每次随机；仅测试时使用）")
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()
    if not args.step.is_file():
        print("用法错误：STEP 文件不存在：%s" % args.step, file=sys.stderr)
        return 2
    if not WEBUI.is_dir():
        print("缺少前端目录 %s" % WEBUI, file=sys.stderr)
        return 1

    ws = Workspace(args.step, args.workdir)
    prepare(ws)
    token = args.token or secrets.token_urlsafe(16)
    url = "http://127.0.0.1:%d/?token=%s" % (args.port, token)
    print("\nPlan Studio 已就绪：\n  %s\n工作目录：%s\n（仅本机可访问；关闭终端即停止）\n"
          % (url, ws.root), flush=True)
    if not args.no_browser:
        threading.Timer(0.8, webbrowser.open, args=(url,)).start()
    uvicorn.run(build_app(ws, token), host="127.0.0.1", port=args.port, log_level="warning")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
