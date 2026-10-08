"""工作台：MCP 协议、大模型输出清洗、端到端（合成装配体）。

端到端部分需要 cadquery-ocp；协议与清洗部分不需要。配置目录一律指到 tmp，
绝不读写本机真实的 ~/.config/patent-workbench。
"""

from __future__ import annotations

import json
import sys
import time
import zipfile
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parents[1]
_SCRIPTS = _ROOT / "scripts"
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import workbench_llm as L  # noqa: E402
import workbench_mcp as M  # noqa: E402

try:
    import OCP  # noqa: F401
    _HAS_OCP = True
except Exception:  # pragma: no cover
    _HAS_OCP = False


@pytest.fixture(autouse=True)
def _isolated_config(tmp_path, monkeypatch):
    monkeypatch.setattr(L, "CONFIG_DIR", tmp_path / "cfg")
    monkeypatch.setattr(L, "CONFIG_FILE", tmp_path / "cfg" / "config.json")
    monkeypatch.delenv("DEEPSEEK_API_KEY", raising=False)
    monkeypatch.setenv("WORKBENCH_TOKEN", "t0k")


# ----------------------------------------------------------------- MCP 协议

def _rpc(method, params=None, mid=1):
    return {"jsonrpc": "2.0", "id": mid, "method": method, "params": params or {}}


def test_mcp_initialize_negotiates_version():
    r = M.handle_message(_rpc("initialize", {"protocolVersion": "2025-03-26"}), {})
    assert r["result"]["protocolVersion"] == "2025-03-26"
    r = M.handle_message(_rpc("initialize", {"protocolVersion": "1999-01-01"}), {})
    assert r["result"]["protocolVersion"] == M.PROTOCOL_VERSIONS[0]
    assert "tools" in r["result"]["capabilities"]


def test_mcp_notifications_get_no_reply_and_batches_work():
    assert M.handle_message({"jsonrpc": "2.0", "method": "notifications/initialized"}, {}) is None
    assert M.handle_payload([{"jsonrpc": "2.0", "method": "notifications/initialized"}], {}) is None
    out = M.handle_payload([_rpc("ping", mid=1), _rpc("tools/list", mid=2)], {})
    assert [o["id"] for o in out] == [1, 2]


def test_mcp_tool_errors_are_results_not_crashes():
    def boom(args):
        raise M.ToolError("坏输入")

    r = M.handle_message(_rpc("tools/call", {"name": "x", "arguments": {}}), {"x": boom})
    assert r["result"]["isError"] is True and "坏输入" in r["result"]["content"][0]["text"]
    r = M.handle_message(_rpc("tools/call", {"name": "nope"}), {})
    assert r["error"]["code"] == -32602


def test_every_tool_has_a_closed_input_schema():
    for tool in M.TOOLS:
        assert tool["inputSchema"]["type"] == "object"
        assert tool["inputSchema"]["additionalProperties"] is False


# ----------------------------------------------------------------- 大模型输出清洗

def test_parse_json_tolerates_fences_and_chatter():
    assert L._parse_json('```json\n{"a": 1}\n```') == {"a": 1}
    assert L._parse_json('好的：{"a": [1, 2]} 以上') == {"a": [1, 2]}
    with pytest.raises(L.LLMError):
        L._parse_json("没有 JSON")


def test_draft_terms_drops_codes_and_dedupes(monkeypatch):
    monkeypatch.setattr(L, "chat_json", lambda *a, **k: {"terms": [
        {"selector": "P1", "term": "电路板", "label": "once", "confidence": "high"},
        {"selector": "P2", "term": "电路板", "label": "once", "confidence": "high"},
        {"selector": "P3", "term": "PCB-K板", "label": "once"},          # 含字母：丢弃
        {"selector": "P4", "term": "连接器", "label": "none"},
        {"selector": "ZZ", "term": "不存在的件", "label": "once"},        # 不在零件表：丢弃
    ]})
    asm = {"parts": [{"name": n, "bbox_size": [s, 1, 1]} for n, s in
                     (("P1", 10), ("P2", 60), ("P3", 5), ("P4", 3))]}
    rows = {r["selector"]: r for r in L.draft_terms(asm, {"terms": []})}
    assert set(rows) == {"P1", "P2", "P4"}
    assert rows["P2"]["term"] == "第一电路板" and rows["P1"]["term"] == "第二电路板"
    assert rows["P4"]["label"] == "none"


def test_draft_terms_never_touches_human_terms(monkeypatch):
    seen = {}

    def fake(system, user, **k):
        seen["user"] = json.loads(user)
        return {"terms": []}

    monkeypatch.setattr(L, "chat_json", fake)
    asm = {"parts": [{"name": "A"}, {"name": "B"}]}
    L.draft_terms(asm, {"terms": [{"selector": "A", "term": "底壳"}]})
    asked = [p["name"] for p in seen["user"]["产品零件（STEP 零件名、包围盒、装配路径）"]]
    assert asked == ["B"]


def test_flow_cleanup_strips_step_numbers():
    spec = L._clean_flow({"nodes": [{"id": "a", "kind": "process", "text": "S101：采集图像"},
                                    {"id": "b", "kind": "process", "text": "步骤2 计算偏差"}],
                          "edges": [{"from": "a", "to": "b", "label": ""}]}, "图名")
    assert [n["text"] for n in spec["nodes"]] == ["采集图像", "计算偏差"]
    assert "label" not in spec["edges"][0]


def test_secrets_file_is_private(tmp_path):
    L.save_config({"deepseek": {"api_key": "sk-test-0000-1234"}})
    assert (L.CONFIG_FILE.stat().st_mode & 0o777) == 0o600
    pub = L.public_settings()
    assert pub["api_key_set"] and pub["api_key_hint"] == "…1234"
    assert "sk-test" not in json.dumps(pub)


# ----------------------------------------------------------------- 端到端


@pytest.mark.skipif(not _HAS_OCP, reason="需要 cadquery-ocp 才能从 STEP 出图")
def test_workbench_end_to_end(tmp_path):
    import workbench as W
    from starlette.testclient import TestClient

    reg = W.Registry(tmp_path / "data", "t0k")
    W.ensure_synthetic(reg)
    app = W.Dispatcher(W.build_workbench(reg, "t0k", 8790), reg)
    c = TestClient(app, base_url="http://127.0.0.1:8790")
    H = {"Authorization": "Bearer t0k"}
    t0 = time.time()
    while True:
        st = c.get("/api/wb/projects/example-synthetic", headers=H).json()
        if st["status"] == "error":
            pytest.fail(st["message"])
        if st["status"] == "ready" and not st["message"]:
            break
        assert time.time() - t0 < 300
        time.sleep(1)

    # 鉴权：主应用与挂载的 Studio 子应用都要 token
    assert c.get("/api/wb/state").status_code == 401
    assert c.get("/p/example-synthetic/api/state").status_code == 401
    assert c.get("/api/wb/state", headers={"Host": "evil.example"}).status_code == 403

    state = c.get("/p/example-synthetic/api/state", headers=H).json()
    figs = state["render"]["figures"]
    assert [f["number"] for f in figs] == [1, 2]
    assert all(set(f["variants"]) == {"raw", "annotated", "filing"} for f in figs)
    assert state["render"]["figure_descriptions"][0].startswith("图1为")
    assert state["flowcharts"][0]["number"] == 3          # 流程图接在结构附图之后

    slides = c.get("/api/wb/carousel", headers=H).json()["slides"]
    assert {s["kind"] for s in slides} == {"structure", "flow"}
    assert c.get(slides[0]["src"], headers=H).headers["content-type"] == "image/png"

    def call(name, args):
        r = c.post("/mcp", headers=H, json=_rpc("tools/call", {"name": name, "arguments": args}))
        return r.json()["result"]

    assert c.post("/mcp", headers=H, json=_rpc("initialize")).status_code == 200
    res = call("list_projects", {})
    assert res["structuredContent"]["projects"][0]["id"] == "example-synthetic"
    res = call("render_flowchart", {"spec": M.FLOW_GUIDE["example"]})
    assert res["structuredContent"]["steps"][0]["step"] == "S101"
    assert any(x["type"] == "image" for x in res["content"])
    res = call("update_terms", {"project_id": "example-synthetic",
                                "terms": [{"selector": "SYN-A01", "term": "新名字"}]})
    assert res["structuredContent"]["kept_human"] == ["SYN-A01"]

    out = c.post("/p/example-synthetic/api/export", headers=H, json={"dwg": False}).json()
    z = Path(out["dir"]).with_suffix(".zip")
    names = zipfile.ZipFile(z).namelist()
    assert "说明书附图_交底版.docx" in names and "附图说明.txt" in names
    assert any(n.startswith("交底版_带件号表/") for n in names)
    assert any(n.startswith("流程图/") for n in names)
