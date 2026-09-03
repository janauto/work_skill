# v2 implementation contract

**Read this file completely before writing any code.** It is the single source of truth for module
boundaries, function signatures, JSON shapes, and determinism rules. Where this file and any other
document disagree, this file wins. If something you need is genuinely unspecified here, pick the
option that is most deterministic and record it in your report's `open_issues` — do not invent a
parallel convention.

Rationale for the whole design is in [iteration-plan-v2.md](iteration-plan-v2.md).

---

## 0. Prime directives

1. **The LLM's only output is `figure-plan.json`.** No layout constants, coordinates, text heights,
   or numerals ever come from a model. Everything geometric is computed by these modules.
2. **Determinism is a correctness property, not a nicety.** Same inputs → byte-identical output
   after normalization. Every algorithm here must be reproducible across machines and runs.
3. **One OCC boundary.** `patent_figure/occ_backend.py` is the *only* module that may
   `import OCP.*`. Every other module operates on plain numpy arrays and dicts, so it is unit
   testable without CAD libraries installed.

## 1. Environment (already verified on the dev machine)

| Package | Version | Note |
| --- | --- | --- |
| `cadquery-ocp` | 7.7.2 | OCCT 7.7.2 bindings, namespace `OCP.*`. **pip-installable.** |
| `ezdxf` | 1.4.2 | |
| `numpy` | 2.0.2 | |
| `jsonschema` | 4.25.1 | plan/assembly validation |
| `matplotlib` | 3.9.4 | preview PNG rendering only |
| `pytest` | 8.4.2 | |

Do **not** import `OCC.Core.*` (that is the conda-only `pythonocc-core`). This project uses `OCP.*`.
Python is 3.9 — **no `match`, no PEP 604 `X | Y` annotations at runtime, no `list[int]` in
signatures unless the module has `from __future__ import annotations`** (add it to every module).

## 2. Frozen files — do not modify

`scripts/cad_hlr_to_dxf.py`, `scripts/clean_patent_dxf.py`, `scripts/validate_clean_dxf.py`,
`scripts/autocad_core_dxf_to_dwg.py`, `scripts/libredwg_dxf_to_dwg.py`.

The new package **copies** the OCC logic it needs into `occ_backend.py` (adapted, not imported).
Legacy refactoring is a separate pass. `cad_hlr_to_dxf.py` stays as the single-view legacy route.

## 3. File map and ownership

```
scripts/patent_figure/__init__.py        version string only
scripts/patent_figure/occ_backend.py     STEP load, view frame, HLR, geometry cache   [OCC]
scripts/patent_figure/analyze.py         assembly.json construction                   [OCC via backend]
scripts/patent_figure/plan.py            plan load, schema + semantic validation      [pure]
scripts/patent_figure/numbering.py       deterministic numeral assignment             [pure]
scripts/patent_figure/layout.py          2D sheet placement of parts                  [pure]
scripts/patent_figure/labels.py          2D leader + numeral placement                [pure]
scripts/patent_figure/sheet.py           DXF emission, preview, hash normalization    [ezdxf]
scripts/patent_figure/qa.py              readability + compliance checks on a DXF     [ezdxf]

scripts/analyze_assembly.py              CLI  step        -> assembly.json
scripts/validate_figure_plan.py          CLI  plan+asm    -> exit 0/1 + errors JSON
scripts/render_patent_figure.py          CLI  plan        -> figures + numerals + previews
scripts/qa_patent_figure.py              CLI  dxf         -> qa.json, exit 0/1
scripts/doctor.py                        CLI  ()          -> capability report
scripts/make_test_assembly.py            CLI  ()          -> synthetic STEP fixture

schemas/assembly.schema.json
schemas/figure-plan.schema.json
tests/test_layout.py  tests/test_labels.py  tests/test_numbering.py
tests/test_plan.py    tests/test_qa.py      tests/test_golden.py
tests/fixtures/                            synthetic only — see §10
requirements-pinned.txt
```

## 4. Data shapes

### 4.1 `assembly.json`

```jsonc
{
  "schema": "patent-assembly/1",
  "source": {"step": "<path>", "sha256": "<hex>", "instances": 164, "distinct": 125},
  "units": "mm",
  "bbox": {"lo": [x,y,z], "hi": [x,y,z], "size": [w,d,h]},
  "parts": [{
    "name": "PRT0001",
    "instances": 4,
    "bbox_size": [w,d,h],           // of one instance, rounded to 3 decimals
    "max_dim": 12.5,
    "centers": [[x,y,z], ...],      // one per instance, sorted lexicographically
    "degenerate": false,            // true when max_dim < 1e-6
    "path_sample": "root/sub/PRT0001"
  }],
  "principal_axis": {"vector": [0,0,1], "nearest": "z", "spread_ratio": 4.21},
  "coaxial_groups": [{
    "id": "g1", "axis": [0,0,1], "origin": [x,y,0],
    "members": ["PRT0001", "PRT0002"], "instances": 12,
    "extent": [z_lo, z_hi]
  }],
  "stack_order": ["PRT0007", "PRT0001", ...],   // by projection of center onto principal axis
  "size_tiers": [{"tier": "large", "members": [...]}, {"tier":"medium",...}, {"tier":"small",...}],
  "split_suggestions": [{
    "id": "coaxial", "strategy": "coaxial",
    "figures": [{"caption_hint": "组件 A 分解", "members": ["PRT0001","PRT0002"], "labels": 14}]
  }],
  "warnings": ["3 part(s) have degenerate bounding boxes: PRT0004, PRT0005, ..."]
}
```

Three `split_suggestions` are always emitted when the assembly exceeds the label cap:
`coaxial` (by coaxial group), `stack` (by breaks in `stack_order` spacing), `size` (by `size_tiers`).
Each suggestion must satisfy `labels <= max_labels_per_figure` per figure or it is not emitted.

### 4.2 `figure-plan.json`

```jsonc
{
  "schema": "patent-figure-plan/1",
  "source": {"step": "<path>", "include": ["glob"], "exclude": ["glob"]},
  "terms": [
    {"selector": "PRT0001", "term": "底座", "label": "once"},
    {"selector": "SCREW*",  "term": "紧固螺钉", "label": "none"}
  ],
  "figures": [
    {"id": "fig1", "caption": "整体结构示意图", "kind": "assembly", "members": ["*"]},
    {"id": "fig2", "caption": "回转组件分解示意图", "kind": "exploded",
     "members": ["PRT0001", "PRT0002"], "layout": {"explode_axis": "z"}}
  ],
  "layout": {                       // defaults; a figure may override any key
    "view": "iso",
    "explode_axis": "auto",
    "axis_angle": 152,
    "density": "normal",
    "max_labels_per_figure": 20,
    "engineering_table": false
  }
}
```

Rules enforced by `plan.py`:

- `terms[].selector` is an fnmatch glob; **every selector must match ≥1 part** in `assembly.json`.
- `terms[].term` must not match any pattern in `forbidden_term_patterns` (default: internal part-code
  shapes, see §7). Terms are Chinese technical nouns.
- `terms[].label` ∈ `once | all | none`. Default `once`.
- `terms` order is the numeral issue order. **The plan never contains numerals.**
- `figures[].kind` ∈ `assembly | exploded`.
- `figures[].members` globs resolve against selected parts; a member matching no part is an error.
- Effective label count per figure ≤ `max_labels_per_figure`, else `E_TOO_MANY_LABELS`.
- `layout.view` ∈ the `VIEWS` enum; `explode_axis` ∈ `x|y|z|auto`; `density` ∈ `compact|normal|loose`;
  `axis_angle` is clamped to `[120, 180]` with a `W_CLAMPED` warning.
- **No other keys are accepted.** `additionalProperties: false` everywhere.

### 4.3 `reference-numerals.json`

```jsonc
{"schema": "patent-numerals/1",
 "numerals": [{"numeral": 1, "term": "底座", "selector": "PRT0001", "figures": ["fig1","fig2"]}],
 "description_zh": "附图标记说明：1—底座；2—回转轴组件；……"}
```

### 4.4 `qa.json`

```jsonc
{"schema": "patent-figure-qa/1", "file": "fig1.dxf", "pass": false,
 "checks": [{"id": "geometry_occupancy", "pass": false, "value": 0.135,
             "threshold": ">=0.55", "detail": "...", "hint": "..."}],
 "summary": {"failed": 5, "passed": 4}}
```

## 5. Module signatures

Every module starts with `from __future__ import annotations`.

### 5.1 `occ_backend.py` — the only OCC module

```python
VIEWS: dict[str, tuple[float, float]]        # name -> (az, el); same 7 entries as the legacy script

class ViewFrame:
    def __init__(self, az: float, el: float, roll: float = 0.0) -> None
    x: np.ndarray; u: np.ndarray; w: np.ndarray      # right, up, view direction (unit, right-handed)
    def to2d(self, pts: np.ndarray) -> np.ndarray    # (...,3) -> (...,2)
    @property
    def key(self) -> str                             # stable cache key, e.g. "35.000_20.000_152.000"

def roll_for_axis(axis, az: float, el: float, target_deg: float) -> float

class PartShape:
    key: str          # f"{name}#{instance_index}" — the stable identity used everywhere downstream
    name: str
    path: str
    instance_index: int
    lo: np.ndarray; hi: np.ndarray                   # 3D bbox
    shape: Any                                       # TopoDS_Shape
    @property
    def center(self) -> np.ndarray
    @property
    def degenerate(self) -> bool                     # max extent < 1e-6

def load_assembly(step: Path) -> list[PartShape]
    """Deterministic order: sorted by (name, path, rounded center tuple). instance_index is
    assigned after sorting, so the same STEP always yields the same keys."""

def part_curves(part: PartShape, view: ViewFrame, deflection: float) -> list[np.ndarray]
    """Per-part HLR. VALID ONLY for parts that are disjoint on the sheet."""

def scene_curves(parts: Sequence[PartShape], view: ViewFrame, deflection: float,
                 want_hidden: bool = False) -> tuple[list[np.ndarray], list[np.ndarray]]
    """Global HLR over all parts together. REQUIRED for kind='assembly' figures, where parts
    occlude each other. Returns (visible, hidden) as lists of (N,2) float arrays."""

class GeometryCache:
    def __init__(self, root: Path, step_sha: str) -> None
    def get(self, part_key: str, view: ViewFrame, deflection: float) -> list[np.ndarray] | None
    def put(self, part_key: str, view: ViewFrame, deflection: float,
            curves: list[np.ndarray]) -> None
```

**HLR compounds to collect** (copy the legacy behaviour): `VCompound`, `Rg1LineVCompound`,
`OutLineVCompound` for visible; `HCompound`, `Rg1LineHCompound`, `OutLineHCompound` for hidden.
Omitting `OutLine*` breaks dome and fillet silhouettes.

**Cache**: `<root>/<step_sha[:12]>/<view.key>_<deflection>/<sha1(part_key)>.npz`, storing curves as
`arr_0..arr_n`. A cache hit must be bit-identical to a recompute. Load 24 s on the reference
assembly makes this mandatory, not optional.

### 5.2 `layout.py` — pure 2D, no OCC, no ezdxf

```python
DENSITY: dict[str, float] = {"compact": 0.035, "normal": 0.055, "loose": 0.085}

@dataclass
class Piece:
    key: str
    name: str
    curves: list[np.ndarray]      # 2D, unplaced (view-projected, model position)
    offset: np.ndarray = field(default_factory=lambda: np.zeros(2))
    @property
    def lo(self) -> np.ndarray    # placed bbox min  (curves min + offset)
    @property
    def hi(self) -> np.ndarray
    def placed(self) -> list[np.ndarray]

@dataclass
class LayoutResult:
    pieces: list[Piece]
    lo: np.ndarray; hi: np.ndarray     # overall geometry bbox after placement
    slot: float                        # min over pieces of max(hi-lo) — drives text height
    rows: int
    diagnostics: dict                  # {"strategy": ..., "gap": ..., "overlaps": 0}

def layout_assembly(pieces: list[Piece]) -> LayoutResult
    """No displacement. Parts keep their projected positions."""

def layout_exploded(pieces: list[Piece], axis2d: np.ndarray, density: str,
                    sheet_aspect: float = 0.75, max_rows: int = 1) -> LayoutResult
    """<<ALGORITHM SPECIFIED IN §11 — implement exactly what is written there.>>"""

def fit_to_frame(result: LayoutResult, frame_w: float, frame_h: float,
                 margin: float = 0.06) -> float
    """Return the uniform scale factor that makes the geometry fill the frame to the target
    occupancy. Never mutates the pieces; the caller applies the scale."""
```

Hard postcondition of `layout_exploded`: **no two pieces' placed bboxes overlap**, verified inside
the function; raise `LayoutError` if the invariant cannot be met.

### 5.3 `labels.py` — pure 2D

```python
@dataclass
class LabelRequest:
    key: str
    numeral: int
    lo: np.ndarray; hi: np.ndarray        # placed part bbox
    anchor_hint: np.ndarray | None = None

@dataclass
class LabelPlacement:
    numeral: int
    key: str
    text_pos: tuple[float, float]
    text_align: str                        # "left" | "right"
    leader: list[tuple[float, float]]      # [anchor_on_part, elbow, landing_end]; 3 points
    score: float

def place_labels(requests: list[LabelRequest], *, obstacles: list[np.ndarray],
                 text_height: float, sheet_lo, sheet_hi) -> list[LabelPlacement]
    """<<ALGORITHM SPECIFIED IN §11 — implement exactly what is written there.>>
    Deterministic: no randomness, stable ordering, stable tie-breaks."""

def text_height_for(slot: float, ratio: float = 0.45) -> float
    """slot * ratio. The QA gate requires height/slot <= 0.6, so 0.45 leaves margin."""
```

Hard postcondition: **zero pairwise overlap among placed numeral boxes**, and no numeral box
intersects any geometry obstacle. If unsatisfiable, raise `LabelError` naming the offending
numerals — the caller turns that into `E_LABELS_UNPLACEABLE` with a "split this figure" hint.

### 5.4 `sheet.py`

```python
LAYERS = ("GEOM", "HIDDEN", "LEADER", "NUM", "TABLE", "CAPTION", "NOTE")

def write_figure(path: Path, *, geometry: list[np.ndarray], hidden: list[np.ndarray] = (),
                 labels: list[LabelPlacement], caption: str, text_height: float,
                 caption_height: float, engineering_rows: list[tuple] | None = None,
                 dxf_version: str = "R2018") -> None
    """All entities and layers CONTINUOUS. Styles: HZ=simfang.ttf, NUM=txt.shx.
    engineering_rows is None for patent figures — the parts table is opt-in only."""

def render_preview(dxf: Path, png: Path, dpi: int = 150) -> None
    """Render FROM THE DXF, not from in-memory geometry."""

VOLATILE_HEADER_VARS = ("$TDCREATE", "$TDUPDATE", "$TDINDWG", "$TDUSRTIMER",
                        "$HANDSEED", "$FINGERPRINTGUID", "$VERSIONGUID", "$ACADVER",
                        "$LASTSAVEDBY", "$MENU", "$DWGCODEPAGE")

def normalized_digest(dxf: Path) -> str
    """SHA-256 over a canonical form: volatile header vars dropped, entity handles dropped,
    entities sorted by (layer, type, rounded coordinates), floats rounded to 6 decimals.
    Two runs of the same plan must produce the same digest."""
```

### 5.5 `qa.py`

```python
@dataclass
class Check:
    id: str; passed: bool; value: Any; threshold: str; detail: str; hint: str

DEFAULT_THRESHOLDS = {
    "geometry_occupancy_min": 0.55,
    "label_overlap_pairs_max": 0,
    "text_slot_ratio_max": 0.6,
    "part_bbox_overlap_pairs_max": 0,     # exploded figures only
    "labels_per_figure_max": 20,
    "non_numeral_text_ratio_max": 0.10,
    "leader_crossing_max": 0,
    "non_continuous_max": 0,
}

FORBIDDEN_TEXT_PATTERNS = [r"^[A-Z]{2,4}[0-9]{4,8}(-|_)", r"^[0-9]{4,6}-[A-Z][0-9]{2}",
                           r"_[0-9]+_[0-9]+$"]

def check_figure(dxf: Path, *, thresholds: dict | None = None,
                 forbidden: list[str] | None = None, kind: str = "exploded") -> dict
    """Returns the qa.json shape from §4.4. Never raises on a bad drawing — it reports."""
```

Every check carries a `hint` that names the plan edit that would fix it, because the model is only
allowed to edit the plan. Example hint for `labels_per_figure`:
`"split fig2 — see assembly.json:split_suggestions"`.

### 5.6 `numbering.py`, `plan.py`

```python
# numbering.py
class Numbering:
    def numeral_of(self, part_name: str) -> int | None
    def label_mode(self, part_name: str) -> str            # once | all | none
    def term_of(self, part_name: str) -> str | None
    def table(self) -> list[tuple[int, str, str]]          # (numeral, term, selector)
    def description_zh(self) -> str

def assign(terms: list[dict], part_names: Sequence[str]) -> Numbering
    """Numerals 1..n issued in `terms` order. A part matched by several selectors takes the
    FIRST matching term (document this precedence). Selectors matching nothing are the caller's
    error to report, not this function's to silently skip."""

# plan.py
@dataclass
class PlanIssue:
    code: str; severity: str; message: str; hint: str; pointer: str   # JSON-pointer-ish

def load_plan(path: Path) -> dict                          # + jsonschema validation
def validate(plan: dict, assembly: dict) -> list[PlanIssue]
def apply_defaults(plan: dict) -> dict                     # merge layout defaults, clamp ranges
```

Error codes: `E_SELECTOR_NO_MATCH`, `E_TERM_LOOKS_LIKE_PART_CODE`, `E_TOO_MANY_LABELS`,
`E_MEMBER_NO_MATCH`, `E_DUPLICATE_FIGURE_ID`, `E_UNKNOWN_ENUM`, `E_SCHEMA`, `W_CLAMPED`,
`W_UNLABELLED_PART`.

## 6. CLI surface

Every CLI: `--help` is the contract for agents; `--json` writes machine-readable output; exit 0 on
pass, 1 on failure, 2 on usage error. No CLI ever prints a stack trace as its only output.

```bash
python3 scripts/analyze_assembly.py ASM.stp -o assembly.json [--include GLOB]... [--exclude GLOB]...
python3 scripts/validate_figure_plan.py plan.json --assembly assembly.json [--json issues.json]
python3 scripts/render_patent_figure.py plan.json --assembly assembly.json -o out/ \
        [--cache .cache/] [--preview] [--only fig2]
python3 scripts/qa_patent_figure.py out/fig1.dxf [--kind exploded] [--json qa.json]
python3 scripts/doctor.py [--json]
python3 scripts/make_test_assembly.py -o tests/fixtures/synthetic.stp
```

`render_patent_figure.py` writes `out/<figure id>.dxf`, `out/<figure id>.png` (with `--preview`),
and `out/reference-numerals.json`. It runs the QA gate on each figure and exits 1 if any fails,
printing the failing checks with their hints.

## 7. Determinism rules (mandatory, testable)

1. No `random`, no `time`, no `uuid`, no `id()`, no `hash()` of objects in any output-affecting path.
2. Never iterate a `set` or a `dict` built from unsorted input — sort explicitly with a total key.
3. Ties break on a stable secondary key (usually the item's index in a sorted list), never on
   insertion order.
4. Round floats before comparison (`round(v, 9)`); never compare raw floats for equality.
5. Numpy: use `np.argsort(..., kind="stable")`.
6. Every module-level constant that affects output lives in **one** place and is named. No magic
   numbers inside function bodies.
7. `normalized_digest` of two runs of the same plan must be equal — `tests/test_golden.py` enforces
   this by rendering twice in one process and once in a subprocess.

## 8. Compliance rules (the R2 fix)

Patent figures carry **only** geometry, leaders, numerals, and the figure caption.
No parts table, no part names, no internal part codes, no dimensions, no title block.
The engineering parts table (`NO./NAME/QTY/REMARK`) is available **only** behind
`layout.engineering_table: true`, is off by default, and files produced with it must be written as
`<id>_engineering.dxf` so a review copy can never be mistaken for a filing copy.

## 9. Reporting format for implementation agents

Return exactly this JSON:

```json
{"files_written": ["scripts/patent_figure/layout.py"],
 "public_api": ["layout_exploded(pieces, axis2d, density, sheet_aspect, max_rows) -> LayoutResult"],
 "decisions": ["gap is a fraction of the median piece extent, not of the total string"],
 "deviations": ["none"],
 "open_issues": ["fit_to_frame margin interacts with label margin — caller must reserve space"],
 "self_test": "python3 -c '...' output summary"}
```

Write real files with the Write tool. Do not print code into the report. Do not run `git`.

## 10. Test fixtures

**Never** use the customer STEP or any real part name in the repository.

`scripts/make_test_assembly.py` is **already written and verified** — do not rewrite it. It emits
`tests/fixtures/synthetic.stp` (64 KB, 8 distinct parts, 11 instances, all named `SYN-*`):

| Part | Shape | Why it is there |
| --- | --- | --- |
| `SYN-A01` | 60×60×6 box | flat base, large footprint |
| `SYN-B02` | ⌀52×18 cylinder | coaxial stack member |
| `SYN-C03` | ⌀18×34 cylinder | tall narrow part — worst case for slot sizing |
| `SYN-D04` | torus R20 r3.5 | silhouette with no B-rep edge (needs `OutLine*`) |
| `SYN-E05` | ⌀44×5 cylinder | thin disc |
| `SYN-F06` | ⌀32 sphere | pure smooth silhouette |
| `SYN-G07` | 44×44×4 box | cover plate |
| `SYN-H08` ×4 | ⌀4×10 cylinder | four instances — exercises qty, `label: once/all/none` |

Verified round trip: `cad_hlr_to_dxf.py tests/fixtures/synthetic.stp --list-parts` reports
`11 instance(s), 8 distinct part(s)` with every `SYN-*` name intact.

**No degenerate part is in the STEP, on purpose.** A bare vertex loses its name in the STEP round
trip and returns labelled with the OCCT version string, which would make the golden digest depend
on the installed OCCT build. Cover the degenerate path in `tests/test_analyze.py` by constructing a
`PartShape` with `lo == hi` directly.

Real-assembly runs happen outside the repo, under the session scratchpad.

## 11. Algorithms — filled from the design panel

> This section is populated after the design workflow completes. Until then, implementers of
> `layout.py` and `labels.py` must wait; every other module is fully specified above.
