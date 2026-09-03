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

Also verified on the dev machine, so do not spend agent time rediscovering it:

- `ezdxf.addons.drawing` with the `matplotlib` backend imports cleanly — use it for `render_preview`.
- `simfang.ttf` is **not installed** here (only `Songti.ttc`). `doctor.py` must report this as a
  real finding, not a hard failure: the DXF still records the style name `simfang.ttf`, and the
  consequence is that Chinese text renders with a fallback locally and may box out on a machine
  without the font. The portable-DWG outlined-text route in `references/workflow.md` §4 exists for
  exactly this case.
- AutoCAD 2026 **is** installed at `/Applications/Autodesk/AutoCAD 2026`, and LibreDWG `dwgread`
  **is** on `PATH` at `/opt/homebrew/bin/dwgread`. Both conversion routes are testable here.
Python is 3.9. Verified on this machine: `list[int]` / `dict[str, int]` **do** work at runtime
(PEP 585 landed in 3.9), but `int | str` raises `TypeError`, and `match` does not parse. Add
`from __future__ import annotations` to every module anyway — it makes all annotations lazy, so the
union syntax used throughout this contract's signatures is legal.

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

When the assembly exceeds the label cap, all three strategies are **attempted** — `coaxial` (by
coaxial group), `stack` (by breaks in `stack_order` spacing), `size` (by `size_tiers`) — and a
strategy is emitted only if every figure it proposes satisfies `labels <= max_labels_per_figure`.
A strategy that cannot reach the cap is reported in `warnings` with the count it got to, never
emitted as a suggestion. If **no** strategy succeeds, `split_suggestions` is `[]` and `warnings`
says so; that is a legitimate outcome, not a bug — the human then has to choose the grouping.

(The earlier wording said three suggestions are "always emitted" *and* that each must satisfy the
cap. Those cannot both hold. Flagged by the cross-model contract audit; see §12.)

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
  `once` labels **exactly one instance**: the one whose `key` sorts first among that part's
  instances in the figure (`key` is `name#instance_index`, and `instance_index` is assigned after
  the global sort in §5.1, so this is stable). `all` labels every instance; `none` labels none and
  the part still gets a numeral in the global table only if some other figure labels it.
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
    def key(self) -> str
        """Exactly f"{az:.4f}_{el:.4f}_{roll:.4f}" — four decimals, no rounding of the stored
        floats themselves. Four is enough that a 1e-4 degree difference busts the cache and not
        enough for float noise to do so."""

def roll_for_axis(axis, az: float, el: float, target_deg: float) -> float
    """Solve in closed form, do NOT search. With e = the axis projected into the unroll ed view
    plane, the sheet angle is atan2(e_u, e_x); the roll that moves it to target is

        roll = target_deg - degrees(atan2(e_u, e_x))

    normalised into [0, 360). The legacy implementation stepped through candidate angles and
    compared raw floats, which violates §7 rule 4; the closed form removes both problems.
    Raise ValueError when |e| < 1e-9 (axis parallel to the view direction)."""

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
    """Deterministic order. The sort key is exactly:

        (name, path, round(cx, 6), round(cy, 6), round(cz, 6))

    where c is the instance centre in model coordinates and 6 is the decimal count — fixed here so
    two implementers cannot pick different precisions. Ties beyond that key keep the order OCCT
    produced, which is NOT guaranteed stable across OCCT builds; if two instances tie on the full
    key they are geometrically identical at micron resolution, so which one takes which
    instance_index cannot change any output. instance_index is assigned AFTER this sort, so the
    same STEP always yields the same keys."""

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
# Gap between adjacent parts on the sheet, as a fraction of the MEDIAN part extent along the
# explode direction — not of the total string length. Keying it to the median is what stops the
# v1 failure where a longer string shrank every part's slot while the gap grew.
DENSITY: dict[str, float] = {"compact": 0.35, "normal": 0.55, "loose": 0.85}

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

TARGET_OCCUPANCY = 0.62   # QA floor is 0.55 (§5.5); 0.62 leaves headroom for the label margin
FRAME_W, FRAME_H = 180.0, 250.0   # A4 portrait minus a 15 mm margin, in millimetres

def fit_to_frame(result: LayoutResult, frame_w: float = FRAME_W, frame_h: float = FRAME_H,
                 margin: float = 0.06) -> float
    """Return the uniform scale factor s such that the geometry bbox, scaled by s and inset by
    `margin` on each side, occupies TARGET_OCCUPANCY of frame_w x frame_h by AREA:

        usable_w = frame_w * (1 - 2*margin);  usable_h = frame_h * (1 - 2*margin)
        s_fit    = min(usable_w / bbox_w, usable_h / bbox_h)        # largest that still fits
        s_target = sqrt(TARGET_OCCUPANCY * usable_w * usable_h / (bbox_w * bbox_h))
        s        = min(s_fit, s_target)

    Taking the min means a figure whose aspect ratio cannot reach the target occupancy is fitted
    rather than overflowed; QA then reports the shortfall instead of the renderer hiding it.
    Never mutates the pieces; the caller applies the scale."""
```

Hard postcondition of `layout_exploded`: **no two pieces' placed bboxes overlap**, verified inside
the function with a tolerance of `1e-6` on each axis (touching edges are not an overlap); raise
`LayoutError` if the invariant cannot be met.

All four audit models asked whether that escape hatch is reachable and well defined. It is, and
this is how: the invariant is always satisfiable for `layout_exploded` on its own, because parts
are laid end to end along one direction with a positive gap, so `LayoutError` can only fire on
degenerate input — a piece with zero extent, or a non-finite coordinate. What is **not** always
satisfiable is the combination of zero overlap *and* the occupancy target *and* the label budget;
that tension is resolved by `fit_to_frame` returning a smaller scale and QA reporting the
shortfall, never by the layout silently overlapping parts.

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
intersects any geometry obstacle. Unlike the layout invariant this one **is** genuinely
unsatisfiable for dense inputs, and that is by design: it is the mechanism that forces a figure to
be split rather than allowing the v1 outcome where 97% of numerals overlapped. Raise `LabelError`
carrying `unplaced: list[int]` (the numerals) and `tried: int` (candidate positions per numeral);
the caller turns it into `E_LABELS_UNPLACEABLE` whose hint names the figure id and points at
`assembly.json:split_suggestions`. `place_labels` must never fall back to placing a label anyway.

`obstacles` is every placed geometry polyline in the figure plus every leader and numeral box
already committed during this call — i.e. it grows as the greedy loop proceeds. The caller passes
only the geometry; `place_labels` accumulates the rest itself.

### 5.4 `sheet.py`

```python
LAYERS = ("GEOM", "HIDDEN", "LEADER", "NUM", "TABLE", "CAPTION", "NOTE")

def write_figure(path: Path, *, geometry: list[np.ndarray], hidden: list[np.ndarray] = (),
                 labels: list[LabelPlacement], caption: str, text_height: float,
                 caption_height: float, engineering_rows: list[tuple] | None = None,
                 dxf_version: str = "R2018") -> None
    """All entities and layers CONTINUOUS. Styles: HZ=simfang.ttf, NUM=txt.shx.
    engineering_rows is None for patent figures — the parts table is opt-in only.

    MUST set `doc.header["$INSUNITS"] = 4` explicitly. ezdxf 1.4.2 defaults a new R2018 document
    to 6 (metres), verified; leaving it would silently label every millimetre figure as metres."""

def render_preview(dxf: Path, png: Path, dpi: int = 150) -> None
    """Render FROM THE DXF, not from in-memory geometry."""

# Verified against ezdxf 1.4.2 on a fresh R2018 document:
#   $TDCREATE/$TDUPDATE  float julian day        -> volatile
#   $HANDSEED            str, e.g. '100'         -> volatile (entity-count dependent)
#   $FINGERPRINTGUID     str '{...}' per document-> volatile
#   $VERSIONGUID         str '{...}' per save    -> volatile
#   $ACADVER             str 'AC1032'            -> NOT volatile. It is the DXF version itself;
#                                                   stripping it would hide a real format change.
VOLATILE_HEADER_VARS = ("$TDCREATE", "$TDUPDATE", "$TDINDWG", "$TDUSRTIMER",
                        "$HANDSEED", "$FINGERPRINTGUID", "$VERSIONGUID",
                        "$LASTSAVEDBY", "$MENU", "$DWGCODEPAGE")

def normalized_digest(dxf: Path) -> str
    """SHA-256 over a canonical form. The recipe is fixed here in full because three of the four
    audit models independently picked different sort keys from the previous one-line description:

    1. Read with ezdxf. Ignore the header entirely except to assert `$INSUNITS == 4`.
    2. For each modelspace entity build a tuple:
         (layer, dxftype, style_or_empty, text_or_empty, tuple_of_rounded_coords)
       where every coordinate is `round(float(v), 6)` and coords are flattened in the entity's
       own vertex order (LINE: start then end; LWPOLYLINE: points in stored order; TEXT: insert
       point then height). Entity handles and owner handles are never read.
    3. Sort the tuples with Python's default tuple ordering — a total order on these values, so no
       tie-break is needed. Do NOT use np.argsort here.
    4. Join with "\n" using repr() of each tuple, encode UTF-8, SHA-256.

    Rounding at 6 decimals and comparing at 6 decimals is the same operation here, so the
    round-then-hash cannot disagree with round-then-compare (a hazard one audit model raised).
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
5. Numpy: `np.argsort` defaults to `kind='quicksort'`, which is **not** stable — verified against
   numpy 2.0.2. Always pass `kind="stable"` explicitly. Prefer plain `sorted()` on tuples when the
   data is small; it is stable by definition and needs no flag.
6. Globs: use `fnmatch.fnmatchcase`, never `fnmatch.fnmatch`. `fnmatch` routes through
   `os.path.normcase`, which is identity on macOS/Linux and lowercasing on Windows — verified — so
   the same plan would select different parts on different platforms. This applies to every glob in
   the system: `terms[].selector`, `figures[].members`, `source.include/exclude`.
7. Every module-level constant that affects output lives in **one** place and is named. No magic
   numbers inside function bodies.
8. `normalized_digest` of two runs of the same plan must be equal — `tests/test_golden.py` enforces
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


---

## 12. Cross-model contract audit (2026-09-03)

Before implementation started, this contract was audited by five models from five vendors, each
given the identical adversarial brief: *falsify the claim that any competent LLM reading this
document produces the same behaviour.* Reports are outside the repo (they quote nothing
confidential, but they are working notes, not deliverables).

| Runtime | Model | Report |
| --- | --- | --- |
| CodeBuddy | GLM-5.1 (智谱) | 25 KB |
| CodeBuddy | Kimi-K2.5 (月之暗面) | 22 KB |
| CodeBuddy | DeepSeek-V3.2 | 26 KB |
| CodeBuddy | MiniMax-M2.7 | 27 KB |
| Codex CLI | GPT-5.x | (see run log) |

**Acceptance rule: a finding counted only if ≥2 models raised it independently, or if one model
raised it and it was reproduced against the actual library.** Every library claim below was
re-verified locally rather than taken on trust — one of them turned out to be wrong.

| Finding | Models | Verified | Fix |
| --- | --- | --- | --- |
| §11 empty — the two hardest algorithms unspecified | 4/4 | known | pending the design review |
| `load_assembly` sort-key rounding precision unspecified | 4/4 | yes | fixed at `round(v, 6)`, §5.1 |
| `normalized_digest` sort key and rounding under-specified | 4/4 | yes | full recipe written out, §5.4 |
| zero-overlap postconditions unsatisfiable on dense input | 4/4 | mathematically | escape hatch made explicit, §5.2/5.3 |
| `fit_to_frame` target occupancy undefined | 3/4 | yes | `TARGET_OCCUPANCY = 0.62` + formula, §5.2 |
| `np.argsort` default is not stable | 3/4 | `quicksort` | §7 rule 5 strengthened |
| `roll_for_axis` algorithm unspecified, legacy compares raw floats | 3/4 | yes | closed form specified, §5.1 |
| `DENSITY` values have no stated semantics | 3/4 | yes | defined against median part extent, §5.2 |
| OCCT traversal order not guaranteed | 3/4 | plausible | already covered by the §5.1 sort; rationale added |
| `$ACADVER` is not a volatile header var | 2/4 | `'AC1032'` | removed from `VOLATILE_HEADER_VARS` |
| `fnmatch` is case-sensitive per platform | 2/4 | yes | `fnmatchcase` mandated, §7 rule 6 |
| `split_suggestions` "always emitted" contradicts its own cap | 1/4 | logical | rewritten, §4.1 |
| `label: "once"` — which instance gets the numeral | 1/4 | yes | first by sorted key, §4.2 |
| `$INSUNITS` defaults to 6 (metres), not absent | 1/4 | `6` | `write_figure` must set 4, §5.4 |
| `list[int]` works at runtime on Python 3.9 | 1/4 | yes | §1 relaxed |
| **`jsonschema` is not installed** | 1/4 | **false — 4.25.1 is installed** | rejected |

The last row is the reason for the verify-locally rule. A single model's confident factual claim
about the environment was simply wrong; had it been actioned, the plan validator would have been
written around a dependency that was there all along.

Two failure modes of the harness itself are worth recording, because they cost a full round:
`codebuddy --disallowedTools` is variadic and swallowed the prompt that followed it, producing four
empty reports with exit code 0 — a silent success. Use the single-valued `--tools` instead. And the
first local verification script called `importlib.metadata` without importing it, which is what
produced the bogus "jsonschema missing" reading on the first pass.
