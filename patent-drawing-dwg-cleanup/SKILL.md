---
name: patent-drawing-dwg-cleanup
description: Generate or clean patent figures as editable DXF/DWG, convert dashed geometry to continuous lines, remove patent reference numerals without deleting functional labels, and validate the result. Use for patent line drawings sourced from a 3D CAD assembly (STEP), DWG, DXF, Matplotlib, PDF, or raster images.
---

# Patent Drawing DWG Cleanup

Produce a clean, editable CAD deliverable while preserving the source files. The expected result is continuous-line geometry, no patent reference numerals requested for removal, retained functional labels, and an AutoCAD-audited DWG plus an editable DXF.

## Route by source type

- 3D CAD assembly (STEP/STP/IGES): this is the highest-fidelity source. Compute the figure with
  analytic hidden-line removal instead of tracing anything, using `scripts/cad_hlr_to_dxf.py`.
  Read [references/cad-source-to-drawing.md](references/cad-source-to-drawing.md) before the first run.
- Existing DXF: inspect the text candidates, clean a copy with `scripts/clean_patent_dxf.py`, then validate it with `scripts/validate_clean_dxf.py`.
- Existing DWG: open or convert a copy to DXF with AutoCAD first. Do not treat a compressed DWG as a text-editable file or rewrite it with `ezdxf`.
- Matplotlib or other vector source: preserve the source geometry and export its artists or paths to DXF. Read [references/workflow.md](references/workflow.md) for the artist-to-CAD mapping.
- PDF, PNG, or JPG only: confirm whether the figure is vector or raster. Extract vector paths when present; otherwise trace or reconstruct the geometry. Do not claim that raster-to-vector output is dimensionally exact unless dimensions or controlled CAD are available.

## Separate reference numerals from functional text

First inventory every text object containing numbers. Treat standalone multi-digit labels such as `100`, `310`, or `3411` as candidates, not automatic proof that they are patent references. Preserve functional identifiers and engineering content such as `A1`, `A2`, `24V`, pin names, dimensions, tolerances, and terminal labels unless the user explicitly says to remove them.

Use an explicit accepted list whenever the drawing contains dimensions or mixed technical notation:

```bash
python3 scripts/clean_patent_dxf.py input.dxf output.dxf \
  --reference-number 100 \
  --reference-number 310 \
  --reference-number 3411 \
  --strip-selected-inline-references \
  --remove-figure-labels \
  --report cleanup-report.json
```

For a drawing known to use every standalone multi-digit text object as an attached reference label, `--remove-standalone-candidates` is available. Run `--dry-run` first and review the report. Do not use broad numeric deletion on dimensioned drawings.

## Make linework continuous

Set graphical entities to `CONTINUOUS` and set their layers to `CONTINUOUS`; checking only the visible plot is insufficient. Preserve geometry, arrows, junctions, and functional symbols. Convert former dashed boxes, centerlines, and chain lines to solid geometry only when the user asks for all-solid figures.

Generate a preview image after cleaning and compare it with the source. The preview must show that no outline, arrowhead, junction, or retained label was lost.

## Create the DWG

Prefer Autodesk's own core engine over third-party DWG converters when AutoCAD is installed. The bundled converter writes to an ASCII temporary path so Chinese output paths do not get corrupted by the command-script prompt:

```bash
python3 scripts/autocad_core_dxf_to_dwg.py cleaned.dxf cleaned.dwg
```

The converter runs `AUDIT`, saves as AutoCAD 2018 DWG, re-audits, and reports entity and non-continuous counts. Pass `--autocad-core` if AutoCAD is installed in a nonstandard location.

When AutoCAD is not installed, fall back to LibreDWG:

```bash
python3 scripts/libredwg_dxf_to_dwg.py cleaned.dxf -o dwg/ --deep-audit --report dwg-report.json
```

It writes R2000 and uses `dwgread` in place of the AUDIT round trip: a file that re-parses cleanly
with a non-zero entity count and intact text passes. Say in the handoff which engine produced the DWG.
LibreDWG does not scale to very large combined sheets; convert per figure and deliver an oversized
combined sheet as DXF rather than reporting a DWG that never completed.

If retained Chinese text is substituted or garbled, keep the editable-text DXF and create a second DWG in which the text is converted to vector outlines before conversion. Label that tradeoff clearly: outlined text is portable but no longer editable as text.

## Validate before delivery

Run:

```bash
python3 scripts/validate_clean_dxf.py cleaned.dxf
python3 scripts/autocad_core_dxf_to_dwg.py cleaned.dxf cleaned.dwg
```

Deliver only after these checks pass:

1. Source files remain unchanged and outputs are in a new folder or new filenames.
2. The DXF contains no rejected reference labels and no non-continuous entity or layer linetypes.
3. The DWG opens in AutoCAD, the final `AUDIT` reports zero errors, and the entity count is nonzero.
4. A visual preview matches the intended source geometry.
5. The handoff states which labels were preserved and whether DWG text is editable text or vector outlines.

Read [references/workflow.md](references/workflow.md) when rebuilding from Matplotlib/raster sources, combining multiple figures, or handling Chinese-font portability.
