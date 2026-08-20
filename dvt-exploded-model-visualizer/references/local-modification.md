# Local Modification Guidance

Read this reference whenever proposed geometry is added around an existing product region.

## Anchor before modeling

1. Identify the nearest source part and its controlled name or source solid ID.
2. Measure its bounding box, center, face direction, and adjacent clearance from source geometry.
3. Identify the local exterior surface and the inside direction.
4. Record which values are measured and which are provisional.

If the local face or section cannot be resolved, create a visibly separate concept view and request the missing section/drawing in the page's input-readiness panel.

## Model physical construction

- Use annular plates or extrusions for flat rings, not round torus tubes.
- Use near-zero-thickness planes/decals for paint, laser ablation, graphics, and coatings.
- Give light guides a real optical section and coupling direction.
- Give light chambers inner and outer opaque walls instead of a generic dark ring.
- Include mounting ears, screw bosses, snap directions, heat stakes, adhesive zones, or compression seals when the proposal depends on them.
- Model FPC as a thin board with a tail, bend path, keep-out region, and connector destination.
- Show fasteners from the actual insertion side. Avoid visible exterior screws when the proposal says internal lock.
- Preserve a central touch/electrode path when adding a halo around an existing capacitive touch element.

## Keep external flush surfaces honest

For a flush, no-through-hole illuminated mark:

1. Use a light-transmitting substrate rather than bulk-pigmented opaque plastic.
2. Keep the external mold surface continuous.
3. Apply an opaque coating and remove only the mark area by laser, or use another verified hidden-graphic process.
4. Put local thinning, diffuser, guide, opaque holder, FPC, LEDs, and fasteners on the internal side.
5. In zero-explosion opaque view, show only the flat illuminated graphic.

State clearly that changing an old opaque shell to an optical substrate can require material, molding, coating, laser, and validation changes even if the external cavity is unchanged.

## DVT deliverables for a local proposal

- Front-to-back section with preliminary dimensions.
- Proposed BOM and reuse boundary.
- Assembly and disassembly path.
- Manufacturing process with inspection gates.
- Mold-change hypothesis: cut steel, add steel/weld, replace insert, or no mold geometry change.
- Risks covering fit, optical uniformity, touch/EMI, thermal, appearance, coating, and reliability.
- Explicit list of drawings and evidence needed before tooling release.

Do not convert a local overlay into a claim of direct fit merely because it looks aligned in the viewer.
