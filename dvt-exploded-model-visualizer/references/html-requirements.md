# Interactive HTML Requirements

Read this reference when generating or updating the exploded-view page.

## Required presentation states

1. **Assembled:** source assembly at zero explosion. Opaque covers hide internal proposed parts.
2. **Exploded:** parts separate along understandable service or assembly directions.
3. **X-ray/cutaway:** external covers become transparent while internal source and proposal parts remain selectable.
4. **Local close-up:** camera targets the modified region without losing orientation.

## Core controls

- Explosion slider with visible percentage.
- X-ray toggle.
- Standard isometric, front, side, and top views.
- Module visibility controls and counts.
- Click-to-inspect with source/status/material/note fields.
- Reset that returns to a predictable review state.

When multiple local concepts exist, provide explicit variants with a recommendation, tradeoffs, risks, and validation requirements. Changing a variant must rebuild only its proposal group and leave source CAD unchanged.

## Input-readiness panel

Use a visible panel titled “开始前请放入这些文件”. It should show:

- `必需`: STEP/STP or visualization-only GLB/GLTF; change request.
- `局部修改冻结前必需`: local section, dimensions/tolerances, mating geometry, material/CMF, and tooling data when relevant.
- `建议`: BOM/name map, photos/renders, existing HTML/metadata, assembly order, and validation targets.
- `状态`: 已找到 / 缺失 / 仅可视化 / 占位文件 / 待确认版本.

Explain that the user should put files in the same project folder or explicitly provide their paths. Do not imply browser upload functionality unless it actually exists.

## Visual truthfulness

- Source CAD should use a consistent neutral material.
- Proposal geometry should use a separate restrained palette and be identified in the legend.
- Surface processes such as coating and laser marking should render as a flat surface layer, not as a thick solid.
- Use real section shapes and thicknesses for plates, rings, guides, FPC, supports, and screws.
- Show internal parts only when exploded or x-rayed if they would be hidden in the physical product.

## Output and browser checks

- Keep the page usable without a development server when the user expects a local HTML artifact.
- Avoid remote dependencies for a self-contained handoff unless the user agrees to network requirements.
- Wait for the model-load completion state before asserting success.
- Exercise every visible control and variant at least once.
- Check a desktop viewport and a narrow mobile viewport; ensure document width does not exceed viewport width.
- Capture screenshots for zero-explosion assembled, exploded, and x-ray/local-close-up states.
- Treat browser console or page errors as failures.
