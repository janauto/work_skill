---
name: product-definition-voc
description: Analyze user-review VOC for product-definition work from Excel workbooks. Use when Claude needs to clean raw review exports, extract what users love and why they churn, tag opinions and hidden needs, identify aha moments, map emotional heat by feature or scene, and generate a formatted Chinese summary workbook for concept definition, competitor learning, or pre-PRD insight synthesis.
---

# Product Definition VOC

Clean review workbooks for product-definition research, then turn raw comments into structured product insight.

## Workflow

Run this workflow in order:

1. Confirm the analysis focus. Accept one product, one category, one scenario, or one curated competitor set.
2. Profile the workbook with `scripts/profile_workbook.py` to identify candidate sheets, headers, comment columns, and image columns.
3. Read `references/input_contract.md`.
4. Clean and dedupe the source with `scripts/clean_voc_comments.py`.
5. Read `references/taxonomy.md` before reviewing or adjusting tags.
6. Read `references/layout.md` and `references/summary_patterns.md` before generating the final workbook.
7. Apply first-pass tagging with `scripts/tag_voc_comments.py`.
8. Build the styled output workbook with `scripts/build_voc_summary_workbook.py`, ensuring `AhaMoments` shows a concise insight line, source raw comment text, and product link when available.
   For `NeedClusters`, prefer semantic Maslow layering via a large model when API credentials are available; otherwise use the built-in semantic fallback.

## Required Input

Require these inputs:

- At least one `.xlsx` workbook
- A focus statement, unless the workbook is already clearly scoped

Acceptable focus examples:

- `LC40`
- `RCA 切换器竞品`
- `桌面双机切换场景`
- `passive preamp`

If the workbook is already a single-scope export, the focus can be omitted. If profiling shows multiple unrelated products or sheets, ask the user to narrow scope only after profiling fails to disambiguate it.

## Focus Rules

Use these rules in order:

1. Prefer rows whose product, scene, keyword, or listing title clearly matches the user focus.
2. If the workbook is already pre-filtered to one product or one category, keep the whole workbook.
3. If several plausible scopes remain, show the candidates and ask the user to choose.
4. Do not silently merge unrelated categories into one final summary.

## Cleaning Rules

Apply only standard, reviewable cleaning:

- Remove duplicates based on normalized comment text inside the chosen focus
- Drop empty rows and obvious garbage such as pure emoji, pure punctuation, and ultra-short generic filler
- Drop rows that are only logistics, seller service, coupon spam, review hijacking, or obvious off-topic content
- Prefer translated Chinese text when available
- Preserve lineage with workbook path, sheet name, row number, rating, and image refs
- Keep dropped rows in an audit sheet when generating the cleaned workbook

Read `references/cleaning_rules.md` for the exact drop policy.

## Tagging Standard

Use these canonical layers:

- `观点类型`
- `一级功能`
- `二级功能需求`
- `底层需求`
- `决策信号`
- `情绪极性`
- `情绪强度`
- `场景标签`
- `嘿哈时刻`

Use `references/taxonomy.md` as the default system. Prefer stable reusable labels over inventing near-synonyms. When a comment is genuinely ambiguous, keep the original text and mark the uncertain field as `待人工复核` or `待补充`.

## Output Structure

Default outputs:

- A formatted Excel workbook
- `CleanedComments`
- `TaggedComments`
- `NeedClusters`
- `AhaMoments`
- `EmotionMap`
- `SceneCards`
- `Summary`

Follow `references/layout.md` for sheet order and block layout.

## Summary Writing Rules

Generate summary text in Chinese.

Always include:

- Total raw rows and valid kept rows
- One-word topic tags for the first summary block
- A second-level functional need under each top-level function
- What users most clearly like
- Why users hesitate, churn, or stop using
- Top hidden needs with percentages
- A separate aha-moment section
- Emotion heat zones by feature or scenario
- Sorted scenario cards with representative chains
- Product-definition opportunities stated as concrete hints, not generic slogans

Follow `references/summary_patterns.md`.

## Scripts

Use these bundled scripts directly:

- `scripts/profile_workbook.py`
- `scripts/clean_voc_comments.py`
- `scripts/tag_voc_comments.py`
- `scripts/build_voc_summary_workbook.py`

## Notes

- Keep the process transparent and auditable.
- Treat rule-based tags as a first-pass scaffold, not automatic ground truth.
- Preserve images or image refs when present because they can strengthen scene and aha interpretation.
- Keep top-level function labels concise, ideally one Chinese noun such as `音量`, `连接`, `音质`, `切换`.
- In `Summary`, keep first-level-function rows visually grouped with restrained light background colors for fast scanning.
- This skill is for product-definition insight, not return-rate attribution. When the user wants return or after-sales analysis, prefer a dedicated return-analysis skill instead.
