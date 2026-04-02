---
name: hifi-comment-tagging
description: Analyze HIFI customer comments, return reasons, and after-sales feedback from Excel workbooks. Use when Claude needs to clean source data, focus analysis on one product such as P4 or ZD3, apply structured Chinese tagging across 1-4 levels, and generate summary tables plus product-manager style conclusions.
---

# Hifi Comment Tagging

Clean Excel source data, focus on one product, apply the team's tagging habits, and produce consistent summary outputs for comment, return, and after-sales analysis.

## Workflow

Run this workflow in order:

1. Confirm the target product. If the user did not name one product, stop and ask.
2. Ask the user if they want to exclude blank data (rows without usable feedback text). If they want to keep blank data, pass `--keep-blank` to `clean_product_comments.py`.
2. Profile the workbook with `scripts/profile_workbook.py` to identify candidate sheets, headers, and product signals.
3. Read the input contract in `references/input_contract.md`.
4. Clean the source data with `scripts/clean_product_comments.py`.
5. Read `references/taxonomy.md` before assigning or reviewing tags.
6. Read `references/layout.md` and `references/summary_patterns.md` before generating summary output.
7. If the workbook is already manually labeled, mine reusable patterns with `scripts/extract_taxonomy_examples.py`.
8. If tagged rows already exist, build a summary scaffold with `scripts/build_summary_scaffold.py`.
9. IMPORTANT: Any comments marked as `待人工判定` or `待人工筛查` must NOT be written to the final summary table. `build_summary_scaffold.py` will intercept them and output them. You must list these problematic comments in the chat dialog and ask the customer for guidance before finishing the analysis.

## Required Input

Require these inputs from the user:

- A target product name, for example `P4`, `ZD3`, `ZA3`, `ZP3`, `LC30`
- At least one `.xlsx` workbook
- If the workbook has many sheets, ask which sheet is raw data only when profiling cannot determine it

Reject ambiguous requests such as “analyze this workbook” when the workbook contains multiple products and the user did not specify one product.

## Product Focus Rules

Focus on one product per run unless the user explicitly asks for multi-product output.

Use these rules in order:

1. Prefer rows whose explicit product column matches the target product.
2. Otherwise prefer sheets whose name contains the target product.
3. Otherwise use the workbook filename only as a weak fallback.
4. If multiple plausible locations remain, show the candidates and ask the user to choose.

Do not silently mix different products into one summary.

## Cleaning Rules

Apply only standard, reviewable cleaning:

- By default, remove empty comments, `NA`, `N/A`, `同上`, `same as above`, and rows without usable feedback text. If the user decides to keep them, use `--keep-blank`.
- Prefer `中文翻译` as the analysis text. Fall back to original comment text when translation is missing.
- Preserve source lineage with workbook path, sheet name, and row number.
- Remove obvious duplicates based on normalized comment text within the target product.
- Keep pre-existing manual labels if they are present in the workbook.

Read `references/cleaning_rules.md` for the exact normalization and invalid-row policy.

## Tagging Standard

Use the canonical chain:

- `一级分类`
- `二级分类`
- `三级问题点`
- `四级归因`

Map older structures into this chain when needed:

- `问题分类：1级 -> 一级分类`
- `问题分类：2级 -> 二级分类`
- `问题分类：3级 -> 三级问题点`

Use the canonical taxonomy in `references/taxonomy.md`. Prefer existing team labels over inventing new ones. If a comment is too vague, keep the original text and mark the uncertain level as `待人工判定`, `暂无法判断`, `NA`, or `未提供` as appropriate.

## Output Structure

Default outputs:

- A formatted Excel workbook, not just plain text
- `CleanedComments`: standardized cleaned rows for the target product
- `TaggedComments`: cleaned rows plus the full label chain
- `Summary`: grouped counts, product-manager style summary text, and a visible convergence path from source classes to sorted sub-classes to re-grouped macro themes

Follow `references/layout.md` for sheet ordering and table layout.

When the source sample is large enough to benefit from visual scanning, generate restrained charts directly inside the Excel workbook. Prefer low-color line charts and keep data tables as the primary evidence surface.

## Summary Writing Rules

Generate summary text in Chinese.

Always include:

- Total records and valid feedback count
- Top 3 categories with percentages
- A short overall diagnosis
- A visible “how the analysis converged” path from big class to small class to macro theme
- Numbered key issues
- Risk emphasis for fault, compatibility, power, silence, protection, or noise issues
- Reasonable demand emphasis when the user is asking for missing but valid features
- Rich-text emphasis in Excel: one-line summary bold, key words red, punctuation left uncolored

Follow the tone and structure in `references/summary_patterns.md`.

Important wording boundary:

- Only rewrite representative comments or example snippets when they appear in table/list fields meant for quick scanning.
- Do not force the final long-form summary text into shorthand, abstraction labels, or rewritten example phrases just to make it shorter.
- The final summary block should stay as normal product-manager narrative Chinese, with complete sentences and clear judgment.

## Scripts

Use these bundled scripts directly:

- `scripts/profile_workbook.py`: inspect workbook structure and candidate products
- `scripts/clean_product_comments.py`: standardize and filter target-product rows into `CleanedComments`
- `scripts/extract_taxonomy_examples.py`: mine label chains and examples from historical labeled workbooks
- `scripts/build_summary_scaffold.py`: build a styled summary workbook from tagged rows

## Notes

- Keep the workflow transparent and reviewable. Do not do opaque clustering or hidden relabeling.
- Treat this skill as an analysis accelerator, not an automatic ground-truth generator.
- When the user supplies a gold-standard manually labeled workbook, use it as the primary pattern source for wording, hierarchy, and summary style.
