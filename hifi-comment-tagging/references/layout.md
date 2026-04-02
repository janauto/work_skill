# Layout Guide

Keep the output close to the existing team workbook style without requiring pixel-perfect cloning.

## Sheet Order

Use this order when generating a new workbook:

1. `CleanedComments`
2. `TaggedComments`
3. `Summary`

If the task only asks for a scaffold, still return a formatted Excel workbook with at least `TaggedComments` and `Summary`.

## Detail Sheet

Preferred left-to-right column order:

- `店铺`
- `国家`
- `中文翻译`
- `一级分类`
- `二级分类`
- `三级问题点`
- `四级归因`

It is acceptable to prepend standard lineage fields such as:

- `record_id`
- `product_name`
- `source_sheet`
- `source_row`

Format detail sheets for direct review:

- Freeze the header row
- Apply bold filled headers
- Enable text wrapping on comment fields
- Set readable widths for text and label columns
- Preserve one row per comment for manual audit

## Summary Sheet Structure

Use four areas and make the convergence logic visible.

### Left block

Columns `A:E`

- Title: `源分类：按大类到小类逐层筛选`
- `一级分类`
- `二级分类`
- `三级分类`
- `数量`
- `占比`

Group by `一级分类`. This block should show the raw classification path.

### Middle block

Columns `G:J`

- Title: `按降序排列：先看高频小类`
- `上级分类`
- `小类`
- `数量`
- `占比`

Sort by count descending so the reader can quickly see where to focus.

### Right-middle block

Columns `L:P`

- Title: `收敛归纳：按类别再次分大类`
- `原分类`
- `收敛大类`
- `数量`
- `占比`
- `说明`

This block is where the workbook shows the thinking process: how raw labels are collapsed into larger decision themes.

### Right block

Columns `R:Y` or another merged area to the far right.

Place the long-form Chinese summary here with:

- One overview paragraph
- Top 3 categories
- Numbered key issues
- Risk callouts
- Optional optimization suggestions

Apply Excel rich-text formatting:

- The one-line summary sentence should be bold
- Key terms and key percentages should be red
- Do not color punctuation marks

### Lower chart area

When valid sample size is at least 20 rows, reserve space below the text summary for simple charts.

Preferred chart types:

- Line chart for Top category share
- Optional second line chart for converged macro-theme share
- Keep the underlying data table visible next to the chart

## Style Cues

Use style cues from current house files:

- Main summary headers: blue fill, white or dark readable text, bold
- Secondary block headers: light blue or light gray fill
- One big category should use one consistent light color block across its rows
- Use only light pastel fills, not saturated colors
- Summary text: merged cells with wrapped text and top alignment
- Percentages: 4 decimal places in data tables, 2 decimal places in prose
- Percentage text in tables should be red
- Charts: low-color, mostly blue/gray lines, keep titles short
- Keep text neat and uniform: same font family, same readable font size, centered headers, top-aligned wrapped body text

## Representative Comment Rule

If a summary sheet needs a representative comment field, do not paste original quotes by default.

Use a short analyst-written abstraction such as:

- `用户希望支持 A+B 同时播放`
- `用户担心切换过程影响音质`
- `用户反馈 VU 表灵敏度不足`

Keep these phrases short, clean, and reusable.

This rule applies only to list-like fields such as:

- `代表评论`
- `概览短语`
- `说明`
- other compact table cells used for scanning

It does not apply to the final long-form summary block on the right side.
The final summary should keep full narrative wording and should not be compressed into shorthand labels just for neatness.

## What Not To Do

- Do not flatten everything into one raw pivot table
- Do not hide the convergence logic from big class to small class to macro theme
- Do not use pie charts or high-saturation charts as the default
- Do not introduce a brand-new column order unless the user gave a fixed template
