# Input Contract

Use this contract before analyzing any workbook.

## Required

- Target product name
- At least one `.xlsx` workbook

## Preferred

- A workbook with one raw-data sheet and one manually labeled sheet
- Existing summary sheets if the user wants to keep the house style
- A note telling you which sheet is raw data when the workbook is large

## Acceptable Source Shapes

### Shape A: Product-specific workbook

Examples:

- `P4_1-2月份退货分析.xlsx`
- `ZA3退货分析.xlsx`

These usually contain:

- One tagged detail sheet
- One summary sheet

### Shape B: Multi-product workbook

Examples:

- `GR70 LC30 BOX X2 GR40亚马逊客户反馈产品近12月反馈总结.xlsx`

These usually contain:

- Shared source data
- Per-product detail sheets
- Per-product summary sheets

Require the user to name the target product for this shape.

### Shape C: After-sales workbook

Examples:

- `ZD3售后反馈总结--WNW20260207.xlsx`

These usually contain:

- Raw remark fields
- Manual category fields
- One final analysis sheet

## Minimum Column Set

At least one of these text columns must exist:

- `中文翻译`
- `原文`
- `英文评论`
- `买家备注`
- `退货原因`

These fields are optional but valuable:

- Product or SKU
- Store
- Country
- Existing level-1 to level-4 tags
- Source row ID

## If the Workbook Is Ambiguous

Ask the user only after profiling fails to resolve it.

Good reasons to ask:

- Many products are present and the target product is not named
- Several sheets look like valid raw-data candidates
- No usable comment column exists
- The workbook contains only summary data and no row-level comments

## Standard Output Contract

When the skill produces structured output, normalize fields to:

- `record_id`
- `product_name`
- `source_type`
- `source_sheet`
- `source_row`
- `return_time`
- `store`
- `country`
- `raw_comment`
- `translated_comment`
- `cleaned_comment`
- `is_valid_feedback`
- `level_1`
- `level_2`
- `level_3`
- `level_4`
- `severity`
- `is_user_reason`
- `is_quality_risk`
