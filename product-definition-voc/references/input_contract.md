# Input Contract

Use this contract before analyzing any workbook.

## Required

- At least one `.xlsx` workbook
- A focus statement when the workbook mixes multiple products, categories, or scenarios

## Preferred

- A workbook with row-level review text
- Rating, title, and image columns when available
- A note describing what should count as in-scope

## Acceptable Focus Shapes

### Shape A: Single product

Examples:

- `LC40`
- `ZA3`

Use when the workbook is already centered on one product or SKU.

### Shape B: Category or competitor set

Examples:

- `RCA 切换器竞品`
- `桌面音频切换器`

Use when the goal is to learn product-definition signals across similar listings.

### Shape C: Scenario-led focus

Examples:

- `电视补耳机口场景`
- `桌面双机切换场景`

Use when the goal is to learn why users adopt or reject a workflow, not just one product.

## Minimum Column Set

At least one usable review-text field must exist:

- `评论内容中文`
- `中文翻译`
- `评论内容`
- `原文`
- `review`
- `comment`

Helpful optional fields:

- Product or listing title
- Scene or keyword field
- Rating
- Review title
- Date
- Image columns
- Existing tags

## If The Workbook Is Ambiguous

Ask the user only after profiling fails to resolve scope.

Good reasons to ask:

- Several unrelated products remain after focus filtering
- Multiple sheets look like valid raw-review sources
- No usable review-text column exists
- The workbook contains only summaries and no row-level comments

## Standard Output Contract

Normalize review rows toward these fields:

- `record_id`
- `focus_name`
- `asin`
- `product_name`
- `product_image`
- `product_link`
- `scene_name`
- `keyword_source`
- `source_sheet`
- `source_row`
- `rating`
- `review_date`
- `cleaned_comment`
- `观点类型`
- `一级功能`
- `二级功能需求`
- `底层需求`
- `决策信号`
- `情绪极性`
- `情绪强度`
- `场景标签`
- `嘿哈时刻`
