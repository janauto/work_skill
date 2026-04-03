# Cleaning Rules

Apply only standard cleaning that another analyst can audit quickly.

## 1. Focus Filtering

- Keep only rows tied to the chosen focus
- Prefer explicit product, scene, keyword, or listing-title matches
- Use workbook-name or sheet-name inference only as a weak fallback

## 2. Text Selection

Choose analysis text in this order:

1. `评论标题中文 + 评论内容中文`
2. `中文翻译`
3. `评论标题 + 评论内容`
4. `原文`

Write the result to `cleaned_comment`.

## 3. Drop Invalid Rows

Drop rows when the chosen text is:

- Empty
- `NA`
- `N/A`
- Pure emoji
- Pure punctuation
- Formatting artifacts
- Ultra-short generic filler such as only `good`, `nice`, `ok`, `不错`, `很好`

## 4. Drop Off-Topic Rows

Drop rows when they are only about:

- Shipping speed
- Delivery timing
- Seller attitude
- Customer-service response
- Coupon spam
- Cashback or review-for-reward signals
- Obvious advertisements, links, or contact invitations

Do not drop a row just because it mentions packaging or logistics if it also contains a real product opinion.

## 5. Duplicate Handling

Normalize before dedupe:

- Lowercase Latin text
- Remove repeated whitespace
- Remove punctuation and separators
- Keep Chinese, letters, and digits

Treat exact normalized duplicates within the same focus as one row.

## 6. Preserve Evidence

Always preserve:

- Source workbook path
- Sheet name
- Source row number
- Rating
- Raw and translated text
- Image refs if they exist

## 7. Conservative Defaults

When information is incomplete:

- Use `未提供` for missing metadata
- Use `待人工复核` for uncertain labels
- Keep the row if it still contains product-definition signal
