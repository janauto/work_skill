# Cleaning Rules

Apply only standard cleaning that another analyst can audit quickly.

## 1. Product Filtering

- Keep only rows tied to the target product.
- Prefer explicit product columns over sheet-name inference.
- Use sheet-name inference only when the workbook is clearly split by product.

## 2. Text Selection

Choose analysis text in this order:

1. `中文翻译`
2. `英文评论`
3. `原文`
4. `买家备注`
5. `退货原因`

Write the chosen value to `cleaned_comment`.

## 3. Invalid Feedback

Mark rows invalid or drop them when the chosen text is (unless `--keep-blank` is specified to keep them):

- Empty
- `NA`
- `N/A`
- `同上`
- `same as above`
- Pure punctuation
- A formatting artifact with no problem statement

If a row has no comment text but already has strong manual labels, keep it only when the user explicitly wants historical labeled data preserved.

## 4. Duplicate Handling

Normalize text before dedupe:

- Lowercase Latin text
- Remove repeated whitespace
- Remove punctuation and separators
- Keep Chinese characters and letters/digits

Treat exact normalized duplicates within the same target product as one record.

Do not merge semantically similar but textually different comments in v1.

## 5. Lineage Preservation

Always preserve:

- Source workbook path if available in the script run context
- Sheet name
- Source row number
- Original text fields

## 6. Existing Labels

- Keep existing manual labels when present.
- Do not overwrite historical labels during cleaning.
- Normalize them into the canonical field names:
  `level_1`, `level_2`, `level_3`, `level_4`

## 7. Conservative Defaults

When information is incomplete:

- Use `未提供` for missing metadata such as store or country
- Use `NA` or `待人工判定` for uncertain labels
- Keep the record if the comment text is still useful
