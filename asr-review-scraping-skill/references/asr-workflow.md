# ASR Workflow

## Recommended Dataset Layout

```text
<dataset-root>/
├── curated_threads.txt
├── raw_threads/
├── thread_index.json
├── thread_summary.md
├── downloaded_images/
├── preview_images/
├── translation_cache.json
└── ASR_切换相关用户内容_打标准备.xlsx
```

## End-to-End Command

```bash
python3 scripts/run_asr_pipeline.py --dataset-root runs/default
```

## Partial Runs

Fetch only:

```bash
python3 scripts/fetch_asr_threads.py --dataset-root runs/default
```

Workbook only:

```bash
python3 scripts/build_asr_workbook.py --dataset-root runs/default
```

## Workbook Columns

- `post_uid`
- `primary_theme`
- `secondary_themes`
- `topic_labels`
- `product_name`
- `thread_title`
- `thread_url`
- `post_no`
- `author`
- `author_role`
- `post_date`
- `is_thread_starter`
- `quote_text`
- `post_text`
- `中文翻译`
- `raw_text`
- `image_count`
- `image_urls`
- `评论图片预览`
- `local_file`
- `一级标签`
- `二级标签`
- `三级标签`
- `情绪倾向`
- `备注`

## Image Handling

- External images are downloaded directly.
- ASR attachment images are extracted from the local Chrome cache.
- If a cache miss occurs, the builder briefly opens an off-screen Chrome window to force a cache fill, then retries extraction.
- Preview images are built from original files, then scaled only for Excel display.

## Known Constraints

- Chrome cache extraction is macOS + local Chrome specific in the current implementation.
- If Chrome cannot open a given ASR image URL in the browser, the builder will not recover that image.
- Thread body fetching relies on `r.jina.ai`, so formatting is only as accurate as the mirror output.
