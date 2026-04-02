---
name: asr-review-scraping-skill
description: Scrape Audio Science Review forum threads, extract comment images from the local Chrome browser cache, translate posts, and build a single-sheet Excel tagging workbook. Also includes the generic Playwright simple and stealth scraping workflows for regular dynamic sites and Cloudflare-protected sites.
---

# ASR Review Scraping Skill

Use this skill when the user wants to:
- scrape Audio Science Review forum threads or attachments
- archive ASR posts about products, switchers, DACs, preamps, AVR workflows, or similar topics
- collect comment images and embed previews into Excel
- build a tagging workbook with topic labels, translation columns, and post metadata
- run generic Playwright scraping against JS-heavy or Cloudflare-protected sites

## Quick Routing

- Regular dynamic pages: `node scripts/playwright-simple.js <URL>`
- Cloudflare or anti-bot pages: `node scripts/playwright-stealth.js <URL>`
- Full ASR pipeline: `python3 scripts/run_asr_pipeline.py --dataset-root <DIR>`

## Prerequisites

- Install Node deps in this skill folder:
```bash
npm install
npx playwright install chromium
```
- Install Python deps:
```bash
python3 -m pip install -r requirements.txt
```
- For ASR attachment images, Chrome on the local machine should already be able to open the image URLs normally. The workbook builder extracts original image responses from the local Chrome cache instead of using screenshots.

## ASR Workflow

1. Create or choose a dataset folder. Default is `runs/default`.
2. Put ASR thread URLs into `curated_threads.txt`.
3. Run:
```bash
python3 scripts/run_asr_pipeline.py --dataset-root runs/default
```
4. Outputs are written into that dataset folder:
- `raw_threads/`
- `thread_index.json`
- `thread_summary.md`
- `downloaded_images/`
- `preview_images/`
- `translation_cache.json`
- `ASR_切换相关用户内容_打标准备.xlsx`

## Targeted Commands

- Fetch threads only:
```bash
python3 scripts/fetch_asr_threads.py --dataset-root runs/default
```
- Rebuild workbook only:
```bash
python3 scripts/build_asr_workbook.py --dataset-root runs/default
```
- Use a custom URL file:
```bash
python3 scripts/run_asr_pipeline.py --dataset-root runs/project-a --urls-file /abs/path/curated_threads.txt
```

## Translation

If any of `ZHIPUAI_API_KEY`, `ZHIPU_API_KEY`, or `BIGMODEL_API_KEY` is present, the workbook builder fills `中文翻译` from the Zhipu API. Without a key, cached translations are reused and uncached rows stay blank.

## Notes

- ASR thread pages are fetched through `r.jina.ai` text mirrors.
- ASR attachment images are recovered from the local Chrome cache, so this workflow is suited to the desktop app environment.
- The workbook is a single-sheet tagging table with topic labels, image previews, and tagging columns.

## References

- Detailed workflow: `references/asr-workflow.md`
- URL list template: `references/curated_threads_template.txt`
