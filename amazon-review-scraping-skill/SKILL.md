---
name: amazon-review-scraping-skill
description: End-to-end Amazon review scraping skill for Amazon product pages. Use when users want to scrape Amazon reviews, review images, logged-in review pages, export bilingual Excel files, or need Playwright anti-bot scraping with the same simple and stealth capabilities provided by playwright-scraper-skill.
---

# Amazon Review Scraping Skill

This skill packages the full Amazon review workflow and keeps the core `playwright-scraper-skill` capabilities inside the same bundle.

Use it when the user wants any of the following:
- Scrape Amazon product reviews from `amazon.sg`, `amazon.com`, or similar marketplaces
- Open a real browser, let the user log in, then continue scraping with the logged-in session
- Download customer review images together with the review text
- Export review data to Excel, with embedded images and Chinese translations
- Use generic Playwright scraping for non-Amazon pages with either simple or stealth mode

## What is included

### 1. Full Playwright scraper modes

These scripts preserve the main capabilities of `playwright-scraper-skill`:
- `scripts/playwright-simple.js`
  Use for ordinary JS-rendered pages without strong anti-bot controls.
- `scripts/playwright-stealth.js`
  Use for Cloudflare, 403, or stronger anti-bot pages.

### 2. Amazon review workflow

- `scripts/amazon-review-login-scrape.js`
  Opens a persistent Playwright browser, waits for manual Amazon login if needed, scrapes multiple review views, deduplicates reviews, and downloads review images.
- `scripts/amazon-reviews-to-excel.py`
  Converts the review JSON into an Excel workbook with:
  - user name
  - rating
  - title
  - review body
  - review time
  - verification flag
  - source views
  - image paths
  - embedded review images
  - Chinese translation columns
- `scripts/amazon-review-workflow.js`
  One-command wrapper that runs the scrape step and then exports Excel.

If you need the field schema, read [references/amazon-review-output.md](references/amazon-review-output.md).

## Install

Run from the skill root:

```bash
cd /Users/linxiansheng/Desktop/工作Skill/work_skill/amazon-review-scraping-skill
npm install
npx playwright install chromium
python3 -m pip install openpyxl pillow requests
```

## Recommended decision rule

1. Non-Amazon page, no obvious anti-bot:
```bash
node scripts/playwright-simple.js "https://example.com"
```

2. Non-Amazon page with Cloudflare, 403, or challenge page:
```bash
HEADLESS=false SAVE_HTML=true node scripts/playwright-stealth.js "https://example.com"
```

3. Amazon review scraping, with login/session reuse, images, and export:
```bash
node scripts/amazon-review-workflow.js "https://www.amazon.sg/dp/B0CYHCGRXM" "./output"
```

## Amazon workflow details

### Fast path

Run the all-in-one workflow:

```bash
node scripts/amazon-review-workflow.js "<amazon_product_url>" "<output_dir>"
```

Outputs:
- JSON review file
- review-image folder
- Excel workbook

### Step-by-step mode

Scrape first:

```bash
node scripts/amazon-review-login-scrape.js "<amazon_product_url>" "./output/amazon_reviews.json"
```

Then export Excel:

```bash
python3 scripts/amazon-reviews-to-excel.py "./output/amazon_reviews.json" "./output/amazon_reviews.xlsx"
```

## Amazon workflow behavior

- Opens a real Playwright browser with a persistent session directory
- Reuses an existing session if already logged in
- Waits for manual login if Amazon redirects to sign-in
- Scrapes these review views:
  - `top_reviews`
  - `most_recent`
  - `positive_reviews`
  - `critical_reviews`
- Repeatedly clicks `Show more reviews` when that is the true pagination path
- Deduplicates reviews by `reviewId`
- Downloads customer review images into a sibling `_media` directory
- Writes bilingual Excel columns by default

## Useful environment variables

### Generic Playwright scripts

- `HEADLESS=false`
- `WAIT_TIME=10000`
- `SCREENSHOT_PATH=/tmp/page.png`
- `SAVE_HTML=true`
- `USER_AGENT="Mozilla/5.0 ..."`

### Amazon review scripts

- `HEADLESS=false`
  Default recommendation, because login may require a visible browser
- `LOGIN_TIMEOUT_MS=900000`
- `WAIT_AFTER_LOAD_MS=3000`
- `DOWNLOAD_IMAGES=false`
- `SESSION_ROOT=./.sessions`
- `SESSION_DIR=/custom/session/path`
- `LOCALE=zh-HK`
- `USER_AGENT="Mozilla/5.0 ..."`
- `EXPORT_EXCEL=false`
  Skip Excel in the wrapper script
- `NO_TRANSLATE=true`
  Export Excel without Chinese translations
- `PYTHON_BIN=python3`

## Operational guidance

### If login is required

Use `HEADLESS=false`. The script will open the review page and poll until Amazon is no longer on the sign-in flow.

### If the review page does not expose normal `pageNumber` paging

Do not assume URL paging works. The Amazon review workflow is designed to click the actual `Show more reviews` control and follow the page state exposed by the live DOM.

### If ratings count is larger than written review count

Treat `global ratings` and written reviews separately. The skill only counts written reviews actually exposed by the logged-in review UI.

### If you only need a quick site scrape

Do not use the Amazon workflow. Use `playwright-simple.js` or `playwright-stealth.js` directly.

## Examples of when this skill should trigger

- “抓取这个 Amazon 商品的全部评论，并导出 Excel”
- “打开 Amazon 页面让我登录，登录后继续抓评论和评论图片”
- “把 Amazon 评论做成带中文翻译的工作簿”
- “这个站点有 Cloudflare，用 Playwright stealth 抓一下”
- “先试简单版 Playwright，不行再切 stealth”
