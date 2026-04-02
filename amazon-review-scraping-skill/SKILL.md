---
name: amazon-review-scraping-skill
description: Preflight-first Amazon review scraping skill for Amazon product pages. Use when users want to search Amazon by 2-3 keywords, generate numbered product-category scope drafts, estimate product and review volume, exclude scenarios before execution, then scrape product reviews, review images, and export bilingual Excel files. Includes the same simple and stealth Playwright capabilities provided by playwright-scraper-skill.
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
- `scripts/amazon-product-discovery.js`
  Searches Amazon result pages by keywords and extracts candidate products with review counts.
- `scripts/amazon-preflight-workflow.js`
  Default entrypoint. Generates numbered scenario drafts, estimates scope, accepts exclusions such as `不搜索：1、8、11`, and only executes after `开始执行`.
- `scripts/amazon-competitor-execute.js`
  Executes the confirmed batch scrape and writes a run manifest plus multi-product workbook.
- `scripts/amazon-competitor-to-excel.py`
  Exports a multi-product workbook with `场景与候选商品`、`评论明细`、`抓取汇总`.

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

3. Amazon review scraping, with preflight confirmation, login/session reuse, images, and export:
```bash
node scripts/amazon-preflight-workflow.js \
  --marketplace amazon.sg \
  --keywords "rca switcher,3.5mm switcher,audio selector" \
  --category Electronics \
  --price-min 10 \
  --price-max 60 \
  --min-rating 4.0 \
  --top-n 5 \
  --output-dir "./output"
```

## Amazon workflow details

### Fast path

Run the preflight workflow:

```bash
node scripts/amazon-preflight-workflow.js \
  --marketplace amazon.sg \
  --keywords "keyword1,keyword2,keyword3" \
  --output-dir "./output"
```

Outputs:
- `preflight_state.json`
- 编号场景清单和估算结果
- 用户回复 `开始执行` 后再生成：
  - run manifest
  - per-product review JSON files
  - review-image folders
  - multi-product Excel workbook

### Step-by-step mode

1. Generate preflight state:

```bash
node scripts/amazon-preflight-workflow.js \
  --marketplace amazon.sg \
  --keywords "rca switcher,3.5mm switcher" \
  --output-dir "./output"
```

2. Exclude scenarios if needed:

```bash
node scripts/amazon-preflight-workflow.js \
  --state "./output/preflight_state.json" \
  --reply "不搜索：1、8、11"
```

3. Start execution only after confirmation:

```bash
node scripts/amazon-preflight-workflow.js \
  --state "./output/preflight_state.json" \
  --reply "开始执行"
```

## Amazon workflow behavior

- Uses a preflight-first flow
- Builds numbered category-style scenario drafts from search results
- Supports exclusions such as `不搜索：1、8、11`
- Prints:
  - kept scenario count
  - candidate product count
  - estimated review volume
  - Top N execution preview
- Starts the real scrape only when the reply is exactly `开始执行`
- Opens a real Playwright browser with a persistent session directory
- Reuses an existing session if already logged in
- Waits for manual login if Amazon redirects to sign-in
- Scrapes these review views for selected products:
  - `top_reviews`
  - `most_recent`
  - `positive_reviews`
  - `critical_reviews`
- Repeatedly clicks `Show more reviews` when that is the true pagination path
- Deduplicates reviews by `reviewId`
- Downloads customer review images into a sibling `_media` directory
- Writes a multi-product workbook by default

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
- `SEARCH_WAIT_MS=2000`
- `pages=2`

## Operational guidance

### If login is required

Use `HEADLESS=false`. The script will open the review page and poll until Amazon is no longer on the sign-in flow.

### If the user wants to narrow the search scope

Let the preflight workflow generate the numbered scenario draft first. Then remove unwanted categories by replying with `不搜索：编号1、编号2`.

### If the review page does not expose normal `pageNumber` paging

Do not assume URL paging works. The review workflow clicks the actual `Show more reviews` control and follows the live page state.

### If ratings count is larger than written review count

Treat search-result `rating_count` as an estimate only. The final written review count comes from the review page itself.

### If you only need a quick site scrape

Do not use the Amazon workflow. Use `playwright-simple.js` or `playwright-stealth.js` directly.

## Examples of when this skill should trigger

- “先给我 2-3 个关键词的 Amazon 搜索范围初稿，确认后再抓”
- “不搜索 1、8、11，剩下的竞品继续保留”
- “用户确认前不要执行，只有回复开始执行才开始抓”
- “抓 Amazon 多个竞品的评论和评论图片，并汇总成 Excel”
- “把 Amazon 评论做成带中文翻译的多商品工作簿”
- “这个站点有 Cloudflare，用 Playwright stealth 抓一下”
- “先试简单版 Playwright，不行再切 stealth”
