#!/usr/bin/env node

const path = require('path');
const { spawnSync } = require('child_process');

const productUrl = process.argv[2];
const outputDir = process.argv[3] ? path.resolve(process.argv[3]) : process.cwd();

if (!productUrl) {
  console.error('Usage: node scripts/amazon-review-workflow.js <amazon_product_url> [output_dir]');
  process.exit(1);
}

function extractAsin(url) {
  const match = url.match(/\/(?:dp|product-reviews)\/([A-Z0-9]{10})/i);
  return match ? match[1].toUpperCase() : null;
}

function sanitizeSegment(value) {
  return String(value).replace(/[^a-zA-Z0-9._-]+/g, '_');
}

const host = new URL(productUrl).host;
const asin = extractAsin(productUrl);

if (!asin) {
  console.error(`Could not extract ASIN from URL: ${productUrl}`);
  process.exit(1);
}

const jsonPath = path.join(outputDir, `amazon_reviews_${sanitizeSegment(host)}_${asin}.json`);
const xlsxPath = path.join(outputDir, `amazon_reviews_${sanitizeSegment(host)}_${asin}.xlsx`);
const scrapeScript = path.join(__dirname, 'amazon-review-login-scrape.js');
const excelScript = path.join(__dirname, 'amazon-reviews-to-excel.py');

const scrape = spawnSync(process.execPath, [scrapeScript, productUrl, jsonPath], {
  stdio: 'inherit',
  env: process.env,
});

if (scrape.status !== 0) {
  process.exit(scrape.status || 1);
}

if (process.env.EXPORT_EXCEL === 'false') {
  console.log(JSON.stringify({ stage: 'done', jsonPath, excelExported: false }, null, 2));
  process.exit(0);
}

const pythonBin = process.env.PYTHON_BIN || 'python3';
const excelArgs = [excelScript, jsonPath, xlsxPath];
if (process.env.NO_TRANSLATE === 'true') {
  excelArgs.push('--no-translate');
}

const excel = spawnSync(pythonBin, excelArgs, {
  stdio: 'inherit',
  env: process.env,
});

if (excel.status !== 0) {
  process.exit(excel.status || 1);
}

console.log(JSON.stringify({ stage: 'done', jsonPath, xlsxPath, excelExported: true }, null, 2));
