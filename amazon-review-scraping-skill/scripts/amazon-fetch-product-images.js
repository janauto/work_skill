#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');
const { chromium } = require('playwright');

function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith('--')) {
      continue;
    }
    const key = token.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith('--')) {
      args[key] = 'true';
      continue;
    }
    args[key] = next;
    index += 1;
  }
  return args;
}

function imageExtension(imageUrl) {
  try {
    const pathname = new URL(imageUrl).pathname;
    const match = pathname.match(/\.([a-zA-Z0-9]+)$/);
    return match ? `.${match[1].toLowerCase()}` : '.jpg';
  } catch (_error) {
    return '.jpg';
  }
}

function fileExistsForAsin(outputDir, asin) {
  return ['.jpg', '.jpeg', '.png', '.webp'].some((extension) => fs.existsSync(path.join(outputDir, `${asin}${extension}`)));
}

function downloadFile(imageUrl, destination) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(destination);
    https
      .get(
        imageUrl,
        {
          headers: {
            'User-Agent': 'Mozilla/5.0',
          },
        },
        (response) => {
          if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
            file.close();
            fs.unlink(destination, () => {});
            resolve(downloadFile(response.headers.location, destination));
            return;
          }

          if (response.statusCode !== 200) {
            file.close();
            fs.unlink(destination, () => {});
            reject(new Error(`HTTP ${response.statusCode} for ${imageUrl}`));
            return;
          }

          response.pipe(file);
          file.on('finish', () => file.close(resolve));
        }
      )
      .on('error', (error) => {
        file.close();
        fs.unlink(destination, () => {});
        reject(error);
      });
  });
}

async function extractProductImage(page, productUrl) {
  await page.goto(productUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(parseInt(process.env.PRODUCT_IMAGE_WAIT_MS || '1800', 10));
  return page.locator('#landingImage, #imgBlkFront, img[data-old-hires], img.a-dynamic-image').first().evaluate((element) => ({
    imageUrl: element.getAttribute('data-old-hires') || element.currentSrc || element.getAttribute('src') || '',
  }));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.state) {
    throw new Error('Usage: node scripts/amazon-fetch-product-images.js --state ./preflight_state.json [--output-dir ./product_images]');
  }

  const statePath = path.resolve(args.state);
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8'));
  const outputDir = path.resolve(args['output-dir'] || path.join(state.outputDir || path.dirname(statePath), 'candidate_product_images'));
  const concurrency = Math.max(1, parseInt(args.concurrency || process.env.IMAGE_FETCH_CONCURRENCY || '4', 10));
  const candidates = state.candidates || [];

  fs.mkdirSync(outputDir, { recursive: true });

  console.log(
    JSON.stringify(
      {
        stage: 'product_image_fetch_start',
        candidateCount: candidates.length,
        outputDir,
        concurrency,
      },
      null,
      2
    )
  );

  const browser = await chromium.launch({ headless: true });
  const pages = await Promise.all(
    Array.from({ length: concurrency }, async () =>
      browser.newPage({
        viewport: { width: 1280, height: 1400 },
        userAgent:
          process.env.USER_AGENT ||
          'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36',
      })
    )
  );

  let index = 0;
  let savedCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  async function worker(page) {
    while (index < candidates.length) {
      const currentIndex = index;
      index += 1;
      const candidate = candidates[currentIndex];
      const asin = candidate.asin;

      if (!asin || !candidate.productUrl) {
        failedCount += 1;
        continue;
      }

      if (fileExistsForAsin(outputDir, asin)) {
        skippedCount += 1;
        continue;
      }

      try {
        let imageUrl = candidate.imageUrl || '';
        if (!imageUrl) {
          const extracted = await extractProductImage(page, candidate.productUrl);
          imageUrl = extracted.imageUrl;
        }
        if (!imageUrl) {
          throw new Error('No product image found');
        }
        const destination = path.join(outputDir, `${asin}${imageExtension(imageUrl)}`);
        await downloadFile(imageUrl, destination);
        savedCount += 1;
        if (savedCount % 20 === 0) {
          console.log(JSON.stringify({ stage: 'product_image_fetch_progress', savedCount, skippedCount, failedCount }));
        }
      } catch (error) {
        failedCount += 1;
        console.log(JSON.stringify({ stage: 'product_image_fetch_failed', asin, message: String(error.message || error) }));
      }
    }
  }

  try {
    await Promise.all(pages.map((page) => worker(page)));
  } finally {
    await Promise.all(pages.map((page) => page.close().catch(() => {})));
    await browser.close();
  }

  console.log(
    JSON.stringify(
      {
        stage: 'product_image_fetch_done',
        outputDir,
        savedCount,
        skippedCount,
        failedCount,
      },
      null,
      2
    )
  );
}

main().catch((error) => {
  console.error(`Failed: ${error.message}`);
  process.exit(1);
});
