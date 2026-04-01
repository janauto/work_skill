#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const productUrl = process.argv[2];

if (!productUrl) {
  console.error('Usage: node scripts/amazon-review-login-scrape.js <amazon_product_url> [output_json_path]');
  process.exit(1);
}

const headless = process.env.HEADLESS === 'true';
const waitAfterLoadMs = parseInt(process.env.WAIT_AFTER_LOAD_MS || '3000', 10);
const loginTimeoutMs = parseInt(process.env.LOGIN_TIMEOUT_MS || `${15 * 60 * 1000}`, 10);
const downloadImages = process.env.DOWNLOAD_IMAGES !== 'false';
const locale = process.env.LOCALE || 'zh-HK';
const userAgent =
  process.env.USER_AGENT ||
  'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1';

function extractAsin(url) {
  const match = url.match(/\/(?:dp|product-reviews)\/([A-Z0-9]{10})/i);
  return match ? match[1].toUpperCase() : null;
}

function sanitizeSegment(value) {
  return String(value).replace(/[^a-zA-Z0-9._-]+/g, '_');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getFileExtension(url) {
  try {
    const pathname = new URL(url).pathname;
    const match = pathname.match(/\.([a-zA-Z0-9]+)$/);
    return match ? `.${match[1].toLowerCase()}` : '.jpg';
  } catch (_error) {
    return '.jpg';
  }
}

async function downloadReviewImages(context, reviews, mediaDir) {
  fs.mkdirSync(mediaDir, { recursive: true });

  let downloadedCount = 0;

  for (const review of reviews) {
    if (!Array.isArray(review.images) || review.images.length === 0) {
      continue;
    }

    for (let index = 0; index < review.images.length; index += 1) {
      const image = review.images[index];
      const imageUrl = image.originalUrl || image.thumbnailUrl;

      if (!imageUrl) {
        continue;
      }

      const extension = getFileExtension(imageUrl);
      const filePath = path.join(mediaDir, `${review.reviewId}_${String(index + 1).padStart(2, '0')}${extension}`);

      if (!fs.existsSync(filePath)) {
        const response = await context.request.get(imageUrl, {
          failOnStatusCode: false,
        });

        if (!response.ok()) {
          console.warn(`Image download failed for ${review.reviewId}: ${imageUrl} (${response.status()})`);
          continue;
        }

        fs.writeFileSync(filePath, await response.body());
        downloadedCount += 1;
      }

      image.localPath = filePath;
    }
  }

  return downloadedCount;
}

async function waitForManualLogin(page, timeoutMs) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    let state;

    try {
      state = await page.evaluate(() => {
        const href = window.location.href;
        const text = document.body?.innerText || '';
        const title = document.title;

        return {
          href,
          title,
          needsLogin:
            href.includes('/ap/signin') ||
            /Sign in|create account|Enter mobile number or email|输入手机号码或邮箱|登入|登录/i.test(text),
        };
      });
    } catch (error) {
      if (
        /Execution context was destroyed|Cannot find context|Target page, context or browser has been closed/i.test(
          String(error.message || error)
        )
      ) {
        await sleep(2000);
        continue;
      }
      throw error;
    }

    console.log(
      JSON.stringify({
        stage: 'waiting_for_login',
        title: state.title,
        url: state.href,
        needsLogin: state.needsLogin,
      })
    );

    if (!state.needsLogin) {
      return;
    }

    await sleep(5000);
  }

  throw new Error('Timed out waiting for manual Amazon login.');
}

async function collectReviewsFromCurrentPage(page) {
  return page.evaluate(() => {
    const nodes = Array.from(
      document.querySelectorAll(
        '[data-hook="mobley-review-content"], [data-hook="review"], [id^="customer_review-"], [id^="R"][data-hook]'
      )
    );

    const reviews = [];
    const seen = new Set();

    for (const node of nodes) {
      const reviewRoot =
        node.matches('[data-hook="mobley-review-content"], [data-hook="review"]')
          ? node
          : node.closest('[data-hook="mobley-review-content"], [data-hook="review"]') || node;

      const reviewId =
        reviewRoot.id ||
        reviewRoot.getAttribute('id') ||
        reviewRoot.querySelector('[id^="R"]')?.id ||
        null;

      if (!reviewId || seen.has(reviewId)) {
        continue;
      }
      seen.add(reviewId);

      const ratingText =
        reviewRoot.querySelector('[data-hook="review-star-rating"] .a-icon-alt')?.textContent?.trim() ||
        reviewRoot.querySelector('.review-rating .a-icon-alt')?.textContent?.trim() ||
        '';

      const ratingMatch = ratingText.match(/([0-9.]+)/);

      reviews.push({
        reviewId,
        author:
          reviewRoot.querySelector('.a-profile-name')?.textContent?.trim() ||
          reviewRoot.querySelector('[data-hook="genome-widget"]')?.textContent?.trim() ||
          '',
        title:
          reviewRoot.querySelector('[data-hook="review-title"]')?.textContent?.replace(/\s+/g, ' ').trim() ||
          '',
        body:
          reviewRoot.querySelector('[data-hook="review-body"]')?.textContent?.replace(/\s+/g, ' ').trim() ||
          reviewRoot.querySelector('.cr-full-content')?.textContent?.replace(/\s+/g, ' ').trim() ||
          '',
        ratingText,
        rating: ratingMatch ? parseFloat(ratingMatch[1]) : null,
        dateText:
          reviewRoot.querySelector('[data-hook="review-date"]')?.textContent?.replace(/\s+/g, ' ').trim() ||
          '',
        verifiedPurchase: Boolean(
          reviewRoot.querySelector('[data-hook="avp-badge"], [data-hook="msrp-avp-badge-linkless"]')
        ),
        images: Array.from(reviewRoot.querySelectorAll('button.review-image-thumbnail')).map((button) => {
          const originalUrl = button.getAttribute('data-image-source') || '';
          const style = button.getAttribute('style') || '';
          const thumbnailMatch = style.match(/url\((.+?)\)/);

          return {
            imageIndex: parseInt(button.getAttribute('data-image-index') || '0', 10),
            originalUrl,
            thumbnailUrl: thumbnailMatch ? thumbnailMatch[1] : '',
          };
        }),
      });
    }

    const showMoreButton = document.querySelector('[data-hook="show-more-button"]');

    return {
      title: document.title,
      url: window.location.href,
      reviews,
      showMoreButton: showMoreButton
        ? {
            text: showMoreButton.textContent?.trim() || '',
            state: showMoreButton.getAttribute('data-reviews-state-param') || '',
          }
        : null,
    };
  });
}

async function collectReviewsForView(page, label, url) {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(waitAfterLoadMs);

  let lastCount = -1;
  let clickCount = 0;

  for (let i = 1; i <= 40; i += 1) {
    const snapshot = await collectReviewsFromCurrentPage(page);

    console.log(
      JSON.stringify({
        stage: 'scraping_view',
        label,
        step: i,
        url: snapshot.url,
        reviewCount: snapshot.reviews.length,
        hasShowMore: Boolean(snapshot.showMoreButton),
      })
    );

    if (!snapshot.showMoreButton || snapshot.reviews.length === lastCount) {
      return {
        label,
        url: snapshot.url,
        clickCount,
        reviews: snapshot.reviews,
      };
    }

    lastCount = snapshot.reviews.length;
    clickCount += 1;

    await page.locator('[data-hook="show-more-button"]').first().click();
    await page.waitForTimeout(1800);
  }

  const finalSnapshot = await collectReviewsFromCurrentPage(page);
  return {
    label,
    url: finalSnapshot.url,
    clickCount,
    reviews: finalSnapshot.reviews,
  };
}

async function main() {
  const parsedProductUrl = new URL(productUrl);
  const host = parsedProductUrl.host;
  const asin = extractAsin(productUrl);

  if (!asin) {
    throw new Error(`Could not extract ASIN from URL: ${productUrl}`);
  }

  const outputPath =
    process.argv[3] ||
    path.join(process.cwd(), `amazon_reviews_${sanitizeSegment(host)}_${asin}.json`);
  const mediaDir = path.join(path.dirname(outputPath), `${path.basename(outputPath, path.extname(outputPath))}_media`);
  const sessionRoot = process.env.SESSION_ROOT || path.join(process.cwd(), '.sessions');
  const sessionDir =
    process.env.SESSION_DIR || path.join(sessionRoot, `amazon-${sanitizeSegment(host)}-${sanitizeSegment(asin)}`);

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.mkdirSync(sessionDir, { recursive: true });

  const reviewsUrl = `https://${host}/product-reviews/${asin}/ref=cm_cr_dp_mb_show_all_btm?ie=UTF8&reviewerType=all_reviews&pageNumber=1`;

  console.log(JSON.stringify({ stage: 'launching_browser', sessionDir, reviewsUrl, outputPath }));

  const context = await chromium.launchPersistentContext(sessionDir, {
    headless,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-blink-features=AutomationControlled',
      '--disable-features=IsolateOrigins,site-per-process',
    ],
    userAgent,
    locale,
    viewport: { width: 375, height: 812 },
    extraHTTPHeaders: {
      'Accept-Language': `${locale},en;q=0.8`,
      Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    },
  });

  await context.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', {
      get: () => false,
    });
    window.chrome = { runtime: {} };
  });

  const page = context.pages()[0] || (await context.newPage());

  try {
    await page.goto(reviewsUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    console.log(
      JSON.stringify({
        stage: 'browser_opened',
        message: 'Use the opened browser window to log in if Amazon asks for authentication.',
        headless,
      })
    );

    await waitForManualLogin(page, loginTimeoutMs);

    const views = [
      {
        label: 'top_reviews',
        url: reviewsUrl,
      },
      {
        label: 'most_recent',
        url: `${reviewsUrl}&sortBy=recent`,
      },
      {
        label: 'positive_reviews',
        url: `https://${host}/product-reviews/${asin}/ref=cm_cr_positive?ie=UTF8&filterByStar=positive&reviewerType=all_reviews#reviews-filter-bar`,
      },
      {
        label: 'critical_reviews',
        url: `https://${host}/product-reviews/${asin}/ref=cm_cr_critical?ie=UTF8&filterByStar=critical&reviewerType=all_reviews#reviews-filter-bar`,
      },
    ];

    const collected = new Map();
    const viewResults = [];

    for (const view of views) {
      const result = await collectReviewsForView(page, view.label, view.url);
      viewResults.push({
        label: result.label,
        url: result.url,
        clickCount: result.clickCount,
        reviewCount: result.reviews.length,
      });

      for (const review of result.reviews) {
        if (!review.reviewId) {
          continue;
        }

        const existing = collected.get(review.reviewId);
        if (existing) {
          if (!existing.sourceViews.includes(result.label)) {
            existing.sourceViews.push(result.label);
          }
          if ((!existing.body || existing.body.length < review.body.length) && review.body) {
            existing.body = review.body;
          }
          if ((!existing.title || existing.title.length < review.title.length) && review.title) {
            existing.title = review.title;
          }
          if ((!existing.author || existing.author.length < review.author.length) && review.author) {
            existing.author = review.author;
          }
          if ((!existing.dateText || existing.dateText.length < review.dateText.length) && review.dateText) {
            existing.dateText = review.dateText;
          }
          if (Array.isArray(review.images) && review.images.length > (existing.images || []).length) {
            existing.images = review.images;
          }
          continue;
        }

        collected.set(review.reviewId, {
          ...review,
          sourceViews: [result.label],
        });
      }
    }

    const output = {
      asin,
      host,
      productUrl,
      reviewsUrl,
      fetchedAt: new Date().toISOString(),
      loggedInFlow: true,
      sessionDir,
      viewResults,
      mediaDir,
      imageReviewCount: 0,
      downloadedImageCount: 0,
      reviewCount: collected.size,
      reviews: Array.from(collected.values()).sort((a, b) => {
        if ((b.rating || 0) !== (a.rating || 0)) {
          return (b.rating || 0) - (a.rating || 0);
        }
        return (a.reviewId || '').localeCompare(b.reviewId || '');
      }),
    };

    output.imageReviewCount = output.reviews.filter((review) => review.images?.length).length;

    if (downloadImages) {
      output.downloadedImageCount = await downloadReviewImages(context, output.reviews, mediaDir);
    }

    fs.writeFileSync(outputPath, `${JSON.stringify(output, null, 2)}\n`);

    console.log(
      JSON.stringify(
        {
          stage: 'done',
          outputPath,
          reviewCount: collected.size,
          imageReviewCount: output.imageReviewCount,
          downloadedImageCount: output.downloadedImageCount,
          mediaDir,
        },
        null,
        2
      )
    );
  } finally {
    await context.close();
  }
}

main().catch((error) => {
  console.error(`Failed: ${error.message}`);
  process.exit(1);
});
