#!/usr/bin/env node

const DEFAULT_USER_AGENT =
  process.env.USER_AGENT ||
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36';

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

function normalizeWhitespace(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function tokenizeCategory(category) {
  if (!category) {
    return [];
  }
  return Array.from(
    new Set(
      normalizeWhitespace(category)
        .split(/[\s,，/|_-]+/)
        .map((token) => token.trim().toLowerCase())
        .filter((token) => token.length >= 2)
    )
  );
}

function keywordToScenarioLabel(keyword) {
  const source = normalizeWhitespace(keyword);
  if (!source) {
    return '其他音频设备';
  }
  if (/[\u4e00-\u9fff]/.test(source)) {
    return source;
  }

  let label = source.toLowerCase();
  const replacements = [
    [/rca/g, 'RCA'],
    [/3\.?5\s*mm|aux/g, '3.5mm'],
    [/\baudio\b/g, '音频'],
    [/\bstereo\b/g, '立体声'],
    [/\bspeaker\b/g, '扬声器'],
    [/\bheadphone\b/g, '耳机'],
    [/\bswitcher\b|\bswitch\b/g, '切换器'],
    [/\bselector\b|\bselect\b/g, '选择器'],
    [/\bmixer\b/g, '混音器'],
    [/\bsplitter\b/g, '分配器'],
    [/\bamp\b|\bamplifier\b/g, '功放'],
    [/\bline\b/g, '线路'],
  ];

  for (const [pattern, replacement] of replacements) {
    label = label.replace(pattern, replacement);
  }

  label = label
    .replace(/\band\b/g, ' ')
    .replace(/[()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return label || source;
}

function inferScenarioLabel(title, matchedKeywords) {
  const haystack = normalizeWhitespace(title).toLowerCase();
  const mappings = [
    { pattern: /\brca\b.*\b(switch|selector|switcher)\b|\b(switch|selector|switcher)\b.*\brca\b/, label: 'RCA切换器' },
    { pattern: /\b(3\.5 ?mm|aux)\b.*\b(switch|selector|switcher)\b|\b(switch|selector|switcher)\b.*\b(3\.5 ?mm|aux)\b/, label: '3.5mm音频切换器' },
    { pattern: /\bspeaker\b.*\b(switch|selector|switcher)\b|\b(switch|selector|switcher)\b.*\bspeaker\b/, label: '扬声器切换器' },
    { pattern: /\bheadphone\b.*\b(switch|selector|switcher)\b|\b(switch|selector|switcher)\b.*\bheadphone\b/, label: '耳机切换器' },
    { pattern: /\b(audio|stereo)\b.*\b(selector|switch|switcher)\b/, label: '音频选择器' },
    { pattern: /\b(line|source)\b.*\b(selector|switch|switcher)\b/, label: '线路切换器' },
    { pattern: /\b(splitter)\b/, label: '音频分配器' },
    { pattern: /\b(mixer)\b/, label: '音频混音器' },
    { pattern: /\bamp|amplifier\b.*\b(selector|switch|switcher)\b/, label: '功放切换器' },
  ];

  for (const mapping of mappings) {
    if (mapping.pattern.test(haystack)) {
      return mapping.label;
    }
  }

  return keywordToScenarioLabel((matchedKeywords || [])[0] || title);
}

function parsePrice(priceText) {
  const normalized = normalizeWhitespace(priceText);
  if (!normalized) {
    return null;
  }
  const match = normalized.replace(/,/g, '').match(/([0-9]+(?:\.[0-9]+)?)/);
  return match ? parseFloat(match[1]) : null;
}

function parseRating(ratingText) {
  const match = normalizeWhitespace(ratingText).match(/([0-9.]+)/);
  return match ? parseFloat(match[1]) : null;
}

function parseRatingCount(ratingCountText) {
  const match = normalizeWhitespace(ratingCountText).replace(/,/g, '').match(/([0-9]+)/);
  return match ? parseInt(match[1], 10) : 0;
}

function buildSearchUrl(marketplace, keyword, category, pageNumber) {
  const query = [normalizeWhitespace(keyword), normalizeWhitespace(category)].filter(Boolean).join(' ');
  const url = new URL(`https://${marketplace}/s`);
  url.searchParams.set('k', query);
  if (pageNumber > 1) {
    url.searchParams.set('page', String(pageNumber));
  }
  return url.toString();
}

function categoryFilterPasses(title, category) {
  const tokens = tokenizeCategory(category);
  if (tokens.length === 0) {
    return true;
  }
  const normalizedTitle = normalizeWhitespace(title).toLowerCase();
  return tokens.some((token) => normalizedTitle.includes(token));
}

function candidatePasses(candidate, options) {
  if (candidate.ratingAverage !== null && options.minRating !== null && candidate.ratingAverage < options.minRating) {
    return false;
  }
  if (candidate.price !== null && options.priceMin !== null && candidate.price < options.priceMin) {
    return false;
  }
  if (candidate.price !== null && options.priceMax !== null && candidate.price > options.priceMax) {
    return false;
  }
  if (!categoryFilterPasses(candidate.title, options.category)) {
    return false;
  }
  return true;
}

function compareCandidates(left, right) {
  if ((right.ratingCount || 0) !== (left.ratingCount || 0)) {
    return (right.ratingCount || 0) - (left.ratingCount || 0);
  }
  if ((right.ratingAverage || 0) !== (left.ratingAverage || 0)) {
    return (right.ratingAverage || 0) - (left.ratingAverage || 0);
  }
  return String(left.title || '').localeCompare(String(right.title || ''));
}

async function createDiscoveryContext(headless, locale) {
  const { chromium } = require('playwright');
  const browser = await chromium.launch({
    headless,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-blink-features=AutomationControlled',
      '--disable-features=IsolateOrigins,site-per-process',
    ],
  });

  const context = await browser.newContext({
    userAgent: DEFAULT_USER_AGENT,
    locale,
    viewport: { width: 1440, height: 1024 },
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

  return { browser, context };
}

async function extractSearchCandidates(page, marketplace, keyword, category, pageNumber) {
  const searchUrl = buildSearchUrl(marketplace, keyword, category, pageNumber);
  await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(parseInt(process.env.SEARCH_WAIT_MS || '2000', 10));

  const rawCandidates = await page.evaluate(() => {
    const cards = Array.from(document.querySelectorAll('[data-component-type="s-search-result"][data-asin]'));
    return cards.map((card) => {
      const titleNode =
        card.querySelector('h2 a span') ||
        card.querySelector('a.a-link-normal.s-no-outline span') ||
        card.querySelector('h2 span');
      const linkNode = card.querySelector('h2 a') || card.querySelector('a.a-link-normal.s-no-outline');
      const ratingNode = card.querySelector('.a-icon-alt');
      const ratingCountText = Array.from(
        card.querySelectorAll('a[href*="#customerReviews"], a[href*="customerReviews"]')
      )
        .map((node) => node.textContent || '')
        .join(' ')
        .trim();
      const priceNode = card.querySelector('.a-price .a-offscreen');
      const brandNode = card.querySelector('h5.s-line-clamp-1 span') || card.querySelector('.s-line-clamp-1 .a-size-base-plus');
      const imageNode = card.querySelector('img.s-image');

      return {
        asin: card.getAttribute('data-asin') || '',
        title: titleNode ? titleNode.textContent || '' : '',
        productUrl: linkNode ? linkNode.href || '' : '',
        ratingText: ratingNode ? ratingNode.textContent || '' : '',
        ratingCountText,
        priceText: priceNode ? priceNode.textContent || '' : '',
        brand: brandNode ? brandNode.textContent || '' : '',
        imageUrl: imageNode ? imageNode.src || '' : '',
      };
    });
  });

  return rawCandidates
    .map((candidate) => ({
      asin: normalizeWhitespace(candidate.asin),
      title: normalizeWhitespace(candidate.title),
      productUrl: normalizeWhitespace(candidate.productUrl),
      priceText: normalizeWhitespace(candidate.priceText),
      price: parsePrice(candidate.priceText),
      ratingText: normalizeWhitespace(candidate.ratingText),
      ratingAverage: parseRating(candidate.ratingText),
      ratingCountText: normalizeWhitespace(candidate.ratingCountText),
      ratingCount: parseRatingCount(candidate.ratingCountText),
      brand: normalizeWhitespace(candidate.brand),
      imageUrl: normalizeWhitespace(candidate.imageUrl),
      categoryHint: normalizeWhitespace(category || ''),
      matchedKeywords: [keyword],
      matchedScenarios: [],
      searchPage: pageNumber,
      marketplace,
    }))
    .filter((candidate) => candidate.asin && candidate.title && candidate.productUrl);
}

function buildScenarios(candidates) {
  const scenarioMap = new Map();

  for (const candidate of candidates) {
    const label = inferScenarioLabel(candidate.title, candidate.matchedKeywords);
    if (!scenarioMap.has(label)) {
      scenarioMap.set(label, {
        id: 0,
        label,
        candidateAsins: new Set(),
        kept: true,
      });
    }
    const scenario = scenarioMap.get(label);
    scenario.candidateAsins.add(candidate.asin);
    candidate.scenarioLabel = label;
    candidate.matchedScenarios = [label];
  }

  const scenarios = Array.from(scenarioMap.values())
    .map((scenario) => ({
      id: 0,
      label: scenario.label,
      kept: true,
      candidateCount: scenario.candidateAsins.size,
      candidateAsins: Array.from(scenario.candidateAsins).sort(),
    }))
    .sort((left, right) => {
      if (right.candidateCount !== left.candidateCount) {
        return right.candidateCount - left.candidateCount;
      }
      return left.label.localeCompare(right.label);
    })
    .map((scenario, index) => ({
      ...scenario,
      id: index + 1,
    }));

  const idByLabel = new Map(scenarios.map((scenario) => [scenario.label, scenario.id]));
  for (const candidate of candidates) {
    candidate.scenarioId = idByLabel.get(candidate.scenarioLabel) || 0;
  }

  return scenarios;
}

function recomputeSelection(payload) {
  const excluded = new Set(payload.excludedScenarioIds || []);
  const scenarioById = new Map(payload.scenarios.map((scenario) => [scenario.id, scenario]));
  for (const scenario of payload.scenarios) {
    scenario.kept = !excluded.has(scenario.id);
  }

  const filteredCandidates = payload.candidates
    .filter((candidate) => {
      const scenario = scenarioById.get(candidate.scenarioId);
      candidate.kept = Boolean(scenario && scenario.kept);
      return candidate.kept;
    })
    .sort(compareCandidates);

  const selected = filteredCandidates.slice(0, payload.topN).map((candidate, index) => {
    candidate.selectedForExecution = true;
    candidate.executionRank = index + 1;
    return candidate;
  });

  for (const candidate of payload.candidates) {
    if (!selected.find((item) => item.asin === candidate.asin)) {
      candidate.selectedForExecution = false;
      candidate.executionRank = null;
    }
  }

  payload.estimate = {
    keptScenarioCount: payload.scenarios.filter((scenario) => scenario.kept).length,
    candidateCount: filteredCandidates.length,
    estimatedReviewCount: filteredCandidates.reduce((sum, candidate) => sum + (candidate.ratingCount || 0), 0),
    executionProductCount: selected.length,
  };

  return payload;
}

async function discoverProducts(options) {
  const { browser, context } = await createDiscoveryContext(options.headless, options.locale);
  const page = await context.newPage();
  const deduped = new Map();
  const rawHits = [];

  try {
    for (const keyword of options.keywords) {
      for (let pageNumber = 1; pageNumber <= options.pages; pageNumber += 1) {
        console.log(
          JSON.stringify({
            stage: 'discovery_search',
            keyword,
            pageNumber,
            marketplace: options.marketplace,
          })
        );

        const candidates = await extractSearchCandidates(page, options.marketplace, keyword, options.category, pageNumber);
        rawHits.push({
          keyword,
          pageNumber,
          candidateCount: candidates.length,
        });

        for (const candidate of candidates) {
          const existing = deduped.get(candidate.asin);
          if (existing) {
            existing.matchedKeywords = Array.from(new Set([...existing.matchedKeywords, keyword]));
            existing.ratingCount = Math.max(existing.ratingCount || 0, candidate.ratingCount || 0);
            existing.ratingAverage = existing.ratingAverage || candidate.ratingAverage;
            existing.price = existing.price === null ? candidate.price : existing.price;
            existing.priceText = existing.priceText || candidate.priceText;
            existing.brand = existing.brand || candidate.brand;
            existing.categoryHint = existing.categoryHint || candidate.categoryHint;
            continue;
          }
          deduped.set(candidate.asin, candidate);
        }
      }
    }
  } finally {
    await context.close();
    await browser.close();
  }

  const filtered = Array.from(deduped.values()).filter((candidate) => candidatePasses(candidate, options));
  filtered.sort(compareCandidates);

  const payload = {
    workflowVersion: 2,
    stage: 'preflight',
    status: 'preflight',
    marketplace: options.marketplace,
    keywords: options.keywords,
    category: options.category || '',
    priceMin: options.priceMin,
    priceMax: options.priceMax,
    minRating: options.minRating,
    topN: options.topN,
    pages: options.pages,
    rawHits,
    excludedScenarioIds: [],
    scenarios: buildScenarios(filtered),
    candidates: filtered,
    estimate: {
      keptScenarioCount: 0,
      candidateCount: 0,
      estimatedReviewCount: 0,
      executionProductCount: 0,
    },
  };

  return recomputeSelection(payload);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const keywords = String(args.keywords || '')
    .split(/[,\n，]/)
    .map((keyword) => normalizeWhitespace(keyword))
    .filter(Boolean);

  if (!args.marketplace || keywords.length === 0) {
    console.error(
      'Usage: node scripts/amazon-product-discovery.js --marketplace amazon.sg --keywords "keyword1,keyword2" [--category Electronics] [--price-min 10] [--price-max 60] [--min-rating 4.0] [--top-n 5] [--pages 2]'
    );
    process.exit(1);
  }

  const payload = await discoverProducts({
    marketplace: String(args.marketplace).replace(/^https?:\/\//, ''),
    keywords,
    category: args.category || '',
    priceMin: args['price-min'] ? parseFloat(args['price-min']) : null,
    priceMax: args['price-max'] ? parseFloat(args['price-max']) : null,
    minRating: args['min-rating'] ? parseFloat(args['min-rating']) : null,
    topN: args['top-n'] ? parseInt(args['top-n'], 10) : 5,
    pages: args.pages ? parseInt(args.pages, 10) : 2,
    headless: args.headless !== 'false',
    locale: args.locale || 'zh-HK',
  });

  console.log(JSON.stringify(payload, null, 2));
}

module.exports = {
  discoverProducts,
  recomputeSelection,
  compareCandidates,
  inferScenarioLabel,
  keywordToScenarioLabel,
};

if (require.main === module) {
  main().catch((error) => {
    console.error(`Failed: ${error.message}`);
    process.exit(1);
  });
}
