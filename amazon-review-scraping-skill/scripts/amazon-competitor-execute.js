#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { compareCandidates } = require('./amazon-product-discovery');

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

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
}

function writeJson(filePath, payload) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${JSON.stringify(payload, null, 2)}\n`);
}

function formatTimestamp(date) {
  return date.toISOString().replace(/[:.]/g, '-');
}

function selectProducts(state) {
  return state.candidates
    .filter((candidate) => candidate.kept)
    .sort(compareCandidates)
    .slice(0, state.topN)
    .map((candidate, index) => ({
      ...candidate,
      executionRank: index + 1,
    }));
}

function loadReviewSummary(jsonPath) {
  const payload = readJson(jsonPath);
  return {
    reviewCount: payload.reviewCount || 0,
    imageReviewCount: payload.imageReviewCount || 0,
    downloadedImageCount: payload.downloadedImageCount || 0,
    mediaDir: payload.mediaDir || '',
  };
}

function canReuseExistingResult(jsonPath) {
  if (!fs.existsSync(jsonPath)) {
    return false;
  }

  try {
    const summary = loadReviewSummary(jsonPath);
    return summary.reviewCount > 0;
  } catch (_error) {
    return false;
  }
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.state) {
    console.error('Usage: node scripts/amazon-competitor-execute.js --state ./preflight_state.json [--output-dir ./output]');
    process.exit(1);
  }

  const statePath = path.resolve(args.state);
  const state = readJson(statePath);
  const outputDir = path.resolve(args['output-dir'] || state.outputDir || path.dirname(statePath));
  const runId = formatTimestamp(new Date());
  const runDir = path.join(outputDir, 'runs', runId);
  const reviewDir = path.join(runDir, 'reviews');
  const manifestPath = path.join(runDir, 'competitor_run_manifest.json');
  const workbookPath = path.join(runDir, 'competitor_reviews.xlsx');
  const selectedProducts = selectProducts(state);
  const reviewScript = path.join(__dirname, 'amazon-review-login-scrape.js');
  const excelScript = path.join(__dirname, 'amazon-competitor-to-excel.py');
  const pythonBin = process.env.PYTHON_BIN || 'python3';

  fs.mkdirSync(reviewDir, { recursive: true });

  const manifest = {
    workflowVersion: 2,
    stage: 'execute',
    createdAt: new Date().toISOString(),
    statePath,
    outputDir,
    runDir,
    workbookPath,
    marketplace: state.marketplace,
    keywords: state.keywords,
    category: state.category || '',
    priceMin: state.priceMin ?? null,
    priceMax: state.priceMax ?? null,
    minRating: state.minRating ?? null,
    topN: state.topN,
    scenarios: state.scenarios,
    excludedScenarioIds: state.excludedScenarioIds || [],
    candidateCount: state.estimate?.candidateCount || 0,
    estimatedReviewCount: state.estimate?.estimatedReviewCount || 0,
    selectedProducts,
    products: [],
  };

  console.log(
    JSON.stringify(
      {
        stage: 'execute_start',
        statePath,
        selectedProductCount: selectedProducts.length,
        runDir,
      },
      null,
      2
    )
  );

  for (const product of selectedProducts) {
    const jsonPath = path.join(reviewDir, `${product.asin}.json`);
    const env = {
      ...process.env,
      SESSION_ROOT: process.env.SESSION_ROOT || path.join(outputDir, '.sessions'),
    };

    console.log(
      JSON.stringify(
        {
          stage: 'execute_product',
          asin: product.asin,
          title: product.title,
          productUrl: product.productUrl,
          jsonPath,
        }
      )
    );

    const record = {
      asin: product.asin,
      title: product.title,
      productUrl: product.productUrl,
      scenarioId: product.scenarioId,
      scenarioLabel: product.scenarioLabel,
      matchedKeywords: product.matchedKeywords,
      ratingAverage: product.ratingAverage,
      ratingCount: product.ratingCount,
      price: product.price,
      priceText: product.priceText,
      brand: product.brand,
      executionRank: product.executionRank,
      status: result.status === 0 ? 'success' : 'failed',
      jsonPath,
      mediaDir: '',
      reviewCount: 0,
      imageReviewCount: 0,
      imageCount: 0,
    };

    if (process.env.SKIP_EXISTING_SUCCESS !== 'false' && canReuseExistingResult(jsonPath)) {
      const summary = loadReviewSummary(jsonPath);
      record.status = 'success';
      record.skippedExisting = true;
      record.reviewCount = summary.reviewCount;
      record.imageReviewCount = summary.imageReviewCount;
      record.imageCount = summary.downloadedImageCount;
      record.mediaDir = summary.mediaDir;
      console.log(
        JSON.stringify({
          stage: 'execute_product_skip',
          asin: product.asin,
          jsonPath,
          reviewCount: summary.reviewCount,
          imageCount: summary.downloadedImageCount,
        })
      );
      manifest.products.push(record);
      writeJson(manifestPath, manifest);
      continue;
    }

    const result = spawnSync(process.execPath, [reviewScript, product.productUrl, jsonPath], {
      stdio: 'inherit',
      env,
    });

    if (fs.existsSync(jsonPath)) {
      const summary = loadReviewSummary(jsonPath);
      record.status = result.status === 0 && summary.reviewCount > 0 ? 'success' : 'failed';
      record.reviewCount = summary.reviewCount;
      record.imageReviewCount = summary.imageReviewCount;
      record.imageCount = summary.downloadedImageCount;
      record.mediaDir = summary.mediaDir;
    }

    manifest.products.push(record);
    writeJson(manifestPath, manifest);
  }

  const excelArgs = [excelScript, manifestPath, workbookPath];
  if (process.env.NO_TRANSLATE === 'true') {
    excelArgs.push('--no-translate');
  }

  const excelResult = spawnSync(pythonBin, excelArgs, {
    stdio: 'inherit',
    env: process.env,
  });

  manifest.excelStatus = excelResult.status === 0 ? 'success' : 'failed';
  manifest.workbookPath = workbookPath;
  writeJson(manifestPath, manifest);

  state.status = 'executed';
  state.lastRunManifest = manifestPath;
  state.lastWorkbookPath = workbookPath;
  state.lastExecutedAt = new Date().toISOString();
  writeJson(statePath, state);

  if (excelResult.status !== 0) {
    process.exit(excelResult.status || 1);
  }

  console.log(
    JSON.stringify(
      {
        stage: 'execute_done',
        manifestPath,
        workbookPath,
        successfulProducts: manifest.products.filter((product) => product.status === 'success').length,
        failedProducts: manifest.products.filter((product) => product.status !== 'success').length,
      },
      null,
      2
    )
  );
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error(`Failed: ${error.message}`);
    process.exit(1);
  }
}
