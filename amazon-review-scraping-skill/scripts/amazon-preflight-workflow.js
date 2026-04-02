#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { discoverProducts, recomputeSelection } = require('./amazon-product-discovery');

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

function writeJson(filePath, payload) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${JSON.stringify(payload, null, 2)}\n`);
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
}

function parseKeywords(rawKeywords) {
  return String(rawKeywords || '')
    .split(/[,\n，]/)
    .map((keyword) => normalizeWhitespace(keyword))
    .filter(Boolean);
}

function parseExcludedScenarioIds(reply) {
  const normalized = normalizeWhitespace(reply);
  const matches = normalized.match(/\d+/g) || [];
  return Array.from(new Set(matches.map((match) => parseInt(match, 10)).filter(Number.isInteger)));
}

function estimatedReviewCountText(count) {
  if (count >= 10000) {
    return `${Math.floor(count / 1000)}万+`;
  }
  if (count >= 1000) {
    return `${Math.floor(count / 1000)}000+`;
  }
  return String(count);
}

function buildSummaryText(state, extraMessage = '') {
  const keptScenarios = state.scenarios.filter((scenario) => scenario.kept);
  const scenarioLines = state.scenarios.map((scenario) => {
    const prefix = scenario.kept ? `${scenario.id}.` : `${scenario.id}. [已排除]`;
    return `${prefix} ${scenario.label} (${scenario.candidateCount}个候选商品)`;
  });

  const topCandidates = state.candidates
    .filter((candidate) => candidate.kept)
    .sort((left, right) => (right.ratingCount || 0) - (left.ratingCount || 0))
    .slice(0, Math.min(state.topN, 10))
    .map(
      (candidate, index) =>
        `${index + 1}. ${candidate.title} | ASIN ${candidate.asin} | 评论数 ${candidate.ratingCount || 0} | 场景 ${candidate.scenarioLabel}`
    );

  const lines = [
    extraMessage,
    '编号场景清单：',
    ...scenarioLines,
    '',
    `当前范围共覆盖 ${state.estimate.keptScenarioCount} 个品类`,
    `预计需抓取 ${state.estimate.candidateCount} 个商品`,
    `预计涉及 ${estimatedReviewCountText(state.estimate.estimatedReviewCount)} 条评论`,
    `正式执行默认抓取评论数最高的 Top ${state.topN} 商品`,
    '',
    '当前保留场景：' + (keptScenarios.length ? keptScenarios.map((scenario) => scenario.id).join('、') : '无'),
    '如需继续排除，请回复：不搜索：编号1、编号2',
    '确认无误后，请回复：开始执行',
  ];

  if (topCandidates.length > 0) {
    lines.push('', '当前入选执行商品预览：', ...topCandidates);
  }

  return lines.filter((line, index, array) => !(line === '' && array[index - 1] === '')).join('\n');
}

async function createInitialState(args) {
  const keywords = parseKeywords(args.keywords);
  if (!args.marketplace || keywords.length === 0) {
    throw new Error(
      'Usage: node scripts/amazon-preflight-workflow.js --marketplace amazon.sg --keywords "keyword1,keyword2" [--category Electronics] [--price-min 10] [--price-max 60] [--min-rating 4.0] [--top-n 5] [--output-dir ./output]'
    );
  }

  const outputDir = path.resolve(args['output-dir'] || path.join(process.cwd(), 'preflight_output'));
  const statePath = path.join(outputDir, 'preflight_state.json');
  const discoveryPayload = await discoverProducts({
    marketplace: String(args.marketplace).replace(/^https?:\/\//, ''),
    keywords,
    category: args.category || '',
    priceMin: args['price-min'] ? parseFloat(args['price-min']) : null,
    priceMax: args['price-max'] ? parseFloat(args['price-max']) : null,
    minRating: args['min-rating'] ? parseFloat(args['min-rating']) : 4.0,
    topN: args['top-n'] ? parseInt(args['top-n'], 10) : 5,
    pages: args.pages ? parseInt(args.pages, 10) : 2,
    headless: args.headless !== 'false',
    locale: args.locale || 'zh-HK',
  });

  const state = {
    ...discoveryPayload,
    outputDir,
    statePath,
    replyHistory: [],
  };

  writeJson(statePath, state);

  if (process.env.NO_FETCH_CANDIDATE_IMAGES !== 'true') {
    const imageScript = path.join(__dirname, 'amazon-fetch-product-images.js');
    const imageResult = spawnSync(process.execPath, [imageScript, '--state', statePath], {
      stdio: 'inherit',
      env: process.env,
    });
    state.productImageFetchStatus = imageResult.status === 0 ? 'success' : 'failed';
    writeJson(statePath, state);
  }

  return state;
}

function handleReply(state, reply) {
  const normalizedReply = normalizeWhitespace(reply);
  state.replyHistory = state.replyHistory || [];
  state.replyHistory.push({
    at: new Date().toISOString(),
    text: normalizedReply,
  });

  if (normalizedReply === '开始执行') {
    state.status = 'ready_to_execute';
    return {
      state,
      shouldExecute: true,
      message: '收到“开始执行”，现在开始正式抓取。',
    };
  }

  const scenarioIds = parseExcludedScenarioIds(normalizedReply);
  if (scenarioIds.length === 0) {
    return {
      state,
      shouldExecute: false,
      message: '未识别到需要排除的编号，已保留当前范围。',
    };
  }

  const validIds = new Set(state.scenarios.map((scenario) => scenario.id));
  const nextExcluded = new Set(state.excludedScenarioIds || []);
  const appliedIds = [];
  for (const id of scenarioIds) {
    if (validIds.has(id)) {
      nextExcluded.add(id);
      appliedIds.push(id);
    }
  }
  if (appliedIds.length === 0) {
    return {
      state,
      shouldExecute: false,
      message: '未识别到有效场景编号，已保留当前范围。',
    };
  }
  state.excludedScenarioIds = Array.from(nextExcluded).sort((left, right) => left - right);
  recomputeSelection(state);

  return {
    state,
    shouldExecute: false,
    message: `已排除场景：${appliedIds.join('、')}`,
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  let state;
  let extraMessage = '';

  if (args.state) {
    const statePath = path.resolve(args.state);
    state = readJson(statePath);
    state.statePath = statePath;

    if (args.reply) {
      const result = handleReply(state, args.reply);
      state = result.state;
      extraMessage = result.message;
      writeJson(state.statePath, state);

      if (result.shouldExecute) {
        console.log(buildSummaryText(state, extraMessage));
        const executeScript = path.join(__dirname, 'amazon-competitor-execute.js');
        const runResult = spawnSync(process.execPath, [executeScript, '--state', state.statePath], {
          stdio: 'inherit',
          env: process.env,
        });
        process.exit(runResult.status || 0);
      }
    }
  } else {
    state = await createInitialState(args);
    extraMessage = '已生成预检范围初稿。';
  }

  recomputeSelection(state);
  writeJson(state.statePath, state);
  console.log(buildSummaryText(state, extraMessage));
  console.log(
    JSON.stringify(
      {
        stage: 'preflight_ready',
        statePath: state.statePath,
        keptScenarioCount: state.estimate.keptScenarioCount,
        candidateCount: state.estimate.candidateCount,
        executionProductCount: state.estimate.executionProductCount,
      },
      null,
      2
    )
  );
}

if (require.main === module) {
  main().catch((error) => {
    console.error(`Failed: ${error.message}`);
    process.exit(1);
  });
}
