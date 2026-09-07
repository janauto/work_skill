// 装配入口：布局、工具条、键盘、右栏页签、校验问题面板。
import {
  boot, bus, state, setSelection, flush,
} from './state.js';
import { initViewer } from './viewer.js';
import { initParts } from './parts.js';
import { initFigures, addSelectionTo } from './figures.js';
import { initTerms } from './terms.js';
import { initPreview } from './preview.js';

const $ = (sel) => document.querySelector(sel);

function renderToolbar() {
  const stepName = state.step.split('/').pop();
  const badge = {
    clean: '', dirty: '未保存…', saving: '保存中…',
    saved: '已保存 ✓', error: '保存失败！',
  }[state.saveState];
  const v = state.validate;
  const errs = (v?.issues || []).filter((i) => i.severity === 'error').length;
  const warns = (v?.issues || []).length - errs;
  $('#toolbar').innerHTML = `
    <div class="brand">图纸规划台 <span class="brand-sub">Plan Studio</span></div>
    <div class="tb-step mono" title="${state.step}">${stepName}</div>
    <div class="tb-validate ${errs ? 'v-err' : (warns ? 'v-warn' : 'v-ok')}"
         title="服务器校验（与模型路线同一套错误码）">
      ${errs ? `${errs} 错` : ''}${errs && warns ? ' · ' : ''}${warns ? `${warns} 提示` : ''}
      ${!errs && !warns ? '计划可渲染' : ''}
    </div>
    <div class="tb-save ${state.saveState}">${badge}</div>`;
}

function renderIssues() {
  const list = state.validate?.issues || [];
  const el = $('#issues');
  if (!list.length) { el.innerHTML = ''; el.hidden = true; return; }
  el.hidden = false;
  el.innerHTML = `<div class="panel-head sub"><span class="stamp">检</span>校验问题
    <span class="head-count">${list.length}</span></div>
    <ul class="issue-list">${list.map((it) => `
      <li class="issue ${it.severity}">
        <b class="mono">${it.code}</b> ${it.message}
        ${it.hint ? `<div class="qa-hint">↳ ${it.hint}</div>` : ''}
      </li>`).join('')}</ul>`;
}

function initTabs() {
  document.querySelectorAll('.rtab').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.rtab').forEach((b) => b.classList.toggle('on', b === btn));
      document.querySelectorAll('.rpane').forEach((p) => {
        p.hidden = p.dataset.pane !== btn.dataset.tab;
      });
    });
  });
}

function initKeyboard() {
  document.addEventListener('keydown', (e) => {
    if (e.target.matches('input, select, textarea')) return;
    if (e.key === 'Escape') setSelection([]);
    const n = Number(e.key);
    if (n >= 1 && n <= 9 && state.selection.size) {
      const fig = state.plan.figures[n - 1];
      if (fig) addSelectionTo(fig.id);
    }
  });
  window.addEventListener('beforeunload', flush);
}

async function start() {
  try {
    await boot();
  } catch (err) {
    document.body.innerHTML = `<div class="fatal">连接不上 Plan Studio 服务器：${err.message}
      <br>请从终端打印的完整地址（含 token）进入。</div>`;
    return;
  }
  renderToolbar(); renderIssues();
  initParts($('#parts'));
  initFigures($('#figures'));
  initTerms($('#terms'), $('#params'));
  initPreview($('#drawer'));
  initTabs(); initKeyboard();
  bus.emit('booted');

  const viewer = await initViewer($('#viewport'));
  $('#btn-frame').addEventListener('click', () => viewer.frameAll());
  $('#btn-isolate').addEventListener('change', (e) => viewer.setIsolate(e.target.checked));

  ['save-state', 'validate', 'booted'].forEach((evt) => bus.on(evt, () => {
    renderToolbar(); renderIssues();
  }));
}

start();
