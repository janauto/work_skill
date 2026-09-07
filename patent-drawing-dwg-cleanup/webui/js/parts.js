// 左栏：零件清单（点选兜底——深腔零件 3D 里点不到，这里永远点得到）+ 拆分建议牌。
import {
  bus, state, mutate, toggleSelect, setSelection, figuresContaining, figureColor,
  unassignedParts, setActiveFigure,
} from './state.js';

let root; let query = '';

function sizeText(p) {
  const [w, d, h] = p.bbox_size || [];
  return (w === undefined) ? '' : `${w.toFixed(0)}×${d.toFixed(0)}×${h.toFixed(0)}`;
}

function render() {
  const parts = state.assembly?.parts || [];
  const missing = new Set(unassignedParts());
  const rows = parts
    .filter((p) => !query || p.name.toLowerCase().includes(query))
    .map((p) => {
      const figs = figuresContaining(p.name);
      const dots = figs.map((f) =>
        `<i class="dot" style="background:${figureColor(f.id)}" title="${f.id}"></i>`).join('');
      const sel = state.selection.has(p.name) ? ' is-selected' : '';
      const warn = missing.has(p.name) ? '<span class="tag-warn">未入图</span>' : '';
      const degen = p.degenerate ? '<span class="tag-warn">退化</span>' : '';
      return `<li class="part-row${sel}" data-name="${p.name}">
        <span class="part-name mono">${p.name}</span>
        <span class="part-meta">×${p.instances} · ${sizeText(p)}</span>
        <span class="part-flags">${dots}${warn}${degen}</span>
      </li>`;
    }).join('');

  const sugg = (state.assembly?.split_suggestions || []).map((s, i) => `
    <button class="sugg" data-sugg="${i}">
      <b>${{ coaxial: '按共轴组', stack: '按装配栈', size: '按尺寸档' }[s.strategy] || s.strategy}</b>
      <span>${s.figures.length} 张图 · 脚本计算</span>
    </button>`).join('');

  root.innerHTML = `
    <div class="panel-head"><span class="stamp">01</span>零件
      <span class="head-count">${parts.length} 种</span></div>
    <input class="search" placeholder="搜索零件名…" value="${query}">
    <ul class="part-list">${rows}</ul>
    <div class="sel-actions">
      <span>${state.selection.size ? `已选 ${state.selection.size} 种` : '点 3D 或列表选件'}</span>
      ${state.selection.size ? '<button class="link" data-act="clear">清空</button>' : ''}
    </div>
    ${sugg ? `<div class="panel-head sub"><span class="stamp">拆</span>拆分建议</div>${sugg}` : ''}
  `;

  root.querySelector('.search').addEventListener('input', (e) => {
    query = e.target.value.trim().toLowerCase();
    render();
    root.querySelector('.search').focus();
    const el = root.querySelector('.search');
    el.setSelectionRange(el.value.length, el.value.length);
  });
  root.querySelectorAll('.part-row').forEach((li) => {
    li.addEventListener('click', () => toggleSelect(li.dataset.name));
  });
  const clear = root.querySelector('[data-act="clear"]');
  if (clear) clear.addEventListener('click', () => setSelection([]));
  root.querySelectorAll('.sugg').forEach((btn) => {
    btn.addEventListener('click', () => applySuggestion(Number(btn.dataset.sugg)));
  });
}

function applySuggestion(index) {
  const s = state.assembly.split_suggestions[index];
  if (!s) return;
  const ok = window.confirm(
    `用「${s.id}」建议重建全部图卡（${s.figures.length} 张）？现有图卡会被替换，术语表不受影响。`);
  if (!ok) return;
  mutate((plan) => {
    plan.figures = s.figures.map((f, i) => ({
      id: `fig${i + 1}`,
      caption: f.caption_hint || `分解示意图${i + 1}`,
      kind: 'exploded',
      members: [...f.members],
    }));
  });
  setActiveFigure('fig1');
}

export function initParts(el) {
  root = el;
  ['booted', 'plan', 'selection', 'active-figure'].forEach((evt) => bus.on(evt, render));
}
