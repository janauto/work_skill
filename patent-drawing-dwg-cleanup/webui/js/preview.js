// 底部抽屉：渲染、QA 红绿灯、SVG 预览。预览里点附图标记数字 → 3D 里对应零件高亮，
// 核对「7 号是不是上盖」不再靠数。导出按钮只在全部 QA 通过后出现——闸门不可绕。
import { api } from './api.js';
import { bus, state, setSelection } from './state.js';
import { flashParts } from './viewer.js';

let root; let activeTab = null; let pollTimer = null;

function checkRow(c) {
  const mark = c.pass ? '<span class="qa-ok">✓</span>' : '<span class="qa-bad">✗</span>';
  const hint = !c.pass && c.hint ? `<div class="qa-hint">↳ ${c.hint}</div>` : '';
  return `<li class="${c.pass ? '' : 'qa-fail'}">${mark}
    <span class="qa-id mono">${c.id}</span>
    <span class="qa-val mono">${c.value}</span>
    <span class="qa-th mono">${c.threshold}</span>${hint}</li>`;
}

function render() {
  const r = state.render;
  const busy = state.rendering;
  const figs = r?.figures || [];
  if (!activeTab && figs.length) activeTab = figs[0].id;
  const allPass = !!(r && r.ok && figs.length && figs.every((f) => f.pass));

  const tabs = figs.map((f) => `
    <button class="tab${f.id === activeTab ? ' on' : ''}${f.pass ? '' : ' tab-fail'}"
            data-tab="${f.id}">${f.id} ${f.pass ? '✓' : '✗'}</button>`).join('');
  const current = figs.find((f) => f.id === activeTab);
  const qa = current?.qa?.checks?.map(checkRow).join('') || '';
  const numeralNote = r?.numerals
    ? `<p class="hint">附图标记说明：${r.numerals.description_zh || ''}</p>` : '';

  root.innerHTML = `
    <div class="drawer-bar">
      <button class="btn-render" data-render ${busy ? 'disabled' : ''}>
        ${busy ? '<span class="spin"></span> 渲染中…' : '渲染出图'}</button>
      <div class="tabs">${tabs}</div>
      <div class="drawer-right">
        ${allPass ? `<label class="check-inline">
            <input type="checkbox" data-dwg checked>同时转 DWG</label>
          <button class="btn-export" data-export>导出交付件</button>`
    : (figs.length ? '<span class="hint">有图未过 QA，按提示修改后重渲</span>' : '')}
      </div>
    </div>
    ${current ? `<div class="drawer-body">
      <div class="svg-pane" data-pane>${'' /* SVG 注入见下 */}</div>
      <div class="qa-pane">
        <div class="qa-title">${current.pass
    ? '<span class="stamp-pass">合格</span>' : '<span class="stamp-fail">未通过</span>'}
          QA ${current.qa ? `${current.qa.summary?.passed ?? ''}项通过` : ''}</div>
        <ul class="qa-list">${qa}</ul>${numeralNote}
      </div>
    </div>` : (busy ? '' : '<p class="hint pad">填好计划后点「渲染出图」。渲染走与模型完全相同的 CLI 与 QA 闸门。</p>')}
    ${r?.log && !figs.length ? `<pre class="log">${r.log.replace(/</g, '&lt;')}</pre>` : ''}`;

  root.querySelector('[data-render]')?.addEventListener('click', startRender);
  root.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => {
    activeTab = t.dataset.tab; render();
  }));
  root.querySelector('[data-export]')?.addEventListener('click', doExport);

  if (current) loadSvg(current);
}

async function loadSvg(fig) {
  const pane = root.querySelector('[data-pane]');
  if (!pane) return;
  try {
    const res = await fetch(api.previewUrl(fig.svg));
    pane.innerHTML = await res.text();      // 本机服务器生成的 SVG，可信来源
    const svg = pane.querySelector('svg');
    if (svg) {
      svg.querySelectorAll('text[data-numeral]').forEach((t) => {
        t.addEventListener('click', () => {
          const numeral = Number(t.dataset.numeral);
          const entry = (state.render?.numerals?.numerals || [])
            .find((n) => n.numeral === numeral);
          if (entry?.parts?.length) {
            setSelection(entry.parts);
            flashParts(entry.parts);
          }
        });
      });
    }
  } catch { pane.innerHTML = '<p class="hint pad">预览加载失败</p>'; }
}

async function startRender() {
  try { await api.render(); } catch (err) {
    window.alert(err.message);
    return;
  }
  state.rendering = true;
  render();
  poll();
}

function poll() {
  clearTimeout(pollTimer);
  pollTimer = setTimeout(async () => {
    const st = await api.renderStatus();
    state.rendering = st.running;
    if (!st.running) {
      state.render = st.result;
      activeTab = null;
      bus.emit('rendered');
      render();
    } else poll();
  }, 900);
}

async function doExport() {
  const dwg = root.querySelector('[data-dwg]')?.checked ?? false;
  const btn = root.querySelector('[data-export]');
  btn.disabled = true; btn.textContent = '导出中…';
  try {
    const res = await api.exportAll(dwg);
    window.alert(`已导出 ${res.files.length} 个文件：\n${res.dir}\n\n`
      + res.files.join('\n') + (res.dwg_log.length ? `\n\n${res.dwg_log.join('\n')}` : ''));
  } catch (err) { window.alert('导出失败：' + err.message); }
  render();
}

export function initPreview(el) {
  root = el;
  bus.on('booted', () => { if (state.rendering) poll(); render(); });
}
