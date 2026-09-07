// 与本机 Plan Studio 服务器的全部通信。token 来自启动时终端打印的 URL。
const token = new URLSearchParams(location.search).get('token') || '';

async function call(method, path, body) {
  const res = await fetch(path, {
    method,
    headers: {
      'X-Studio-Token': token,
      ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}),
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    let detail = res.statusText;
    try { detail = (await res.json()).detail || detail; } catch { /* 保底用 statusText */ }
    const err = new Error(detail);
    err.status = res.status;
    throw err;
  }
  return res.json();
}

export const api = {
  state: () => call('GET', '/api/state'),
  savePlan: (plan) => call('PUT', '/api/plan', { plan }),
  render: () => call('POST', '/api/render'),
  renderStatus: () => call('GET', '/api/render/status'),
  exportAll: (dwg) => call('POST', '/api/export', { dwg }),
  modelUrl: () => `/api/model.glb?token=${encodeURIComponent(token)}`,
  previewUrl: (p) => `${p}?token=${encodeURIComponent(token)}`,
};
