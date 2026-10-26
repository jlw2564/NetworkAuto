'use strict';

const $ = (id) => document.getElementById(id);
let currentJob = null;
let eventSource = null;

// ---------------- 初始化 ----------------
async function init() {
  try {
    const cfg = await (await fetch('/api/config')).json();
    $('switchFile').placeholder = cfg.default_switch_file + (cfg.default_switch_exists ? '（已存在）' : '（缺失）');
    $('terminalFile').placeholder = cfg.default_terminal_file + (cfg.default_terminal_exists ? '（已存在）' : '（缺失）');

    const box = $('sections');
    box.innerHTML = '';
    cfg.sections.forEach((s) => {
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = s.key;
      cb.checked = true;
      cb.dataset.section = s.key;
      label.appendChild(cb);
      label.appendChild(document.createTextNode(s.label));
      box.appendChild(label);
    });
  } catch (e) {
    appendLog('ERROR', '读取配置失败: ' + e);
  }
  refreshHistory();
}

// ---------------- 日志 ----------------
function appendLog(level, message) {
  const pre = $('log');
  if (pre.dataset.placeholder !== '1') {
    pre.textContent = '';
    pre.dataset.placeholder = '1';
  }
  const line = document.createElement('span');
  const ts = new Date().toLocaleTimeString('zh-CN', { hour12: false });
  line.textContent = `[${ts}] ${message}\n`;
  if (level === 'WARNING' || level === 'WARN') line.className = 'warn';
  if (level === 'ERROR' || level === 'CRITICAL') line.className = 'error';
  pre.appendChild(line);
  pre.scrollTop = pre.scrollHeight;
}

// ---------------- 进度 ----------------
function setStatus(status) {
  const badge = $('statusBadge');
  badge.textContent = { running: '运行中', done: '已完成', failed: '失败' }[status] || status;
  badge.className = 'badge ' + status;
}

function setProgress(stage, current, total, detail) {
  const pct = total > 0 ? Math.min(100, Math.round((current / total) * 100)) : 0;
  $('progressBar').style.width = pct + '%';
  const stageName = {
    load: '读取清单', collect: '采集交换机', discover: '发现未知交换机',
    analyze: '分析', report: '生成报告',
  }[stage] || stage;
  $('stageText').textContent = `${stageName} ${current}/${total}${detail ? ' — ' + detail : ''}`;
}

// ---------------- 上传 ----------------
function bindUpload(buttonId, inputId, targetId) {
  $(buttonId).onclick = () => $(inputId).click();
  $(inputId).onchange = async () => {
    const file = $(inputId).files[0];
    if (!file) return;
    const fd = new FormData();
    fd.append('file', file);
    appendLog('INFO', '正在上传 ' + file.name + ' ...');
    try {
      const res = await fetch('/api/upload', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || '上传失败');
      $(targetId).value = data.path;
      appendLog('INFO', '上传成功: ' + data.path);
    } catch (e) {
      appendLog('ERROR', '上传失败: ' + e.message);
    }
  };
}

// ---------------- 运行 ----------------
async function startRun() {
  const sections = Array.from(document.querySelectorAll('#sections input:checked'))
    .map((cb) => cb.dataset.section);
  if (sections.length === 0) {
    alert('请至少勾选一个功能模块');
    return;
  }
  const body = {
    switch_file: $('switchFile').value.trim() || null,
    terminal_file: $('terminalFile').value.trim() || null,
    sections,
    concurrency: parseInt($('concurrency').value, 10) || 6,
    discover: $('discover').value === 'true',
  };

  $('btnRun').disabled = true;
  $('btnDownload').disabled = true;
  $('log').textContent = '';
  $('log').dataset.placeholder = '1';
  $('sheetTabs').innerHTML = '';
  $('tableWrap').innerHTML = '<div class="muted" style="padding:12px">任务运行中…</div>';
  $('stats').innerHTML = '';

  try {
    const res = await fetch('/api/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || '启动失败');
    currentJob = data.job_id;
    setStatus('running');
    setProgress('load', 0, 1, '任务已提交');
    subscribe(currentJob);
  } catch (e) {
    appendLog('ERROR', '启动失败: ' + e.message);
    $('btnRun').disabled = false;
  }
}

function subscribe(jobId) {
  if (eventSource) eventSource.close();
  eventSource = new EventSource(`/api/jobs/${jobId}/events`);
  eventSource.onmessage = (ev) => {
    let data;
    try { data = JSON.parse(ev.data); } catch (e) { return; }
    handleEvent(data);
  };
  eventSource.onerror = () => {
    if (currentJob) {
      eventSource.close();
      eventSource = null;
    }
  };
}

function handleEvent(ev) {
  if (ev.kind === 'log') {
    appendLog(ev.level, ev.message);
  } else if (ev.kind === 'progress') {
    setProgress(ev.stage, ev.current, ev.total, ev.detail);
  } else if (ev.kind === 'done') {
    finish(ev);
  }
}

async function finish(ev) {
  if (eventSource) { eventSource.close(); eventSource = null; }
  $('btnRun').disabled = false;
  const jobId = currentJob;
  const snap = await (await fetch(`/api/jobs/${jobId}`)).json();
  const ok = snap.status === 'done';
  setStatus(snap.status);
  if (ok) {
    $('progressBar').style.width = '100%';
    $('stageText').textContent = `完成，用时 ${snap.result?.duration ?? '-'} 秒`;
    renderStats(snap.result?.stats || {}, snap.result?.devices || [], snap.result?.unknown_switches || []);
    $('btnDownload').disabled = false;
    $('btnDownload').onclick = () => window.open(`/api/jobs/${jobId}/download`, '_blank');
    loadSheets(jobId);
  } else {
    $('stageText').textContent = '任务失败: ' + (snap.error || '未知错误');
  }
  refreshHistory();
}

function renderStats(stats, devices, unknown) {
  const box = $('stats');
  box.innerHTML = '';
  Object.entries(stats).forEach(([k, v]) => {
    const d = document.createElement('div');
    d.className = 'stat';
    d.innerHTML = `<div class="k">${k}</div><div class="v">${v}</div>`;
    box.appendChild(d);
  });
  const okCount = devices.filter((d) => d.ok).length;
  const d = document.createElement('div');
  d.className = 'stat';
  d.innerHTML = `<div class="k">交换机采集</div><div class="v">${okCount}/${devices.length}</div>`;
  box.appendChild(d);
  if (unknown.length) {
    const u = document.createElement('div');
    u.className = 'stat';
    u.innerHTML = `<div class="k">清单外交换机</div><div class="v">${unknown.length}</div>`;
    box.appendChild(u);
  }
}

// ---------------- 结果预览 ----------------
async function loadSheets(jobId) {
  try {
    const res = await fetch(`/api/jobs/${jobId}/sheets`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail);
    const tabs = $('sheetTabs');
    tabs.innerHTML = '';
    data.sheets.forEach((name, idx) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.onclick = () => {
        Array.from(tabs.children).forEach((c) => c.classList.remove('active'));
        btn.classList.add('active');
        previewSheet(`/api/jobs/${jobId}/sheets/${encodeURIComponent(name)}`);
      };
      tabs.appendChild(btn);
      if (idx === 0) { btn.classList.add('active'); previewSheet(`/api/jobs/${jobId}/sheets/${encodeURIComponent(name)}`); }
    });
  } catch (e) {
    $('tableWrap').innerHTML = `<div class="muted" style="padding:12px">加载结果失败: ${e.message}</div>`;
  }
}

async function previewSheet(url) {
  const wrap = $('tableWrap');
  wrap.innerHTML = '<div class="muted" style="padding:12px">加载中…</div>';
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail);
    if (!data.rows.length) {
      wrap.innerHTML = '<div class="muted" style="padding:12px">该 Sheet 无数据</div>';
      return;
    }
    const thead = '<tr>' + data.columns.map((c) => `<th>${escapeHtml(c)}</th>`).join('') + '</tr>';
    const tbody = data.rows.map((row) =>
      '<tr>' + row.map((v) => `<td title="${escapeHtml(String(v))}">${escapeHtml(String(v))}</td>`).join('') + '</tr>'
    ).join('');
    wrap.innerHTML = `<table><thead>${thead}</thead><tbody>${tbody}</tbody></table>`
      + `<div class="hint" style="padding:8px 4px">共 ${data.total} 行，展示前 ${data.rows.length} 行。完整数据请下载 Excel。</div>`;
  } catch (e) {
    wrap.innerHTML = `<div class="muted" style="padding:12px">加载失败: ${e.message}</div>`;
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

// ---------------- 历史 ----------------
async function refreshHistory() {
  const box = $('history');
  try {
    const res = await fetch('/api/history');
    const data = await res.json();
    if (!data.history.length) {
      box.innerHTML = '<div class="muted">暂无记录</div>';
      return;
    }
    box.innerHTML = '';
    data.history.forEach((h) => {
      const row = document.createElement('div');
      row.className = 'history-item';
      const left = document.createElement('div');
      left.innerHTML = `<div>${h.created_at || h.job_id}</div>`
        + `<div class="muted">${(h.options?.sections || []).length} 个模块 · ${h.status || ''}</div>`;
      const right = document.createElement('div');
      const view = document.createElement('button');
      view.textContent = '查看';
      view.style.padding = '2px 8px';
      view.onclick = () => loadHistorySheets(h.job_id);
      right.appendChild(view);
      if (h.has_output) {
        const dl = document.createElement('button');
        dl.textContent = '下载';
        dl.style.padding = '2px 8px';
        dl.style.marginLeft = '4px';
        dl.onclick = () => window.open(`/api/history/${h.job_id}/download`, '_blank');
        right.appendChild(dl);
      }
      row.appendChild(left);
      row.appendChild(right);
      box.appendChild(row);
    });
  } catch (e) {
    box.innerHTML = `<div class="muted">读取历史失败: ${e}</div>`;
  }
}

async function loadHistorySheets(jobId) {
  try {
    const res = await fetch(`/api/history/${jobId}/sheets`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail);
    const tabs = $('sheetTabs');
    tabs.innerHTML = '';
    data.sheets.forEach((name, idx) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.onclick = () => {
        Array.from(tabs.children).forEach((c) => c.classList.remove('active'));
        btn.classList.add('active');
        previewSheet(`/api/history/${jobId}/sheets/${encodeURIComponent(name)}`);
      };
      tabs.appendChild(btn);
      if (idx === 0) { btn.classList.add('active'); previewSheet(`/api/history/${jobId}/sheets/${encodeURIComponent(name)}`); }
    });
  } catch (e) {
    $('tableWrap').innerHTML = `<div class="muted" style="padding:12px">加载历史失败: ${e.message}</div>`;
  }
}

// ---------------- 事件绑定 ----------------
bindUpload('btnUploadSwitch', 'fileSwitch', 'switchFile');
bindUpload('btnUploadTerminal', 'fileTerminal', 'terminalFile');
$('btnRun').onclick = startRun;
$('btnRefreshHistory').onclick = refreshHistory;
$('btnClearLog').onclick = () => { $('log').textContent = '已清屏'; $('log').dataset.placeholder = ''; };

init();
