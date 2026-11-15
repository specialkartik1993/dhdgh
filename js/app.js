/* MP WCD Commissioner Dashboard – UI, charts and live data sync. */
(function () {
  'use strict';

  /* ================= Config ================= */
  const DATA_URL = new URLSearchParams(location.search).get('src') || 'data/dashboard.xlsx';
  const SERVER_POLL_MS = 15000;   // check hosted workbook for changes
  const LOCAL_POLL_MS = 2000;     // check a locally-picked file for saves
  const PAGE_SIZE = 25;

  const INDICATORS = [
    { k: 'me', label: 'Measurement Efficiency', short: 'Measurement' },
    { k: 'thr', label: 'THR Coverage', short: 'THR' },
    { k: 'hcm', label: 'HCM Coverage', short: 'HCM' },
    { k: 'op', label: 'AWC Operational Rate', short: 'AWC operational' },
    { k: 'hv', label: 'Home Visit Completion', short: 'Home visits' },
    { k: 'infra', label: 'Basic Infrastructure Coverage', short: 'Infrastructure' },
    { k: 'anc', label: 'ANC Coverage', short: 'ANC' },
  ];
  const BURDEN = [
    { k: 'samPct', label: 'SAM % of Measured', lowerBetter: true },
    { k: 'mamPct', label: 'MAM % of Measured', lowerBetter: true },
    { k: 'anaemiaPct', label: 'Anaemia % of Pregnant Women', lowerBetter: true },
  ];
  const METRICS = [...INDICATORS, ...BURDEN];
  const metricOf = (k) => METRICS.find((m) => m.k === k);

  /* ================= State ================= */
  const state = {
    data: null,
    view: 'overview',
    division: '',
    divMetric: 'me',
    rankMetric: 'me',
    sorts: {},
    search: { district: '', nut: '', mat: '', awc: '' },
    awcPage: 0,
    awcDistrict: '',
    awcBuilding: '',
    alertPrio: '',
    alertStatus: '',
    alertCat: '',
    source: { mode: 'server', handle: null, key: null, sig: null, lastMod: null, fileName: '' },
    loadedAt: null,
    prevKpi: {},
    drawerDistrict: null,
  };
  const charts = {};

  /* ================= Utils ================= */
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = new Intl.NumberFormat('en-IN');
  const fInt = (v) => (v === null || v === undefined || !isFinite(v) ? '—' : nf.format(Math.round(v)));
  const fPct = (v, d = 1) => (v === null || v === undefined || !isFinite(v) ? '—' : (v * 100).toFixed(d) + '%');
  const compact = (v) => {
    if (v === null || !isFinite(v)) return '—';
    if (Math.abs(v) >= 1e7) return (v / 1e7).toFixed(2) + ' Cr';
    if (Math.abs(v) >= 1e5) return (v / 1e5).toFixed(2) + ' L';
    return nf.format(Math.round(v));
  };
  const rag = (v) => (v === null || v === undefined ? 'n' : v >= 0.9 ? 'g' : v >= 0.75 ? 'a' : 'r');
  const RAG_LABEL = { g: 'On track', a: 'Watch', r: 'Off track', n: 'No data' };
  const ragChip = (v, text) => `<span class="rag ${rag(v)}" title="${RAG_LABEL[rag(v)]}">${text ?? fPct(v)}</span>`;
  const cssVar = (n) => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
  const debounce = (fn, ms = 180) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
  const timeAgo = (d) => {
    const s = Math.round((Date.now() - d) / 1000);
    if (s < 10) return 'just now';
    if (s < 60) return s + 's ago';
    if (s < 3600) return Math.round(s / 60) + 'm ago';
    return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  };

  function animateNumber(el, to, fmt, from) {
    const start = from ?? 0;
    if (!isFinite(to)) { el.textContent = fmt(to); return; }
    if (matchMedia('(prefers-reduced-motion: reduce)').matches || start === to) { el.textContent = fmt(to); return; }
    const dur = 1000, t0 = performance.now();
    const step = (t) => {
      const p = Math.min(1, (t - t0) / dur);
      const e = 1 - Math.pow(1 - p, 3);
      el.textContent = fmt(start + (to - start) * e);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  async function digest(buf) {
    try {
      if (crypto.subtle) {
        const h = await crypto.subtle.digest('SHA-1', buf);
        return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, '0')).join('');
      }
    } catch (e) { /* fall through */ }
    // FNV-1a fallback for non-secure contexts
    const a = new Uint8Array(buf); let h = 0x811c9dc5;
    for (let i = 0; i < a.length; i++) { h ^= a[i]; h = Math.imul(h, 0x01000193); }
    return (h >>> 0).toString(16) + ':' + a.length;
  }

  /* ================= Toasts & status ================= */
  function toast(title, msg = '', kind = 'ok', ms = 4200) {
    const el = document.createElement('div');
    el.className = 'toast ' + kind;
    el.innerHTML = `<span class="t-ic">${kind === 'err' ? '!' : kind === 'info' ? 'i' : '✓'}</span><div><b>${esc(title)}</b>${msg ? esc(msg) : ''}</div>`;
    $('#toasts').appendChild(el);
    setTimeout(() => { el.classList.add('out'); el.addEventListener('animationend', () => el.remove(), { once: true }); }, ms);
  }
  const progress = (on) => $('#progress').classList.toggle('on', on);

  function setSync(kind, text) {
    const s = $('#sync');
    s.classList.toggle('live', kind === 'live');
    s.classList.toggle('error', kind === 'error');
    $('#syncText').textContent = text;
    s.title = text;
  }
  function refreshSyncLabel() {
    if (!state.loadedAt) return;
    const src = state.source;
    const where = src.mode === 'local' ? `Watching ${src.fileName}` : src.mode === 'upload' ? `Uploaded ${src.fileName}` : 'Live';
    setSync(src.mode === 'upload' ? 'idle' : 'live', `${where} · updated ${timeAgo(state.loadedAt)}`);
    $('#srcServerState').textContent = src.mode === 'server' ? 'Active' : '';
    $('#srcLocalState').textContent = src.mode === 'local' ? 'Active' : '';
  }

  /* ================= Data loading ================= */
  async function applyBuffer(buf, origin) {
    const sig = await digest(buf);
    if (sig === state.source.sig) return false;
    const model = WCD.parseWorkbook(buf);
    const prev = state.data;
    state.source.sig = sig;
    state.data = model;
    state.loadedAt = new Date();
    onData(prev, origin);
    return true;
  }

  async function loadServer({ force = false, quiet = false } = {}) {
    if (!quiet) progress(true);
    try {
      // Cheap change check via HEAD (ETag / Last-Modified) before downloading.
      if (!force && state.source.key) {
        try {
          const h = await fetch(DATA_URL, { method: 'HEAD', cache: 'no-store' });
          if (h.ok) {
            const key = [h.headers.get('etag'), h.headers.get('last-modified'), h.headers.get('content-length')].join('|');
            if (key !== '||' && key === state.source.key) { refreshSyncLabel(); return false; }
          }
        } catch (e) { /* HEAD unsupported: fall back to full GET */ }
      }
      const res = await fetch(DATA_URL + (DATA_URL.includes('?') ? '&' : '?') + 't=' + Date.now(), { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status} while fetching ${DATA_URL}`);
      state.source.key = [res.headers.get('etag'), res.headers.get('last-modified'), res.headers.get('content-length')].join('|');
      const changed = await applyBuffer(await res.arrayBuffer(), 'server');
      refreshSyncLabel();
      return changed;
    } catch (err) {
      if (!state.data) showError(err);
      else if (!quiet) toast('Refresh failed', ' ' + err.message, 'err');
      setSync('error', 'Sync error · retrying');
      return false;
    } finally {
      progress(false);
    }
  }

  async function pollLocal() {
    const h = state.source.handle;
    if (!h) return;
    try {
      const file = await h.getFile();
      if (file.lastModified === state.source.lastMod) { refreshSyncLabel(); return; }
      progress(true);
      const buf = await file.arrayBuffer();
      await applyBuffer(buf, 'local');
      state.source.lastMod = file.lastModified;
      refreshSyncLabel();
    } catch (err) {
      // Excel may hold a lock or be mid-save; try again on the next tick.
      setSync('error', 'Waiting for file…');
    } finally { progress(false); }
  }

  let pollTimer = null;
  function startPolling() {
    clearInterval(pollTimer);
    const tick = () => {
      if (document.hidden) return;
      if (state.source.mode === 'server') loadServer({ quiet: true });
      else if (state.source.mode === 'local') pollLocal();
    };
    const ms = state.source.mode === 'local' ? LOCAL_POLL_MS : SERVER_POLL_MS;
    if (state.source.mode !== 'upload') pollTimer = setInterval(tick, ms);
  }

  async function pickLocalFile() {
    if (window.showOpenFilePicker) {
      try {
        const [handle] = await window.showOpenFilePicker({
          types: [{ description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'], 'application/vnd.ms-excel': ['.xls'] } }],
        });
        const file = await handle.getFile();
        progress(true);
        state.source.sig = null;
        await applyBuffer(await file.arrayBuffer(), 'local-first');
        Object.assign(state.source, { mode: 'local', handle, lastMod: file.lastModified, fileName: file.name });
        closeModal();
        toast('Watching ' + file.name, ' Save the workbook in Excel and the dashboard updates automatically.', 'info', 6000);
        startPolling(); refreshSyncLabel();
      } catch (err) {
        if (err.name !== 'AbortError') toast('Could not open file', ' ' + err.message, 'err');
      } finally { progress(false); }
    } else {
      $('#fileInput').click();
    }
  }

  async function loadUploadedFile(file) {
    if (!file) return;
    progress(true);
    try {
      state.source.sig = null;
      await applyBuffer(await file.arrayBuffer(), 'upload');
      Object.assign(state.source, { mode: 'upload', handle: null, fileName: file.name });
      closeModal();
      toast('Loaded ' + file.name, ' This is a one-off copy. Use "Choose Excel file" in Chrome/Edge to keep it live.', 'info', 6000);
      startPolling(); refreshSyncLabel();
    } catch (err) {
      toast('Could not read workbook', ' ' + err.message, 'err');
    } finally { progress(false); }
  }

  function useServer() {
    Object.assign(state.source, { mode: 'server', handle: null, key: null, sig: null, fileName: '' });
    closeModal();
    loadServer({ force: true }).then(() => toast('Using hosted workbook', ' ' + DATA_URL, 'info'));
    startPolling();
  }

  function showError(err) {
    $('#errorMsg').textContent = err.message || String(err);
    $$('.view').forEach((v) => v.classList.remove('active'));
    $('#errorView').classList.add('active');
    setSync('error', 'Not loaded');
  }

  /* ================= On new data ================= */
  function diffCount(prev, next) {
    if (!prev) return 0;
    const pm = new Map(prev.districts.map((d) => [d.district, JSON.stringify(d)]));
    let n = 0;
    next.districts.forEach((d) => { if (pm.get(d.district) !== JSON.stringify(d)) n++; });
    n += Math.abs(prev.districts.length - next.districts.length);
    const awcChanged = JSON.stringify(prev.awcs) !== JSON.stringify(next.awcs);
    const alertChanged = JSON.stringify(prev.alerts) !== JSON.stringify(next.alerts);
    return { districts: n, awcChanged, alertChanged };
  }

  function onData(prev, origin) {
    const d = state.data;
    $('#skeleton').classList.remove('active');
    $('#errorView').classList.remove('active');

    // Division filter options
    const divs = [...new Set(d.districts.map((x) => x.division))].sort();
    if (state.division && !divs.includes(state.division)) state.division = '';
    $('#divisionFilter').innerHTML = '<option value="">All divisions</option>' +
      divs.map((v) => `<option ${v === state.division ? 'selected' : ''}>${esc(v)}</option>`).join('');

    $('#period').textContent = d.meta.period ? 'Reporting period · ' + d.meta.period : '';
    $('#footNote').textContent = (d.meta.caution ? d.meta.caution + ' ' : '') +
      `${d.districts.length} districts · ${d.awcs.length} AWC records · ${d.alerts.length} alerts. RAG: ≥90% on track · 75–90% watch · <75% off track.`;

    if (prev && origin === 'server' || prev && origin === 'local') {
      const c = diffCount(prev, d);
      const parts = [];
      if (c.districts) parts.push(`${c.districts} district row${c.districts > 1 ? 's' : ''} changed`);
      if (c.awcChanged) parts.push('AWC records changed');
      if (c.alertChanged) parts.push('alerts changed');
      toast('Dashboard updated from Excel', parts.length ? ' ' + parts.join(' · ') + '.' : ' New workbook loaded.');
    }

    renderBadge();
    showView(state.view, { keepAnim: !!prev });
    if (state.drawerDistrict) openDrawer(state.drawerDistrict, true);
  }

  /* ================= Filters ================= */
  const fDistricts = () => state.data.districts.filter((d) => !state.division || d.division === state.division);
  const fAwcs = () => state.data.awcs.filter((a) => !state.division || a.division === state.division);
  const fAlerts = () => state.data.alerts.filter((a) => !state.division || a.division === state.division);

  /* ================= Views ================= */
  function showView(v, { keepAnim = false } = {}) {
    state.view = v;
    $$('#tabs button').forEach((b) => {
      const on = b.dataset.view === v;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on);
      if (on) {
        const t = $('#tabs');
        if (t.scrollWidth > t.clientWidth) t.scrollTo({ left: b.offsetLeft - (t.clientWidth - b.offsetWidth) / 2, behavior: 'smooth' });
      }
    });
    if (!state.data) return;
    $$('.view').forEach((s) => s.classList.toggle('active', s.id === 'view-' + v));
    const sec = $('#view-' + v);
    $$('.card', sec).forEach((c, i) => c.style.setProperty('--i', i));
    if (keepAnim) $$('.card', sec).forEach((c) => (c.style.animation = 'none'));
    else $$('.card', sec).forEach((c) => (c.style.animation = ''));
    RENDER[v]();
    try { history.replaceState(null, '', '#' + v); } catch (e) { /* ignore */ }
  }

  function renderBadge() {
    const open = fAlerts().filter((a) => !/closed|resolved|done/i.test(a.status)).length;
    $('#alertBadge').textContent = open || '';
  }

  /* ---------- KPI tiles ---------- */
  const ICONS = {
    home: '<path d="M3 11l9-7 9 7M5 10v10h14V10"/>',
    kids: '<circle cx="9" cy="7" r="3"/><path d="M3 20a6 6 0 0 1 12 0M16 4a3 3 0 0 1 0 6M21 20a6 6 0 0 0-4-5.6"/>',
    heart: '<path d="M12 20s-7-4.4-7-10a4 4 0 0 1 7-2.6A4 4 0 0 1 19 10c0 5.6-7 10-7 10z"/>',
    ruler: '<path d="M4 16L16 4l4 4L8 20zM8 12l2 2M11 9l2 2M14 6l2 2"/>',
    alert: '<path d="M12 3l10 18H2zM12 10v4m0 3v.01"/>',
    drop: '<path d="M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z"/>',
    user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    check: '<path d="M4 12l5 5L20 6"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    water: '<path d="M7 3h10l-1 6H8zM8 9c0 6 1 12 4 12s4-6 4-12"/>',
    wc: '<path d="M6 3v18M18 3v18M6 12h12"/>',
  };

  function renderKpis(container, items) {
    const el = $(container);
    if (el.children.length !== items.length) {
      el.innerHTML = items.map((it, i) => `
        <article class="card kpi" style="--i:${i}" data-key="${it.key}">
          <div class="kpi-label"><span class="kpi-icon"><svg viewBox="0 0 24 24">${ICONS[it.icon] || ICONS.chart}</svg></span>${esc(it.label)}</div>
          <div class="kpi-value">—</div>
          <div class="kpi-sub"></div>
          ${it.meter !== undefined ? '<div class="meter"><i></i></div>' : ''}
        </article>`).join('');
    }
    items.forEach((it, i) => {
      const card = el.children[i];
      card.dataset.key = it.key;
      $('.kpi-label', card).lastChild.textContent = it.label;
      const valEl = $('.kpi-value', card);
      const prev = state.prevKpi[it.key];
      animateNumber(valEl, it.value, it.fmt || fInt, prev);
      $('.kpi-sub', card).innerHTML = it.sub || '';
      if (prev !== undefined && prev !== it.value && isFinite(it.value)) {
        card.classList.remove('flash'); void card.offsetWidth; card.classList.add('flash');
        const diff = it.value - prev;
        const good = it.lowerBetter ? diff < 0 : diff > 0;
        const txt = (diff > 0 ? '▲ ' : '▼ ') + (it.fmt === fPctK ? (Math.abs(diff) * 100).toFixed(1) + ' pts' : fInt(Math.abs(diff)));
        $('.kpi-sub', card).insertAdjacentHTML('afterbegin', `<span class="delta ${good ? 'up' : 'down'}">${txt}</span>`);
      }
      state.prevKpi[it.key] = it.value;
      const m = $('.meter i', card);
      if (m) {
        m.style.background = it.meterColor || cssVar('--series-1');
        requestAnimationFrame(() => (m.style.width = Math.max(0, Math.min(1, it.meter)) * 100 + '%'));
      }
    });
  }
  const fPctK = (v) => fPct(v);
  const ragColor = (v) => ({ g: cssVar('--good'), a: cssVar('--warn'), r: cssVar('--crit'), n: cssVar('--muted') }[rag(v)]);

  /* ---------- Chart helpers ---------- */
  function chartTheme() {
    Chart.defaults.font.family = 'Inter, system-ui, sans-serif';
    Chart.defaults.font.size = 12;
    Chart.defaults.color = cssVar('--text-2');
    Chart.defaults.borderColor = cssVar('--grid');
    Chart.defaults.animation.duration = 800;
    Chart.defaults.animation.easing = 'easeOutQuart';
    Chart.defaults.maintainAspectRatio = false;
    Chart.defaults.plugins.legend.display = false;
    Object.assign(Chart.defaults.plugins.tooltip, {
      backgroundColor: cssVar('--text'), titleColor: cssVar('--bg'), bodyColor: cssVar('--bg'),
      padding: 10, cornerRadius: 8, boxPadding: 4, usePointStyle: true,
    });
  }

  function upsertChart(id, cfg) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const ex = charts[id];
    if (ex && ex.config.type === cfg.type && ex.canvas === canvas) {
      ex.data.labels = cfg.data.labels;
      cfg.data.datasets.forEach((ds, i) => {
        if (ex.data.datasets[i]) Object.assign(ex.data.datasets[i], ds);
        else ex.data.datasets.push(ds);
      });
      ex.data.datasets.length = cfg.data.datasets.length;
      ex.options = cfg.options;
      ex.update();
      return ex;
    }
    if (ex) ex.destroy();
    charts[id] = new Chart(canvas, cfg);
    return charts[id];
  }
  function destroyCharts() { Object.keys(charts).forEach((k) => { charts[k].destroy(); delete charts[k]; }); }

  const pctAxis = (max) => ({
    beginAtZero: true, max, grid: { color: cssVar('--grid') }, border: { display: false },
    ticks: { callback: (v) => Math.round(v * 100) + '%' },
  });
  const catAxis = () => ({ grid: { display: false }, border: { display: false } });
  const barStyle = (color) => ({
    backgroundColor: color, hoverBackgroundColor: color, borderRadius: 4, borderSkipped: false,
    borderColor: cssVar('--surface'), borderWidth: 0, maxBarThickness: 26,
  });
  const legendHTML = (items) => items.map(([c, t, line]) => `<span><i class="${line ? 'line' : ''}" style="${line ? '' : 'background:' + c}"></i>${esc(t)}</span>`).join('');
  const niceMax = (vals) => { const m = Math.max(0.0001, ...vals.filter(isFinite)); return Math.min(1, Math.ceil(m * 20 + 0.0001) / 20); };

  /* ================= OVERVIEW ================= */
  function renderOverview() {
    const ds = fDistricts();
    const A = WCD.aggregate(ds);
    const al = fAlerts();
    const active = al.filter((a) => !/closed|resolved|done/i.test(a.status));
    renderKpis('#kpis', [
      { key: 'awcs', label: 'Total AWCs', icon: 'home', value: A.awcs, sub: `${A.n} district${A.n !== 1 ? 's' : ''}${state.division ? ' in ' + esc(state.division) : ''}` },
      { key: 'children', label: 'Active children (0–6)', icon: 'kids', value: A.children, sub: `${compact(A.children)} enrolled` },
      { key: 'pw', label: 'Pregnant women', icon: 'heart', value: A.pw, sub: `${fInt(A.lm)} lactating mothers` },
      { key: 'me', label: 'Measurement efficiency', icon: 'ruler', value: A.me, fmt: fPctK, meter: A.me, meterColor: ragColor(A.me), sub: `${ragChip(A.me, RAG_LABEL[rag(A.me)])} ${fInt(A.measured)} measured` },
      { key: 'sam', label: 'SAM cases', icon: 'alert', value: A.sam, lowerBetter: true, sub: `${fPct(A.samPct)} of measured` },
      { key: 'mam', label: 'MAM cases', icon: 'drop', value: A.mam, lowerBetter: true, sub: `${fPct(A.mamPct)} of measured` },
      { key: 'awwV', label: 'AWW vacancies', icon: 'user', value: A.awwV, lowerBetter: true, sub: `${fPct(A.awwVPct)} of posts · AWH ${fInt(A.awhV)} (${fPct(A.awhVPct)})` },
      { key: 'alerts', label: 'Active alerts', icon: 'alert', value: active.length, lowerBetter: true, sub: `${active.filter((a) => /high/i.test(a.priority)).length} high priority` },
    ]);

    // RAG health
    const good = cssVar('--good'), warn = cssVar('--warn'), crit = cssVar('--crit');
    $('#ragLegend').innerHTML = legendHTML([[good, 'On track ≥90%'], [warn, 'Watch 75–90%'], [crit, 'Off track <75%']]);
    const counts = INDICATORS.map((m) => {
      const c = { g: 0, a: 0, r: 0 };
      ds.forEach((d) => { const r = rag(d[m.k]); if (c[r] !== undefined) c[r]++; });
      return c;
    });
    const seg = (color) => ({ ...barStyle(color), borderColor: cssVar('--surface'), borderWidth: { left: 1, right: 1 }, borderRadius: 0, maxBarThickness: 22 });
    upsertChart('ragChart', {
      type: 'bar',
      data: {
        labels: INDICATORS.map((m) => m.short),
        datasets: [
          { label: 'On track', data: counts.map((c) => c.g), ...seg(good) },
          { label: 'Watch', data: counts.map((c) => c.a), ...seg(warn) },
          { label: 'Off track', data: counts.map((c) => c.r), ...seg(crit) },
        ],
      },
      options: {
        indexAxis: 'y',
        scales: { x: { stacked: true, max: ds.length, grid: { color: cssVar('--grid') }, border: { display: false }, ticks: { precision: 0 } }, y: { stacked: true, ...catAxis() } },
        plugins: { tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${c.raw} district${c.raw !== 1 ? 's' : ''}` } } },
      },
    });

    // Division comparison
    const metric = metricOf(state.divMetric);
    const state0 = WCD.aggregate(state.data.districts)[metric.k];
    let labels, vals, mode;
    if (state.division) {
      mode = 'district';
      const sorted = [...ds].sort((a, b) => (b[metric.k] ?? -1) - (a[metric.k] ?? -1));
      labels = sorted.map((d) => d.district); vals = sorted.map((d) => d[metric.k]);
      $('#divChartSub').textContent = `${metric.label} · districts in ${state.division} vs state average`;
    } else {
      mode = 'division';
      const g = WCD.groupBy(ds, 'division');
      const rows = [...g.entries()].map(([k, arr]) => [k, WCD.aggregate(arr)[metric.k]]).sort((a, b) => (b[1] ?? -1) - (a[1] ?? -1));
      labels = rows.map((r) => r[0]); vals = rows.map((r) => r[1]);
      $('#divChartSub').textContent = `${metric.label} · tap a division to filter`;
    }
    const s1 = metric.lowerBetter ? cssVar('--series-2') : cssVar('--series-1');
    upsertChart('divChart', {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { type: 'bar', label: metric.label, data: vals, ...barStyle(s1), order: 2 },
          { type: 'line', label: 'State average', data: labels.map(() => state0), borderColor: cssVar('--text-2'), borderDash: [5, 4], borderWidth: 1.5, pointRadius: 0, pointHitRadius: 0, order: 1 },
        ],
      },
      options: {
        scales: { y: pctAxis(niceMax([...vals, state0])), x: { ...catAxis(), ticks: { autoSkip: false, maxRotation: 50, minRotation: 0 } } },
        interaction: { mode: 'index', intersect: false },
        plugins: { tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${fPct(c.raw)}` } } },
        onHover: (e, el) => (e.native.target.style.cursor = el.length ? 'pointer' : 'default'),
        onClick: (e, el) => {
          if (!el.length) return;
          const lab = labels[el[0].index];
          if (mode === 'division') setDivision(lab); else openDrawer(lab);
        },
      },
    });

    // Top / bottom
    const ranked = ds.filter((d) => d[metric.k] !== null).sort((a, b) => b[metric.k] - a[metric.k]);
    const best = metric.lowerBetter ? [...ranked].reverse() : ranked;
    const worst = [...best].reverse();
    $('#topSub').textContent = `Best ${metric.label}`;
    $('#bottomSub').textContent = `Lowest ${metric.label}`.replace('Lowest', metric.lowerBetter ? 'Highest' : 'Lowest');
    const li = (d, i) => {
      const v = d[metric.k];
      const color = metric.lowerBetter ? cssVar('--series-2') : ragColor(v);
      const w = metric.lowerBetter ? Math.min(1, v / (ranked[0][metric.k] || 1)) : v;
      return `<li style="--i:${i}" data-district="${esc(d.district)}"><span class="rank-n">${i + 1}</span>
        <div class="rank-name">${esc(d.district)}<small>${esc(d.division)}</small><div class="rank-bar"><i style="width:${w * 100}%;background:${color}"></i></div></div>
        ${metric.lowerBetter ? `<b>${fPct(v)}</b>` : ragChip(v)}</li>`;
    };
    $('#topList').innerHTML = best.slice(0, 5).map(li).join('');
    $('#bottomList').innerHTML = worst.slice(0, 5).map(li).join('');

    // Alerts summary
    const cats = WCD.groupBy(active, 'category');
    const prio = (arr, p) => arr.filter((a) => a.priority.toLowerCase() === p).length;
    const rows = [...cats.entries()].sort((a, b) => b[1].length - a[1].length);
    $('#alertSummary').innerHTML = rows.length ? rows.map(([cat, arr]) => `
      <div class="as-row" data-cat="${esc(cat)}" style="cursor:pointer">
        <span class="as-name">${esc(cat)}</span><span class="as-count">${arr.length}</span>
        <div class="stack" title="High ${prio(arr, 'high')} · Medium ${prio(arr, 'medium')} · Low ${prio(arr, 'low')}">
          <i style="flex-grow:${prio(arr, 'high')};background:var(--crit)"></i>
          <i style="flex-grow:${prio(arr, 'medium')};background:var(--warn)"></i>
          <i style="flex-grow:${prio(arr, 'low')};background:var(--good)"></i>
        </div>
      </div>`).join('') + `<div class="as-foot legend" style="margin:0">${legendHTML([[crit, 'High'], [warn, 'Medium'], [good, 'Low']])}<span style="margin-left:auto">${active.length} active of ${al.length}</span></div>`
      : '<p class="hint">No active alerts 🎉</p>';
  }

  /* ================= Generic table ================= */
  function renderTable(id, columns, rows, { onRow, defaultSort, empty = 'No matching rows' } = {}) {
    const el = document.getElementById(id);
    const s = state.sorts[id] || defaultSort || { k: columns[0].k, dir: 1 };
    state.sorts[id] = s;
    const col = columns.find((c) => c.k === s.k) || columns[0];
    const val = col.val || ((r) => r[col.k]);
    const sorted = [...rows].sort((a, b) => {
      const x = val(a), y = val(b);
      if (x === y) return 0;
      if (x === null || x === undefined) return 1;
      if (y === null || y === undefined) return -1;
      return (typeof x === 'string' ? x.localeCompare(y) : x - y) * s.dir;
    });
    el.innerHTML = `<thead><tr>${columns.map((c) => `<th data-k="${c.k}" class="${c.num ? 'num' : ''} ${c.k === s.k ? 'sorted' : ''}">${esc(c.label)}<span class="arrow">${c.k === s.k ? (s.dir > 0 ? '▲' : '▼') : '↕'}</span></th>`).join('')}</tr></thead>
      <tbody>${sorted.length ? sorted.map((r, i) => `<tr class="${onRow ? 'click' : ''}" data-i="${rows.indexOf(r)}" style="animation-delay:${Math.min(i, 30) * 12}ms">${columns.map((c) => `<td class="${c.num ? 'num' : ''} ${c.cls || ''}">${c.fmt ? c.fmt(r) : esc(r[c.k])}</td>`).join('')}</tr>`).join('')
        : `<tr class="empty-row"><td colspan="${columns.length}">${empty}</td></tr>`}</tbody>`;
    el.onclick = (e) => {
      const th = e.target.closest('th');
      if (th) {
        const k = th.dataset.k;
        state.sorts[id] = { k, dir: s.k === k ? -s.dir : (columns.find((c) => c.k === k).num ? -1 : 1) };
        renderTable(id, columns, rows, { onRow, defaultSort, empty });
        return;
      }
      const tr = e.target.closest('tr[data-i]');
      if (tr && onRow) onRow(rows[+tr.dataset.i]);
    };
    return sorted;
  }
  const bar = (v, color) => `<span class="cell-bar">${fPct(v)}<b><i style="width:${Math.max(0, Math.min(1, v || 0)) * 100}%;background:${color}"></i></b></span>`;
  const pctCol = (k, label) => ({ k, label, num: true, fmt: (r) => ragChip(r[k]) });
  const burdenCol = (k, label) => ({ k, label, num: true, fmt: (r) => fPct(r[k]) });
  const intCol = (k, label) => ({ k, label, num: true, fmt: (r) => fInt(r[k]) });
  const districtCol = { k: 'district', label: 'District', cls: 'strong', fmt: (r) => `${esc(r.district)}<span class="sub">${esc(r.division)}</span>` };
  const matchQ = (q, ...fields) => !q || fields.some((f) => String(f || '').toLowerCase().includes(q.toLowerCase()));

  /* ================= DISTRICTS ================= */
  const DISTRICT_COLS = [
    districtCol,
    intCol('awcs', 'AWCs'), intCol('children', 'Children 0–6'),
    pctCol('me', 'Measurement'), pctCol('thr', 'THR'), pctCol('hcm', 'HCM'), pctCol('op', 'AWC op.'),
    pctCol('hv', 'Home visits'), pctCol('infra', 'Infra.'), pctCol('anc', 'ANC'),
    burdenCol('samPct', 'SAM %'), burdenCol('mamPct', 'MAM %'), burdenCol('anaemiaPct', 'Anaemia %'),
    intCol('awwV', 'AWW vacant'), intCol('awhV', 'AWH vacant'),
  ];

  function renderDistricts() {
    const ds = fDistricts();
    const metric = metricOf(state.rankMetric);
    const sorted = ds.filter((d) => d[metric.k] !== null).sort((a, b) => b[metric.k] - a[metric.k]);
    const s2 = cssVar('--series-2');
    $('#rankLegend').innerHTML = metric.lowerBetter
      ? legendHTML([[s2, metric.label + ' (higher = more burden)'], [null, 'State average', true]])
      : legendHTML([[cssVar('--good'), 'On track ≥90%'], [cssVar('--warn'), 'Watch 75–90%'], [cssVar('--crit'), 'Off track <75%'], [null, 'State average', true]]);
    $('#rankSub').textContent = `${sorted.length} districts by ${metric.label} · tap a bar to drill down`;
    $('#rankBox').style.height = Math.max(220, sorted.length * 22 + 50) + 'px';
    const stateAvg = WCD.aggregate(state.data.districts)[metric.k];
    const colors = sorted.map((d) => (metric.lowerBetter ? s2 : ragColor(d[metric.k])));
    upsertChart('rankChart', {
      type: 'bar',
      data: {
        labels: sorted.map((d) => d.district),
        datasets: [
          { label: metric.label, data: sorted.map((d) => d[metric.k]), ...barStyle(colors), maxBarThickness: 14, order: 2 },
          { type: 'line', label: 'State average', data: sorted.map(() => stateAvg), borderColor: cssVar('--text-2'), borderDash: [5, 4], borderWidth: 1.5, pointRadius: 0, pointHitRadius: 0, order: 1 },
        ],
      },
      options: {
        indexAxis: 'y',
        scales: {
          x: { ...pctAxis(niceMax([...sorted.map((d) => d[metric.k]), stateAvg])), position: 'top' },
          y: { ...catAxis(), ticks: { autoSkip: false, font: { size: 11 } } },
        },
        interaction: { mode: 'index', intersect: false, axis: 'y' },
        plugins: { tooltip: { callbacks: { title: (c) => `${c[0].label} · ${sorted[c[0].dataIndex].division}`, label: (c) => ` ${c.dataset.label}: ${fPct(c.raw)}` } } },
        onHover: (e, el) => (e.native.target.style.cursor = el.length ? 'pointer' : 'default'),
        onClick: (e, el) => { if (el.length) openDrawer(sorted[el[0].index].district); },
      },
    });

    const rows = ds.filter((d) => matchQ(state.search.district, d.district, d.division));
    $('#tableSub').textContent = `${rows.length} of ${ds.length} districts · tap a row for details`;
    renderTable('districtTable', DISTRICT_COLS, rows, { onRow: (r) => openDrawer(r.district), empty: 'No district matches your search' });
  }

  function exportCsv() {
    const ds = fDistricts();
    const cols = [['division', 'Division'], ['district', 'District'], ['awcs', 'Total AWCs'], ['children', 'Active Children (0-6)'], ['pw', 'Pregnant Women'], ['lm', 'Lactating Mothers'], ['measured', 'Children Measured'],
      ...METRICS.map((m) => [m.k, m.label]), ['sam', 'SAM Cases'], ['mam', 'MAM Cases'], ['awwV', 'AWW Vacant'], ['awhV', 'AWH Vacant']];
    const q = (v) => { const s = v === null || v === undefined ? '' : String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
    const csv = [cols.map((c) => q(c[1])).join(','), ...ds.map((d) => cols.map((c) => q(d[c[0]])).join(','))].join('\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    a.download = `districts${state.division ? '-' + state.division : ''}.csv`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    toast('CSV exported', ` ${ds.length} districts`);
  }

  /* ================= NUTRITION ================= */
  function renderNutrition() {
    const ds = fDistricts();
    const A = WCD.aggregate(ds);
    renderKpis('#nutKpis', [
      { key: 'n_sam', label: 'SAM cases', icon: 'alert', value: A.sam, lowerBetter: true, sub: `${fPct(A.samPct)} of ${fInt(A.measured)} measured` },
      { key: 'n_mam', label: 'MAM cases', icon: 'drop', value: A.mam, lowerBetter: true, sub: `${fPct(A.mamPct)} of measured` },
      { key: 'n_new', label: 'New SAM this month', icon: 'clock', value: A.newSam, lowerBetter: true, sub: `${fInt(A.samRef)} referred · ${fInt(A.samFollow)} follow-ups` },
      { key: 'n_stunt', label: 'Severe stunting (avg)', icon: 'ruler', value: A.sevStunt, fmt: fPctK, lowerBetter: true, sub: `Moderate ${fPct(A.modStunt)}` },
    ]);

    const s1 = cssVar('--series-1'), s2 = cssVar('--series-2');
    $('#nutDivLegend').innerHTML = legendHTML([[s1, 'SAM %'], [s2, 'MAM %']]);
    let labels, rowsA;
    if (state.division) { labels = ds.map((d) => d.district); rowsA = ds; }
    else {
      const g = WCD.groupBy(ds, 'division');
      labels = [...g.keys()].sort(); rowsA = labels.map((k) => WCD.aggregate(g.get(k)));
    }
    upsertChart('nutDivChart', {
      type: 'bar',
      data: { labels, datasets: [
        { label: 'SAM %', data: rowsA.map((r) => r.samPct), ...barStyle(s1), borderWidth: 1, maxBarThickness: 18 },
        { label: 'MAM %', data: rowsA.map((r) => r.mamPct), ...barStyle(s2), borderWidth: 1, maxBarThickness: 18 },
      ] },
      options: {
        scales: { y: pctAxis(niceMax(rowsA.map((r) => r.mamPct))), x: { ...catAxis(), ticks: { autoSkip: false, maxRotation: 50 } } },
        interaction: { mode: 'index', intersect: false },
        plugins: { tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${fPct(c.raw)}` } } },
        onClick: (e, el) => { if (!el.length) return; const l = labels[el[0].index]; state.division ? openDrawer(l) : setDivision(l); },
      },
    });

    const top = [...ds].filter((d) => d.samPct !== null).sort((a, b) => b.samPct - a.samPct).slice(0, 10);
    upsertChart('samTopChart', {
      type: 'bar',
      data: { labels: top.map((d) => d.district), datasets: [{ label: 'SAM %', data: top.map((d) => d.samPct), ...barStyle(s2), maxBarThickness: 16 }] },
      options: {
        indexAxis: 'y',
        scales: { x: pctAxis(niceMax(top.map((d) => d.samPct))), y: catAxis() },
        plugins: { tooltip: { callbacks: { label: (c) => ` SAM ${fPct(c.raw)} · ${fInt(top[c.dataIndex].sam)} cases` } } },
        onClick: (e, el) => { if (el.length) openDrawer(top[el[0].index].district); },
      },
    });

    const pts = ds.filter((d) => d.sevStunt !== null && d.modStunt !== null);
    upsertChart('stuntChart', {
      type: 'scatter',
      data: { datasets: [{ label: 'District', data: pts.map((d) => ({ x: d.sevStunt, y: d.modStunt, d: d.district })),
        backgroundColor: s1 + 'cc', borderColor: cssVar('--surface'), borderWidth: 2, pointRadius: 6, pointHoverRadius: 9 }] },
      options: {
        scales: {
          x: { ...pctAxis(), beginAtZero: false, max: undefined, title: { display: true, text: 'Severely stunted %' } },
          y: { ...pctAxis(), beginAtZero: false, max: undefined, title: { display: true, text: 'Moderately stunted %' } },
        },
        plugins: { tooltip: { callbacks: { title: (c) => c[0].raw.d, label: (c) => ` Severe ${fPct(c.raw.x)} · Moderate ${fPct(c.raw.y)}` } } },
        onClick: (e, el) => { if (el.length) openDrawer(pts[el[0].index].district); },
      },
    });

    const rows = ds.filter((d) => matchQ(state.search.nut, d.district, d.division));
    renderTable('nutTable', [
      districtCol, intCol('measured', 'Measured'), intCol('sam', 'SAM'), burdenCol('samPct', 'SAM %'), intCol('mam', 'MAM'),
      intCol('newSam', 'New SAM'), intCol('samFollow', 'Follow-ups'), intCol('samRef', 'Referred'),
    ], rows, { onRow: (r) => openDrawer(r.district), defaultSort: { k: 'samPct', dir: -1 } });
  }

  /* ================= MATERNAL ================= */
  function renderMaternal() {
    const ds = fDistricts();
    const A = WCD.aggregate(ds);
    renderKpis('#matKpis', [
      { key: 'm_pw', label: 'Pregnant women', icon: 'heart', value: A.pw, sub: `${fInt(A.lm)} lactating mothers` },
      { key: 'm_anc', label: 'ANC coverage (avg)', icon: 'check', value: A.anc, fmt: fPctK, meter: A.anc, meterColor: ragColor(A.anc), sub: ragChip(A.anc, RAG_LABEL[rag(A.anc)]) },
      { key: 'm_an', label: 'Anaemia among PW', icon: 'drop', value: A.anaemiaCount, lowerBetter: true, sub: `${fPct(A.anaemiaPct)} of pregnant women` },
      { key: 'm_pend', label: 'Pending due ANC', icon: 'clock', value: A.pendingAnc, lowerBetter: true, sub: `Optimum weight gain ${fPct(A.gwg)}` },
    ]);
    const ramp = ['--seq-1', '--seq-2', '--seq-3', '--seq-4'].map(cssVar);
    const vals = [A.anc1, A.anc2, A.anc3, A.anc4];
    upsertChart('ancChart', {
      type: 'bar',
      data: { labels: ['ANC 1', 'ANC 2', 'ANC 3', 'ANC 4'], datasets: [{ label: 'Coverage', data: vals, ...barStyle(ramp), maxBarThickness: 64 }] },
      options: {
        scales: { y: pctAxis(1), x: catAxis() },
        plugins: { tooltip: { callbacks: { label: (c) => ` Avg coverage ${fPct(c.raw)}` } } },
      },
      plugins: [valueLabels(fPct)],
    });
    const top = [...ds].filter((d) => d.anaemiaPct !== null).sort((a, b) => b.anaemiaPct - a.anaemiaPct).slice(0, 10);
    upsertChart('anaemiaChart', {
      type: 'bar',
      data: { labels: top.map((d) => d.district), datasets: [{ label: 'Anaemia %', data: top.map((d) => d.anaemiaPct), ...barStyle(cssVar('--series-2')), maxBarThickness: 16 }] },
      options: {
        indexAxis: 'y',
        scales: { x: pctAxis(niceMax(top.map((d) => d.anaemiaPct))), y: catAxis() },
        plugins: { tooltip: { callbacks: { label: (c) => ` ${fPct(c.raw)} · ${fInt(top[c.dataIndex].anaemiaCount)} women` } } },
        onClick: (e, el) => { if (el.length) openDrawer(top[el[0].index].district); },
      },
    });
    const rows = ds.filter((d) => matchQ(state.search.mat, d.district, d.division));
    renderTable('matTable', [
      districtCol, intCol('pw', 'Pregnant women'), pctCol('anc1', 'ANC 1'), pctCol('anc2', 'ANC 2'), pctCol('anc3', 'ANC 3'), pctCol('anc4', 'ANC 4'),
      intCol('anaemiaCount', 'Anaemic'), burdenCol('anaemiaPct', 'Anaemia %'), intCol('pendingAnc', 'Pending ANC'), pctCol('gwg', 'Optimum GWG'),
    ], rows, { onRow: (r) => openDrawer(r.district) });
  }

  // Small plugin: value labels on top of bars (used for the 4-bar ANC cascade only).
  function valueLabels(fmt) {
    return {
      id: 'valueLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        ctx.save();
        ctx.fillStyle = cssVar('--text');
        ctx.font = '600 12px Inter, system-ui, sans-serif';
        ctx.textAlign = 'center';
        chart.getDatasetMeta(0).data.forEach((b, i) => {
          const v = chart.data.datasets[0].data[i];
          if (v !== null) ctx.fillText(fmt(v), b.x, b.y - 6);
        });
        ctx.restore();
      },
    };
  }

  /* ================= AWC ================= */
  const goodStatus = (s) => /^(functional|available|present|yes)$/i.test(s);
  const warnStatus = (s) => /repair|partial/i.test(s);
  const statusChip = (s) => `<span class="rag ${goodStatus(s) ? 'g' : warnStatus(s) ? 'a' : s === '—' ? 'n' : 'r'}">${esc(s)}</span>`;

  function renderAwc() {
    const base = fAwcs();
    // District select options limited to current division
    const dists = [...new Set(base.map((a) => a.district))].sort();
    if (state.awcDistrict && !dists.includes(state.awcDistrict)) state.awcDistrict = '';
    $('#awcDistrict').innerHTML = '<option value="">All districts</option>' + dists.map((d) => `<option ${d === state.awcDistrict ? 'selected' : ''}>${esc(d)}</option>`).join('');
    const blds = [...new Set(state.data.awcs.map((a) => a.building))].sort();
    $('#awcBuilding').innerHTML = '<option value="">Any building</option>' + blds.map((d) => `<option ${d === state.awcBuilding ? 'selected' : ''}>${esc(d)}</option>`).join('');

    const scoped = base.filter((a) => !state.awcDistrict || a.district === state.awcDistrict);
    const share = (k) => (scoped.length ? scoped.filter((a) => goodStatus(a[k])).length / scoped.length : null);
    const avgOp = scoped.length ? scoped.reduce((s, a) => s + (a.op || 0), 0) / scoped.length : null;
    renderKpis('#awcKpis', [
      { key: 'a_b', label: 'Functional buildings', icon: 'home', value: share('building'), fmt: fPctK, meter: share('building'), meterColor: ragColor(share('building')), sub: `${scoped.filter((a) => goodStatus(a.building)).length} of ${scoped.length} centres` },
      { key: 'a_t', label: 'Toilet available', icon: 'wc', value: share('toilet'), fmt: fPctK, meter: share('toilet'), meterColor: ragColor(share('toilet')), sub: `${scoped.filter((a) => goodStatus(a.toilet)).length} of ${scoped.length} centres` },
      { key: 'a_w', label: 'Drinking water', icon: 'water', value: share('water'), fmt: fPctK, meter: share('water'), meterColor: ragColor(share('water')), sub: `${scoped.filter((a) => goodStatus(a.water)).length} of ${scoped.length} centres` },
      { key: 'a_op', label: 'Avg operational rate', icon: 'clock', value: avgOp, fmt: fPctK, meter: avgOp, meterColor: ragColor(avgOp), sub: 'Days open ÷ 26 reporting days' },
    ]);

    const good = cssVar('--good'), warn = cssVar('--warn'), crit = cssVar('--crit');
    $('#awcLegend').innerHTML = legendHTML([[good, 'Functional / available / present'], [warn, 'Needs repair'], [crit, 'Non-functional / unavailable / vacant']]);
    const dims = [['building', 'Building'], ['toilet', 'Toilet'], ['water', 'Drinking water'], ['aww', 'AWW'], ['awh', 'AWH']];
    const cnt = dims.map(([k]) => {
      const g = scoped.filter((a) => goodStatus(a[k])).length;
      const w = scoped.filter((a) => warnStatus(a[k])).length;
      return { g, w, r: scoped.length - g - w, n: scoped.length || 1 };
    });
    const seg = (color) => ({ ...barStyle(color), borderColor: cssVar('--surface'), borderWidth: { left: 1, right: 1 }, borderRadius: 0, maxBarThickness: 22 });
    upsertChart('awcStatusChart', {
      type: 'bar',
      data: { labels: dims.map((d) => d[1]), datasets: [
        { label: 'Good', data: cnt.map((c) => c.g / c.n), raw: cnt.map((c) => c.g), ...seg(good) },
        { label: 'Needs repair', data: cnt.map((c) => c.w / c.n), raw: cnt.map((c) => c.w), ...seg(warn) },
        { label: 'Not available', data: cnt.map((c) => c.r / c.n), raw: cnt.map((c) => c.r), ...seg(crit) },
      ] },
      options: {
        indexAxis: 'y',
        scales: { x: { ...pctAxis(1), stacked: true }, y: { ...catAxis(), stacked: true } },
        plugins: { tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${fPct(c.raw)} (${c.dataset.raw[c.dataIndex]})` } } },
      },
    });

    const rows = scoped.filter((a) => (!state.awcBuilding || a.building === state.awcBuilding) && matchQ(state.search.awc, a.code, a.name, a.project, a.sector, a.district));
    $('#awcSub').textContent = `${rows.length} of ${base.length} sample centre records`;
    const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    state.awcPage = Math.min(state.awcPage, pages - 1);
    const cols = AWC_COLS;
    // sort whole set then page
    const sortedAll = sortRows('awcTable', cols, rows, { k: 'code', dir: 1 });
    const page = sortedAll.slice(state.awcPage * PAGE_SIZE, (state.awcPage + 1) * PAGE_SIZE);
    renderTable('awcTable', cols, page, { defaultSort: state.sorts.awcTable, empty: 'No centres match these filters' });
    // header clicks should re-sort whole set
    $('#awcTable').onclick = (e) => {
      const th = e.target.closest('th'); if (!th) return;
      const s = state.sorts.awcTable; const k = th.dataset.k;
      state.sorts.awcTable = { k, dir: s.k === k ? -s.dir : (cols.find((c) => c.k === k).num ? -1 : 1) };
      state.awcPage = 0; renderAwc();
    };
    $('#awcPager').innerHTML = `<span>Showing ${rows.length ? state.awcPage * PAGE_SIZE + 1 : 0}–${Math.min(rows.length, (state.awcPage + 1) * PAGE_SIZE)} of ${rows.length}</span>
      <div class="pages"><button class="btn small" data-page="-1" ${state.awcPage === 0 ? 'disabled' : ''}>← Prev</button>
      <button class="btn small" disabled>${state.awcPage + 1} / ${pages}</button>
      <button class="btn small" data-page="1" ${state.awcPage >= pages - 1 ? 'disabled' : ''}>Next →</button></div>`;
  }
  const AWC_COLS = [
    { k: 'code', label: 'AWC', cls: 'strong', fmt: (r) => `${esc(r.code)}<span class="sub">${esc(r.name)}</span>` },
    { k: 'district', label: 'District / project', fmt: (r) => `${esc(r.district)}<span class="sub">${esc(r.project)} · ${esc(r.sector)}</span>` },
    intCol('active', 'Children'), pctCol('me', 'Measured'), intCol('sam', 'SAM'), intCol('mam', 'MAM'),
    { k: 'building', label: 'Building', fmt: (r) => statusChip(r.building) },
    { k: 'toilet', label: 'Toilet', fmt: (r) => statusChip(r.toilet) },
    { k: 'water', label: 'Water', fmt: (r) => statusChip(r.water) },
    intCol('days', 'Days open'), pctCol('op', 'Op. rate'), pctCol('hvRate', 'Home visits'),
    { k: 'aww', label: 'AWW', fmt: (r) => statusChip(r.aww) },
    { k: 'awh', label: 'AWH', fmt: (r) => statusChip(r.awh) },
  ];
  function sortRows(id, columns, rows, def) {
    const s = state.sorts[id] || def; state.sorts[id] = s;
    const col = columns.find((c) => c.k === s.k) || columns[0];
    return [...rows].sort((a, b) => {
      const x = a[col.k], y = b[col.k];
      if (x === y) return 0; if (x === null) return 1; if (y === null) return -1;
      return (typeof x === 'string' ? x.localeCompare(y) : x - y) * s.dir;
    });
  }

  /* ================= ALERTS ================= */
  const prioClass = (p) => (/high/i.test(p) ? 'high' : /low/i.test(p) ? 'low' : 'medium');
  const statusClass = (s) => (/progress/i.test(s) ? 'progress' : /closed|resolved|done/i.test(s) ? 'closed' : 'open');

  function renderAlerts() {
    const al = fAlerts();
    renderKpis('#alertKpis', [
      { key: 'al_t', label: 'Total alerts', icon: 'alert', value: al.length, lowerBetter: true, sub: `${new Set(al.map((a) => a.district)).size} districts affected` },
      { key: 'al_h', label: 'High priority', icon: 'alert', value: al.filter((a) => /high/i.test(a.priority)).length, lowerBetter: true, sub: `${al.filter((a) => /medium/i.test(a.priority)).length} medium · ${al.filter((a) => /low/i.test(a.priority)).length} low` },
      { key: 'al_o', label: 'Open', icon: 'clock', value: al.filter((a) => statusClass(a.status) === 'open').length, lowerBetter: true, sub: 'Awaiting action' },
      { key: 'al_p', label: 'In progress', icon: 'check', value: al.filter((a) => statusClass(a.status) === 'progress').length, sub: `${al.filter((a) => statusClass(a.status) === 'closed').length} closed` },
    ]);
    const prios = [...new Set(state.data.alerts.map((a) => a.priority))];
    const order = { high: 0, medium: 1, low: 2 };
    prios.sort((a, b) => (order[a.toLowerCase()] ?? 9) - (order[b.toLowerCase()] ?? 9));
    $('#prioChips').innerHTML = ['', ...prios].map((p) => `<button class="chip ${p === state.alertPrio ? 'on' : ''}" data-prio="${esc(p)}">${p ? esc(p) : 'All'}</button>`).join('');
    const sts = [...new Set(state.data.alerts.map((a) => a.status))].sort();
    $('#alertStatus').innerHTML = '<option value="">Any status</option>' + sts.map((s) => `<option ${s === state.alertStatus ? 'selected' : ''}>${esc(s)}</option>`).join('');
    const cats = [...new Set(state.data.alerts.map((a) => a.category))].sort();
    $('#alertCat').innerHTML = '<option value="">Any category</option>' + cats.map((s) => `<option ${s === state.alertCat ? 'selected' : ''}>${esc(s)}</option>`).join('');

    const rows = al.filter((a) => (!state.alertPrio || a.priority === state.alertPrio) && (!state.alertStatus || a.status === state.alertStatus) && (!state.alertCat || a.category === state.alertCat))
      .sort((a, b) => (order[a.priority.toLowerCase()] ?? 9) - (order[b.priority.toLowerCase()] ?? 9) || a.district.localeCompare(b.district));
    $('#alertSub').textContent = `${rows.length} of ${al.length} alerts${state.division ? ' in ' + state.division : ''}`;
    $('#alertList').innerHTML = rows.length ? rows.map((a, i) => alertCard(a, i)).join('') : '<p class="hint">No alerts match these filters.</p>';
  }
  const alertCard = (a, i) => `
    <div class="alert-card ${prioClass(a.priority)}" style="--i:${i}" data-district="${esc(a.district)}">
      <div class="ac-top"><h4>${esc(a.district)}</h4><span class="rag ${prioClass(a.priority) === 'high' ? 'r' : prioClass(a.priority) === 'low' ? 'g' : 'a'}">${esc(a.priority)}</span></div>
      <div class="ac-cat">${esc(a.category)}${a.division ? ' · ' + esc(a.division) : ''}</div>
      <div class="ac-issue">${esc(a.issue)}</div>
      <div class="ac-foot"><span>${esc(a.officer)}</span><span class="status ${statusClass(a.status)}">${esc(a.status)}</span></div>
    </div>`;

  /* ================= MAP ================= */
  const MAP_METRICS = [
    ...INDICATORS.map((m) => ({ ...m, rag: true })),
    ...BURDEN,
    { k: 'awcs', label: 'Total AWCs', count: true },
    { k: 'children', label: 'Active Children (0–6)', count: true },
    { k: 'sam', label: 'SAM Cases', count: true, lowerBetter: true },
    { k: 'alertsN', label: 'Active Alerts', count: true, lowerBetter: true },
  ];
  const mapState = { metric: 'me', selected: null, built: false, vb: null, full: null, labels: false, anim: 0 };
  const keyName = (s) => String(s || '').toLowerCase().replace(/[^a-z]/g, '');
  const mapFmt = (m, v) => (m.count ? fInt(v) : fPct(v));

  function mapRows() {
    const active = state.data.alerts.filter((a) => !/closed|resolved|done/i.test(a.status));
    const cnt = new Map();
    active.forEach((a) => cnt.set(a.district, (cnt.get(a.district) || 0) + 1));
    return state.data.districts.map((d) => ({ ...d, alertsN: cnt.get(d.district) || 0 }));
  }

  // Five quantile classes over ALL districts, so a district's colour doesn't change with the division filter.
  function mapScale(m, rows) {
    if (m.rag) return { type: 'rag', color: (v) => (v === null || v === undefined ? 'var(--map-empty)' : `var(--${{ g: 'good', a: 'warn', r: 'crit' }[rag(v)]})`) };
    const vals = rows.map((r) => r[m.k]).filter((v) => v !== null && v !== undefined && isFinite(v)).sort((a, b) => a - b);
    const q = (p) => vals[Math.min(vals.length - 1, Math.floor(p * vals.length))];
    const breaks = [q(0.2), q(0.4), q(0.6), q(0.8)];
    const cls = (v) => breaks.filter((b) => v >= b).length;
    return {
      type: 'seq', breaks, min: vals[0], max: vals[vals.length - 1],
      color: (v) => (v === null || v === undefined || !isFinite(v) ? 'var(--map-empty)' : `var(--map-${cls(v) + 1})`),
    };
  }

  function buildMap() {
    const M = window.MP_MAP;
    const svg = $('#mapSvg');
    if (!M) { $('#mapStage').innerHTML = '<p class="hint" style="padding:20px">Map data (js/mp-map.js) not found.</p>'; return; }
    mapState.full = { x: -10, y: -10, w: M.width + 20, h: M.height + 20 };
    mapState.vb = { ...mapState.full };
    const ns = 'http://www.w3.org/2000/svg';
    const minX = Math.min(...M.districts.map((d) => d.cx)), maxX = Math.max(...M.districts.map((d) => d.cx));
    let html = '<g id="mapShapes">';
    M.districts.forEach((d) => {
      const delay = Math.round(((d.cx - minX) / (maxX - minX)) * 600);
      html += `<path class="d" d="${d.d}" data-name="${esc(d.name)}" tabindex="0" role="button" aria-label="${esc(d.name)}" style="--d:${delay}"></path>`;
    });
    html += '</g><path class="map-hi" id="mapHi" style="opacity:0"></path><path class="map-sel" id="mapSel" style="opacity:0"></path><g id="mapPts">';
    M.points.forEach((p) => {
      html += `<g class="pt" data-name="${esc(p.name)}" tabindex="0" role="button" aria-label="${esc(p.name)} (new district)">
        <circle class="map-pt-pulse" cx="${p.cx}" cy="${p.cy}" r="12" style="opacity:0"></circle>
        <circle class="map-pt" cx="${p.cx}" cy="${p.cy}" r="8"></circle>
        <circle class="map-pt-ring" cx="${p.cx}" cy="${p.cy}" r="11.5"></circle></g>`;
    });
    html += '</g><g id="mapLabelsG">';
    M.districts.forEach((d) => { html += `<text class="map-label" x="${d.cx}" y="${d.cy}">${esc(d.name)}</text>`; });
    M.points.forEach((p) => { html += `<text class="map-label pt-label" x="${p.cx}" y="${p.cy + 20}" data-cy="${p.cy}">${esc(p.name)}</text>`; });
    html += '</g>';
    svg.innerHTML = html;
    svg.setAttribute('xmlns', ns);
    setViewBox(mapState.vb);
    $('#mapSource').textContent = M.source + ' Boundaries are indicative.';
    wireMap();
    mapState.built = true;
  }

  // Keep labels and markers a constant on-screen size while zooming.
  function setViewBox(v) {
    const svg = $('#mapSvg');
    svg.setAttribute('viewBox', `${v.x} ${v.y} ${v.w} ${v.h}`);
    const k = v.w / mapState.full.w;
    svg.style.setProperty('--k', k.toFixed(4));
    $$('#mapPts .pt', svg).forEach((g) => {
      const [pulse, dot, ring] = g.querySelectorAll('circle');
      pulse.setAttribute('r', 12 * k); dot.setAttribute('r', 8 * k); ring.setAttribute('r', 11.5 * k);
    });
    $$('#mapLabelsG text.pt-label', svg).forEach((t) => t.setAttribute('y', +t.dataset.cy + 20 * k));
  }
  function animateViewBox(to, dur = 650) {
    const from = { ...mapState.vb }, t0 = performance.now(), id = ++mapState.anim;
    const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
    const step = (t) => {
      if (id !== mapState.anim) return;
      const p = reduce ? 1 : Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 3);
      mapState.vb = { x: from.x + (to.x - from.x) * e, y: from.y + (to.y - from.y) * e, w: from.w + (to.w - from.w) * e, h: from.h + (to.h - from.h) * e };
      setViewBox(mapState.vb);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }
  function fitTo(els) {
    if (!els.length) return animateViewBox(mapState.full);
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    els.forEach((el) => { const b = el.getBBox(); x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y); x1 = Math.max(x1, b.x + b.width); y1 = Math.max(y1, b.y + b.height); });
    const pad = 30, F = mapState.full, ar = F.w / F.h;
    let w = x1 - x0 + pad * 2, h = y1 - y0 + pad * 2;
    if (w / h > ar) h = w / ar; else w = h * ar;
    animateViewBox({ x: (x0 + x1) / 2 - w / 2, y: (y0 + y1) / 2 - h / 2, w, h });
  }
  function zoomBy(f, cx, cy) {
    const v = mapState.vb, F = mapState.full;
    const w = Math.min(F.w, Math.max(F.w / 8, v.w / f)), h = w * (v.h / v.w);
    cx = cx ?? v.x + v.w / 2; cy = cy ?? v.y + v.h / 2;
    const to = { x: cx - (cx - v.x) * (w / v.w), y: cy - (cy - v.y) * (h / v.h), w, h };
    animateViewBox(to, 300);
  }

  function districtOfMapName(name) {
    const k = keyName(name);
    return mapRows().find((d) => keyName(d.district) === k) || null;
  }

  function wireMap() {
    const stage = $('#mapStage'), svg = $('#mapSvg'), tip = $('#mapTip');
    const toSvg = (cx, cy) => { const r = svg.getBoundingClientRect(); const v = mapState.vb;
      const s = Math.max(v.w / r.width, v.h / r.height); const ox = (r.width - v.w / s) / 2, oy = (r.height - v.h / s) / 2;
      return { x: v.x + (cx - r.left - ox) * s, y: v.y + (cy - r.top - oy) * s, s }; };
    const targetName = (el) => el && el.closest && (el.closest('.d') || el.closest('.pt'))?.dataset.name;

    // Hover tooltip (mouse/pen only)
    stage.addEventListener('pointermove', (e) => {
      if (e.pointerType === 'touch' || drag) return;
      const name = targetName(e.target);
      if (!name) { hideHover(); return; }
      showHover(name);
      const r = stage.getBoundingClientRect();
      tip.style.left = Math.min(r.width - 100, Math.max(100, e.clientX - r.left)) + 'px';
      tip.style.top = Math.max(110, e.clientY - r.top) + 'px';
    });
    stage.addEventListener('pointerleave', hideHover);

    // Pan / pinch / click
    const ptrs = new Map(); let drag = null, moved = 0, pinch = null;
    stage.addEventListener('pointerdown', (e) => {
      if (e.target.closest('.map-zoom')) return;
      ptrs.set(e.pointerId, { x: e.clientX, y: e.clientY });
      stage.setPointerCapture(e.pointerId);
      moved = 0;
      if (ptrs.size === 1) drag = { x: e.clientX, y: e.clientY, vb: { ...mapState.vb }, name: targetName(e.target) };
      if (ptrs.size === 2) { const [a, b] = [...ptrs.values()]; pinch = { d: Math.hypot(a.x - b.x, a.y - b.y), vb: { ...mapState.vb } }; drag = null; }
    });
    stage.addEventListener('pointermove', (e) => {
      if (!ptrs.has(e.pointerId)) return;
      ptrs.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pinch && ptrs.size === 2) {
        const [a, b] = [...ptrs.values()]; const d = Math.hypot(a.x - b.x, a.y - b.y);
        const mid = toSvg((a.x + b.x) / 2, (a.y + b.y) / 2); const f = d / pinch.d;
        const F = mapState.full; const w = Math.min(F.w, Math.max(F.w / 8, pinch.vb.w / f)); const h = w * (pinch.vb.h / pinch.vb.w);
        mapState.anim++; mapState.vb = { x: mid.x - (mid.x - mapState.vb.x) * (w / mapState.vb.w), y: mid.y - (mid.y - mapState.vb.y) * (h / mapState.vb.h), w, h };
        setViewBox(mapState.vb); moved = 99; return;
      }
      if (!drag) return;
      const dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      moved = Math.max(moved, Math.hypot(dx, dy));
      if (moved > 6 && mapState.vb.w < mapState.full.w - 1) {
        stage.classList.add('dragging'); hideHover();
        const s = toSvg(0, 0).s; mapState.anim++;
        mapState.vb = { ...mapState.vb, x: drag.vb.x - dx * s, y: drag.vb.y - dy * s };
        setViewBox(mapState.vb);
      }
    });
    const end = (e) => {
      if (!ptrs.has(e.pointerId)) return;
      ptrs.delete(e.pointerId);
      if (ptrs.size < 2) pinch = null;
      stage.classList.remove('dragging');
      if (drag && moved < 6 && ptrs.size === 0) {
        if (drag.name) selectMapDistrict(drag.name, { scroll: e.pointerType === 'touch' });
        else selectMapDistrict(null);
      }
      if (ptrs.size === 0) drag = null;
    };
    stage.addEventListener('pointerup', end);
    stage.addEventListener('pointercancel', end);
    stage.addEventListener('wheel', (e) => {
      e.preventDefault();
      const p = toSvg(e.clientX, e.clientY);
      zoomBy(e.deltaY < 0 ? 1.35 : 1 / 1.35, p.x, p.y);
    }, { passive: false });
    stage.addEventListener('dblclick', (e) => { const p = toSvg(e.clientX, e.clientY); zoomBy(1.8, p.x, p.y); });

    svg.addEventListener('keydown', (e) => {
      const name = targetName(e.target);
      if (name && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); selectMapDistrict(name); }
    });
    svg.addEventListener('focusin', (e) => { const n = targetName(e.target); if (n) showHover(n); });

    $('.map-zoom').addEventListener('click', (e) => {
      const b = e.target.closest('[data-zoom]'); if (!b) return;
      if (b.dataset.zoom === 'in') zoomBy(1.5); else if (b.dataset.zoom === 'out') zoomBy(1 / 1.5); else fitDivision();
    });
  }

  function shapeFor(name) {
    return $(`#mapShapes .d[data-name="${CSS.escape(name)}"]`) || $(`#mapPts .pt[data-name="${CSS.escape(name)}"]`);
  }
  function outline(id, name) {
    const el = $(id), sh = name && shapeFor(name);
    if (!sh) { el.style.opacity = 0; return; }
    if (sh.classList.contains('pt')) { const c = $('.map-pt', sh); const cx = +c.getAttribute('cx'), cy = +c.getAttribute('cy');
      el.setAttribute('d', `M${cx - 10},${cy}a10,10 0 1,0 20,0a10,10 0 1,0 -20,0`); }
    else el.setAttribute('d', sh.getAttribute('d'));
    el.style.opacity = 1;
  }

  function showHover(name) {
    const d = districtOfMapName(name); if (!d) return;
    const m = MAP_METRICS.find((x) => x.k === mapState.metric);
    const rows = mapRows().filter((r) => r[m.k] !== null && r[m.k] !== undefined);
    const rank = [...rows].sort((a, b) => (m.lowerBetter ? a[m.k] - b[m.k] : b[m.k] - a[m.k])).findIndex((r) => r.district === d.district) + 1;
    const pt = (window.MP_MAP.points || []).find((p) => keyName(p.name) === keyName(name));
    $('#mapTip').innerHTML = `<b>${esc(d.district)}</b><small>${esc(d.division)} division${pt ? ' · new district' : ''}</small>
      <div class="tv">${mapFmt(m, d[m.k])}</div><small>${esc(m.label)} · rank ${rank} of ${rows.length}${m.rag ? ' · ' + RAG_LABEL[rag(d[m.k])] : ''}</small>`;
    $('#mapTip').classList.add('on');
    outline('#mapHi', name);
  }
  function hideHover() { $('#mapTip').classList.remove('on'); $('#mapHi').style.opacity = 0; }

  function selectMapDistrict(name, { scroll = false, zoom = false } = {}) {
    const d = name && districtOfMapName(name);
    mapState.selected = d ? d.district : null;
    $$('#mapPts .map-pt-pulse').forEach((c) => (c.style.opacity = 0));
    if (d && shapeFor(name)?.classList.contains('pt')) $('.map-pt-pulse', shapeFor(name)).style.opacity = 1;
    outline('#mapSel', d ? name : null);
    if (d && zoom) fitTo([shapeFor(name)]);
    renderMapPanel();
    if (scroll && d) setTimeout(() => $('#mapPanel').scrollIntoView({ behavior: 'smooth', block: 'start' }), 60);
  }

  function fitDivision() {
    if (!state.division) return animateViewBox(mapState.full);
    const names = new Set(state.data.districts.filter((d) => d.division === state.division).map((d) => keyName(d.district)));
    fitTo($$('#mapShapes .d').filter((p) => names.has(keyName(p.dataset.name))));
  }

  function renderMap() {
    if (!window.MP_MAP) { buildMap(); return; }
    if (!mapState.built) { buildMap(); if (state.division) setTimeout(fitDivision, 60); }
    const m = MAP_METRICS.find((x) => x.k === mapState.metric);
    const rows = mapRows();
    const byKey = new Map(rows.map((r) => [keyName(r.district), r]));
    const scale = mapScale(m, rows);
    const inDiv = (r) => !state.division || (r && r.division === state.division);
    hideHover();
    if (state.division && mapState.selected) {
      const sel = rows.find((r) => r.district === mapState.selected);
      if (!sel || sel.division !== state.division) selectMapDistrict(null);
    }

    $$('#mapShapes .d').forEach((p) => {
      const r = byKey.get(keyName(p.dataset.name));
      p.style.fill = scale.color(r ? r[m.k] : null);
      p.classList.toggle('dim', !inDiv(r));
    });
    $$('#mapPts .pt').forEach((g) => {
      const r = byKey.get(keyName(g.dataset.name));
      $('.map-pt', g).style.fill = scale.color(r ? r[m.k] : null);
      g.style.opacity = inDiv(r) ? 1 : 0.2;
      g.style.display = r ? '' : 'none';
    });
    $('#mapStage').classList.toggle('labels-on', mapState.labels);

    const mapped = new Set([...window.MP_MAP.districts, ...window.MP_MAP.points].map((x) => keyName(x.name)));
    const missing = rows.filter((r) => !mapped.has(keyName(r.district))).map((r) => r.district);
    $('#mapSub').textContent = `${m.label} · ${state.division ? state.division + ' division' : 'all ' + rows.length + ' districts'}` +
      (missing.length ? ` · not on map: ${missing.join(', ')}` : '');
    $('#mapDistricts').innerHTML = rows.map((r) => `<option value="${esc(r.district)}">`).join('');

    // Legend
    const newKey = '<span class="key"><i class="dot"></i>New district (2023), shown at HQ</span>';
    if (scale.type === 'rag') {
      $('#mapLegend').innerHTML = `<span class="key"><i style="background:var(--good)"></i>On track ≥90%</span><span class="key"><i style="background:var(--warn)"></i>Watch 75–90%</span><span class="key"><i style="background:var(--crit)"></i>Off track &lt;75%</span>${newKey}`;
    } else {
      const b = [scale.min, ...scale.breaks, scale.max];
      $('#mapLegend').innerHTML = `<span>${esc(m.label)}${m.lowerBetter ? ' (darker = higher burden)' : ' (darker = higher)'}</span><div class="scale">${[1, 2, 3, 4, 5].map((i) => `<span><i style="background:var(--map-${i})"></i>${mapFmt(m, b[i - 1])}${i === 5 ? '–' + mapFmt(m, b[5]) : '+'}</span>`).join('')}</div>${newKey}`;
    }
    if (mapState.selected && !rows.some((r) => r.district === mapState.selected)) mapState.selected = null;
    renderMapPanel();
  }

  function renderMapPanel() {
    const m = MAP_METRICS.find((x) => x.k === mapState.metric);
    const rows = mapRows();
    const scope = rows.filter((r) => !state.division || r.division === state.division);
    const valid = scope.filter((r) => r[m.k] !== null && r[m.k] !== undefined);
    const sorted = [...valid].sort((a, b) => (m.lowerBetter ? a[m.k] - b[m.k] : b[m.k] - a[m.k]));
    const panel = $('#mapPanel');
    const d = mapState.selected && rows.find((r) => r.district === mapState.selected);

    if (!d) {
      const A = WCD.aggregate(scope);
      const total = m.count ? valid.reduce((s, r) => s + r[m.k], 0) : null;
      const headline = m.count ? total : A[m.k];
      const li = (r, i) => `<div class="mp-bar" data-pick="${esc(r.district)}"><span>${i + 1}. ${esc(r.district)}</span><b>${mapFmt(m, r[m.k])}</b></div>`;
      const ragCounts = m.rag ? ['g', 'a', 'r'].map((c) => valid.filter((r) => rag(r[m.k]) === c).length) : null;
      panel.innerHTML = `
        <div class="mp-card"><div class="mp-kicker">${state.division ? esc(state.division) + ' division' : 'Madhya Pradesh'}</div>
          <div class="mp-title">${esc(m.label)}</div>
          <div class="mp-value">${mapFmt(m, headline)}</div>
          <div class="mp-meta">${m.count ? 'Total across ' + valid.length + ' districts' : (m.rag ? ragChip(headline, RAG_LABEL[rag(headline)]) + ' ' : '') + 'Average across ' + valid.length + ' districts'}</div>
          ${ragCounts ? `<div class="mp-facts"><div><b style="color:var(--good-ink)">${ragCounts[0]}</b><span>On track</span></div><div><b style="color:var(--warn-ink)">${ragCounts[1]}</b><span>Watch</span></div><div><b style="color:var(--crit-ink)">${ragCounts[2]}</b><span>Off track</span></div><div><b>${scope.length - valid.length}</b><span>No data</span></div></div>` : ''}
          <div class="mp-note">Tap a district on the map, or pick one below.</div></div>
        <div class="mp-card"><h4>${m.lowerBetter ? 'Lowest' : 'Highest'} ${esc(m.label)}</h4><div class="mp-bars">${sorted.slice(0, 5).map(li).join('')}</div></div>
        <div class="mp-card"><h4>${m.lowerBetter ? 'Highest' : 'Lowest'} ${esc(m.label)}</h4><div class="mp-bars">${[...sorted].reverse().slice(0, 5).map(li).join('')}</div></div>`;
      return;
    }

    const rank = sorted.findIndex((r) => r.district === d.district) + 1;
    const pt = (window.MP_MAP.points || []).find((p) => keyName(p.name) === keyName(d.district));
    const S = WCD.aggregate(rows);
    panel.innerHTML = `
      <div class="mp-card"><div class="mp-kicker">${esc(d.division)} division</div>
        <div class="mp-title">${esc(d.district)} ${m.rag ? ragChip(d[m.k], RAG_LABEL[rag(d[m.k])]) : ''}</div>
        <div class="mp-value">${mapFmt(m, d[m.k])}</div>
        <div class="mp-meta">${esc(m.label)}${rank ? ` · rank ${rank} of ${valid.length}${state.division ? ' in division' : ''}` : ''}${!m.count && S[m.k] !== undefined ? ' · state ' + fPct(S[m.k]) : ''}</div>
        ${pt ? `<div class="mp-note">New district (formed ${esc(pt.formed)}, from ${esc(pt.parent)}). Its boundary isn't in the map source yet, so it is shown as a marker at its headquarters.</div>` : ''}
        <div class="mp-facts">
          <div><b>${fInt(d.awcs)}</b><span>AWCs</span></div><div><b>${compact(d.children)}</b><span>Children 0–6</span></div>
          <div><b>${fInt(d.sam)}</b><span>SAM (${fPct(d.samPct)})</span></div><div><b>${d.alertsN}</b><span>Active alerts</span></div>
        </div>
        <button class="btn primary" data-open="${esc(d.district)}">Open full district profile →</button>
      </div>
      <div class="mp-card"><h4>All coverage indicators</h4><div class="mp-bars">
        ${INDICATORS.map((x) => `<div class="mp-bar ${x.k === m.k ? 'on' : ''}" data-metric="${x.k}"><span>${esc(x.label)}</span>${ragChip(d[x.k])}<div class="track"><i data-w="${(d[x.k] || 0) * 100}" style="background:${ragColor(d[x.k])}"></i></div></div>`).join('')}
      </div><div class="compare-note" style="margin-top:8px">Tap an indicator to recolour the map</div></div>`;
    requestAnimationFrame(() => setTimeout(() => $$('#mapPanel .track i').forEach((i) => (i.style.width = i.dataset.w + '%')), 30));
  }

  function setMapMetric(k) {
    mapState.metric = k;
    $('#mapMetric').value = k;
    renderMap();
  }

  const RENDER = { overview: renderOverview, map: renderMap, districts: renderDistricts, nutrition: renderNutrition, maternal: renderMaternal, awc: renderAwc, alerts: renderAlerts };

  /* ================= Drawer ================= */
  function openDrawer(name, refresh = false) {
    const d = state.data.districts.find((x) => x.district === name);
    if (!d) { closeDrawer(); return; }
    state.drawerDistrict = name;
    const S = WCD.aggregate(state.data.districts);
    const awcs = state.data.awcs.filter((a) => a.district === name);
    const alerts = state.data.alerts.filter((a) => a.district === name);
    $('#drawerKicker').textContent = `${d.division} division`;
    $('#drawerTitle').textContent = d.district;
    const rank = (k) => {
      const arr = state.data.districts.filter((x) => x[k] !== null).sort((a, b) => b[k] - a[k]);
      return arr.findIndex((x) => x.district === name) + 1;
    };
    const anc = [d.anc1, d.anc2, d.anc3, d.anc4];
    $('#drawerBody').innerHTML = `
      <div class="card" style="--i:0"><div class="mini-kpis">
        <div><b>${fInt(d.awcs)}</b><span>AWCs</span></div>
        <div><b>${compact(d.children)}</b><span>Children 0–6</span></div>
        <div><b>${fInt(d.pw)}</b><span>Pregnant women</span></div>
        <div><b>${fInt(d.sam)}</b><span>SAM (${fPct(d.samPct)})</span></div>
        <div><b>${fInt(d.mam)}</b><span>MAM (${fPct(d.mamPct)})</span></div>
        <div><b>${fInt(d.awwV)} / ${fInt(d.awhV)}</b><span>AWW / AWH vacant</span></div>
      </div></div>
      <div class="card" style="--i:1">
        <header class="card-head"><div><h3>Coverage indicators</h3><p>District value vs state average · rank of ${state.data.districts.length}</p></div></header>
        <div class="compare">${INDICATORS.map((m) => `
          <div class="compare-row"><span>${esc(m.label)} <span class="sub" style="display:inline;color:var(--muted);font-size:11px">#${rank(m.k)}</span></span>${ragChip(d[m.k])}
            <div class="track"><i data-w="${(d[m.k] || 0) * 100}" style="background:${ragColor(d[m.k])}"></i><u style="left:calc(${(S[m.k] || 0) * 100}% - 1px)" title="State ${fPct(S[m.k])}"></u></div></div>`).join('')}
        </div>
        <div class="compare-note"><u></u> State average</div>
      </div>
      <div class="card" style="--i:2">
        <header class="card-head"><div><h3>Maternal health</h3><p>ANC cascade, anaemia &amp; weight gain</p></div></header>
        <div class="compare">${anc.map((v, i) => `
          <div class="compare-row"><span>ANC ${i + 1}</span>${ragChip(v)}
          <div class="track"><i data-w="${(v || 0) * 100}" style="background:${ragColor(v)}"></i><u style="left:calc(${(S['anc' + (i + 1)] || 0) * 100}% - 1px)"></u></div></div>`).join('')}
        </div>
        <div class="mini-kpis" style="margin-top:12px">
          <div><b>${fPct(d.anaemiaPct)}</b><span>Anaemia (${fInt(d.anaemiaCount)})</span></div>
          <div><b>${fInt(d.pendingAnc)}</b><span>Pending due ANC</span></div>
          <div><b>${fPct(d.gwg)}</b><span>Optimum GWG</span></div>
        </div>
      </div>
      <div class="card" style="--i:3">
        <header class="card-head"><div><h3>Nutrition</h3><p>Stunting &amp; SAM management</p></div></header>
        <div class="mini-kpis">
          <div><b>${fPct(d.sevStunt)}</b><span>Severely stunted</span></div>
          <div><b>${fPct(d.modStunt)}</b><span>Moderately stunted</span></div>
          <div><b>${fInt(d.newSam)}</b><span>New SAM this month</span></div>
          <div><b>${fInt(d.samFollow)}</b><span>SAM follow-ups</span></div>
          <div><b>${fInt(d.samRef)}</b><span>SAM referred</span></div>
          <div><b>${fInt(d.measured)}</b><span>Children measured</span></div>
        </div>
      </div>
      <div class="card" style="--i:4">
        <header class="card-head"><div><h3>Alerts</h3><p>${alerts.length} for this district</p></div></header>
        <div class="alert-list">${alerts.length ? alerts.map(alertCard).join('') : '<p class="hint">No alerts for this district.</p>'}</div>
      </div>
      <div class="card" style="--i:5">
        <header class="card-head"><div><h3>Anganwadi centres</h3><p>${awcs.length} sample records</p></div></header>
        <div class="table-wrap short"><table class="table" id="drawerAwc"></table></div>
      </div>`;
    renderTable('drawerAwc', [AWC_COLS[0], intCol('active', 'Children'), pctCol('me', 'Measured'), intCol('sam', 'SAM'),
      { k: 'building', label: 'Building', fmt: (r) => statusChip(r.building) }, pctCol('op', 'Op. rate'), { k: 'aww', label: 'AWW', fmt: (r) => statusChip(r.aww) }], awcs);
    $$('#drawerBody .alert-card').forEach((c) => c.removeAttribute('data-district'));
    if (refresh) $$('#drawerBody .card').forEach((c) => (c.style.animation = 'none'));
    $('#drawer').classList.add('on'); $('#scrim').classList.add('on');
    $('#drawer').setAttribute('aria-hidden', 'false');
    requestAnimationFrame(() => setTimeout(() => $$('#drawerBody .track i').forEach((i) => (i.style.width = i.dataset.w + '%')), refresh ? 0 : 200));
  }
  function closeDrawer() {
    state.drawerDistrict = null;
    $('#drawer').classList.remove('on'); $('#scrim').classList.remove('on');
    $('#drawer').setAttribute('aria-hidden', 'true');
  }

  /* ================= Modal ================= */
  const openModal = () => { $('#sourceModal').classList.add('on'); $('#sourceModal').setAttribute('aria-hidden', 'false'); };
  const closeModal = () => { $('#sourceModal').classList.remove('on'); $('#sourceModal').setAttribute('aria-hidden', 'true'); };

  /* ================= Theme ================= */
  function applyTheme(t) {
    if (t) document.documentElement.setAttribute('data-theme', t);
    else document.documentElement.removeAttribute('data-theme');
    $('meta[name=theme-color]').content = cssVar('--header');
    chartTheme();
    destroyCharts();
    if (state.data) RENDER[state.view]();
  }
  function currentTheme() {
    return document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  }

  function setDivision(v) {
    state.division = v;
    $('#divisionFilter').value = v;
    state.awcPage = 0;
    renderBadge();
    RENDER[state.view]();
    if (mapState.built) fitDivision();
  }

  /* ================= Wire up ================= */
  function init() {
    try { const t = localStorage.getItem('wcd-theme'); if (t) document.documentElement.setAttribute('data-theme', t); } catch (e) { /* ignore */ }
    chartTheme();
    $('#pollSecs').textContent = SERVER_POLL_MS / 1000;
    $('#srcPath').textContent = DATA_URL;

    const metricOpts = (sel) => METRICS.map((m) => `<option value="${m.k}" ${m.k === sel ? 'selected' : ''}>${esc(m.label)}</option>`).join('');
    $('#divMetric').innerHTML = metricOpts(state.divMetric);
    $('#rankMetric').innerHTML = metricOpts(state.rankMetric);
    $('#mapMetric').innerHTML = MAP_METRICS.map((m) => `<option value="${m.k}" ${m.k === mapState.metric ? 'selected' : ''}>${esc(m.label)}</option>`).join('');
    mapState.labels = matchMedia('(min-width: 1100px)').matches;
    $('#mapLabels').classList.toggle('on', mapState.labels);
    $('#mapLabels').setAttribute('aria-pressed', mapState.labels);
    $('#mapMetric').addEventListener('change', (e) => setMapMetric(e.target.value));
    $('#mapLabels').addEventListener('click', (e) => {
      mapState.labels = !mapState.labels;
      e.currentTarget.classList.toggle('on', mapState.labels);
      e.currentTarget.setAttribute('aria-pressed', mapState.labels);
      $('#mapStage').classList.toggle('labels-on', mapState.labels);
    });
    $('#mapSearch').addEventListener('change', (e) => {
      const d = state.data && state.data.districts.find((x) => keyName(x.district) === keyName(e.target.value));
      if (!d) { if (e.target.value) toast('District not found', ' ' + e.target.value, 'err', 2500); return; }
      if (state.division && d.division !== state.division) setDivision('');
      const shape = shapeFor((window.MP_MAP.districts.concat(window.MP_MAP.points).find((x) => keyName(x.name) === keyName(d.district)) || {}).name || '');
      selectMapDistrict(shape ? shape.dataset.name : d.district, { zoom: true });
      e.target.value = '';
      e.target.blur();
    });
    $('#mapPanel').addEventListener('click', (e) => {
      const pick = e.target.closest('[data-pick]');
      if (pick) {
        const n = pick.dataset.pick;
        const shape = shapeFor((window.MP_MAP.districts.concat(window.MP_MAP.points).find((x) => keyName(x.name) === keyName(n)) || {}).name || '');
        selectMapDistrict(shape ? shape.dataset.name : n);
        return;
      }
      const met = e.target.closest('[data-metric]');
      if (met) { setMapMetric(met.dataset.metric); return; }
      const op = e.target.closest('[data-open]');
      if (op) openDrawer(op.dataset.open);
    });

    const hash = location.hash.slice(1);
    if (RENDER[hash]) state.view = hash;
    $$('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.view === state.view));

    $('#tabs').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) { showView(b.dataset.view); window.scrollTo({ top: 0, behavior: 'smooth' }); } });
    $('#divisionFilter').addEventListener('change', (e) => setDivision(e.target.value));
    $('#divMetric').addEventListener('change', (e) => { state.divMetric = e.target.value; renderOverview(); });
    $('#rankMetric').addEventListener('change', (e) => { state.rankMetric = e.target.value; renderDistricts(); });
    $('#districtSearch').addEventListener('input', debounce((e) => { state.search.district = e.target.value; renderDistricts(); }));
    $('#nutSearch').addEventListener('input', debounce((e) => { state.search.nut = e.target.value; renderNutrition(); }));
    $('#matSearch').addEventListener('input', debounce((e) => { state.search.mat = e.target.value; renderMaternal(); }));
    $('#awcSearch').addEventListener('input', debounce((e) => { state.search.awc = e.target.value; state.awcPage = 0; renderAwc(); }));
    $('#awcDistrict').addEventListener('change', (e) => { state.awcDistrict = e.target.value; state.awcPage = 0; renderAwc(); });
    $('#awcBuilding').addEventListener('change', (e) => { state.awcBuilding = e.target.value; state.awcPage = 0; renderAwc(); });
    $('#awcPager').addEventListener('click', (e) => { const b = e.target.closest('[data-page]'); if (b) { state.awcPage += +b.dataset.page; renderAwc(); } });
    $('#prioChips').addEventListener('click', (e) => { const b = e.target.closest('.chip'); if (b) { state.alertPrio = b.dataset.prio; renderAlerts(); } });
    $('#alertStatus').addEventListener('change', (e) => { state.alertStatus = e.target.value; renderAlerts(); });
    $('#alertCat').addEventListener('change', (e) => { state.alertCat = e.target.value; renderAlerts(); });
    $('#exportCsv').addEventListener('click', exportCsv);

    document.addEventListener('click', (e) => {
      const go = e.target.closest('[data-goto]');
      if (go) { showView(go.dataset.goto); return; }
      const cat = e.target.closest('.as-row[data-cat]');
      if (cat) { state.alertCat = cat.dataset.cat; state.alertPrio = ''; state.alertStatus = ''; showView('alerts'); return; }
      const dl = e.target.closest('[data-district]');
      if (dl && !e.target.closest('.drawer')) { openDrawer(dl.dataset.district); return; }
      const act = e.target.closest('[data-action]');
      if (act) {
        const a = act.dataset.action;
        if (a === 'pick') pickLocalFile();
        if (a === 'retry') { $('#errorView').classList.remove('active'); $('#skeleton').classList.add('active'); loadServer({ force: true }); }
        if (a === 'server') useServer();
      }
      if (e.target.closest('[data-close]') || e.target === $('#sourceModal')) closeModal();
    });

    $('#drawerClose').addEventListener('click', closeDrawer);
    $('#scrim').addEventListener('click', closeDrawer);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { closeDrawer(); closeModal(); } });

    $('#refreshBtn').addEventListener('click', async () => {
      const b = $('#refreshBtn'); b.classList.add('spin');
      let changed = false;
      if (state.source.mode === 'local') { state.source.lastMod = null; await pollLocal(); }
      else if (state.source.mode === 'server') changed = await loadServer({ force: true });
      else toast('Uploaded copy', ' Re-upload the file or switch to a live source to refresh.', 'info');
      setTimeout(() => b.classList.remove('spin'), 500);
      if (state.source.mode === 'server' && !changed && state.data) toast('Already up to date', ' No changes in the workbook.', 'info', 2500);
    });
    $('#sourceBtn').addEventListener('click', openModal);
    $('#themeBtn').addEventListener('click', () => {
      const t = currentTheme() === 'dark' ? 'light' : 'dark';
      try { localStorage.setItem('wcd-theme', t); } catch (e) { /* ignore */ }
      applyTheme(t);
    });
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (!document.documentElement.getAttribute('data-theme')) applyTheme(null); });

    // File input + drag & drop
    $('#fileInput').addEventListener('change', (e) => { loadUploadedFile(e.target.files[0]); e.target.value = ''; });
    let dragDepth = 0;
    window.addEventListener('dragenter', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) { dragDepth++; $('#dropOverlay').classList.add('on'); } });
    window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('#dropOverlay').classList.remove('on'); } });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault(); dragDepth = 0; $('#dropOverlay').classList.remove('on');
      const f = e.dataTransfer.files[0]; if (f) loadUploadedFile(f);
    });

    document.addEventListener('visibilitychange', () => {
      if (!document.hidden) { if (state.source.mode === 'server') loadServer({ quiet: true }); else if (state.source.mode === 'local') pollLocal(); }
    });
    setInterval(refreshSyncLabel, 10000);

    // Re-flow charts on resize/orientation change
    window.addEventListener('resize', debounce(() => Object.values(charts).forEach((c) => c.resize()), 150));

    loadServer({ force: true });
    startPolling();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
