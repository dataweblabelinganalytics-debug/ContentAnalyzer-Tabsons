// ── Content Analyzer Dashboard — Main JS ────────────────────────────────────

const API = '';
const API_TIMEOUT_MS = 30000;
const SESSION_STATE_KEY = 'contentAnalyzer.sessionState';
const TABLE_WIDTHS_KEY = 'contentAnalyzer.tableWidths';
// Exact workbook fill colors from barc_nct_comparison.py.
const PROGRAM_TYPE_COLORS = {
  commercial: '#E2EFDA',
  promo: '#FFEB9C',
  promo_sponsor: '#EDE7F6',
  program: '#DEEAF1',
};
const KPI_DEFS = [
  { key: 'total', label: 'Total', color: '#2E75B6', textColor: '#2E75B6' },
  { key: 'commercial', label: 'Commercial', color: PROGRAM_TYPE_COLORS.commercial, textColor: '#1F2937' },
  { key: 'promo', label: 'Promo', color: PROGRAM_TYPE_COLORS.promo, textColor: '#1F2937' },
  { key: 'promo_sponsor', label: 'PromoSponsor', color: PROGRAM_TYPE_COLORS.promo_sponsor, textColor: '#1F2937' },
  { key: 'program', label: 'Program', color: PROGRAM_TYPE_COLORS.program, textColor: '#1F2937' },
];

let currentView = 'dashboard';
let globalChannel = '';
let globalDate = '';
let pieChart = null;
let pieChartBarc = null;
let pieChartTabsons = null;
let syncedHoverIndex = null;

function getSessionState() {
  try {
    return JSON.parse(sessionStorage.getItem(SESSION_STATE_KEY) || '{}');
  } catch (e) {
    console.warn('Unable to read session state', e);
    return {};
  }
}

function saveSessionState() {
  try {
    sessionStorage.setItem(SESSION_STATE_KEY, JSON.stringify({
      channel: globalChannel,
      date: globalDate,
      source: document.getElementById('dash-source')?.value || 'TABSONS-BARC',
      dataType: document.getElementById('dash-datatype')?.value || 'COUNT',
      activeView: currentView,
    }));
  } catch (e) {
    console.warn('Unable to save session state', e);
  }
}

function getProgramLabels() {
  return ['Commercial', 'Promo', 'PromoSponsor', 'Program'];
}

function getProgramColors() {
  return [
    PROGRAM_TYPE_COLORS.commercial,
    PROGRAM_TYPE_COLORS.promo,
    PROGRAM_TYPE_COLORS.promo_sponsor,
    PROGRAM_TYPE_COLORS.program,
  ];
}

function notify(type, message, error) {
  if (type === 'error') console.error(message, error || '');
  if (type === 'warning') console.warn(message, error || '');
  if (type === 'success') console.log(message);

  let center = document.getElementById('notification-center');
  if (!center) {
    center = document.createElement('div');
    center.id = 'notification-center';
    center.className = 'notification-center';
    document.body.appendChild(center);
  }

  const item = document.createElement('div');
  item.className = `notice ${type}`;
  item.setAttribute('role', type === 'error' ? 'alert' : 'status');

  const text = document.createElement('div');
  text.className = 'notice-text';
  text.textContent = message;

  const close = document.createElement('button');
  close.className = 'notice-close';
  close.type = 'button';
  close.setAttribute('aria-label', 'Dismiss notification');
  close.textContent = 'x';
  close.onclick = () => item.remove();

  item.append(text, close);
  center.appendChild(item);

  if (type !== 'error') {
    setTimeout(() => item.remove(), type === 'success' ? 3500 : 6500);
  }
}

function showError(message, error) { notify('error', message, error); }
function showSuccess(message) { notify('success', message); }
function showWarning(message, error) { notify('warning', message, error); }

async function readErrorMessage(res) {
  const text = await res.text();
  if (!text) return `Request failed (${res.status})`;
  try {
    const data = JSON.parse(text);
    return data.error || data.message || text;
  } catch (e) {
    return text.slice(0, 300);
  }
}

async function readJsonResponse(res) {
  const text = await res.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch (e) {
    throw new Error(`Server returned non-JSON (${res.status}): ${text.slice(0, 300)}`);
  }
  if (!res.ok || data.error) {
    throw new Error(data.error || data.message || `Request failed (${res.status})`);
  }
  return data;
}

async function safeFetch(url, options = {}, timeoutMs = API_TIMEOUT_MS) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (e) {
    if (e.name === 'AbortError') throw new Error(`Request timed out after ${Math.round(timeoutMs / 1000)} seconds`);
    throw new Error(`Network request failed: ${e.message}`);
  } finally {
    clearTimeout(timeoutId);
  }
}

async function safeApiCall(label, fn, options = {}) {
  const { loading = false, rethrow = false } = options;
  if (loading) showLoading(true);
  try {
    return await fn();
  } catch (e) {
    showError(`${label}: ${e.message}`, e);
    if (rethrow) throw e;
    return null;
  } finally {
    if (loading) showLoading(false);
  }
}

async function fetchJson(url, options) {
  const res = await safeFetch(url, options);
  return readJsonResponse(res);
}

async function fetchBlobResponse(url, options) {
  const res = await safeFetch(url, options);
  if (!res.ok) throw new Error(await readErrorMessage(res));
  return res;
}

// ── Init ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', initApp);

async function initApp() {
  const state = getSessionState();
  const sourceSel = document.getElementById('dash-source');
  const dataTypeSel = document.getElementById('dash-datatype');

  if (sourceSel && state.source) sourceSel.value = state.source;
  if (dataTypeSel && state.dataType) dataTypeSel.value = state.dataType;

  if (sourceSel) sourceSel.onchange = () => { saveSessionState(); loadDashboard(); };
  if (dataTypeSel) dataTypeSel.onchange = () => { saveSessionState(); loadDashboard(); };

  await loadChannelDates();
  navigate(state.activeView || 'dashboard');
}

// ── Navigation ────────────────────────────────────────────────────────────────
function navigate(view) {
  currentView = view;
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const el = document.getElementById('view-' + view);
  if (el) el.classList.add('active');
  const nav = document.querySelector(`[data-view="${view}"]`);
  if (nav) nav.classList.add('active');
  const titles = { dashboard: 'Dashboard', analysis: 'Sheet Data', 'commercial-comparison': 'Commercial Comparison', 'compare-report': 'Compare Report', download: 'Download Reports' };
  document.getElementById('topbar-title').textContent = titles[view] || 'Dashboard';
  saveSessionState();
  if (view === 'dashboard') loadDashboard();
  if (view === 'analysis') loadAnalysisSheets();
  if (view === 'commercial-comparison') loadCommercialComparison();
}

// ── Global Filters ────────────────────────────────────────────────────────────
async function loadChannelDates() {
  return safeApiCall('Failed to load channels and dates', async () => {
    const data = await fetchJson(API + '/api/channels-dates');
    if (!Array.isArray(data)) throw new Error('Unexpected response from server');

    const chSel = document.getElementById('global-channel');
    const dtSel = document.getElementById('global-date');
    const channels = [...new Set(data.map(d => d.channel_name))];
    const state = getSessionState();

    chSel.innerHTML = '<option value="">Select Channel</option>' + channels.map(c => `<option value="${escHtml(c)}">${escHtml(c)}</option>`).join('');

    const populateDates = (restore = false) => {
      globalChannel = chSel.value;
      const dates = data.filter(d => d.channel_name === globalChannel).map(d => d.date);
      dtSel.innerHTML = '<option value="">Select Date</option>' + dates.map(d => `<option value="${escHtml(d)}">${escHtml(d)}</option>`).join('');
      globalDate = '';
      if (restore && dates.includes(state.date)) {
        dtSel.value = state.date;
        globalDate = state.date;
      }
    };

    if (channels.includes(state.channel)) chSel.value = state.channel;
    populateDates(true);

    chSel.onchange = () => {
      populateDates();
      saveSessionState();
      onGlobalFilterChange();
    };

    dtSel.onchange = () => {
      globalDate = dtSel.value;
      saveSessionState();
      onGlobalFilterChange();
    };
  });
}

function onGlobalFilterChange() {
  if (!globalChannel || !globalDate) return;
  if (currentView === 'dashboard') loadDashboard();
  if (currentView === 'analysis') loadAnalysisSheets();
  if (currentView === 'commercial-comparison') loadCommercialComparison();
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
async function loadDashboard() {
  if (!globalChannel || !globalDate) return;
  const source = document.getElementById('dash-source').value;
  const dataType = document.getElementById('dash-datatype').value;
  saveSessionState();
  return safeApiCall('Failed to load dashboard', async () => {
    const d = await fetchJson(`${API}/api/dashboard?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}&source=${encodeURIComponent(source)}&data_type=${encodeURIComponent(dataType)}`);
    renderDashboardKPIs(d, source, dataType);
    renderPieChart(d, source, dataType);
  });
}

function renderKPIs(d, source, dataType) {
  const container = document.getElementById('kpi-row');
  if (source === 'TABSONS-BARC') {
    // Helper: add two values (numbers or HH:MM:SS strings)
    const addVals = (a, b) => {
      const na = parseNum(a), nb = parseNum(b);
      if (typeof a === 'string' && a.includes(':')) {
        // duration mode — convert back to HH:MM:SS
        const total = na + nb;
        const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = total % 60;
        return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
      }
      return na + nb;
    };
    const totalLineItem   = addVals(d.tabsons_total,        d.barc_total);
    const totalCommercial = addVals(d.tabsons_commercial,   d.barc_commercial);
    const totalPromo      = addVals(d.tabsons_promo,        d.barc_promo);
    const totalPromoSp    = addVals(d.tabsons_promo_sponsor,d.barc_promo_sponsor);
    const totalProgram    = addVals(d.tabsons_program !== undefined ? d.tabsons_program : 0,
                                   d.barc_program    !== undefined ? d.barc_program    : 0);
    container.innerHTML = `
      <div class="kpi-card"><div class="kpi-label">Total</div><div class="kpi-value">${totalLineItem}</div><div class="kpi-sub">TABSONS: ${d.tabsons_total} &nbsp;|&nbsp; BARC: ${d.barc_total}</div></div>
      <div class="kpi-card"><div class="kpi-label">Commercial</div><div class="kpi-value">${totalCommercial}</div><div class="kpi-sub">TABSONS: ${d.tabsons_commercial} &nbsp;|&nbsp; BARC: ${d.barc_commercial}</div></div>
      <div class="kpi-card"><div class="kpi-label">Promo</div><div class="kpi-value">${totalPromo}</div><div class="kpi-sub">TABSONS: ${d.tabsons_promo} &nbsp;|&nbsp; BARC: ${d.barc_promo}</div></div>
      <div class="kpi-card"><div class="kpi-label">PromoSponsor</div><div class="kpi-value">${totalPromoSp}</div><div class="kpi-sub">TABSONS: ${d.tabsons_promo_sponsor} &nbsp;|&nbsp; BARC: ${d.barc_promo_sponsor}</div></div>
      <div class="kpi-card kpi-program"><div class="kpi-label">Program</div><div class="kpi-value">${totalProgram}</div><div class="kpi-sub">TABSONS: ${d.tabsons_program !== undefined ? d.tabsons_program : '—'} &nbsp;|&nbsp; BARC: ${d.barc_program !== undefined ? d.barc_program : '—'}</div></div>`;
  } else {
    container.innerHTML = `
      <div class="kpi-card"><div class="kpi-label">Total</div><div class="kpi-value">${d.total_line_item}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card"><div class="kpi-label">Commercial</div><div class="kpi-value">${d.commercial}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card"><div class="kpi-label">Promo</div><div class="kpi-value">${d.promo}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card"><div class="kpi-label">PromoSponsor</div><div class="kpi-value">${d.promo_sponsor}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card kpi-program"><div class="kpi-label">Program</div><div class="kpi-value">${d.program !== undefined ? d.program : '—'}</div><div class="kpi-sub">${dataType}</div></div>`;
  }
}

function renderDashboardKPIs(d, source, dataType) {
  const container = document.getElementById('kpi-row');
  if (!container) return;

  const isDuration = dataType === 'DURATION';
  const displayMetric = (value) => isDuration ? formatDuration(value) : escHtml(value ?? 0);
  const addVals = (a, b) => isDuration ? formatDuration(parseDurationSeconds(a) + parseDurationSeconds(b)) : parseCountValue(a) + parseCountValue(b);
  const sourcePair = (tabsons, barc) => `
        <span class="kpi-source-pair">
          <span class="kpi-source-item"><span class="kpi-source-label">TABSONS:</span> ${displayMetric(tabsons)}</span>
          <span class="kpi-separator">|</span>
          <span class="kpi-source-item"><span class="kpi-source-label">BARC:</span> ${displayMetric(barc)}</span>
        </span>`;
  const card = (def, primary, sub) => `
      <div class="kpi-card" style="--kpi-color:${def.color};--kpi-bg:${hexToRgba(def.color, 0.14)};--kpi-text:${def.textColor}">
        <div class="kpi-label">${def.label}</div>
        <div class="kpi-value">${primary}</div>
        <div class="kpi-sub">${sub}</div>
      </div>`;

  if (source === 'TABSONS-BARC') {
    const values = {
      total: { tabsons: d.tabsons_total, barc: d.barc_total },
      commercial: { tabsons: d.tabsons_commercial, barc: d.barc_commercial },
      promo: { tabsons: d.tabsons_promo, barc: d.barc_promo },
      promo_sponsor: { tabsons: d.tabsons_promo_sponsor, barc: d.barc_promo_sponsor },
      program: { tabsons: d.tabsons_program ?? 0, barc: d.barc_program ?? 0 },
    };

    container.innerHTML = KPI_DEFS.map(def => {
      const v = values[def.key];
      return card(
        def,
        sourcePair(v.tabsons, v.barc),
        `TOTAL: ${displayMetric(addVals(v.tabsons, v.barc))}`
      );
    }).join('');
    return;
  }

  const values = {
    total: d.total_line_item,
    commercial: d.commercial,
    promo: d.promo,
    promo_sponsor: d.promo_sponsor,
    program: d.program ?? 0,
  };
  container.innerHTML = KPI_DEFS.map(def => card(def, displayMetric(values[def.key]), escHtml(dataType))).join('');
}

function renderPieChart(d, source, dataType) {
  const singleWrap = document.getElementById('chart-single');
  const dualWrap   = document.getElementById('chart-dual');
  const isDuration = dataType === 'DURATION';
  const chartValue = (value) => isDuration ? parseDurationSeconds(value) : parseCountValue(value);
  syncedHoverIndex = null;

  // Destroy existing charts
  if (pieChart)        { pieChart.destroy();        pieChart = null; }
  if (pieChartBarc)    { pieChartBarc.destroy();    pieChartBarc = null; }
  if (pieChartTabsons) { pieChartTabsons.destroy(); pieChartTabsons = null; }

  // ── Colour palettes ────────────────────────────────────────────────────────
  // Workbook colors from barc_nct_comparison.py.
  const palette1 = getProgramColors();
  const palette2 = getProgramColors();

  // 4 segments (including Program)
  const makeLabels = getProgramLabels;

  const hoverSyncPlugin = {
    id: 'dashboardHoverSync',
    afterEvent(chart, args) {
      if (args.event.type === 'mouseout' && (chart === pieChartBarc || chart === pieChartTabsons)) {
        syncDualChartHover(null);
      }
    }
  };

  const chartOpts = (labels, values, colors, syncHover = false) => ({
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors,
        hoverBackgroundColor: colors,
        hoverBorderColor: colors,
        hoverOffset: 8,
        borderColor: 'transparent',
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#475569', padding: 14, font: { family: 'DM Sans', size: 11 } }
        }
      },
      onHover: syncHover ? (event, active) => {
        syncDualChartHover(active.length ? active[0].index : null);
      } : undefined
    },
    plugins: syncHover ? [hoverSyncPlugin] : []
  });

  if (source === 'TABSONS-BARC') {
    // Show dual, hide single
    if (singleWrap) singleWrap.style.display = 'none';
    if (dualWrap)   dualWrap.style.display = 'flex';

    const barcCtx    = document.getElementById('pieChartBarc');
    const tabsonsCtx = document.getElementById('pieChartTabsons');

    // BARC chart
    const barcValues = [
      chartValue(d.barc_commercial),
      chartValue(d.barc_promo),
      chartValue(d.barc_promo_sponsor),
      chartValue(d.barc_program)
    ];
    // TABSONS chart
    const tabsonsValues = [
      chartValue(d.tabsons_commercial),
      chartValue(d.tabsons_promo),
      chartValue(d.tabsons_promo_sponsor),
      chartValue(d.tabsons_program)
    ];

    if (barcCtx)    pieChartBarc    = new Chart(barcCtx,    chartOpts(makeLabels(), barcValues,    palette1, true));
    if (tabsonsCtx) pieChartTabsons = new Chart(tabsonsCtx, chartOpts(makeLabels(), tabsonsValues, palette2, true));

  } else {
    // Show single, hide dual
    if (singleWrap) singleWrap.style.display = 'flex';
    if (dualWrap)   dualWrap.style.display = 'none';

    const ctx = document.getElementById('pieChart');
    if (!ctx) return;

    const values = [
      chartValue(d.commercial),
      chartValue(d.promo),
      chartValue(d.promo_sponsor),
      chartValue(d.program)
    ];
    const colors = source === 'BARC XML' ? palette1 : palette2;

    pieChart = new Chart(ctx, chartOpts(makeLabels(), values, colors));
  }
}

function syncDualChartHover(index) {
  if (syncedHoverIndex === index) return;
  syncedHoverIndex = index;
  [pieChartBarc, pieChartTabsons].forEach(chart => setChartHoverIndex(chart, index));
}

function setChartHoverIndex(chart, index) {
  if (!chart) return;
  if (index === null || index === undefined) {
    chart.setActiveElements([]);
    chart.tooltip.setActiveElements([], { x: 0, y: 0 });
    chart.update('none');
    return;
  }

  const element = chart.getDatasetMeta(0)?.data?.[index];
  if (!element) return;
  const position = element.tooltipPosition();
  chart.setActiveElements([{ datasetIndex: 0, index }]);
  chart.tooltip.setActiveElements([{ datasetIndex: 0, index }], position);
  chart.update('none');
}

function parseNum(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  const s = String(v).replace(/[^0-9:.]/g, '');
  if (s.includes(':')) { const p = s.split(':'); return (parseInt(p[0])||0)*3600+(parseInt(p[1])||0)*60+(parseInt(p[2])||0); }
  return parseInt(s) || 0;
}

function parseCountValue(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const n = parseFloat(String(value ?? '').replace(/,/g, '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function formatDuration(value) {
  const totalSeconds = Math.max(0, Math.floor(parseDurationSeconds(value)));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

function parseDurationSeconds(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return 0;

  const dayMatch = raw.match(/(\d+)\s*days?,\s*(\d{1,3}):(\d{1,2})(?::(\d{1,2}))?/);
  if (dayMatch) {
    return (parseInt(dayMatch[1], 10) || 0) * 86400
      + (parseInt(dayMatch[2], 10) || 0) * 3600
      + (parseInt(dayMatch[3], 10) || 0) * 60
      + (parseInt(dayMatch[4], 10) || 0);
  }

  const excelDateMatch = raw.match(/^1900-01-(\d{2})[ t](\d{1,2}):(\d{1,2}):(\d{1,2})/);
  if (excelDateMatch) {
    const elapsedDays = Math.max(0, (parseInt(excelDateMatch[1], 10) || 1) - 1);
    return elapsedDays * 86400
      + (parseInt(excelDateMatch[2], 10) || 0) * 3600
      + (parseInt(excelDateMatch[3], 10) || 0) * 60
      + (parseInt(excelDateMatch[4], 10) || 0);
  }

  const timeMatches = [...raw.matchAll(/(\d{1,6}):(\d{1,2})(?::(\d{1,2}))?/g)];
  if (timeMatches.length) {
    const match = timeMatches[timeMatches.length - 1];
    return (parseInt(match[1], 10) || 0) * 3600
      + (parseInt(match[2], 10) || 0) * 60
      + (parseInt(match[3], 10) || 0);
  }

  const n = parseFloat(raw.replace(/,/g, '').replace(/[^0-9.]/g, '')) || 0;
  if (/\b(h|hr|hour|hours)\b/.test(raw)) return n * 3600;
  if (/\b(m|min|mins|minute|minutes)\b/.test(raw)) return n * 60;
  return n;
}

function isDurationColumn(labelOrKey) {
  return /duration/i.test(String(labelOrKey || ''));
}

function hexToRgba(hex, alpha) {
  const clean = String(hex).replace('#', '');
  const value = parseInt(clean.length === 3 ? clean.split('').map(c => c + c).join('') : clean, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// ── Analysis ──────────────────────────────────────────────────────────────────
async function loadAnalysisSheets() {
  if (!globalChannel || !globalDate) return;
  return safeApiCall('Failed to load sheets', async () => {
    const sheets = await fetchJson(`${API}/api/sheets?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}`);
    const sel = document.getElementById('analysis-sheet');
    sel.innerHTML = '<option value="">Select Sheet</option>' + sheets.map(s => `<option value="${escHtml(s.sheet_name)}">${escHtml(s.sheet_name)} (${escHtml(s.row_count)} rows)</option>`).join('');
  });
}

async function loadSheetData() {
  const sheet = document.getElementById('analysis-sheet').value;
  if (!sheet || !globalChannel || !globalDate) return;
  if (sheet === 'COMMERCIAL COMPARISION') { navigate('commercial-comparison'); return; }
  return safeApiCall('Failed to load data', async () => {
    const data = await fetchJson(`${API}/api/sheet-data?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}&sheet=${encodeURIComponent(sheet)}`);
    renderDataTable('analysis-table-container', data.rows, sheet);
  }, { loading: true });
}

function renderDataTable(containerId, rows, title) {
  const container = document.getElementById(containerId);
  if (!rows || rows.length < 2) { container.innerHTML = '<p style="color:var(--muted);padding:20px">No data available</p>'; return; }
  const headerIdx = rows[0].some(h => h && h.length > 0 && h !== 'None') ? 0 : 1;
  const headers = rows[headerIdx];
  const dataRows = rows.slice(headerIdx + 1);
  const tableKey = makeTableKey(containerId, title);
  let html = `<div class="data-table-wrap"><div class="table-header"><h3>${escHtml(title)}</h3><span style="color:var(--muted);font-size:11px;margin-left:auto">${dataRows.length} rows</span></div><div class="table-scroll"><table class="data-table" data-table-key="${tableKey}"><thead><tr>`;
  headers.forEach(h => { html += `<th>${escHtml(h || '')}</th>`; });
  html += '</tr></thead><tbody>';
  dataRows.forEach(row => {
    html += '<tr>';
    row.forEach((cell, idx) => {
      html += renderTableCell(cell, headers[idx], idx === row.length - 1);
    });
    html += '</tr>';
  });
  html += '</tbody></table></div></div>';
  container.innerHTML = html;
  initResizableTables(container);
}

// ── Commercial Comparison ─────────────────────────────────────────────────────
async function loadCommercialComparison() {
  if (!globalChannel || !globalDate) { return; }
  return safeApiCall('Failed to load commercial data', async () => {
    const data = await fetchJson(`${API}/api/commercial-comparison?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}`);
    renderCommercialTables(data);
  }, { loading: true });
}

function renderCommercialTables(data) {
  const container = document.getElementById('commercial-tables');

  // Columns: BARC Brand, TABSONS Brand, BARC Count, TABSONS Count, BARC Duration, TABSONS Duration, Remarks, Action
  const colDefs = [
    { label: 'BARC Brand',      key: 'barc_brand',     cls: 'th-barc'   },
    { label: 'TABSONS Brand',   key: 'nct_brand',      cls: 'th-nct'    },
    { label: 'BARC Count',      key: 'barc_count',     cls: 'th-barc'   },
    { label: 'TABSONS Count',   key: 'nct_count',      cls: 'th-nct'    },
    { label: 'BARC Duration',   key: 'barc_duration',  cls: 'th-barc'   },
    { label: 'TABSONS Duration',key: 'nct_duration',   cls: 'th-nct'    },
    { label: 'Remarks',         key: 'remarks',        cls: 'th-common' },
    { label: 'Action',          key: null,             cls: 'th-action' },
  ];

  const theadHTML = colDefs.map(c => `<th class="${c.cls}">${escHtml(c.label)}</th>`).join('');

  // ── helper: build tbody rows ──────────────────────────────────────────────
  const buildRows = (rows, actionFn) => rows.map(r => {
    const cells = colDefs.map(c => {
      if (c.key === null) return `<td>${actionFn(r)}</td>`;
      const v = r[c.key] || '';
      const display = (v === '—' || v === '\uFFFD' || v === '?') ? '<span style="color:var(--muted)">—</span>' : escHtml(v);
      // Remarks column — show full text, no truncation
      const extraCls = (c.key === 'remarks') ? ' td-full' : '';
      return renderTableCell(v, c.label, c.key === 'remarks');
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  // ── Totals ────────────────────────────────────────────────────────────────
  const sumField = (arr, key) => arr.reduce((acc, r) => {
    const n = isDurationColumn(key) ? parseDurationSeconds(r[key]) : parseInt((r[key]||'0').replace(/[^0-9]/g,'')) || 0;
    return acc + n;
  }, 0);

  // Count distinct BARC brands and NCT brands
  const uniqueBarcBrands = (arr) => new Set(arr.map(r => r.barc_brand).filter(b => b && b !== '—')).size;
  const uniqueNctBrands  = (arr) => new Set(arr.map(r => r.nct_brand).filter(b => b && b !== '—')).size;

  const totalRow = (arr, label, cssClass) => {
    const bCount = sumField(arr, 'barc_count');
    const nCount = sumField(arr, 'nct_count');
    const bDur   = formatDuration(sumField(arr, 'barc_duration'));
    const nDur   = formatDuration(sumField(arr, 'nct_duration'));
    const bUniq  = uniqueBarcBrands(arr);
    const nUniq  = uniqueNctBrands(arr);
    return `<tr class="total-row ${cssClass}">
      <td colspan="2"><strong>${label} TOTAL</strong></td>
      <td>${bCount || '—'}</td><td>${nCount || '—'}</td>
      <td>${bDur || '—'}</td><td>${nDur || '—'}</td>
      <td colspan="2"><span class="total-badge barc-badge">${bUniq} BARC unique</span> <span class="total-badge nct-badge">${nUniq} NCT matched</span></td>
    </tr>`;
  };

  const matched   = data.matched   || [];
  const unmatched = data.unmatched || [];

  // Grand total
  const allBarcBrands = new Set([...matched, ...unmatched].map(r => r.barc_brand).filter(b => b && b !== '—')).size;
  const allNctBrands  = new Set([...matched, ...unmatched].map(r => r.nct_brand).filter(b => b && b !== '—')).size;
  const totalBarcCount = sumField(matched, 'barc_count') + sumField(unmatched, 'barc_count');
  const totalNctCount  = sumField(matched, 'nct_count')  + sumField(unmatched, 'nct_count');
  const totalBarcDur   = formatDuration(sumField(matched, 'barc_duration') + sumField(unmatched, 'barc_duration'));
  const totalNctDur    = formatDuration(sumField(matched, 'nct_duration')  + sumField(unmatched, 'nct_duration'));
  const unmatchedNct   = uniqueNctBrands(unmatched);

  // ── Build HTML ────────────────────────────────────────────────────────────
  let html = `
  <!-- Legend -->
  <div class="comm-legend">
    <div class="comm-legend-item"><div class="comm-legend-dot" style="background:#3B82F6"></div><span style="color:#2563EB">BARC columns</span></div>
    <div class="comm-legend-item"><div class="comm-legend-dot" style="background:#10B981"></div><span style="color:#047857">TABSONS columns</span></div>
    <div class="comm-legend-item"><div class="comm-legend-dot" style="background:#8B5CF6"></div><span style="color:#7C3AED">Shared / Remarks</span></div>
    <div class="comm-legend-item"><div class="comm-legend-dot" style="background:#F59E0B"></div><span style="color:#B45309">Action</span></div>
  </div>

  <!-- MATCHED section -->
  <div class="comm-section">
    <div class="comm-section-header matched">
      <div>
        <span class="comm-section-title">✓ MATCHED — BARC vs TABSONS Commercial</span>
        <span class="comm-section-count">${matched.length} brands</span>
      </div>
    </div>
    <div class="data-table-wrap">
      <div class="table-scroll">
        <table class="data-table comm-table">
          <thead><tr>${theadHTML}</tr></thead>
          <tbody>
            ${matched.length ? buildRows(matched, r => `<button class="action-btn remove" onclick="removeBrand('${esc(r.barc_brand)}')">✕ Unmatch</button>`) : '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:20px">No matched brands</td></tr>'}
            ${matched.length ? totalRow(matched, 'MATCHING COMMERCIAL', 'total-matched') : ''}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- UNMATCHED section -->
  <div class="comm-section">
    <div class="comm-section-header unmatched">
      <div>
        <span class="comm-section-title">✕ NOT MATCHED — TABSONS Brands Not in BARC</span>
        <span class="comm-section-count">${unmatched.length} brands</span>
      </div>
    </div>
    <div class="data-table-wrap">
      <div class="table-scroll">
        <table class="data-table comm-table">
          <thead><tr>${theadHTML}</tr></thead>
          <tbody>
            ${unmatched.length ? buildRows(unmatched, r => `<button class="action-btn merge" onclick="mergeBrand('${esc(r.nct_brand)}')">✓ Match to BARC</button>`) : '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:20px">No unmatched brands</td></tr>'}
            ${unmatched.length ? totalRow(unmatched, 'NCT UNMATCHED', 'total-unmatched') : ''}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- GRAND TOTAL section -->
  <div class="comm-section">
    <div class="data-table-wrap">
      <div class="table-scroll">
        <table class="data-table comm-table">
          <thead><tr>${theadHTML}</tr></thead>
          <tbody>
            <tr class="total-row total-grand">
              <td colspan="2"><strong>GRAND TOTAL</strong></td>
              <td>${totalBarcCount || '—'}</td>
              <td>${totalNctCount  || '—'}</td>
              <td>${totalBarcDur   || '—'}</td>
              <td>${totalNctDur    || '—'}</td>
              <td colspan="2">
                <span class="total-badge barc-badge">${allBarcBrands} BARC unique</span>
                <span class="total-badge nct-badge">${allNctBrands} NCT unique</span>
                ${unmatchedNct > 0 ? `<span class="total-note">NCT Tagging less than BARC by ${unmatchedNct}</span>` : ''}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>`;

  container.innerHTML = html;
  initResizableTables(container);
}

function makeTableKey(...parts) {
  return parts
    .map(part => String(part || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, ''))
    .filter(Boolean)
    .join(':') || 'table';
}

function renderTableCell(value, header, forceFull = false) {
  const raw = String(value ?? '');
  const trimmed = raw.trim();
  const isMissing = !trimmed || trimmed === '-' || trimmed === '?' || trimmed.includes('\uFFFD');
  const displayValue = isMissing ? '-' : (isDurationColumn(header) ? formatDuration(raw) : raw);
  const escaped = escHtml(displayValue);
  const title = escHtml(isMissing ? '-' : displayValue);
  const isLong = displayValue.length > 72 || displayValue.includes('\n');
  const cls = forceFull ? ' class="td-full"' : '';

  if (!isMissing && (isLong || forceFull)) {
    const preview = escHtml(displayValue.length > 110 ? `${displayValue.slice(0, 110)}...` : displayValue);
    return `<td${cls} title="${title}"><details class="cell-expand"><summary>${preview}</summary><div>${escaped}</div></details></td>`;
  }

  if (isMissing) return `<td${cls} title="${title}"><span style="color:var(--muted)">-</span></td>`;
  return `<td${cls} title="${title}"><span class="cell-text">${escaped}</span></td>`;
}

function getStoredTableWidths() {
  try {
    return JSON.parse(sessionStorage.getItem(TABLE_WIDTHS_KEY) || '{}');
  } catch (e) {
    console.warn('Unable to read table widths', e);
    return {};
  }
}

function saveTableWidth(tableKey, colIndex, width) {
  try {
    const widths = getStoredTableWidths();
    widths[tableKey] = widths[tableKey] || {};
    widths[tableKey][colIndex] = width;
    sessionStorage.setItem(TABLE_WIDTHS_KEY, JSON.stringify(widths));
  } catch (e) {
    console.warn('Unable to save table width', e);
  }
}

function ensureColGroup(table, columnCount) {
  let colgroup = table.querySelector('colgroup');
  if (!colgroup) {
    colgroup = document.createElement('colgroup');
    for (let i = 0; i < columnCount; i++) colgroup.appendChild(document.createElement('col'));
    table.insertBefore(colgroup, table.firstChild);
  }
  return colgroup;
}

function initResizableTables(root = document) {
  const tables = root.querySelectorAll('table.data-table');
  const stored = getStoredTableWidths();

  tables.forEach((table, tableIndex) => {
    const headers = Array.from(table.querySelectorAll('thead th'));
    if (!headers.length) return;

    const tableKey = table.dataset.tableKey || makeTableKey(currentView, tableIndex);
    table.dataset.tableKey = tableKey;
    const colgroup = ensureColGroup(table, headers.length);
    const cols = Array.from(colgroup.children);

    headers.forEach((th, colIndex) => {
      const savedWidth = stored[tableKey]?.[colIndex];
      if (savedWidth) {
        cols[colIndex].style.width = `${savedWidth}px`;
        th.style.width = `${savedWidth}px`;
      }

      if (!th.querySelector('.col-resizer')) {
        const handle = document.createElement('span');
        handle.className = 'col-resizer';
        handle.setAttribute('aria-hidden', 'true');
        th.appendChild(handle);

        handle.addEventListener('mousedown', (event) => {
          event.preventDefault();
          event.stopPropagation();

          const startX = event.clientX;
          const startWidth = th.getBoundingClientRect().width;
          table.classList.add('is-resizing');

          const onMove = (moveEvent) => {
            const nextWidth = Math.max(72, Math.round(startWidth + moveEvent.clientX - startX));
            cols[colIndex].style.width = `${nextWidth}px`;
            th.style.width = `${nextWidth}px`;
          };

          const onUp = () => {
            const finalWidth = Math.round(th.getBoundingClientRect().width);
            saveTableWidth(tableKey, colIndex, finalWidth);
            table.classList.remove('is-resizing');
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
          };

          document.addEventListener('mousemove', onMove);
          document.addEventListener('mouseup', onUp);
        });
      }
    });
  });
}

function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function esc(s) { return (s||'').replace(/'/g,"\\'").replace(/"/g,'&quot;'); }

// ── Brand actions ─────────────────────────────────────────────────────────────
async function removeBrand(brandName) {
  if (!confirm(`Unmatch "${brandName}" — move it back to NOT MATCHED?`)) return;
  showLoading(true);
  try {
    const d = await fetchJson(API+'/api/commercial/move-brand', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ channel: globalChannel, date: globalDate, action: 'remove_from_matched', brand_name: brandName })
    });
    if (d.error) throw new Error(d.error);
    showToast(`"${brandName}" moved to Not Matched`, 'success');
    loadCommercialComparison();
  } catch(e) { showToast('Error: '+e.message,'error'); }
  showLoading(false);
}

async function mergeBrand(brandName) {
  const target = prompt(`Match "${brandName}" to which BARC brand?\n\nEnter the exact BARC brand name:`);
  if (!target || !target.trim()) return;
  showLoading(true);
  try {
    const d = await fetchJson(API+'/api/commercial/move-brand', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ channel: globalChannel, date: globalDate, action: 'merge_to_matched', brand_name: brandName, target_barc_brand: target.trim() })
    });
    if (d.error) throw new Error(d.error);
    showToast(`"${brandName}" matched to "${target.trim()}"`, 'success');
    loadCommercialComparison();
  } catch(e) { showToast('Error: '+e.message,'error'); }
  showLoading(false);
}

// ── Compare Report ────────────────────────────────────────────────────────────
function onCompareFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('compare-file-info').innerHTML = `<div style="display:flex;align-items:center;gap:10px;padding:10px;background:var(--surface2);border-radius:12px;margin-top:10px;border:1px solid var(--border)"><span style="color:#2563EB;font-family:monospace;font-size:12px">${file.name}</span><span style="color:var(--muted);font-size:11px">${fmtSize(file.size)}</span></div>`;
  document.getElementById('compare-run-btn').style.display = 'inline-flex';
}

async function runCompare() {
  const input = document.getElementById('compare-file-input');
  const file = input.files[0];
  if (!file) { showToast('Please select a file','error'); return; }
  showLoading(true);
  try {
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetchBlobResponse(API+'/api/compare',{method:'POST',body:fd});
    const blob = await res.blob();
    let filename = 'comparison_result.xlsx';
    const disp = res.headers.get('Content-Disposition');
    if (disp) { const m = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disp); if(m) filename = m[1].replace(/['"]/g,''); }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download=filename; a.click();
    setTimeout(()=>URL.revokeObjectURL(url),60000);
    showToast('Comparison complete! File downloaded.','success');
    loadChannelDates();
  } catch(e) { showToast('Error: '+e.message,'error'); }
  showLoading(false);
}

async function downloadTemplate() {
  try {
    const res = await fetchBlobResponse(API+'/api/template');
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download='brand_comparison_template.xlsx'; a.click();
    showToast('Template downloaded','success');
  } catch(e) { showToast('Error downloading template','error'); }
}

// ── Downloads ─────────────────────────────────────────────────────────────────
async function downloadReport(type) {
  if (!globalChannel || !globalDate) { showToast('Select channel and date first','error'); return; }
  showLoading(true);
  try {
    const url = `${API}/api/download/${type}?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}`;
    const res = await fetchBlobResponse(url);
    const blob = await res.blob();
    let filename = `report_${type}.xlsx`;
    const disp = res.headers.get('Content-Disposition');
    if (disp) { const m = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disp); if(m) filename = m[1].replace(/['"]/g,''); }
    const u = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=u; a.download=filename; a.click();
    showToast(`${type} report downloaded!`,'success');
  } catch(e) { showToast('Error: '+e.message,'error'); }
  showLoading(false);
}

// ── Upload for processing ─────────────────────────────────────────────────────
function onUploadFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('upload-file-info').innerHTML = `<div style="display:flex;align-items:center;gap:10px;padding:10px;background:var(--surface2);border-radius:12px;margin-top:10px;border:1px solid var(--border)"><span style="color:#2563EB;font-family:monospace;font-size:12px">${file.name}</span><span style="color:var(--muted);font-size:11px">${fmtSize(file.size)}</span></div>`;
  document.getElementById('upload-run-btn').style.display = 'inline-flex';
}

async function runUpload() {
  const input = document.getElementById('upload-file-input');
  const file = input.files[0];
  if (!file) return;
  showLoading(true);
  try {
    const fd = new FormData(); fd.append('files', file);
    const res = await fetchBlobResponse(API+'/analyze',{method:'POST',body:fd});
    const blob = await res.blob();
    let filename = 'output.xlsx';
    const disp = res.headers.get('Content-Disposition');
    if (disp) { const m = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disp); if(m) filename = m[1].replace(/['"]/g,''); }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download=filename; a.click();
    showToast('File processed and saved to database!','success');
    loadChannelDates();
  } catch(e) { showToast('Error: '+e.message,'error'); }
  showLoading(false);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtSize(b) { if(b<1024)return b+' B'; if(b<1048576)return(b/1024).toFixed(1)+' KB'; return(b/1048576).toFixed(1)+' MB'; }

function showToast(msg, type='info') {
  if (type === 'error') { showError(msg); return; }
  if (type === 'success') { showSuccess(msg); return; }
  if (type === 'warning' || type === 'warn') { showWarning(msg); return; }
  notify('info', msg);
}

function showLoading(show) {
  let l = document.getElementById('loading-overlay');
  if (!l) { l=document.createElement('div'); l.id='loading-overlay'; l.className='loading-overlay'; l.innerHTML='<div class="spinner"></div>'; document.body.appendChild(l); }
  l.classList.toggle('show', show);
}
