// ── Content Analyzer Dashboard — Main JS ────────────────────────────────────

const API = '';
let currentView = 'dashboard';
let globalChannel = '';
let globalDate = '';
let pieChart = null;
let pieChartBarc = null;
let pieChartTabsons = null;

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

async function fetchJson(url, options) {
  const res = await fetch(url, options);
  return readJsonResponse(res);
}

async function fetchBlobResponse(url, options) {
  const res = await fetch(url, options);
  if (!res.ok) throw new Error(await readErrorMessage(res));
  return res;
}

// ── Init ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => { loadChannelDates(); });

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
  if (view === 'dashboard') loadDashboard();
  if (view === 'analysis') loadAnalysisSheets();
  if (view === 'commercial-comparison') loadCommercialComparison();
}

// ── Global Filters ────────────────────────────────────────────────────────────
async function loadChannelDates() {
  try {
    const data = await fetchJson(API + '/api/channels-dates');
    if (data.error) { showToast(data.error, 'error'); return; }
    if (!Array.isArray(data)) { showToast('Unexpected response from server', 'error'); return; }
    const chSel = document.getElementById('global-channel');
    const dtSel = document.getElementById('global-date');
    const channels = [...new Set(data.map(d => d.channel_name))];
    chSel.innerHTML = '<option value="">Select Channel</option>' + channels.map(c => `<option value="${c}">${c}</option>`).join('');
    chSel.onchange = () => {
      globalChannel = chSel.value;
      const dates = data.filter(d => d.channel_name === globalChannel).map(d => d.date);
      dtSel.innerHTML = '<option value="">Select Date</option>' + dates.map(d => `<option value="${d}">${d}</option>`).join('');
      dtSel.onchange = () => { globalDate = dtSel.value; onGlobalFilterChange(); };
    };
  } catch (e) {
    showToast('Failed to connect to server: ' + e.message, 'error');
  }
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
  try {
    const d = await fetchJson(`${API}/api/dashboard?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}&source=${encodeURIComponent(source)}&data_type=${encodeURIComponent(dataType)}`);
    if (d.error) { showToast(d.error, 'error'); return; }
    renderKPIs(d, source, dataType);
    renderPieChart(d, source, dataType);
  } catch (e) { showToast('Failed to load dashboard: ' + e.message, 'error'); }
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
      <div class="kpi-card"><div class="kpi-label">Total Line Item</div><div class="kpi-value">${totalLineItem}</div><div class="kpi-sub">TABSONS: ${d.tabsons_total} &nbsp;|&nbsp; BARC: ${d.barc_total}</div></div>
      <div class="kpi-card"><div class="kpi-label">Commercial</div><div class="kpi-value">${totalCommercial}</div><div class="kpi-sub">TABSONS: ${d.tabsons_commercial} &nbsp;|&nbsp; BARC: ${d.barc_commercial}</div></div>
      <div class="kpi-card"><div class="kpi-label">Promo</div><div class="kpi-value">${totalPromo}</div><div class="kpi-sub">TABSONS: ${d.tabsons_promo} &nbsp;|&nbsp; BARC: ${d.barc_promo}</div></div>
      <div class="kpi-card"><div class="kpi-label">PromoSponsor</div><div class="kpi-value">${totalPromoSp}</div><div class="kpi-sub">TABSONS: ${d.tabsons_promo_sponsor} &nbsp;|&nbsp; BARC: ${d.barc_promo_sponsor}</div></div>
      <div class="kpi-card kpi-program"><div class="kpi-label">Program</div><div class="kpi-value">${totalProgram}</div><div class="kpi-sub">TABSONS: ${d.tabsons_program !== undefined ? d.tabsons_program : '—'} &nbsp;|&nbsp; BARC: ${d.barc_program !== undefined ? d.barc_program : '—'}</div></div>`;
  } else {
    container.innerHTML = `
      <div class="kpi-card"><div class="kpi-label">Total Line Item</div><div class="kpi-value">${d.total_line_item}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card"><div class="kpi-label">Commercial</div><div class="kpi-value">${d.commercial}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card"><div class="kpi-label">Promo</div><div class="kpi-value">${d.promo}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card"><div class="kpi-label">PromoSponsor</div><div class="kpi-value">${d.promo_sponsor}</div><div class="kpi-sub">${dataType}</div></div>
      <div class="kpi-card kpi-program"><div class="kpi-label">Program</div><div class="kpi-value">${d.program !== undefined ? d.program : '—'}</div><div class="kpi-sub">${dataType}</div></div>`;
  }
}

function renderPieChart(d, source, dataType) {
  const singleWrap = document.getElementById('chart-single');
  const dualWrap   = document.getElementById('chart-dual');

  // Destroy existing charts
  if (pieChart)        { pieChart.destroy();        pieChart = null; }
  if (pieChartBarc)    { pieChartBarc.destroy();    pieChartBarc = null; }
  if (pieChartTabsons) { pieChartTabsons.destroy(); pieChartTabsons = null; }

  // ── Colour palettes ────────────────────────────────────────────────────────
  // Palette 1 — Classic Bold  (BARC XML)
  // Commercial → #3B82F6, Promo → #F59E0B, PromoSponsor → #14B8A6, Program → #8B5CF6
  const palette1 = ['#3B82F6', '#F59E0B', '#14B8A6', '#8B5CF6'];
  // Palette 2 — Soft Modern   (TABSONS)
  // Commercial → #10B981, Promo → #F472B6, PromoSponsor → #FBBF24, Program → #6366F1
  const palette2 = ['#10B981', '#F472B6', '#FBBF24', '#6366F1'];

  // 4 segments (including Program)
  const makeLabels = () => ['Commercial', 'Promo', 'PromoSponsor', 'Program'];

  const chartOpts = (labels, values, colors) => ({
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors,
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
      }
    }
  });

  if (source === 'TABSONS-BARC') {
    // Show dual, hide single
    if (singleWrap) singleWrap.style.display = 'none';
    if (dualWrap)   dualWrap.style.display = 'flex';

    const barcCtx    = document.getElementById('pieChartBarc');
    const tabsonsCtx = document.getElementById('pieChartTabsons');

    // BARC chart — Palette 1
    const barcValues = [
      parseNum(d.barc_commercial),
      parseNum(d.barc_promo),
      parseNum(d.barc_promo_sponsor),
      parseNum(d.barc_program)
    ];
    // TABSONS chart — Palette 2
    const tabsonsValues = [
      parseNum(d.tabsons_commercial),
      parseNum(d.tabsons_promo),
      parseNum(d.tabsons_promo_sponsor),
      parseNum(d.tabsons_program)
    ];

    if (barcCtx)    pieChartBarc    = new Chart(barcCtx,    chartOpts(makeLabels(), barcValues,    palette1));
    if (tabsonsCtx) pieChartTabsons = new Chart(tabsonsCtx, chartOpts(makeLabels(), tabsonsValues, palette2));

  } else {
    // Show single, hide dual
    if (singleWrap) singleWrap.style.display = 'flex';
    if (dualWrap)   dualWrap.style.display = 'none';

    const ctx = document.getElementById('pieChart');
    if (!ctx) return;

    const values = [
      parseNum(d.commercial),
      parseNum(d.promo),
      parseNum(d.promo_sponsor),
      parseNum(d.program)
    ];
    // BARC XML → Palette 1 (Classic Bold), TABSONS → Palette 2 (Neon Dark)
    const colors = source === 'BARC XML' ? palette1 : palette2;

    pieChart = new Chart(ctx, chartOpts(makeLabels(), values, colors));
  }
}

function parseNum(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  const s = String(v).replace(/[^0-9:.]/g, '');
  if (s.includes(':')) { const p = s.split(':'); return (parseInt(p[0])||0)*3600+(parseInt(p[1])||0)*60+(parseInt(p[2])||0); }
  return parseInt(s) || 0;
}

// ── Analysis ──────────────────────────────────────────────────────────────────
async function loadAnalysisSheets() {
  if (!globalChannel || !globalDate) return;
  try {
    const sheets = await fetchJson(`${API}/api/sheets?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}`);
    const sel = document.getElementById('analysis-sheet');
    sel.innerHTML = '<option value="">Select Sheet</option>' + sheets.map(s => `<option value="${s.sheet_name}">${s.sheet_name} (${s.row_count} rows)</option>`).join('');
  } catch (e) { showToast('Failed to load sheets', 'error'); }
}

async function loadSheetData() {
  const sheet = document.getElementById('analysis-sheet').value;
  if (!sheet || !globalChannel || !globalDate) return;
  if (sheet === 'COMMERCIAL COMPARISION') { navigate('commercial-comparison'); return; }
  showLoading(true);
  try {
    const data = await fetchJson(`${API}/api/sheet-data?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}&sheet=${encodeURIComponent(sheet)}`);
    if (data.error) { showToast(data.error, 'error'); return; }
    renderDataTable('analysis-table-container', data.rows, sheet);
  } catch (e) { showToast('Failed to load data', 'error'); }
  showLoading(false);
}

function renderDataTable(containerId, rows, title) {
  const container = document.getElementById(containerId);
  if (!rows || rows.length < 2) { container.innerHTML = '<p style="color:var(--muted);padding:20px">No data available</p>'; return; }
  const headerIdx = rows[0].some(h => h && h.length > 0 && h !== 'None') ? 0 : 1;
  const headers = rows[headerIdx];
  const dataRows = rows.slice(headerIdx + 1);
  let html = `<div class="data-table-wrap"><div class="table-header"><h3>${title}</h3><span style="color:var(--muted);font-size:11px;margin-left:auto">${dataRows.length} rows</span></div><div class="table-scroll"><table class="data-table"><thead><tr>`;
  headers.forEach(h => { html += `<th>${h||''}</th>`; });
  html += '</tr></thead><tbody>';
  dataRows.forEach(row => {
    html += '<tr>';
    row.forEach((cell, idx) => {
      // Last column gets td-full so remarks/conclusions show completely
      const cls = (idx === row.length - 1) ? ' class="td-full"' : '';
      html += `<td${cls} title="${(cell||'').replace(/"/g,'&quot;')}">${cell||''}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table></div></div>';
  container.innerHTML = html;
}

// ── Commercial Comparison ─────────────────────────────────────────────────────
async function loadCommercialComparison() {
  if (!globalChannel || !globalDate) { return; }
  showLoading(true);
  try {
    const data = await fetchJson(`${API}/api/commercial-comparison?channel=${encodeURIComponent(globalChannel)}&date=${encodeURIComponent(globalDate)}`);
    if (data.error) { showToast(data.error, 'error'); showLoading(false); return; }
    renderCommercialTables(data);
  } catch (e) { showToast('Failed to load commercial data', 'error'); }
  showLoading(false);
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

  const theadHTML = colDefs.map(c => `<th class="${c.cls}">${c.label}</th>`).join('');

  // ── helper: build tbody rows ──────────────────────────────────────────────
  const buildRows = (rows, actionFn) => rows.map(r => {
    const cells = colDefs.map(c => {
      if (c.key === null) return `<td>${actionFn(r)}</td>`;
      const v = r[c.key] || '';
      const display = (v === '—' || v === '\uFFFD' || v === '?') ? '<span style="color:var(--muted)">—</span>' : escHtml(v);
      // Remarks column — show full text, no truncation
      const extraCls = (c.key === 'remarks') ? ' td-full' : '';
      return `<td class="${extraCls.trim()}" title="${escHtml(v)}">${display}</td>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  // ── Totals ────────────────────────────────────────────────────────────────
  const sumField = (arr, key) => arr.reduce((acc, r) => {
    const n = parseInt((r[key]||'0').replace(/[^0-9]/g,'')) || 0;
    return acc + n;
  }, 0);

  // Count distinct BARC brands and NCT brands
  const uniqueBarcBrands = (arr) => new Set(arr.map(r => r.barc_brand).filter(b => b && b !== '—')).size;
  const uniqueNctBrands  = (arr) => new Set(arr.map(r => r.nct_brand).filter(b => b && b !== '—')).size;

  const totalRow = (arr, label, cssClass) => {
    const bCount = sumField(arr, 'barc_count');
    const nCount = sumField(arr, 'nct_count');
    const bDur   = sumField(arr, 'barc_duration');
    const nDur   = sumField(arr, 'nct_duration');
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
  const totalBarcDur   = sumField(matched, 'barc_duration') + sumField(unmatched, 'barc_duration');
  const totalNctDur    = sumField(matched, 'nct_duration')  + sumField(unmatched, 'nct_duration');
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
  let t = document.getElementById('toast');
  if (!t) { t=document.createElement('div'); t.id='toast'; t.className='toast'; document.body.appendChild(t); }
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(()=>{ t.className='toast'; }, 3500);
}

function showLoading(show) {
  let l = document.getElementById('loading-overlay');
  if (!l) { l=document.createElement('div'); l.id='loading-overlay'; l.className='loading-overlay'; l.innerHTML='<div class="spinner"></div>'; document.body.appendChild(l); }
  l.classList.toggle('show', show);
}
