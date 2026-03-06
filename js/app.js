
const state = {
  allRows: [],
  filteredRows: [],
  charts: { gauge: null, pie: null },
};
const numberFormat = new Intl.NumberFormat('en-US');
const dateFormat = new Intl.DateTimeFormat('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

document.addEventListener('DOMContentLoaded', () => {
  setupTabs();
  setupFilters();
  loadDefaultWorkbook();
});

function setupTabs() {
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.nav-btn').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    });
  });
}

function setupFilters() {
  ['officeFilter', 'reconcilerFilter', 'statusFilter'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });
  document.getElementById('searchInput').addEventListener('input', applyFilters);
  document.getElementById('resetFilters').addEventListener('click', () => {
    document.getElementById('officeFilter').value = '';
    document.getElementById('reconcilerFilter').value = '';
    document.getElementById('statusFilter').value = '';
    document.getElementById('searchInput').value = '';
    applyFilters();
  });
  document.getElementById('fileInput').addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    hydrateFromWorkbook(workbook, file.name);
  });
}

async function loadDefaultWorkbook() {
  const response = await fetch('./data/Team Recon Report NQ 2.xlsx');
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  hydrateFromWorkbook(workbook, 'Team Recon Report NQ 2.xlsx');
}

function hydrateFromWorkbook(workbook, fileName) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  state.allRows = json.map(row => ({
    unit: row['Unit'] || '',
    groupId: row['Group ID'] || '',
    numberGroupId: row['# Group ID'] || '',
    title: row['Title'] || '',
    office: row['Office'] || '',
    reconcilerId: row['Reconciler ID'] || '',
    supvId: row['Supv ID'] || '',
    reconciler: row['Reconciler'] || '',
    nqItems: Number(row['NQ Items'] || 0),
    lastReconDate: normalizeDate(row['Last Recon Date']),
    status: row['Status'] || '',
    comments: row['Comments'] || '',
    submitDate: normalizeDate(row['Submit Date']),
    approveDate: normalizeDate(row['Approve Date']),
  }));
  populateFilterOptions();
  applyFilters();
  document.getElementById('updatedAt').textContent = fileName;
}

function normalizeDate(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function populateFilterOptions() {
  setOptions('officeFilter', uniqueValues(state.allRows.map(r => r.office)));
  setOptions('reconcilerFilter', uniqueValues(state.allRows.map(r => r.reconciler)));
  setOptions('statusFilter', uniqueValues(state.allRows.map(r => r.status)));
}

function setOptions(id, values) {
  const select = document.getElementById(id);
  const current = select.value;
  select.innerHTML = `<option value="">All ${id.replace('Filter', '').toLowerCase()}</option>`;
  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  select.value = current;
}

function uniqueValues(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function applyFilters() {
  const office = document.getElementById('officeFilter').value;
  const reconciler = document.getElementById('reconcilerFilter').value;
  const status = document.getElementById('statusFilter').value;
  const search = document.getElementById('searchInput').value.trim().toLowerCase();

  state.filteredRows = state.allRows.filter(row => {
    const text = [row.unit, row.title, row.office, row.reconciler, row.status, row.numberGroupId].join(' ').toLowerCase();
    return (!office || row.office === office) &&
      (!reconciler || row.reconciler === reconciler) &&
      (!status || row.status === status) &&
      (!search || text.includes(search));
  });
  renderAll();
}

function renderAll() {
  document.getElementById('rowCount').textContent = numberFormat.format(state.filteredRows.length);
  renderTotals();
  renderReconcilerSummary();
  renderOfficeSummary();
  renderDetailTable();
  renderCharts();
}

function renderTotals() {
  const total = sumBy(state.filteredRows, row => row.nqItems);
  document.getElementById('totalNQ').textContent = numberFormat.format(total);
  document.getElementById('reconcilerTotal').textContent = numberFormat.format(total);
  document.getElementById('officeTotal').textContent = numberFormat.format(total);
}

function renderReconcilerSummary() {
  const grouped = groupSum(state.filteredRows, 'reconciler').sort((a, b) => b.value - a.value);
  document.getElementById('reconcilerSummaryBody').innerHTML = grouped.map(item => `
    <tr><td>${escapeHtml(item.key || 'Unknown')}</td><td class="num">${numberFormat.format(item.value)}</td></tr>
  `).join('');
}

function renderOfficeSummary() {
  const grouped = groupSum(state.filteredRows, 'office').sort((a, b) => a.key.localeCompare(b.key));
  document.getElementById('officeSummaryBody').innerHTML = grouped.map(item => `
    <tr><td>${escapeHtml(item.key || 'Unknown')}</td><td class="num">${numberFormat.format(item.value)}</td></tr>
  `).join('');
}

function renderDetailTable() {
  const rows = [...state.filteredRows].sort((a, b) => b.nqItems - a.nqItems || a.office.localeCompare(b.office));
  document.getElementById('detailBody').innerHTML = rows.map(row => `
    <tr>
      <td>${escapeHtml(row.office)}</td>
      <td>${escapeHtml(row.reconciler)}</td>
      <td>${escapeHtml(row.unit)}</td>
      <td>${escapeHtml(String(row.numberGroupId))}</td>
      <td>${escapeHtml(row.title)}</td>
      <td class="num">${numberFormat.format(row.nqItems)}</td>
      <td>${row.lastReconDate ? dateFormat.format(row.lastReconDate) : ''}</td>
      <td><span class="status-pill">${escapeHtml(row.status || 'N/A')}</span></td>
    </tr>
  `).join('');
}

function renderCharts() {
  renderGaugeChart();
  renderPieChart();
}

function renderGaugeChart() {
  const total = sumBy(state.filteredRows, row => row.nqItems);
  const max = 1000;
  const capped = Math.min(total, max);
  const remaining = Math.max(max - capped, 0);
  if (state.charts.gauge) state.charts.gauge.destroy();
  const ctx = document.getElementById('gaugeChart');
  state.charts.gauge = new Chart(ctx, {
    type: 'doughnut',
    data: { labels: ['Total NQ', 'Remaining'], datasets: [{ data: [capped, remaining], backgroundColor: ['#1d4ed8', '#e5e7eb'], borderWidth: 0, circumference: 180, rotation: 270, cutout: '72%' }] },
    options: { responsive: true, plugins: { legend: { display: false }, tooltip: { enabled: false } } },
    plugins: [{
      id: 'markerLine',
      afterDraw(chart) {
        const {ctx} = chart;
        const meta = chart.getDatasetMeta(0).data[0];
        if (!meta) return;
        const centerX = meta.x, centerY = meta.y, outerRadius = meta.outerRadius, innerRadius = meta.innerRadius;
        const targetRatio = 400 / max;
        const angle = Math.PI * (1 - targetRatio);
        const x1 = centerX + Math.cos(Math.PI + angle) * innerRadius;
        const y1 = centerY + Math.sin(Math.PI + angle) * innerRadius;
        const x2 = centerX + Math.cos(Math.PI + angle) * outerRadius;
        const y2 = centerY + Math.sin(Math.PI + angle) * outerRadius;
        ctx.save();
        ctx.strokeStyle = '#16a34a';
        ctx.lineWidth = 3;
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
        ctx.restore();
      }
    }]
  });
}

function renderPieChart() {
  const grouped = groupSum(state.filteredRows, 'office').sort((a, b) => b.value - a.value);
  if (state.charts.pie) state.charts.pie.destroy();
  const ctx = document.getElementById('officePieChart');
  state.charts.pie = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: grouped.map(x => x.key),
      datasets: [{
        data: grouped.map(x => x.value),
        backgroundColor: ['#2583eb','#ea6c33','#8a0da8','#2133a8','#227a7f','#1faa47','#d8b100','#5d9cff','#18b6e6','#b85fd0','#e14a57','#7b5dc8'],
        borderColor: '#ffffff',
        borderWidth: 2,
      }]
    },
    options: {
      plugins: {
        legend: { position: 'right', labels: { boxWidth: 14, usePointStyle: true, pointStyle: 'circle' } },
        tooltip: {
          callbacks: {
            label(context) {
              const total = context.dataset.data.reduce((a, b) => a + b, 0);
              const value = context.raw;
              const pct = total ? ((value / total) * 100).toFixed(1) : 0;
              return `${context.label}: ${numberFormat.format(value)} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function groupSum(rows, field) {
  const map = new Map();
  rows.forEach(row => {
    const key = row[field] || 'Unknown';
    map.set(key, (map.get(key) || 0) + row.nqItems);
  });
  return [...map.entries()].map(([key, value]) => ({ key, value }));
}

function sumBy(rows, selector) {
  return rows.reduce((sum, row) => sum + selector(row), 0);
}

function escapeHtml(value) {
  return String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
}
