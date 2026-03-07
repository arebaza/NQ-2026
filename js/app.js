let allRows = [];
let officePieChart = null;

document.addEventListener("DOMContentLoaded", () => {
  const officeFilter = document.getElementById("officeFilter");
  const reconcilerFilter = document.getElementById("reconcilerFilter");
  const searchBox = document.getElementById("searchBox");

  officeFilter.addEventListener("change", () => {
    updateDependentFilters("office");
    applyFilters();
  });

  reconcilerFilter.addEventListener("change", () => {
    updateDependentFilters("reconciler");
    applyFilters();
  });

  searchBox.addEventListener("input", applyFilters);

  loadDefaultExcel();
});

async function loadDefaultExcel() {
  try {
    const filePath = "data/latest-nq.xlsx";
    const response = await fetch(filePath);

    if (!response.ok) {
      throw new Error("Default Excel file not found.");
    }

    const arrayBuffer = await response.arrayBuffer();

    document.getElementById("systemFileName").textContent = "latest-nq.xlsx";
    document.getElementById("fileName").textContent = "latest-nq.xlsx";
    document.getElementById("fileModified").textContent = "Current system file from /data folder";

    processExcel(arrayBuffer);
    updateDashboardRefreshTime();
  } catch (error) {
    console.error("Error loading default Excel file:", error);
    document.getElementById("fileName").textContent = "No system file found";
    document.getElementById("fileModified").textContent = "N/A";
    document.getElementById("dashboardRefresh").textContent = "N/A";
    document.getElementById("systemFileName").textContent = "No file";
  }
}

function processExcel(data) {
  const workbook = XLSX.read(data, { type: "array" });

  let parsedRows = [];

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const sheetRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    const headerRowIndex = findHeaderRow(sheetRows);

    if (headerRowIndex === -1) continue;

    const headers = sheetRows[headerRowIndex].map(h => normalizeHeader(h));
    const dataRows = sheetRows.slice(headerRowIndex + 1);

    parsedRows = dataRows
      .filter(row => row.some(cell => String(cell).trim() !== ""))
      .map(row => mapRowByHeaders(headers, row));

    if (parsedRows.length > 0) break;
  }

  allRows = parsedRows.map(row => ({
    office: getValue(row, ["office"]),
    reconciler: getValue(row, ["reconciler"]),
    unit: getValue(row, ["unit"]),
    groupId: getValue(row, ["#groupid", "groupid", "groupidnumber", "group id", "# group id", "numbergroupid"]),
    title: getValue(row, ["title"]),
    nqItems: toNumber(getValue(row, ["nqitems", "nqitem", "sumofnqitems", "sum nq items"]))
  }));

  allRows = allRows.filter(row =>
    row.office || row.reconciler || row.unit || row.groupId || row.title || row.nqItems
  );

  console.log("Parsed rows:", allRows.slice(0, 20));

  initializeFilters();
  applyFilters();
}

function initializeFilters() {
  fillSelect("officeFilter", getUniqueSorted(allRows.map(r => r.office)), "All Offices");
  fillSelect("reconcilerFilter", getUniqueSorted(allRows.map(r => r.reconciler)), "All Reconcilers");
}

function updateDependentFilters(changedFilter) {
  const officeFilter = document.getElementById("officeFilter");
  const reconcilerFilter = document.getElementById("reconcilerFilter");

  const selectedOffice = officeFilter.value;
  const selectedReconciler = reconcilerFilter.value;

  if (changedFilter === "office") {
    const rowsForOffice = selectedOffice
      ? allRows.filter(row => row.office === selectedOffice)
      : allRows;

    const reconcilerOptions = getUniqueSorted(rowsForOffice.map(r => r.reconciler));
    fillSelect("reconcilerFilter", reconcilerOptions, "All Reconcilers", selectedReconciler);

    if (selectedReconciler && !reconcilerOptions.includes(selectedReconciler)) {
      reconcilerFilter.value = "";
    }
  }

  if (changedFilter === "reconciler") {
    const rowsForReconciler = selectedReconciler
      ? allRows.filter(row => row.reconciler === selectedReconciler)
      : allRows;

    const officeOptions = getUniqueSorted(rowsForReconciler.map(r => r.office));
    fillSelect("officeFilter", officeOptions, "All Offices", selectedOffice);

    if (selectedOffice && !officeOptions.includes(selectedOffice)) {
      officeFilter.value = "";
    }
  }

  if (!selectedOffice && !selectedReconciler) {
    fillSelect("officeFilter", getUniqueSorted(allRows.map(r => r.office)), "All Offices");
    fillSelect("reconcilerFilter", getUniqueSorted(allRows.map(r => r.reconciler)), "All Reconcilers");
  }
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const normalized = rows[i].map(cell => normalizeHeader(cell));

    const hasOffice = normalized.includes("office");
    const hasReconciler = normalized.includes("reconciler");
    const hasTitle = normalized.includes("title");
    const hasNQ =
      normalized.includes("nqitems") ||
      normalized.includes("nqitem") ||
      normalized.includes("sumofnqitems");

    if (hasOffice && hasReconciler && hasTitle && hasNQ) {
      return i;
    }
  }
  return -1;
}

function mapRowByHeaders(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index];
  });
  return obj;
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/_/g, "")
    .replace(/\./g, "")
    .replace(/-/g, "");
}

function getValue(obj, possibleKeys) {
  for (const key of possibleKeys) {
    if (obj[key] !== undefined && obj[key] !== null && obj[key] !== "") {
      return String(obj[key]).trim();
    }
  }
  return "";
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/\$/g, "")
    .trim();

  const num = Number(cleaned);
  return isNaN(num) ? 0 : num;
}

function fillSelect(selectId, values, defaultText, selectedValue = "") {
  const select = document.getElementById(selectId);
  select.innerHTML = "";

  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = defaultText;
  select.appendChild(defaultOption);

  values.forEach(value => {
    if (!value) return;
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  if (selectedValue && values.includes(selectedValue)) {
    select.value = selectedValue;
  } else {
    select.value = "";
  }
}

function applyFilters() {
  const officeValue = document.getElementById("officeFilter").value.toLowerCase();
  const reconcilerValue = document.getElementById("reconcilerFilter").value.toLowerCase();
  const searchValue = document.getElementById("searchBox").value.toLowerCase();

  const filteredRows = allRows.filter(row => {
    const matchesOffice = !officeValue || row.office.toLowerCase() === officeValue;
    const matchesReconciler = !reconcilerValue || row.reconciler.toLowerCase() === reconcilerValue;
    const matchesSearch =
      !searchValue ||
      row.title.toLowerCase().includes(searchValue) ||
      row.unit.toLowerCase().includes(searchValue) ||
      String(row.groupId).toLowerCase().includes(searchValue);

    return matchesOffice && matchesReconciler && matchesSearch;
  });

  renderDashboard(filteredRows);
}

function renderDashboard(rows) {
  const totalNQ = rows.reduce((sum, row) => sum + row.nqItems, 0);

  document.getElementById("totalNQ").textContent = formatNumber(totalNQ);
  document.getElementById("detailTotalFooter").textContent = formatNumber(totalNQ);
  document.getElementById("reconcilerTotalFooter").textContent = formatNumber(totalNQ);
  document.getElementById("officeTotalFooter").textContent = formatNumber(totalNQ);

  updateGauge(totalNQ);
  renderOfficePieChart(rows);
  renderReconcilerSummary(rows);
  renderOfficeSummary(rows);
  renderDetailTable(rows);
}

function updateGauge(totalNQ) {
  const gaugeMax = Math.max(1000, roundUpToNearest(totalNQ, 100));
  const percentage = gaugeMax > 0 ? Math.min((totalNQ / gaugeMax) * 100, 100) : 0;

  document.getElementById("gaugeFill").style.width = `${percentage}%`;
  document.getElementById("gaugeMax").textContent = formatNumber(gaugeMax);
  document.getElementById("gaugeMiddle").textContent = formatNumber(Math.round(gaugeMax / 2));
}

function renderOfficePieChart(rows) {
  const officeTotals = groupAndSum(rows, "office");
  const labels = officeTotals.map(item => item.key || "(Blank)");
  const data = officeTotals.map(item => item.total);

  const canvas = document.getElementById("officePieChart");
  const ctx = canvas.getContext("2d");

  if (officePieChart) {
    officePieChart.destroy();
  }

  officePieChart = new Chart(ctx, {
    type: "pie",
    data: {
      labels,
      datasets: [{ data }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "right" },
        tooltip: {
          callbacks: {
            label: function (context) {
              const total = data.reduce((a, b) => a + b, 0);
              const value = context.raw;
              const percent = total ? ((value / total) * 100).toFixed(1) : 0;
              return `${context.label}: ${formatNumber(value)} (${percent}%)`;
            }
          }
        }
      }
    }
  });
}

function renderReconcilerSummary(rows) {
  const tbody = document.querySelector("#reconcilerSummaryTable tbody");
  tbody.innerHTML = "";

  const summary = groupAndSum(rows, "reconciler");

  if (summary.length === 0) {
    tbody.innerHTML = `<tr><td>(Blank)</td><td>0</td></tr>`;
    return;
  }

  summary.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.key || "(Blank)")}</td>
      <td>${formatNumber(item.total)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderOfficeSummary(rows) {
  const tbody = document.querySelector("#officeSummaryTable tbody");
  tbody.innerHTML = "";

  const summary = groupAndSum(rows, "office");

  if (summary.length === 0) {
    tbody.innerHTML = `<tr><td>(Blank)</td><td>0</td></tr>`;
    return;
  }

  summary.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.key || "(Blank)")}</td>
      <td>${formatNumber(item.total)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderDetailTable(rows) {
  const tbody = document.querySelector("#detailTable tbody");
  tbody.innerHTML = "";

  const sortedRows = [...rows].sort((a, b) => b.nqItems - a.nqItems);

  if (sortedRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6">No data found.</td></tr>`;
    return;
  }

  sortedRows.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.office)}</td>
      <td>${escapeHtml(row.reconciler)}</td>
      <td>${escapeHtml(row.unit)}</td>
      <td>${escapeHtml(String(row.groupId))}</td>
      <td>${escapeHtml(row.title)}</td>
      <td>${formatNumber(row.nqItems)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function groupAndSum(rows, field) {
  const map = new Map();

  rows.forEach(row => {
    const key = row[field] || "";
    const current = map.get(key) || 0;
    map.set(key, current + row.nqItems);
  });

  return [...map.entries()]
    .map(([key, total]) => ({ key, total }))
    .sort((a, b) => b.total - a.total);
}

function getUniqueSorted(array) {
  return [...new Set(array.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function formatNumber(value) {
  return new Intl.NumberFormat("en-US").format(value || 0);
}

function updateDashboardRefreshTime() {
  const now = new Date();
  const formatted = now.toLocaleString("en-US", {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });

  document.getElementById("dashboardRefresh").textContent = formatted;
  document.getElementById("updateDate").textContent = formatted;
}

function roundUpToNearest(value, nearest) {
  return Math.ceil(value / nearest) * nearest;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}