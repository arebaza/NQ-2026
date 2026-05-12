const converterForm = document.querySelector('#converterForm');
const itemType = document.querySelector('#itemType');
const sourceFile = document.querySelector('#sourceFile');
const convertMessage = document.querySelector('#convertMessage');

const reportForm = document.querySelector('#reportForm');
const reportFiles = document.querySelector('#reportFiles');
const reportMessage = document.querySelector('#reportMessage');
const previewArea = document.querySelector('#previewArea');

const tabButtons = document.querySelectorAll('.tab-button');
const panels = document.querySelectorAll('.panel');

const outputHeaders = {
  team: [
    'Unit',
    'Group ID',
    '# Group ID',
    'Title',
    'Office',
    'Reconciler ID',
    'Supv ID',
    'Reconciler',
    'NQ Items',
    'Last Recon Date',
    'Status',
    'Comments'
  ],
  accounts: [
    'Unit',
    'Group ID',
    '# Group ID',
    'Account',
    'Title',
    'Office',
    'Reconciler ID',
    'Supv ID',
    'Reconciler',
    'NQ Items',
    'Last Recon Date',
    'Status'
  ]
};

const processSettings = {
  nq: {
    expectedEnding: '- NQ',
    alternativeEnding: '- O',
    teamFileName: 'Team Recon Report NQ.xlsx',
    accountsFileName: 'NQ by Accounts.xlsx',
    label: 'NQ Items'
  },
  outstanding: {
    expectedEnding: '- OU',
    alternativeEnding: null,
    teamFileName: 'Team Recon Report Outstanding.xlsx',
    accountsFileName: 'Outstanding by Accounts.xlsx',
    label: 'Outstanding Items'
  }
};

const monthNumbers = {
  // English months and abbreviations
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,

  // Spanish months and abbreviations
  ene: 0,
  enero: 0,
  febrero: 1,
  mar: 2,
  marzo: 2,
  abr: 3,
  abril: 3,
  mayo: 4,
  jun: 5,
  junio: 5,
  jul: 6,
  julio: 6,
  ago: 7,
  agosto: 7,
  set: 8,
  sep: 8,
  sept: 8,
  septiembre: 8,
  setiembre: 8,
  oct: 9,
  octubre: 9,
  nov: 10,
  noviembre: 10,
  dic: 11,
  diciembre: 11
};

setupTabs();

converterForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  clearMessage(convertMessage);

  const file = sourceFile.files[0];
  const selectedProcess = processSettings[itemType.value];

  if (!file) {
    showMessage(convertMessage, 'Please upload an Excel file first.', 'error');
    return;
  }

  if (!isCorrectFileName(file.name, selectedProcess)) {
    const endingText = selectedProcess.alternativeEnding
      ? `${selectedProcess.expectedEnding}.xlsx or ${selectedProcess.alternativeEnding}.xlsx`
      : `${selectedProcess.expectedEnding}.xlsx`;

    showMessage(
      convertMessage,
      `The selected process is ${selectedProcess.label}. Please upload the original file that ends in ${endingText}.`,
      'error'
    );
    return;
  }

  try {
    const workbook = await readWorkbook(file);
    const rows = extractRows(workbook);

    if (rows.length === 0) {
      showMessage(convertMessage, 'No valid data rows were found in the uploaded file.', 'error');
      return;
    }

    const teamRows = buildTeamReconRows(rows);
    const accountRows = buildAccountRows(rows);

    downloadWorkbook(selectedProcess.teamFileName, outputHeaders.team, teamRows);
    downloadWorkbook(selectedProcess.accountsFileName, outputHeaders.accounts, accountRows);

    showMessage(convertMessage, `Success. ${rows.length} rows were converted and two Excel files were downloaded.`, 'success');
  } catch (error) {
    console.error(error);
    showMessage(convertMessage, `Error: ${error.message}`, 'error');
  }
});

reportForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  clearMessage(reportMessage);
  previewArea.innerHTML = '';

  const files = Array.from(reportFiles.files || []);

  if (files.length < 2) {
    showMessage(reportMessage, 'Please upload at least two dated files to compare performance.', 'error');
    return;
  }

  try {
    const dailyFiles = await readDailyFiles(files);
    const reportData = buildWeeklyPerformanceReport(dailyFiles);

    if (reportData.dailyPerformance.length === 0) {
      showMessage(reportMessage, 'No day-to-day comparison was created. Please upload at least two dates for the same item type.', 'warning');
      return;
    }

    downloadPerformanceReport(reportData);
    renderPreview(reportData.outstandingRanking.slice(0, 10));

    showMessage(
      reportMessage,
      `Success. Weekly Performance Report.xlsx was created with ${reportData.dailyPerformance.length} daily comparison rows.`,
      'success'
    );
  } catch (error) {
    console.error(error);
    showMessage(reportMessage, `Error: ${error.message}`, 'error');
  }
});

function setupTabs() {
  tabButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const targetId = button.dataset.tab;

      tabButtons.forEach((item) => item.classList.remove('active'));
      panels.forEach((panel) => panel.classList.remove('active'));

      button.classList.add('active');
      document.querySelector(`#${targetId}`).classList.add('active');
    });
  });
}

function isCorrectFileName(fileName, settings) {
  const cleanName = fileName.toLowerCase().replace(/\.xlsx$|\.xls$/i, '').trim();
  const expected = settings.expectedEnding.toLowerCase();
  const alternative = settings.alternativeEnding?.toLowerCase();

  return cleanName.endsWith(expected) || (alternative && cleanName.endsWith(alternative));
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {
          type: 'array',
          cellDates: false,
          raw: false
        });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('The file could not be read.'));
    reader.readAsArrayBuffer(file);
  });
}

function extractRows(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const sheetRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: false
  });

  const headerRowIndex = findHeaderRow(sheetRows);
  if (headerRowIndex === -1) {
    throw new Error('The header row was not found. The file must include Unit, Group ID, Title, Office, Reconciler, NQ Items, Last Recon Date, and Status.');
  }

  const headers = sheetRows[headerRowIndex].map(normalizeHeader);
  const dataRows = sheetRows.slice(headerRowIndex + 1);

  return dataRows
    .map((row) => mapRowToObject(headers, row))
    .filter((row) => row.Unit && row['Group ID'] && row.Title);
}

function findHeaderRow(rows) {
  return rows.findIndex((row) => {
    const normalized = row.map(normalizeHeader);
    return normalized.includes('Unit') &&
      normalized.includes('Group ID') &&
      normalized.includes('Title') &&
      normalized.includes('Office') &&
      normalized.includes('Reconciler') &&
      normalized.includes('NQ Items') &&
      normalized.includes('Last Recon Date') &&
      normalized.includes('Status');
  });
}

function normalizeHeader(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function mapRowToObject(headers, row) {
  const object = {};

  headers.forEach((header, index) => {
    if (header) {
      object[header] = cleanCell(row[index]);
    }
  });

  return object;
}

function cleanCell(value) {
  const text = String(value ?? '').trim();

  if (/^\d+\.0$/.test(text)) {
    return text.replace('.0', '');
  }

  return text;
}

function buildTeamReconRows(rows) {
  return rows.map((row) => ({
    Unit: row.Unit,
    'Group ID': row['Group ID'],
    '# Group ID': row['Group ID'],
    Title: row.Title,
    Office: row.Office,
    'Reconciler ID': row['Reconciler ID'],
    'Supv ID': row['Supv ID'],
    Reconciler: row.Reconciler,
    'NQ Items': toNumber(row['NQ Items']),
    'Last Recon Date': row['Last Recon Date'],
    Status: row.Status,
    Comments: row.Comments || ''
  }));
}

function buildAccountRows(rows) {
  return rows.map((row) => ({
    Unit: row.Unit,
    'Group ID': row['Group ID'],
    '# Group ID': row['Group ID'],
    Account: extractAccount(row.Title),
    Title: row.Title,
    Office: row.Office,
    'Reconciler ID': row['Reconciler ID'],
    'Supv ID': row['Supv ID'],
    Reconciler: row.Reconciler,
    'NQ Items': toNumber(row['NQ Items']),
    'Last Recon Date': row['Last Recon Date'],
    Status: row.Status
  }));
}

function extractAccount(title) {
  const match = String(title || '').trim().match(/^\d{4}/);
  return match ? match[0] : '';
}

async function readDailyFiles(files) {
  const dailyFiles = [];

  for (const file of files) {
    const dateInfo = parseDateFromFileName(file.name);
    const reportType = getReportTypeFromFileName(file.name);

    if (!dateInfo || !reportType) {
      throw new Error(`The file name was not recognized: ${file.name}. Use names like NQ by Accounts 7 May 2026.xlsx, NQ by Accounts 24 Abr 2026.xlsx, or Outstanding by Accounts 7 May 2026.xlsx.`);
    }

    const workbook = await readWorkbook(file);
    const rows = extractRows(workbook);
    const totalsByReconciler = summarizeByReconciler(rows);

    dailyFiles.push({
      fileName: file.name,
      reportType,
      date: dateInfo.date,
      dateKey: dateInfo.key,
      displayDate: dateInfo.display,
      totalsByReconciler,
      totalItems: Array.from(totalsByReconciler.values()).reduce((sum, item) => sum + item.totalItems, 0)
    });
  }

  return dailyFiles.sort((a, b) => a.date - b.date || a.reportType.localeCompare(b.reportType));
}

function parseDateFromFileName(fileName) {
  const cleanName = fileName.replace(/\.xlsx$|\.xls$/i, '').trim();
  const match = cleanName.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);

  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = monthNumbers[match[2].toLowerCase()];
  const year = Number(match[3]);

  if (Number.isNaN(day) || month === undefined || Number.isNaN(year)) {
    return null;
  }

  const date = new Date(year, month, day);
  const key = date.toISOString().slice(0, 10);

  return {
    date,
    key,
    display: `${day} ${capitalize(match[2])} ${year}`
  };
}

function getReportTypeFromFileName(fileName) {
  const lowerName = fileName.toLowerCase();

  if (lowerName.includes('outstanding') || lowerName.includes('oustanding')) {
    return 'Outstanding';
  }

  if (lowerName.includes('nq')) {
    return 'NQ';
  }

  return null;
}

function summarizeByReconciler(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const reconciler = row.Reconciler || 'Blank Reconciler';
    const office = row.Office || '';
    const reconcilerId = row['Reconciler ID'] || '';
    const itemCount = toNumber(row['NQ Items']);

    if (!map.has(reconciler)) {
      map.set(reconciler, {
        reconciler,
        reconcilerId,
        offices: new Set(),
        accountRows: 0,
        totalItems: 0
      });
    }

    const record = map.get(reconciler);
    record.accountRows += 1;
    record.totalItems += itemCount;

    if (office) {
      record.offices.add(office);
    }
  });

  return map;
}

function buildWeeklyPerformanceReport(dailyFiles) {
  const adjustedFiles = buildSeparatedDailyFiles(dailyFiles);
  const dailyPerformance = [];
  const dailyTotals = [];
  const rankingMap = new Map();
  const groupedByType = groupBy(adjustedFiles, 'reportType');

  Object.entries(groupedByType).forEach(([reportType, files]) => {
    const sortedFiles = files.sort((a, b) => a.date - b.date);

    sortedFiles.forEach((file) => {
      dailyTotals.push({
        'Item Type': reportType,
        Date: file.displayDate,
        'Total Items': file.totalItems,
        'Calculation': file.calculation,
        'File Name': file.fileName
      });
    });

    for (let index = 1; index < sortedFiles.length; index += 1) {
      const previousFile = sortedFiles[index - 1];
      const currentFile = sortedFiles[index];
      const reconcilers = new Set([
        ...previousFile.totalsByReconciler.keys(),
        ...currentFile.totalsByReconciler.keys()
      ]);

      reconcilers.forEach((reconciler) => {
        const previous = previousFile.totalsByReconciler.get(reconciler) || emptyReconciler(reconciler);
        const current = currentFile.totalsByReconciler.get(reconciler) || emptyReconciler(reconciler);
        const netReduction = previous.totalItems - current.totalItems;
        const reducedItems = Math.max(netReduction, 0);
        const increasedItems = Math.max(current.totalItems - previous.totalItems, 0);

        dailyPerformance.push({
          'Item Type': reportType,
          'From Date': previousFile.displayDate,
          'To Date': currentFile.displayDate,
          Reconciler: reconciler,
          'Reconciler ID': current.reconcilerId || previous.reconcilerId,
          Offices: formatOffices(current.offices, previous.offices),
          'Previous Items': previous.totalItems,
          'Current Items': current.totalItems,
          'Net Reduction': netReduction,
          'Reduced Items': reducedItems,
          'Increased Items': increasedItems,
          'Previous Account Rows': previous.accountRows,
          'Current Account Rows': current.accountRows,
          'Calculation': currentFile.calculation
        });

        const key = `${reportType}|${reconciler}`;
        if (!rankingMap.has(key)) {
          rankingMap.set(key, {
            'Item Type': reportType,
            Reconciler: reconciler,
            'Reconciler ID': current.reconcilerId || previous.reconcilerId,
            'Total Reduced Items': 0,
            'Total Increased Items': 0,
            'Week Start Items': null,
            'Week End Items': null,
            'Week Net Reduction': 0,
            'Calculation': currentFile.calculation
          });
        }

        const ranking = rankingMap.get(key);
        ranking['Total Reduced Items'] += reducedItems;
        ranking['Total Increased Items'] += increasedItems;
      });
    }

    updateWeekStartEndRanking(rankingMap, reportType, sortedFiles);
  });

  const weeklyRanking = Array.from(rankingMap.values())
    .sort((a, b) => b['Total Reduced Items'] - a['Total Reduced Items'] || b['Week Net Reduction'] - a['Week Net Reduction']);

  const outstandingRanking = weeklyRanking
    .filter((row) => row['Item Type'] === 'Outstanding Items without NQ')
    .map((row, index) => ({ Rank: index + 1, ...row }));

  const nqRanking = weeklyRanking
    .filter((row) => row['Item Type'] === 'NQ Items')
    .map((row, index) => ({ Rank: index + 1, ...row }));

  return {
    dailyPerformance: dailyPerformance.sort((a, b) => a['Item Type'].localeCompare(b['Item Type']) || a.Reconciler.localeCompare(b.Reconciler)),
    dailyTotals,
    outstandingRanking,
    nqRanking,
    weeklyRanking: weeklyRanking.map((row, index) => ({ Rank: index + 1, ...row }))
  };
}

function buildSeparatedDailyFiles(dailyFiles) {
  const groupedByDate = groupBy(dailyFiles, 'dateKey');
  const separatedFiles = [];

  Object.values(groupedByDate).forEach((filesForDate) => {
    const nqFile = filesForDate.find((file) => file.reportType === 'NQ');
    const outstandingFile = filesForDate.find((file) => file.reportType === 'Outstanding');

    if (nqFile) {
      separatedFiles.push({
        ...nqFile,
        reportType: 'NQ Items',
        calculation: 'NQ only'
      });
    }

    if (outstandingFile) {
      const adjustedOutstandingMap = subtractNqFromOutstanding(
        outstandingFile.totalsByReconciler,
        nqFile ? nqFile.totalsByReconciler : new Map()
      );

      separatedFiles.push({
        ...outstandingFile,
        reportType: 'Outstanding Items without NQ',
        totalsByReconciler: adjustedOutstandingMap,
        totalItems: Array.from(adjustedOutstandingMap.values()).reduce((sum, item) => sum + item.totalItems, 0),
        calculation: 'Outstanding total minus NQ items'
      });
    }
  });

  return separatedFiles.sort((a, b) => a.date - b.date || a.reportType.localeCompare(b.reportType));
}

function subtractNqFromOutstanding(outstandingMap, nqMap) {
  const result = new Map();
  const reconcilers = new Set([
    ...outstandingMap.keys(),
    ...nqMap.keys()
  ]);

  reconcilers.forEach((reconciler) => {
    const outstanding = outstandingMap.get(reconciler) || emptyReconciler(reconciler);
    const nq = nqMap.get(reconciler) || emptyReconciler(reconciler);
    const adjustedTotal = Math.max(outstanding.totalItems - nq.totalItems, 0);

    result.set(reconciler, {
      reconciler,
      reconcilerId: outstanding.reconcilerId || nq.reconcilerId,
      offices: new Set([...(outstanding.offices || []), ...(nq.offices || [])]),
      accountRows: Math.max(outstanding.accountRows - nq.accountRows, 0),
      totalItems: adjustedTotal
    });
  });

  return result;
}

function updateWeekStartEndRanking(rankingMap, reportType, sortedFiles) {
  if (sortedFiles.length === 0) {
    return;
  }

  const firstFile = sortedFiles[0];
  const lastFile = sortedFiles[sortedFiles.length - 1];
  const reconcilers = new Set([
    ...firstFile.totalsByReconciler.keys(),
    ...lastFile.totalsByReconciler.keys()
  ]);

  reconcilers.forEach((reconciler) => {
    const first = firstFile.totalsByReconciler.get(reconciler) || emptyReconciler(reconciler);
    const last = lastFile.totalsByReconciler.get(reconciler) || emptyReconciler(reconciler);
    const key = `${reportType}|${reconciler}`;

    if (!rankingMap.has(key)) {
      rankingMap.set(key, {
        'Item Type': reportType,
        Reconciler: reconciler,
        'Reconciler ID': last.reconcilerId || first.reconcilerId,
        'Total Reduced Items': 0,
        'Total Increased Items': 0,
        'Week Start Items': null,
        'Week End Items': null,
        'Week Net Reduction': 0
      });
    }

    const ranking = rankingMap.get(key);
    ranking['Week Start Items'] = first.totalItems;
    ranking['Week End Items'] = last.totalItems;
    ranking['Week Net Reduction'] = first.totalItems - last.totalItems;
  });
}

function emptyReconciler(reconciler) {
  return {
    reconciler,
    reconcilerId: '',
    offices: new Set(),
    accountRows: 0,
    totalItems: 0
  };
}

function formatOffices(currentOffices, previousOffices) {
  const offices = new Set([...(previousOffices || []), ...(currentOffices || [])]);
  return Array.from(offices).sort().join(', ');
}

function groupBy(items, key) {
  return items.reduce((groups, item) => {
    const groupKey = item[key];
    groups[groupKey] = groups[groupKey] || [];
    groups[groupKey].push(item);
    return groups;
  }, {});
}

function downloadPerformanceReport(reportData) {
  const workbook = XLSX.utils.book_new();

  appendSheet(workbook, 'Outstanding without NQ Ranking', reportData.outstandingRanking);
  appendSheet(workbook, 'NQ Ranking', reportData.nqRanking);
  appendSheet(workbook, 'Daily Performance', reportData.dailyPerformance);
  appendSheet(workbook, 'Daily Totals', reportData.dailyTotals);
  appendSheet(workbook, 'All Weekly Ranking', reportData.weeklyRanking);

  XLSX.writeFile(workbook, 'Weekly Performance Report.xlsx');
}

function appendSheet(workbook, sheetName, rows) {
  const safeRows = rows.length > 0 ? rows : [{ Message: 'No data found' }];
  const worksheet = XLSX.utils.json_to_sheet(safeRows);
  const headers = Object.keys(safeRows[0]);

  worksheet['!cols'] = headers.map((header) => ({ wch: getReportColumnWidth(header) }));
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
}

function downloadWorkbook(fileName, headers, rowObjects) {
  const worksheetRows = [headers];

  rowObjects.forEach((rowObject) => {
    worksheetRows.push(headers.map((header) => rowObject[header] ?? ''));
  });

  const worksheet = XLSX.utils.aoa_to_sheet(worksheetRows);
  worksheet['!cols'] = headers.map((header) => ({ wch: getColumnWidth(header) }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'sheet1');
  XLSX.writeFile(workbook, fileName);
}

function getColumnWidth(header) {
  const widths = {
    Unit: 10,
    'Group ID': 12,
    '# Group ID': 12,
    Account: 10,
    Title: 34,
    Office: 12,
    'Reconciler ID': 15,
    'Supv ID': 12,
    Reconciler: 24,
    'NQ Items': 12,
    'Last Recon Date': 16,
    Status: 18,
    Comments: 28
  };

  return widths[header] || 16;
}

function getReportColumnWidth(header) {
  const widths = {
    Rank: 8,
    'Item Type': 14,
    'From Date': 18,
    'To Date': 18,
    Date: 18,
    Reconciler: 26,
    'Reconciler ID': 15,
    Offices: 28,
    'Previous Items': 16,
    'Current Items': 16,
    'Net Reduction': 16,
    'Reduced Items': 16,
    'Increased Items': 16,
    'Total Reduced Items': 20,
    'Total Increased Items': 20,
    'Week Start Items': 18,
    'Week End Items': 18,
    'Week Net Reduction': 20,
    'Calculation': 34,
    'File Name': 40
  };

  return widths[header] || 18;
}

function renderPreview(rows) {
  if (rows.length === 0) {
    previewArea.innerHTML = '<p class="small-note">No Outstanding without NQ ranking data found.</p>';
    return;
  }

  const headers = ['Rank', 'Reconciler', 'Total Reduced Items', 'Week Start Items', 'Week End Items', 'Week Net Reduction'];
  const headerHtml = headers.map((header) => `<th>${header}</th>`).join('');
  const bodyHtml = rows.map((row) => {
    const cells = headers.map((header) => `<td>${row[header] ?? ''}</td>`).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  previewArea.innerHTML = `
    <h3>Top Outstanding Reduction Preview — excluding NQ</h3>
    <table>
      <thead><tr>${headerHtml}</tr></thead>
      <tbody>${bodyHtml}</tbody>
    </table>
  `;
}

function toNumber(value) {
  const number = Number(String(value ?? '0').replace(/,/g, '').trim());
  return Number.isFinite(number) ? number : 0;
}

function capitalize(text) {
  const clean = String(text || '').toLowerCase();
  return clean.charAt(0).toUpperCase() + clean.slice(1);
}

function showMessage(element, text, type) {
  element.textContent = text;
  element.className = `message ${type}`;
}

function clearMessage(element) {
  element.textContent = '';
  element.className = 'message';
}
