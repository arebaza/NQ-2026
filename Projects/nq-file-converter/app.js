const inputFile = document.querySelector('#inputFile');
const convertBtn = document.querySelector('#convertBtn');
const message = document.querySelector('#message');

let selectedFile = null;

const TEAM_RECON_HEADERS = [
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
  'Comments',
  'Submit Date',
  'Approve Date'
];

const NQ_BY_ACCOUNTS_HEADERS = [
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
  'Status',
  'Comments',
  'Submit Date',
  'Approve Date'
];

inputFile.addEventListener('change', (event) => {
  selectedFile = event.target.files[0];
  convertBtn.disabled = !selectedFile;
  clearMessage();
});

convertBtn.addEventListener('click', async () => {
  try {
    if (!selectedFile) {
      showMessage('Please select the - O file first.', 'error');
      return;
    }

    const workbook = await readWorkbook(selectedFile);
    const sourceRows = getSourceRows(workbook);

    if (sourceRows.length === 0) {
      showMessage('No valid rows were found in the file.', 'error');
      return;
    }

    const teamReconRows = buildTeamReconRows(sourceRows);
    const nqByAccountsRows = buildNqByAccountsRows(sourceRows);

    downloadWorkbook('Team Recon Report NQ.xlsx', TEAM_RECON_HEADERS, teamReconRows);
    downloadWorkbook('NQ by Accounts.xlsx', NQ_BY_ACCOUNTS_HEADERS, nqByAccountsRows);

    showMessage('Success. The two converted files were downloaded.', 'success');
  } catch (error) {
    console.error(error);
    showMessage(`Error: ${error.message}`, 'error');
  }
});

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('The file could not be read.'));
    reader.readAsArrayBuffer(file);
  });
}

function getSourceRows(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: true
  });

  const headerRowIndex = rows.findIndex((row) => normalize(row[0]) === 'unit' && normalize(row[1]) === 'group id');

  if (headerRowIndex === -1) {
    throw new Error('The header row was not found. The file must include Unit and Group ID columns.');
  }

  const headers = rows[headerRowIndex].map((header) => String(header).trim());
  const dataRows = rows.slice(headerRowIndex + 1);

  return dataRows
    .filter((row) => row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== ''))
    .map((row) => rowToObject(headers, row));
}

function rowToObject(headers, row) {
  const item = {};

  headers.forEach((header, index) => {
    if (header) {
      item[header] = row[index] ?? '';
    }
  });

  return item;
}

function buildTeamReconRows(sourceRows) {
  return sourceRows.map((row) => ({
    'Unit': clean(row['Unit']),
    'Group ID': clean(row['Group ID']),
    '# Group ID': clean(row['Group ID']),
    'Title': clean(row['Title']),
    'Office': clean(row['Office']),
    'Reconciler ID': clean(row['Reconciler ID']),
    'Supv ID': clean(row['Supv ID']),
    'Reconciler': clean(row['Reconciler']),
    'NQ Items': row['NQ Items'] || 0,
    'Last Recon Date': row['Last Recon Date'] || '',
    'Status': clean(row['Status']),
    'Comments': clean(row['Comments']),
    'Submit Date': row['Submit Date'] || '',
    'Approve Date': row['Approve Date'] || ''
  }));
}

function buildNqByAccountsRows(sourceRows) {
  return sourceRows.map((row) => {
    const title = clean(row['Title']);

    return {
      'Unit': clean(row['Unit']),
      'Group ID': clean(row['Group ID']),
      '# Group ID': clean(row['Group ID']),
      'Account': getAccountFromTitle(title),
      'Title': title,
      'Office': clean(row['Office']),
      'Reconciler ID': clean(row['Reconciler ID']),
      'Supv ID': clean(row['Supv ID']),
      'Reconciler': clean(row['Reconciler']),
      'NQ Items': row['NQ Items'] || 0,
      'Last Recon Date': row['Last Recon Date'] || '',
      'Status': clean(row['Status']),
      'Comments': clean(row['Comments']),
      'Submit Date': row['Submit Date'] || '',
      'Approve Date': row['Approve Date'] || ''
    };
  });
}

function getAccountFromTitle(title) {
  const match = String(title).match(/^\s*(\d{4})/);
  return match ? match[1] : '';
}

function downloadWorkbook(fileName, headers, rows) {
  const formattedRows = rows.map((row) => {
    const orderedRow = {};
    headers.forEach((header) => {
      orderedRow[header] = row[header] ?? '';
    });
    return orderedRow;
  });

  const worksheet = XLSX.utils.json_to_sheet(formattedRows, { header: headers });

  worksheet['!cols'] = headers.map((header) => ({
    wch: Math.max(header.length + 2, 14)
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'sheet1');
  XLSX.writeFile(workbook, fileName);
}

function normalize(value) {
  return String(value || '').trim().toLowerCase();
}

function clean(value) {
  return String(value ?? '').trim();
}

function showMessage(text, type) {
  message.textContent = text;
  message.className = `message ${type}`;
}

function clearMessage() {
  message.textContent = '';
  message.className = 'message';
}
