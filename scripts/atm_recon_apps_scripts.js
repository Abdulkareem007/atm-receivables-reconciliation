/**
 * ================================================================
 *  ATM RECEIVABLES RECONCILIATION — Google Apps Script
 *  Version 1.0  |  Abeokuta Branch  |  All 5 GL Accounts
 * ================================================================
 *
 *  HOW TO USE (every reconciliation period):
 *  1. Copy GL Activity Report (.xlsx) files into your "GL_Reports"
 *     folder in Google Drive
 *  2. Open this Google Sheet
 *  3. Click:  🏧 ATM Recon  →  ▶ Run Reconciliation
 *  4. Check the SUMMARY sheet — all rows must show ✅ BALANCED
 *
 * ================================================================
 */


// ════════════════════════════════════════════════════════════════
//  CONFIGURATION  —  Only edit this section if GL accounts change
// ════════════════════════════════════════════════════════════════

const CFG = {

  // Name of the Google Drive folder where you drop GL files
  GL_FOLDER_NAME  : 'GL_Reports',

  // Name of the folder where processed files are moved after running
  ARCHIVE_FOLDER  : 'GL_Archive',

  // Maps GL account number (read from inside each file) → sheet name
  // Update this if accounts are added, removed, or renamed
  GL_SHEET_MAP : {
    '119110010' : '16436-119110010',   // ISW ATM Settlement
    '119110038' : '16484-119110038',   // MC Domestic ATM
    '119130021' : '16533-119130021',   // Appzone ZS ATM
    '119110026' : '16459-119110026',   // VISA / V-Pay
    '119110093' : '119110093'          // AFRIGO
  },

  // Transaction codes treated as valid (others are skipped)
  VALID_TRN_CODES  : ['ATI', 'EDB', 'ESR', 'ESS'],

  // Keywords used to find and remove old footer rows on each sheet
  FOOTER_KEYWORDS  : ['PROOF BALANCE', 'SYSTEM BALANCE', 'DIFFERENCE']
};


// ════════════════════════════════════════════════════════════════
//  MENU  —  Creates the "🏧 ATM Recon" menu when the sheet opens
// ════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏧 ATM Recon')
    .addItem('▶  Run Reconciliation',  'runReconciliation')
    .addSeparator()
    .addItem('📋  View Last Run Log',  'viewLogs')
    .addToUi();
}


// ════════════════════════════════════════════════════════════════
//  MAIN  —  Entry point called by the menu
// ════════════════════════════════════════════════════════════════

function runReconciliation() {
  const ui  = SpreadsheetApp.getUi();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // ── 1. Find the GL_Reports folder in Google Drive ──────────────
  const folderIter = DriveApp.getFoldersByName(CFG.GL_FOLDER_NAME);
  if (!folderIter.hasNext()) {
    ui.alert(
      '❌  Folder Not Found',
      'Could not find a folder named "GL_Reports" in your Google Drive.\n\n' +
      'Please create it and add your GL Activity Report (.xlsx) files, then try again.',
      ui.ButtonSet.OK
    );
    return;
  }
  const glFolder = folderIter.next();

  // ── 2. Find or create the Archive folder ───────────────────────
  const archIter   = DriveApp.getFoldersByName(CFG.ARCHIVE_FOLDER);
  const archFolder = archIter.hasNext()
    ? archIter.next()
    : DriveApp.createFolder(CFG.ARCHIVE_FOLDER);

  // ── 3. Collect .xlsx files from GL_Reports ─────────────────────
  const glFiles  = [];
  const fileIter = glFolder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  while (fileIter.hasNext()) glFiles.push(fileIter.next());

  if (glFiles.length === 0) {
    ui.alert(
      '⚠️  No Files Found',
      'No .xlsx files were found in the "GL_Reports" folder.\n\n' +
      'Please add your GL Activity Report files and try again.',
      ui.ButtonSet.OK
    );
    return;
  }

  log.push(`──────────────────────────────────────────`);
  log.push(`Run started: ${new Date().toLocaleString()}`);
  log.push(`Files found: ${glFiles.length}`);
  log.push(`──────────────────────────────────────────`);

  const results = [];

  // ── 4. Process each GL file ────────────────────────────────────
  for (const file of glFiles) {
    log.push(`\n📄 Processing: ${file.getName()}`);
    try {
      const result = processOneGLFile(file, ss, log);
      results.push(result);
      // Move the file to Archive after successful processing
      archFolder.addFile(file);
      glFolder.removeFile(file);
      log.push(`   ✓ Moved to GL_Archive folder.`);
    } catch (err) {
      log.push(`   ✗ ERROR: ${err.message}`);
      results.push({
        glAccount : file.getName(),
        sheetName : '—',
        rowsAdded : 0,
        proofBal  : 'N/A',
        sysBal    : 'N/A',
        diff      : 'N/A',
        status    : `❌ ${err.message}`
      });
    }
  }

  // ── 5. Write the Summary sheet ─────────────────────────────────
  updateSummarySheet(ss, results);

  // ── 6. Save the run log for later viewing ─────────────────────
  log.push(`\n──────────────────────────────────────────`);
  log.push(`Run finished: ${new Date().toLocaleString()}`);
  PropertiesService.getScriptProperties().setProperty('LAST_LOG', log.join('\n'));

  // ── 7. Show completion message ─────────────────────────────────
  const allOk = results.every(r => String(r.status).startsWith('✅'));
  ui.alert(
    allOk ? '✅  Reconciliation Complete' : '⚠️  Done — Check Issues',
    allOk
      ? `All ${results.length} GL account(s) reconciled with zero difference.\n\nCheck the SUMMARY sheet for details.`
      : `Completed but some accounts need attention.\n\nCheck the SUMMARY sheet and use "View Last Run Log" from the menu for full details.`,
    ui.ButtonSet.OK
  );
}


// ════════════════════════════════════════════════════════════════
//  PROCESS ONE GL FILE
// ════════════════════════════════════════════════════════════════

function processOneGLFile(file, ss, log) {

  // Convert the .xlsx file to a temporary Google Sheet so we can read it
  const converted = Drive.Files.create(
    { name: '_TEMP_RECON_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
    file.getBlob(),
    { convert: true }
  );

  try {
    const tempSheet = SpreadsheetApp.openById(converted.id).getSheets()[0];
    const rawData   = tempSheet.getDataRange().getValues();

    // ── Find which GL account this file belongs to ─────────────
    const glAccount = findGLAccount(rawData);
    if (!glAccount) throw new Error('GL account number not found in file');
    log.push(`   GL Account  : ${glAccount}`);

    // ── Match to the correct sheet in the master workbook ───────
    const sheetName   = CFG.GL_SHEET_MAP[glAccount];
    if (!sheetName)   throw new Error(`No sheet mapping configured for GL: ${glAccount}`);

    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) throw new Error(`Sheet "${sheetName}" not found in this workbook`);

    // ── Read data from the GL file ──────────────────────────────
    const sysBal      = findSystemClosingBalance(rawData);
    const transactions = extractTransactions(rawData, log);

    log.push(`   Transactions: ${transactions.length} rows`);
    log.push(`   System Bal  : ${formatNumber(sysBal)}`);

    // ── Update the target sheet ─────────────────────────────────
    removeFooterRows(targetSheet);
    const lastDataRow = appendTransactions(targetSheet, transactions);
    addFooterRows(targetSheet, lastDataRow, sysBal);

    // ── Verify the balance ──────────────────────────────────────
    SpreadsheetApp.flush();
    const proofBal = getProofBalance(targetSheet, lastDataRow);
    const diff     = Math.round((proofBal - sysBal) * 100) / 100;

    log.push(`   Proof Bal   : ${formatNumber(proofBal)}`);
    log.push(`   Difference  : ${diff}`);
    log.push(`   Status      : ${Math.abs(diff) < 0.01 ? '✅ BALANCED' : '⚠️ DIFFERENCE = ' + diff}`);

    return {
      glAccount,
      sheetName,
      rowsAdded : transactions.length,
      proofBal,
      sysBal,
      diff,
      status    : Math.abs(diff) < 0.01 ? '✅ BALANCED' : `⚠️  DIFF = ${formatNumber(diff)}`
    };

  } finally {
    // Always delete the temporary Google Sheet — regardless of success or failure
    Drive.Files.remove(converted.id);
  }
}


// ════════════════════════════════════════════════════════════════
//  FIND GL ACCOUNT NUMBER (reads the file header section)
// ════════════════════════════════════════════════════════════════

function findGLAccount(data) {
  for (let r = 0; r < Math.min(data.length, 20); r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim().toLowerCase().includes('account number')) {
        // The account value is in the next non-empty cell on the same row
        for (let k = c + 1; k < data[r].length; k++) {
          const val = String(data[r][k]).trim();
          if (val && val !== 'null') {
            const m = val.match(/(\d{9})/);  // 9-digit GL number
            if (m) return m[1];
          }
        }
      }
    }
  }
  return null;
}


// ════════════════════════════════════════════════════════════════
//  FIND HEADER ROW & BUILD COLUMN INDEX MAP
// ════════════════════════════════════════════════════════════════

function findHeaderRow(data) {
  for (let r = 0; r < data.length; r++) {
    const upperRow = data[r].map(v => String(v).trim().toUpperCase());
    // Confirm it is the real header row (must have all three of these)
    if (upperRow.includes('CREATE DATE') &&
        upperRow.includes('TRN CODE') &&
        upperRow.includes('DESCRIPTION')) {
      const colMap = {};
      upperRow.forEach((name, idx) => {
        // Normalise name: replace spaces and dots with underscores
        const key = name.replace(/[\s.]+/g, '_');
        if (key) colMap[key] = idx;
      });
      return { headerRowIdx: r, colMap };
    }
  }
  return null;
}


// ════════════════════════════════════════════════════════════════
//  EXTRACT TRANSACTIONS FROM GL FILE
// ════════════════════════════════════════════════════════════════

function extractTransactions(data, log) {
  const found = findHeaderRow(data);
  if (!found) {
    if (log) log.push('   ⚠ Could not locate header row in GL file');
    return [];
  }

  const { headerRowIdx, colMap } = found;
  const transactions = [];

  // Transactions start 2 rows after the header
  // (row +1 is blank, row +2 is the Opening Balance row — both skipped)
  for (let r = headerRowIdx + 2; r < data.length; r++) {
    const row     = data[r];
    const trnCode = String(row[colMap['TRN_CODE']] || '').trim().toUpperCase();

    // Skip rows that are not valid transaction types
    if (!CFG.VALID_TRN_CODES.includes(trnCode)) continue;

    const description = String(row[colMap['DESCRIPTION']] || '').trim();
    const amount      = toNum(row[colMap['AMOUNT']]);
    const debit       = toNum(row[colMap['DEBIT']]);
    const credit      = toNum(row[colMap['CREDIT']]);

    // Skip rows with no monetary movement
    if (debit === 0 && credit === 0) continue;

    transactions.push({
      rrn         : extractRRN(description),
      date1       : fmtDate(row[colMap['CREATE_DATE']]),
      date2       : fmtDate(row[colMap['EFFECTIVE_DATE']]),
      trnCode,
      description,
      amount,
      dr          : debit,
      cr          : credit,
      poster      : String(row[colMap['POSTER']] || '').trim(),
      branch      : String(row[colMap['BRANCH']] || '').trim()
    });
  }

  return transactions;
}


// ════════════════════════════════════════════════════════════════
//  EXTRACT RRN FROM DESCRIPTION
// ════════════════════════════════════════════════════════════════

function extractRRN(desc) {
  if (!desc || typeof desc !== 'string') return '';

  // Strip leading asterisks and reversal prefixes
  let d = desc.trim()
              .replace(/^\*+/, '')
              .replace(/^(?:RSVL|RVSL)\s+/i, '')
              .trim();
  let m;

  // Old ATI format:  ISW|GL|RRN|...  or  MC/VC/VGATE|GL|RRN|...
  m = d.match(/^(?:ISW|MC|VC|VGATE)\|\d+\|0*(\d{6,})\|/i);
  if (m) return parseInt(m[1], 10);

  // New ATI format:  RRN|GL|TERMINAL|...  (starts with digits then pipe)
  m = d.match(/^0*(\d{6,})\|/);
  if (m) return parseInt(m[1], 10);

  // ZS ATM format:   ZS ATM-RRN-...
  m = d.match(/^ZS ATM[-\s]0*(\d{6,})/i);
  if (m) return parseInt(m[1], 10);

  // ATM WDL format:  ATM WDL-RRN-...
  m = d.match(/^ATM WDL-0*(\d+)/i);
  if (m) return parseInt(m[1], 10);

  // EDB settlement:  RRN-TEXT-...  (AFRIGO, MC ATM SETT, ZS ATM SETT etc.)
  m = d.match(/^0*(\d{6,})-/);
  if (m) return parseInt(m[1], 10);

  return '';
}


// ════════════════════════════════════════════════════════════════
//  FIND SYSTEM CLOSING BALANCE  (last row of BALANCE column in GL)
// ════════════════════════════════════════════════════════════════

function findSystemClosingBalance(data) {
  const found = findHeaderRow(data);
  if (!found) return 0;

  const { headerRowIdx, colMap } = found;
  const balIdx = colMap['BALANCE'];
  if (balIdx === undefined) return 0;

  let lastBal = 0;
  // Start from headerRow + 2 (skip blank + Opening Balance rows)
  for (let r = headerRowIdx + 2; r < data.length; r++) {
    const val = data[r][balIdx];
    if (val !== null && val !== '' && val !== undefined) {
      const num = toNum(val);
      if (!isNaN(num)) lastBal = num;
    }
  }
  return lastBal;
}


// ════════════════════════════════════════════════════════════════
//  REMOVE OLD FOOTER ROWS FROM TARGET SHEET
// ════════════════════════════════════════════════════════════════

function removeFooterRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Read column E (DESCRIPTION column, index 5) to find footer keywords
  const colE     = sheet.getRange(1, 5, lastRow, 1).getValues();
  const toDelete = [];

  for (let i = lastRow; i >= 1; i--) {
    const val = String(colE[i - 1][0]).toUpperCase();
    if (CFG.FOOTER_KEYWORDS.some(kw => val.includes(kw))) {
      toDelete.push(i);
    }
  }

  // Delete from bottom up — critical to avoid row number shifting
  toDelete.sort((a, b) => b - a);
  toDelete.forEach(row => sheet.deleteRow(row));
}


// ════════════════════════════════════════════════════════════════
//  APPEND NEW TRANSACTIONS (batch write — fast even for 1500+ rows)
// ════════════════════════════════════════════════════════════════

function appendTransactions(sheet, transactions) {
  if (transactions.length === 0) return sheet.getLastRow();

  // ── Duplicate guard: collect all RRNs already on this sheet ──
  const lastExisting = sheet.getLastRow();
  const existingRRNs = new Set();
  if (lastExisting >= 2) {
    sheet.getRange(2, 1, lastExisting - 1, 1)
      .getValues()
      .forEach(r => { if (r[0]) existingRRNs.add(String(r[0])); });
  }

  // Keep only transactions whose RRN is not already on the sheet
  const newOnly = transactions.filter(t => t.rrn && !existingRRNs.has(String(t.rrn)));

  const skipped = transactions.length - newOnly.length;
  if (skipped > 0) Logger.log(`   ⚠ Skipped ${skipped} duplicate RRN(s) already on sheet`);

  if (newOnly.length === 0) return sheet.getLastRow();
  transactions = newOnly;

  const startRow = sheet.getLastRow() + 1;

  // Build the data array (11 columns per row, balance col = 0 placeholder)
  const values = transactions.map(t => [
    t.rrn,          // A — RRN
    t.date1,        // B — Create Date
    t.date2,        // C — Effective Date
    t.trnCode,      // D — TRN Code
    t.description,  // E — Description
    t.amount,       // F — Amount
    t.dr,           // G — Debit
    t.cr,           // H — Credit
    0,              // I — Balance (overwritten with formula below)
    t.poster,       // J — Poster
    t.branch        // K — Branch
  ]);

  // Write all rows in a single API call (much faster than row-by-row)
  sheet.getRange(startRow, 1, values.length, 11).setValues(values);

  // Write balance formulas in one batch call
  const formulas = transactions.map((_, i) => [`=H${startRow + i}-G${startRow + i}`]);
  sheet.getRange(startRow, 9, formulas.length, 1).setFormulas(formulas);

  return startRow + transactions.length - 1; // returns the last data row number
}


// ════════════════════════════════════════════════════════════════
//  ADD FOOTER ROWS  (Proof Balance / System Balance / Difference)
// ════════════════════════════════════════════════════════════════

function addFooterRows(sheet, lastDataRow, sysBal) {
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const proofRow = lastDataRow + 1;
  const sysRow   = lastDataRow + 2;
  const diffRow  = lastDataRow + 3;

  // Proof Balance row
  sheet.getRange(proofRow, 5).setValue(`PROOF BALANCE AS AT ${today}`);
  sheet.getRange(proofRow, 9).setFormula(`=SUM(I2:I${lastDataRow})`);

  // System Balance row
  sheet.getRange(sysRow, 5).setValue(`SYSTEM BALANCE AS AT ${today}`);
  sheet.getRange(sysRow, 9).setValue(sysBal);

  // Difference row
  sheet.getRange(diffRow, 5).setValue('DIFFERENCE');
  sheet.getRange(diffRow, 9).setFormula(`=I${proofRow}-I${sysRow}`);

  // Bold all footer labels and values
  const footerCells = [
    [proofRow, 5], [proofRow, 9],
    [sysRow,   5], [sysRow,   9],
    [diffRow,  5], [diffRow,  9]
  ];
  footerCells.forEach(([r, c]) => sheet.getRange(r, c).setFontWeight('bold'));

  // Green (balanced) / Red (difference) conditional formatting
  const diffCell     = sheet.getRange(diffRow, 9);
  const existingRules = sheet.getConditionalFormatRules();

  existingRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#c6efce').setFontColor('#276221')
      .setRanges([diffCell]).build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(0)
      .setBackground('#ffc7ce').setFontColor('#9c0006')
      .setRanges([diffCell]).build()
  );
  sheet.setConditionalFormatRules(existingRules);
}


// ════════════════════════════════════════════════════════════════
//  GET PROOF BALANCE  (sums column I after formulas are calculated)
// ════════════════════════════════════════════════════════════════

function getProofBalance(sheet, lastDataRow) {
  // Read column I from row 2 to the last data row
  const vals = sheet.getRange(2, 9, lastDataRow - 1, 1).getValues();
  return vals.reduce((sum, row) => sum + (typeof row[0] === 'number' ? row[0] : 0), 0);
}


// ════════════════════════════════════════════════════════════════
//  UPDATE SUMMARY SHEET
// ════════════════════════════════════════════════════════════════

function updateSummarySheet(ss, results) {
  // Get or create the SUMMARY sheet — always place it first
  let summary = ss.getSheetByName('SUMMARY');
  if (!summary) {
    summary = ss.insertSheet('SUMMARY', 0);
  }
  summary.clearContents();
  summary.clearFormats();

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');

  // Title block
  summary.getRange(1, 1).setValue('ATM RECEIVABLES RECONCILIATION SUMMARY')
    .setFontWeight('bold').setFontSize(14).setFontColor('#1a4a7a');
  summary.getRange(2, 1).setValue('Last run: ' + now)
    .setFontColor('#666666').setFontSize(10);

  // Column headers
  const headers = [
    'GL Account', 'Sheet Name', 'New Rows Added',
    'Proof Balance', 'System Balance', 'Difference', 'Status'
  ];
  const hRange = summary.getRange(4, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight('bold')
        .setBackground('#1a4a7a')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');

  // Data rows
  results.forEach((r, i) => {
    const dataRow = [
      r.glAccount,
      r.sheetName,
      r.rowsAdded,
      r.proofBal,
      r.sysBal,
      r.diff,
      r.status
    ];
    const rowRange = summary.getRange(5 + i, 1, 1, dataRow.length);
    rowRange.setValues([dataRow]);

    // Alternate row background for readability
    if (i % 2 === 0) rowRange.setBackground('#f2f2f2');

    // Colour the Status cell
    const statusCell = summary.getRange(5 + i, 7);
    if (String(r.status).startsWith('✅')) {
      statusCell.setBackground('#c6efce').setFontColor('#276221').setFontWeight('bold');
    } else {
      statusCell.setBackground('#ffc7ce').setFontColor('#9c0006').setFontWeight('bold');
    }

    // Number formatting for balance and difference columns
    ['D', 'E', 'F'].forEach((col, ci) => {
      summary.getRange(5 + i, 4 + ci).setNumberFormat('#,##0.00');
    });
  });

  // Auto-size all columns
  headers.forEach((_, c) => summary.autoResizeColumn(c + 1));

  // Add a border around the table
  const tableRange = summary.getRange(4, 1, results.length + 1, headers.length);
  tableRange.setBorder(true, true, true, true, true, true, '#cccccc',
    SpreadsheetApp.BorderStyle.SOLID);
}


// ════════════════════════════════════════════════════════════════
//  VIEW LAST RUN LOG
// ════════════════════════════════════════════════════════════════

function viewLogs() {
  const log = PropertiesService.getScriptProperties().getProperty('LAST_LOG')
    || 'No logs available yet. Run a reconciliation first.';
  SpreadsheetApp.getUi().alert('Last Run Log', log, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ════════════════════════════════════════════════════════════════
//  UTILITIES
// ════════════════════════════════════════════════════════════════

/** Convert a value to a number — handles string amounts with commas */
function toNum(val) {
  if (typeof val === 'number') return val;
  if (val === null || val === undefined || val === '') return 0;
  const n = parseFloat(String(val).replace(/,/g, '').trim());
  return isNaN(n) ? 0 : n;
}

/** Format a date value to DD-MMM-YYYY string */
function fmtDate(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MMM-yyyy').toUpperCase();
  }
  // Clean up non-breaking spaces (\xa0) from banking system exports
  return String(val).trim().replace(/^\xa0/, '').trim();
}

/** Format a number with commas for display in alerts/logs */
function formatNumber(n) {
  if (typeof n !== 'number') return String(n);
  return n.toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
