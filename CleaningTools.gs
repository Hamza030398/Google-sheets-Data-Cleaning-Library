/** === AUTO CLEANING MODULE === */
function runAutoClean() {
  const sheet = SpreadsheetApp.getActiveSheet();
  createBackup(sheet, "AutoClean");
  const range = sheet.getDataRange();
  const values = range.getValues();

  const headers = values[0];
  const cleaned = [headers];
  const seenRows = new Set();

  for (let i = 1; i < values.length; i++) {
    let row = values[i].map(v => typeof v === 'string' ? v.trim() : v);
    row = row.map(cell => normalizeDateCell(cell));
    if (row.every(v => v === '')) continue;
    const rowKey = JSON.stringify(row);
    if (!seenRows.has(rowKey)) {
      seenRows.add(rowKey);
      cleaned.push(row);
    }
  }

  range.clearContent();
  sheet.getRange(1, 1, cleaned.length, cleaned[0].length).setValues(cleaned);
  SpreadsheetApp.getUi().alert('✅ Sheet cleaned and normalized successfully!');
}

/** Trims only spaces in the active selection */
function trimSelectionSides() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return;
  const values = range.getValues();
  const trimmed = values.map(r => r.map(v => typeof v === 'string' ? v.trim() : v));
  range.setValues(trimmed);
}

/** Removes duplicate rows (based on all columns) */
function removeDuplicates() {
  const sheet = SpreadsheetApp.getActiveSheet();
  createBackup(sheet, "RemoveDuplicates");
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values.shift();

  const unique = [];
  const seen = new Set();
  for (const row of values) {
    const key = JSON.stringify(row);
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(row);
    }
  }

  const cleaned = [headers, ...unique];
  range.clearContent();
  sheet.getRange(1, 1, cleaned.length, cleaned[0].length).setValues(cleaned);
  SpreadsheetApp.getUi().alert(`✅ Removed duplicates — kept ${unique.length} unique rows.`);
}

/** Normalizes date/time formats across the sheet */
function normalizeDateTime() {
  const sheet = SpreadsheetApp.getActiveSheet();
  createBackup(sheet, "NormalizeDateTime");
  const range = sheet.getDataRange();
  const values = range.getValues();

  for (let i = 1; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      values[i][j] = normalizeDateCell(values[i][j]);
    }
  }

  range.setValues(values);
  SpreadsheetApp.getUi().alert('✅ Dates & times normalized!');
}

/** --- Utility: Normalize a single cell’s date/time value --- */
function normalizeDateCell(val) {
  if (!val) return val;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  if (typeof val === 'number') {
    try {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    } catch (e) {}
  }
  if (typeof val === 'string') {
    const d = Date.parse(val);
    if (!isNaN(d)) return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return val;
}

/** --- Backup before running any data-destructive task --- */
function createBackup(sheet, tag) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = sheet.getName();
  const backupName = `Backup_${name}_${tag}_${new Date().toISOString().replace(/[:.]/g, '-')}`;
  const newSheet = sheet.copyTo(ss);
  newSheet.setName(backupName);
  newSheet.hideSheet();
  Logger.log(`🧷 Backup created: ${backupName}`);
}

/** --- Restore the latest backup for current sheet --- */
function restoreLatestBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const allSheets = ss.getSheets();

  const backups = allSheets
    .map(s => s.getName())
    .filter(n => n.startsWith(`Backup_${sheetName}_`));

  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No backups found for this sheet.');
    return;
  }

  // Sort backups by timestamp in name (latest last)
  backups.sort((a, b) => {
    const ta = a.split('_').pop();
    const tb = b.split('_').pop();
    return ta.localeCompare(tb);
  });

  const latestName = backups[backups.length - 1];
  const backupSheet = ss.getSheetByName(latestName);
  if (!backupSheet) {
    SpreadsheetApp.getUi().alert('⚠️ Backup sheet could not be found.');
    return;
  }

  const data = backupSheet.getDataRange().getValues();

  const confirm = SpreadsheetApp.getUi().alert(
    'Restore Backup',
    `This will overwrite current sheet "${sheetName}" with data from "${latestName}". Continue?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  // Overwrite current sheet content
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getUi().alert(`✅ Restored data from: ${latestName}`);
}
