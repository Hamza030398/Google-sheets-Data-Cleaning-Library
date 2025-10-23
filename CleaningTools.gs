/** === AUTO CLEANING MODULE (SAFE FORMATTING VERSION) === */
function runAutoClean() {
  const sheet = SpreadsheetApp.getActiveSheet();
  createBackup(sheet, "AutoClean");

  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0].map(h => h.toString().trim());
  const cleaned = [headers];
  const seenRows = new Set();

  for (let i = 1; i < values.length; i++) {
    let row = values[i].map(v => (typeof v === 'string' ? v.trim() : v));

    // --- Header-aware formatting ---
    for (let j = 0; j < row.length; j++) {
      const header = headers[j].toLowerCase();
      const cell = row[j];
      if (isDateHeader(header)) {
        row[j] = formatDateCell(cell); // format only
      } else if (isTimeHeader(header)) {
        row[j] = formatTimeCell(cell); // format only
      }
    }

    // Skip empty rows
    if (row.every(v => v === "")) continue;

    const rowKey = JSON.stringify(row);
    if (!seenRows.has(rowKey)) {
      seenRows.add(rowKey);
      cleaned.push(row);
    }
  }

  range.clearContent();
  sheet.getRange(1, 1, cleaned.length, cleaned[0].length).setValues(cleaned);
  SpreadsheetApp.getUi().alert('✅ Sheet cleaned and formatted safely!');
}

/** Trims only spaces in the active selection */
function trimSelectionSides() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return;
  const values = range.getValues();
  const trimmed = values.map(r => r.map(v => (typeof v === 'string' ? v.trim() : v)));
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

/** Normalizes date/time formats across the sheet (explicit user action only) */
function normalizeDateTime() {
  const sheet = SpreadsheetApp.getActiveSheet();
  createBackup(sheet, "NormalizeDateTime");

  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0].map(h => h.toString().trim().toLowerCase());

  for (let i = 1; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const header = headers[j];
      const val = values[i][j];
      if (val === "" || val === null || val === undefined) continue;

      if (isDateHeader(header)) {
        values[i][j] = formatDateCell(val);
      } else if (isTimeHeader(header)) {
        values[i][j] = formatTimeCell(val);
      }
    }
  }

  range.setValues(values);
  SpreadsheetApp.getUi().alert('✅ Dates & times formatted (no values changed).');
}

/** --- Detects whether a header likely represents a date column --- */
function isDateHeader(header) {
  const dateKeywords = [
    "date", "day", "dob", "birth", "created", "joined",
    "processed", "updated", "submitted", "timestamp"
  ];
  const skipKeywords = [
    "amount", "price", "balance", "total", "cost",
    "fare", "qty", "quantity", "account", "id", "number"
  ];
  return dateKeywords.some(k => header.includes(k)) && !skipKeywords.some(k => header.includes(k));
}

/** --- Detects whether a header likely represents a time column --- */
function isTimeHeader(header) {
  const timeKeywords = ["time", "hour", "duration", "clock", "timing"];
  return timeKeywords.some(k => header.includes(k));
}

/** --- Format a date value safely (no reinterpretation) --- */
function formatDateCell(val) {
  try {
    if (!val) return val;

    // If it's a true Date object
    if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val)) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }

    // If it's numeric serial (Excel/Sheets style)
    if (typeof val === 'number' && val > 25000 && val < 60000) {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }

    // If it's a string, only normalize *formatting*, not meaning
    if (typeof val === 'string') {
      const trimmed = val.trim();

      // Match recognizable short date forms like 1/1/24 → pad to 01/01/2024
      const shortPattern = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/;
      const m = trimmed.match(shortPattern);
      if (m) {
        let [_, d, mth, y] = m;
        if (y.length === 2) y = '20' + y; // assume 20xx if 2-digit year
        return `${d.padStart(2, '0')}/${mth.padStart(2, '0')}/${y}`;
      }
      return trimmed;
    }

    return val;
  } catch (e) {
    Logger.log("formatDateCell() error for " + val + ": " + e.message);
    return val;
  }
}

/** --- Format a time value safely (no reinterpretation) --- */
function formatTimeCell(val) {
  try {
    if (!val) return val;

    // If numeric fraction of a day (Sheets style)
    if (typeof val === 'number' && val >= 0 && val < 1) {
      const totalSeconds = Math.round(val * 24 * 60 * 60);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      return [hours, minutes, seconds].map(v => v.toString().padStart(2, '0')).join(':');
    }

    // If string, format to HH:MM:SS
    if (typeof val === 'string') {
      const str = val.trim();
      const timePattern = /^(\d{1,2})([:\-])?(\d{0,2})?([:\-]?(\d{0,2}))?$/;
      const m = str.match(timePattern);
      if (m) {
        const h = m[1].padStart(2, '0');
        const mi = (m[3] || '00').padStart(2, '0');
        const s = (m[5] || '00').padStart(2, '0');
        return `${h}:${mi}:${s}`;
      }
    }

    return val;
  } catch (e) {
    Logger.log("formatTimeCell() error for " + val + ": " + e.message);
    return val;
  }
}

/** --- Backup utility --- */
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

  backups.sort((a, b) => a.localeCompare(b));
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

  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getUi().alert(`✅ Restored data from: ${latestName}`);
}
