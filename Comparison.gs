/** === COMPARISON TOOL ===
 * Creates a side-by-side comparison sheet for current sheet vs a chosen backup sheet.
 * Highlights:
 *  🔴 Red → value mismatch
 *  🟡 Yellow → extra rows or columns found only in one sheet
 */
function compareWithBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const parentSheet = ss.getActiveSheet();
  const parentName = parentSheet.getName();

  // --- Step 1: Ask user which backup to compare ---
  const backups = ss.getSheets()
    .map(s => s.getName())
    .filter(name => {
      const n = name.toLowerCase();
      const parent = parentName.toLowerCase();

      // Old style: "Backup_<Parent>_yyyy..."
      if (n.startsWith(`backup_${parent}_`)) return true;

      // Newer style examples: "Data Backup (AutoClean) 10:43:04"
      if (n.includes('backup') && n.includes(parent)) return true;

      // Generic data backups
      if (n.includes('data backup')) return true;

      return false;
    });

  const choice = ui.prompt(
    "Compare with Backup",
    "Available backups:\n" + backups.join("\n") +
      "\n\nEnter exact name of the backup sheet to compare:",
    ui.ButtonSet.OK_CANCEL
  );

  if (choice.getSelectedButton() !== ui.Button.OK) return;
  const backupName = choice.getResponseText().trim();
  const backupSheet = ss.getSheetByName(backupName);
  if (!backupSheet) {
    ui.alert("⚠️ Backup sheet not found: " + backupName);
    return;
  }

  // --- Step 2: Prepare data ---
  const parentData = parentSheet.getDataRange().getValues();
  const backupData = backupSheet.getDataRange().getValues();

  const parentHeaders = parentData[0].map(h => h.toString().trim());
  const backupHeaders = backupData[0].map(h => h.toString().trim());

  // --- Step 3: Determine column structure ---
  const allHeaders = Array.from(new Set([...backupHeaders, ...parentHeaders]));
  const extraParentCols = parentHeaders.filter(h => !backupHeaders.includes(h));
  const extraBackupCols = backupHeaders.filter(h => !parentHeaders.includes(h));

  const orderedHeaders = [
    ...allHeaders.filter(h => backupHeaders.includes(h) && parentHeaders.includes(h)),
    ...extraParentCols,
    ...extraBackupCols
  ];

  // --- Step 4: Create comparison sheet ---
  const compName = `Comparison_${parentName}_vs_${backupName}`;
  const oldComp = ss.getSheetByName(compName);
  if (oldComp) ss.deleteSheet(oldComp);

  const compSheet = ss.insertSheet(compName);
  compSheet.activate();

  // --- Step 5: Write headers (Backup–Current pairs) ---
  const pairedHeaders = [];
  orderedHeaders.forEach(h => {
    pairedHeaders.push(`${h} (Backup)`);
    pairedHeaders.push(`${h} (Current)`);
  });
  compSheet.getRange(1, 1, 1, pairedHeaders.length)
    .setValues([pairedHeaders])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");

  // --- Step 6: Compare row by row ---
  const maxRows = Math.max(parentData.length, backupData.length) - 1; // exclude header
  const comparison = [];
  const backupBody = backupData.slice(1);
  const parentBody = parentData.slice(1);

  for (let i = 0; i < maxRows; i++) {
    const backupRow = backupBody[i] || [];
    const parentRow = parentBody[i] || [];
    const backupMap = {};
    const parentMap = {};

    orderedHeaders.forEach((h) => {
      const bi = backupHeaders.indexOf(h);
      const pi = parentHeaders.indexOf(h);
      backupMap[h] = bi >= 0 ? backupRow[bi] : "";
      parentMap[h] = pi >= 0 ? parentRow[pi] : "";
    });

    const rowOut = [];
    orderedHeaders.forEach(h => {
      rowOut.push(backupMap[h]);
      rowOut.push(parentMap[h]);
    });
    comparison.push(rowOut);
  }

  // --- Step 7: Paste data ---
  if (comparison.length > 0)
    compSheet.getRange(2, 1, comparison.length, pairedHeaders.length).setValues(comparison);

  // --- Step 8: Highlight discrepancies & extras ---
  const lastRow = compSheet.getLastRow();
  const lastCol = compSheet.getLastColumn();
  const range = compSheet.getRange(2, 1, lastRow - 1, lastCol);
  const bgColors = range.getBackgrounds();
  const values = range.getValues();

  // 🔹 Helper: fuzzy equality for date/time precision issues
  function isNearlyEqual(a, b) {
    if (a === b) return true;
    if (a == null || b == null || a === "" || b === "") return a === b;

    // numeric (for serial dates/times)
    if (typeof a === "number" && typeof b === "number") {
      return Math.abs(a - b) < 0.0001;
    }

    // both are valid Date objects
    if (a instanceof Date && b instanceof Date) {
      return Math.abs(a.getTime() - b.getTime()) < 1000; // within 1s tolerance
    }

    return false;
  }

  // Compare every Backup–Current pair
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < pairedHeaders.length; c += 2) {
      const backupVal = values[r][c];
      const parentVal = values[r][c + 1];

      // Apply tolerance for date/time precision only
      if (!isNearlyEqual(backupVal, parentVal)) {
        bgColors[r][c] = "#ffb3b3";     // light red
        bgColors[r][c + 1] = "#ffb3b3"; // light red
      }
    }
  }

  // Highlight extra columns
  const yellow = "#fff3b3";
  orderedHeaders.forEach((h, i) => {
    const backupCol = 2 * i + 1;
    const currentCol = 2 * i + 2;
    if (extraParentCols.includes(h)) {
      compSheet.getRange(1, currentCol, lastRow).setBackground(yellow);
    }
    if (extraBackupCols.includes(h)) {
      compSheet.getRange(1, backupCol, lastRow).setBackground(yellow);
    }
  });

  // Highlight extra rows
  if (parentBody.length > backupBody.length) {
    compSheet.getRange(backupBody.length + 2, 1, parentBody.length - backupBody.length, lastCol)
      .setBackground(yellow);
  }

  range.setBackgrounds(bgColors);
  ui.alert(`✅ Comparison sheet created:\n"${compName}"\n🔴 Red = Mismatch | 🟡 Yellow = New/Extra`);
}
