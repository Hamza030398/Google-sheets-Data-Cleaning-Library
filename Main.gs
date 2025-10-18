/**
 * @OnlyCurrentDoc
 * Data_CleaningTools — Modular Library with Backup Safety
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧹 Data Cleaning Tools')
    .addItem('Clean Sheet (Auto Normalize)', 'runAutoClean')
    .addItem('Trim Selection (Spaces Only)', 'trimSelectionSides')
    .addSeparator()
    .addItem('Remove Duplicates', 'removeDuplicates')
    .addItem('Format Dates & Times', 'normalizeDateTime')
    .addSeparator()
    .addItem('Restore Latest Backup', 'restoreLatestBackup')  // 🆕 new feature
    .addSeparator()
    .addItem('Open Regex Tool', 'openRegexSidebar')
    .addItem('View Regex Logs', 'viewLogs')
    .addToUi();
}
