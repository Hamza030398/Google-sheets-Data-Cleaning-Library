/**
 * @fileoverview Contains functions for the Regex Sidebar, including
 * the core execution logic and pattern storage using PropertiesService.
 */

// ----------------------------------------------------------------------
// UTILITY FUNCTIONS (Implemented from placeholders)
// ----------------------------------------------------------------------

/**
 * Creates a simple backup sheet by duplicating the current sheet.
 * NOTE: This is a robust safety feature before destructive operations.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {string} action Description of the action causing the backup.
 */
function createBackup(sheet, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Create a uniquely named backup sheet
  const backupSheetName = sheet.getName() + ' Backup (' + action + ') ' + new Date().toLocaleTimeString();
  sheet.copyTo(ss).setName(backupSheetName);
  Logger.log("Created backup sheet: " + backupSheetName);
}

/**
 * Logs the successful regex action configuration to User Properties.
 * This log is retrieved by the viewLogs() function in SidebarUI.gs.
 * @param {object} config The configuration object passed to runRegexAction.
 */
function logAction(config) {
  const properties = PropertiesService.getUserProperties();
  const logKey = 'REGEX_ACTION_LOGS'; // Key must match SidebarUI.gs > getLogs()
  
  // Load existing logs or initialize an empty array
  const logs = JSON.parse(properties.getProperty(logKey) || '[]');
  
  const logEntry = {
    timestamp: new Date().toLocaleString(),
    pattern: config.pattern,
    action: config.action,
    replacement: config.replacement,
    flags: config.flags,
    outputMode: config.outputMode
  };
  
  // Add new entry to the start of the array (most recent first)
  logs.unshift(logEntry);
  
  // Limit the log size to prevent property storage limits (e.g., last 50 entries)
  while (logs.length > 50) {
    logs.pop(); 
  }
  
  // Save the updated log array
  properties.setProperty(logKey, JSON.stringify(logs));
  Logger.log("Logged regex action: " + JSON.stringify(logEntry));
}

// ----------------------------------------------------------------------
// CORE REGEX EXECUTION
// ----------------------------------------------------------------------

function runRegexAction(config) {
  const { pattern, action, replacement, flags, preview, outputMode } = config;
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();

  // === Preserve formatting before any changes ===
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontWeights = range.getFontWeights();

  // --- Always clear old highlights ---
  range.setBackground(null);

  // --- Create safety backup before applying any changes ---
  if (!preview && ['replace', 'newColumn', 'newSheet'].includes(outputMode)) {
    createBackup(sheet, "RegexRun");
  }

  // Use try-catch to safely handle invalid regex patterns
  let regex;
  try {
    regex = new RegExp(pattern, flags || 'g');
  } catch (e) {
    // If the regex is invalid, return an error message to the client
    return { error: true, message: "Invalid Regular Expression: " + e.message };
  }
  
  const output = []; // 2D array for final output
  const previewResults = []; // 2D array for preview

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const newRow = []; // Array for current output row
    const previewRow = []; // Array for current preview row

    for (let c = 0; c < row.length; c++) {
      let cell = row[c];
      
      // Convert non-strings to string for regex matching
      const cellString = String(cell);
      
      let newCellValue = cell; // Default value if no change
      let previewCellValue = cell;

      if (typeof cellString === 'string') {
        
        // --- Highlighting (Needs separate regex instance for highlighting) ---
        // Use a temporary regex instance for test/highlighting
        let highlightRegex = new RegExp(pattern, flags || 'g');
        const matched = highlightRegex.test(cellString);
        
        if (matched) {
          range.getCell(r + 1, c + 1).setBackground('#FFF59D');
        }

        // --- Action Logic ---
        if (action === 'extract') {
          // Extraction: Find all matches or just the first match (depends on 'g' flag)
          // We must use a fresh matchAll call since test() consumes the match state
          const matchAllRegex = new RegExp(pattern, flags || 'g');
          const matches = [...cellString.matchAll(matchAllRegex)].map(m => m[0]);
          if (matches.length > 0) {
            newCellValue = matches.join(', '); // Join multiple matches
            previewCellValue = newCellValue;
          } else {
            newCellValue = '';
            previewCellValue = '';
          }
        } else if (action === 'replace') {
          // Replacement uses the original regex instance
          const replaced = cellString.replace(regex, replacement);
          previewCellValue = replaced;
          newCellValue = replaced;
        } else if (action === 'test') {
          // Re-test to get the true/false result after the highlighting test consumed it
          const testRegex = new RegExp(pattern, flags);
          const testResult = testRegex.test(cellString);
          previewCellValue = testResult; // true/false
          newCellValue = testResult;
        } else {
          // Default to original cell value
        }
      }
      
      newRow.push(newCellValue);
      previewRow.push(previewCellValue);
    }
    
    // Add the processed row to the 2D output arrays
    output.push(newRow);
    previewResults.push(previewRow);
  }

  if (preview) return previewResults;

  // --- Apply Changes based on Output Mode ---
  if (outputMode === 'replace') {
    range.setValues(output);

    // === Restore original formatting ===
    range.setBackgrounds(backgrounds);
    range.setFontColors(fontColors);
    range.setFontWeights(fontWeights);

  } else if (outputMode === 'newColumn') {
    const startCol = range.getLastColumn() + 1;
    // output is a 2D array, so we use its dimensions
    sheet.getRange(range.getRow(), startCol, output.length, output[0].length).setValues(output);
  } else if (outputMode === 'newSheet') {
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('RegexOutput_' + new Date().getTime());
    newSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  }

  // Log successful action only if it was executed (not preview)
  logAction(config);
  return 'done';
}


// ----------------------------------------------------------------------
// PATTERN STORAGE FUNCTIONS
// ----------------------------------------------------------------------

const PROPERTIES_KEY = 'SAVED_REGEX_PATTERNS';

/**
 * Loads all user-saved regex patterns from the PropertiesService.
 * @returns {object} An object mapping pattern names to their configuration.
 */
function loadPatterns() {
  const properties = PropertiesService.getUserProperties();
  const rawData = properties.getProperty(PROPERTIES_KEY);
  // Returns an empty object if no patterns are found
  return rawData ? JSON.parse(rawData) : {};
}

/**
 * Saves a new regex pattern configuration.
 * @param {string} name - The name to save the pattern under.
 * @param {string} pattern - The regex pattern string.
 * @param {string} action - The action (e.g., 'extract', 'replace').
 * @param {string} replacement - The replacement string.
 * @param {string} flags - The regex flags (e.g., 'gi').
 * @returns {void}
 */
function savePattern(name, pattern, action, replacement, flags) {
  const properties = PropertiesService.getUserProperties();
  const patterns = loadPatterns();
  
  patterns[name] = {
    pattern: pattern,
    action: action,
    replacement: replacement,
    flags: flags
  };
  
  properties.setProperty(PROPERTIES_KEY, JSON.stringify(patterns));
}

// ----------------------------------------------------------------------
// SIDEBAR & LOGS ENTRYPOINTS (Needed for bound script + menu to work)
// ----------------------------------------------------------------------

/**
 * Opens the Regex Sidebar UI inside the sheet.
 */
function openRegexSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Advanced Regex Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens a simple alert letting the user know where to find logs.
 * (Optional: could later show a proper log viewer sidebar)
 */
function viewLogs() {
  const ui = SpreadsheetApp.getUi();
  const logKey = 'REGEX_ACTION_LOGS';
  const props = PropertiesService.getUserProperties();
  const logs = JSON.parse(props.getProperty(logKey) || '[]');
  if (logs.length === 0) {
    ui.alert('No regex logs found yet.');
    return;
  }
  ui.alert(`📘 You have ${logs.length} logged Regex actions. Viewable in the Script Properties.`);
}

