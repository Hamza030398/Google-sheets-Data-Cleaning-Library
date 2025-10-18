/**
 * @fileoverview This script is bound to the Google Sheet and provides
 * global proxy functions. These proxies allow the custom menu items
 * and the Regex Sidebar (which uses direct google.script.run calls)
 * to correctly call the functions defined within the Data_Cleaning_tools_Lib library.
 *
 * NOTE: It is assumed the library is added with the identifier 'DataCleaningTools'.
 */

// -----------------------------------------------------------------------------
// PROXY FUNCTIONS FOR CUSTOM MENU (Must match the menu item handlers)
// -----------------------------------------------------------------------------

function runAutoClean() {
  DataCleaningTools.runAutoClean();
}

function trimSelectionSides() {
  DataCleaningTools.trimSelectionSides();
}

function removeDuplicates() {
  DataCleaningTools.removeDuplicates();
}

function normalizeDateTime() {
  DataCleaningTools.normalizeDateTime();
}

function restoreLatestBackup() {
  DataCleaningTools.restoreLatestBackup();
}

/**
 * PROXY: Called by the custom menu to open the Regex sidebar.
 */
function openRegexSidebar() {
  // Now calls the function defined in SidebarUI.gs
  DataCleaningTools.openRegexSidebar();
}

/**
 * PROXY: Called by the custom menu to view the logs.
 */
function viewLogs() {
  // Now calls the function defined in SidebarUI.gs
  DataCleaningTools.viewLogs();
}

// -----------------------------------------------------------------------------
// PROXY FUNCTIONS FOR REGEX SIDEBAR (Must match the google.script.run calls in Sidebar.html)
// -----------------------------------------------------------------------------

/**
 * Proxy for the core regex execution function called by the sidebar.
 * @param {object} config - The configuration object passed from the HTML.
 * @returns {Array|string|object} The result (preview data, 'done', or error object).
 */
function runRegexAction(config) {
  return DataCleaningTools.runRegexAction(config);
}

/**
 * Proxy for loading saved patterns.
 * @returns {object} Map of saved patterns.
 */
function loadPatterns() {
  return DataCleaningTools.loadPatterns();
}

/**
 * Proxy for saving a new pattern.
 * @param {string} name 
 * @param {string} pattern 
 * @param {string} action 
 * @param {string} replacement 
 * @param {string} flags 
 */
function savePattern(name, pattern, action, replacement, flags) {
  DataCleaningTools.savePattern(name, pattern, action, replacement, flags);
}

/**
 * PROXY ADDED: Called by Logs.html to retrieve the log history.
 */
function getLogs() {
  return DataCleaningTools.getLogs();
}

// -----------------------------------------------------------------------------
// onOpen Wrapper (Crucial for menu display and sidebar function accessibility)
// -----------------------------------------------------------------------------
/**
 * Executes automatically when the spreadsheet opens, ensuring the custom menu
 * is created by the library.
 */
function onOpen(e) {
  try {
    DataCleaningTools.onOpen(e);
  } catch (error) {
    // Basic error handling if the library is missing or misconfigured
    console.error('Failed to run library onOpen:', error);
  }
}
