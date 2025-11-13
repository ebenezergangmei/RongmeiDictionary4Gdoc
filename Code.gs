// Code.gs — Rongmei Dictionary for Google Sheets (robust poll pattern)

// Add menu when spreadsheet opens
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Rongmei Dictionary')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

// Show the sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Rongmei Dictionary');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Return the currently selected text (active range display value)
// This is called by the sidebar (client) via google.script.run
function getSelectedText() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getActiveRange();
    if (!range) return '';
    // If multi-cell selected, join values with spaces
    const values = range.getDisplayValues();
    if (values.length === 1 && values[0].length === 1) {
      return String(values[0][0]).trim();
    } else {
      // flatten 2D array to single string
      return values.map(row => row.join(' ')).join(' ').trim();
    }
  } catch (e) {
    // Log for debugging
    Logger.log('getSelectedText error: ' + e);
    return '';
  }
}
