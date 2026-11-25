// UIManager.js - the spreadsheet menu.

/**
 * An onOpen simple trigger that runs automatically when the spreadsheet is opened.
 * It creates a custom menu for easy access to the tool's functions.
 * @param {object} e The event object.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('QA Ticket Generator')
    .addItem('▶️ Run Test Generation', 'QAGEN_runTestGeneration')
    .addItem('▶️ Run Campaign...', 'QAGEN_runCampaign')
    .addItem('🔬 Run Diagnostics', 'QAGEN_runDiagnostics')
    .addSeparator()
    .addItem('🔄 Sync YouTrack Statuses', 'QAGEN_syncYouTrackStatuses')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Configuration')
      .addItem('🔑 Set YouTrack Token', 'QAGEN_setYouTrackToken')
      .addSeparator()
      .addItem('➡ Apply Data Validation Rules', 'QAGEN_applyDataValidationRules')
      .addSeparator()
      .addItem('🔄 Force Schema Update', 'QAGEN_updateYouTrackSchemaCache')
      .addItem('🗑️ Clear Schema Hash', 'QAGEN_clearSchemaHash'))
    .addToUi();
}

/**
 * Runs on installation to create the menu immediately.
 * @param {object} e The event object.
 */
function onInstall(e) {
  onOpen(e);
}