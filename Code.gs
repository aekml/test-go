// ═══════════════════════════════════════════════════════════════════════════
// Code.gs — entry point (doGet, include, getBootstrapData)
// ═══════════════════════════════════════════════════════════════════════════

var APP_NAME          = 'G.O Compliance - Service Progress for Residential Premises';
var WORKFLOW_STATUSES = ['Pending', 'Verified', 'Re-Submit', 'Rejected'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Single boot call — returns only what the client needs at startup
function getBootstrapData() {
  var email = '';
  try { email = Session.getActiveUser().getEmail()    || ''; } catch(e) {}
  try { if (!email) email = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}

  return {
    appName:          APP_NAME,
    userEmail:        email,
    masters:          getMastersPayload(),
    workflowStatuses: WORKFLOW_STATUSES
  };
}


// ── Spreadsheet menu ─────────────────────────────────────────────────────────
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('G.O Admin')
      .addItem('Run Initial Setup',  'runInitialSetup')
      .addItem('Import Master Data', 'importMasterDataFromImportSheets')
      .addToUi();
  } catch(e) { Logger.log('onOpen: ' + e.message); }
}
