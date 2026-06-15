// ═══════════════════════════════════════════════════════════════════════════
// Utils.gs — shared helpers used by ALL other .gs files
// This file MUST be present in the clasp project alongside the other .gs files
// ═══════════════════════════════════════════════════════════════════════════

var _ssCache = null;

function getSpreadsheet_() {
  if (_ssCache) return _ssCache;
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty('SS_ID');
  if (id) {
    _ssCache = SpreadsheetApp.openById(id);
    return _ssCache;
  }
  var active = SpreadsheetApp.getActive();
  if (active) {
    props.setProperty('SS_ID', active.getId());
    _ssCache = active;
    return _ssCache;
  }
  throw new Error('Run Initial Setup from the spreadsheet menu first.');
}

function setSsId(id) {
  PropertiesService.getScriptProperties().setProperty('SS_ID', id);
}

function checkSsId() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  var ss = getSpreadsheet_();
  return { id: id, name: ss.getName(), url: ss.getUrl() };
}

// Always returns [] — never throws, never returns null
function getSheetObjects_(sheetName) {
  try {
    var sh = getSpreadsheet_().getSheetByName(sheetName);
    if (!sh) return [];
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) return [];
    var vals = sh.getRange(1, 1, lr, lc).getValues();
    if (!vals || vals.length < 2) return [];
    var headers = vals[0].map(function(h) { return String(h || '').trim(); });
    return vals.slice(1)
      .filter(function(r) { return r.join('').trim() !== ''; })
      .map(function(r) {
        var o = {};
        headers.forEach(function(h, i) { o[h] = r[i] !== undefined ? r[i] : ''; });
        return o;
      });
  } catch(e) {
    Logger.log('getSheetObjects_ [' + sheetName + ']: ' + e.message);
    return [];
  }
}

// Always returns [] — never throws
function getSheetHeaders_(sheetName) {
  try {
    var sh = getSpreadsheet_().getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return [];
    var vals = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues();
    return (vals && vals[0]) ? vals[0].map(function(h) { return String(h || '').trim(); }) : [];
  } catch(e) {
    Logger.log('getSheetHeaders_ [' + sheetName + ']: ' + e.message);
    return [];
  }
}

function cloneObject_(obj) {
  try { return JSON.parse(JSON.stringify(obj || {})); } catch(e) { return {}; }
}

function trimObjectStrings_(obj) {
  var out = {};
  Object.keys(obj || {}).forEach(function(k) {
    var v = obj[k];
    out[k] = (typeof v === 'string') ? v.trim() : (v == null ? '' : v);
  });
  return out;
}

function parseDateSafe_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateKey_(val) {
  var d = parseDateSafe_(val);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

function audit_(action, recordId, details) {
  try {
    var sh = getSpreadsheet_().getSheetByName('AUDIT_LOG');
    if (!sh) return;
    sh.appendRow([
      new Date(),
      Session.getActiveUser().getEmail() || '',
      action,
      recordId || '',
      details  || ''
    ]);
  } catch(e) {}
}


// ═══════════════════════════════════════════════════════════════════════════
// importMasterDataFromImportSheets()
// Reads staging sheets IMPORT_BESCOM and IMPORT_GEO from the same spreadsheet
// and populates MASTER_BESCOM and MASTER_GEO respectively.
// Run from Spreadsheet menu: G.O Admin > Import Master Data
// Safe: aborts without touching master sheets if staging sheets or
//       required headers are missing.
// ═══════════════════════════════════════════════════════════════════════════
function importMasterDataFromImportSheets() {
  var ss = getSpreadsheet_();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { ui = null; }

  function readSheet_(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return null;
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) return { headers: [], rows: [] };
    var vals = sh.getRange(1, 1, lr, lc).getValues();
    return {
      headers: vals[0].map(function(h){ return String(h||'').trim().toLowerCase(); }),
      rows:    vals.slice(1)
    };
  }

  function normalizeActive_(v) {
    var s = String(v||'').trim().toLowerCase();
    return (s==='yes'||s==='y'||s==='1'||s==='true') ? 'yes' : 'no';
  }

  function trimRow_(arr) {
    return arr.map(function(v){ return typeof v==='string' ? v.trim() : (v==null?'':v); });
  }

  function validateHeaders_(headers, required, name) {
    var missing = required.filter(function(c){ return headers.indexOf(c)===-1; });
    if (missing.length) throw new Error(name+' is missing columns: '+missing.join(', '));
  }

  function alert_(title, msg) {
    if (ui) { try { ui.alert(title, msg, ui.ButtonSet.OK); return; } catch(e){} }
    Logger.log(title + ': ' + msg);
  }

  try {
    // ── BESCOM ──────────────────────────────────────────────────────────────
    var bescomReq = ['zone','circle','division','subdivision','active'];
    var bData = readSheet_('IMPORT_BESCOM');
    if (!bData) throw new Error('Sheet "IMPORT_BESCOM" not found. Create it first.');
    validateHeaders_(bData.headers, bescomReq, 'IMPORT_BESCOM');

    var bescomRows = [], bSkipped = 0;
    bData.rows.forEach(function(r) {
      var row = trimRow_(r);
      if (!row.join('').replace(/\s/g,'')) { bSkipped++; return; }
      var o = {};
      bData.headers.forEach(function(h,i){ o[h] = row[i]!==undefined?row[i]:''; });
      bescomRows.push([o['zone']||'', o['circle']||'', o['division']||'', o['subdivision']||'', normalizeActive_(o['active'])]);
    });

    // ── GEO ─────────────────────────────────────────────────────────────────
    var geoReq = ['district','taluk','active'];
    var gData = readSheet_('IMPORT_GEO');
    if (!gData) throw new Error('Sheet "IMPORT_GEO" not found. Create it first.');
    validateHeaders_(gData.headers, geoReq, 'IMPORT_GEO');

    var geoRows = [], gSkipped = 0;
    gData.rows.forEach(function(r) {
      var row = trimRow_(r);
      if (!row.join('').replace(/\s/g,'')) { gSkipped++; return; }
      var o = {};
      gData.headers.forEach(function(h,i){ o[h] = row[i]!==undefined?row[i]:''; });
      geoRows.push([o['district']||'', o['taluk']||'', normalizeActive_(o['active'])]);
    });

    // ── Write MASTER_BESCOM ──────────────────────────────────────────────────
    var bSh = ss.getSheetByName('MASTER_BESCOM');
    if (!bSh) throw new Error('"MASTER_BESCOM" not found. Run Initial Setup first.');
    var bLr = bSh.getLastRow();
    if (bLr > 1) bSh.getRange(2,1,bLr-1,bSh.getLastColumn()).clearContent();
    if (bescomRows.length) bSh.getRange(2,1,bescomRows.length,5).setValues(bescomRows);

    // ── Write MASTER_GEO ─────────────────────────────────────────────────────
    var gSh = ss.getSheetByName('MASTER_GEO');
    if (!gSh) throw new Error('"MASTER_GEO" not found. Run Initial Setup first.');
    var gLr = gSh.getLastRow();
    if (gLr > 1) gSh.getRange(2,1,gLr-1,gSh.getLastColumn()).clearContent();
    if (geoRows.length) gSh.getRange(2,1,geoRows.length,3).setValues(geoRows);

    var msg = '✓ Import complete!\n\n' +
              'MASTER_BESCOM : ' + bescomRows.length + ' rows written' +
              (bSkipped ? ' (' + bSkipped + ' blank rows skipped)' : '') + '\n' +
              'MASTER_GEO    : ' + geoRows.length    + ' rows written' +
              (gSkipped  ? ' (' + gSkipped  + ' blank rows skipped)' : '');
    alert_('Import Complete', msg);

  } catch(err) {
    alert_('Import Error', '✗ ' + err.message);
    Logger.log('importMasterDataFromImportSheets error: ' + err.message);
  }
}
