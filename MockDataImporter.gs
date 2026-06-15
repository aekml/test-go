// ═══════════════════════════════════════════════════════════════════════════
// MockDataImporter.gs — G.O Compliance (multi-entry build)
// MASTER_BESCOM  cols: zone, circle, division, subdivision, active
// MASTER_GEO     cols: district, taluk, active
// ═══════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ G.O Utilities')
    .addItem('1 · Run Initial Setup',        'runInitialSetup')
    .addSeparator()
    .addItem('2 · Load ALL Mock Data',        'loadAllMockData')
    .addItem(' Load Mock Masters Only',       'loadMockMastersOnly')
    .addItem(' Load Mock Transactions Only',  'loadMockTransactionsOnly')
    .addSeparator()
    .addItem('3 · Clear Transaction Data',    'clearTransactionMockData')
    .addSeparator()
    .addItem('Dev · Who Am I?',   'whoAmI_alert')
    .addItem('Dev · Check SS ID', 'checkSsId_alert')
    .addToUi();
}

function loadAllMockData() {
  runInitialSetup();
  _writeMasters();
  _writeTransactions();
  SpreadsheetApp.getUi().alert(
    '✓ Mock data loaded!\n\n' +
    'MASTER_BESCOM — 8 rows\nMASTER_GEO — 8 rows\n' +
    'MASTER_CORPORATION — 2 rows\nMASTER_ULB — 5 rows\n' +
    'DATA_SERVICE_PROGRESS — 6 Form1 rows\nDATA_SURVEY — 10 Form2 rows\n\n' +
    'Next: Deploy → New Deployment → Web App\n' +
    'Execute as: Me | Access: Anyone'
  );
}

function loadMockMastersOnly() {
  _ensureSheets_();
  _writeMasters();
  audit_('LOAD_MOCK_MASTERS', '', 'Master tables loaded');
  SpreadsheetApp.getUi().alert('✓ Master data loaded (transactions untouched).');
}

function loadMockTransactionsOnly() {
  _ensureSheets_();
  _writeTransactions();
  audit_('LOAD_MOCK_TXN', '', '6 Form1 + 10 Form2 rows loaded');
  SpreadsheetApp.getUi().alert('✓ Transaction data loaded (masters untouched).');
}

function clearTransactionMockData() {
  mock_clearRows_('DATA_SERVICE_PROGRESS');
  mock_clearRows_('DATA_SURVEY');
  audit_('CLEAR_MOCK_TXN', '', 'Transaction data cleared, masters untouched');
  SpreadsheetApp.getUi().alert('✓ Transaction sheets cleared.');
}

function whoAmI_alert() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  SpreadsheetApp.getUi().alert('Logged-in email: ' + (email || '(not available)'));
}

function checkSsId_alert() {
  var r = checkSsId();
  SpreadsheetApp.getUi().alert('SS_ID : ' + r.id + '\nName : ' + r.name + '\nURL : ' + r.url);
}

// ── Private helpers ──────────────────────────────────────────────────────────
function _writeMasters() {
  mock_overwrite_('MASTER_BESCOM',      MOCK_BESCOM_);
  mock_overwrite_('MASTER_GEO',         MOCK_GEO_);
  mock_overwrite_('MASTER_CORPORATION', MOCK_CORP_);
  mock_overwrite_('MASTER_ULB',         MOCK_ULB_);
}

function _writeTransactions() {
  mock_overwrite_('DATA_SERVICE_PROGRESS', MOCK_FORM1_);
  mock_overwrite_('DATA_SURVEY',           MOCK_FORM2_);
}

function _ensureSheets_() {
  var ss = getSpreadsheet_();
  var needed = {
    'MASTER_BESCOM':         ['zone','circle','division','subdivision','active'],
    'MASTER_GEO':            ['district','taluk','active'],
    'MASTER_CORPORATION':    ['corporation_name','active'],
    'MASTER_ULB':            ['ulb_body','body_type','active'],
    'DATA_SERVICE_PROGRESS': getForm1Headers_(),
    'DATA_SURVEY':           getForm2Headers_(),
    'AUDIT_LOG':             ['timestamp','user_email','action','record_id','details']
  };
  Object.keys(needed).forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      var sh = ss.insertSheet(name);
      var h  = needed[name];
      sh.getRange(1,1,1,h.length).setValues([h]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#dbeafe');
    }
  });
}

function mock_overwrite_(sheetName, data) {
  var sh = getSpreadsheet_().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName + '. Run Initial Setup first.');
  if (!data || data.length === 0) return;
  sh.clearContents();
  sh.getRange(1,1,data.length,data[0].length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,data[0].length).setFontWeight('bold').setBackground('#dbeafe');
  sh.autoResizeColumns(1, data[0].length);
}

function mock_clearRows_(sheetName) {
  var sh = getSpreadsheet_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() <= 1) return;
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
}

// ══════════════════════════════════════════════════════════════════════════════
// MOCK MASTER DATA
// ══════════════════════════════════════════════════════════════════════════════

// MASTER_BESCOM: zone | circle | division | subdivision | active
var MOCK_BESCOM_ = [
  ['zone','circle','division','subdivision','active'],
  ['Bangalore Zone','Bangalore Circle','Bangalore North Division','Hebbal Subdivision','TRUE'],
  ['Bangalore Zone','Bangalore Circle','Bangalore North Division','Yelahanka Subdivision','TRUE'],
  ['Bangalore Zone','Bangalore Circle','Bangalore North Division','RT Nagar Subdivision','TRUE'],
  ['Bangalore Zone','Bangalore Circle','Bangalore South Division','Jayanagar Subdivision','TRUE'],
  ['Bangalore Zone','Bangalore Circle','Bangalore South Division','JP Nagar Subdivision','TRUE'],
  ['Bangalore Zone','Bangalore Circle','Bangalore South Division','BTM Layout Subdivision','TRUE'],
  ['Mysore Zone','Mysore Circle','Mysore Division','Mysore North Subdivision','TRUE'],
  ['Mysore Zone','Mysore Circle','Mysore Division','Mysore South Subdivision','TRUE']
];

// MASTER_GEO: district | taluk | active
var MOCK_GEO_ = [
  ['district','taluk','active'],
  ['Bangalore Urban','Bangalore North','TRUE'],
  ['Bangalore Urban','Bangalore South','TRUE'],
  ['Bangalore Urban','Bangalore East','TRUE'],
  ['Bangalore Rural','Devanahalli','TRUE'],
  ['Bangalore Rural','Nelamangala','TRUE'],
  ['Mysore','Mysore','TRUE'],
  ['Mysore','Hunsur','TRUE'],
  ['Tumkur','Tumkur','TRUE']
];

var MOCK_CORP_ = [
  ['corporation_name','active'],
  ['BBMP','TRUE'],
  ['MCC','TRUE']
];

var MOCK_ULB_ = [
  ['ulb_body','body_type','active'],
  ['Anekal Town Municipal Council','ULB','TRUE'],
  ['Nelamangala Town Municipal Council','ULB','TRUE'],
  ['Mysore City Corporation','ULB','TRUE'],
  ['Hunsur Rural Local Body','RLB','TRUE'],
  ['Devanahalli Town Municipal Council','ULB','TRUE']
];

// ══════════════════════════════════════════════════════════════════════════════
// MOCK TRANSACTION DATA
// ══════════════════════════════════════════════════════════════════════════════

var F1H_ = [
  'record_id','batch_id','form_key','submitted_at','submitted_by_email',
  'officer_name','officer_phone',
  'zone_name','circle_name','division_name','subdivision_name',
  'district','taluk','region_type','corporation_name','ulb_body',
  'permanent_conn_applied',
  'app_no','app_date',
  'building_ready',
  'plot_sqft',
  'no_of_floors','no_of_connections','total_sanctioned_kw','date_of_service','any_other_info',
  'entry_flag',
  'workflow_status','workflow_remarks','reviewed_at','reviewed_by'
];

var MOCK_FORM1_ = [
  F1H_,
  // Batch 1 — ELIGIBLE entry
  ['SP-20260610-0001','BATCH-001','FORM1',new Date('2026-06-10T09:10:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001',
   'Bangalore Zone','Bangalore Circle','Bangalore North Division','Hebbal Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'Yes','APP-2024-1001',new Date('2024-03-01'),
   'Yes',1800,'G+2',2,4.5,new Date('2026-06-10'),'',
   'ELIGIBLE','Pending','','',''],
  // Batch 1 — BLDG_NOT_READY entry
  ['SP-20260610-0002','BATCH-001','FORM1',new Date('2026-06-10T09:11:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001',
   'Bangalore Zone','Bangalore Circle','Bangalore North Division','Hebbal Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'Yes','APP-2024-1002',new Date('2024-03-05'),
   'No','','','','','','',
   'BLDG_NOT_READY','Pending','','',''],
  // Batch 1 — OVERSIZED entry
  ['SP-20260610-0003','BATCH-001','FORM1',new Date('2026-06-10T09:12:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001',
   'Bangalore Zone','Bangalore Circle','Bangalore North Division','Hebbal Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'Yes','APP-2024-1003',new Date('2024-03-10'),
   'Yes',3200,'','','','','',
   'OVERSIZED','Rejected','Premises exceed GO limits',new Date('2026-06-10'),'div@bescom.org'],
  // Batch 2 — Gate 1 NO
  ['SP-20260610-0004','BATCH-002','FORM1',new Date('2026-06-10T10:00:00'),'officer2@bescom.org',
   'Suma Patil','9845001002',
   'Bangalore Zone','Bangalore Circle','Bangalore South Division','Jayanagar Subdivision',
   'Bangalore Urban','Bangalore South','GBA','BBMP','',
   'No','','','','','','','','','',
   'NO_PERM_CONN','Pending','','',''],
  // Batch 3 — ELIGIBLE ULB
  ['SP-20260611-0005','BATCH-003','FORM1',new Date('2026-06-11T09:30:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001',
   'Mysore Zone','Mysore Circle','Mysore Division','Mysore North Subdivision',
   'Mysore','Mysore','ULB','','Mysore City Corporation',
   'Yes','APP-2024-2001',new Date('2024-04-01'),
   'Yes',1600,'G+1',1,2.0,new Date('2026-06-11'),'ULB area',
   'ELIGIBLE','Verified','Field verified OK',new Date('2026-06-11'),'div@bescom.org'],
  // Batch 3 — ELIGIBLE RLB
  ['SP-20260611-0006','BATCH-003','FORM1',new Date('2026-06-11T09:31:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001',
   'Mysore Zone','Mysore Circle','Mysore Division','Mysore North Subdivision',
   'Mysore','Mysore','RLB','','Hunsur Rural Local Body',
   'Yes','APP-2024-2002',new Date('2024-04-05'),
   'Yes',1400,'G',1,1.5,new Date('2026-06-11'),'Ground floor',
   'ELIGIBLE','Pending','','','']
];

var F2H_ = getForm2Headers_();

var MOCK_FORM2_ = [
  F2H_,
  // 1) upto_2880 ELIGIBLE temp No Pending GBA
  ['SV-20260612-0001','FORM2',new Date('2026-06-12T09:00:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001','Bangalore Zone','Bangalore Circle','Bangalore North Division','Hebbal Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'upto_2880',1800,'G+2',3,9.0,
   'No','','','Yes','Pending connection',
   'ELIGIBLE','Pending','','',''],
  // 2) upto_2880 ELIGIBLE temp Yes Verified GBA
  ['SV-20260612-0002','FORM2',new Date('2026-06-12T09:15:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001','Bangalore Zone','Bangalore Circle','Bangalore North Division','Yelahanka Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'upto_2880',2200,'G+3',4,12.0,
   'Yes','TEMP-001',new Date('2026-04-15'),'Yes','Active temp supply',
   'ELIGIBLE','Verified','All docs verified',new Date('2026-06-10'),'div@bescom.org'],
  // 3) upto_2880 ELIGIBLE temp No Re-Submit GBA
  ['SV-20260612-0003','FORM2',new Date('2026-06-12T09:30:00'),'officer2@bescom.org',
   'Suma Patil','9845001002','Bangalore Zone','Bangalore Circle','Bangalore South Division','Jayanagar Subdivision',
   'Bangalore Urban','Bangalore South','GBA','BBMP','',
   'upto_2880',2450,'G+3',5,14.0,
   'No','','','No','Building under construction',
   'ELIGIBLE','Re-Submit','Photographs needed',new Date('2026-06-10'),'div@bescom.org'],
  // 4) upto_2880 ELIGIBLE temp Yes Pending RLB
  ['SV-20260612-0004','FORM2',new Date('2026-06-12T10:00:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001','Mysore Zone','Mysore Circle','Mysore Division','Mysore North Subdivision',
   'Mysore','Mysore','RLB','','Hunsur Rural Local Body',
   'upto_2880',1400,'Ground',1,4.0,
   'Yes','TEMP-002',new Date('2026-03-20'),'Yes','RLB area',
   'ELIGIBLE','Pending','','',''],
  // 5) upto_2880 ELIGIBLE temp No Verified ULB
  ['SV-20260612-0005','FORM2',new Date('2026-06-12T10:15:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001','Mysore Zone','Mysore Circle','Mysore Division','Mysore South Subdivision',
   'Mysore','Mysore','ULB','Anekal Town Municipal Council','',
   'upto_2880',1600,'G+1',2,6.0,
   'No','','','Yes','ULB area',
   'ELIGIBLE','Verified','Field OK',new Date('2026-06-11'),'div@bescom.org'],
  // 6) 2881_to_4800 NOT-ELIGIBLE temp No Rejected GBA
  ['SV-20260612-0006','FORM2',new Date('2026-06-12T11:00:00'),'officer2@bescom.org',
   'Suma Patil','9845001002','Bangalore Zone','Bangalore Circle','Bangalore South Division','JP Nagar Subdivision',
   'Bangalore Urban','Bangalore South','GBA','BBMP','',
   '2881_to_4800',3100,'G+4',6,18.0,
   'No','','','No','Awaiting service',
   'NOT-ELIGIBLE','Rejected','Exceeds GO limits',new Date('2026-06-11'),'div@bescom.org'],
  // 7) 2881_to_4800 NOT-ELIGIBLE temp Yes Pending GBA
  ['SV-20260612-0007','FORM2',new Date('2026-06-12T11:15:00'),'officer2@bescom.org',
   'Suma Patil','9845001002','Bangalore Zone','Bangalore Circle','Bangalore South Division','BTM Layout Subdivision',
   'Bangalore Urban','Bangalore South','GBA','BBMP','',
   '2881_to_4800',3500,'G+4',5,16.0,
   'Yes','TEMP-003',new Date('2026-04-10'),'No','Temp only',
   'NOT-ELIGIBLE','Pending','','',''],
  // 8) beyond_4800 NOT-ELIGIBLE temp No Pending GBA
  ['SV-20260612-0008','FORM2',new Date('2026-06-12T11:30:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001','Bangalore Zone','Bangalore Circle','Bangalore North Division','RT Nagar Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'beyond_4800',5200,'G+5',8,25.0,
   'No','','','No','Large layout',
   'NOT-ELIGIBLE','Pending','','',''],
  // 9) beyond_4800 NOT-ELIGIBLE temp Yes Pending GBA
  ['SV-20260612-0009','FORM2',new Date('2026-06-12T11:45:00'),'officer1@bescom.org',
   'Ravi Kumar','9845001001','Bangalore Zone','Bangalore Circle','Bangalore North Division','Hebbal Subdivision',
   'Bangalore Urban','Bangalore North','GBA','BBMP','',
   'beyond_4800',6000,'G+4',10,30.0,
   'Yes','TEMP-004',new Date('2026-05-01'),'Yes','Mega project',
   'NOT-ELIGIBLE','Pending','','','']
];

