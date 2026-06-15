function runInitialSetup() {
  var ss = getSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  var f1Headers = getForm1Headers_();
  var f2Headers = getForm2Headers_();

  createSheet_('MASTER_BESCOM',         ['zone','circle','division','subdivision','active']);
  createSheet_('MASTER_GEO',            ['district','taluk','active']);
  createSheet_('MASTER_CORPORATION',    ['corporation_name','active']);
  createSheet_('MASTER_ULB',            ['ulb_body','body_type','active']);
  createSheet_('DATA_SERVICE_PROGRESS', f1Headers);
  createSheet_('DATA_SURVEY',           f2Headers);
  createSheet_('AUDIT_LOG',             ['timestamp','user_email','action','record_id','details']);

  try {
    SpreadsheetApp.getUi().alert(
      '✓ Setup complete! SS_ID: ' + ss.getId() + '\n\n' +
      'Next: Run "Load ALL Mock Data" then Deploy as Web App.'
    );
  } catch(e) { Logger.log('Setup done. SS_ID: ' + ss.getId()); }
}

function createSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  if (!headers || !headers.length) return;
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#dbeafe');
  sh.autoResizeColumns(1, headers.length);
}

// ─── FORM 1 HEADERS ─────────────────────────────────────────────────────────
// Unchanged — do not modify
function getForm1Headers_() {
  return [
    'record_id', 'batch_id', 'form_key', 'submitted_at', 'submitted_by_email',
    'officer_name', 'officer_phone',
    'zone_name', 'circle_name', 'division_name', 'subdivision_name',
    'district', 'taluk', 'region_type', 'corporation_name', 'ulb_body',
    'permanent_conn_applied',
    'app_no', 'app_date',
    'building_ready',
    'plot_sqft',
    'no_of_floors', 'no_of_connections', 'total_sanctioned_kw',
    'date_of_service', 'any_other_info',
    'entry_flag',
    'workflow_status', 'workflow_remarks', 'reviewed_at', 'reviewed_by'
  ];
}

// ─── FORM 2 HEADERS ─────────────────────────────────────────────────────────
// Survey form — one row per premises surveyed.
// plot_category : 'upto_2880' | '2881_to_4800' | 'beyond_4800'
// temp_conn_applied : Yes | No
// bldg_ready_occupied : Yes | No
function getForm2Headers_() {
  return [
    // ── Meta ────────────────────────────────────────────────────────────
    'record_id', 'form_key', 'submitted_at', 'submitted_by_email',
    // ── Section 1: Officer ──────────────────────────────────────────────
    'officer_name', 'officer_phone',
    // ── Section 2: BESCOM Hierarchy ─────────────────────────────────────
    'zone_name', 'circle_name', 'division_name', 'subdivision_name',
    // ── Section 3: Location ─────────────────────────────────────────────
    'district', 'taluk', 'region_type', 'corporation_name', 'ulb_body',
    // ── Section 4: Plot & Premises ──────────────────────────────────────
    'plot_category', 'plot_sqft', 'no_of_floors',
    'expected_connections', 'expected_load_kw',
    // ── Section 5: Connection & Occupancy ───────────────────────────────
    'temp_conn_applied', 'temp_account_id', 'app_reg_date',
    'bldg_ready_occupied',
    'remarks',
    // eligibility flag for survey
    'entry_flag',
    // ── Workflow ────────────────────────────────────────────────────────
    'workflow_status', 'workflow_remarks', 'reviewed_at', 'reviewed_by'
  ];
}
