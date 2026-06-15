var FORM_CONFIGS_ = {
  FORM1: { key: 'FORM1', id_prefix: 'SP-', data_sheet: 'DATA_SERVICE_PROGRESS' },
  FORM2: { key: 'FORM2', id_prefix: 'SV-', data_sheet: 'DATA_SURVEY' }
};

// ── submitForm ───────────────────────────────────────────────────────────────
// Submits a single entry row. Called per application entry from the client.
// payload must include form_key and all relevant fields for this entry.
function submitForm(payload) {
  payload = trimObjectStrings_(payload || {});
  var formKey = String(payload.form_key || '').toUpperCase();
  var cfg = FORM_CONFIGS_[formKey];
  if (!cfg) throw new Error('Unknown form key: ' + formKey);

  if (!payload.officer_name)  throw new Error('Officer Name is required.');
  if (!payload.officer_phone) throw new Error('Officer Phone is required.');

  var headers = formKey === 'FORM1' ? getForm1Headers_() : getForm2Headers_();
  var row = buildRow_(payload, formKey);

  validateRequired_(row, formKey);
  validateBusinessRules_(row, formKey);

  var sh = getSpreadsheet_().getSheetByName(cfg.data_sheet);
  if (!sh) throw new Error('Sheet ' + cfg.data_sheet + ' not found. Run Initial Setup.');
  sh.appendRow(headers.map(function(h) { return row[h] !== undefined ? row[h] : ''; }));
  audit_('SUBMIT', row.record_id, 'form:' + formKey + ' flag:' + (row.entry_flag||'') + ' by:' + row.officer_name);
  return { success: true, record_id: row.record_id, entry_flag: row.entry_flag || '' };
}

// ── submitBatch ──────────────────────────────────────────────────────────────
// Submits multiple Form1 entry rows in one call (all entries from one session).
// Each row in the array already has the shared header data merged in.
function submitBatch(formKey, rows) {
  formKey = String(formKey || '').toUpperCase();
  var cfg = FORM_CONFIGS_[formKey];
  if (!cfg) throw new Error('Unknown form key: ' + formKey);
  rows = Array.isArray(rows) ? rows : [];
  var headers = formKey === 'FORM1' ? getForm1Headers_() : getForm2Headers_();
  var errors = [], values = [];

  rows.forEach(function(raw, idx) {
    try {
      var p = trimObjectStrings_(Object.assign({}, raw, { form_key: formKey }));
      var row = buildRow_(p, formKey);
      validateRequired_(row, formKey);
      validateBusinessRules_(row, formKey);
      values.push(headers.map(function(h) { return row[h] !== undefined ? row[h] : ''; }));
    } catch(e) {
      errors.push({ row: idx + 1, message: e.message });
    }
  });

  if (errors.length) return { success: false, count: 0, error_count: errors.length, errors: errors.slice(0, 25) };

  if (values.length) {
    var sh = getSpreadsheet_().getSheetByName(cfg.data_sheet);
    if (!sh) throw new Error('Sheet ' + cfg.data_sheet + ' not found. Run Initial Setup.');
    sh.getRange(sh.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    audit_('BATCH_SUBMIT', '', formKey + ' count:' + values.length);
  }
  return { success: true, count: values.length, error_count: 0, errors: [] };
}

// ── submitBulkRows (CSV upload — unchanged) ───────────────────────────────────
function submitBulkRows(formKey, rows) {
  return submitBatch(formKey, rows);
}

// ── reviewWorkflow ────────────────────────────────────────────────────────────
function reviewWorkflow(formKey, recordId, status, remarks) {
  formKey = String(formKey || '').toUpperCase();
  var cfg = FORM_CONFIGS_[formKey];
  if (!cfg) throw new Error('Unknown form key: ' + formKey);
  var headers = getSheetHeaders_(cfg.data_sheet);
  if (!headers.length) throw new Error('No headers in ' + cfg.data_sheet);
  var sh = getSpreadsheet_().getSheetByName(cfg.data_sheet);
  var data = sh.getDataRange().getValues();
  var ridIdx = headers.indexOf('record_id');
  if (ridIdx < 0) throw new Error('record_id column not found.');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ridIdx]) === String(recordId)) {
      var set = function(col, val) {
        var idx = headers.indexOf(col);
        if (idx >= 0) sh.getRange(i + 1, idx + 1).setValue(val);
      };
      set('workflow_status',  status   || 'Pending');
      set('workflow_remarks', remarks  || '');
      set('reviewed_at',      new Date());
      set('reviewed_by',      Session.getActiveUser().getEmail() || '');
      audit_('REVIEW', recordId, 'status:' + status);
      return { success: true };
    }
  }
  throw new Error('Record not found: ' + recordId);
}

// ── getRecordsPage ────────────────────────────────────────────────────────────
function getRecordsPage(formKey, filters, page, pageSize, sortCol, sortDir) {
  formKey  = String(formKey  || 'FORM1').toUpperCase();
  var cfg  = FORM_CONFIGS_[formKey];
  if (!cfg) return { rows: [], page: 1, total_pages: 1, total: 0 };
  filters  = filters  || {};
  page     = Math.max(1, parseInt(page)     || 1);
  pageSize = Math.max(1, Math.min(200, parseInt(pageSize) || 10));
  sortCol  = sortCol  || 'submitted_at';
  sortDir  = String(sortDir || 'desc').toLowerCase();

  var rows = getSheetObjects_(cfg.data_sheet);
  if (!rows.length) return { rows: [], page: 1, total_pages: 1, total: 0 };

  var f = rows.filter(function(r) {
    if (filters.circle_name      && String(r.circle_name      || '') !== filters.circle_name)      return false;
    if (filters.division_name    && String(r.division_name    || '') !== filters.division_name)    return false;
    if (filters.subdivision_name && String(r.subdivision_name || '') !== filters.subdivision_name) return false;
    if (filters.district         && String(r.district         || '') !== filters.district)         return false;
    if (filters.workflow_status  && String(r.workflow_status  || '') !== filters.workflow_status)  return false;
    if (filters.entry_flag       && String(r.entry_flag       || '') !== filters.entry_flag)       return false;
    if (filters.q) {
      var q    = String(filters.q).toLowerCase();
      var blob = [r.record_id, r.officer_name, r.circle_name, r.division_name,
                  r.district, r.app_no, r.submitted_by_email].join(' ').toLowerCase();
      if (blob.indexOf(q) < 0) return false;
    }
    return true;
  });

  f.sort(function(a, b) {
    var av = a[sortCol], bv = b[sortCol];
    if (av instanceof Date && bv instanceof Date) return sortDir === 'asc' ? av - bv : bv - av;
    av = String(av || '').toLowerCase();
    bv = String(bv || '').toLowerCase();
    return sortDir === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av);
  });

  var total       = f.length;
  var total_pages = Math.max(1, Math.ceil(total / pageSize));
  page            = Math.min(page, total_pages);
  var pageRows    = f.slice((page - 1) * pageSize, page * pageSize).map(serializeRow_);
  return { rows: pageRows, page: page, total_pages: total_pages, total: total };
}

// ── getFilterOptions ──────────────────────────────────────────────────────────
function getFilterOptions(formKey) {
  try {
    formKey = String(formKey || 'FORM1').toUpperCase();
    var cfg = FORM_CONFIGS_[formKey];
    if (!cfg) return { circle_name: [], division_name: [], subdivision_name: [], district: [], workflow_status: [], entry_flag: [] };
    var rows = getSheetObjects_(cfg.data_sheet) || [];
    var uniq = function(k) {
      var seen = {}, out = [];
      rows.forEach(function(r) {
        var v = String(r[k] || '').trim();
        if (v && !seen[v]) { seen[v] = 1; out.push(v); }
      });
      return out.sort();
    };
    return {
      circle_name:      uniq('circle_name'),
      division_name:    uniq('division_name'),
      subdivision_name: uniq('subdivision_name'),
      district:         uniq('district'),
      workflow_status:  uniq('workflow_status'),
      entry_flag:       uniq('entry_flag')
    };
  } catch(e) {
    Logger.log('getFilterOptions error: ' + e.message);
    return { circle_name: [], division_name: [], subdivision_name: [], district: [], workflow_status: [], entry_flag: [] };
  }
}

// ── buildRow_ ────────────────────────────────────────────────────────────────
function buildRow_(p, formKey) {
  var now   = new Date();
  var ts    = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var rid   = (formKey === 'FORM1' ? 'SP-' : 'SV-') + ts + '-' + (Math.floor(Math.random() * 9000) + 1000);
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}

  var base = {
    record_id:          rid,
    batch_id:           p.batch_id           || rid,
    form_key:           formKey,
    submitted_at:       now,
    submitted_by_email: email,
    // Section 1
    officer_name:       p.officer_name  || '',
    officer_phone:      p.officer_phone || '',
    // Section 2
    zone_name:          p.zone_name          || '',
    circle_name:        p.circle_name        || '',
    division_name:      p.division_name      || '',
    subdivision_name:   p.subdivision_name   || '',
    // Section 3
    district:           p.district           || '',
    taluk:              p.taluk              || '',
    region_type:        p.region_type        || '',
    corporation_name:   p.corporation_name   || '',
    ulb_body:           p.ulb_body           || '',
    // Gate 1 (plot_sqft stored per-row from shared batch value)
    // Workflow
    workflow_status:    'Pending',
    workflow_remarks:   '',
    reviewed_at:        '',
    reviewed_by:        ''
  };

  if (formKey === 'FORM1') {
    Object.assign(base, {
      // Section 4
      app_no:              p.app_no              || '',
      app_date:            p.app_date            || '',
      // Gate 2 removed
      // Gate 3
      plot_sqft:           p.plot_sqft           || '',
      // Section 5
      no_of_floors:        p.no_of_floors        || '',
      no_of_connections:   p.no_of_connections   || '',
      total_sanctioned_kw: p.total_sanctioned_kw || '',
      date_of_service:     p.date_of_service     || '',
      any_other_info:      p.any_other_info      || '',
      // Entry flag
      entry_flag:          p.entry_flag          || 'SERVICED'
    });
  } else {
    // Derive eligibility flag for Survey (Form 2)
    var cat    = String(p.plot_category || '').trim();
    var floors = String(p.no_of_floors || '').trim();
    var eligibleFloors = ['Ground','G+1','G+2','G+3','S+1','S+2','S+3','S+4'];
    var isEligible = (cat === 'upto_2880') && (eligibleFloors.indexOf(floors) !== -1);

    Object.assign(base, {
      plot_category:        p.plot_category                  || '',
      plot_sqft:            parseFloat(p.plot_sqft)          || '',
      no_of_floors:         p.no_of_floors                   || '',
      expected_connections: parseInt(p.expected_connections)  || '',
      expected_load_kw:     parseFloat(p.expected_load_kw)   || '',
      temp_conn_applied:    p.temp_conn_applied               || '',
      temp_account_id:      p.temp_account_id                || '',
      app_reg_date:         p.app_reg_date                   || '',
      bldg_ready_occupied:  p.bldg_ready_occupied            || '',
      remarks:              p.remarks                        || '',
      entry_flag:           isEligible ? 'ELIGIBLE' : 'NOT-ELIGIBLE'
    });
  }

  return base;
}


// Helper: check duplicate Application Number (Form1) across existing data
function isDuplicateAppNo_(appNoDigits) {
  if (!appNoDigits) return false;
  var rows = getSheetObjects_('DATA_SERVICE_PROGRESS') || [];
  var target = String(appNoDigits);
  for (var i = 0; i < rows.length; i++) {
    var existing = String(rows[i].app_no || '').replace(/[^0-9]/g, '');
    if (existing && existing === target) return true;
  }
  return false;
}

// ── validateRequired_ ─────────────────────────────────────────────────────────
function validateRequired_(row, formKey) {
  var base = ['officer_name', 'officer_phone', 'circle_name', 'division_name',
              'subdivision_name', 'district', 'taluk', 'region_type'];

  if (formKey === 'FORM1') {
    // plot_sqft is a shared batch value required for every entry
    base = base.concat(['plot_sqft', 'app_no', 'app_date',
                        'no_of_floors', 'no_of_connections']);
    // date_of_service only required for eligible (building ready) entries
    base = base.concat(['date_of_service']);
  }

  if (formKey === 'FORM2') {
    base = base.concat(['plot_category', 'plot_sqft', 'no_of_floors',
                        'temp_conn_applied', 'bldg_ready_occupied']);
    // Expected connections and load only required for plots up to 2880 sq ft
    if (row.plot_category === 'upto_2880') {
      base = base.concat(['expected_connections', 'expected_load_kw']);
    }
    // corporation_name and ulb_body dependencies removed
  }

  var missing = base.filter(function(field) {
    return !row[field] && row[field] !== 0;
  });
  if (missing.length) throw new Error('Required fields missing: ' + missing.join(', '));
}

// ── validateBusinessRules_ ──────────────────────────────────────────────────
function validateBusinessRules_(row, formKey) {
  if (formKey === 'FORM1') {
    // Officer phone — must be 10 digits
    var appDigits = String(row.app_no || '').replace(/[^0-9]/g,'');
    if (!appDigits) throw new Error('Application Number must contain only digits.');
    if (appDigits.length > 12) throw new Error('Application Number cannot exceed 12 digits.');
    if (isDuplicateAppNo_(appDigits)) throw new Error('Application Number already exists in submitted records.');
    // Officer phone — must be 10 digits
    var phone = String(row.officer_phone || '').replace(/\D/g, '');
    if (phone && phone.length !== 10) throw new Error('Officer Phone must be a 10-digit mobile number.');

    // Plot size — must be > 0 and <= 2880 (ineligible plots blocked client-side but validated server-side too)
    var sqft = parseFloat(row.plot_sqft || 0);
    if (!sqft || sqft <= 0) throw new Error('Plot Size must be greater than zero.');
    if (sqft > 2880) throw new Error('Plot size above 2880 sq ft is not eligible for service.');

    // Connections — must be > 0
    var conns = parseFloat(row.no_of_connections || 0);
    if (row.no_of_connections !== '' && row.no_of_connections !== null && row.no_of_connections !== undefined) {
      if (!conns || conns <= 0) throw new Error('Number of Connections must be greater than zero.');
    }

    // Date of Service — always required, must be >= 2026-06-06
    if (!row.date_of_service) throw new Error('Date of Service is required.');
    var dosStr = String(row.date_of_service || '').substring(0, 10);
    if (dosStr && dosStr < '2026-06-06') throw new Error('Date of Service cannot be earlier than 06.06.2026.');
  }

  if (formKey === 'FORM2') {
    var cat   = row.plot_category || '';
    var psqft = parseFloat(row.plot_sqft || 0);

    // For plots up to 2880 sq ft, keep strict validation.
    if (cat === 'upto_2880') {
      if (!psqft || psqft <= 0)
        throw new Error('Plot Area is required.');
      if (psqft > 2880)
        throw new Error('Plot area exceeds 2880 sq ft for "Up to 2880" category.');
    }
    // For 2881 and beyond categories, skip all plot-area range validations.

    // Temp connection date validation (applies to all categories when temp = Yes)
    if (String(row.temp_conn_applied).toLowerCase() === 'yes') {
      var appDate = parseDateSafe_(row.app_reg_date);
      if (!appDate) throw new Error('Application Registration Date is required when Temporary Connection = Yes.');
      appDate.setHours(0,0,0,0);
      if (appDate > new Date('2026-05-31T00:00:00'))
        throw new Error('Application Registration Date cannot be beyond 31.05.2026.');
    }
  }
}

// ── serializeRow_ ─────────────────────────────────────────────────────────────
function serializeRow_(row) {
  var tz  = Session.getScriptTimeZone();
  var out = {};
  Object.keys(row || {}).forEach(function(k) {
    var v = row[k];
    if (v instanceof Date) {
      out[k] = isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
    } else if (v === null || v === undefined) {
      out[k] = '';
    } else {
      out[k] = v;
    }
  });
  return out;
}
