// ── DashboardService.gs ──────────────────────────────────────────────────────
function getDashboardData(formKey) {
  try {
    formKey = String(formKey || 'FORM1').toUpperCase();
    var sheetMap = { FORM1: 'DATA_SERVICE_PROGRESS', FORM2: 'DATA_SURVEY' };
    var sheetName = sheetMap[formKey];
    if (!sheetName) throw new Error('Unknown formKey: ' + formKey);
    var rows = getSheetObjects_(sheetName) || [];

    // KPIs
    var kpi = {};
    if (formKey === 'FORM1') {
      kpi.total_buildings   = rows.length;
      kpi.total_connections = rows.reduce(function(s,r){ return s+(parseFloat(r.no_of_connections)||0); },0);
      kpi.total_load_kw     = rows.reduce(function(s,r){ return s+(parseFloat(r.total_sanctioned_kw)||0); },0);
      kpi.eligible          = rows.filter(function(r){ return String(r.entry_flag||'')===('ELIGIBLE'); }).length;
      kpi.bldg_not_ready    = rows.filter(function(r){ return String(r.entry_flag||'')===('BLDG_NOT_READY'); }).length;
      kpi.oversized         = rows.filter(function(r){ return String(r.entry_flag||'')===('OVERSIZED'); }).length;
    } else {
      kpi.total_buildings       = rows.length;
      kpi.total_exp_connections = rows.reduce(function(s,r){ return s+(parseFloat(r.expected_connections)||0); },0);
      kpi.total_exp_load_kw     = rows.reduce(function(s,r){ return s+(parseFloat(r.expected_load_kw)||0); },0);
      kpi.app_registered        = rows.filter(function(r){ return String(r.temp_conn_applied||'').toLowerCase()==='yes'; }).length;
      kpi.bldg_occupied         = rows.filter(function(r){ return String(r.bldg_ready_occupied||'').toLowerCase()==='yes'; }).length;
    }

    // helpers
    function groupCount(field) {
      var m={};
      rows.forEach(function(r){ var v=String(r[field]||'—').trim(); m[v]=(m[v]||0)+1; });
      var keys=Object.keys(m).sort(function(a,b){ return m[b]-m[a]; });
      return { labels:keys, values:keys.map(function(k){ return m[k]; }) };
    }
    function submissionsOverTime() {
      var m={};
      rows.forEach(function(r){
        if(!r.submitted_at) return;
        var d=new Date(r.submitted_at); if(isNaN(d)) return;
        var key=Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
        m[key]=(m[key]||0)+1;
      });
      var keys=Object.keys(m).sort();
      return { labels:keys, values:keys.map(function(k){ return m[k]; }) };
    }
    function appRegisteredBy(groupField) {
      var yes={},no={};
      rows.forEach(function(r){
        var v=String(r[groupField]||'—').trim();
        var filled=String(r.app_no||'').trim().length>0;
        if(filled){ yes[v]=(yes[v]||0)+1; } else { no[v]=(no[v]||0)+1; }
      });
      var keys=Object.keys(Object.assign({},yes,no)).sort(function(a,b){
        return ((yes[b]||0)+(no[b]||0))-((yes[a]||0)+(no[a]||0));
      });
      return { labels:keys, registered:keys.map(function(k){ return yes[k]||0; }), not_registered:keys.map(function(k){ return no[k]||0; }) };
    }

    var charts={};
    charts.by_circle      = groupCount('circle_name');
    charts.by_division    = groupCount('division_name');
    charts.by_subdivision = groupCount('subdivision_name');
    charts.region_type    = groupCount('region_type');
    charts.over_time      = submissionsOverTime();
    if (formKey==='FORM1') {
      charts.entry_flag        = groupCount('entry_flag');
      charts.app_reg_by_circle = appRegisteredBy('circle_name');
      charts.app_reg_by_div    = appRegisteredBy('division_name');
    } else {
      charts.plot_category = groupCount('plot_category');
    }

    return { success:true, formKey:formKey, kpi:kpi, charts:charts };
  } catch(e) {
    Logger.log('getDashboardData error: '+e.message);
    return { success:false, error:e.message };
  }
}

function getRecentSubmissions(filters) {
  try {
    filters = filters||{};
    var f1 = (getSheetObjects_('DATA_SERVICE_PROGRESS')||[]).map(function(r){ r._form='FORM1'; return r; });
    var f2 = (getSheetObjects_('DATA_SURVEY')||[]).map(function(r){ r._form='FORM2'; return r; });
    var all = f1.concat(f2);
    all.sort(function(a,b){
      var da=a.submitted_at?new Date(a.submitted_at):new Date(0);
      var db=b.submitted_at?new Date(b.submitted_at):new Date(0);
      return db-da;
    });
    if (filters.form && filters.form!=='all')
      all=all.filter(function(r){ return r._form===filters.form.toUpperCase(); });
    if (filters.circle)
      all=all.filter(function(r){ return String(r.circle_name||'')===filters.circle; });
    if (filters.workflow_status)
      all=all.filter(function(r){ return String(r.workflow_status||'')===filters.workflow_status; });

    var total=all.length;
    var tz=Session.getScriptTimeZone();
    var rows=all.slice(0,100).map(function(r){
      var d=r.submitted_at?new Date(r.submitted_at):null;
      return {
        record_id:       r.record_id||'',
        form:            r._form||'',
        officer_name:    r.officer_name||'',
        circle_name:     r.circle_name||'',
        division_name:   r.division_name||'',
        subdivision_name:r.subdivision_name||'',
        district:        r.district||'',
        taluk:           r.taluk||'',
        region_type:     r.region_type||'',
        entry_flag:      r.entry_flag||'',
        workflow_status: r.workflow_status||'Pending',
        submitted_at:    d?Utilities.formatDate(d,tz,'dd-MM-yyyy HH:mm'):''
      };
    });
    return { success:true, total:total, rows:rows };
  } catch(e) {
    return { success:false, error:e.message, rows:[] };
  }
}
