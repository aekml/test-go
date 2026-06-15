// ═══════════════════════════════════════════════════════════════════════════
// MasterService.gs — loads master data for bootstrap
// MASTER_BESCOM  cols: zone, circle, division, subdivision, active
// MASTER_GEO     cols: district, taluk, active
// ═══════════════════════════════════════════════════════════════════════════

function getMastersPayload() {
  function active(arr) {
    return (arr || []).filter(function(r) {
      return String(r.active || 'TRUE').toUpperCase() !== 'FALSE';
    });
  }
  return {
    bescom:      active(getSheetObjects_('MASTER_BESCOM')),
    geo:         active(getSheetObjects_('MASTER_GEO')),
    corporation: active(getSheetObjects_('MASTER_CORPORATION')),
    ulb:         active(getSheetObjects_('MASTER_ULB'))
  };
}
