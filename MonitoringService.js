'use strict';

const SYSTEM_HEALTH_REQUIRED_SHEETS_ = [
  CONFIG_SHEETS.YEARS,
  CONFIG_SHEETS.ACTIVITIES,
  CONFIG_SHEETS.SUB_ACTIVITIES,
  CONFIG_SHEETS.METADATA_FIELDS,
  CONFIG_SHEETS.ARCHIVE_LOG,
  CONFIG_SHEETS.ADMIN_AUDIT_LOG,
  CONFIG_SHEETS.ACCOUNTS
];

const MonitoringService = {
  check: function () {
    const startedAt = Date.now();
    const checks = [];
    const metrics = {};
    function add(name, ok, severity, detail) {
      checks.push({ name: name, ok: !!ok, severity: severity || 'ERROR', detail: String(detail || '') });
    }

    const props = PropertiesService.getScriptProperties();
    const propertyKeys = props.getKeys();
    metrics.scriptPropertyCount = propertyKeys.length;
    metrics.activeSessionCount = propertyKeys.filter(function (key) { return key.indexOf('sess_') === 0; }).length;
    metrics.resumableSessionCount = propertyKeys.filter(function (key) { return key.indexOf(RESUMABLE_SESSION_PROPERTY_PREFIX) === 0; }).length;

    const configId = cleanId_(props.getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID));
    add('config_spreadsheet_property', !!configId, 'CRITICAL', configId ? 'configured' : 'missing');

    let ss = null;
    if (configId) {
      try {
        ss = openSpreadsheetById_(configId);
        add('config_spreadsheet_access', !!ss, 'CRITICAL', ss ? ss.getName() : 'unavailable');
      } catch (error) {
        add('config_spreadsheet_access', false, 'CRITICAL', error.message);
      }
    }

    if (ss) {
      const names = ss.getSheets().map(function (sheet) { return sheet.getName(); });
      const missing = SYSTEM_HEALTH_REQUIRED_SHEETS_.filter(function (name) { return names.indexOf(name) < 0; });
      add('required_config_sheets', missing.length === 0, 'ERROR', missing.length ? 'missing: ' + missing.join(', ') : 'complete');
      metrics.configSheetCount = names.length;
    }

    let settings = {};
    try {
      settings = ConfigService.getSettings() || {};
      add('settings_read', true, 'CRITICAL', 'ok');
      add('current_year', !!settings.currentYear, 'ERROR', settings.currentYear || 'missing');
    } catch (error) {
      add('settings_read', false, 'CRITICAL', error.message);
    }

    let roots = {};
    try {
      roots = _collectWorkspaceRootIds_(settings.currentYear);
      const rootIds = Object.keys(roots);
      metrics.workspaceRootCount = rootIds.length;
      add('workspace_roots_configured', rootIds.length > 0, 'CRITICAL', rootIds.length + ' root(s)');
      rootIds.slice(0, 5).forEach(function (id, index) {
        try {
          const meta = Drive.Files.get(id, { fields: 'id,name,mimeType,trashed', supportsAllDrives: true });
          const valid = !!meta && meta.mimeType === 'application/vnd.google-apps.folder' && !meta.trashed;
          add('workspace_root_' + (index + 1), valid, 'CRITICAL', valid ? String(meta.name || 'folder') : 'invalid/trashed');
        } catch (error) {
          add('workspace_root_' + (index + 1), false, 'CRITICAL', error.message);
        }
      });
    } catch (error) {
      add('workspace_roots_read', false, 'CRITICAL', error.message);
    }

    try {
      const handlers = ScriptApp.getProjectTriggers().map(function (trigger) {
        return trigger.getHandlerFunction ? trigger.getHandlerFunction() : '';
      });
      metrics.triggerCount = handlers.length;
      add('maintenance_trigger', handlers.indexOf('runArchiveMaintenance') >= 0, 'WARN', handlers.join(', ') || 'none');
    } catch (error) {
      add('trigger_inventory', false, 'WARN', error.message);
    }

    let status = 'HEALTHY';
    if (checks.some(function (check) { return !check.ok && check.severity === 'CRITICAL'; })) status = 'CRITICAL';
    else if (checks.some(function (check) { return !check.ok; })) status = 'DEGRADED';

    return {
      status: status,
      checkedAt: new Date().toISOString(),
      durationMs: Date.now() - startedAt,
      checks: checks,
      metrics: metrics
    };
  }
};

function runSystemHealthMonitor_() {
  const health = MonitoringService.check();
  const failed = health.checks.filter(function (check) { return !check.ok; });
  const metadata = {
    status: health.status,
    durationMs: health.durationMs,
    failedChecks: failed,
    metrics: health.metrics
  };
  if (health.status === 'HEALTHY') {
    SystemLogger.info('SYSTEM_HEALTH', 'Pemeriksaan kesehatan sistem lulus.', metadata);
  } else {
    SystemLogger.warn('SYSTEM_HEALTH', 'Pemeriksaan kesehatan sistem menemukan masalah.', metadata);
  }
  return health;
}
