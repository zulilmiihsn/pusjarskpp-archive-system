'use strict';

const SYSTEM_LOG_SHEET_NAME = 'system_logs';
const SYSTEM_LOG_HEADERS_ = [
  'timestamp', 'level', 'request_id', 'endpoint', 'user', 'role', 'action',
  'duration_ms', 'outcome', 'message', 'metadata'
];

const SystemLogger = {
  log: function (level, action, message, metadata) {
    try {
      const ssId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID);
      if (!ssId) return;

      const ss = openSpreadsheetById_(ssId);
      const sheet = getOrCreateSheet_(ss, SYSTEM_LOG_SHEET_NAME);
      ensureHeaders_(sheet, SYSTEM_LOG_HEADERS_);

      const safeMeta = typeof redactLogMetadata_ === 'function'
        ? redactLogMetadata_(metadata && typeof metadata === 'object' ? metadata : { detail: metadata })
        : (metadata && typeof metadata === 'object' ? metadata : { detail: metadata });
      const context = typeof getApiRequestContext_ === 'function' ? (getApiRequestContext_() || {}) : {};
      const row = {
        timestamp: new Date().toISOString(),
        level: String(level || 'INFO').toUpperCase(),
        request_id: String(safeMeta.requestId || context.requestId || ''),
        endpoint: String(safeMeta.endpoint || context.endpoint || ''),
        user: String(safeMeta.portalUsername || context.actor || 'system'),
        role: String(safeMeta.portalRole || context.role || 'system'),
        action: String(action || ''),
        duration_ms: Number(safeMeta.durationMs || 0) || '',
        outcome: String(safeMeta.outcome || ''),
        message: String(message || '').slice(0, 2000),
        metadata: JSON.stringify(safeMeta).slice(0, 8000)
      };
      const headers = getHeaders_(sheet);
      sheet.appendRow(headers.map(function (header) {
        return row[header] === undefined ? '' : row[header];
      }));
    } catch (error) {
      console.error('SystemLogger failed:', error);
    }
  },

  info: function (action, message, metadata) {
    this.log('INFO', action, message, metadata || {});
  },

  warn: function (action, message, metadata) {
    this.log('WARN', action, message, metadata || {});
  },

  error: function (action, message, errorOrMetadata) {
    let metadata = errorOrMetadata || {};
    if (errorOrMetadata instanceof Error) {
      metadata = {
        errorName: errorOrMetadata.name,
        errorMessage: errorOrMetadata.message,
        errorCode: errorOrMetadata.errorCode || ''
      };
    }
    this.log('ERROR', action, message, metadata);
  }
};
