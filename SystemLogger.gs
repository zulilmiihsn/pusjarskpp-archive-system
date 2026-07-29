'use strict';

const SYSTEM_LOG_SHEET_NAME = 'system_logs';

const SystemLogger = {
  log: function(level, action, message, metadata) {
    try {
      const ssId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID);
      if (!ssId) return; // Silent fallback if config SS is not set
      
      const ss = openSpreadsheetById_(ssId);
      const sheet = getOrCreateSheet_(ss, SYSTEM_LOG_SHEET_NAME);
      ensureHeaders_(sheet, ['timestamp', 'level', 'user', 'action', 'message', 'metadata']);
      
      const metaObject = metadata && typeof metadata === 'object' ? metadata : null;
      const user = metaObject && metaObject.portalUsername ? String(metaObject.portalUsername) : 'system';
      const metaStr = metadata ? (typeof metadata === 'string' ? metadata : JSON.stringify(metadata)) : '';
      
      sheet.appendRow([
        new Date().toISOString(),
        level,
        user,
        action,
        message,
        metaStr
      ]);
    } catch (e) {
      console.error('SystemLogger failed:', e);
    }
  },
  
  info: function(action, message, metadata) {
    this.log('INFO', action, message, metadata);
  },
  
  warn: function(action, message, metadata) {
    this.log('WARN', action, message, metadata);
  },
  
  error: function(action, message, error) {
    this.log('ERROR', action, message, error ? error.toString() : '');
  }
};
