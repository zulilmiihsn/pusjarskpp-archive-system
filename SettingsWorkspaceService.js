'use strict';

/** @private Workspace lifecycle and maintenance operations. */
const SettingsWorkspaceImpl_ = {
  resetWorkspace: function () {
    const settings = ConfigService.getSettings();
    let removed = 0;

    // Trash config spreadsheet
    const configId = settings.configSpreadsheetId;
    if (configId) {
      try { DriveApp.getFileById(cleanId_(configId)).setTrashed(true); removed++; } catch (e) {}
    }

    // Trash activity spreadsheets
    try {
      const config = CacheHelper.getConfig(settings.currentYear || DEFAULT_YEAR);
      if (config && config.activities) {
        config.activities.forEach(function (a) {
          if (a.spreadsheet_file_id) {
            try { DriveApp.getFileById(cleanId_(a.spreadsheet_file_id)).setTrashed(true); removed++; } catch (e) {}
          }
        });
      }
    } catch (e) {}

    // Clear all workspace settings
    ConfigService.saveSettings({
      configSpreadsheetId: '',
      workspaceRootFolderId: '',
      referenceRootFolderId: '',
      systemFolderParentId: '',
      systemFolderName: '00. Sistem Portal',
      configFolderId: '',
      currentYear: DEFAULT_YEAR
    });

    CacheHelper.invalidateAll();
    return { success: true, removed: removed };
  },

  initializeWorkspace: function(payload) {
    payload = payload || {};
    Validator.requireYear(payload.year);
    if (!payload.workspaceRootFolderId && !payload.workspaceRootFolderUrl) {
      throw new Error('Folder Utama Workspace wajib diisi.');
    }

    // Lock saat provisioning: cegah dua init bersamaan saling balapan list->create
    // (TOCTOU) yang bisa bikin folder/spreadsheet duplikat. Aman secara availability
    // karena saat init workspace memang belum siap dipakai mengarsip (B3).
    const result = withLock_(function () {
      return WorkspaceSetupService.initialize(payload);
    }, 30000);
    CacheHelper.invalidateAll();

    const adminResult = AuthService.saveDefaultAdmin();

    try {
      result.cleanupTrigger = this.ensureArchiveMaintenanceTrigger();
    } catch (error) {
      result.cleanupTrigger = {
        installed: false,
        error: error.message || String(error)
      };
    }

    try {
      result.docTypesTrigger = this.ensureDocumentTypesSyncTrigger();
    } catch (error) {
      result.docTypesTrigger = {
        installed: false,
        error: error.message || String(error)
      };
    }

    if (adminResult && adminResult.created) {
      result.defaultPassword = adminResult.defaultPassword;
      if (adminResult.sessionId) result.sessionId = adminResult.sessionId;
    }

    const bootstrap = this.getBootstrap();
    bootstrap.workspaceSetup = result;
    return bootstrap;
  },

  ensureArchiveMaintenanceTrigger: function () {
    const handlerName = 'runArchiveMaintenance';
    const legacyHandler = 'cleanupTrashedSubActivities';
    const triggers = ScriptApp.getProjectTriggers();

    // Migrasi: hapus trigger lama yang menunjuk endpoint ber-auth — di konteks
    // trigger (tanpa sesi) endpoint itu akan selalu throw.
    triggers.forEach(function (trigger) {
      if (trigger.getHandlerFunction && trigger.getHandlerFunction() === legacyHandler) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    const exists = triggers.some(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === handlerName);
    if (exists) {
      return { installed: false, exists: true, handler: handlerName };
    }

    ScriptApp.newTrigger(handlerName)
      .timeBased()
      .everyDays(1)
      .atHour(2)
      .create();

    return { installed: true, exists: false, handler: handlerName };
  },

  // Pasang trigger onEdit pada config spreadsheet supaya perubahan di sheet
  // `config_document_types` langsung menyinkronkan kolom Rekap (tanpa nunggu lazy).
  ensureDocumentTypesSyncTrigger: function () {
    const handlerName = 'onConfigDocumentTypesEdit';
    const settings = ConfigService.getSettings();
    if (!settings.configSpreadsheetId) return { installed: false, reason: 'no_config_spreadsheet', handler: handlerName };
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === handlerName);
    if (exists) return { installed: false, exists: true, handler: handlerName };

    ScriptApp.newTrigger(handlerName)
      .forSpreadsheet(cleanId_(settings.configSpreadsheetId))
      .onEdit()
      .create();

    return { installed: true, exists: false, handler: handlerName };
  },

  // Sapu semua spreadsheet Rekap (per kegiatan/sub-kegiatan) untuk tahun terkait,
  // tambah kolom tipe aktif + hapus kolom tipe nonaktif. Idempoten.
  syncDocumentTypeColumns: function (year) {
    const selectedYear = Number(year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    return withLock_(() => {
      const config = CacheHelper.getConfig(selectedYear);
      const ids = {};
      (config.activities || []).forEach(function (a) {
        const id = cleanId_(a.spreadsheet_file_id || '');
        if (id) ids[id] = true;
      });
      (config.subActivities || []).forEach(function (s) {
        const id = cleanId_(s.spreadsheet_file_id || '');
        if (id) ids[id] = true;
      });

      let spreadsheetsSynced = 0;
      let errors = 0;
      Object.keys(ids).forEach(function (id) {
        try {
          const ss = openSpreadsheetById_(id);
          const rekap = findRekapSheet_(ss);
          if (rekap) {
            ensureRekapDocumentColumns_(rekap);
            spreadsheetsSynced++;
          }
        } catch (e) {
          errors++;
          console.warn('syncDocumentTypeColumns gagal utk ' + id + ': ' + e.message);
        }
      });

      CacheHelper.invalidate(selectedYear);
      bumpVersion();
      return { year: selectedYear, spreadsheetsSynced: spreadsheetsSynced, errors: errors };
    });
  },

};
