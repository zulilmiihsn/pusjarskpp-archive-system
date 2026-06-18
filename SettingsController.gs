'use strict';

const SettingsController = {
  getBootstrap: function() {
    const settings = ConfigService.getSettings();
    if (!settings.configSpreadsheetId) {
      return {
        configured: false,
        settings: settings,
        message: 'Aplikasi belum punya ruang kerja. Pilih Folder Ruang Kerja di Pengaturan.'
      };
    }

    const config = CacheHelper.getConfig(settings.currentYear);
    const activities = config.activities.map((activity) => {
      const subs = config.subActivities
        .filter(sub => sub.activity_id === activity.activity_id)
        .map((sub) => {
          const effectiveFormalName = sub.formal_archive_name || sub.sub_activity_name || '';
          const effectiveSheetName = sub.target_sheet_name || sub.sub_activity_name || '';
          const mappingStatus = sub.mapping_status || inferSubActivityMappingStatus_(sub);
          return Object.assign({}, sub, {
            effective_formal_archive_name: effectiveFormalName,
            effective_target_sheet_name: effectiveSheetName,
            effective_mapping_status: mappingStatus,
            folder: DriveService.folderDtoFromConfig(sub.folder_id, sub.folder_path || sub.sub_activity_name),
            spreadsheetFile: DriveService.fileDtoFromConfig(
              sub.spreadsheet_file_id || activity.spreadsheet_file_id,
              effectiveSheetName || activity.activity_name,
              MimeType.GOOGLE_SHEETS
            )
          });
        });
      const fields = config.fields.filter(
        field => field.activity_id === activity.activity_id && isTrue_(field.is_visible_in_form)
      );
      return Object.assign({}, activity, {
        subActivities: subs,
        fields: fields,
        laciFolder: DriveService.folderDtoFromConfig(activity.laci_folder_id, activity.activity_name || activity.laci_no),
        targetFolder: DriveService.folderDtoFromConfig(activity.target_folder_id, activity.activity_name),
        spreadsheetFile: DriveService.fileDtoFromConfig(
          activity.spreadsheet_file_id,
          activity.activity_name ? 'Daftar Arsip - ' + activity.activity_name : activity.spreadsheet_file_id,
          MimeType.GOOGLE_SHEETS
        )
      });
    });

    let triggerInstalled = false;
    try {
      const handlerName = 'cleanupTrashedSubActivities';
      const triggers = ScriptApp.getProjectTriggers();
      triggerInstalled = triggers.some(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === handlerName);
    } catch (e) {
      console.warn('Error checking triggers: ' + e.message);
    }

    let trashedCount = 0;
    let totalSubActivitiesCount = 0;
    try {
      const ss = ConfigRepository.getConfigSpreadsheet();
      const allSubs = readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
        .filter(row => Number(row.year) === Number(settings.currentYear));
      totalSubActivitiesCount = allSubs.length;
      trashedCount = allSubs.filter(row => !isTrue_(row.is_active) && row.inactive_reason === 'DRIVE_TRASHED').length;
    } catch (e) {
      console.warn('Error loading inactive sub-activities: ' + e.message);
    }

    return {
      configured: true,
      settings: settings,
      selectedYear: config.selectedYear,
      years: config.years,
      activities: activities,
      history: config.history,
      progress: this.buildProgress(config.history, activities),
      maintenance: {
        triggerInstalled: triggerInstalled,
        trashedCount: trashedCount,
        totalCount: totalSubActivitiesCount
      }
    };
  },

  getSettings: function () { return ConfigService.getSettings(); },

  saveSettings: function (payload) {
    const settings = ConfigService.saveSettings(payload || {});
    if (settings.configSpreadsheetId) {
      ConfigRepository.getConfigSpreadsheet();
    }
    CacheHelper.invalidateAll();
    return this.getBootstrap();
  },

  resetWorkspace: function () {
    var settings = ConfigService.getSettings();
    var removed = 0;

    // Trash config spreadsheet
    var configId = settings.configSpreadsheetId;
    if (configId) {
      try { DriveApp.getFileById(cleanId_(configId)).setTrashed(true); removed++; } catch (e) {}
    }

    // Trash activity spreadsheets
    try {
      var config = CacheHelper.getConfig(settings.currentYear || DEFAULT_YEAR);
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

    const result = WorkspaceSetupService.initialize(payload);
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

    if (adminResult && adminResult.created) {
      result.defaultPassword = adminResult.defaultPassword;
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

  buildProgress: function (history, activities) {
    const now = new Date();
    const todayStr = now.toISOString().slice(0, 10);
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    const totals = {
      total: history.length, completed: 0, failed: 0, draft: 0,
      thisWeek: 0, today: 0, completionRate: 0,
      mostActiveActivity: null, latestEntry: null, byActivity: {}
    };

    activities.forEach(function (activity) {
      totals.byActivity[activity.activity_id] = {
        activityName: activity.activity_name, total: 0, completed: 0, failed: 0, draft: 0
      };
    });

    history.forEach(function (row) {
      const status = String(row.status || '').toUpperCase();
      if (status === STATUS.COMPLETED) totals.completed++;
      if (status === STATUS.FAILED) totals.failed++;
      if (status === STATUS.DRAFT) totals.draft++;

      if (row.created_at) {
        try {
          const d = new Date(row.created_at);
          if (!isNaN(d.getTime())) {
            if (d >= weekAgo) totals.thisWeek++;
            if (d.toISOString().slice(0, 10) === todayStr) totals.today++;
            if (!totals.latestEntry || d > new Date(totals.latestEntry.created_at)) {
              totals.latestEntry = { created_at: row.created_at, status: status, activity_id: row.activity_id };
            }
          }
        } catch (e) {
          console.warn('SettingsController.buildProgress date parsing error: ' + e.message);
        }
      }

      if (totals.byActivity[row.activity_id]) {
        totals.byActivity[row.activity_id].total++;
        if (status === STATUS.COMPLETED) totals.byActivity[row.activity_id].completed++;
        if (status === STATUS.FAILED) totals.byActivity[row.activity_id].failed++;
        if (status === STATUS.DRAFT) totals.byActivity[row.activity_id].draft++;
      }
    });

    const finalizedCount = totals.completed + totals.failed;
    totals.completionRate = finalizedCount > 0 ? Math.round((totals.completed / finalizedCount) * 100) : 0;

    let maxTotal = 0;
    Object.keys(totals.byActivity).forEach(function (key) {
      const act = totals.byActivity[key];
      if (act.total > maxTotal) {
        maxTotal = act.total;
        totals.mostActiveActivity = { id: key, name: act.activityName, total: act.total };
      }
    });

    return totals;
  },

  updateActivityMapping: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireId(payload.targetFolderId, 'Target Folder ID');
    Validator.requireId(payload.spreadsheetFileId, 'Spreadsheet File ID');

    const updated = ConfigRepository.updateActivityMapping({
      year: year, activityId: payload.activityId,
      laciFolderId: payload.laciFolderId, targetFolderId: payload.targetFolderId,
      spreadsheetFileId: payload.spreadsheetFileId
    });
    CacheHelper.invalidate(year);
    return { activity: updated, bootstrap: this.getBootstrap() };
  },

  updateSubActivityMapping: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');
    Validator.requireId(payload.folderId, 'Folder ID');

    const config = CacheHelper.getConfig(year);
    const activity = ConfigService.findActivity(config, payload.activityId);
    if (!activity) throw new Error('Kegiatan tidak ditemukan.');
    const previous = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
    if (!previous) throw new Error('Sub-kegiatan tidak ditemukan.');

    if (payload.targetSheetName && payload.targetSheetName !== previous.target_sheet_name) {
      SpreadsheetService.renameSubActivitySheet(activity, previous.target_sheet_name, payload.targetSheetName, previous);
    }

    const updated = ConfigRepository.updateSubActivityMapping({
      year: year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      folderId: payload.folderId,
      folderPath: payload.folderPath,
      targetSheetName: payload.targetSheetName,
      formalArchiveName: payload.formalArchiveName,
      noFolder: payload.noFolder,
      mappingStatus: payload.mappingStatus || inferSubActivityMappingStatus_({
        sub_activity_name: payload.subActivityName,
        formal_archive_name: payload.formalArchiveName,
        target_sheet_name: payload.targetSheetName
      }),
      mappingNote: payload.mappingNote,
      rekapRowNumber: payload.rekapRowNumber,
      subActivityName: payload.subActivityName,
      parentFolderId: payload.parentFolderId,
      parentFolderName: payload.parentFolderName,
      parentFolderPath: payload.parentFolderPath,
      spreadsheetFileId: payload.spreadsheetFileId
    });
    SpreadsheetService.ensureSubActivitySheet(activity, updated);
    SpreadsheetService.updateRekapSubActivityIdentity(activity, previous, updated);
    CacheHelper.invalidate(year);
    return { subActivity: updated, bootstrap: this.getBootstrap() };
  },

  renameDriveItem: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.itemId, 'Item ID');
    Validator.requireString(payload.name, 'Nama baru');

    const item = (String(payload.type || 'folder').toLowerCase() === 'file')
      ? DriveService.renameFile(payload.itemId, payload.name)
      : DriveService.renameFolder(payload.itemId, payload.name);
    return { item: item, bootstrap: this.getBootstrap() };
  },

  listDriveFolders: function (payload) { return DriveService.listFolders(payload || {}); },

  getHistory: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    return CacheHelper.getConfig(year).history;
  },

  getTemplates: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    return DriveService.listTemplates(year);
  },

  getTemplatesData: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const cacheKey = 'tpl_data_' + year;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.categories) return parsed;
      }
    } catch (_) {}
    const driveData = DriveService.listTemplatesByCategory(year);
    const categoryMetaRows = ConfigRepository.getTemplateCategories();
    const catMetaById = {};
    categoryMetaRows.forEach(function (row) {
      if (row.category_id) catMetaById[row.category_id] = row;
    });

    const result = {
      categories: (driveData.categories || []).map(function (category) {
        const meta = catMetaById[category.id] || {};
        return Object.assign({}, category, {
          color: normalizeHexColor_(meta.color, DEFAULT_TEMPLATE_CATEGORY_COLOR),
          sort_order: meta.sort_order || ''
        });
      }).sort(function (a, b) {
        return Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
          String(a.name || '').localeCompare(String(b.name || ''), 'id', { sensitivity: 'base' });
      }),
      uncategorized: driveData.uncategorized || []
    };
    try {
      const serialized = JSON.stringify(result);
      if (serialized.length < 95000) {
        CacheService.getScriptCache().put(cacheKey, serialized, 600);
      }
    } catch (_) {}
    return result;
  },

  getAdminAuditLogs: function (payload) {
    payload = payload || {};
    return ConfigRepository.getAdminAuditLogs(Number(payload.limit || 100));
  },

  getTemplateCategories: function () {
    return ConfigRepository.getTemplateCategories();
  },

  saveTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.name, 'Nama kategori');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const folder = DriveService.createTemplateCategory(year, payload.name);
    const meta = ConfigRepository.saveTemplateCategory({
      categoryId: folder.id,
      name: folder.name,
      color: payload.color
    });
    return {
      id: folder.id,
      name: folder.name,
      color: meta.color
    };
  },

  renameTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.categoryId, 'Category ID');
    Validator.requireString(payload.name, 'Nama baru');
    const folder = DriveService.renameTemplateCategoryFolder(payload.categoryId, payload.name);
    const meta = ConfigRepository.saveTemplateCategory({
      categoryId: folder.id,
      name: folder.name,
      color: payload.color
    });
    return {
      id: folder.id,
      name: folder.name,
      color: meta.color
    };
  },

  deleteTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.categoryId, 'Category ID');
    const result = DriveService.deleteTemplateCategoryFolder(payload.categoryId);
    ConfigRepository.deleteTemplateCategory(payload.categoryId);
    return result;
  },

  deleteYear: function (payload) {
    payload = payload || {};
    const year = Number(payload.year);
    if (!year || isNaN(year)) {
      throw new Error('Tahun kerja tidak valid.');
    }

    const settings = ConfigService.getSettings();
    const ss = ConfigRepository.getConfigSpreadsheet();

    // 1. Find the year record in config_years
    const yearsSheet = ss.getSheetByName(CONFIG_SHEETS.YEARS);
    if (!yearsSheet) throw new Error('Sheet config_years tidak ditemukan.');
    const yearsValues = yearsSheet.getDataRange().getDisplayValues();
    const yearsHeaders = yearsValues[0].map(normalizeHeader_);
    const yearCol = yearsHeaders.indexOf('year');
    
    let yearRowIndex = -1;
    let spreadsheetYearFolderId = '';
    let persuratanYearFolderId = '';

    if (yearCol !== -1) {
      for (let i = 1; i < yearsValues.length; i++) {
        if (Number(yearsValues[i][yearCol]) === year) {
          yearRowIndex = i + 1;
          const spreadsheetFolderCol = yearsHeaders.indexOf('spreadsheet_year_folder_id');
          const persuratanFolderCol = yearsHeaders.indexOf('persuratan_year_folder_id');
          if (spreadsheetFolderCol !== -1) spreadsheetYearFolderId = yearsValues[i][spreadsheetFolderCol];
          if (persuratanFolderCol !== -1) persuratanYearFolderId = yearsValues[i][persuratanFolderCol];
          break;
        }
      }
    }

    // 2. Move associated folders in Google Drive to Trash
    const trashReport = [];
    
    // Trash spreadsheet year folder (ARSIP DIKLAT {year})
    if (spreadsheetYearFolderId) {
      try {
        const folder = DriveApp.getFolderById(spreadsheetYearFolderId);
        folder.setTrashed(true);
        trashReport.push('Folder Arsip Diklat berhasil dibuang ke sampah.');
      } catch (e) {
        console.warn('Error trashing spreadsheet year folder: ' + e.message);
        trashReport.push('Gagal membuang Folder Arsip Diklat: ' + e.message);
      }
    }

    // Trash persuratan year folder (Tahun {year} under Persuratan) and try to find/trash dokumen year folder (Tahun {year} under Dokumen)
    if (persuratanYearFolderId) {
      try {
        const folder = DriveApp.getFolderById(persuratanYearFolderId);
        
        // Find document year folder under naskahDinas/2. Dokumen/Tahun {year}
        try {
          const parentFolders = folder.getParentFolders();
          if (parentFolders.hasNext()) {
            const persuratanFolder = parentFolders.next();
            const naskahDinasFolders = persuratanFolder.getParentFolders();
            if (naskahDinasFolders.hasNext()) {
              const naskahDinas = naskahDinasFolders.next();
              const docFolders = naskahDinas.getFoldersByName('2. Dokumen');
              if (docFolders.hasNext()) {
                const docFolder = docFolders.next();
                const docYearFolders = docFolder.getFoldersByName('Tahun ' + year);
                if (docYearFolders.hasNext()) {
                  docYearFolders.next().setTrashed(true);
                  trashReport.push('Folder Dokumen Tahun ' + year + ' berhasil dibuang ke sampah.');
                }
              }
            }
          }
        } catch (innerErr) {
          console.warn('Error locating and trashing doc year folder: ' + innerErr.message);
        }

        folder.setTrashed(true);
        trashReport.push('Folder Persuratan Tahun ' + year + ' berhasil dibuang ke sampah.');
      } catch (e) {
        console.warn('Error trashing persuratan year folder: ' + e.message);
        trashReport.push('Gagal membuang Folder Persuratan: ' + e.message);
      }
    }

    // 3. Delete config rows matching the year in all configuration sheets
    const sheetsToDelete = [
      { name: CONFIG_SHEETS.YEARS, col: 'year' },
      { name: CONFIG_SHEETS.ACTIVITIES, col: 'year' },
      { name: CONFIG_SHEETS.SUB_ACTIVITIES, col: 'year' },
      { name: CONFIG_SHEETS.METADATA_FIELDS, col: 'year' },
      { name: CONFIG_SHEETS.ARCHIVE_LOG, col: 'year' }
    ];

    sheetsToDelete.forEach(function (sheetConfig) {
      try {
        const sheet = ss.getSheetByName(sheetConfig.name);
        if (sheet) {
          const values = sheet.getDataRange().getValues();
          if (values.length >= 2) {
            const headers = values[0].map(normalizeHeader_);
            const colIdx = headers.indexOf(sheetConfig.col);
            if (colIdx !== -1) {
              let deleteCount = 0;
              for (let i = values.length - 1; i >= 1; i--) {
                if (Number(values[i][colIdx]) === year) {
                  sheet.deleteRow(i + 1);
                  deleteCount++;
                }
              }
              console.log('Deleted ' + deleteCount + ' rows from ' + sheetConfig.name);
            }
          }
        }
      } catch (err) {
        console.warn('Error deleting rows from ' + sheetConfig.name + ': ' + err.message);
      }
    });

    // 4. Invalidate Script Cache
    CacheHelper.invalidateAll();

    // 5. If the deleted year was the current active year, switch to another available year
    let nextActiveYear = Number(settings.currentYear);
    if (nextActiveYear === year) {
      nextActiveYear = 0;
      try {
        const remainingYearsValues = yearsSheet.getDataRange().getDisplayValues();
        const yearColIdx = remainingYearsValues[0].map(normalizeHeader_).indexOf('year');
        if (yearColIdx !== -1 && remainingYearsValues.length > 1) {
          // Select the first remaining year in the list
          nextActiveYear = Number(remainingYearsValues[1][yearColIdx]);
        }
      } catch (e) {
        console.warn('Error finding next active year: ' + e.message);
      }

      if (nextActiveYear && !isNaN(nextActiveYear)) {
        ConfigService.saveSettings({ currentYear: nextActiveYear });
      } else {
        // No remaining years, reset currentYear settings or fallback to DEFAULT_YEAR
        ConfigService.saveSettings({ currentYear: '' });
      }
    }

    return {
      success: true,
      year: year,
      nextActiveYear: nextActiveYear || null,
      trashReport: trashReport
    };
  },

  setTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.fileId, 'File template');
    return ConfigRepository.setTemplateCategory(payload.fileId, payload.categoryId || '');
  }
};

function inferSubActivityMappingStatus_(sub) {
  if (!sub) return 'PERLU_REVIEW';
  const folderName = String(sub.sub_activity_name || '').trim();
  const formalName = String(sub.formal_archive_name || '').trim();
  const sheetName = String(sub.target_sheet_name || '').trim();
  if (!formalName || !sheetName) return 'PERLU_REVIEW';
  if (normalizeSettingsMappingText_(formalName) === normalizeSettingsMappingText_(folderName)) return 'PERLU_REVIEW';
  if (normalizeSettingsMappingText_(sheetName) === normalizeSettingsMappingText_(folderName)) return 'PERLU_REVIEW';
  return 'AUTO_MATCHED';
}

function normalizeSettingsMappingText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/^\d+\s*[\.\-\)]\s*/, '')
    .replace(/\btahun\s+20\d{2}\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
