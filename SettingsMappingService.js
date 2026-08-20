'use strict';

/** @private Progress, mapping, and Drive picker operations. */
const SettingsMappingImpl_ = {
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
      sortOrder: payload.sortOrder,
      defaultKodeKlasifikasi: payload.defaultKodeKlasifikasi,
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
    SpreadsheetService.reconcileGlobalArchiveNumbers(year, { apply: true });
    CacheHelper.invalidate(year);
    return { subActivity: updated, bootstrap: this.getBootstrap() };
  },

  batchUpdateSubActivityMappings: function (payload) {
    payload = payload || {};
    const items = payload.items || [];
    if (!items.length) return { success: true, bootstrap: this.getBootstrap() };
    const year = Validator.requireYear(payload.year || items[0].year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const activity = ConfigService.findActivity(config, item.activityId);
        if (!activity) continue;
        const previous = ConfigService.findSubActivity(config, item.activityId, item.subActivityId);
        if (!previous) continue;

        if (item.targetSheetName && previous.target_sheet_name && item.targetSheetName !== previous.target_sheet_name) {
          SpreadsheetService.renameSubActivitySheet(activity, previous.target_sheet_name, item.targetSheetName, previous);
        }
        if (item.sortOrder !== undefined) previous.sort_order = item.sortOrder;
        if (item.formalArchiveName !== undefined) previous.formal_archive_name = item.formalArchiveName;
        if (item.targetSheetName !== undefined) previous.target_sheet_name = item.targetSheetName;
        if (item.defaultKodeKlasifikasi !== undefined) {
          previous.default_kode_klasifikasi = item.defaultKodeKlasifikasi;
        } else if (!isPelatihanActivity_(item.activityId, previous.sub_activity_name, previous.formal_archive_name) && String(previous.default_kode_klasifikasi || '').trim() === 'PDP.07.1') {
          previous.default_kode_klasifikasi = '';
          item.defaultKodeKlasifikasi = '';
        }
        if (item.noFolder !== undefined) previous.no_folder = item.noFolder;
      }

      ConfigRepository.updateSubActivityMappingsBatch(items);
      SpreadsheetService.reconcileGlobalArchiveNumbers(year, { apply: true, quick: true });
      CacheHelper.invalidate(year);
      const freshConfig = CacheHelper.getConfig(year);
      return { success: true, bootstrap: this.getBootstrap(year, freshConfig) };
    }, 30000);
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

};
