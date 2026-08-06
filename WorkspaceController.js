'use strict';

const WorkspaceController = {
  getBootstrap: function () { return SettingsController.getBootstrap(); },
  getClientBootstrap: function (payload) {
    const full = SettingsController.getBootstrap();
    const user = AuthService.getCurrentUser(payload || {});
    return user && user.role !== 'guest' ? full : redactBootstrapForGuest_(full);
  },
  getSystemVersion: function () { return getVersion(); },

  saveSettings: function (payload) {
    const actor = requireAdminIfWorkspaceSecured_(payload);
    const r = SettingsController.saveSettings(payload);
    bumpVersion();
    auditAction_(actor, 'SETTINGS_SAVED', { message: 'Mengubah pengaturan ruang kerja' });
    return r;
  },
  initializeWorkspace: function (payload) {
    SystemLogger.info('WORKSPACE_INIT', 'Initializing workspace', { year: payload.year });
    payload = payload || {};
    const actor = requireAdminIfWorkspaceSecured_(payload);
    const r = SettingsController.initializeWorkspace(payload);
    bumpVersion();
    auditAction_(actor, 'WORKSPACE_INIT', { year: payload.year, message: 'Inisialisasi ruang kerja tahun ' + (payload.year || '-') });
    return r;
  },
  deleteYear: function (payload) {
    SystemLogger.warn('WORKSPACE_DELETE_YEAR', 'Deleting year config', { year: payload.year });
    payload = payload || {};
    const actor = requireAdmin_(payload);
    const r = SettingsController.deleteYear(payload);
    bumpVersion();
    auditAction_(actor, 'YEAR_DELETED', { year: payload.year, message: 'Menghapus tahun kerja ' + (payload.year || '-') });
    return r;
  },
  installMaintenanceTrigger: function (payload) {
    const actor = requireAdmin_(payload);
    const r = SettingsController.ensureArchiveMaintenanceTrigger();
    try { r.docTypesTrigger = SettingsController.ensureDocumentTypesSyncTrigger(); } catch (e) { r.docTypesTrigger = { installed: false, error: e.message }; }
    auditAction_(actor, 'TRIGGER_INSTALLED', { message: 'Memasang trigger maintenance harian + sinkron tipe dokumen' });
    return r;
  },
  updateSubActivityMapping: function (payload) {
    payload = payload || {};
    const actor = requireAdmin_(payload);
    const r = SettingsController.updateSubActivityMapping(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_MAPPING_UPDATED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, folderId: payload.folderId, message: 'Memperbarui mapping sub-kegiatan' });
    return r;
  },
  resetWorkspace: function (payload) {
    SystemLogger.error('WORKSPACE_RESET', 'Workspace completely reset by admin', {});
    const actor = requireAdmin_(payload);
    const r = SettingsController.resetWorkspace();
    bumpVersion();
    auditAction_(actor, 'WORKSPACE_RESET', { message: 'Mereset ruang kerja' + (r && typeof r.removed === 'number' ? ' (' + r.removed + ' file)' : '') });
    return r;
  }
};

function redactBootstrapForGuest_(full) {
  full = full || {};
  const currentYear = full.settings && full.settings.currentYear
    ? full.settings.currentYear
    : (full.selectedYear || DEFAULT_YEAR);
  const safeActivities = (full.activities || []).map(function (activity) {
    return {
      activity_id: activity.activity_id || '',
      activity_name: activity.activity_name || '',
      laci_no: activity.laci_no || '',
      folder_no: activity.folder_no || '',
      sort_order: activity.sort_order || '',
      laciFolder: null,
      targetFolder: null,
      spreadsheetFile: {},
      fields: [],
      subActivities: (activity.subActivities || []).map(function (sub) {
        return {
          sub_activity_id: sub.sub_activity_id || '',
          activity_id: sub.activity_id || activity.activity_id || '',
          sub_activity_name: sub.sub_activity_name || '',
          parent_folder_name: sub.parent_folder_name || '',
          sort_order: sub.sort_order || '',
          folder_id: '',
          folder: null,
          spreadsheetFile: {},
          effective_formal_archive_name: sub.effective_formal_archive_name || '',
          effective_target_sheet_name: sub.effective_target_sheet_name || ''
        };
      })
    };
  });
  return {
    configured: !!full.configured,
    settings: { currentYear: currentYear },
    selectedYear: full.selectedYear || currentYear,
    years: (full.years || []).map(function (yearRow) { return { year: yearRow.year }; }),
    activities: safeActivities,
    documentTypes: full.documentTypes || [],
    history: [],
    historyMeta: { total: 0, page: 1, totalPages: 0 },
    progress: full.progress || { total: 0, completed: 0, failed: 0, draft: 0, byActivity: {} },
    maintenance: {},
    message: full.message || ''
  };
}

