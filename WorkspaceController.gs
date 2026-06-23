'use strict';

const WorkspaceController = {
  getBootstrap: function () { return SettingsController.getBootstrap(); },
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
