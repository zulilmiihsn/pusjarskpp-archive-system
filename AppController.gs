'use strict';

const AppController = {
  getBootstrap: function () { return SettingsController.getBootstrap(); },
  getSettings: function () { return SettingsController.getSettings(); },
  getSystemVersion: function () { return getVersion(); },

  saveSettings: function (payload) {
    const actor = requireAdminIfWorkspaceSecured_(payload);
    const r = SettingsController.saveSettings(payload);
    bumpVersion();
    auditAction_(actor, 'SETTINGS_SAVED', { message: 'Mengubah pengaturan ruang kerja' });
    return r;
  },
  initializeWorkspace: function (payload) {
    payload = payload || {};
    const actor = requireAdminIfWorkspaceSecured_(payload);
    const r = SettingsController.initializeWorkspace(payload);
    bumpVersion();
    auditAction_(actor, 'WORKSPACE_INIT', { year: payload.year, message: 'Inisialisasi ruang kerja tahun ' + (payload.year || '-') });
    return r;
  },
  deleteYear: function (payload) {
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
    auditAction_(actor, 'TRIGGER_INSTALLED', { message: 'Memasang trigger maintenance harian' });
    return r;
  },
  updateActivityMapping: function (payload) {
    payload = payload || {};
    const actor = requireAdmin_(payload);
    const r = SettingsController.updateActivityMapping(payload);
    bumpVersion();
    auditAction_(actor, 'ACTIVITY_MAPPING_UPDATED', { year: payload.year, activityId: payload.activityId, folderId: payload.targetFolderId, message: 'Memperbarui mapping kegiatan' });
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
  renameDriveItem: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = SettingsController.renameDriveItem(payload);
    bumpVersion();
    auditAction_(actor, 'DRIVE_ITEM_RENAMED', { folderId: payload.itemId, message: 'Mengganti nama item Drive: ' + (payload.name || '-') });
    return r;
  },
  listDriveFolders: function (payload) { return SettingsController.listDriveFolders(payload); },
  listArchiveFolder: function (payload) {
    payload = payload || {};
    Validator.requireId(payload.folderId, 'Folder ID');
    const result = DriveService.listFolderContent(payload);
    try {
      const mirrorFolder = resolveMirrorForFolder_(payload.folderId);
      if (mirrorFolder) {
        const mirrorPayload = {
          folderId: mirrorFolder.getId(),
          pageSize: payload.pageSize,
          pageToken: payload.pageToken
        };
        const mirrorContent = DriveService.listFolderContent(mirrorPayload);
        const shortcutFiles = (mirrorContent.files || []).filter(function (f) {
          return f.mimeType === 'application/vnd.google-apps.shortcut';
        });
        result.files = (result.files || []).concat(shortcutFiles);
      }
    } catch (e) {
      console.error('Failed to load mirror dokumen: ' + e.message);
    }
    return result;
  },
  addArchiveChildFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.parentFolderId, 'Folder induk');
    Validator.requireString(payload.name, 'Nama folder');
    const r = DriveService.createChildFolder(payload.parentFolderId, payload.name);
    bumpVersion();
    auditAction_(actor, 'FOLDER_CREATED', { folderId: payload.parentFolderId, message: 'Membuat folder turunan: ' + payload.name });
    return r;
  },
  bulkAddArchiveDocumentLinks: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const items = payload.items || [];
    if (!items.length) return { successCount: 0, failCount: 0, errors: [] };
    
    let spreadsheet = null;
    let activity = null;
    let subActivity = null;
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    
    if (payload.activityId && payload.subActivityId) {
      const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      const config = CacheHelper.getConfig(year);
      activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');
      subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
      if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
    }

    const files = [];
    items.forEach(function(item) {
      try {
        const file = DriveService.addArchiveDocumentLink(item);
        if (activity && subActivity) {
          var linkFile = { name: file.name, url: item.url || file.url };
          spreadsheet = SpreadsheetService.updateArchiveDocumentLink(activity, subActivity, item.name, linkFile);
        }
        files.push(file);
        successCount++;
      } catch (e) {
        failCount++;
        errors.push(item.name + ': ' + e.message);
      }
    });

    bumpVersion();
    auditAction_(actor, 'DOCLINK_BULK_ADDED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      message: 'Tambah tautan dokumen massal: ' + successCount + ' sukses, ' + failCount + ' gagal'
    });
    return { successCount: successCount, failCount: failCount, errors: errors, files: files };
  },
  addArchiveDocumentLink: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.parentFolderId, 'Folder induk');
    Validator.requireString(payload.name, 'Kategori dokumen');
    Validator.requireString(payload.url, 'Link Google Drive');
    const file = DriveService.addArchiveDocumentLink(payload);
    let spreadsheet = null;
    if (payload.activityId && payload.subActivityId) {
      const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');
      const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
      if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
      var linkFile = { name: file.name, url: payload.url || file.url };
      spreadsheet = SpreadsheetService.updateArchiveDocumentLink(activity, subActivity, payload.name, linkFile);
    }
    bumpVersion();
    auditAction_(actor, 'DOCLINK_ADDED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      folderId: payload.parentFolderId, message: 'Menambah tautan dokumen: ' + payload.name
    });
    return Object.assign({}, file, { spreadsheet: spreadsheet });
  },
  getShortcutTargetInfo: function (payload) {
    return DriveService.getShortcutTargetInfo(payload || {});
  },
  updateArchiveDocumentLink: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    Validator.requireString(payload.url, 'URL Drive');
    const result = DriveService.updateArchiveDocumentLink(payload);
    if (payload.activityId && payload.subActivityId && payload.categoryName) {
      const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');
      const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
      if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
      var linkFile = { name: result.name, url: payload.url || result.url };
      SpreadsheetService.updateArchiveDocumentLink(activity, subActivity, payload.categoryName, linkFile);
    }
    bumpVersion();
    auditAction_(actor, 'DOCLINK_UPDATED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      folderId: payload.fileId, message: 'Memperbarui tautan dokumen: ' + (payload.categoryName || '-')
    });
    return result;
  },
  renameArchiveFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    Validator.requireString(payload.name, 'Nama baru');
    const r = DriveService.renameFolder(payload.folderId, payload.name);
    bumpVersion();
    auditAction_(actor, 'FOLDER_RENAMED', { folderId: payload.folderId, message: 'Mengganti nama folder menjadi: ' + payload.name });
    return r;
  },
  trashArchiveFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    const r = DriveService.trashFolder(payload.folderId);
    bumpVersion();
    auditAction_(actor, 'FOLDER_TRASHED', { folderId: payload.folderId, message: 'Memindahkan folder ke Tempat Sampah' });
    return r;
  },
  renameArchiveFile: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    Validator.requireString(payload.name, 'Nama baru');
    const r = DriveService.renameFile(payload.fileId, payload.name);
    bumpVersion();
    auditAction_(actor, 'FILE_RENAMED', { folderId: payload.fileId, message: 'Mengganti nama file menjadi: ' + payload.name });
    return r;
  },
  trashArchiveFile: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    const log = ConfigRepository.getArchiveLogByFileId(payload.fileId);
    if (log && log.archive_id) {
       ArchiveController.deleteArchive({ archiveId: log.archive_id, year: payload.year });
       auditAction_(actor, 'ARCHIVE_DELETED', {
         year: payload.year || log.year, activityId: log.activity_id, subActivityId: log.sub_activity_id,
         folderId: log.final_file_id || payload.fileId,
         message: 'Menghapus surat: ' + (log.final_file_name || '-')
       });
    } else {
       let fileName = '';
       try {
         const file = DriveApp.getFileById(cleanId_(payload.fileId));
         fileName = file.getName();
       } catch (e) {
         console.warn('Failed to read file name before trash: ' + e.message);
       }
       DriveService.trashFile(payload.fileId);
       if (payload.activityId && payload.subActivityId && fileName) {
         try {
           const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
           const config = CacheHelper.getConfig(year);
           const activity = ConfigService.findActivity(config, payload.activityId);
           const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
           if (activity && subActivity) {
             SpreadsheetService.clearArchiveDocumentLink(activity, subActivity, fileName);
           }
         } catch (e) {
           console.error('Failed to clear archive document link: ' + e.message);
         }
       }
       auditAction_(actor, 'FILE_TRASHED', {
         year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
         folderId: payload.fileId, message: 'Menghapus file: ' + (fileName || '-')
       });
    }
    bumpVersion();
    return { success: true };
  },
  getHistory: function (payload) { return SettingsController.getHistory(payload); },
  getTemplates: function (payload) { return SettingsController.getTemplates(payload); },
  getTemplatesData: function (payload) { return SettingsController.getTemplatesData(payload); },
  uploadTemplate: function (payload) { requireAuth_(payload); const r = DriveService.uploadTemplateFile(payload); invalidateTemplatesCache_(payload.year); bumpVersion(); return r; },
  deleteTemplate: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    invalidateTemplatesCache_(payload.year);
    bumpVersion(); return DriveService.trashTemplateFile(payload.fileId);
  },
  getTemplateCategories: function () { return SettingsController.getTemplateCategories(); },
  saveTemplateCategory: function (payload) { const r = SettingsController.saveTemplateCategory(payload); invalidateTemplatesCache_(payload && payload.year); bumpVersion(); return r; },
  deleteTemplateCategory: function (payload) { const r = SettingsController.deleteTemplateCategory(payload); invalidateTemplatesCache_(payload && payload.year); bumpVersion(); return r; },
  renameTemplateCategory: function (payload) { const r = SettingsController.renameTemplateCategory(payload); invalidateTemplatesCache_(payload && payload.year); bumpVersion(); return r; },
  setTemplateCategory: function (payload) { const r = SettingsController.setTemplateCategory(payload); invalidateTemplatesCache_(payload && payload.year); bumpVersion(); return r; },
  getAdminAuditLogs: function (payload) { requireAdmin_(payload); return SettingsController.getAdminAuditLogs(payload); },
  listAccounts: function (payload) { requireAdmin_(payload); return ConfigService.listAccounts().map(stripAccountSecrets_); },
  saveAccount: function (payload) {
    payload = payload || {};
    var adminUser = requireAdmin_(payload);
    Validator.requireString(payload.username, 'Username');
    const currentUser = adminUser;
    const isUpdate = !!payload.accountId;
    const result = ConfigService.saveAccount(payload);
    bumpVersion();
    ConfigRepository.appendAdminAudit({
      created_at: new Date().toISOString(),
      actor: currentUser.displayName || currentUser.username || '',
      action: isUpdate ? 'ACCOUNT_UPDATED' : 'ACCOUNT_CREATED',
      status: 'SUCCESS',
      message: (isUpdate ? 'Update akun: ' : 'Buat akun: ') + (payload.displayName || payload.username)
    });
    return result;
  },
  deleteAccount: function (payload) {
    payload = payload || {};
    var adminUser = requireAdmin_(payload);
    Validator.requireString(payload.accountId, 'Account ID');
    const currentUser = adminUser;
    const allAccounts = ConfigService.listAccounts();
    const target = allAccounts.find(function (a) { return a.account_id === payload.accountId; });
    const result = ConfigService.deleteAccount(payload.accountId) > 0;
    bumpVersion();
    if (result) {
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: currentUser.displayName || currentUser.username || '',
        action: 'ACCOUNT_DELETED',
        status: 'SUCCESS',
        message: 'Nonaktifkan akun: ' + (target ? (target.displayName || target.username) : payload.accountId)
      });
    }
    return result;
  },
  initInboxResumableUpload: function (payload) { return ArchiveController.initInboxResumableUpload(payload); },
  initTemplateResumableUpload: function (payload) { return DriveService.initTemplateResumableUpload(payload); },
  uploadResumableChunk: function (payload) { return DriveService.uploadResumableChunk(payload); },
  uploadSourceFile: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = ArchiveController.uploadSourceFile(payload);
    auditAction_(actor, 'SOURCE_UPLOADED', { message: 'Mengunggah file sumber: ' + (payload.name || (r && r.file && r.file.name) || '-') });
    return r;
  },
  parseDocumentContent: function (payload) { requireAuth_(payload); return ArchiveController.parseDocumentContent(payload); },
  createDraft: function (payload) { requireAuth_(payload); return ArchiveController.createDraft(payload); },
  getArchiveMetadataDefaults: function (payload) { return ArchiveController.getMetadataDefaults(payload); },
  saveDraftToLog: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = ArchiveController.saveDraftToLog(payload);
    auditAction_(actor, 'DRAFT_SAVED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Menyimpan draf arsip' });
    return r;
  },
  deleteDraft: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = ArchiveController.deleteDraft(payload);
    auditAction_(actor, 'DRAFT_DELETED', { year: payload.year, message: 'Menghapus draf arsip: ' + (payload.archiveId || '-') });
    return r;
  },
  deleteArchive: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    let logCtx = null;
    try { logCtx = ConfigRepository.getArchiveLog(payload.archiveId); } catch (e) { /* ignore */ }
    const r = ArchiveController.deleteArchive(payload);
    bumpVersion();
    auditAction_(actor, 'ARCHIVE_DELETED', {
      year: payload.year || (logCtx && logCtx.year),
      activityId: logCtx && logCtx.activity_id,
      subActivityId: logCtx && logCtx.sub_activity_id,
      folderId: logCtx && logCtx.final_file_id,
      message: 'Menghapus surat: ' + ((logCtx && logCtx.final_file_name) || payload.archiveId || '-')
    });
    return r;
  },
  finalizeArchive: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = ArchiveController.finalizeArchive(payload);
    bumpVersion();
    auditAction_(actor, 'ARCHIVE_CREATED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      folderId: (r && r.finalFile && r.finalFile.id) || '',
      message: 'Mengarsipkan surat: ' + ((r && r.finalFile && r.finalFile.name) || '-')
    });
    return r;
  },
  adoptExistingArchives: function (payload) {
    payload = payload || {};
    const r = ArchiveController.adoptExistingArchives(payload);
    if (!payload.dryRun) {
      bumpVersion();
      auditAction_(auditActor_(payload), 'ARCHIVE_ADOPTED', {
        year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
        message: 'Adopsi arsip lama' + (r && typeof r.imported === 'number' ? ': ' + r.imported + ' item' : '')
      });
    }
    return r;
  },
  getArchiveLogByFileId: function (payload) {
    payload = payload || {};
    Validator.requireId(payload.fileId, 'File ID');
    const log = ConfigRepository.getArchiveLogByFileId(payload.fileId);
    return log || null;
  },
  getArchiveMetadata: function (payload) { return ArchiveController.getArchiveMetadata(payload); },
  validateArchiveFields: function (payload) { return ArchiveController.validateArchiveFields(payload); },
  editMetadata: function (payload) { requireAuth_(payload); const r = ArchiveController.editMetadata(payload); bumpVersion(); return r; },
  listInboxFiles: function (payload) { return ArchiveController.listInboxFiles(payload); },
  addSubActivity: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = SubActivityController.addSubActivity(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_CREATED', {
      year: payload.year, activityId: payload.activityId,
      subActivityId: (r && r.newSubId) || '',
      folderId: (r && r.folder && r.folder.id) || '',
      message: 'Membuat sub-kegiatan: ' + (payload.subActivityName || (r && r.subActivity && r.subActivity.sub_activity_name) || '-')
    });
    return r;
  },
  createParentFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = SubActivityController.createParentFolder(payload);
    bumpVersion();
    auditAction_(actor, 'PARENT_FOLDER_CREATED', { year: payload.year, activityId: payload.activityId, message: 'Membuat folder induk: ' + (payload.folderName || '-') });
    return r;
  },
  convertSubActivityToParent: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const nama = auditSubName_(payload);
    const r = SubActivityController.convertSubActivityToParent(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_CONVERTED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Mengubah "' + (nama || '-') + '" jadi folder induk, sub baru: ' + (payload.newSubActivityName || '-') });
    return r;
  },
  syncSubActivities: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = SubActivityController.syncSubActivities(payload);
    bumpVersion();
    if (r && r.updated) auditAction_(actor, 'SUBACT_SYNCED', { year: payload.year, activityId: payload.activityId, message: 'Sinkronisasi sub-kegiatan dari folder Drive' });
    return r;
  },
  deleteSubActivity: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const nama = auditSubName_(payload);
    const r = SubActivityController.deleteSubActivity(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_DELETED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Menghapus sub-kegiatan: ' + (nama || '-') });
    return r;
  },
  trashSubActivityFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const nama = auditSubName_(payload);
    const r = SubActivityController.trashSubActivityFolder(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_TRASHED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, folderId: (r && r.folderId) || '', message: 'Menghapus sub-kegiatan (folder ke Sampah): ' + (nama || '-') });
    return r;
  },
  cleanupTrashedSubActivities: function (payload) {
    const actor = requireAuth_(payload);
    const r = SubActivityController.cleanupTrashedSubActivities(payload);
    auditAction_(actor, 'MAINTENANCE_CLEANUP', { message: 'Menjalankan pembersihan sub-kegiatan terhapus' });
    return r;
  },
  renameSubActivity: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const r = SubActivityController.renameSubActivity(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_RENAMED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Mengganti nama sub-kegiatan menjadi: ' + (payload.newName || '-') });
    return r;
  },
  updateSubActivityMetadata: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const nama = auditSubName_(payload);
    const r = SubActivityController.updateSubActivityMetadata(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_META_UPDATED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Memperbarui metadata sub-kegiatan: ' + (nama || '-') });
    return r;
  },
  getInactiveSubActivities: function (payload) { return SubActivityController.getInactiveSubActivities(payload); },
  restoreSubActivity: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const nama = auditSubName_(payload);
    const r = SubActivityController.restoreSubActivity(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_RESTORED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Memulihkan sub-kegiatan' + (nama ? ': ' + nama : '') });
    return r;
  },
  purgeSubActivity: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const nama = auditSubName_(payload);
    const r = SubActivityController.purgeSubActivity(payload);
    bumpVersion();
    auditAction_(actor, 'SUBACT_PURGED', { year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId, message: 'Menghapus permanen sub-kegiatan' + (nama ? ': ' + nama : ': ' + (payload.subActivityId || '-')) });
    return r;
  },
  login: function (payload) { return AuthService.login(payload); },
  logout: function (payload) { const r = AuthService.logout(payload); bumpVersion(); return r; },
  resetWorkspace: function (payload) {
    const actor = requireAdmin_(payload);
    const r = SettingsController.resetWorkspace();
    bumpVersion();
    auditAction_(actor, 'WORKSPACE_RESET', { message: 'Mereset ruang kerja' + (r && typeof r.removed === 'number' ? ' (' + r.removed + ' file)' : '') });
    return r;
  },
  getCurrentUser: function (payload) { return AuthService.getCurrentUser(payload); },
  saveDefaultAdmin: function (payload) { requireAdmin_(payload); const r = AuthService.saveDefaultAdmin(); bumpVersion(); return r; },
  getUserEmail: function () { return AuthService.getUserEmail(); },
  getFinalFileName: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.metadata, 'Metadata JSON');
    Validator.requireString(payload.sourceName, 'Nama file sumber');
    return MetadataService.buildFinalFileName(JSON.parse(payload.metadata), payload.sourceName);
  }
};

/**
 * Tulis satu baris audit. Tidak pernah melempar error — audit yang gagal tidak
 * boleh menggagalkan aksi utama.
 * @param {{displayName?:string,username?:string}|null} actor
 * @param {string} action
 * @param {{year?:*,activityId?:string,subActivityId?:string,folderId?:string,status?:string,message?:string}} [opts]
 */
function auditAction_(actor, action, opts) {
  opts = opts || {};
  try {
    ConfigRepository.appendAdminAudit({
      created_at: new Date().toISOString(),
      actor: (actor && (actor.displayName || actor.username)) || 'Sistem',
      action: action,
      year: opts.year || '',
      activity_id: opts.activityId || '',
      sub_activity_id: opts.subActivityId || '',
      folder_id: opts.folderId || '',
      status: opts.status || 'SUCCESS',
      message: opts.message || ''
    });
  } catch (error) {
    console.error('audit gagal (' + action + '): ' + error.message);
  }
}

/**
 * Ambil user aktif dari sesi pada payload (untuk endpoint tanpa requireAuth_).
 * @param {object} payload
 * @return {object|null}
 */
function auditActor_(payload) {
  try {
    const user = AuthService.getCurrentUser(payload);
    return (user && user.role !== 'guest') ? user : null;
  } catch (error) {
    return null;
  }
}

/**
 * Resolusi nama sub-kegiatan dari config — dipanggil SEBELUM aksi destruktif
 * agar nama masih bisa ditemukan untuk pesan audit.
 * @param {object} payload
 * @return {string}
 */
function auditSubName_(payload) {
  try {
    const year = (payload && payload.year) || ConfigService.getSettings().currentYear || DEFAULT_YEAR;
    const config = CacheHelper.getConfig(year);
    const sub = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
    return sub ? (sub.sub_activity_name || '') : '';
  } catch (error) {
    return '';
  }
}

/**
 * Strip sensitive fields (password hash) from an account row before returning to client.
 * Internal callers (hasActiveAdminAccount_, deleteAccount) use ConfigService.listAccounts
 * directly and never expose the raw row to the client.
 * @param {object} account
 * @return {object}
 */
function stripAccountSecrets_(account) {
  const safe = Object.assign({}, account);
  delete safe.password_hash;
  return safe;
}

function requireAdminIfWorkspaceSecured_(payload) {
  const settings = ConfigService.getSettings();
  if (!settings.configSpreadsheetId) return null;
  if (!hasActiveAdminAccount_()) return null;
  return requireAdmin_(payload);
}

function hasActiveAdminAccount_() {
  try {
    return ConfigService.listAccounts().some(function (account) {
      return String(account.role || '').trim().toLowerCase() === 'admin';
    });
  } catch (error) {
    console.warn('Admin account check failed: ' + error.message);
    return false;
  }
}
