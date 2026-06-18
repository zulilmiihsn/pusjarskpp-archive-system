'use strict';

const AppController = {
  getBootstrap: function () { return SettingsController.getBootstrap(); },
  getSettings: function () { return SettingsController.getSettings(); },
  getSystemVersion: function () { return getVersion(); },

  saveSettings: function (payload) { requireAdminIfWorkspaceSecured_(payload); const r = SettingsController.saveSettings(payload); bumpVersion(); return r; },
  initializeWorkspace: function (payload) { requireAdminIfWorkspaceSecured_(payload); const r = SettingsController.initializeWorkspace(payload); bumpVersion(); return r; },
  deleteYear: function (payload) { requireAdmin_(payload); const r = SettingsController.deleteYear(payload); bumpVersion(); return r; },
  installMaintenanceTrigger: function (payload) { requireAdmin_(payload); return SettingsController.ensureArchiveMaintenanceTrigger(); },
  updateActivityMapping: function (payload) { requireAdmin_(payload); const r = SettingsController.updateActivityMapping(payload); bumpVersion(); return r; },
  updateSubActivityMapping: function (payload) { requireAdmin_(payload); const r = SettingsController.updateSubActivityMapping(payload); bumpVersion(); return r; },
  renameDriveItem: function (payload) { requireAuth_(payload); const r = SettingsController.renameDriveItem(payload); bumpVersion(); return r; },
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
    requireAuth_(payload);
    Validator.requireId(payload.parentFolderId, 'Folder induk');
    Validator.requireString(payload.name, 'Nama folder');
    bumpVersion(); return DriveService.createChildFolder(payload.parentFolderId, payload.name);
  },
  bulkAddArchiveDocumentLinks: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
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
    return { successCount: successCount, failCount: failCount, errors: errors, files: files };
  },
  addArchiveDocumentLink: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
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
    return Object.assign({}, file, { spreadsheet: spreadsheet });
  },
  getShortcutTargetInfo: function (payload) {
    return DriveService.getShortcutTargetInfo(payload || {});
  },
  updateArchiveDocumentLink: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
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
    return result;
  },  
  renameArchiveFolder: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    Validator.requireString(payload.name, 'Nama baru');
    bumpVersion(); return DriveService.renameFolder(payload.folderId, payload.name);
  },
  trashArchiveFolder: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    bumpVersion(); return DriveService.trashFolder(payload.folderId);
  },
  renameArchiveFile: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    Validator.requireString(payload.name, 'Nama baru');
    bumpVersion(); return DriveService.renameFile(payload.fileId, payload.name);
  },
  trashArchiveFile: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    const log = ConfigRepository.getArchiveLogByFileId(payload.fileId);
    if (log && log.archive_id) {
       ArchiveController.deleteArchive({ archiveId: log.archive_id, year: payload.year });
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
  uploadSourceFile: function (payload) { requireAuth_(payload); return ArchiveController.uploadSourceFile(payload); },
  parseDocumentContent: function (payload) { requireAuth_(payload); return ArchiveController.parseDocumentContent(payload); },
  createDraft: function (payload) { requireAuth_(payload); return ArchiveController.createDraft(payload); },
  getArchiveMetadataDefaults: function (payload) { return ArchiveController.getMetadataDefaults(payload); },
  saveDraftToLog: function (payload) { requireAuth_(payload); return ArchiveController.saveDraftToLog(payload); },
  deleteDraft: function (payload) { requireAuth_(payload); return ArchiveController.deleteDraft(payload); },
  deleteArchive: function (payload) { requireAuth_(payload); const r = ArchiveController.deleteArchive(payload); bumpVersion(); return r; },
  finalizeArchive: function (payload) { requireAuth_(payload); const r = ArchiveController.finalizeArchive(payload); bumpVersion(); return r; },
  adoptExistingArchives: function (payload) {
    const r = ArchiveController.adoptExistingArchives(payload);
    if (!(payload && payload.dryRun)) bumpVersion();
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
  addSubActivity: function (payload) { requireAuth_(payload); const r = SubActivityController.addSubActivity(payload); bumpVersion(); return r; },
  createParentFolder: function (payload) { requireAuth_(payload); const r = SubActivityController.createParentFolder(payload); bumpVersion(); return r; },
  convertSubActivityToParent: function (payload) { requireAuth_(payload); const r = SubActivityController.convertSubActivityToParent(payload); bumpVersion(); return r; },
  syncSubActivities: function (payload) { requireAuth_(payload); const r = SubActivityController.syncSubActivities(payload); bumpVersion(); return r; },
  deleteSubActivity: function (payload) { requireAuth_(payload); const r = SubActivityController.deleteSubActivity(payload); bumpVersion(); return r; },
  trashSubActivityFolder: function (payload) { requireAuth_(payload); const r = SubActivityController.trashSubActivityFolder(payload); bumpVersion(); return r; },
  cleanupTrashedSubActivities: function (payload) { requireAuth_(payload); return SubActivityController.cleanupTrashedSubActivities(payload); },
  renameSubActivity: function (payload) { requireAuth_(payload); const r = SubActivityController.renameSubActivity(payload); bumpVersion(); return r; },
  updateSubActivityMetadata: function (payload) { requireAuth_(payload); const r = SubActivityController.updateSubActivityMetadata(payload); bumpVersion(); return r; },
  getInactiveSubActivities: function (payload) { return SubActivityController.getInactiveSubActivities(payload); },
  restoreSubActivity: function (payload) { requireAuth_(payload); const r = SubActivityController.restoreSubActivity(payload); bumpVersion(); return r; },
  purgeSubActivity: function (payload) { requireAuth_(payload); const r = SubActivityController.purgeSubActivity(payload); bumpVersion(); return r; },
  login: function (payload) { return AuthService.login(payload); },
  logout: function (payload) { const r = AuthService.logout(payload); bumpVersion(); return r; },
  resetWorkspace: function (payload) { requireAdmin_(payload); const r = SettingsController.resetWorkspace(); bumpVersion(); return r; },
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
