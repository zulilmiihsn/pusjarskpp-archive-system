'use strict';

/**
 * Dokumentasi otorisasi yang dapat diuji mesin. Controller tetap menjadi sumber
 * enforcement utama; policy ini mencegah endpoint baru lolos tanpa klasifikasi.
 */
const POLICY_GUEST_ENDPOINTS_ = [
  'getBootstrap', 'getDecorativeAssets', 'getTemplatesData', 'getSystemVersion',
  'login', 'logout', 'getCurrentUser'
];

const POLICY_AUTH_ENDPOINTS_ = [
  'getArchiveMetadataDefaults', 'getArchiveMetadata', 'saveDraftToLog',
  'deleteDraft', 'finalizeArchive', 'archiveStep_validate',
  'archiveStep_copyFile', 'archiveStep_writeAndLog', 'validateArchiveFields',
  'editMetadata', 'adoptExistingArchives', 'previewExistingArchives',
  'initInboxResumableUpload', 'initTemplateResumableUpload',
  'uploadResumableChunk', 'uploadSourceFile', 'parseDocumentContent',
  'addSubActivity', 'syncSubActivities', 'repairActivityMapping', 'getHistory', 'uploadTemplate',
  'deleteTemplate', 'saveTemplateCategory', 'renameTemplateCategory',
  'setTemplateCategory', 'listArchiveFolder', 'addArchiveChildFolder',
  'bulkAddArchiveDocumentLinks', 'addArchiveDocumentLink',
  'getShortcutTargetInfo', 'updateArchiveDocumentLink', 'renameArchiveFolder',
  'renameArchiveFile', 'trashArchiveFile', 'getInactiveSubActivities',
  'restoreSubActivity', 'renameSubActivity', 'updateSubActivityMetadata',
  'getArchiveLogByFileId', 'createParentFolder',
  'convertSubActivityToParent', 'syncExistingPhysicalFiles'
];

const POLICY_ADMIN_ENDPOINTS_ = [
  'installMaintenanceTrigger', 'deleteYear', 'updateSubActivityMapping',
  'renameDriveItem', 'trashArchiveFolder', 'deleteSubActivity',
  'trashSubActivityFolder', 'cleanupTrashedSubActivities', 'cleanupOrphanedLainLainSheets', 'cleanupAllOrphanedSheets',
  'deleteTemplateCategory', 'purgeSubActivity', 'getAdminAuditLogs',
  'resetWorkspace', 'listAccounts', 'saveAccount', 'deleteAccount'
];

// Boleh tanpa sesi hanya saat bootstrap benar-benar belum memiliki admin.
const POLICY_BOOTSTRAP_ADMIN_ENDPOINTS_ = [
  'saveSettings', 'initializeWorkspace', 'listDriveFolders'
];

const POLICY_SYSTEM_HANDLERS_ = [
  'runArchiveMaintenance', 'onConfigDocumentTypesEdit', 'onRekapSheetEdit'
];

const POLICY_MUTATING_ENDPOINTS_ = [
  'saveSettings', 'installMaintenanceTrigger', 'saveDraftToLog', 'deleteDraft',
  'finalizeArchive', 'archiveStep_copyFile', 'archiveStep_writeAndLog',
  'editMetadata', 'adoptExistingArchives', 'initInboxResumableUpload',
  'initTemplateResumableUpload', 'uploadResumableChunk', 'uploadSourceFile',
  'addSubActivity', 'syncSubActivities', 'uploadTemplate', 'deleteTemplate',
  'saveTemplateCategory', 'renameTemplateCategory', 'deleteTemplateCategory',
  'setTemplateCategory', 'initializeWorkspace', 'deleteYear',
  'updateSubActivityMapping', 'renameDriveItem', 'addArchiveChildFolder',
  'bulkAddArchiveDocumentLinks', 'addArchiveDocumentLink',
  'updateArchiveDocumentLink', 'renameArchiveFolder', 'trashArchiveFolder',
  'renameArchiveFile', 'trashArchiveFile', 'deleteSubActivity',
  'trashSubActivityFolder', 'cleanupTrashedSubActivities',
  'restoreSubActivity', 'purgeSubActivity', 'resetWorkspace',
  'renameSubActivity', 'updateSubActivityMetadata', 'saveAccount',
  'deleteAccount', 'createParentFolder', 'convertSubActivityToParent',
  'syncExistingPhysicalFiles'
];

const ENDPOINT_ACCESS_POLICY_ = (function () {
  const policy = {};
  const mutations = {};
  POLICY_MUTATING_ENDPOINTS_.forEach(function (name) { mutations[name] = true; });
  function add(names, mode, roles) {
    names.forEach(function (name) {
      if (policy[name]) throw new Error('Policy endpoint duplikat: ' + name);
      policy[name] = {
        mode: mode,
        roles: roles.slice(),
        mutates: !!mutations[name]
      };
    });
  }
  add(POLICY_GUEST_ENDPOINTS_, 'PUBLIC', ['guest', 'user', 'admin']);
  add(POLICY_AUTH_ENDPOINTS_, 'AUTHENTICATED', ['user', 'admin']);
  add(POLICY_ADMIN_ENDPOINTS_, 'ADMIN', ['admin']);
  add(POLICY_BOOTSTRAP_ADMIN_ENDPOINTS_, 'BOOTSTRAP_ADMIN', ['admin']);
  add(POLICY_SYSTEM_HANDLERS_, 'SYSTEM_TRIGGER', ['system']);
  return policy;
})();

function getEndpointPolicy_(endpoint) {
  return ENDPOINT_ACCESS_POLICY_[String(endpoint || '')] || null;
}

function isRoleAllowedByPolicy_(endpoint, role, workspaceSecured) {
  const policy = getEndpointPolicy_(endpoint);
  if (!policy) return false;
  if (policy.mode === 'BOOTSTRAP_ADMIN') {
    return workspaceSecured ? role === 'admin' : true;
  }
  return policy.roles.indexOf(String(role || 'guest').toLowerCase()) >= 0;
}

function validateAuthorizationPolicy_() {
  const groups = [
    POLICY_GUEST_ENDPOINTS_, POLICY_AUTH_ENDPOINTS_, POLICY_ADMIN_ENDPOINTS_,
    POLICY_BOOTSTRAP_ADMIN_ENDPOINTS_, POLICY_SYSTEM_HANDLERS_
  ];
  const seen = {};
  const duplicates = [];
  groups.forEach(function (group) {
    group.forEach(function (name) {
      if (seen[name]) duplicates.push(name);
      seen[name] = true;
    });
  });
  const unknownMutations = POLICY_MUTATING_ENDPOINTS_.filter(function (name) {
    return !seen[name];
  });
  const invalidEntries = Object.keys(ENDPOINT_ACCESS_POLICY_).filter(function (name) {
    const p = ENDPOINT_ACCESS_POLICY_[name];
    return !p || !p.mode || !Array.isArray(p.roles) || !p.roles.length ||
      typeof p.mutates !== 'boolean';
  });
  return {
    ok: duplicates.length === 0 && unknownMutations.length === 0 && invalidEntries.length === 0,
    endpointCount: Object.keys(ENDPOINT_ACCESS_POLICY_).length,
    duplicates: duplicates,
    unknownMutations: unknownMutations,
    invalidEntries: invalidEntries
  };
}
