'use strict';

/**
 * Entrypoint: serve SPA shell.
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal Arsip PUSJARSKPP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Load partial HTML file for <?!= include() ?>.
 * @param {string} filename
 * @return {string} Raw HTML content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Strip internal infrastructure details from error messages.
 * Prevents leaking Drive IDs, spreadsheet IDs, stack traces, etc.
 * @param {string} message
 * @return {string}
 */
function sanitizeError_(message) {
  if (!message) return 'Terjadi kesalahan tidak dikenal.';
  let msg = String(message);
  // Strip Google Drive / Spreadsheet IDs (25+ char base64url-like strings).
  // Ambang disamakan dengan cleanId_ ({25,}) agar ID 25-32 char tidak lolos.
  msg = msg.replace(/\b[A-Za-z0-9_-]{25,}\b/g, '[ID]');
  // Strip file paths
  msg = msg.replace(/(?:\/|[A-Z]:\\)[^\s"']+/g, '[path]');
  // Strip URLs
  msg = msg.replace(/https?:\/\/[^\s"']+/g, '[URL]');
  // Strip stack traces
  msg = msg.replace(/\n\s*at\s+.+/g, '');
  if (msg.length > 500) msg = msg.slice(0, 500) + '...';
  return msg;
}

/**
 * Determine error code from error message content.
 * @param {string} msg
 * @return {string}
 */
function getErrorCode_(msg) {
  const m = String(msg || '').toLowerCase();
  if (m.indexOf('sesi login') >= 0 || m.indexOf('akses ditolak') >= 0 || m.indexOf('hanya admin') >= 0) return 'AUTH_ERROR';
  if (m.indexOf('terlalu banyak percobaan') >= 0) return 'RATE_LIMIT_ERROR';
  if (m.indexOf('wajib diisi') >= 0 || m.indexOf('tidak valid') >= 0 || m.indexOf('harus berupa teks') >= 0) return 'VALIDATION_ERROR';
  if (m.indexOf('tidak ditemukan') >= 0) return 'NOT_FOUND_ERROR';
  if (m.indexOf('sedang sibuk') >= 0 || m.indexOf('lock') >= 0) return 'LOCK_ERROR';
  if (m.indexOf('limit') >= 0 || m.indexOf('quota') >= 0 || m.indexOf('timeout') >= 0) return 'QUOTA_ERROR';
  return 'INTERNAL_ERROR';
}

/**
 * Wrap a GAS-callable action in {success, data|error, errorCode}.
 * Every public API function uses this for uniform error handling.
 * @param {function(): *} action
 * @return {{success: boolean, data?: *, error?: string, errorCode?: string}}
 */
function wrapApi(action) {
  try {
    return { success: true, data: action() };
  } catch (error) {
    console.error(error.message);
    const sanitizedMsg = sanitizeError_(error.message);
    return { 
      success: false, 
      error: sanitizedMsg,
      errorCode: getErrorCode_(error.message)
    };
  }
}

function getBootstrap() {
  return wrapApi(() => WorkspaceController.getBootstrap());
}

/**
 * Lazy-load aset dekoratif berat (background gedung & ornamen) di luar payload
 * HTML awal. Dipanggil client hanya saat toggle background aktif.
 * @return {{success: boolean, data?: {gedungUtama: string, ornamen: string}, error?: string}}
 */
function getDecorativeAssets() {
  return wrapApi(() => JSON.parse(HtmlService.createHtmlOutputFromFile('ClientAssetsHeavy').getContent()));
}

function saveSettings(payload) {
  return wrapApi(() => WorkspaceController.saveSettings(payload));
}

function installMaintenanceTrigger(payload) {
  return wrapApi(() => WorkspaceController.installMaintenanceTrigger(payload || {}));
}

function getArchiveMetadataDefaults(payload) {
  return wrapApi(() => ArchiveController.getArchiveMetadataDefaults(payload || {}));
}

function getArchiveMetadata(payload) {
  return wrapApi(() => ArchiveController.getArchiveMetadata(payload || {}));
}

function saveDraftToLog(payload) {
  return wrapApi(() => ArchiveController.saveDraftToLog(payload));
}

function deleteDraft(payload) {
  return wrapApi(() => ArchiveController.deleteDraft(payload));
}

function finalizeArchive(payload) {
  return wrapApi(() => ArchiveController.finalizeArchive(payload));
}

function validateArchiveFields(payload) {
  return wrapApi(() => ArchiveController.validateArchiveFields(payload || {}));
}

function editMetadata(payload) {
  return wrapApi(() => ArchiveController.editMetadata(payload || {}));
}

function adoptExistingArchives(payload) {
  return wrapApi(() => ArchiveController.adoptExistingArchives(payload || {}));
}

function previewExistingArchives(payload) {
  payload = payload || {};
  payload.dryRun = true;
  return wrapApi(() => ArchiveController.adoptExistingArchives(payload));
}

function initInboxResumableUpload(payload) {
  return wrapApi(() => ArchiveController.initInboxResumableUpload(payload || {}));
}

function initTemplateResumableUpload(payload) {
  return wrapApi(() => {
    requireAuth_(payload || {});
    return DriveService.initTemplateResumableUpload(payload || {});
  });
}

function uploadResumableChunk(payload) {
  return wrapApi(() => {
    requireAuth_(payload || {});
    return DriveService.uploadResumableChunk(payload || {});
  });
}

function uploadSourceFile(payload) {
  return wrapApi(() => ArchiveController.uploadSourceFile(payload));
}

function parseDocumentContent(payload) {
  return wrapApi(() => ArchiveController.parseDocumentContent(payload));
}

function addSubActivity(payload) {
  return wrapApi(() => SubActivityController.addSubActivity(payload));
}

function syncSubActivities(payload) {
  return wrapApi(() => SubActivityController.syncSubActivities(payload));
}

function getHistory(payload) {
  return wrapApi(() => AccountController.getHistory(payload || {}));
}

function uploadTemplate(payload) {
  return wrapApi(() => TemplateController.uploadTemplate(payload || {}));
}

function deleteTemplate(payload) {
  return wrapApi(() => TemplateController.deleteTemplate(payload || {}));
}

function saveTemplateCategory(payload) {
  return wrapApi(() => TemplateController.saveTemplateCategory(payload || {}));
}

function renameTemplateCategory(payload) {
  return wrapApi(() => TemplateController.renameTemplateCategory(payload || {}));
}

function deleteTemplateCategory(payload) {
  return wrapApi(() => TemplateController.deleteTemplateCategory(payload || {}));
}

function setTemplateCategory(payload) {
  return wrapApi(() => TemplateController.setTemplateCategory(payload || {}));
}

function getTemplatesData(payload) {
  return wrapApi(() => TemplateController.getTemplatesData(payload || {}));
}

function initializeWorkspace(payload) {
  return wrapApi(() => WorkspaceController.initializeWorkspace(payload || {}));
}

function deleteYear(payload) {
  return wrapApi(() => WorkspaceController.deleteYear(payload || {}));
}

function updateSubActivityMapping(payload) {
  return wrapApi(() => WorkspaceController.updateSubActivityMapping(payload || {}));
}

function renameDriveItem(payload) {
  return wrapApi(() => DriveController.renameDriveItem(payload || {}));
}

function listDriveFolders(payload) {
  return wrapApi(() => DriveController.listDriveFolders(payload || {}));
}

function listArchiveFolder(payload) {
  return wrapApi(() => DriveController.listArchiveFolder(payload || {}));
}

function addArchiveChildFolder(payload) {
  return wrapApi(() => DriveController.addArchiveChildFolder(payload || {}));
}

function bulkAddArchiveDocumentLinks(payload) {
  return wrapApi(() => ArchiveController.bulkAddArchiveDocumentLinks(payload || {}));
}

function addArchiveDocumentLink(payload) {
  return wrapApi(() => ArchiveController.addArchiveDocumentLink(payload || {}));
}

function getShortcutTargetInfo(payload) {
  return wrapApi(() => ArchiveController.getShortcutTargetInfo(payload || {}));
}

function updateArchiveDocumentLink(payload) {
  return wrapApi(() => ArchiveController.updateArchiveDocumentLink(payload || {}));
}

function renameArchiveFolder(payload) {
  return wrapApi(() => DriveController.renameArchiveFolder(payload || {}));
}

function trashArchiveFolder(payload) {
  return wrapApi(() => DriveController.trashArchiveFolder(payload || {}));
}

function renameArchiveFile(payload) {
  return wrapApi(() => DriveController.renameArchiveFile(payload || {}));
}

function trashArchiveFile(payload) {
  return wrapApi(() => DriveController.trashArchiveFile(payload || {}));
}

function deleteSubActivity(payload) {
  return wrapApi(() => SubActivityController.deleteSubActivity(payload || {}));
}

function trashSubActivityFolder(payload) {
  return wrapApi(() => SubActivityController.trashSubActivityFolder(payload || {}));
}

function cleanupTrashedSubActivities(payload) {
  // Guard di sini (endpoint user), BUKAN di controller — sebab runArchiveMaintenance
  // (time-trigger, konteks owner tanpa sesi) memanggil controller-nya langsung.
  return wrapApi(() => {
    requireAdmin_(payload || {});
    return SubActivityController.cleanupTrashedSubActivities(payload || {});
  });
}

/**
 * Handler maintenance harian yang dipanggil time-trigger (konteks owner, TANPA
 * sesi login). Sengaja TIDAK requireAuth_ dan terpisah dari endpoint
 * cleanupTrashedSubActivities (yang dipanggil user dan butuh sesi).
 */
function runArchiveMaintenance() {
  try {
    SubActivityController.cleanupTrashedSubActivities({});
    auditAction_({ displayName: 'Sistem' }, 'MAINTENANCE_CLEANUP', { message: 'Pembersihan harian sub-kegiatan terhapus (trigger)' });
  } catch (error) {
    console.error('runArchiveMaintenance: cleanup gagal: ' + error.message);
  }
}

/**
 * Handler onEdit (installable) pada config spreadsheet. Saat admin menambah/
 * mengubah baris di sheet `config_document_types`, kolom dokumen di SEMUA
 * spreadsheet Rekap langsung disinkronkan. Konteks trigger (owner, TANPA sesi).
 */
function onConfigDocumentTypesEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== CONFIG_SHEETS.DOCUMENT_TYPES) return;

    // Debounce: rentetan edit (mis. isi banyak sel) cukup memicu satu sweep.
    const cache = CacheService.getScriptCache();
    if (cache.get('doctypes_sync_lock')) return;
    cache.put('doctypes_sync_lock', '1', 8);

    const year = ConfigService.getSettings().currentYear || DEFAULT_YEAR;
    const r = SettingsController.syncDocumentTypeColumns(year);
    auditAction_({ displayName: 'Sistem' }, 'DOCTYPES_SYNCED', { year: r.year, message: 'Sinkron otomatis tipe dokumen (onEdit): ' + r.spreadsheetsSynced + ' spreadsheet' });
  } catch (error) {
    console.error('onConfigDocumentTypesEdit gagal: ' + error.message);
  }
}

/**
 * Handler onEdit (installable) pada spreadsheet arsip (Activity Spreadsheet).
 * Sinkronisasi 2 arah: Jika admin/staff mengubah "Kode Klasifikasi" di sheet Rekap,
 * perubahan akan disimpan kembali ke config aplikasi.
 */
function onRekapSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== REKAP_SHEET_NAME) return;

    // Kolom Kode Klasifikasi di Rekap adalah kolom ke-3 (C)
    const col = e.range.getColumn();
    if (col !== 3) return;

    // Data Rekap dimulai dari baris 8
    const row = e.range.getRow();
    if (row < 8) return;

    const newValue = String(e.value || e.range.getValue() || '').trim();
    if (!newValue) return;

    // Uraian ada di kolom 4 (D)
    const uraian = String(sheet.getRange(row, 4).getValue() || '').trim();
    if (!uraian) return;

    const year = ConfigService.getSettings().currentYear || new Date().getFullYear();
    const activities = ConfigRepository.getActivities(year);
    
    let matchedSubActivity = null;
    let matchedActivityId = null;

    // Cari sub-kegiatan berdasarkan Uraian (sub_activity_name atau formal_archive_name)
    for (let i = 0; i < activities.length; i++) {
      const act = activities[i];
      const subActivities = ConfigRepository.getSubActivities(year, act.activity_id);
      for (let j = 0; j < subActivities.length; j++) {
        const sub = subActivities[j];
        const formalName = (sub.formal_archive_name || sub.sub_activity_name || '').trim();
        if (formalName.toLowerCase() === uraian.toLowerCase() || (sub.sub_activity_name || '').trim().toLowerCase() === uraian.toLowerCase()) {
          matchedSubActivity = sub;
          matchedActivityId = act.activity_id;
          break;
        }
      }
      if (matchedSubActivity) break;
    }

    if (matchedSubActivity) {
      if (matchedSubActivity.default_kode_klasifikasi !== newValue) {
        matchedSubActivity.default_kode_klasifikasi = newValue;
        ConfigRepository.updateSubActivity(year, matchedActivityId, matchedSubActivity.sub_activity_id, matchedSubActivity);
      }
    }
  } catch (error) {
    console.error('onRekapSheetEdit gagal: ' + error.message);
  }
}

function getInactiveSubActivities(payload) {
  return wrapApi(() => SubActivityController.getInactiveSubActivities(payload || {}));
}

function restoreSubActivity(payload) {
  return wrapApi(() => SubActivityController.restoreSubActivity(payload || {}));
}

function purgeSubActivity(payload) {
  return wrapApi(() => SubActivityController.purgeSubActivity(payload || {}));
}

function getSystemVersion() {
  return wrapApi(() => WorkspaceController.getSystemVersion());
}

function getAdminAuditLogs(payload) {
  return wrapApi(() => AccountController.getAdminAuditLogs(payload || {}));
}

function login(payload) {
  return wrapApi(() => AccountController.login(payload || {}));
}

function logout(payload) {
  return wrapApi(() => AccountController.logout(payload || {}));
}

function getCurrentUser(payload) {
  return wrapApi(() => AccountController.getCurrentUser(payload || {}));
}

function resetWorkspace(payload) {
  return wrapApi(() => WorkspaceController.resetWorkspace(payload || {}));
}

function renameSubActivity(payload) {
  return wrapApi(() => SubActivityController.renameSubActivity(payload || {}));
}

function updateSubActivityMetadata(payload) {
  return wrapApi(() => SubActivityController.updateSubActivityMetadata(payload || {}));
}

function listAccounts(payload) {
  return wrapApi(() => AccountController.listAccounts(payload || {}));
}

function saveAccount(payload) {
  return wrapApi(() => AccountController.saveAccount(payload || {}));
}

function deleteAccount(payload) {
  return wrapApi(() => AccountController.deleteAccount(payload || {}));
}

function getArchiveLogByFileId(payload) {
  return wrapApi(() => ArchiveController.getArchiveLogByFileId(payload || {}));
}

function createParentFolder(payload) {
  return wrapApi(() => SubActivityController.createParentFolder(payload || {}));
}

function convertSubActivityToParent(payload) {
  return wrapApi(() => SubActivityController.convertSubActivityToParent(payload || {}));
}

function syncExistingPhysicalFiles(payload) {
  return wrapApi(() => ArchiveController.syncExistingPhysicalFiles(payload || {}));
}

