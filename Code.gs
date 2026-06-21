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
  var msg = String(message);
  // Strip Google Drive / Spreadsheet IDs (33+ char base64url-like strings)
  msg = msg.replace(/\b[A-Za-z0-9_-]{33,}\b/g, '[ID]');
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
  return wrapApi(() => AppController.getBootstrap());
}

function getSettings() {
  return wrapApi(() => AppController.getSettings());
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
  return wrapApi(() => AppController.saveSettings(payload));
}

function installMaintenanceTrigger(payload) {
  return wrapApi(() => AppController.installMaintenanceTrigger(payload || {}));
}

function createDraft(payload) {
  return wrapApi(() => AppController.createDraft(payload));
}

function getArchiveMetadataDefaults(payload) {
  return wrapApi(() => AppController.getArchiveMetadataDefaults(payload || {}));
}

function getArchiveMetadata(payload) {
  return wrapApi(() => AppController.getArchiveMetadata(payload || {}));
}

function saveDraftToLog(payload) {
  return wrapApi(() => AppController.saveDraftToLog(payload));
}

function deleteDraft(payload) {
  return wrapApi(() => AppController.deleteDraft(payload));
}

function finalizeArchive(payload) {
  return wrapApi(() => AppController.finalizeArchive(payload));
}

function deleteArchive(payload) {
  return wrapApi(() => AppController.deleteArchive(payload || {}));
}

function validateArchiveFields(payload) {
  return wrapApi(() => AppController.validateArchiveFields(payload || {}));
}

function editMetadata(payload) {
  return wrapApi(() => AppController.editMetadata(payload || {}));
}

function adoptExistingArchives(payload) {
  return wrapApi(() => AppController.adoptExistingArchives(payload || {}));
}

function previewExistingArchives(payload) {
  payload = payload || {};
  payload.dryRun = true;
  return wrapApi(() => AppController.adoptExistingArchives(payload));
}

function initInboxResumableUpload(payload) {
  return wrapApi(() => AppController.initInboxResumableUpload(payload || {}));
}

function initTemplateResumableUpload(payload) {
  return wrapApi(() => AppController.initTemplateResumableUpload(payload || {}));
}

function uploadResumableChunk(payload) {
  return wrapApi(() => AppController.uploadResumableChunk(payload || {}));
}

function uploadSourceFile(payload) {
  return wrapApi(() => AppController.uploadSourceFile(payload));
}

function parseDocumentContent(payload) {
  return wrapApi(() => AppController.parseDocumentContent(payload));
}

function listInboxFiles(payload) {
  return wrapApi(() => AppController.listInboxFiles(payload));
}

function addSubActivity(payload) {
  return wrapApi(() => AppController.addSubActivity(payload));
}

function syncSubActivities(payload) {
  return wrapApi(() => AppController.syncSubActivities(payload));
}

function getHistory(payload) {
  return wrapApi(() => AppController.getHistory(payload || {}));
}

function getTemplates(payload) {
  return wrapApi(() => AppController.getTemplates(payload || {}));
}

function uploadTemplate(payload) {
  return wrapApi(() => AppController.uploadTemplate(payload || {}));
}

function deleteTemplate(payload) {
  return wrapApi(() => AppController.deleteTemplate(payload || {}));
}

function getTemplateCategories(payload) {
  return wrapApi(() => AppController.getTemplateCategories(payload || {}));
}

function saveTemplateCategory(payload) {
  return wrapApi(() => AppController.saveTemplateCategory(payload || {}));
}

function renameTemplateCategory(payload) {
  return wrapApi(() => AppController.renameTemplateCategory(payload || {}));
}

function deleteTemplateCategory(payload) {
  return wrapApi(() => AppController.deleteTemplateCategory(payload || {}));
}

function setTemplateCategory(payload) {
  return wrapApi(() => AppController.setTemplateCategory(payload || {}));
}

function getTemplatesData(payload) {
  return wrapApi(() => AppController.getTemplatesData(payload || {}));
}

function initializeWorkspace(payload) {
  return wrapApi(() => AppController.initializeWorkspace(payload || {}));
}

function deleteYear(payload) {
  return wrapApi(() => AppController.deleteYear(payload || {}));
}

function updateActivityMapping(payload) {
  return wrapApi(() => AppController.updateActivityMapping(payload || {}));
}

function updateSubActivityMapping(payload) {
  return wrapApi(() => AppController.updateSubActivityMapping(payload || {}));
}

function renameDriveItem(payload) {
  return wrapApi(() => AppController.renameDriveItem(payload || {}));
}

function listDriveFolders(payload) {
  return wrapApi(() => AppController.listDriveFolders(payload || {}));
}

function listArchiveFolder(payload) {
  return wrapApi(() => AppController.listArchiveFolder(payload || {}));
}

function addArchiveChildFolder(payload) {
  return wrapApi(() => AppController.addArchiveChildFolder(payload || {}));
}

function bulkAddArchiveDocumentLinks(payload) {
  return wrapApi(() => AppController.bulkAddArchiveDocumentLinks(payload || {}));
}

function addArchiveDocumentLink(payload) {
  return wrapApi(() => AppController.addArchiveDocumentLink(payload || {}));
}

function getShortcutTargetInfo(payload) {
  return wrapApi(() => AppController.getShortcutTargetInfo(payload || {}));
}

function updateArchiveDocumentLink(payload) {
  return wrapApi(() => AppController.updateArchiveDocumentLink(payload || {}));
}

function renameArchiveFolder(payload) {
  return wrapApi(() => AppController.renameArchiveFolder(payload || {}));
}

function trashArchiveFolder(payload) {
  return wrapApi(() => AppController.trashArchiveFolder(payload || {}));
}

function renameArchiveFile(payload) {
  return wrapApi(() => AppController.renameArchiveFile(payload || {}));
}

function trashArchiveFile(payload) {
  return wrapApi(() => AppController.trashArchiveFile(payload || {}));
}

function deleteSubActivity(payload) {
  return wrapApi(() => AppController.deleteSubActivity(payload || {}));
}

function trashSubActivityFolder(payload) {
  return wrapApi(() => AppController.trashSubActivityFolder(payload || {}));
}

function cleanupTrashedSubActivities(payload) {
  return wrapApi(() => AppController.cleanupTrashedSubActivities(payload || {}));
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

/** Endpoint manual (admin) untuk paksa sinkron kolom tipe dokumen. */
function syncDocumentTypes(payload) {
  return wrapApi(() => AppController.syncDocumentTypes(payload || {}));
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

function getInactiveSubActivities(payload) {
  return wrapApi(() => AppController.getInactiveSubActivities(payload || {}));
}

function restoreSubActivity(payload) {
  return wrapApi(() => AppController.restoreSubActivity(payload || {}));
}

function purgeSubActivity(payload) {
  return wrapApi(() => AppController.purgeSubActivity(payload || {}));
}

function getSystemVersion() {
  return wrapApi(() => AppController.getSystemVersion());
}

function getAdminAuditLogs(payload) {
  return wrapApi(() => AppController.getAdminAuditLogs(payload || {}));
}

function login(payload) {
  return wrapApi(() => AppController.login(payload || {}));
}

function logout(payload) {
  return wrapApi(() => AppController.logout(payload || {}));
}

function getCurrentUser(payload) {
  return wrapApi(() => AppController.getCurrentUser(payload || {}));
}

function saveDefaultAdmin(payload) {
  return wrapApi(() => AppController.saveDefaultAdmin(payload || {}));
}

function resetWorkspace(payload) {
  return wrapApi(() => AppController.resetWorkspace(payload || {}));
}

function renameSubActivity(payload) {
  return wrapApi(() => AppController.renameSubActivity(payload || {}));
}

function updateSubActivityMetadata(payload) {
  return wrapApi(() => AppController.updateSubActivityMetadata(payload || {}));
}

function listAccounts(payload) {
  return wrapApi(() => AppController.listAccounts(payload || {}));
}

function saveAccount(payload) {
  return wrapApi(() => AppController.saveAccount(payload || {}));
}

function deleteAccount(payload) {
  return wrapApi(() => AppController.deleteAccount(payload || {}));
}

function getArchiveLogByFileId(payload) {
  return wrapApi(() => AppController.getArchiveLogByFileId(payload || {}));
}

function getFinalFileName(payload) {
  return wrapApi(() => AppController.getFinalFileName(payload || {}));
}

function createParentFolder(payload) {
  return wrapApi(() => AppController.createParentFolder(payload || {}));
}

function convertSubActivityToParent(payload) {
  return wrapApi(() => AppController.convertSubActivityToParent(payload || {}));
}

function forceResetAdmin() {
  return wrapApi(() => AppController.forceResetAdmin());
}

function syncExistingPhysicalFiles(payload) {
  return wrapApi(() => AppController.syncExistingPhysicalFiles(payload || {}));
}

function getUserEmail() {
  return wrapApi(() => AppController.getUserEmail());
}

