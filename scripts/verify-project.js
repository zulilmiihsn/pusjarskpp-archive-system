const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

function read(file) {
  return fs.readFileSync(path.join(root, file), 'utf8');
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const appController = read('WorkspaceController.gs') + '\n' + read('DriveController.gs') + '\n' + read('TemplateController.gs') + '\n' + read('AccountController.gs') + '\n' + read('SecurityHelpers.gs') + '\n' + read('ArchiveController.gs') + '\n' + read('SubActivityController.gs') + '\n' + read('SettingsController.gs') + '\n' + read('PureFunctions.gs');
const appsscript = JSON.parse(read('appsscript.json'));
const client = read('ClientState.html') + '\n' +
               read('ClientApi.html') + '\n' +
               read('ClientRouter.html') + '\n' +
               read('ClientDashboard.html') + '\n' +
               read('ClientActivityDetail.html') + '\n' +
               read('ClientArchiveFolder.html') + '\n' +
               read('ClientProcess.html') + '\n' +
               read('ClientDocumentProcess.html') + '\n' +
               read('ClientHistory.html') + '\n' +
               read('ClientTemplates.html') + '\n' +
               read('ClientSettings.html') + '\n' +
               read('ClientFolderPicker.html') + '\n' +
               read('ClientUtils.html');
const clientApi = read('ClientApi.html');
const clientRouter = read('ClientRouter.html');
const clientLogin = read('ClientLogin.html');
const authService = read('AuthService.gs');
['ClientState.html', 'ClientApi.html', 'ClientRouter.html', 'ClientLogin.html', 'ClientSettings.html'].forEach(function (file) {
  const source = read(file).replace(/^\s*<script>\s*/, '').replace(/\s*<\/script>\s*$/, '');
  new Function(source);
});
const configConstants = read('ConfigConstants.gs');
const configHelpers = read('ConfigHelpers.gs');
const configRepo = read('ConfigRepository.gs');
const configService = read('ConfigService.gs');
const config = configConstants + '\n' + configHelpers + '\n' + configRepo + '\n' + configService;
const driveService = read('DriveService.gs');
const readme = read('README.md');
const spreadsheetService = read('SpreadsheetService.gs') + '\n' + read('SheetHelpers.gs');
const workspaceSetup = read('WorkspaceSetup.gs');

assert(
  workspaceSetup.includes("sheet.getRange('B2:N2').merge().setValue('Daftar Isi Berkas Arsip Aktip')"),
  'Workspace detail template must use the B:N official detail layout.'
);
assert(!workspaceSetup.includes("sheet.getRange('C1:P1')"), 'Workspace detail template must not start at C:P.');
assert(configConstants.includes('DETAIL_NOTE_FALLBACK_ROW = 34'), 'Detail notes must start below the official 24-row table.');
assert(spreadsheetService.includes('function getDetailStartColumn_'), 'Spreadsheet writes must detect the detail start column.');
assert(configConstants.includes('DETAIL_FALLBACK_START_COL = 2'), 'Spreadsheet detail fallback column must be B.');
assert(!workspaceSetup.includes('hasRekapSheet: false'), 'Every laci workbook must include a rekap sheet.');
assert(workspaceSetup.includes("'error_message', 'metadata_json'"), 'Workspace setup must create archive_log with metadata_json.');
assert(
  workspaceSetup.includes('folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)'),
  'Archive folders must grant Anyone with link Viewer access.'
);
assert(
  /wsEnsureAnyoneWithLinkViewer_\(daftarArsip, report\)[\s\S]*wsEnsureAnyoneWithLinkViewer_\(naskahDinas, report\)/.test(workspaceSetup),
  'Both archive branches must receive link-viewer access.'
);
assert(
  !/wsEnsureAnyoneWithLinkViewer_\(systemFolder, report\)/.test(workspaceSetup),
  'System/config folder must remain private.'
);
assert(
  /deleteArchive: function \(payload\) \{[\s\S]*?requireArchiveDeletionRole_\(payload\)/.test(appController),
  'Archive deletion must allow authenticated portal workers through the dedicated role guard.'
);
assert(
  /trashArchiveFile: function \(payload\) \{[\s\S]*?requireArchiveDeletionRole_\(payload\)/.test(appController),
  'Archive item deletion must use the worker/admin role guard.'
);
const initWorkspaceBody = read('WorkspaceController.gs').match(/initializeWorkspace: function \(payload\) \{([\s\S]*?)\n {2}\},/);
assert(
  !!initWorkspaceBody && /requireAdminIfWorkspaceSecured_\(payload\)/.test(initWorkspaceBody[1]) && !/requireAdmin_\(payload\)/.test(initWorkspaceBody[1]),
  'First workspace initialization must not require an existing admin session.'
);
assert(read('SecurityHelpers.gs').includes('function requireAdminIfWorkspaceSecured_') && read('SecurityHelpers.gs').includes('hasActiveAdminAccount_'), 'Workspace auth guard must allow first-run or half-initialized recovery before an admin exists.');
assert(read('SettingsController.gs').includes('const adminResult = AuthService.saveDefaultAdmin();'), 'Workspace initialization must create the default admin through the internal auth service.');
assert(!read('SettingsController.gs').includes('AccountController.saveDefaultAdmin()'), 'Workspace initialization must not call the admin-protected saveDefaultAdmin endpoint.');

assert(appController.includes('LockService.getScriptLock()'), 'Archive finalization must use LockService.');
assert(appController.includes('SpreadsheetService.getNextItemNumber(activity, subActivity)'), 'Finalization must allocate item numbers inside the locked path.');
assert(appController.includes('STATUS.FAILED'), 'Finalization must write failed archive logs.');
assert(read('Code.gs').includes('function adoptExistingArchives'), 'Existing archive adoption must be exposed to Apps Script clients.');
assert(read('Code.gs').includes('function previewExistingArchives'), 'Existing archive adoption must support a preview endpoint.');
assert(appController.includes('adoptExistingArchives: function'), 'Existing archive adoption must have an application controller endpoint.');
assert(appController.includes('payload.dryRun'), 'Existing archive preview must avoid version bumping.');
assert(appController.includes('SpreadsheetService.listExistingArchiveRows'), 'Existing archive adoption must read filled detail rows.');
assert(configRepo.includes('getArchiveLogKeyMap'), 'Existing archive adoption must deduplicate previously imported rows.');
assert(spreadsheetService.includes('listExistingArchiveRows'), 'Spreadsheet service must expose a read-only existing row scan.');
assert(read('ClientSettings.html').includes('preview-existing-archives'), 'Settings UI must expose the existing data preview action.');
assert(read('ClientSettings.html').includes('Konfirmasi Adapt'), 'Settings UI must require confirmation before importing existing data.');

assert(driveService.includes('function getUniqueFileName_'), 'Drive copy must protect against duplicate final names.');
assert(appController.includes('trashSubActivityFolder'), 'There must be an explicit action for moving a Drive folder to Bin.');
assert(appController.includes('setTrashed(true)'), 'The explicit Bin action must move the Drive folder to Bin.');
assert(appController.includes("SUB_ACTIVITY_INACTIVE_REASON.MANUAL"), 'Manual deactivation must be tracked separately.');
assert(appController.includes("SUB_ACTIVITY_INACTIVE_REASON.DRIVE_TRASHED"), 'Drive Bin deactivation must be tracked separately.');
assert(appController.includes("inactiveReason === SUB_ACTIVITY_INACTIVE_REASON.DRIVE_MISSING"), 'Drive restore sync must only reactivate Drive-origin deactivations.');
assert(appController.includes('ensureArchiveMaintenanceTrigger'), 'Workspace setup must install the archive maintenance trigger.');
assert(appController.includes('cleanupTrashedSubActivities'), 'There must be a cleanup handler for trashed sub-activity config rows.');
assert(appController.includes('getInactiveSubActivities'), 'There must be an inactive sub-activity endpoint.');
assert(appController.includes('restoreSubActivity'), 'There must be a restore sub-activity endpoint.');
assert(appController.includes('purgeSubActivity'), 'There must be a purge sub-activity endpoint.');
assert(appController.includes('getAdminAuditLogs'), 'There must be an admin audit log endpoint.');
assert(!config.includes('deleteSubActivityRows'), 'Hard-delete config helper must not be present.');
assert(!workspaceSetup.includes('deleteSheet('), 'Workspace setup must not delete unknown config sheets.');
assert(config.includes('deactivateSubActivity'), 'Config must support soft-deactivating sub-activities.');
assert(config.includes('inactive_reason'), 'Sub-activity config must track inactive reason.');
assert(config.includes('inactive_at'), 'Sub-activity config must track inactive timestamp.');
assert(config.includes('purgeInactiveSubActivities'), 'Config must support retention cleanup for inactive rows.');
assert(config.includes('ADMIN_AUDIT_LOG'), 'Config must define an admin audit log sheet.');
assert(config.includes('readRecentSheetObjects_'), 'History reads must be bounded for Apps Script performance.');
assert(config.includes("'category_id', 'name', 'color', 'sort_order', 'created_at', 'updated_at'"), 'Template category config must persist colors.');
assert(read('SettingsController.gs').includes('normalizeHexColor_(meta.color'), 'Template data must merge stored category colors.');
assert(read('SettingsController.gs').includes('ConfigRepository.setTemplateCategory'), 'Template category assignment must persist to config.');
assert(read('SpreadsheetService.gs').includes('updateRekapSubActivityIdentity'), 'Sub-activity mapping changes must update the rekap identity row.');
assert(read('SpreadsheetService.gs').includes('reconcileGlobalArchiveNumbers'), 'Sync Drive must reconcile one global archive sequence across the year.');
assert(read('SpreadsheetService.gs').includes('const existingRow = findRekapRowForSubActivity_'), 'Rekap row creation must be idempotent.');
assert(read('SpreadsheetService.gs').includes('findPossibleRekapRowsFromLookup_') && read('SpreadsheetService.gs').includes('SpreadsheetService.updateRekapSummary(activity, sub, {})'), 'Sync Drive must safely recreate truly missing rekap rows and calculate them from detail sheets.');
assert(read('SheetHelpers.gs').includes('PORTAL_ARSIP_ROW_ID|'), 'Rekap rows must carry a stable sub-activity identity marker.');
assert(read('SheetHelpers.gs').includes('findRekapRowForSubActivityFromLookup_'), 'Rekap lookup must validate stale row hints against stable identity.');
assert(!read('SheetHelpers.gs').includes('uraian.indexOf(targetName)') &&
  !read('SheetHelpers.gs').includes('targetName.indexOf(uraian)'),
  'Rekap identity must never use fuzzy substring matching for numbered or Roman-numeral sub-activities.');
assert(!read('SheetHelpers.gs').includes('getDocumentProperties()') &&
  !read('SheetHelpers.gs').includes("setFormula('=1/2')"),
  'Formula separator detection must work in standalone web apps without Document Properties or cell probes.');
assert(read('PureFunctions.gs').includes('buildGlobalArchiveNumberPlan_') &&
  read('SpreadsheetService.gs').includes('planGlobalArchiveNumbers') &&
  read('SheetHelpers.gs').includes('writeDetailArchiveNumber_'),
  'Archive numbering must use one global plan and mirror it to rekap/detail sheets.');
assert(!read('PureFunctions.gs').includes('normalizeRekapArchiveSequence_') &&
  !read('SheetHelpers.gs').includes('repairRekapArchiveSequence_'),
  'Per-spreadsheet 1..N normalization must not return.');
assert(!read('SpreadsheetService.gs').includes('cascadeNomorBerkasShift'),
  'Legacy per-activity numbering cascade must not return.');
assert(!read('SheetHelpers.gs').includes('locks.nomorBerkas !== false && expectedNumber'), 'Canonical Nomor Berkas repair must not be bypassed by manual metadata locks.');
assert(read('MetadataService.gs').includes('normalized.no_berkas = normalized.no_berkas || resolveSubActivityArchiveNumber_'), 'Existing detail Nomor Berkas must not be overwritten by a config fallback.');
assert(read('SpreadsheetService.gs').includes('REKAP_SUMMARY_COLUMNS.nomorBerkas') &&
  read('SpreadsheetService.gs').includes('REKAP_SUMMARY_COLUMNS.noFolder'),
  'Global archive number must be mirrored to both Rekap identity columns.');
assert(read('ConfigRepository.gs').includes('bulkUpdateSubActivityNumbering') &&
  read('ConfigRepository.gs').includes('local_sort_order'),
  'Global and local archive ordering must be persisted together in config.');
assert(read('SubActivityController.gs').includes('formalArchiveName') && read('SubActivityController.gs').includes('rekapRowNumber'), 'Sub-activity rename must accept mapping metadata, not only the folder name.');
assert(read('ClientArchiveFolder.html').includes("mode: 'edit-sub-activity'"), 'Mapped sub-activity folders must open the edit-sub-activity flow instead of raw folder rename.');
assert(configRepo.includes('target_folder_id') && configRepo.includes('headers.map(function (header)'), 'Archive log writes must persist target folder fields by header name.');
assert(read('ArchiveController.gs').includes('buildArchiveFileIndex_') && read('ArchiveController.gs').includes('cur.depth + 1'), 'Existing archive adoption must search nested target folders (file index built once per sub-activity).');
assert(read('SubActivityController.gs').includes('persistRekapRowNumber_'), 'New/synced sub-activities must persist their rekap row number.');
assert(read('SubActivityController.gs').includes('markRekapSubActivityInactive'), 'Sub-activity deactivation must mark the rekap row.');
assert(read('SpreadsheetService.gs').includes('clearRekapSubActivityInactiveMark'), 'Sub-activity restore must clear stale inactive rekap notes.');
assert(read('EditSubActivity.html').includes('rename-sub-formal-name'), 'Add/edit sub-activity modal must expose optional formal rekap mapping.');
assert(read('EditSubActivity.html').includes("api('addSubActivity'") && read('EditSubActivity.html').includes("noFolder: (document.getElementById('rename-sub-no-folder')"), 'Add sub-activity submit must pass Nomor Folder into config.');
assert(read('ClientHistory.html').includes('Folder Tujuan'), 'History UI must show final target folder context.');
assert(spreadsheetService.includes("detailSummary.sumLembar ? detailSummary.sumLembar + ' lembar'"), 'Rekap Jumlah must be the SUM of detail-sheet "jumlah" (lembar), not a row count.');
assert(spreadsheetService.includes('sumLembar += n') && spreadsheetService.includes('SUM('), 'Rekap Jumlah must sum the detail "jumlah" (lembar) with consistent JS snapshot and formula.');
assert(!spreadsheetService.includes('detailSummary.filingCabinet'), 'Rekap location columns must strictly inherit top-down and NOT prefer values summarized from the detail sheet.');
assert(spreadsheetService.includes('formatAccessSummary_') && spreadsheetService.includes("values.join(' & ')"), 'Rekap security access must aggregate unique detail-sheet values with ampersand formatting.');
assert(read('ArchiveController.gs').includes('SpreadsheetService.updateRekapSummary(activity, subActivity, {})'), 'Existing-data adoption must refresh rekap summaries from filled detail sheets.');
assert(spreadsheetService.includes('getDetailMetadataDefaults'), 'Archive form defaults must be readable from the detail sheet.');
assert(read('Code.gs').includes('function getArchiveMetadataDefaults'), 'Archive metadata defaults must be exposed to the client.');
assert(read('ClientProcess.html').includes('ensureProcessRequiredFields_') && read('ClientProcess.html').includes('applyArchiveMetadataDefaults_'), 'Archive form must always show and default location/access fields.');
assert(read('ClientProcess.html').includes("key === 'lokasi_simpan'") && read('ClientProcess.html').includes("' readonly'"), 'Archive form Lokasi Simpan must be read-only.');
assert(read('MetadataService.gs').includes('normalized.jumlah') && read('MetadataService.gs').includes('? metadata.jumlah : 1'), 'Archive metadata normalization must default Jumlah to one document per detail row.');
assert(read('EditSubActivity.html').includes('id="rename-sub-no-filing-cabinet"'), 'Edit metadata form must show location metadata fields.');
assert(read('EditSubActivity.html').includes('id="rename-sub-kurun-start"') && read('EditSubActivity.html').includes('id="rename-sub-kurun-end"'), 'Edit metadata form must use date range inputs.');
assert(read('Code.gs').includes('function updateSubActivityMetadata'), 'updateSubActivityMetadata endpoint must be exposed to Apps Script clients.');
assert(read('SpreadsheetService.gs').includes('updateRekapDocumentMetadata'), 'Sub-activity metadata editor must be able to update the rekap sheet.');
assert(!read('ClientDocumentProcess.html').includes('nomorBerkas') && !read('ClientDocumentProcess.html').includes('formatKurunWaktuText_'), 'Document-process form must be link-only; sub-activity metadata is edited separately.');
assert(read('ClientProcess.html').includes("['nomor_surat', 'Nomor Surat'") && read('ClientProcess.html').includes('extractNomorSuratFromText_'), 'Archive-letter form must show editable Nomor Surat with filename parsing.');
assert(read('WorkspaceSetup.gs').includes("['nomor_surat', 'Nomor Surat'") && read('ClientState.html').includes("nomor_surat: 'Nomor Surat'"), 'Nomor Surat must be available in metadata config and labels.');
assert(read('ClientProcess.html').includes('Object.assign({}, state.draft && state.draft.metadata'), 'Archive-letter submit must preserve hidden automatic metadata such as no_berkas.');
assert(read('ClientProcess.html').includes('requiredMark') && read('ClientProcess.html').includes('required-dot'), 'Archive-letter required fields must be marked with a visible star.');
assert(read('ClientProcess.html').includes("if (f.field_name === 'klasifikasi_akses') val = 'Terbatas'") && read('WorkspaceSetup.gs').includes("['klasifikasi_akses', 'Klasifikasi Keamanan & Akses Arsip', true, 'Terbatas'"), 'Archive-letter access classification must default to Terbatas.');
assert(read('WorkspaceSetup.gs').includes('wsBuildLeafSubActivityEntries_'), 'Workspace setup must map leaf folders as final sub-activities.');
assert(read('SubActivityController.gs').includes('wsBuildLeafSubActivityEntries_') && read('SubActivityController.gs').includes('entry.groupName || entry.parentFolderName'), 'Sub-activity sync must use leaf folders and preserve leadership group context.');

// USER_DEPLOYING: app jadi perantara, spreadsheet/folder backend tetap privat ke owner,
// gerbang = login+RBAC app. Aman setelah auth guard + requireWithinWorkspace_ (IDOR) terpasang.
// Catatan deploy: butuh deployment versi baru + re-auth owner + un-share sheet dari 10 user.
assert(appsscript.webapp.executeAs === 'USER_DEPLOYING', 'Web app must execute as deploying user (shared 10-user model; backend sheets stay private).');
assert(appsscript.webapp.access === 'ANYONE', 'Web app must allow signed-in users from any Google domain; portal authentication remains authoritative.');
assert(appsscript.oauthScopes.includes('https://www.googleapis.com/auth/script.external_request'), 'Manifest must include external request scope for UrlFetchApp resumable uploads.');
assert(appsscript.oauthScopes.includes('https://www.googleapis.com/auth/script.scriptapp'), 'Manifest must allow trigger installation.');

assert(client.includes("result.warning || 'Selesai:"), 'Client must surface finalization warnings.');
assert(clientApi.includes('.withFailureHandler('), 'Every google.script.run call must use the centralized failure handler.');
assert(clientApi.includes('API_TIMEOUT_MS') && clientApi.includes('CLIENT_TIMEOUT'), 'Every client RPC must have a default timeout.');
assert(clientApi.includes('API_TIMEOUTS_MS[name]') && read('ClientState.html').includes('initializeWorkspace: 330000'), 'Long-running workspace operations must not inherit the short RPC timeout.');
assert(authService.includes("console.info('LOGIN_PERF '") && read('SettingsController.gs').includes("console.info('BOOTSTRAP_PERF '"), 'Login and bootstrap must emit phase timing diagnostics.');
assert(!clientApi.includes("localStorage.getItem('portal_session_id')") && !clientLogin.includes("localStorage.setItem('portal_session_id'"), 'Authentication token must stay in client memory, not iframe storage.');
assert(clientRouter.includes('rememberPendingRoute_') && clientRouter.includes('state.pendingRoute'), 'Deep-link route must be retained across login.');
assert(clientRouter.includes('cancelGuestLoginPrompt_') && clientRouter.includes("classList.remove('auth-pending')"), 'Closing login as guest must restore a visible safe route.');
assert(clientRouter.includes('state.loginReturnRoute') && clientRouter.includes('google.script.history.replace'), 'Guest login cancellation must restore the previous safe route and replace restricted history.');
assert(clientRouter.includes('prepareRestrictedGuestRoute_') && clientRouter.includes("{ view: 'activity_detail', activityId: activityId }"), 'Restricted guest deep-links must render a safe activity background before showing login.');
assert(clientLogin.includes("close(true)") && clientLogin.includes('cancelGuestLoginPrompt_'), 'Login close actions must invoke guest-route recovery.');
assert(/if \(!isGuest\(\)\) \{[\s\S]*syncDrive_\(activity, true\)/.test(read('ClientActivityDetail.html')), 'Guest activity pages must not trigger authenticated silent Drive sync and hide the safe background.');
assert(clientRouter.includes('requestGeneration !== getAuthGeneration()'), 'Stale auth responses must not overwrite a newer login.');
assert(read('Index.html').includes('auth-pending') && read('StylesBase.html').includes('.app-shell.auth-pending'), 'Main shell must stay hidden until authentication is validated.');
assert(clientLogin.includes('Memverifikasi...') && clientLogin.includes('Memuat data...'), 'Login must show staged progress labels.');
assert(authService.includes('saveSession_') && authService.includes('CacheService.getScriptCache()'), 'Server sessions must use durable properties plus cache.');
assert(authService.includes('sheet.getDataRange().getValues()'), 'Login account scan must use one batched values read.');
assert(!authService.includes('activeEmail:') && !authService.includes('session.activeEmail'), 'Portal authorization must not bind sessions to Google email.');
assert((authService.match(/Session\.getActiveUser\(\)\.getEmail\(\)/g) || []).length === 1 && authService.includes('detectedGoogleEmailForDiagnostics_'), 'Google active email may only be read for diagnostics.');
assert(!read('SystemLogger.gs').includes('Session.getActiveUser'), 'System logger identity must come from portal session context.');
assert(read('Code.gs').includes("return 'ACCESS_DENIED'") && read('Code.gs').includes("return 'AUTH_ERROR'"), 'Authorization denial must be distinct from invalid session.');
assert(clientApi.includes('if (isAccessDeniedError_(error)) return;') && clientApi.indexOf('if (isAccessDeniedError_(error)) return;') < clientApi.indexOf('clearSessionToken();', clientApi.indexOf('function handleSessionError_')), 'Client must preserve session on nonfatal access denial.');
assert(read('SecurityHelpers.gs').includes("SystemLogger.warn('ACCESS_DENIED'") && read('SecurityHelpers.gs').includes('failedCheck'), 'Access denials must write diagnostic context to system_logs.');
assert(configHelpers.includes('sha256Hex_(key) + activeSalt') && !/function pbkdf2Like_[\s\S]*?Utilities\.computeDigest/.test(configHelpers), 'Password verification must run SHA-256 rounds inside V8, not call a GAS service per round.');
assert(client.includes('install-maintenance-trigger'), 'Settings UI must expose maintenance trigger install.');
assert(client.includes('run-cleanup-now'), 'Settings UI must expose manual cleanup.');
assert(client.includes('renderInactiveSubActivities'), 'Client must expose inactive sub-activity page.');
assert(client.includes('renderAuditLog'), 'Client must expose audit log page.');
assert(read('ClientTemplates.html').includes('TEMPLATE_CATEGORY_COLORS'), 'Template UI must define persistent category color presets.');
assert(read('ClientTemplates.html').includes('tall-card-template-cat'), 'Template UI must render category blocks using the tall-card layout.');
assert(client.includes('tab-inactive') || read('Index.html').includes('data-view="inactive"'), 'Inactive page must be accessible.');
assert(client.includes('tab-audit') || read('Index.html').includes('data-view="audit"'), 'Audit page must be accessible.');
assert(read('ArchiveController.gs').includes('parseDocumentContent'), 'Auto-parse backend endpoint must exist in ArchiveController.');
assert(read('ArchiveController.gs').includes('ParseEngine.analyze'), 'Auto-parse must use ParseEngine for extraction.');
assert(read('Code.gs').includes('function parseDocumentContent'), 'Auto-parse must be exposed as a public Apps Script endpoint.');
assert(read('ParseEngine.gs').includes('analyze'), 'ParseEngine module must exist with analyze entry point.');
assert(read('ParseEngine.gs').includes('analyzeStructure'), 'ParseEngine must include document structure analysis.');
assert(read('ParseEngine.gs').includes('classifyDocumentType'), 'ParseEngine must include document type classification.');
assert(read('ClientProcess.html').includes('triggerAutoParse_'), 'Auto-parse trigger function must exist in client.');
assert(read('ClientProcess.html').includes('autoFillParseFields_'), 'Auto-parse auto-fill function must exist in client.');
assert(read('StylesTables.html').includes('auto-parse-filled'), 'Auto-parse green highlight CSS must exist in Styles.');
assert(read('MetadataService.gs').includes('extractKlasifikasiAkses_'), 'Klasifikasi akses extractor must exist in MetadataService.');

// Regresi audit holistik: AppController.gs dihapus saat refactor split — tak boleh ada
// referensi tersisa (akan ReferenceError di runtime). Gunakan WorkspaceController/DriveService.
assert(!read('Code.gs').includes('AppController.') && !read('SubActivityController.gs').includes('AppController.'),
  'AppController (dihapus) tidak boleh direferensikan; gunakan WorkspaceController.getBootstrap / DriveService.*');
// ConfigRepository wajib mendefinisikan metode yang dipakai rekonsiliasi global & maintenance.
assert(read('ConfigRepository.gs').includes('getActivities: function') &&
  read('ConfigRepository.gs').includes('getSubActivities: function') &&
  read('ConfigRepository.gs').includes('getSubActivityById: function'),
  'ConfigRepository harus mendefinisikan getActivities/getSubActivities/getSubActivityById.');
// CONFIG_SHEETS.DOCUMENT_TYPES harus ada (kalau tidak, getOrCreateSheet membuat sheet "undefined").
assert(read('ConfigConstants.gs').includes("DOCUMENT_TYPES: 'config_document_types'"),
  'CONFIG_SHEETS.DOCUMENT_TYPES harus terdefinisi.');

assert(!readme.includes('apps-script/portal-arsip/README.md'), 'README must not point to missing apps-script path.');
assert(!readme.includes('scripts/google-apps-script'), 'README must not point to missing setup script paths.');

const redundantRootMirrors = fs.readdirSync(root).filter(file =>
  file.endsWith('.js') && fs.existsSync(path.join(root, file.replace(/\.js$/, '.gs')))
);
assert(redundantRootMirrors.length === 0,
  'Root-level .js mirrors of .gs files are redundant: ' + redundantRootMirrors.join(', '));
assert(!fs.existsSync(path.join(root, '_DebugAutofill.gs')),
  'Temporary unauthenticated debug helpers must not be shipped to Apps Script.');
assert(!fs.existsSync(path.join(root, 'scripts', 'html-splitter.js')),
  'One-time HTML migration tooling must not return.');

console.log('Portal Arsip verification passed.');
