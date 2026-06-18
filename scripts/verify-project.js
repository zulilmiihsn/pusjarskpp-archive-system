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

const appController = read('AppController.gs') + '\n' + read('ArchiveController.gs') + '\n' + read('SubActivityController.gs') + '\n' + read('SettingsController.gs') + '\n' + read('PureFunctions.gs');
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
assert(read('AppController.gs').includes('initializeWorkspace: function (payload) { requireAdminIfWorkspaceSecured_(payload);'), 'First workspace initialization must not require an existing admin session.');
assert(read('AppController.gs').includes('function requireAdminIfWorkspaceSecured_') && read('AppController.gs').includes('hasActiveAdminAccount_'), 'Workspace auth guard must allow first-run or half-initialized recovery before an admin exists.');
assert(read('SettingsController.gs').includes('const adminResult = AuthService.saveDefaultAdmin();'), 'Workspace initialization must create the default admin through the internal auth service.');
assert(!read('SettingsController.gs').includes('AppController.saveDefaultAdmin()'), 'Workspace initialization must not call the admin-protected saveDefaultAdmin endpoint.');

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
assert(spreadsheetService.includes("detailSummary.count ? detailSummary.count + ' dokumen'"), 'Rekap Jumlah must be derived from detail-sheet document count.');
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
assert(read('ClientProcess.html').includes("if (f.field_name === 'klasifikasi_akses') val = 'Terbatas'") && read('WorkspaceSetup.gs').includes("['klasifikasi_akses', 'Klasifikasi Keamanan & Akses Arsip', false, 'Terbatas'"), 'Archive-letter access classification must default to Terbatas.');
assert(read('WorkspaceSetup.gs').includes('wsBuildLeafSubActivityEntries_'), 'Workspace setup must map leaf folders as final sub-activities.');
assert(read('SubActivityController.gs').includes('wsBuildLeafSubActivityEntries_') && read('SubActivityController.gs').includes('entry.groupName || entry.parentFolderName'), 'Sub-activity sync must use leaf folders and preserve leadership group context.');

assert(appsscript.webapp.executeAs === 'USER_ACCESSING', 'Web app must execute as the accessing user.');
assert(appsscript.webapp.access === 'DOMAIN', 'Web app access must be domain-restricted.');
assert(appsscript.oauthScopes.includes('https://www.googleapis.com/auth/script.external_request'), 'Manifest must include external request scope for UrlFetchApp resumable uploads.');
assert(appsscript.oauthScopes.includes('https://www.googleapis.com/auth/script.scriptapp'), 'Manifest must allow trigger installation.');

assert(client.includes("result.warning || 'Selesai:"), 'Client must surface finalization warnings.');
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
assert(read('Styles.html').includes('auto-parse-filled'), 'Auto-parse green highlight CSS must exist in Styles.');
assert(read('MetadataService.gs').includes('extractKlasifikasiAkses_'), 'Klasifikasi akses extractor must exist in MetadataService.');

assert(!readme.includes('apps-script/portal-arsip/README.md'), 'README must not point to missing apps-script path.');
assert(!readme.includes('scripts/google-apps-script'), 'README must not point to missing setup script paths.');

console.log('Portal Arsip verification passed.');
