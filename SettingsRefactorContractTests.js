'use strict';

/** Private staging-safe contract test for SettingsController. */
function runSettingsRefactorContractTests_() {
  const expected = ["getBootstrap","saveSettings","resetWorkspace","initializeWorkspace","ensureArchiveMaintenanceTrigger","ensureDocumentTypesSyncTrigger","syncDocumentTypeColumns","buildProgress","updateSubActivityMapping","renameDriveItem","listDriveFolders","getTemplatesData","getAdminAuditLogs","getTemplateCategories","saveTemplateCategory","renameTemplateCategory","deleteTemplateCategory","deleteYear","setTemplateCategory"];
  const mapping = {"getBootstrap":"SettingsBootstrapImpl_","saveSettings":"SettingsBootstrapImpl_","resetWorkspace":"SettingsWorkspaceImpl_","initializeWorkspace":"SettingsWorkspaceImpl_","ensureArchiveMaintenanceTrigger":"SettingsWorkspaceImpl_","ensureDocumentTypesSyncTrigger":"SettingsWorkspaceImpl_","syncDocumentTypeColumns":"SettingsWorkspaceImpl_","buildProgress":"SettingsMappingImpl_","updateSubActivityMapping":"SettingsMappingImpl_","renameDriveItem":"SettingsMappingImpl_","listDriveFolders":"SettingsMappingImpl_","getTemplatesData":"SettingsTemplateImpl_","getAdminAuditLogs":"SettingsTemplateImpl_","getTemplateCategories":"SettingsTemplateImpl_","saveTemplateCategory":"SettingsTemplateImpl_","renameTemplateCategory":"SettingsTemplateImpl_","deleteTemplateCategory":"SettingsTemplateImpl_","deleteYear":"SettingsYearImpl_","setTemplateCategory":"SettingsTemplateImpl_"};
  const implementations = { SettingsBootstrapImpl_: SettingsBootstrapImpl_, SettingsWorkspaceImpl_: SettingsWorkspaceImpl_, SettingsMappingImpl_: SettingsMappingImpl_, SettingsTemplateImpl_: SettingsTemplateImpl_, SettingsYearImpl_: SettingsYearImpl_ };
  expected.forEach(function (name) {
    if (typeof SettingsController[name] !== 'function') throw new Error('Facade method hilang: ' + name);
    if (!implementations[mapping[name]] || typeof implementations[mapping[name]][name] !== 'function') throw new Error('Implementasi method hilang: ' + name);
  });
  if (Object.keys(SettingsController).length !== expected.length) throw new Error('Jumlah method facade berubah.');
  const progress = SettingsController.buildProgress([], []);
  if (!progress || progress.total !== 0 || progress.completionRate !== 0) throw new Error('Delegasi buildProgress gagal.');
  if (inferSubActivityMappingStatus_({ sub_activity_name: 'A', formal_archive_name: 'A', target_sheet_name: 'A' }) !== 'PERLU_REVIEW') throw new Error('Mapping helper tidak terhubung.');
  return { ok: true, facadeMethods: expected.length, implementationGroups: Object.keys(implementations).length, helperCheck: true, delegationCheck: true };
}
