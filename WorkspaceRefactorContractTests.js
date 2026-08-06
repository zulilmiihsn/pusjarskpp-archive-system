'use strict';

/** Private staging-safe contract test for the WorkspaceSetup refactor. */
function runWorkspaceRefactorContractTests_() {
  const expected = ["parseLeadingNumber","initialize","initializeSingleYear_","scanAndImportPhysicalYears_"];
  expected.forEach(function (name) {
    if (typeof WorkspaceSetupService[name] !== 'function') throw new Error('Facade method hilang: ' + name);
    if (typeof WorkspaceSetupImpl_[name] !== 'function') throw new Error('Implementasi method hilang: ' + name);
  });
  if (Object.keys(WorkspaceSetupService).length !== expected.length) throw new Error('Jumlah method facade berubah.');
  if (WorkspaceSetupService.parseLeadingNumber('12. Folder') !== '12') throw new Error('Delegasi parseLeadingNumber gagal.');
  if (wsNormalize_(' Halo, DUNIA! ') !== 'halo dunia') throw new Error('Workspace Drive helpers tidak terhubung.');
  if (wsBuildPath_('Root', ['A', 'B']) !== 'Root > A > B') throw new Error('Path helper tidak terhubung.');
  const report = []; wsPushReport_(report, 'ok', 'uji');
  if (report.length !== 1 || report[0].label !== 'uji') throw new Error('Report helper tidak terhubung.');
  return { ok: true, facadeMethods: expected.length, helperCheck: true, delegationCheck: true };
}
