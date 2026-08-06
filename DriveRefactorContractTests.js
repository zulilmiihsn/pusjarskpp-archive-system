'use strict';

/**
 * Private staging-safe contract test for the DriveService refactor.
 * The trailing underscore prevents google.script.run from exposing it to clients.
 */
function runDriveRefactorContractTests_() {
  const expected = ["folderDtoFromConfig","resolveFolderPathAndUrl","fileDtoFromConfig","safeFolderSummary","safeFileSummary","getOrCreateChildFolder","resolveSystemFolder","fileToDto","folderToDto","getDownloadUrl","uploadToInbox","getFileFromInput","_isFileUnderInbox_","copyToFinalFolder","createSubFolder","createChildFolder","addArchiveDocumentLink","getShortcutTargetInfo","updateArchiveDocumentLink","getFolderSummary","getFileSummary","renameFolder","renameFile","trashFolder","trashFile","listFolderContent","listFolders","listTemplates","listTemplatesByCategory","getTemplateFolder_","uploadTemplateFile","trashTemplateFile","createTemplateCategory","renameTemplateCategoryFolder","deleteTemplateCategoryFolder","initTemplateResumableUpload","uploadResumableChunk"];
  const mapping = {"folderDtoFromConfig":"DriveCoreImpl_","resolveFolderPathAndUrl":"DriveCoreImpl_","fileDtoFromConfig":"DriveCoreImpl_","safeFolderSummary":"DriveCoreImpl_","safeFileSummary":"DriveCoreImpl_","getOrCreateChildFolder":"DriveCoreImpl_","resolveSystemFolder":"DriveCoreImpl_","fileToDto":"DriveCoreImpl_","folderToDto":"DriveCoreImpl_","getDownloadUrl":"DriveCoreImpl_","uploadToInbox":"DriveCoreImpl_","getFileFromInput":"DriveCoreImpl_","_isFileUnderInbox_":"DriveCoreImpl_","copyToFinalFolder":"DriveArchiveImpl_","createSubFolder":"DriveArchiveImpl_","createChildFolder":"DriveArchiveImpl_","addArchiveDocumentLink":"DriveArchiveImpl_","getShortcutTargetInfo":"DriveArchiveImpl_","updateArchiveDocumentLink":"DriveArchiveImpl_","getFolderSummary":"DriveCrudImpl_","getFileSummary":"DriveCrudImpl_","renameFolder":"DriveCrudImpl_","renameFile":"DriveCrudImpl_","trashFolder":"DriveCrudImpl_","trashFile":"DriveCrudImpl_","listFolderContent":"DriveCrudImpl_","listFolders":"DriveCrudImpl_","listTemplates":"DriveTemplateImpl_","listTemplatesByCategory":"DriveTemplateImpl_","getTemplateFolder_":"DriveTemplateImpl_","uploadTemplateFile":"DriveTemplateImpl_","trashTemplateFile":"DriveTemplateImpl_","createTemplateCategory":"DriveTemplateImpl_","renameTemplateCategoryFolder":"DriveTemplateImpl_","deleteTemplateCategoryFolder":"DriveTemplateImpl_","initTemplateResumableUpload":"DriveTemplateImpl_","uploadResumableChunk":"DriveTemplateImpl_"};
  const implementations = {
    DriveCoreImpl_: DriveCoreImpl_,
    DriveArchiveImpl_: DriveArchiveImpl_,
    DriveCrudImpl_: DriveCrudImpl_,
    DriveTemplateImpl_: DriveTemplateImpl_
  };

  expected.forEach(function (name) {
    if (typeof DriveService[name] !== 'function') {
      throw new Error('Facade method hilang: ' + name);
    }
    const implementation = implementations[mapping[name]];
    if (!implementation || typeof implementation[name] !== 'function') {
      throw new Error('Implementasi method hilang: ' + name);
    }
  });

  const actual = Object.keys(DriveService);
  if (actual.length !== expected.length) {
    throw new Error('Jumlah method facade berubah: ' + actual.length + ' dari ' + expected.length);
  }

  const folder = DriveService.folderDtoFromConfig('folder-test', 'Folder Test');
  if (!folder || folder.id !== 'folder-test' || folder.name !== 'Folder Test') {
    throw new Error('Delegasi folderDtoFromConfig gagal.');
  }

  const file = DriveService.fileDtoFromConfig('file-test', 'File Test', 'application/pdf');
  if (!file || file.id !== 'file-test' || file.downloadUrl.indexOf('file-test') < 0) {
    throw new Error('Delegasi berantai fileDtoFromConfig/getDownloadUrl gagal.');
  }

  if (validateDriveItemName_('arsip.pdf', 'Nama file') !== 'arsip.pdf') {
    throw new Error('DriveHelpers tidak terhubung.');
  }

  return {
    ok: true,
    facadeMethods: actual.length,
    implementationGroups: Object.keys(implementations).length,
    helperCheck: true,
    delegationCheck: true
  };
}
