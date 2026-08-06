'use strict';

const MAX_UPLOAD_MB = 50;
const ALLOWED_EXTENSIONS = ['.pdf', '.doc', '.docx'];
const ALLOWED_MIME_TYPES = ['application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
const DRIVE_FOLDER_MIME_TYPE = 'application/vnd.google-apps.folder';
const DRIVE_LIST_PAGE_SIZE = 500;
const FOLDER_PICKER_PAGE_SIZE = 500;

function _validateFileName_(name) {
  if (!name) throw new Error('Nama file tidak lengkap.');
  const safeName = String(name || '').trim();
  if (!safeName || safeName.length > MAX_FILENAME_LENGTH) {
    throw new Error('Nama file tidak valid atau terlalu panjang (maksimal ' + MAX_FILENAME_LENGTH + ' karakter).');
  }
  const extMatch = safeName.match(/\.[a-z0-9]+$/i);
  const ext = extMatch ? extMatch[0].toLowerCase() : '';
  if (!ext || ALLOWED_EXTENSIONS.indexOf(ext) === -1) {
    throw new Error('Jenis file tidak diizinkan. Hanya file PDF, DOC, dan DOCX yang dapat diupload.');
  }
  return safeName;
}

function validateFilePayload_(payload, extraHint) {
  if (!payload || !payload.name || !payload.dataUrl) {
    throw new Error('File upload tidak lengkap.');
  }
  const safeName = _validateFileName_(payload.name);
  const MAX_BASE64_LENGTH = Math.round(MAX_UPLOAD_MB * 1.33) * 1024 * 1024;
  if (payload.dataUrl.length > MAX_BASE64_LENGTH) {
    const fileMB = Math.round(payload.dataUrl.length / 1024 / 1024 / 1.33);
    throw new Error(
      'Ukuran file (' + fileMB + ' MB) melebihi batas upload langsung. ' +
      'Silakan gunakan menu upload di web app untuk upload otomatis bertahap. ' +
      (extraHint || '')
    );
  }
  return safeName;
}

function validateFilePayloadForResumable_(payload) {
  if (!payload || !payload.name) {
    throw new Error('Nama file tidak lengkap.');
  }
  return _validateFileName_(payload.name);
}

/**
 * All Google Drive operations: upload, copy, rename, trash, folder traversal.
 * Works within the workspace folder hierarchy.
 */
const DriveService = {
  folderDtoFromConfig: function (folderId, folderName) {
    return DriveCoreImpl_.folderDtoFromConfig.apply(this, arguments);
  },

  resolveFolderPathAndUrl: function (folderId) {
    return DriveCoreImpl_.resolveFolderPathAndUrl.apply(this, arguments);
  },

  fileDtoFromConfig: function (fileId, fileName, mimeType) {
    return DriveCoreImpl_.fileDtoFromConfig.apply(this, arguments);
  },

  safeFolderSummary: function (folderId) {
    return DriveCoreImpl_.safeFolderSummary.apply(this, arguments);
  },

  safeFileSummary: function (fileId) {
    return DriveCoreImpl_.safeFileSummary.apply(this, arguments);
  },

  getOrCreateChildFolder: function (parent, name) {
    return DriveCoreImpl_.getOrCreateChildFolder.apply(this, arguments);
  },

  resolveSystemFolder: function (settings, rootFolder) {
    return DriveCoreImpl_.resolveSystemFolder.apply(this, arguments);
  },

  fileToDto: function (file) {
    return DriveCoreImpl_.fileToDto.apply(this, arguments);
  },

  folderToDto: function (folder) {
    return DriveCoreImpl_.folderToDto.apply(this, arguments);
  },

  getDownloadUrl: function (fileId, mimeType) {
    return DriveCoreImpl_.getDownloadUrl.apply(this, arguments);
  },

  uploadToInbox: function (payload) {
    return DriveCoreImpl_.uploadToInbox.apply(this, arguments);
  },

  getFileFromInput: function (payload) {
    return DriveCoreImpl_.getFileFromInput.apply(this, arguments);
  },

  _isFileUnderInbox_: function (fileId, year) {
    return DriveCoreImpl_._isFileUnderInbox_.apply(this, arguments);
  },

  copyToFinalFolder: function (sourceFile, targetFolderId, finalFileName, year) {
    return DriveArchiveImpl_.copyToFinalFolder.apply(this, arguments);
  },

  createSubFolder: function (parentFolderId, name) {
    return DriveArchiveImpl_.createSubFolder.apply(this, arguments);
  },

  createChildFolder: function (parentFolderId, name) {
    return DriveArchiveImpl_.createChildFolder.apply(this, arguments);
  },

  addArchiveDocumentLink: function (payload) {
    return DriveArchiveImpl_.addArchiveDocumentLink.apply(this, arguments);
  },

  getShortcutTargetInfo: function (payload) {
    return DriveArchiveImpl_.getShortcutTargetInfo.apply(this, arguments);
  },

  updateArchiveDocumentLink: function (payload) {
    return DriveArchiveImpl_.updateArchiveDocumentLink.apply(this, arguments);
  },

  getFolderSummary: function (folderId) {
    return DriveCrudImpl_.getFolderSummary.apply(this, arguments);
  },

  getFileSummary: function (fileId) {
    return DriveCrudImpl_.getFileSummary.apply(this, arguments);
  },

  renameFolder: function (folderId, name) {
    return DriveCrudImpl_.renameFolder.apply(this, arguments);
  },

  renameFile: function (fileId, name) {
    return DriveCrudImpl_.renameFile.apply(this, arguments);
  },

  trashFolder: function (folderId) {
    return DriveCrudImpl_.trashFolder.apply(this, arguments);
  },

  trashFile: function (fileId) {
    return DriveCrudImpl_.trashFile.apply(this, arguments);
  },

  listFolderContent: function (payload) {
    return DriveCrudImpl_.listFolderContent.apply(this, arguments);
  },

  listFolders: function (payload) {
    return DriveCrudImpl_.listFolders.apply(this, arguments);
  },

  listTemplates: function (year, preloadedConfig) {
    return DriveTemplateImpl_.listTemplates.apply(this, arguments);
  },

  listTemplatesByCategory: function (year) {
    return DriveTemplateImpl_.listTemplatesByCategory.apply(this, arguments);
  },

  getTemplateFolder_: function (year) {
    return DriveTemplateImpl_.getTemplateFolder_.apply(this, arguments);
  },

  uploadTemplateFile: function (payload) {
    return DriveTemplateImpl_.uploadTemplateFile.apply(this, arguments);
  },

  trashTemplateFile: function (fileId, year) {
    return DriveTemplateImpl_.trashTemplateFile.apply(this, arguments);
  },

  createTemplateCategory: function (year, name) {
    return DriveTemplateImpl_.createTemplateCategory.apply(this, arguments);
  },

  renameTemplateCategoryFolder: function (folderId, newName, year) {
    return DriveTemplateImpl_.renameTemplateCategoryFolder.apply(this, arguments);
  },

  deleteTemplateCategoryFolder: function (folderId, year) {
    return DriveTemplateImpl_.deleteTemplateCategoryFolder.apply(this, arguments);
  },

  initTemplateResumableUpload: function (payload) {
    return DriveTemplateImpl_.initTemplateResumableUpload.apply(this, arguments);
  },

  uploadResumableChunk: function (payload) {
    return DriveTemplateImpl_.uploadResumableChunk.apply(this, arguments);
  }
};
