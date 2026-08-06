'use strict';

/** @private Drive core metadata, upload, and inbox operations. */
const DriveCoreImpl_ = {
  folderDtoFromConfig: function (folderId, folderName) {
    if (!folderId) return null;
    return {
      id: folderId,
      name: folderName || folderId,
      url: 'https://drive.google.com/drive/folders/' + folderId,
      mimeType: DRIVE_FOLDER_MIME_TYPE
    };
  },

  resolveFolderPathAndUrl: function(folderId) {
    if (!folderId) return null;
    try {
      const folder = DriveApp.getFolderById(folderId);
      const url = folder.getUrl();
      const pathParts = [folder.getName()];
      return { path: pathParts[0], url: url };
    } catch (e) {
      console.error('resolveFolderPathAndUrl failed: ' + e.message);
      return null;
    }
  },

  fileDtoFromConfig: function (fileId, fileName, mimeType) {
    if (!fileId) return null;
    const type = mimeType || MimeType.GOOGLE_SHEETS;
    return {
      id: fileId,
      name: fileName || fileId,
      url: 'https://drive.google.com/file/d/' + fileId + '/view',
      downloadUrl: this.getDownloadUrl(fileId, type),
      mimeType: type,
      size: 0
    };
  },

  safeFolderSummary: function (folderId) {
    try {
      return this.getFolderSummary(folderId);
    } catch (error) {
      console.error('DriveService.safeFolderSummary failed for ID ' + folderId + ': ' + error.message);
      return null;
    }
  },

  safeFileSummary: function (fileId) {
    try {
      return this.getFileSummary(fileId);
    } catch (error) {
      console.error('DriveService.safeFileSummary failed for ID ' + fileId + ': ' + error.message);
      return null;
    }
  },

  getOrCreateChildFolder: function (parent, name) {
    return withRetry_(() => {
      const folders = parent.getFoldersByName(name);
      if (folders.hasNext()) return folders.next();
      return parent.createFolder(name);
    });
  },

  resolveSystemFolder: function (settings, rootFolder) {
    let systemFolder;
    if (settings.configFolderId) {
      try {
        systemFolder = DriveApp.getFolderById(cleanId_(settings.configFolderId));
      } catch (e) {
        console.error('DriveService.resolveSystemFolder configFolderId lookup failed: ' + e.message);
        systemFolder = null;
      }
    }
    if (!systemFolder) {
      let systemFolderParent = rootFolder;
      if (settings.systemFolderParentId) {
        try {
          systemFolderParent = DriveApp.getFolderById(cleanId_(settings.systemFolderParentId));
        } catch (e) {
          console.error('DriveService.resolveSystemFolder systemFolderParentId lookup failed: ' + e.message);
        }
      }
      const systemFolderName = settings.systemFolderName || '00. Sistem Portal';
      systemFolder = this.getOrCreateChildFolder(systemFolderParent, systemFolderName);
    }
    return systemFolder;
  },

  fileToDto: function (file) {
    if (!file) return null;
    const mimeType = file.getMimeType();
    return {
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      downloadUrl: this.getDownloadUrl(file.getId(), mimeType),
      mimeType: mimeType,
      size: file.getSize()
    };
  },

  folderToDto: function (folder) {
    if (!folder) return null;
    return {
      id: folder.getId(),
      name: folder.getName(),
      url: folder.getUrl(),
      mimeType: 'application/vnd.google-apps.folder'
    };
  },

  getDownloadUrl: function (fileId, mimeType) {
    if (mimeType === MimeType.GOOGLE_DOCS) {
      return 'https://docs.google.com/document/d/' + fileId + '/export?format=docx';
    }
    if (mimeType === MimeType.GOOGLE_SHEETS) {
      return 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    }
    if (mimeType === MimeType.GOOGLE_SLIDES) {
      return 'https://docs.google.com/presentation/d/' + fileId + '/export/pptx';
    }
    return 'https://drive.google.com/uc?export=download&id=' + fileId;
  },

  uploadToInbox: function (payload) {
    const safeName = validateFilePayload_(payload, 'Upload ke Google Drive dulu lalu tempel link-nya di kolom "Link / ID File Google Drive".');

    const year = Number(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const config = CacheHelper.getConfig(year);
    const yearConfig = ConfigService.getYearConfig(config, year);
    const rootFolder = DriveApp.getFolderById(yearConfig.root_folder_id);
    const settings = ConfigService.getSettings();
    const systemFolder = this.resolveSystemFolder(settings, rootFolder);
    const inbox = yearConfig.inbox_folder_id
      ? DriveApp.getFolderById(cleanId_(yearConfig.inbox_folder_id))
      : this.getOrCreateChildFolder(systemFolder, 'Inbox Dokumen Masuk');
    const target = this.getOrCreateChildFolder(inbox, chooseInboxSubfolder_(payload.name));

    if (payload.mimeType && ALLOWED_MIME_TYPES.indexOf(payload.mimeType) === -1) {
      throw new Error('Jenis file tidak diizinkan. Hanya file PDF, DOC, dan DOCX yang dapat diupload.');
    }

    const blob = dataUrlToBlob_(payload.dataUrl, payload.name, payload.mimeType);
    const file = withRetry_(() => target.createFile(blob));

    return this.fileToDto(file);
  },

  getFileFromInput: function (payload) {
    if (payload.sourceFileId) {
      return DriveApp.getFileById(cleanId_(payload.sourceFileId));
    }
    throw new Error('Pilih atau upload file terlebih dahulu.');
  },

  _isFileUnderInbox_: function (fileId, year) {
    try {
      const file = DriveApp.getFileById(cleanId_(fileId));
      const config = CacheHelper.getConfig(year);
      const yearConfig = ConfigService.getYearConfig(config, year);

      if (!yearConfig.inbox_folder_id) return false;

      let parent = file;
      while (parent) {
        const parents = parent.getParents();
        if (!parents.hasNext()) break;
        const p = parents.next();
        if (p.getId() === yearConfig.inbox_folder_id) return true;
        parent = p;
      }
      return false;
    } catch (e) {
      console.error('_isFileUnderInbox_: ' + e.message);
      return false;
    }
  },

};
