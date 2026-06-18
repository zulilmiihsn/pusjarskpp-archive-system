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
      'File terlalu besar (' + fileMB + ' MB) untuk diupload langsung. ' +
      'Maksimal upload langsung adalah ' + MAX_UPLOAD_MB + ' MB. ' +
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
    if (!folderId) return null;
    return {
      id: folderId,
      name: folderName || folderId,
      url: 'https://drive.google.com/drive/folders/' + folderId,
      mimeType: DRIVE_FOLDER_MIME_TYPE
    };
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

  copyToFinalFolder: function (sourceFile, targetFolderId, finalFileName, year) {
    return withRetry_(() => {
      const targetFolder = DriveApp.getFolderById(targetFolderId);
      const fileName = getUniqueFileName_(targetFolder, finalFileName || sourceFile.getName());

      let isAlreadyInTarget = false;
      const parents = sourceFile.getParents();
      while (parents.hasNext()) {
        if (parents.next().getId() === targetFolderId) {
          isAlreadyInTarget = true;
          break;
        }
      }

      if (isAlreadyInTarget) {
        if (sourceFile.getName() !== fileName) {
           sourceFile.setName(fileName);
        }
        return sourceFile;
      }

      const isFromInbox = year && DriveService._isFileUnderInbox_(sourceFile.getId(), year);
      if (isFromInbox) {
        const moved = sourceFile.moveTo(targetFolder);
        moved.setName(fileName);
        return moved;
      }
      return sourceFile.makeCopy(fileName, targetFolder);
    });
  },

  createSubFolder: function (parentFolderId, name) {
    return this.getOrCreateChildFolder(DriveApp.getFolderById(parentFolderId), name);
  },

  createChildFolder: function (parentFolderId, name) {
    const parent = DriveApp.getFolderById(cleanId_(parentFolderId));
    const folderName = validateDriveItemName_(name, 'Nama folder');
    return this.folderToDto(this.getOrCreateChildFolder(parent, folderName));
  },

  addArchiveDocumentLink: function (payload) {
    const parentId = cleanId_(payload.parentFolderId);
    const name = validateDriveItemName_(payload.name, 'Kategori dokumen');
    const url = payload.url || '';
    let target = null;

    try {
      target = resolveArchiveDocumentShortcutTarget_(url);
      
      const persuratanSubFolder = DriveApp.getFolderById(parentId);
      let current = persuratanSubFolder;
      const pathParts = [];
      let naskahDinasFolder = null;
      
      while (current) {
        const parents = current.getParents();
        if (!parents.hasNext()) break;
        const parent = parents.next();
        const pName = parent.getName().toLowerCase();
        
        if (pName.indexOf('persuratan') >= 0 || pName.indexOf('1. persuratan') >= 0) {
          const ndParents = parent.getParents();
          if (ndParents.hasNext()) {
            naskahDinasFolder = ndParents.next();
          }
          pathParts.unshift(current.getName());
          break;
        }
        
        pathParts.unshift(current.getName());
        current = parent;
      }
      
      let targetFolder = persuratanSubFolder;
      if (naskahDinasFolder && pathParts.length > 0) {
        const docFolders = naskahDinasFolder.getFoldersByName('2. Dokumen');
        let dokumenFolder;
        if (docFolders.hasNext()) {
          dokumenFolder = docFolders.next();
        } else {
          const docFolders2 = naskahDinasFolder.getFoldersByName('Dokumen');
          if (docFolders2.hasNext()) {
            dokumenFolder = docFolders2.next();
          } else {
            dokumenFolder = naskahDinasFolder.createFolder('2. Dokumen');
          }
        }
        
        let tempFolder = dokumenFolder;
        for (let i = 0; i < pathParts.length; i++) {
          tempFolder = DriveService.getOrCreateChildFolder(tempFolder, pathParts[i]);
        }
        targetFolder = tempFolder;
      }

      const targetFolderId = targetFolder.getId();
      const existingShortcuts = Drive.Files.list({
        q: "'" + cleanId_(targetFolderId) + "' in parents and mimeType='application/vnd.google-apps.shortcut' and name='" + escapeDriveQueryValue_(name) + "' and trashed=false",
        fields: 'files(id,name)',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });

      if (existingShortcuts.files && existingShortcuts.files.length > 0) {
        existingShortcuts.files.forEach(f => {
          try {
            Drive.Files.update({ trashed: true }, f.id, null, { supportsAllDrives: true });
          } catch (ignore) {}
        });
      }

      const shortcut = Drive.Files.create({
        name: name,
        mimeType: 'application/vnd.google-apps.shortcut',
        shortcutDetails: { targetId: target.id }
      }, null, { supportsAllDrives: true });
      
      const file = DriveApp.getFileById(shortcut.id);
      file.moveTo(targetFolder);
      
      return this.fileToDto(file);
    } catch (e) {
      console.error('Failed to create shortcut: ' + e.message);
      throw buildArchiveDocumentShortcutError_(e, target);
    }
  },

  getShortcutTargetInfo: function (payload) {
    payload = payload || {};
    const fileId = cleanId_(payload.fileId);
    if (!fileId) throw new Error('File ID wajib diisi.');
    const meta = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,shortcutDetails,webViewLink'
    });
    if (meta.mimeType !== 'application/vnd.google-apps.shortcut') {
      throw new Error('File ini bukan shortcut Drive.');
    }
    if (!meta.shortcutDetails || !meta.shortcutDetails.targetId) {
      throw new Error('Shortcut tidak memiliki target.');
    }
    const targetId = meta.shortcutDetails.targetId;
    let targetName = '';
    let targetUrl = '';
    let targetMime = '';
    try {
      const target = Drive.Files.get(targetId, {
        supportsAllDrives: true,
        fields: 'id,name,mimeType,webViewLink'
      });
      targetName = target.name || '';
      targetUrl = target.webViewLink || 'https://drive.google.com/file/d/' + targetId + '/view';
      targetMime = target.mimeType || '';
    } catch (e) {
      targetName = '(target tidak dapat dibaca)';
      targetUrl = 'https://drive.google.com/file/d/' + targetId + '/view';
    }
    return {
      fileId: meta.id,
      fileName: meta.name,
      fileUrl: meta.webViewLink || '',
      targetId: targetId,
      targetName: targetName,
      targetUrl: targetUrl,
      targetMimeType: targetMime
    };
  },

  updateArchiveDocumentLink: function (payload) {
    payload = payload || {};
    const fileId = cleanId_(payload.fileId);
    if (!fileId) throw new Error('File ID wajib diisi.');
    const url = payload.url || '';
    if (!url) throw new Error('URL wajib diisi.');
    const meta = Drive.Files.get(fileId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,shortcutDetails'
    });
    if (meta.mimeType !== 'application/vnd.google-apps.shortcut') {
      throw new Error('File ini bukan shortcut Drive.');
    }
    const target = resolveArchiveDocumentShortcutTarget_(url);
    Drive.Files.update({
      shortcutDetails: { targetId: target.id }
    }, fileId, null, { supportsAllDrives: true });
    const file = DriveApp.getFileById(fileId);
    return this.fileToDto(file);
  },

  getFolderSummary: function (folderId) {
    if (!folderId) return null;
    return this.folderToDto(DriveApp.getFolderById(cleanId_(folderId)));
  },

  getFileSummary: function (fileId) {
    if (!fileId) return null;
    return this.fileToDto(DriveApp.getFileById(cleanId_(fileId)));
  },

  renameFolder: function (folderId, name) {
    if (!folderId || !name) throw new Error('Folder dan nama baru wajib diisi.');
    const folder = DriveApp.getFolderById(cleanId_(folderId));
    folder.setName(validateDriveItemName_(name, 'Nama folder'));
    return this.folderToDto(folder);
  },

  renameFile: function (fileId, name) {
    if (!fileId || !name) throw new Error('File dan nama baru wajib diisi.');
    const file = DriveApp.getFileById(cleanId_(fileId));
    file.setName(validateDriveItemName_(name, 'Nama file'));
    return this.fileToDto(file);
  },

  trashFolder: function (folderId) {
    DriveApp.getFolderById(cleanId_(folderId)).setTrashed(true);
    return { id: folderId, trashed: true };
  },

  trashFile: function (fileId) {
    DriveApp.getFileById(cleanId_(fileId)).setTrashed(true);
    return { id: fileId, trashed: true };
  },

  listFolderContent: function (payload) {
    payload = payload || {};
    const folderId = cleanId_(payload.folderId);
    const pageSize = normalizeDrivePageSize_(payload.pageSize, DRIVE_LIST_PAGE_SIZE, 1000);
    const pageToken = payload.pageToken ? String(payload.pageToken) : undefined;

    const listOptions = {
      q: "'" + cleanId_(folderId) + "' in parents and trashed=false",
      fields: "nextPageToken,files(id,name,mimeType,size,webViewLink,createdTime,modifiedTime)",
      pageSize: pageSize,
      orderBy: "folder,name_natural",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    };
    if (pageToken) listOptions.pageToken = pageToken;

    const result = Drive.Files.list(listOptions);

    const folders = [];
    const files = [];
    (result.files || []).forEach(function (item) {
      const dto = driveFileResourceToDto_(item);
      if (item.mimeType === DRIVE_FOLDER_MIME_TYPE) {
        folders.push(dto);
      } else {
        files.push(dto);
      }
    });

    // Optimasi: Gunakan Drive REST API (Drive.Files.get) untuk mendapatkan metadata folder saat ini
    // dan folder induknya, menghindari overhead pemanggilan RPC DriveApp yang lambat.
    let currentDto = { id: folderId, name: '', url: '', mimeType: 'application/vnd.google-apps.folder' };
    let parentDto = null;
    try {
      const folderMeta = Drive.Files.get(folderId, {
        fields: "id,name,parents,webViewLink",
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
      currentDto = {
        id: folderMeta.id,
        name: folderMeta.name || '',
        url: folderMeta.webViewLink || '',
        mimeType: 'application/vnd.google-apps.folder'
      };
      if (folderMeta.parents && folderMeta.parents.length > 0) {
        parentDto = {
          id: folderMeta.parents[0],
          name: '',
          url: '',
          mimeType: 'application/vnd.google-apps.folder'
        };
      }
    } catch (e) {
      console.warn('Gagal memuat metadata folder dengan REST API:', e);
      try {
        // Fallback jika terjadi kendala
        const folder = DriveApp.getFolderById(folderId);
        currentDto = this.folderToDto(folder);
        parentDto = getFirstParentFolder_(folder);
      } catch (_) {}
    }

    return {
      current: currentDto,
      parent: parentDto,
      folders: folders,
      files: files,
      nextPageToken: result.nextPageToken || '',
      pageSize: pageSize,
      partial: !!result.nextPageToken
    };
  },

  listFolders: function (payload) {
    payload = payload || {};
    const isRoot = !payload.parentFolderId || String(payload.parentFolderId) === 'root';
    const folder = getFolderByPickerId_(payload.parentFolderId);
    const folderId = isRoot ? 'root' : folder.getId();
    const pageSize = normalizeDrivePageSize_(payload.pageSize, FOLDER_PICKER_PAGE_SIZE, 1000);
    const pageToken = payload.pageToken ? String(payload.pageToken) : undefined;

    const listOptions = {
      q: "'" + (folderId === 'root' ? 'root' : cleanId_(folderId)) + "' in parents and mimeType='" + DRIVE_FOLDER_MIME_TYPE + "' and trashed=false",
      fields: "nextPageToken,files(id,name,mimeType,webViewLink)",
      pageSize: pageSize,
      orderBy: "name_natural",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    };
    if (pageToken) listOptions.pageToken = pageToken;

    const result = Drive.Files.list(listOptions);

    const children = [];
    (result.files || []).forEach(function (item) {
      children.push(driveFileResourceToDto_(item));
    });

    return {
      current: Object.assign(this.folderToDto(folder), { isRoot: isRoot }),
      parent: getFirstParentFolder_(folder),
      folders: children,
      nextPageToken: result.nextPageToken || '',
      pageSize: pageSize,
      partial: !!result.nextPageToken
    };
  },

  listTemplates: function (year, preloadedConfig) {
    try {
      const config = preloadedConfig || CacheHelper.getConfig(year);
      const yearConfig = ConfigService.getYearConfig(config, year);
      if (!yearConfig.template_folder_id) return [];
      const folder = DriveApp.getFolderById(cleanId_(yearConfig.template_folder_id));
      const files = folder.getFiles();
      const result = [];
      while (files.hasNext()) {
        const file = files.next();
        result.push(this.fileToDto(file));
      }
      return result.sort(function (a, b) { return a.name.localeCompare(b.name); });
    } catch (error) {
      console.error('DriveService.listTemplates failed: ' + error.message);
      return [];
    }
  },

  /**
   * List templates grouped by category folder.
   * Uses Drive API v3 batch queries for speed.
   * Returns { categories: [{id, name, files: [...]}], uncategorized: [fileDto...] }
   */
  listTemplatesByCategory: function (year) {
    try {
      const root = this.getTemplateFolder_(year);
      const rootId = root.getId();

      const allById = {};
      function addFile(f) {
        allById[f.id] = true;
        return f;
      }
      function addFiles(arr) {
        arr.forEach(function(f) { allById[f.id] = true; });
        return arr;
      }

      const folderIds = [];
      const foldersByName = {};
      let pageToken;

      do {
        const folderResult = Drive.Files.list({
          q: "'" + cleanId_(rootId) + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
          fields: "nextPageToken,files(id,name)",
          pageToken: pageToken,
          pageSize: 100,
          supportsAllDrives: true
        });
        if (folderResult.files) {
          folderResult.files.forEach(function (f) {
            folderIds.push(f.id);
            foldersByName[f.id] = f;
          });
        }
        pageToken = folderResult.nextPageToken;
      } while (pageToken);

      const categories = folderIds.map(function (fid) {
        const files = [];
        let pt;
        do {
          const fileResult = Drive.Files.list({
            q: "'" + cleanId_(fid) + "' in parents and trashed=false",
            fields: "nextPageToken,files(id,name,mimeType,size,webViewLink)",
            pageToken: pt,
            pageSize: 100,
            supportsAllDrives: true
          });
          if (fileResult.files) {
            fileResult.files.forEach(function (f) {
              files.push({
                id: f.id,
                name: f.name,
                url: f.webViewLink,
                downloadUrl: DriveService.getDownloadUrl(f.id, f.mimeType),
                mimeType: f.mimeType,
                size: f.size
              });
            });
          }
          pt = fileResult.nextPageToken;
        } while (pt);
        files.sort(function (a, b) { return a.name.localeCompare(b.name); });
        const folder = foldersByName[fid];
        return { id: fid, name: folder ? folder.name : fid, files: files };
      });

      categories.sort(function (a, b) { return a.name.localeCompare(b.name, 'id'); });

      const uncategorized = [];
      let pt2;
      do {
        const rootResult = Drive.Files.list({
          q: "'" + cleanId_(rootId) + "' in parents and mimeType!='application/vnd.google-apps.folder' and trashed=false",
          fields: "nextPageToken,files(id,name,mimeType,size,webViewLink)",
          pageToken: pt2,
          pageSize: 100,
          supportsAllDrives: true
        });
        if (rootResult.files) {
          rootResult.files.forEach(function (f) {
            uncategorized.push({
              id: f.id,
              name: f.name,
              url: f.webViewLink,
              downloadUrl: DriveService.getDownloadUrl(f.id, f.mimeType),
              mimeType: f.mimeType,
              size: f.size
            });
          });
        }
        pt2 = rootResult.nextPageToken;
      } while (pt2);
      uncategorized.sort(function (a, b) { return a.name.localeCompare(b.name); });

      return { categories: categories, uncategorized: uncategorized };
    } catch (error) {
      console.error('DriveService.listTemplatesByCategory failed: ' + error.message);
      return { categories: [], uncategorized: [] };
    }
  },

  getTemplateFolder_: function (year) {
    const config = CacheHelper.getConfig(year);
    const yearConfig = ConfigService.getYearConfig(config, year);
    if (!yearConfig.template_folder_id) {
      throw new Error('Folder template belum dikonfigurasi untuk tahun ini.');
    }
    return DriveApp.getFolderById(cleanId_(yearConfig.template_folder_id));
  },

  uploadTemplateFile: function (payload) {
    return withRetry_(() => {
      const safeName = validateFilePayload_(payload);
      const year = Number(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      let targetFolder = this.getTemplateFolder_(year);
      if (payload.categoryFolderId) {
        try {
          targetFolder = DriveApp.getFolderById(cleanId_(payload.categoryFolderId));
        } catch (e) {
          console.warn('uploadTemplateFile: category folder not found, fallback to root. ' + e.message);
        }
      }
      return this.fileToDto(targetFolder.createFile(dataUrlToBlob_(payload.dataUrl, safeName, payload.mimeType)));
    });
  },

  trashTemplateFile: function (fileId) {
    if (!fileId) throw new Error('File ID wajib diisi.');
    DriveApp.getFileById(cleanId_(fileId)).setTrashed(true);
    return { id: fileId, trashed: true };
  },

  createTemplateCategory: function (year, name) {
    const root = this.getTemplateFolder_(year);
    const folder = this.getOrCreateChildFolder(root, name);
    return this.folderToDto(folder);
  },

  renameTemplateCategoryFolder: function (folderId, newName) {
    return this.renameFolder(folderId, newName);
  },

  deleteTemplateCategoryFolder: function (folderId) {
    return this.trashFolder(folderId);
  },

  initTemplateResumableUpload: function (payload) {
    payload = payload || {};
    const safeName = validateDriveItemName_(payload.name, 'Nama file');
    const year = Number(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const mimeType = payload.mimeType || 'application/octet-stream';
    const totalSize = Number(payload.totalSize);
    if (!totalSize || totalSize < 1) throw new Error('Ukuran file tidak valid.');

    let targetFolder = this.getTemplateFolder_(year);
    if (payload.categoryFolderId) {
      try {
        targetFolder = DriveApp.getFolderById(cleanId_(payload.categoryFolderId));
      } catch (e) {
        console.warn('category folder not found, fallback to root. ' + e.message);
      }
    }
    return resumableUploadInit_(safeName, mimeType, totalSize, targetFolder.getId());
  },

  uploadResumableChunk: function (payload) {
    payload = payload || {};
    const sessionUrl = payload.sessionUrl;
    const chunkBase64 = payload.chunk;
    const startByte = Number(payload.startByte);
    const totalSize = Number(payload.totalSize);
    if (!sessionUrl || !chunkBase64) throw new Error('Parameter upload tidak lengkap.');
    return resumableUploadChunk_(sessionUrl, chunkBase64, startByte, totalSize);
  }
};

// ============================================================
// Resumable Upload (>50 MB) via Drive API v3
// ============================================================

function assertResumableSession_(sessionUrl) {
  if (!sessionUrl || sessionUrl.indexOf('https://') !== 0) {
    throw new Error('Session upload tidak valid.');
  }
}

function resumableUploadInit_(name, mimeType, totalSize, folderId) {
  const safeName = validateDriveItemName_(name, 'Nama file');
  const metadata = { name: safeName };
  if (folderId) metadata.parents = [cleanId_(folderId)];

  const token = ScriptApp.getOAuthToken();
  const headers = {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json; charset=UTF-8',
    'X-Upload-Content-Type': mimeType || 'application/octet-stream',
    'X-Upload-Content-Length': String(totalSize)
  };

  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable',
    {
      method: 'POST',
      headers: headers,
      payload: JSON.stringify(metadata),
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    const body = response.getContentText();
    throw new Error('Gagal mulai upload besar: ' + (body ? body.slice(0, 200) : 'HTTP ' + code));
  }

  const location = response.getHeaders()['Location'];
  if (!location) throw new Error('Server tidak mengembalikan URL upload.');

  return { sessionUrl: location, totalSize: totalSize };
}

function resumableUploadChunk_(sessionUrl, chunkBase64, startByte, totalSize) {
  assertResumableSession_(sessionUrl);
  const chunkBytes = Utilities.base64Decode(chunkBase64);
  const chunkSize = chunkBytes.length;
  const endByte = startByte + chunkSize - 1;

  const response = UrlFetchApp.fetch(sessionUrl, {
    method: 'PUT',
    headers: { 'Content-Range': 'bytes ' + startByte + '-' + endByte + '/' + totalSize },
    payload: chunkBytes,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code === 200 || code === 201) {
    const fileData = JSON.parse(response.getContentText());
    return {
      done: true,
      fileId: fileData.id,
      name: fileData.name,
      url: 'https://drive.google.com/file/d/' + fileData.id + '/view',
      mimeType: fileData.mimeType,
      size: Number(fileData.size || 0)
    };
  }

  if (code === 308) {
    const range = response.getHeaders()['Range'];
    let nextByte = startByte + chunkSize;
    if (range) {
      const m = range.match(/(\d+)$/);
      if (m) nextByte = parseInt(m[1], 10) + 1;
    }
    return { done: false, nextByte: nextByte };
  }

  const body = response.getContentText();
  throw new Error('Upload gagal (HTTP ' + code + '): ' + (body ? body.slice(0, 200) : ''));
}

/**
 * Given an activity's target_folder_id (which lives inside 1. Persuratan > Tahun XXXX),
 * resolve the corresponding mirror folder inside 2. Dokumen > Tahun XXXX > same activity folder.
 * Returns the mirror activity folder, or null if the structure cannot be resolved.
 */
function resolveDokumenMirrorForActivity_(activityTargetFolderId) {
  try {
    // activityTargetFolderId points to e.g. "Folder 2 (Latsar CPNS)" inside "1. Persuratan > Tahun 2026"
    const activityFolder = DriveApp.getFolderById(activityTargetFolderId);
    const activityFolderName = activityFolder.getName();

    // Go up: activityFolder -> tahunFolder -> persuratanFolder -> naskahDinasFolder
    const tahunParents = activityFolder.getParents();
    if (!tahunParents.hasNext()) return null;
    const tahunFolder = tahunParents.next();
    const tahunFolderName = tahunFolder.getName();

    const persuratanParents = tahunFolder.getParents();
    if (!persuratanParents.hasNext()) return null;
    const persuratanFolder = persuratanParents.next();

    const naskahDinasParents = persuratanFolder.getParents();
    if (!naskahDinasParents.hasNext()) return null;
    const naskahDinasFolder = naskahDinasParents.next();

    // Navigate into 2. Dokumen > Tahun XXXX > same activity folder name
    const dokumenFolder = DriveService.getOrCreateChildFolder(naskahDinasFolder, '2. Dokumen');
    const dokumenTahunFolder = DriveService.getOrCreateChildFolder(dokumenFolder, tahunFolderName);
    return DriveService.getOrCreateChildFolder(dokumenTahunFolder, activityFolderName);
  } catch (e) {
    console.error('resolveDokumenMirrorForActivity_ failed: ' + e.message);
    return null;
  }
}

/**
 * Ensure a sub-activity folder also exists inside the 2. Dokumen mirror tree.
 * parentFolderName is for grouped activities (e.g. PKN/PKA/PKP inside Kepemimpinan).
 */
function ensureDokumenMirrorSubActivity_(activityTargetFolderId, subActivityName, parentFolderName) {
  try {
    const mirrorActivityFolder = resolveDokumenMirrorForActivity_(activityTargetFolderId);
    if (!mirrorActivityFolder) return null;

    if (parentFolderName) {
      const groupFolder = DriveService.getOrCreateChildFolder(mirrorActivityFolder, parentFolderName);
      return DriveService.getOrCreateChildFolder(groupFolder, subActivityName);
    }
    return DriveService.getOrCreateChildFolder(mirrorActivityFolder, subActivityName);
  } catch (e) {
    console.error('ensureDokumenMirrorSubActivity_ failed: ' + e.message);
    return null;
  }
}

/**
 * Rename a sub-activity folder inside the 2. Dokumen mirror tree.
 */
function renameDokumenMirrorSubActivity_(activityTargetFolderId, oldName, newName, parentFolderName) {
  try {
    const mirrorActivityFolder = resolveDokumenMirrorForActivity_(activityTargetFolderId);
    if (!mirrorActivityFolder) return;

    let searchParent = mirrorActivityFolder;
    if (parentFolderName) {
      const groupFolders = mirrorActivityFolder.getFoldersByName(parentFolderName);
      if (groupFolders.hasNext()) {
        searchParent = groupFolders.next();
      } else {
        return; // group folder doesn't exist, nothing to rename
      }
    }

    const folders = searchParent.getFoldersByName(oldName);
    if (folders.hasNext()) {
      folders.next().setName(newName);
    }
  } catch (e) {
    console.error('renameDokumenMirrorSubActivity_ failed: ' + e.message);
  }
}

/**
 * Resolve the "2. Dokumen" mirror folder for any folder under "1. Persuratan".
 * Read-only: returns null if mirror path doesn't exist.
 */
function resolveMirrorForFolder_(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const pathParts = [];
    let current = folder;
    let naskahDinasFolder = null;

    while (current) {
      const parents = current.getParents();
      if (!parents.hasNext()) break;
      const parent = parents.next();
      const pName = parent.getName().toLowerCase();

      if (pName.indexOf('persuratan') >= 0 || pName.indexOf('1. persuratan') >= 0) {
        const ndParents = parent.getParents();
        if (ndParents.hasNext()) {
          naskahDinasFolder = ndParents.next();
        }
        pathParts.unshift(current.getName());
        break;
      }

      pathParts.unshift(current.getName());
      current = parent;
    }

    if (!naskahDinasFolder || pathParts.length === 0) return null;

    const dokumenFolders = naskahDinasFolder.getFoldersByName('2. Dokumen');
    if (!dokumenFolders.hasNext()) return null;
    let mirror = dokumenFolders.next();

    for (let i = 0; i < pathParts.length; i++) {
      const subFolders = mirror.getFoldersByName(pathParts[i]);
      if (!subFolders.hasNext()) return null;
      mirror = subFolders.next();
    }
    return mirror;
  } catch (e) {
    console.error('resolveMirrorForFolder_ failed: ' + e.message);
    return null;
  }
}

function validateDriveItemName_(name, label) {
  const cleanName = String(name || '').trim();
  if (!cleanName) throw new Error((label || 'Nama') + ' wajib diisi.');
  if (cleanName.length > MAX_FILENAME_LENGTH) throw new Error((label || 'Nama') + ' maksimal ' + MAX_FILENAME_LENGTH + ' karakter.');
  if (/[\\/:*?"<>|]/.test(cleanName)) {
    throw new Error((label || 'Nama') + ' tidak boleh berisi karakter \\ / : * ? " < > |.');
  }
  return cleanName;
}

/**
 * Escape a value for safe inclusion in a Drive API v2 `q` query string.
 * Escapes backslashes and single quotes to prevent query injection.
 */
function escapeDriveQueryValue_(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function extractDriveTargetId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const patterns = [
    /\/d\/([-\w]{25,})/i,
    /\/folders\/([-\w]{25,})/i,
    /[?&]id=([-\w]{25,})/i
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match) return match[1];
  }

  const fallback = text.match(/[-\w]{25,}/);
  return fallback ? fallback[0] : '';
}

function resolveArchiveDocumentShortcutTarget_(url) {
  let targetId = extractDriveTargetId_(url);
  if (!targetId) throw new Error('Link Google Drive tidak valid.');

  for (let depth = 0; depth < 3; depth++) {
    const target = Drive.Files.get(targetId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,shortcutDetails'
    });

    if (target.mimeType !== 'application/vnd.google-apps.shortcut') {
      return target;
    }

    if (!target.shortcutDetails || !target.shortcutDetails.targetId) {
      throw new Error('Gagal membuat link dokumen: Link mengarah ke shortcut Drive yang target aslinya tidak dapat dibaca.');
    }
    targetId = target.shortcutDetails.targetId;
  }

  throw new Error('Gagal membuat link dokumen: Link Drive berisi rangkaian shortcut terlalu dalam. Salin link file atau folder asli, bukan shortcut.');
}

function buildArchiveDocumentShortcutError_(error, target) {
  const message = String(error && error.message ? error.message : error || '');
  if (message.indexOf('Gagal membuat link dokumen:') === 0 || message.indexOf('Link Google Drive tidak valid.') === 0) {
    return new Error(message);
  }

  const targetLabel = target && target.name ? '"' + target.name + '"' : 'target link';
  const targetMime = target && target.mimeType ? ' (' + target.mimeType + ')' : '';
  if (/shortcutTargetInvalid|not an allowed shortcut target type/i.test(message)) {
    return new Error(
      'Gagal membuat link dokumen: Link mengarah ke ' + targetLabel + targetMime +
      ', tetapi tipe item ini tidak bisa dijadikan shortcut Drive. Salin link file atau folder asli di Google Drive.'
    );
  }
  if (/not found|File not found|insufficient|permission|access/i.test(message)) {
    return new Error('Gagal membuat link dokumen: Target link tidak dapat dibaca oleh akun ini. Pastikan file/folder asli masih ada dan akun Anda memiliki akses.');
  }
  return new Error('Gagal membuat link dokumen: Google Drive menolak pembuatan shortcut. Detail: ' + message);
}

function getUniqueFileName_(folder, requestedName) {
  const cleanName = String(requestedName || 'Dokumen Arsip.pdf').trim() || 'Dokumen Arsip.pdf';
  if (!fileNameExistsInFolder_(folder, cleanName)) return cleanName;

  const match = cleanName.match(/^(.*?)(\.[^.]+)?$/);
  const base = (match && match[1] ? match[1] : cleanName).trim();
  const ext = match && match[2] ? match[2] : '';
  let counter = 2;
  let candidate = base + ' (' + counter + ')' + ext;
  while (fileNameExistsInFolder_(folder, candidate)) {
    counter++;
    candidate = base + ' (' + counter + ')' + ext;
  }
  return candidate;
}

function fileNameExistsInFolder_(folder, name) {
  const files = folder.getFilesByName(name);
  while (files.hasNext()) {
    const file = files.next();
    if (!file.isTrashed()) return true;
  }
  return false;
}

function getFolderByPickerId_(folderId) {
  if (!folderId || String(folderId) === 'root') return DriveApp.getRootFolder();
  return DriveApp.getFolderById(cleanId_(folderId));
}

function normalizeDrivePageSize_(value, fallback, max) {
  const parsed = Number(value || fallback);
  if (!parsed || isNaN(parsed)) return fallback;
  return Math.max(10, Math.min(Math.floor(parsed), max || fallback));
}

function driveFileResourceToDto_(item) {
  item = item || {};
  const mimeType = item.mimeType || '';
  const id = item.id || '';
  const modified = item.modifiedTime ? new Date(item.modifiedTime).getTime() : 0;
  const created = item.createdTime ? new Date(item.createdTime).getTime() : 0;
  return {
    id: id,
    name: item.name || id,
    url: item.webViewLink || (mimeType === DRIVE_FOLDER_MIME_TYPE
      ? 'https://drive.google.com/drive/folders/' + id
      : 'https://drive.google.com/file/d/' + id + '/view'),
    downloadUrl: mimeType === DRIVE_FOLDER_MIME_TYPE ? '' : DriveService.getDownloadUrl(id, mimeType),
    mimeType: mimeType,
    size: Number(item.size || 0),
    created: created,
    updated: modified || created
  };
}

function getFirstParentFolder_(folder) {
  try {
    const parents = folder.getParents();
    if (!parents.hasNext()) return null;
    return DriveService.folderToDto(parents.next());
  } catch (error) {
    console.error('getFirstParentFolder_ failed: ' + error.message);
    return null;
  }
}

function chooseInboxSubfolder_(name) {
  if (/\.docx?$/i.test(name)) return 'DOCX';
  if (/\.pdf$/i.test(name)) return 'PDF';
  return 'Lainnya';
}

function dataUrlToBlob_(dataUrl, name, mimeType) {
  const parts = String(dataUrl).split(',');
  if (parts.length < 2 || !parts[1]) throw new Error('Format file upload tidak valid.');
  const header = parts[0];
  const contentType = mimeType || (header.match(/data:(.*?);base64/) || [])[1] || MimeType.BINARY;
  const bytes = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, contentType, name);
}
