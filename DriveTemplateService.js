'use strict';

/** @private Template workspace operations. */
const DriveTemplateImpl_ = {
  listTemplates: function (year, preloadedConfig) {
    try {
      const config = preloadedConfig || CacheHelper.getConfig(year);
      const yearConfig = ConfigService.getYearConfig(config, year);
      if (!yearConfig.template_folder_id) return [];
      const folderId = cleanId_(yearConfig.template_folder_id);
      
      const result = [];
      let pageToken = null;
      do {
        let res;
        try {
          res = Drive.Files.list({
            q: "'" + folderId + "' in parents and trashed = false",
            fields: "nextPageToken, files(id, name, mimeType, webViewLink)",
            pageSize: 1000,
            pageToken: pageToken,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
          });
        } catch (e) {
          break;
        }
        
        const items = res.files || [];
        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          result.push({
            id: item.id,
            name: item.name,
            mimeType: item.mimeType,
            url: item.webViewLink,
            downloadUrl: DriveService.getDownloadUrl(item.id, item.mimeType)
          });
        }
        pageToken = res.nextPageToken;
      } while (pageToken);
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
      payload = payload || {};
      const safeName = validateFilePayload_(payload);
      const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      let targetFolder = this.getTemplateFolder_(year);
      if (payload.categoryFolderId) {
        requireWithinTemplateWorkspace_(payload.categoryFolderId, year);
        targetFolder = DriveApp.getFolderById(cleanId_(payload.categoryFolderId));
      }
      return this.fileToDto(targetFolder.createFile(dataUrlToBlob_(payload.dataUrl, safeName, payload.mimeType)));
    });
  },

  trashTemplateFile: function (fileId, year) {
    if (!fileId) throw new Error('File ID wajib diisi.');
    const scope = requireWithinTemplateWorkspace_(fileId, year);
    DriveApp.getFileById(cleanId_(fileId)).setTrashed(true);
    return { id: fileId, year: scope.year, trashed: true };
  },

  createTemplateCategory: function (year, name) {
    const root = this.getTemplateFolder_(year);
    const folder = this.getOrCreateChildFolder(root, name);
    return this.folderToDto(folder);
  },

  renameTemplateCategoryFolder: function (folderId, newName, year) {
    requireWithinTemplateWorkspace_(folderId, year);
    return this.renameFolder(folderId, newName);
  },

  deleteTemplateCategoryFolder: function (folderId, year) {
    requireWithinTemplateWorkspace_(folderId, year);
    return this.trashFolder(folderId);
  },

  initTemplateResumableUpload: function (payload) {
    payload = payload || {};
    const safeName = validateDriveItemName_(payload.name, 'Nama file');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const mimeType = payload.mimeType || 'application/octet-stream';
    const totalSize = Number(payload.totalSize);
    if (!totalSize || totalSize < 1) throw new Error('Ukuran file tidak valid.');

    let targetFolder = this.getTemplateFolder_(year);
    if (payload.categoryFolderId) {
      requireWithinTemplateWorkspace_(payload.categoryFolderId, year);
      targetFolder = DriveApp.getFolderById(cleanId_(payload.categoryFolderId));
    }
    return resumableUploadInit_(safeName, mimeType, totalSize, targetFolder.getId());
  },

  uploadResumableChunk: function (payload) {
    payload = payload || {};
    const sessionToken = payload.sessionUrl;
    const chunkBase64 = payload.chunk;
    const startByte = Number(payload.startByte);
    const totalSize = Number(payload.totalSize);
    if (!sessionToken || !chunkBase64) throw new Error('Parameter upload tidak lengkap.');
    return resumableUploadChunk_(sessionToken, chunkBase64, startByte, totalSize);
  }
};
