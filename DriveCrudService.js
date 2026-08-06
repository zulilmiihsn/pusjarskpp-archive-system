'use strict';

/** @private Drive file/folder CRUD and listing operations. */
const DriveCrudImpl_ = {
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

};
