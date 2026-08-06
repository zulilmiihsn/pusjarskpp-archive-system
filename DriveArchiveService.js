'use strict';

/** @private Archive placement and shortcut operations. */
const DriveArchiveImpl_ = {
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

};
