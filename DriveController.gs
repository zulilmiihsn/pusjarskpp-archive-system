'use strict';

const DriveController = {
  renameDriveItem: function (payload) {
    payload = payload || {};
    const actor = requireAdmin_(payload);
    requireWithinWorkspace_(payload.itemId, payload.year);
    const r = SettingsController.renameDriveItem(payload);
    bumpVersion();
    auditAction_(actor, 'DRIVE_ITEM_RENAMED', { folderId: payload.itemId, message: 'Mengganti nama item Drive: ' + (payload.name || '-') });
    return r;
  },
  listDriveFolders: function (payload) { requireAuth_(payload || {}); return SettingsController.listDriveFolders(payload); },
  listArchiveFolder: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    requireWithinWorkspace_(payload.folderId, payload.year);
    const result = DriveService.listFolderContent(payload);
    try {
      const mirrorFolder = resolveMirrorForFolder_(payload.folderId);
      if (mirrorFolder) {
        const mirrorPayload = {
          folderId: mirrorFolder.getId(),
          pageSize: payload.pageSize,
          pageToken: payload.pageToken
        };
        const mirrorContent = DriveService.listFolderContent(mirrorPayload);
        const shortcutFiles = (mirrorContent.files || []).filter(function (f) {
          return f.mimeType === 'application/vnd.google-apps.shortcut';
        });
        result.files = (result.files || []).concat(shortcutFiles);
      }
    } catch (e) {
      console.error('Failed to load mirror dokumen: ' + e.message);
    }
    return result;
  },
  addArchiveChildFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.parentFolderId, 'Folder induk');
    Validator.requireString(payload.name, 'Nama folder');
    requireWithinWorkspace_(payload.parentFolderId, payload.year);
    const r = DriveService.createChildFolder(payload.parentFolderId, payload.name);
    bumpVersion();
    auditAction_(actor, 'FOLDER_CREATED', { folderId: payload.parentFolderId, message: 'Membuat folder turunan: ' + payload.name });
    return r;
  },
  renameArchiveFolder: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    Validator.requireString(payload.name, 'Nama baru');
    requireWithinWorkspace_(payload.folderId, payload.year);
    const r = DriveService.renameFolder(payload.folderId, payload.name);
    bumpVersion();
    auditAction_(actor, 'FOLDER_RENAMED', { folderId: payload.folderId, message: 'Mengganti nama folder menjadi: ' + payload.name });
    return r;
  },
  trashArchiveFolder: function (payload) {
    payload = payload || {};
    // Trash folder bisa menyeret arsip final di dalamnya → admin-only.
    const actor = requireAdmin_(payload);
    Validator.requireId(payload.folderId, 'Folder ID');
    requireWithinWorkspace_(payload.folderId, payload.year);
    // B3: cascade — hapus baris detail/rekap + log untuk SEMUA file ber-log di dalam
    // folder SEBELUM folder di-trash, agar tidak meninggalkan baris/log yatim yang
    // menunjuk file yang sudah masuk Tempat Sampah.
    const cascadeCount = cascadeTrashLoggedArchivesInFolder_(payload.folderId, payload.year);
    const r = DriveService.trashFolder(payload.folderId);
    bumpVersion();
    auditAction_(actor, 'FOLDER_TRASHED', { folderId: payload.folderId, message: 'Memindahkan folder ke Tempat Sampah (arsip terkait dibersihkan: ' + cascadeCount + ')' });
    return r;
  },
  renameArchiveFile: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    Validator.requireString(payload.name, 'Nama baru');
    requireWithinWorkspace_(payload.fileId, payload.year);
    const r = DriveService.renameFile(payload.fileId, payload.name);
    // B7: jaga sinkron — bila file ini arsip ber-log, perbarui nama di archive_log
    // agar tampilan & proses re-adopsi tidak memakai nama lama yang basi.
    try {
      const log = ConfigRepository.getArchiveLogByFileId(payload.fileId);
      if (log && log.archive_id) {
        ConfigRepository.updateArchiveLog(log.archive_id, { final_file_name: (r && r.name) || payload.name });
      }
    } catch (syncError) {
      console.warn('renameArchiveFile: gagal sinkron nama ke archive_log: ' + syncError.message);
    }
    bumpVersion();
    auditAction_(actor, 'FILE_RENAMED', { folderId: payload.fileId, message: 'Mengganti nama file menjadi: ' + payload.name });
    return r;
  },
  trashArchiveFile: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    requireWithinWorkspace_(payload.fileId, payload.year);
    const log = ConfigRepository.getArchiveLogByFileId(payload.fileId);
    if (log && log.archive_id) {
       // Menghapus arsip final = admin-only. Escalate di sini, lalu pakai core
       // (yang sudah atomik: hapus baris sheet + log + trash file).
       requireAdmin_(payload);
       ArchiveController._deleteArchiveCore_({ archiveId: log.archive_id, year: payload.year });
       auditAction_(actor, 'ARCHIVE_DELETED', {
         year: payload.year || log.year, activityId: log.activity_id, subActivityId: log.sub_activity_id,
         folderId: log.final_file_id || payload.fileId,
         message: 'Menghapus surat: ' + (log.final_file_name || '-')
       });
    } else {
       let fileName = '';
       try {
         const file = DriveApp.getFileById(cleanId_(payload.fileId));
         fileName = file.getName();
       } catch (e) {
         console.warn('Failed to read file name before trash: ' + e.message);
       }
       DriveService.trashFile(payload.fileId);
       if (payload.activityId && payload.subActivityId && fileName) {
         try {
           const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
           const config = CacheHelper.getConfig(year);
           const activity = ConfigService.findActivity(config, payload.activityId);
           const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
           if (activity && subActivity) {
             SpreadsheetService.clearArchiveDocumentLink(activity, subActivity, fileName);
           }
         } catch (e) {
           console.error('Failed to clear archive document link: ' + e.message);
         }
       }
       auditAction_(actor, 'FILE_TRASHED', {
         year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
         folderId: payload.fileId, message: 'Menghapus file: ' + (fileName || '-')
       });
    }
    bumpVersion();
    return { success: true };
  }
};

/**
 * B3: Sebelum sebuah folder di-trash, telusuri semua file di dalamnya (rekursif,
 * batas kedalaman 8) dan hapus arsip ber-log lewat ArchiveController._deleteArchiveCore_
 * (atomik: baris detail/rekap + log + trash file). Mencegah baris/log yatim.
 * @param {string} folderId
 * @param {*} year
 * @return {number} jumlah arsip ber-log yang dibersihkan
 */
function cascadeTrashLoggedArchivesInFolder_(folderId, year) {
  let cleaned = 0;
  const rootId = cleanId_(folderId);
  if (!rootId) return cleaned;
  const stack = [{ id: rootId, depth: 0 }];
  while (stack.length) {
    const cur = stack.pop();
    if (cur.depth > 8) continue;
    let pageToken = null;
    do {
      let result;
      try {
        result = Drive.Files.list({
          q: "'" + cur.id + "' in parents and trashed = false",
          fields: 'nextPageToken, files(id, mimeType)',
          pageSize: 1000,
          pageToken: pageToken,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true
        });
      } catch (e) {
        break;
      }
      const items = result.files || [];
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.mimeType === 'application/vnd.google-apps.folder') {
          stack.push({ id: item.id, depth: cur.depth + 1 });
          continue;
        }
        try {
          const log = ConfigRepository.getArchiveLogByFileId(item.id);
          if (log && log.archive_id) {
            ArchiveController._deleteArchiveCore_({ archiveId: log.archive_id, year: year || log.year });
            cleaned++;
          }
        } catch (e) {
          console.warn('cascadeTrashLoggedArchivesInFolder_: gagal membersihkan arsip untuk file ' + item.id + ': ' + e.message);
        }
      }
      pageToken = result.nextPageToken;
    } while (pageToken);
  }
  return cleaned;
}
