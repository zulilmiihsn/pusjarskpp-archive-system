'use strict';

const SubActivityController = {
  addSubActivity: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityName, 'Nama sub-kegiatan');
    sanitizeNameForSheetAndDrive_(payload.subActivityName, 'Nama sub-kegiatan');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');

      let subActivityName = String(payload.subActivityName || '').trim();
      const existing = config.subActivities.filter(function (sub) { return sub.activity_id === activity.activity_id; });
      const nextIndex = existing.length + 1;

      if (activity.activity_id === 'latsar_cpns') {
        if (!subActivityName.toLowerCase().startsWith('angkatan') && !subActivityName.toLowerCase().startsWith('latsar')) {
          subActivityName = 'Angkatan ' + subActivityName;
        }
      } else if (activity.activity_id === 'kepemimpinan') {
        subActivityName = subActivityName;
      } else {
        if (!/^\d+\.\s*/.test(subActivityName)) {
          subActivityName = nextIndex + '. ' + subActivityName;
        }
      }

      const targetFolder = DriveApp.getFolderById(activity.target_folder_id);
      const parentFolder = resolveSubActivityParentFolder_(targetFolder, activity, subActivityName);
      const folder = DriveService.createSubFolder(parentFolder ? parentFolder.getId() : activity.target_folder_id, subActivityName);
      const parentFolderName = parentFolder ? parentFolder.getName() : '';
      const spreadsheet = wsResolveSpreadsheetForSubActivity_(activity, subActivityName, year, null, parentFolderName);
      const formalArchiveName = payload.formalArchiveName || wsSuggestFormalArchiveName_(activity, subActivityName);
      const targetSheetName = payload.targetSheetName || wsSuggestDetailSheetName_(activity, subActivityName, spreadsheet);
      let subRow = ConfigRepository.addSubActivityRow({
        year: year,
        activityId: activity.activity_id,
        subActivityId: parentFolderName ? slug_(parentFolderName + '_' + subActivityName) : slug_(subActivityName),
        subActivityName: subActivityName,
        formalArchiveName: formalArchiveName,
        noFolder: payload.noFolder || WorkspaceSetupService.parseLeadingNumber(subActivityName) || activity.folder_no,
        defaultKodeKlasifikasi: payload.defaultKodeKlasifikasi || DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI,
        allowNonLetterDocument: payload.allowNonLetterDocument,
        targetSheetName: targetSheetName,
        mappingStatus: payload.mappingStatus || wsInferMappingStatus_(subActivityName, formalArchiveName, targetSheetName),
        mappingNote: payload.mappingNote,
        folderPath: parentFolderName ? parentFolderName + ' > ' + folder.getName() : folder.getName(),
        parentFolderId: parentFolder ? cleanId_(parentFolder.getId()) : '',
        laciFolderId: payload.laciFolderId ? cleanId_(payload.laciFolderId) : undefined,
        parentFolderName: parentFolderName,
        parentFolderPath: parentFolderName,
        spreadsheetFileId: spreadsheet ? spreadsheet.getId() : activity.spreadsheet_file_id
      }, folder);

      if (subRow && subRow._shiftedSubActivities && subRow._shiftedSubActivities.length > 0) {
        SpreadsheetService.cascadeNomorBerkasShift(year, subRow._shiftedSubActivities);
      }

      SpreadsheetService.ensureSubActivitySheet(activity, subRow);
      const rekap = subRow.rekap_row_number
        ? SpreadsheetService.updateRekapSubActivityIdentity(activity, null, subRow)
        : SpreadsheetService.appendRekapRowIfPresent(activity, subRow);
      subRow = persistRekapRowNumber_(year, activity, subRow, rekap) || subRow;

      // Mirror folder to 2. Dokumen
      ensureDokumenMirrorSubActivity_(activity.target_folder_id, subActivityName, parentFolderName);

      CacheHelper.invalidate(year);
      return {
        subActivity: subRow,
        newSubId: subRow.sub_activity_id,
        folder: DriveService.folderToDto(folder),
        rekap: rekap,
        bootstrap: WorkspaceController.getBootstrap()
      };
    }, 30000);
  },
  createParentFolder: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.folderName, 'Nama Folder Utama');
    sanitizeNameForSheetAndDrive_(payload.folderName, 'Nama Folder Utama');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');

      const targetFolder = DriveApp.getFolderById(activity.target_folder_id);
      const folderName = String(payload.folderName).trim();
      
      // Buat folder induk di Persuratan
      const newParentFolder = DriveService.createSubFolder(targetFolder.getId(), folderName);
      
      // Temukan docTargetFolder untuk Dokumen
      const persuratan = DriveApp.getFolderById(CacheHelper.getConfig(year).years[0].persuratan_year_folder_id).getParents().next();
      const naskahDinas = persuratan.getParents().next();
      let docTargetFolder = null;
      try {
        const docRoot = naskahDinas.getFoldersByName("2. Dokumen").next();
        const docYear = docRoot.getFoldersByName("Tahun " + year).next();
        docTargetFolder = docYear.getFoldersByName(targetFolder.getName()).next();
      } catch (e) { }

      let docNewParentFolder = null;
      if (docTargetFolder) {
        docNewParentFolder = DriveService.createSubFolder(docTargetFolder.getId(), folderName);
      }

      const subActivityIds = payload.subActivityIds || [];
      const subsToMove = config.subActivities.filter(function(s) { 
        return s.activity_id === activity.activity_id && subActivityIds.indexOf(s.sub_activity_id) !== -1; 
      });

      subsToMove.forEach(function(sub) {
        if (sub.folder_id) {
          try {
            const subFolder = DriveApp.getFolderById(sub.folder_id);
            subFolder.moveTo(newParentFolder);
          } catch(e) { }
        }
        if (docTargetFolder && docNewParentFolder) {
          try {
            const docSubFolder = docTargetFolder.getFoldersByName(sub.sub_activity_name).next();
            docSubFolder.moveTo(docNewParentFolder);
          } catch(e) { }
        }
      });

      // Teruskan _sessionId: syncSubActivities kini meminta sesi; ini panggilan internal.
      return this.syncSubActivities({ year: year, activityId: activity.activity_id, _sessionId: payload._sessionId });
    }, 30000);
  },

  convertSubActivityToParent: function(payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');
    Validator.requireString(payload.newSubActivityName, 'Nama Anak Baru');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');

      const existingSub = config.subActivities.find(function (sub) { return sub.sub_activity_id === payload.subActivityId; });
      if (!existingSub) throw new Error('Sub-kegiatan tidak ditemukan.');
      if (!existingSub.folder_id) throw new Error('Sub-kegiatan belum punya folder Drive.');

      const newSubName = String(payload.newSubActivityName).trim();
      const parentFolderName = existingSub.sub_activity_name;
      const parentFolderId = existingSub.folder_id;

      // 1. Buat folder anak baru di dalam folder lama (di Persuratan)
      const newChildFolder = DriveService.createSubFolder(parentFolderId, newSubName);
      const newChildFolderId = newChildFolder.getId();

      const moveFilesToNewFolder_ = (oldParentId, newParentId) => {
        let pageToken = null;
        do {
          let result;
          try {
            result = Drive.Files.list({
              q: "'" + oldParentId + "' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false",
              fields: "nextPageToken, files(id)",
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
            try {
              Drive.Files.update({}, items[i].id, null, {
                addParents: newParentId,
                removeParents: oldParentId,
                supportsAllDrives: true
              });
            } catch (e) {
              console.warn('Failed to move file ' + items[i].id + ': ' + e.message);
            }
          }
          pageToken = result.nextPageToken;
        } while (pageToken);
      };

      // Pindahkan semua file yang ada di parentFolderId ke dalam newChildFolderId
      // (Kita tidak memindahkan folder, hanya file agar struktur aman)
      moveFilesToNewFolder_(parentFolderId, newChildFolderId);

      // 2. Buat folder anak baru di Dokumen
      let docParentFolderId = null;
      try {
        const persuratan = DriveApp.getFolderById(CacheHelper.getConfig(year).years[0].persuratan_year_folder_id).getParents().next();
        const naskahDinas = persuratan.getParents().next();
        const docRoot = naskahDinas.getFoldersByName("2. Dokumen").next();
        const docYear = docRoot.getFoldersByName("Tahun " + year).next();
        const docTargetFolder = docYear.getFoldersByName(activity.targetFolderName || activity.activity_name).next();
        
        let docParentFolder = null;
        if (existingSub.parent_folder_name) {
          const docGroup = docTargetFolder.getFoldersByName(existingSub.parent_folder_name).next();
          docParentFolder = docGroup.getFoldersByName(parentFolderName).next();
        } else {
          docParentFolder = docTargetFolder.getFoldersByName(parentFolderName).next();
        }
        
        if (docParentFolder) {
          docParentFolderId = docParentFolder.getId();
          const docNewChildFolder = DriveService.createSubFolder(docParentFolderId, newSubName);
          moveFilesToNewFolder_(docParentFolderId, docNewChildFolder.getId());
        }
      } catch (e) {
        // Abaikan jika struktur dokumen tidak standar
      }

      // 3. Rename sheet Detail dulu (bawa semua baris data) dan TANGKAP hasilnya.
      //    Kalau rename gagal (mis. nama tujuan sudah dipakai / sheet lama tak ada),
      //    config harus tetap menunjuk ke sheet tempat data sebenarnya berada, bukan
      //    ke nama baru yang kosong — kalau tidak, data lama jadi orphan & numbering reset.
      const oldSheetName = existingSub.target_sheet_name || existingSub.sub_activity_name;
      const renameOk = SpreadsheetService.renameSubActivitySheet(
        activity, oldSheetName, newSubName, { spreadsheet_file_id: existingSub.spreadsheet_file_id }
      );
      const effectiveSheetName = renameOk ? newSubName : oldSheetName;

      // 4. Update konfigurasi (Transfer Ownership) — target_sheet_name ikut hasil rename.
      const updatedSub = ConfigRepository.updateSubActivityMapping({
        year: year,
        activityId: activity.activity_id,
        subActivityId: existingSub.sub_activity_id,
        folderId: newChildFolder.getId(),
        folderPath: existingSub.parent_folder_name
          ? existingSub.parent_folder_name + ' > ' + parentFolderName + ' > ' + newSubName
          : parentFolderName + ' > ' + newSubName,
        targetSheetName: effectiveSheetName,
        formalArchiveName: wsSuggestFormalArchiveName_(activity, newSubName),
        subActivityName: newSubName,
        parentFolderId: parentFolderId,
        parentFolderName: parentFolderName,
        parentFolderPath: existingSub.parent_folder_name
          ? existingSub.parent_folder_name + ' > ' + parentFolderName
          : parentFolderName,
        spreadsheetFileId: existingSub.spreadsheet_file_id
      });

      // 5. Pastikan sheet Detail ada (kalau rename gagal/old belum ada) + segarkan identitas baris Rekap.
      SpreadsheetService.ensureSubActivitySheet(activity, updatedSub);
      SpreadsheetService.updateRekapSubActivityIdentity(activity, existingSub, updatedSub);

      CacheHelper.invalidate(year);
      return { success: true, bootstrap: WorkspaceController.getBootstrap() };
    }, 30000);
  },

  syncSubActivities: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');

      const targetFolder = DriveApp.getFolderById(activity.target_folder_id);
      const subFolderEntries = buildSyncSubActivityEntries_(targetFolder, activity);
      const liveFolderIds = subFolderEntries.map(function (entry) { return entry.folder.getId(); });
      const entryByFolderId = {};
      subFolderEntries.forEach(function (entry) { entryByFolderId[entry.folder.getId()] = entry; });

      const existingSubs = ConfigRepository.getSubActivitiesForSync(year, activity.activity_id);

      let updated = _syncResolveSpreadsheet_(activity, year);

      // Nesting: sub-kegiatan yang foldernya kini punya anak (jadi container) dimigrasi turun
      // ke leaf ahli warisnya — sheet Detail + baris Rekap + data ikut, tanpa bikin duplikat.
      const migratedHeirIds = {};
      if (_syncMigrateParents_(activity, year, existingSubs, liveFolderIds, entryByFolderId, migratedHeirIds)) updated = true;

      subFolderEntries.forEach(function (entry) {
        if (migratedHeirIds[entry.folder.getId()]) return; // sudah ditangani migrasi
        if (_syncProcessEntry_(entry, activity, year, existingSubs)) updated = true;
      });

      if (_syncDeactivateMissing_(year, activity, existingSubs, liveFolderIds, migratedHeirIds)) updated = true;

      if (updated) CacheHelper.invalidate(year);
      return { updated: updated, bootstrap: WorkspaceController.getBootstrap() };
    }, 30000);
  },

  deleteSubActivity: function (payload) {
    payload = payload || {};
    requireAdmin_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');

      const existingSub = config.subActivities.find(function (sub) { return sub.sub_activity_id === payload.subActivityId; });
      if (!existingSub) throw new Error('Sub-kegiatan tidak ditemukan.');

      const removed = ConfigRepository.deactivateSubActivity(year, activity.activity_id, existingSub.sub_activity_id, SUB_ACTIVITY_INACTIVE_REASON.MANUAL);
      if (removed > 0) {
        SpreadsheetService.markRekapSubActivityInactive(activity, existingSub, SUB_ACTIVITY_INACTIVE_REASON.MANUAL);
      }
      CacheHelper.invalidate(year);
      return { success: removed > 0, bootstrap: WorkspaceController.getBootstrap() };
    }, 30000);
  },

  trashSubActivityFolder: function (payload) {
    payload = payload || {};
    requireAdmin_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');

      const existingSub = ConfigRepository.getSubActivitiesForSync(year, activity.activity_id).find(function (sub) { return sub.sub_activity_id === payload.subActivityId; });
      if (!existingSub) throw new Error('Sub-kegiatan tidak ditemukan.');
      if (!existingSub.folder_id) throw new Error('Sub-kegiatan belum punya folder Drive.');

      try {
        DriveApp.getFolderById(existingSub.folder_id).setTrashed(true);
      } catch (error) {
        CacheHelper.invalidate(year);
        throw new Error('Folder tidak bisa dipindahkan ke Bin: ' + error.message);
      }

      const removed = ConfigRepository.deactivateSubActivity(year, activity.activity_id, existingSub.sub_activity_id, SUB_ACTIVITY_INACTIVE_REASON.DRIVE_TRASHED);
      if (removed > 0) {
        SpreadsheetService.markRekapSubActivityInactive(activity, existingSub, SUB_ACTIVITY_INACTIVE_REASON.DRIVE_TRASHED);
      }
      CacheHelper.invalidate(year);
      return { success: removed > 0, folderId: existingSub.folder_id, bootstrap: WorkspaceController.getBootstrap() };
    }, 30000);
  },

  cleanupTrashedSubActivities: function (payload) {
    payload = payload || {};
    return ConfigRepository.purgeInactiveSubActivities({
      reason: SUB_ACTIVITY_INACTIVE_REASON.DRIVE_TRASHED,
      retentionDays: payload.retentionDays || TRASHED_SUB_ACTIVITY_RETENTION_DAYS
    });
  },

  renameSubActivity: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');
    Validator.requireString(payload.newName, 'Nama baru');
    sanitizeNameForSheetAndDrive_(payload.newName, 'Nama baru');

    return withLock_(() => {
    const config = CacheHelper.getConfig(year);
    const activity = ConfigService.findActivity(config, payload.activityId);
    if (!activity) throw new Error('Kegiatan tidak ditemukan.');

    const newName = String(payload.newName || '').trim();
    const existingSub = config.subActivities.find(function (sub) { return sub.sub_activity_id === payload.subActivityId; });
    if (!existingSub) throw new Error('Sub-kegiatan tidak ditemukan.');
    const previousSub = Object.assign({}, existingSub);

    let targetSheetName = existingSub.target_sheet_name;
    let formalArchiveName = existingSub.formal_archive_name;
    let noFolder = existingSub.no_folder;
    let rekapRowNumber = existingSub.rekap_row_number;
    let mappingStatus = existingSub.mapping_status;

    if (payload.targetSheetName !== undefined) {
      targetSheetName = String(payload.targetSheetName || '').trim();
    }
    if (payload.formalArchiveName !== undefined) {
      formalArchiveName = String(payload.formalArchiveName || '').trim();
    }
    if (payload.noFolder !== undefined) {
      noFolder = String(payload.noFolder || '').trim();
    }
    if (payload.rekapRowNumber !== undefined) {
      rekapRowNumber = String(payload.rekapRowNumber || '').trim();
    }
    if (payload.mappingStatus !== undefined) {
      mappingStatus = String(payload.mappingStatus || '').trim();
    }

    if (existingSub.folder_id) {
      try {
        const oldFolderName = DriveApp.getFolderById(existingSub.folder_id).getName();
        DriveApp.getFolderById(existingSub.folder_id).setName(newName);
        if (!targetSheetName || targetSheetName === oldFolderName) {
          targetSheetName = wsSuggestDetailSheetName_(activity, newName, wsResolveSpreadsheetForSubActivity_(activity, newName, year, null, existingSub.parent_folder_name));
        }
        if (!formalArchiveName || formalArchiveName === oldFolderName) {
          formalArchiveName = wsSuggestFormalArchiveName_(activity, newName);
        }

        // Mirror rename to 2. Dokumen
        renameDokumenMirrorSubActivity_(activity.target_folder_id, oldFolderName, newName, existingSub.parent_folder_name);
      } catch (e) {
        throw new Error('Gagal mengganti nama folder di Google Drive: ' + e.message);
      }
    }

    if (targetSheetName !== existingSub.target_sheet_name) {
      const renameSpreadsheet = wsResolveSpreadsheetForSubActivity_(activity, newName, year, null, existingSub.parent_folder_name);
      if (renameSpreadsheet && renameSpreadsheet.getId() !== existingSub.spreadsheet_file_id) {
        existingSub.spreadsheet_file_id = renameSpreadsheet.getId();
      }
      SpreadsheetService.renameSubActivitySheet(activity, existingSub.target_sheet_name, targetSheetName, existingSub);
    }

    const updatedSub = ConfigRepository.updateSubActivityMapping({
      year: year, activityId: activity.activity_id, subActivityId: existingSub.sub_activity_id,
      folderId: existingSub.folder_id,
      folderPath: existingSub.parent_folder_name ? existingSub.parent_folder_name + ' > ' + newName : newName,
      targetSheetName: targetSheetName,
      formalArchiveName: formalArchiveName,
      noFolder: noFolder,
      rekapRowNumber: rekapRowNumber,
      mappingStatus: mappingStatus || wsInferMappingStatus_(newName, formalArchiveName, targetSheetName),
      subActivityName: newName,
      parentFolderId: existingSub.parent_folder_id,
      parentFolderName: existingSub.parent_folder_name,
      parentFolderPath: existingSub.parent_folder_path,
      spreadsheetFileId: existingSub.spreadsheet_file_id
    });
    SpreadsheetService.updateRekapSubActivityIdentity(activity, previousSub, updatedSub);

    CacheHelper.invalidate(year);
    return { success: true, bootstrap: WorkspaceController.getBootstrap() };
    }, 30000);
  },

  getInactiveSubActivities: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    return ConfigRepository.getInactiveSubActivities(year);
  },

  restoreSubActivity: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');

    const config = CacheHelper.getConfig(year);
    const activity = ConfigService.findActivity(config, payload.activityId);
    const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId) ||
      ConfigRepository.getSubActivitiesForSync(year, payload.activityId).find(function (sub) {
        return sub.sub_activity_id === payload.subActivityId;
      });
    const result = ConfigRepository.restoreSubActivity(year, payload.activityId, payload.subActivityId);
    if (result > 0 && activity && subActivity) {
      SpreadsheetService.updateRekapSubActivityIdentity(activity, subActivity, subActivity);
      SpreadsheetService.clearRekapSubActivityInactiveMark(activity, subActivity);
      CacheHelper.invalidate(year);
    } else if (result > 0) {
      CacheHelper.invalidate(year);
    }
    return { success: result > 0, bootstrap: WorkspaceController.getBootstrap() };
  },

  purgeSubActivity: function (payload) {
    payload = payload || {};
    requireAdmin_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');

    const result = ConfigRepository.purgeSubActivity(year, payload.activityId, payload.subActivityId);
    if (result > 0) CacheHelper.invalidate(year);
    return { success: result > 0, bootstrap: WorkspaceController.getBootstrap() };
  },

  updateSubActivityMetadata: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');

    return withLock_(() => {
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');
      const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
      if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');

      if (payload.metadataLocks !== undefined || payload.kodeKlasifikasi !== undefined) {
        const updates = {
          year: year,
          activityId: payload.activityId,
          subActivityId: payload.subActivityId
        };
        if (payload.metadataLocks !== undefined) updates.metadataLocks = payload.metadataLocks;
        if (payload.kodeKlasifikasi !== undefined) updates.defaultKodeKlasifikasi = payload.kodeKlasifikasi;
        
        ConfigRepository.updateSubActivityMapping(updates);
        if (payload.kodeKlasifikasi !== undefined) subActivity.default_kode_klasifikasi = payload.kodeKlasifikasi;
      }

      const rekapResult = SpreadsheetService.updateRekapDocumentMetadata(activity, subActivity, {
        nomorBerkas: payload.nomorBerkas,
        kodeKlasifikasi: payload.kodeKlasifikasi,
        kurunWaktu: payload.kurunWaktu,
        noFilingCabinet: payload.noFilingCabinet,
        noLaci: payload.noLaci,
        noFolder: payload.noFolder,
        klasifikasiAkses: payload.klasifikasiAkses,
        ket: payload.ket,
        jumlah: payload.jumlah
      }, payload.metadataLocks ? JSON.parse(payload.metadataLocks) : {});

      CacheHelper.invalidate(year);
      return { success: true, rekap: rekapResult, bootstrap: WorkspaceController.getBootstrap() };
    }, 30000);
  }
};

/**
 * Extensibility point: resolves a parent grouping folder for a new sub-activity.
 * Currently returns null (sub-activities are created directly under the target folder).
 * Override this to implement parent-folder grouping (e.g., PKN/PKA/PKP containers
 * for the kepemimpinan activity type).
 * @param {Folder} targetFolder - The activity's target folder in Drive.
 * @param {Object} activity - The activity configuration row.
 * @param {string} subActivityName - The name of the new sub-activity.
 * @returns {Folder|null} Parent folder, or null to use targetFolder directly.
 */
function resolveSubActivityParentFolder_(targetFolder, activity, subActivityName) {
  return null;
}

function buildSyncSubActivityEntries_(targetFolder, activity) {
  return wsBuildLeafSubActivityEntries_('', [], targetFolder, activity).map(function (entry) {
    return Object.assign({}, entry, {
      folderPath: String(entry.folderPath || '').replace(/^\s*>\s*/, ''),
      parentFolderPath: String(entry.parentFolderPath || '').replace(/^\s*>\s*/, '')
    });
  });
}

function persistRekapRowNumber_(year, activity, subActivity, rekapResult) {
  if (!rekapResult || !rekapResult.rowNumber || !subActivity || subActivity.rekap_row_number) return null;
  return ConfigRepository.updateSubActivityMapping({
    year: year,
    activityId: activity.activity_id,
    subActivityId: subActivity.sub_activity_id,
    folderId: subActivity.folder_id,
    folderPath: subActivity.folder_path,
    targetSheetName: subActivity.target_sheet_name,
    formalArchiveName: subActivity.formal_archive_name,
    noFolder: subActivity.no_folder,
    defaultKodeKlasifikasi: subActivity.default_kode_klasifikasi,
    mappingStatus: subActivity.mapping_status,
    mappingNote: subActivity.mapping_note,
    rekapRowNumber: rekapResult.rowNumber,
    subActivityName: subActivity.sub_activity_name,
    parentFolderId: subActivity.parent_folder_id,
    parentFolderName: subActivity.parent_folder_name,
    parentFolderPath: subActivity.parent_folder_path,
    spreadsheetFileId: subActivity.spreadsheet_file_id
  });
}

function _syncResolveSpreadsheet_(activity, year) {
  const ss = wsResolveSpreadsheetForSubActivity_(activity, '', year, null, '');
  if (ss && ss.getId() !== activity.spreadsheet_file_id) {
    ConfigRepository.updateActivityMapping({
      year: year,
      activityId: activity.activity_id,
      spreadsheetFileId: ss.getId()
    });
    activity.spreadsheet_file_id = ss.getId();
    return true;
  }
  return false;
}

function _syncDeactivateMissing_(year, activity, existingSubs, liveFolderIds, migratedHeirIds) {
  migratedHeirIds = migratedHeirIds || {};
  const deletedIds = existingSubs
    .filter(function (sub) {
      if (!sub.folder_id) return false;
      if (!isTrue_(sub.is_active)) return false;             // hanya yang masih aktif
      if (liveFolderIds.indexOf(sub.folder_id) !== -1) return false; // masih leaf
      if (migratedHeirIds[sub.folder_id]) return false;      // sudah dimigrasi ke heir
      // Nonaktifkan HANYA jika folder benar-benar hilang dari Drive (terhapus/trashed),
      // bukan sekadar berubah jadi parent (itu sudah ditangani _syncMigrateParents_).
      try {
        return DriveApp.getFolderById(sub.folder_id).isTrashed();
      } catch (e) {
        return true; // tidak ditemukan
      }
    })
    .map(function (sub) { return sub.folder_id; });
  if (deletedIds.length === 0) return false;
  const removed = ConfigRepository.deactivateSubActivityRows(
    year, activity.activity_id, deletedIds, SUB_ACTIVITY_INACTIVE_REASON.DRIVE_MISSING
  );
  existingSubs.forEach(function (sub) {
    if (deletedIds.indexOf(sub.folder_id) !== -1) {
      SpreadsheetService.markRekapSubActivityInactive(activity, sub, SUB_ACTIVITY_INACTIVE_REASON.DRIVE_MISSING);
    }
  });
  return removed > 0;
}

function _syncFindHeirEntry_(parentFolder, activity, entryByFolderId) {
  // Leaf-leaf di bawah parentFolder (urutan deterministik); ambil yang ada di daftar leaf utama.
  const leaves = wsBuildLeafSubActivityEntries_('', [], parentFolder, activity);
  for (let i = 0; i < leaves.length; i++) {
    const id = leaves[i].folder.getId();
    if (entryByFolderId[id]) return entryByFolderId[id];
  }
  return null;
}

function _syncMigrateParents_(activity, year, existingSubs, liveFolderIds, entryByFolderId, migratedHeirIds) {
  let changed = false;
  existingSubs.forEach(function (sub) {
    if (!sub.folder_id) return;
    if (!isTrue_(sub.is_active)) return;
    if (liveFolderIds.indexOf(sub.folder_id) !== -1) return; // masih leaf → jalur normal

    let folder = null;
    try { folder = DriveApp.getFolderById(sub.folder_id); } catch (e) { folder = null; }
    if (!folder || folder.isTrashed()) return;            // hilang → biar _syncDeactivateMissing_ yang urus
    if (wsListChildFolders_(folder).length === 0) return; // ada tapi bukan parent → lewati

    const heirEntry = _syncFindHeirEntry_(folder, activity, entryByFolderId);
    if (!heirEntry) {
      console.warn('Sync: folder "' + sub.sub_activity_name + '" jadi parent tapi belum ada leaf; data dibiarkan.');
      return;
    }
    const heirFolder = heirEntry.folder;
    const heirFolderId = heirFolder.getId();
    if (migratedHeirIds[heirFolderId]) return; // heir sudah diklaim parent lain pada sync ini

    const heirAlreadyTracked = existingSubs.some(function (other) {
      return other !== sub && isTrue_(other.is_active) && other.folder_id === heirFolderId;
    });
    if (heirAlreadyTracked) {
      console.warn('Sync: heir "' + heirFolder.getName() + '" sudah punya sub-kegiatan; migrasi "' + sub.sub_activity_name + '" dilewati.');
      return;
    }

    const heirName = heirFolder.getName();
    const oldSheetName = sub.target_sheet_name || sub.sub_activity_name;
    // Data tetap di workbook asal sub; cuma sheet Detail di-rename + config di-repoint ke leaf.
    const spreadsheetFileId = sub.spreadsheet_file_id;
    const newSheetName = wsSuggestDetailSheetName_(activity, heirName, null) || heirName;

    // Rename sheet Detail (bawa semua baris data; referensi formula Rekap auto-ikut saat rename).
    const renameOk = (oldSheetName && newSheetName && oldSheetName !== newSheetName)
      ? SpreadsheetService.renameSubActivitySheet(activity, oldSheetName, newSheetName, { spreadsheet_file_id: spreadsheetFileId })
      : true;
    const effectiveSheetName = renameOk ? newSheetName : oldSheetName;

    const updatedSub = ConfigRepository.updateSubActivityMapping({
      year: year, activityId: activity.activity_id, subActivityId: sub.sub_activity_id,
      folderId: heirFolderId,
      folderPath: heirEntry.folderPath,
      targetSheetName: effectiveSheetName,
      formalArchiveName: sub.formal_archive_name || wsSuggestFormalArchiveName_(activity, heirName),
      noFolder: WorkspaceSetupService.parseLeadingNumber(heirName) || sub.no_folder,
      defaultKodeKlasifikasi: sub.default_kode_klasifikasi,
      mappingStatus: sub.mapping_status,
      subActivityName: heirName,
      parentFolderId: heirEntry.parentFolderId,
      parentFolderName: heirEntry.parentFolderName,
      parentFolderPath: heirEntry.parentFolderPath,
      spreadsheetFileId: spreadsheetFileId,
      rekapRowNumber: sub.rekap_row_number
    });

    // Pastikan sheet Detail ada (kalau rename gagal/old belum ada) + segarkan identitas baris Rekap (baris sama).
    SpreadsheetService.ensureSubActivitySheet(activity, updatedSub);
    SpreadsheetService.updateRekapSubActivityIdentity(activity, sub, updatedSub);

    // Sinkronkan objek in-memory supaya pass berikutnya melihat row yang sudah dipindah.
    Object.assign(sub, updatedSub);
    migratedHeirIds[heirFolderId] = true;
    changed = true;
    console.info('Sync: migrasi "' + oldSheetName + '" → leaf "' + heirName + '" (data dipertahankan).');
  });
  return changed;
}

function _syncProcessEntry_(entry, activity, year, existingSubs) {
  const folder = entry.folder;
  const folderId = folder.getId();
  const folderName = folder.getName();
  const existingSub = existingSubs.find(function (sub) { return sub.folder_id === folderId; });
  const spreadsheet = wsResolveSpreadsheetForSubActivity_(activity, folderName, year, null, entry.groupName || entry.parentFolderName);
  const spreadsheetFileId = spreadsheet ? spreadsheet.getId() : activity.spreadsheet_file_id;

  if (!existingSub) return _syncAddNewEntry_(entry, folder, folderName, activity, year, spreadsheet, spreadsheetFileId);
  if (!isTrue_(existingSub.is_active)) return _syncRestoreEntry_(entry, folder, folderName, existingSub, activity, year, spreadsheet, spreadsheetFileId);
  return _syncUpdateEntry_(entry, folderName, existingSub, activity, year, spreadsheet, spreadsheetFileId);
}

function _syncAddNewEntry_(entry, folder, folderName, activity, year, spreadsheet, spreadsheetFileId) {
  const formalArchiveName = wsSuggestFormalArchiveName_(activity, folderName);
  const targetSheetName = wsSuggestDetailSheetName_(activity, folderName, spreadsheet);
  const subRow = ConfigRepository.addSubActivityRow({
    year: year,
    activityId: activity.activity_id,
    subActivityId: entry.subActivityId,
    subActivityName: folderName,
    formalArchiveName: formalArchiveName,
    noFolder: WorkspaceSetupService.parseLeadingNumber(folderName) || activity.folder_no,
    defaultKodeKlasifikasi: '',
    allowNonLetterDocument: activity.allowNonLetter,
    targetSheetName: targetSheetName,
    mappingStatus: wsInferMappingStatus_(folderName, formalArchiveName, targetSheetName),
    mappingNote: '',
    folderPath: entry.folderPath,
    parentFolderId: entry.parentFolderId,
    parentFolderName: entry.parentFolderName,
    parentFolderPath: entry.parentFolderPath,
    spreadsheetFileId: spreadsheetFileId
  }, folder);

  if (subRow && subRow._shiftedSubActivities && subRow._shiftedSubActivities.length > 0) {
    SpreadsheetService.cascadeNomorBerkasShift(year, subRow._shiftedSubActivities);
  }
  SpreadsheetService.ensureSubActivitySheet(activity, subRow);
  const rekap = SpreadsheetService.appendRekapRowIfPresent(activity, subRow);
  persistRekapRowNumber_(year, activity, subRow, rekap);
  ensureDokumenMirrorSubActivity_(activity.target_folder_id, folderName, entry.parentFolderName);
  return true;
}

function _syncRestoreEntry_(entry, folder, folderName, existingSub, activity, year, spreadsheet, spreadsheetFileId) {
  const inactiveReason = String(existingSub.inactive_reason || '');
  if (inactiveReason === SUB_ACTIVITY_INACTIVE_REASON.DRIVE_MISSING ||
      inactiveReason === SUB_ACTIVITY_INACTIVE_REASON.DRIVE_TRASHED) {
    // proceed to restore
  } else {
    return false;
  }
  const formalArchiveName = existingSub.formal_archive_name || wsSuggestFormalArchiveName_(activity, folderName);
  const targetSheetName = existingSub.target_sheet_name || wsSuggestDetailSheetName_(activity, folderName, spreadsheet);
  const restoredSub = ConfigRepository.updateSubActivityMapping({
    year: year, activityId: activity.activity_id, subActivityId: existingSub.sub_activity_id,
    folderId: folder.getId(),
    folderPath: entry.folderPath,
    targetSheetName: targetSheetName,
    formalArchiveName: formalArchiveName,
    mappingStatus: existingSub.mapping_status || wsInferMappingStatus_(folderName, formalArchiveName, targetSheetName),
    subActivityName: folderName,
    parentFolderId: entry.parentFolderId,
    parentFolderName: entry.parentFolderName,
    parentFolderPath: entry.parentFolderPath,
    spreadsheetFileId: spreadsheetFileId,
    isActive: true, inactiveReason: '', inactiveAt: ''
  });
  SpreadsheetService.ensureSubActivitySheet(activity, Object.assign({}, existingSub, restoredSub, {
    sub_activity_name: folderName
  }));
  SpreadsheetService.updateRekapSubActivityIdentity(activity, existingSub, restoredSub);
  SpreadsheetService.clearRekapSubActivityInactiveMark(activity, restoredSub);
  return true;
}

function _syncUpdateEntry_(entry, folderName, existingSub, activity, year, spreadsheet, spreadsheetFileId) {
  if (existingSub.sub_activity_name === folderName &&
      existingSub.folder_path === entry.folderPath &&
      existingSub.parent_folder_id === entry.parentFolderId &&
      existingSub.spreadsheet_file_id === spreadsheetFileId) return false;
  const formalArchiveName = existingSub.formal_archive_name || wsSuggestFormalArchiveName_(activity, folderName);
  const targetSheetName = existingSub.target_sheet_name || wsSuggestDetailSheetName_(activity, folderName, spreadsheet);
  const updatedSub = ConfigRepository.updateSubActivityMapping({
    year: year, activityId: activity.activity_id, subActivityId: existingSub.sub_activity_id,
    folderId: existingSub.folder_id,
    folderPath: entry.folderPath,
    targetSheetName: targetSheetName,
    formalArchiveName: formalArchiveName,
    mappingStatus: existingSub.mapping_status || wsInferMappingStatus_(folderName, formalArchiveName, targetSheetName),
    subActivityName: folderName,
    parentFolderId: entry.parentFolderId,
    parentFolderName: entry.parentFolderName,
    parentFolderPath: entry.parentFolderPath,
    spreadsheetFileId: spreadsheetFileId
  });
  SpreadsheetService.ensureSubActivitySheet(activity, updatedSub);
  SpreadsheetService.updateRekapSubActivityIdentity(activity, existingSub, updatedSub);
  return true;
}
