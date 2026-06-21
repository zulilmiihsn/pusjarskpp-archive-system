'use strict';

/** Approximate bytes per page for scanned PDF size-based estimation. */
var PDF_SIZE_PER_PAGE_BYTES = 81920;
/** MIME types eligible for OCR text extraction. */
var PARSEABLE_MIME_TYPES = {
  'application/pdf': true,
  'application/msword': true,
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': true,
  'application/vnd.oasis.opendocument.text': true,
  'application/rtf': true,
  'text/plain': true
};

const ArchiveController = {
  initInboxResumableUpload: function (payload) {
    payload = payload || {};
    const safeName = validateFilePayloadForResumable_(payload);
    const year = Number(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const config = CacheHelper.getConfig(year);
    const yearConfig = ConfigService.getYearConfig(config, year);
    const rootFolder = DriveApp.getFolderById(yearConfig.root_folder_id);
    const settings = ConfigService.getSettings();
    const systemFolder = DriveService.resolveSystemFolder(settings, rootFolder);
    const inbox = yearConfig.inbox_folder_id
      ? DriveApp.getFolderById(cleanId_(yearConfig.inbox_folder_id))
      : DriveService.getOrCreateChildFolder(systemFolder, 'Inbox Dokumen Masuk');
    const target = DriveService.getOrCreateChildFolder(inbox, chooseInboxSubfolder_(safeName));

    const mimeType = payload.mimeType || 'application/octet-stream';
    const totalSize = Number(payload.totalSize);
    if (!totalSize || totalSize < 1) throw new Error('Ukuran file tidak valid.');

    return resumableUploadInit_(safeName, mimeType, totalSize, target.getId());
  },

  uploadSourceFile: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.name, 'Nama file');
    Validator.requireLongString(payload.dataUrl, 'Data URL berkas');

    return DriveService.uploadToInbox(payload);
  },

  createDraft: function (payload) {
    var ctx = _resolveArchiveContext_(payload);
    const config = ctx.config, activity = ctx.activity, subActivity = ctx.subActivity;

    const file = DriveService.getFileFromInput(payload);
    const fields = config.fields.filter(function (field) { return field.activity_id === payload.activityId && isTrue_(field.is_visible_in_form); });
    const nextItem = SpreadsheetService.getNextItemNumber(activity, subActivity);
    
    let fileLastUpdatedStr = '';
    try {
      const lastUpdated = file.getLastUpdated();
      fileLastUpdatedStr = Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
      console.warn('Gagal membaca lastUpdated berkas saat createDraft: ' + e.message);
    }

    const enrichedPayload = Object.assign({}, payload, {
      sourceFileName: file.getName(),
      nomorItemArsip: payload.nomorItemArsip || nextItem,
      noBerkas: payload.noBerkas || subActivity.sort_order || 1,
      fileLastUpdatedStr: fileLastUpdatedStr
    });
    const draft = MetadataService.createDraft(enrichedPayload, activity, subActivity, fields);
    const locationDefaults = SpreadsheetService.getDetailMetadataDefaults(activity, subActivity);
    ['no_filing_cabinet', 'no_laci', 'no_folder', 'klasifikasi_akses'].forEach(function (key) {
      if (locationDefaults[key]) draft.metadata[key] = locationDefaults[key];
    });

    return {
      year: ctx.year, activity: activity, subActivity: subActivity, sourceFile: DriveService.fileToDto(file),
      fields: fields, metadata: draft.metadata, confidence: draft.confidence,
      notes: draft.notes, finalFileName: draft.metadata.lokasi_simpan
    };
  },

  getArchiveMetadata: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.archiveId, 'Archive ID');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const log = payload.logData || ConfigRepository.getArchiveLog(payload.archiveId);
    if (!log) throw new Error('Log arsip tidak ditemukan.');

    const config = CacheHelper.getConfig(year);
    const activity = ConfigService.findActivity(config, log.activity_id);
    const subActivity = activity ? ConfigService.findSubActivity(config, log.activity_id, log.sub_activity_id) : null;
    const fields = activity ? config.fields.filter(function (field) { return field.activity_id === log.activity_id && isTrue_(field.is_visible_in_form); }) : [];

    let metadata = {};
    if (activity && subActivity && (log.final_file_id || log.spreadsheet_row_number)) {
       const rowHint = log.spreadsheet_row_number ? parseInt(log.spreadsheet_row_number, 10) : null;
       metadata = SpreadsheetService.getArchiveRowByFileId(activity, subActivity, log.final_file_id, rowHint);
       // nomor_surat, satuan, ket gak ada di detail sheet — ambil dari log
       if (log.metadata_json) {
         try {
           var savedMeta = JSON.parse(log.metadata_json);
           if (savedMeta.metadata) {
             if (!metadata.nomor_surat && savedMeta.metadata.nomor_surat) metadata.nomor_surat = savedMeta.metadata.nomor_surat;
             if (!metadata.satuan && savedMeta.metadata.satuan) metadata.satuan = savedMeta.metadata.satuan;
             if (!metadata.ket && savedMeta.metadata.ket) metadata.ket = savedMeta.metadata.ket;
           }
         } catch (e) { console.warn('getArchiveMetadata: failed to parse metadata_json: ' + e.message); }
       }
    } else if (log.metadata_json) {
       try {
         const draftObj = JSON.parse(log.metadata_json);
         if (draftObj.metadata) metadata = draftObj.metadata;
       } catch (e) { console.warn('getArchiveMetadata: failed to parse draft metadata_json: ' + e.message); }
    }

    return {
       archiveId: log.archive_id,
       year: year,
       activity: activity || null,
       subActivity: subActivity || null,
       fields: fields,
       metadata: metadata,
        sourceFile: (function() {
          var sf = { id: log.final_file_id || log.source_file_id || '', name: log.final_file_name || '', url: '', fileSize: 0 };
          if (log.final_file_id) {
            sf.url = 'https://drive.google.com/file/d/' + log.final_file_id + '/view';
          }
          return sf;
        })(),
       finalFileName: log.final_file_name || '',
       targetFolderId: log.target_folder_id || '',
       targetFolderName: log.target_folder_name || '',
       targetLocked: true,
       folderBreadcrumb: [activity ? activity.activity_name : '', subActivity ? subActivity.sub_activity_name : ''].filter(Boolean)
    };
  },

  getMetadataDefaults: function (payload) {
    var ctx = _resolveArchiveContext_(payload);

    return {
      year: ctx.year,
      activityId: ctx.activity.activity_id,
      subActivityId: ctx.subActivity.sub_activity_id,
      metadata: SpreadsheetService.getDetailMetadataDefaults(ctx.activity, ctx.subActivity, payload.forceCalculate === true)
    };
  },

  saveDraftToLog: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    Validator.requireString(payload.activityId, 'Activity ID');
    Validator.requireString(payload.subActivityId, 'Sub Activity ID');
    Validator.requireLongString(payload.draftState, 'Draft State JSON');

    let draftStateObj;
    try {
      draftStateObj = JSON.parse(payload.draftState);
    } catch (e) {
      throw new Error('Format draft state tidak valid');
    }

    const draftUser = AuthService.getCurrentUser(payload);
    const draftCreatedBy = draftUser.displayName || draftUser.username || '';

    const archiveId = 'DRF-' + Utilities.getUuid().slice(0, 8).toUpperCase();

    ConfigRepository.appendArchiveLog({
      archive_id: archiveId, year: year, activity_id: payload.activityId,
      sub_activity_id: payload.subActivityId,
      source_file_id: draftStateObj.sourceFile ? draftStateObj.sourceFile.id || '' : '',
      final_file_id: '', final_file_name: draftStateObj.finalFileName || '',
      target_folder_id: draftStateObj.targetFolderId || '',
      target_folder_name: draftStateObj.targetFolderName || '',
      target_folder_path: Array.isArray(draftStateObj.folderBreadcrumb) ? draftStateObj.folderBreadcrumb.join(' > ') : '',
      spreadsheet_file_id: '', spreadsheet_row_number: '',
      status: STATUS.DRAFT, created_at: new Date().toISOString(),
      created_by: draftCreatedBy, error_message: '',
      metadata_json: payload.draftState
    });
    CacheHelper.invalidate(year);

    return { archiveId: archiveId, status: STATUS.DRAFT };
  },

  deleteDraft: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.archiveId, 'Archive ID');

    return withLock_(() => {
      const year = payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR;
      const sourceFileId = ConfigRepository.deleteDraftLog(payload.archiveId);
      if (sourceFileId) {
        CacheHelper.invalidate(year);
        try {
          DriveApp.getFileById(sourceFileId).setTrashed(true);
        } catch (e) {
          console.error('deleteDraft: failed to trash source file ' + sourceFileId + ': ' + e.message);
        }
      }
      return { success: true };
    }, 30000);
  },

  deleteArchive: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.archiveId, 'Archive ID');

    return withLock_(() => {
      const year = payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR;
      const log = ConfigRepository.getArchiveLog(payload.archiveId);
      if (!log) throw new Error('Log arsip tidak ditemukan.');

      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, log.activity_id);
      const subActivity = ConfigService.findSubActivity(config, log.activity_id, log.sub_activity_id);

      if (!activity || !subActivity) {
         throw new Error('Konfigurasi kegiatan/sub-kegiatan arsip tidak ditemukan.');
      }

      let detailDeleteSuccess = true;
      if (log.spreadsheet_row_number || log.final_file_id) {
         const actualDeletedRow = SpreadsheetService.deleteArchiveRowAndReorder(activity, subActivity, Number(log.spreadsheet_row_number) || 0, log.final_file_id);
         if (actualDeletedRow) {
             SpreadsheetService.updateRekapSummary(activity, subActivity, {});
             ConfigRepository.decrementArchiveLogRows(log.spreadsheet_file_id, actualDeletedRow);
         } else {
             // If we couldn't delete from the sheet (e.g. out of bounds and fileId not found), it might already be gone.
             // We allow the config log deletion to proceed so the UI clears it.
         }
      }

      ConfigRepository.deleteArchiveLog(payload.archiveId);

      if (log.final_file_id) {
         try { DriveApp.getFileById(log.final_file_id).setTrashed(true); } catch(e) { console.warn('deleteArchive: failed to trash final file ' + log.final_file_id + ': ' + e.message); }
      }
      if (log.source_file_id) {
         try { DriveApp.getFileById(log.source_file_id).setTrashed(true); } catch(e) { console.warn('deleteArchive: failed to trash source file ' + log.source_file_id + ': ' + e.message); }
      }

      CacheHelper.invalidate(year);
      return { success: true };
    }, 30000);
  },

  _prepareMetadata: function (payload, activity, subActivity, sourceFileName) {
    const metadata = MetadataService.normalize(payload.metadata, activity, subActivity, sourceFileName);
    metadata.lokasi_simpan = MetadataService.buildFinalFileName(metadata, sourceFileName);
    return metadata;
  },

  _validateUniqueMetadata: function(payload, activity, subActivity, currentRowNumber) {
    const existingPairs = SpreadsheetService.listExistingItemUraianPairs(activity, subActivity);
    const rawIncomingItem = String(payload.metadata.nomor_item_arsip || '').trim();
    const incomingItemNumber = rawIncomingItem ? rawIncomingItem.replace(/^0+/, '').padStart(2, '0') : '';
    const incomingUraian = String(payload.metadata.uraian_informasi_item || '').trim().toLowerCase();

    existingPairs.forEach(function(pair) {
      if (currentRowNumber && pair.rowNumber === currentRowNumber) return;

      var existingItem = String(pair.nomor_item_arsip || '').trim();
      var existingItemNumber = existingItem ? existingItem.replace(/^0+/, '').padStart(2, '0') : '';
      var existingUraian = String(pair.uraian_informasi_item || '').trim().toLowerCase();

      if (incomingUraian && existingUraian === incomingUraian) {
        throw new Error('Uraian informasi item "' + payload.metadata.uraian_informasi_item + '" sudah ada. Harap gunakan uraian yang berbeda.');
      }

      if (incomingItemNumber && existingItemNumber === incomingItemNumber) {
        throw new Error('Nomor Item Arsip "' + rawIncomingItem + '" sudah digunakan. Biarkan kosong agar sistem mengisi otomatis angka selanjutnya.');
      }
    });
  },

  _processArchiveInLock: function (payload, activity, subActivity, metadata, sourceFile, logContext) {
    return withLock_(() => {
      // Uniqueness diperiksa di dalam lock agar read-validate-write atomik
      // terhadap arsip lain yang dikerjakan bersamaan.
      ArchiveController._validateUniqueMetadata(payload, activity, subActivity, null);
      if (!metadata.nomor_item_arsip) {
        const nextNum = SpreadsheetService.getNextItemNumber(activity, subActivity);
        metadata.nomor_item_arsip = String(nextNum).padStart(2, '0');
        metadata.no_berkas = String(nextNum);
        metadata.lokasi_simpan = MetadataService.buildFinalFileName(metadata, sourceFile.getName());
      }
      const targetFolderId = payload.targetFolderId || subActivity.folder_id;
      const targetFolder = DriveApp.getFolderById(cleanId_(targetFolderId));
      const targetFolderInfo = getArchiveTargetFolderInfo_(targetFolder);
      const finalFile = DriveService.copyToFinalFolder(sourceFile, targetFolderId, metadata.lokasi_simpan, payload.year);

      metadata.lokasi_simpan = finalFile.getName();
      metadata._lokasi_simpan_url = finalFile.getUrl();

      const writeResult = SpreadsheetService.appendArchiveRow(activity, subActivity, metadata);
      SpreadsheetApp.flush();
      const rekapResult = SpreadsheetService.updateRekapSummary(activity, subActivity, metadata);

      // Catat log penyelesaian DI DALAM lock yang sama dengan penulisan baris.
      // Kalau ditulis setelah lock dilepas, ada celah crash: file + baris sheet
      // sudah ada tapi log arsip belum — arsip jadi tak tercatat / bisa di-import
      // ganda saat sync. (LockService reentrant untuk eksekusi yang sama.)
      let logWarning = '';
      if (logContext) {
        logWarning = ArchiveController._logArchiveCompletion(
          logContext.archiveId, logContext.year, activity, subActivity,
          logContext.sourceFileId, finalFile, writeResult, logContext.createdBy,
          targetFolderInfo, metadata
        );
      }

      return {
        finalFile: finalFile,
        writeResult: writeResult,
        rekapResult: rekapResult,
        metadata: metadata,
        targetFolder: targetFolderInfo,
        logWarning: logWarning
      };
    }, 30000);
  },

  _logArchiveCompletion: function (archiveId, year, activity, subActivity, sourceFileId, finalFile, writeResult, createdBy, targetFolder, metadata) {
    try {
      ConfigRepository.appendArchiveLog({
        archive_id: archiveId, year: year,
        activity_id: activity.activity_id, sub_activity_id: subActivity.sub_activity_id,
        source_file_id: sourceFileId, final_file_id: finalFile.getId(),
        final_file_name: finalFile.getName(),
        target_folder_id: targetFolder ? targetFolder.id : '',
        target_folder_name: targetFolder ? targetFolder.name : '',
        target_folder_path: targetFolder ? targetFolder.path : '',
        spreadsheet_file_id: writeResult.spreadsheetId,
        spreadsheet_row_number: writeResult.rowNumber, status: STATUS.COMPLETED,
        created_at: new Date().toISOString(), created_by: createdBy || '', error_message: '',
        metadata_json: JSON.stringify({ metadata: metadata })
      });
      CacheHelper.invalidate(year);
      return '';
    } catch (logError) {
      console.error('Failed to write archive error log: ' + logError.message);
      return 'Arsip tersimpan tetapi riwayat gagal dicatat: ' + logError.message;
    }
  },

  finalizeArchive: function (payload) {
    var ctx = _resolveArchiveContext_(payload);
    const year = ctx.year, config = ctx.config, activity = ctx.activity, subActivity = ctx.subActivity;

    const fields = config.fields.filter(function (field) { return field.activity_id === payload.activityId && isTrue_(field.is_visible_in_form); });
    Validator.requireMetadata(payload.metadata, fields);
    // Catatan: cek uniqueness (uraian/nomor item) dilakukan DI DALAM lock pada
    // _processArchiveInLock — kalau di sini, dua submit bersamaan sama-sama lolos
    // sebelum baris pertama tertulis, lalu menghasilkan duplikat.

    const archiveId = 'ARC-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const sourceFile = DriveService.getFileFromInput(payload);
    const archiveUser = AuthService.getCurrentUser(payload);
    const archiveCreatedBy = archiveUser.displayName || archiveUser.username || '';

    let result;
    try {
      const initialMetadata = ArchiveController._prepareMetadata(payload, activity, subActivity, sourceFile.getName());
      result = ArchiveController._processArchiveInLock(payload, activity, subActivity, initialMetadata, sourceFile, {
        archiveId: archiveId,
        year: year,
        sourceFileId: sourceFile.getId(),
        createdBy: archiveCreatedBy
      });
      var logWarning = result.logWarning || '';
    } catch (error) {
      ConfigRepository.appendArchiveLog({
        archive_id: archiveId,
        year: year,
        activity_id: activity.activity_id,
        sub_activity_id: subActivity.sub_activity_id,
        source_file_id: sourceFile.getId(),
        final_file_id: '',
        final_file_name: '',
        target_folder_id: payload.targetFolderId || subActivity.folder_id || '',
        target_folder_name: '',
        target_folder_path: '',
        spreadsheet_file_id: '',
        spreadsheet_row_number: '',
        status: STATUS.FAILED,
        created_at: new Date().toISOString(),
        created_by: archiveCreatedBy,
        error_message: error.message || String(error),
        metadata_json: JSON.stringify({ failedPayload: payload.metadata || {}, sourceFileName: sourceFile.getName() })
      });
      CacheHelper.invalidate(year);
      throw error;
    }

    return {
      archiveId: archiveId, status: STATUS.COMPLETED,
      finalFile: DriveService.fileToDto(result.finalFile),
      spreadsheet: result.writeResult, rekapSpreadsheet: result.rekapResult,
      metadata: result.metadata, warning: logWarning || ''
    };
  },

  adoptExistingArchives: function (payload) {
    payload = payload || {};
    const dryRun = !!payload.dryRun;
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const config = CacheHelper.getConfig(year);
    const archiveUser = AuthService.getCurrentUser(payload);
    const createdBy = archiveUser.displayName || archiveUser.username || '';
    const existingKeys = ConfigRepository.getArchiveLogKeyMap(year);
    const activityFilter = payload.activityId || '';
    const subFilter = payload.subActivityId || '';
    const result = {
      year: year,
      scannedSheets: 0,
      scannedRows: 0,
      imported: 0,
      skippedExisting: 0,
      missingSheets: 0,
      missingFileMatches: 0,
      dryRun: dryRun,
      byActivity: []
    };

    config.activities.forEach(function (activity) {
      if (activityFilter && activity.activity_id !== activityFilter) return;
      const subActivities = config.subActivities.filter(function (sub) {
        return sub.activity_id === activity.activity_id && (!subFilter || sub.sub_activity_id === subFilter);
      });
      const activityResult = {
        activityId: activity.activity_id,
        activityName: activity.activity_name,
        scannedRows: 0,
        imported: 0,
        skippedExisting: 0,
        missingSheets: 0,
        missingFileMatches: 0,
        subActivities: []
      };

      subActivities.forEach(function (subActivity) {
        const subResult = _adoptSubActivity_(activity, subActivity, result, activityResult, existingKeys, year, createdBy, dryRun);
        activityResult.subActivities.push(subResult);
      });

      result.byActivity.push(activityResult);
    });

    if (!dryRun) CacheHelper.invalidate(year);
    return result;
  },

  validateArchiveFields: function (payload) {
    var ctx = _resolveArchiveContext_(payload);
    var activity = ctx.activity, subActivity = ctx.subActivity;

    var existingData = SpreadsheetService.listExistingArchiveRows(activity, subActivity);
    var errors = [];

    var incomingItem = String(payload.nomor_item_arsip || '').trim();
    if (incomingItem) {
      var normalizedItem = incomingItem.replace(/^0+/, '').padStart(2, '0');
      existingData.rows.forEach(function (row) {
        var existingItem = String(row.metadata.nomor_item_arsip || '').trim();
        var normalizedExisting = existingItem ? existingItem.replace(/^0+/, '').padStart(2, '0') : '';
        if (normalizedExisting === normalizedItem) {
          errors.push({ field: 'nomor_item_arsip', message: 'Nomor Item Arsip "' + incomingItem + '" sudah digunakan.' });
        }
      });
    }

    var incomingUraian = String(payload.uraian_informasi_item || '').trim().toLowerCase();
    if (incomingUraian) {
      existingData.rows.forEach(function (row) {
        var existingUraian = String(row.metadata.uraian_informasi_item || '').trim().toLowerCase();
        if (existingUraian === incomingUraian) {
          errors.push({ field: 'uraian_informasi_item', message: 'Uraian informasi item "' + payload.uraian_informasi_item + '" sudah ada.' });
        }
      });
    }

    return { valid: errors.length === 0, errors: errors };
  },

  editMetadata: function (payload) {
    var ctx = _resolveArchiveContext_(payload);
    const year = ctx.year, config = ctx.config, activity = ctx.activity, subActivity = ctx.subActivity;
    Validator.requireString(payload.archiveId, 'Archive ID');

    const fields = config.fields.filter(function (field) { return field.activity_id === payload.activityId && isTrue_(field.is_visible_in_form); });
    Validator.requireMetadata(payload.metadata, fields);

    const log = ConfigRepository.getArchiveLog(payload.archiveId);
    if (!log) throw new Error('Log arsip tidak ditemukan.');

    ArchiveController._validateUniqueMetadata(payload, activity, subActivity, Number(log.spreadsheet_row_number));

    const metadata = ArchiveController._prepareMetadata(payload, activity, subActivity, log.final_file_name);
    let finalFileDto = null;
    let finalFileName = log.final_file_name;

    return withLock_(() => {
      let fileUrl = '';
      if (log.final_file_id) {
        try {
          const file = DriveApp.getFileById(log.final_file_id);
          fileUrl = file.getUrl();
          if (metadata.lokasi_simpan !== log.final_file_name) {
            file.setName(metadata.lokasi_simpan);
            finalFileDto = DriveService.fileToDto(file);
            finalFileName = file.getName();
          }
        } catch (e) {
          console.warn('Failed to handle drive file: ' + e.message);
          metadata.lokasi_simpan = log.final_file_name;
        }
      }
      metadata._lokasi_simpan_url = fileUrl;

      if (metadata._resolved_row_number) {
        SpreadsheetService.updateArchiveRow(activity, subActivity, parseInt(metadata._resolved_row_number, 10), metadata);
      } else if (log.spreadsheet_row_number) {
        SpreadsheetService.updateArchiveRow(activity, subActivity, parseInt(log.spreadsheet_row_number, 10), metadata);
      }
      SpreadsheetService.updateRekapSummary(activity, subActivity, metadata);

      let newMetadataJson = log.metadata_json;
      try {
        let draftStateObj = JSON.parse(log.metadata_json);
        draftStateObj.metadata = metadata;
        draftStateObj.finalFileName = finalFileName;
        newMetadataJson = JSON.stringify(draftStateObj);
      } catch (e) {
        newMetadataJson = JSON.stringify({ metadata: metadata, finalFileName: finalFileName });
      }

      ConfigRepository.updateArchiveLog(payload.archiveId, {
        metadata_json: newMetadataJson,
        final_file_name: finalFileName
      });

      CacheHelper.invalidate(year);

      var editWarning = '';
      try {
        const editUser = AuthService.getCurrentUser(payload);
        const editCreatedBy = editUser.displayName || editUser.username || 'unknown';
        ConfigRepository.appendAdminAudit({
          created_at: new Date().toISOString(),
          actor: editCreatedBy,
          action: 'METADATA_UPDATED',
          year: year,
          activity_id: payload.activityId,
          sub_activity_id: payload.subActivityId,
          folder_id: log.final_file_id || '',
          status: 'SUCCESS',
          message: 'Mengubah metadata arsip ID: ' + payload.archiveId + ' (' + finalFileName + ')'
        });
      } catch (auditError) {
        console.error('Failed to log metadata edit in audit: ' + auditError.message);
        editWarning = 'Metadata tersimpan tetapi log audit gagal dicatat.';
      }

      return {
        archiveId: payload.archiveId,
        status: STATUS.COMPLETED,
        finalFile: finalFileDto || { id: log.final_file_id, name: finalFileName },
        metadata: metadata,
        warning: editWarning
      };
    }, 30000);
  },

  listInboxFiles: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);

    const settings = ConfigService.getSettings();
    const config = CacheHelper.getConfig(year);
    const yearConfig = ConfigService.getYearConfig(config, year);
    const rootFolder = DriveApp.getFolderById(yearConfig.root_folder_id);
    const systemFolder = DriveService.resolveSystemFolder(settings, rootFolder);

    const inboxId = yearConfig.inbox_folder_id
      ? cleanId_(yearConfig.inbox_folder_id)
      : DriveService.getOrCreateChildFolder(systemFolder, 'Inbox Dokumen Masuk').getId();

    const list = [];

    function traverse(folderId) {
      let pageToken = null;
      do {
        let result;
        try {
          result = Drive.Files.list({
            q: "'" + folderId + "' in parents and trashed = false",
            fields: "nextPageToken, files(id, name, webViewLink, mimeType, size, createdTime)",
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
            traverse(item.id);
          } else {
            list.push({
              id: item.id, 
              name: item.name, 
              url: item.webViewLink,
              downloadUrl: DriveService.getDownloadUrl(item.id, item.mimeType),
              mimeType: item.mimeType, 
              size: item.size || 0,
              created: new Date(item.createdTime).getTime()
            });
          }
        }
        pageToken = result.nextPageToken;
      } while (pageToken);
    }

    traverse(inboxId);
    list.sort(function (a, b) { return b.created - a.created; });
    return list;
  },

  parseDocumentContent: function (payload) {
    Validator.requireString(payload.fileId, 'File ID');
    var startTime = Date.now();
    var file = DriveApp.getFileById(cleanId_(payload.fileId));
    var mimeType = file.getMimeType();
    var fileName = file.getName();
    var fileSize = file.getSize();

    if (!PARSEABLE_MIME_TYPES[mimeType]) {
      return {
        fields: {},
        rawTextLength: 0,
        parseDuration: Date.now() - startTime,
        fieldCount: 0,
        totalFields: 6,
        skipped: true,
        skipReason: 'Tipe file tidak didukung untuk parsing otomatis (' + (mimeType || 'unknown') + ').'
      };
    }

    var text = '';
    var pageCount = 0;
    if (mimeType === 'application/pdf') {
      pageCount = extractPdfPageCount_(file);
    }
    var tempDocId = '';

    // Try multiple extraction methods
    var extractionMethods = [];
    var extractionErrors = [];

    try {
      // Method 1: Direct text extraction (works for text/plain files)
      if (mimeType === 'text/plain') {
        text = file.getBlob().getDataAsString();
        extractionMethods.push('direct_text');
      }
    } catch (e) {
      console.warn('Method 1 failed: ' + e.message);
    }

    if (!text || text.length < 100) {
      // Method 2: OCR via Drive.Files.create has been moved to OcrBackgroundJob
      // to prevent UI freezing. The background job will fill missing data later.
      extractionMethods.push('ocr_deferred');
    }

    // Method 3 removed because it unsafely calls getDataAsString on binary files like PDFs

    // Fallback to filename if nothing extracted
    if (!text) {
      text = fileName;
      extractionMethods.push('filename_only');
    }



    // Cleanup temp doc
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (e) {}
    }

    console.log('Extraction methods tried: ' + extractionMethods.join(', ') + ', final text length: ' + text.length);

    // Run ParseEngine for scored multi-pass extraction
    var engineResult = ParseEngine.analyze(text, fileName, { activity: {}, subActivity: {} });
    var fields = engineResult.fields;

    // Use file modified date as fallback for date field
    if (!fields.tanggal) {
      try {
        var lastUpdated = file.getLastUpdated();
        var lastUpdatedStr = Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        fields.tanggal = {
          value: lastUpdatedStr,
          score: 0.5,
          confidence: 'medium',
          source: 'file_metadata_last_updated'
        };
      } catch (e) {
        console.warn('Gagal membaca lastUpdated berkas: ' + e.message);
      }
    }

    // Debug: log OCR text sample and extraction results
    var debugInfo = {
      rawText: text,
      errors: extractionErrors,
      ocrTextLength: text ? text.length : 0,
      ocrTextSample: text.substring(0, 500),
      extractionMethods: extractionMethods,
      structure: engineResult.structure,
      documentType: engineResult.documentType,
      fieldsFound: Object.keys(fields),
      fieldsMissing: ['nomor_surat','kode_klasifikasi','tanggal','uraian_informasi_item','klasifikasi_akses','pengirim','penerima','tanda_tangan','lampiran']
        .filter(function(k) { return !fields[k]; })
    };
    console.log('ParseEngine debug:', JSON.stringify(debugInfo, null, 2));

    // Flatten complex fields for client compatibility
    if (fields.tanda_tangan && fields.tanda_tangan.value) {
      var tt = fields.tanda_tangan.value;
      var ttParts = [];
      if (tt.jabatan) ttParts.push(tt.jabatan);
      if (tt.nama) ttParts.push(tt.nama);
      fields.tanda_tangan.value = ttParts.join(', ') || ttParts.join('');
    }
    if (fields.pengirim && typeof fields.pengirim.value === 'object') {
      delete fields.pengirim; // skip if no string value
    }

    // Add page count (estimated separately from OCR)
    if (pageCount > 0) {
      fields.jumlah = { value: String(pageCount), score: 0.6, confidence: 'medium', source: 'page_estimation' };
    }

    return {
      fields: fields,
      rawTextLength: engineResult.rawTextLength || text.length,
      parseDuration: Date.now() - startTime,
      fieldCount: Object.keys(fields).length,
      totalFields: engineResult.totalFields || 9,
      documentType: engineResult.documentType || '',
      structure: engineResult.structure || {},
      debug: debugInfo
    };
  }
};

function _adoptSubActivity_(activity, subActivity, result, activityResult, existingKeys, year, createdBy, dryRun) {
  const listing = SpreadsheetService.listExistingArchiveRows(activity, subActivity);
  result.scannedSheets++;
  const subResult = {
    subActivityId: subActivity.sub_activity_id,
    subActivityName: subActivity.sub_activity_name,
    sheetName: listing.sheetName || '',
    scannedRows: 0,
    importable: 0,
    skippedExisting: 0,
    missingFileMatches: 0,
    missingSheet: !!listing.missingSheet
  };
  if (listing.missingSheet) {
    result.missingSheets++;
    activityResult.missingSheets++;
    return subResult;
  }

  if (!dryRun) {
    try {
      SpreadsheetService.updateRekapSummary(activity, subActivity, {});
    } catch (summaryError) {
      console.warn('adoptExistingArchives: gagal refresh rangkuman ' + subActivity.sub_activity_name + ': ' + summaryError.message);
    }
  }

  // Index nama file dibangun SEKALI per sub-kegiatan (bukan rekursi folder per baris).
  // Sebelumnya tiap baris menelusuri ulang seluruh pohon folder → O(baris × pohon).
  const fileIndex = listing.rows.length ? buildArchiveFileIndex_(subActivity.folder_id) : null;
  listing.rows.forEach(function (row) {
    _adoptRow_(row, listing, activity, subActivity, result, activityResult, subResult, existingKeys, year, createdBy, dryRun, fileIndex);
  });
  return subResult;
}

function _adoptRow_(row, listing, activity, subActivity, result, activityResult, subResult, existingKeys, year, createdBy, dryRun, fileIndex) {
  result.scannedRows++;
  activityResult.scannedRows++;
  subResult.scannedRows++;

  const sheetKey = cleanId_(listing.spreadsheetId) + ':' + row.rowNumber;
  if (existingKeys[sheetKey]) {
    result.skippedExisting++;
    activityResult.skippedExisting++;
    subResult.skippedExisting++;
    return;
  }

  const finalFileName = row.metadata.lokasi_simpan || row.metadata.uraian_informasi_item || row.metadata.uraian_informasi_item || '';
  const finalFileMatch = lookupArchiveFile_(fileIndex, finalFileName);
  const finalFile = finalFileMatch ? finalFileMatch.file : null;
  const targetFolderInfo = finalFileMatch ? finalFileMatch.folderInfo : null;
  if (!finalFile && finalFileName) {
    result.missingFileMatches++;
    activityResult.missingFileMatches++;
    subResult.missingFileMatches++;
  }
  subResult.importable++;
  result.imported++;
  activityResult.imported++;
  if (dryRun) return;

  const archiveId = 'IMP-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  ConfigRepository.appendArchiveLog({
    archive_id: archiveId,
    year: year,
    activity_id: activity.activity_id,
    sub_activity_id: subActivity.sub_activity_id,
    source_file_id: '',
    final_file_id: finalFile ? finalFile.getId() : '',
    final_file_name: finalFile ? finalFile.getName() : finalFileName,
    target_folder_id: targetFolderInfo ? targetFolderInfo.id : '',
    target_folder_name: targetFolderInfo ? targetFolderInfo.name : '',
    target_folder_path: targetFolderInfo ? targetFolderInfo.path : '',
    spreadsheet_file_id: listing.spreadsheetId,
    spreadsheet_row_number: row.rowNumber,
    status: STATUS.COMPLETED,
    created_at: new Date().toISOString(),
    created_by: createdBy,
    error_message: '',
    metadata_json: JSON.stringify({
      importedFromExisting: true,
      metadata: row.metadata,
      finalFileName: finalFile ? finalFile.getName() : finalFileName,
      spreadsheet: {
        spreadsheetId: listing.spreadsheetId,
        spreadsheetUrl: listing.spreadsheetUrl,
        sheetName: listing.sheetName,
        rowNumber: row.rowNumber
      }
    })
  });
  existingKeys[sheetKey] = true;
  if (finalFile) existingKeys['file:' + finalFile.getId()] = true;
}

/**
 * Telusuri pohon folder sub-kegiatan SATU KALI (depth <= 6) dan bangun index
 * nama-file → { file, folder }. Dipakai untuk lookup per baris tanpa rekursi ulang.
 * @param {string} folderId
 * @return {{exact: Object, normalized: Object}}
 */
function buildArchiveFileIndex_(folderId) {
  const index = { exact: {}, normalized: {} };
  if (!folderId) return index;

  let rootName = 'Unknown';
  try {
    const meta = Drive.Files.get(cleanId_(folderId), {fields: 'name', supportsAllDrives: true});
    rootName = meta.name;
  } catch (e) {}

  const stack = [{ id: cleanId_(folderId), name: rootName, path: rootName, depth: 0 }];
  
  while (stack.length) {
    const cur = stack.pop();
    if (cur.depth > 6) continue;

    let pageToken = null;
    do {
      let result;
      try {
        result = Drive.Files.list({
          q: "'" + cur.id + "' in parents and trashed = false",
          fields: "nextPageToken, files(id, name, mimeType)",
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
          stack.push({
            id: item.id,
            name: item.name,
            path: cur.path + ' > ' + item.name,
            depth: cur.depth + 1
          });
        } else {
          const entry = { 
            file: { getId: () => item.id, getName: () => item.name }, 
            folderInfo: { id: cur.id, name: cur.name, path: cur.path } 
          };
          if (!(item.name in index.exact)) index.exact[item.name] = entry;
          const norm = normalizeFileLookupName_(item.name);
          if (norm && !(norm in index.normalized)) index.normalized[norm] = entry;
        }
      }
      pageToken = result.nextPageToken;
    } while (pageToken);
  }
  return index;
}

/**
 * Cari file di index: exact dulu, lalu normalized. Null jika tidak ada.
 * @param {{exact: Object, normalized: Object}} index
 * @param {string} fileName
 * @return {?{file: GoogleAppsScript.Drive.File, folder: GoogleAppsScript.Drive.Folder}}
 */
function lookupArchiveFile_(index, fileName) {
  if (!index || !fileName) return null;
  if (index.exact[fileName]) return index.exact[fileName];
  const norm = normalizeFileLookupName_(fileName);
  if (norm && index.normalized[norm]) return index.normalized[norm];
  return null;
}

function getArchiveTargetFolderInfo_(folder) {
  if (!folder) return null;
  return {
    id: folder.getId(),
    name: folder.getName(),
    path: buildFolderPathForLog_(folder)
  };
}

function buildFolderPathForLog_(folder) {
  const names = [];
  let current = folder;
  let guard = 0;
  while (current && guard < 12) {
    names.unshift(current.getName());
    const parents = current.getParents();
    if (!parents.hasNext()) break;
    current = parents.next();
    guard++;
  }
  return names.join(' > ');
}

function normalizeFileLookupName_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function _resolveArchiveContext_(payload) {
  payload = payload || {};
  var year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
  Validator.requireString(payload.activityId, 'Activity ID');
  Validator.requireString(payload.subActivityId, 'Sub Activity ID');
  var config = CacheHelper.getConfig(year);
  var activity = ConfigService.findActivity(config, payload.activityId);
  if (!activity) throw new Error('Kegiatan tidak ditemukan.');
  var subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
  if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
  return { year: year, config: config, activity: activity, subActivity: subActivity };
}

function extractPdfPageCount_(file) {
  try {
    var blob = file.getBlob();
    var pdfText = blob.getDataAsString('ISO-8859-1');
    
    // Try to find /Count [number]
    var countMatch = pdfText.match(/\/Count\s+(\d+)/);
    if (countMatch && countMatch[1]) {
      var count = parseInt(countMatch[1], 10);
      if (count > 0 && count < 1000) return count;
    }
    
    // Fallback: count /Type /Page
    var pageMatch = pdfText.match(/\/Type\s*\/Page\b/g);
    if (pageMatch && pageMatch.length > 0) {
      return pageMatch.length;
    }
  } catch (e) {
    console.warn('extractPdfPageCount_ failed: ' + e.message);
  }
  return 0;
}
