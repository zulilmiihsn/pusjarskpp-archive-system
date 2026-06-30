'use strict';

/** Approximate bytes per page for scanned PDF size-based estimation. */
const PDF_SIZE_PER_PAGE_BYTES = 81920;
/** MIME types eligible for OCR text extraction. */
const PARSEABLE_MIME_TYPES = {
  'application/pdf': true,
  'application/msword': true,
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': true,
  'application/vnd.oasis.opendocument.text': true,
  'application/rtf': true,
  'text/plain': true
};

// Batas ukuran file untuk ekstraksi isi (jaga kuota & waktu eksekusi GAS).
const PARSE_EXTRACT_MAX_BYTES = 25 * 1024 * 1024; // 25MB

/**
 * Ekstrak teks isi dokumen via konversi Drive → Google Doc, lalu export text/plain.
 * PDF: pakai OCR bahasa Indonesia (menangani surat hasil scan).
 * Word/ODT/RTF: konversi biasa (cepat, tanpa OCR).
 * Doc sementara selalu dihapus. Export via UrlFetch + token OAuth agar cukup pakai
 * scope Drive yang sudah ada (tak perlu scope Documents tambahan).
 * @return {{text: string, method: string}}
 */
function extractTextViaConversion_(file, mimeType) {
  try {
    const size = file.getSize ? file.getSize() : 0;
    if (size && size > PARSE_EXTRACT_MAX_BYTES) return { text: '', method: 'skipped_too_large' };

    const isPdf = (mimeType === 'application/pdf');
    const resource = { mimeType: 'application/vnd.google-apps.document', name: 'tmp_parse' };
    const optArgs = { supportsAllDrives: true };
    if (isPdf) optArgs.ocrLanguage = 'id'; // OCR untuk scan; PDF digital tetap terbaca

    const copied = Drive.Files.copy(resource, cleanId_(file.getId()), optArgs);
    const docId = copied && copied.id;
    if (!docId) return { text: '', method: 'conversion_failed' };

    let text = '';
    try {
      const token = ScriptApp.getOAuthToken();
      const url = 'https://www.googleapis.com/drive/v3/files/' + docId +
        '/export?mimeType=' + encodeURIComponent('text/plain');
      const resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) text = resp.getContentText();
    } finally {
      try { Drive.Files.remove(docId, { supportsAllDrives: true }); }
      catch (e) { try { DriveApp.getFileById(docId).setTrashed(true); } catch (e2) {} }
    }
    return { text: String(text || ''), method: isPdf ? 'ocr_id' : 'convert' };
  } catch (e) {
    console.warn('extractTextViaConversion_ failed: ' + e.message);
    return { text: '', method: 'conversion_failed' };
  }
}

const ArchiveController = {
  initInboxResumableUpload: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
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
    requireAuth_(payload);
    Validator.requireString(payload.name, 'Nama file');
    Validator.requireLongString(payload.dataUrl, 'Data URL berkas');

    return DriveService.uploadToInbox(payload);
  },

  getArchiveMetadata: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireString(payload.archiveId, 'Archive ID');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    // Selalu muat log dari sumber tepercaya berdasarkan archiveId; JANGAN percaya
    // payload.logData dari client (bisa dipalsukan untuk menunjuk file/baris lain).
    const log = ConfigRepository.getArchiveLog(payload.archiveId);
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
           const savedMeta = JSON.parse(log.metadata_json);
           if (savedMeta.metadata) {
             if (!metadata.nomor_surat && savedMeta.metadata.nomor_surat) metadata.nomor_surat = savedMeta.metadata.nomor_surat;
             if (!metadata.satuan && savedMeta.metadata.satuan) metadata.satuan = savedMeta.metadata.satuan;
             if (!metadata.ket && savedMeta.metadata.ket) metadata.ket = savedMeta.metadata.ket;
             if (savedMeta.metadata.uraian_informasi_item) metadata.uraian_informasi_item = savedMeta.metadata.uraian_informasi_item;
             if (savedMeta.metadata.kepada) metadata.kepada = savedMeta.metadata.kepada;
             if (savedMeta.metadata.dari) metadata.dari = savedMeta.metadata.dari;
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
          const sf = { id: log.final_file_id || log.source_file_id || '', name: log.final_file_name || '', url: '', fileSize: 0 };
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

  getArchiveMetadataDefaults: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const ctx = _resolveArchiveContext_(payload);

    return {
      year: ctx.year,
      activityId: ctx.activity.activity_id,
      subActivityId: ctx.subActivity.sub_activity_id,
      metadata: SpreadsheetService.getDetailMetadataDefaults(ctx.activity, ctx.subActivity, payload.forceCalculate === true)
    };
  },

  saveDraftToLog: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
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
    // Draft = work-in-progress milik user; cukup wajib login (bukan admin-only),
    // agar alur normal pengarsip tidak terganggu.
    requireAuth_(payload);
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
    // Menghapus arsip final = aksi sensitif → admin-only.
    requireAdmin_(payload);
    return ArchiveController._deleteArchiveCore_(payload);
  },

  // Logika hapus tanpa cek otorisasi. Dipanggil oleh endpoint deleteArchive (admin)
  // dan oleh DriveController.trashArchiveFile (yang sudah meng-escalate ke admin).
  // JANGAN ekspos langsung ke client.
  _deleteArchiveCore_: function (payload) {
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

      const detailDeleteSuccess = true;
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
    const existingPairs = SpreadsheetService.listExistingItemNumbers(activity, subActivity);
    const rawIncomingItem = String(payload.metadata.nomor_item_arsip || '').trim();
    const incomingItemNumber = rawIncomingItem ? rawIncomingItem.replace(/^0+/, '').padStart(2, '0') : '';

    existingPairs.forEach(function(pair) {
      if (currentRowNumber && pair.rowNumber === currentRowNumber) return;

      const existingItem = String(pair.nomor_item_arsip || '').trim();
      const existingItemNumber = existingItem ? existingItem.replace(/^0+/, '').padStart(2, '0') : '';

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
        metadata.lokasi_simpan = MetadataService.buildFinalFileName(metadata, sourceFile.getName());
      }
      const targetFolderId = payload.targetFolderId || subActivity.folder_id;
      const targetFolder = DriveApp.getFolderById(cleanId_(targetFolderId));
      const targetFolderInfo = getArchiveTargetFolderInfo_(targetFolder);
      const finalFile = DriveService.copyToFinalFolder(sourceFile, targetFolderId, metadata.lokasi_simpan, payload.year);

      metadata.lokasi_simpan = finalFile.getName();
      metadata._lokasi_simpan_url = finalFile.getUrl();

      // Kompensasi partial-failure: file salinan sudah dibuat di Drive. Bila penulisan
      // baris/rekap gagal, buang salinan itu agar tidak jadi file yatim tanpa baris.
      let writeResult, rekapResult;
      try {
        writeResult = SpreadsheetService.appendArchiveRow(activity, subActivity, metadata);
        SpreadsheetApp.flush();
        rekapResult = SpreadsheetService.updateRekapSummary(activity, subActivity, metadata);
      } catch (sheetError) {
        try { finalFile.setTrashed(true); } catch (cleanupError) {
          console.warn('finalizeArchive: gagal membersihkan file salinan yatim: ' + cleanupError.message);
        }
        throw sheetError;
      }

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
    payload = payload || {};
    requireAuth_(payload);
    const ctx = _resolveArchiveContext_(payload);
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
    let logWarning = '';
    try {
      const initialMetadata = ArchiveController._prepareMetadata(payload, activity, subActivity, sourceFile.getName());
      result = ArchiveController._processArchiveInLock(payload, activity, subActivity, initialMetadata, sourceFile, {
        archiveId: archiveId,
        year: year,
        sourceFileId: sourceFile.getId(),
        createdBy: archiveCreatedBy
      });
      logWarning = result.logWarning || '';
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

  // ── Stepped archive flow ──────────────────────────────────────────────
  // Pecah finalizeArchive jadi 3 panggilan berurutan agar client bisa
  // menampilkan progress nyata. Atomisitas tulis tetap dijaga oleh lock
  // di step 3. Step 1-2 tanpa lock — jika step 3 gagal, file salinan
  // yang sudah dibuat di step 2 akan di-trash otomatis (rollback).
  //
  // Step 1: Validasi metadata + cek uniqueness (tanpa lock)
  // Step 2: Copy/move file ke folder tujuan (tanpa lock, operasi lambat)
  // Step 3: Tulis baris spreadsheet + rekap + log (dengan lock)

  archiveStep_validate: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const ctx = _resolveArchiveContext_(payload);
    const year = ctx.year, config = ctx.config, activity = ctx.activity, subActivity = ctx.subActivity;

    const fields = config.fields.filter(function (field) {
      return field.activity_id === payload.activityId && isTrue_(field.is_visible_in_form);
    });
    Validator.requireMetadata(payload.metadata, fields);

    // Cek uniqueness tanpa lock (early-exit; akan dicek ulang di step 3 dalam lock).
    ArchiveController._validateUniqueMetadata(payload, activity, subActivity, null);

    // Siapkan metadata + nama file final.
    const sourceFile = DriveService.getFileFromInput(payload);
    const metadata = ArchiveController._prepareMetadata(payload, activity, subActivity, sourceFile.getName());

    // Preview nomor item berikutnya (read-only, bisa berubah di step 3).
    var autoItemNumber = false;
    if (!metadata.nomor_item_arsip) {
      autoItemNumber = true;
      const nextNum = SpreadsheetService.getNextItemNumber(activity, subActivity);
      metadata.nomor_item_arsip = String(nextNum).padStart(2, '0');
      metadata.lokasi_simpan = MetadataService.buildFinalFileName(metadata, sourceFile.getName());
    }

    const archiveId = 'ARC-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const archiveUser = AuthService.getCurrentUser(payload);

    return {
      archiveId: archiveId,
      metadata: metadata,
      autoItemNumber: autoItemNumber,
      sourceFileName: sourceFile.getName(),
      sourceFileId: sourceFile.getId(),
      createdBy: archiveUser.displayName || archiveUser.username || ''
    };
  },

  archiveStep_copyFile: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireString(payload.sourceFileId, 'Source File ID');
    Validator.requireString(payload.targetFolderId, 'Target Folder ID');
    Validator.requireString(payload.finalFileName, 'Final File Name');

    const sourceFile = DriveApp.getFileById(cleanId_(payload.sourceFileId));
    const finalFile = DriveService.copyToFinalFolder(
      sourceFile, payload.targetFolderId, payload.finalFileName, payload.year || ''
    );

    return {
      finalFileId: finalFile.getId(),
      finalFileName: finalFile.getName(),
      finalFileUrl: finalFile.getUrl()
    };
  },

  archiveStep_writeAndLog: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const ctx = _resolveArchiveContext_(payload);
    const year = ctx.year, activity = ctx.activity, subActivity = ctx.subActivity;

    Validator.requireString(payload.archiveId, 'Archive ID');
    Validator.requireString(payload.finalFileId, 'Final File ID');
    Validator.requireString(payload.sourceFileId, 'Source File ID');

    const finalFile = DriveApp.getFileById(cleanId_(payload.finalFileId));
    const metadata = payload.metadata || {};
    const createdBy = payload.createdBy || '';

    let result;
    let logWarning = '';
    try {
      result = withLock_(function () {
        // Re-validasi uniqueness di dalam lock (atomik).
        ArchiveController._validateUniqueMetadata(payload, activity, subActivity, null);

        // Re-assign nomor item final di dalam lock. Jika autoItemNumber true,
        // selalu ambil ulang nomor terbaru agar concurrent submit gak bentrok.
        if (payload.autoItemNumber || !metadata.nomor_item_arsip) {
          var nextNum = SpreadsheetService.getNextItemNumber(activity, subActivity);
          var originalNum = metadata.nomor_item_arsip;
          metadata.nomor_item_arsip = String(nextNum).padStart(2, '0');
          var newExpectedName = MetadataService.buildFinalFileName(metadata, finalFile.getName());
          
          // Jika nomor item bergeser karena antrean, RENAME file di Drive agar sesuai.
          if (originalNum !== metadata.nomor_item_arsip && finalFile.getName() !== newExpectedName) {
            try {
              finalFile.setName(newExpectedName);
            } catch (renameErr) {
              console.warn('Gagal merename file yang bergeser antreannya: ' + renameErr.message);
            }
          }
        }

        // Ambil nama final yang benar-benar tersimpan di Drive.
        metadata.lokasi_simpan = finalFile.getName();
        metadata._lokasi_simpan_url = finalFile.getUrl();

        var targetFolderId = payload.targetFolderId || subActivity.folder_id;
        var targetFolder = DriveApp.getFolderById(cleanId_(targetFolderId));
        var targetFolderInfo = getArchiveTargetFolderInfo_(targetFolder);

        var writeResult, rekapResult;
        try {
          writeResult = SpreadsheetService.appendArchiveRow(activity, subActivity, metadata);
          SpreadsheetApp.flush();
          rekapResult = SpreadsheetService.updateRekapSummary(activity, subActivity, metadata);
        } catch (sheetError) {
          // Rollback: trash file salinan yang sudah dibuat di step 2.
          try { finalFile.setTrashed(true); } catch (cleanupError) {
            console.warn('archiveStep_writeAndLog: gagal membersihkan file salinan yatim: ' + cleanupError.message);
          }
          throw sheetError;
        }

        var _logWarning = ArchiveController._logArchiveCompletion(
          payload.archiveId, year, activity, subActivity,
          payload.sourceFileId, finalFile, writeResult, createdBy,
          targetFolderInfo, metadata
        );

        return {
          finalFile: finalFile,
          writeResult: writeResult,
          rekapResult: rekapResult,
          metadata: metadata,
          targetFolder: targetFolderInfo,
          logWarning: _logWarning
        };
      }, 30000);
      logWarning = result.logWarning || '';
    } catch (error) {
      ConfigRepository.appendArchiveLog({
        archive_id: payload.archiveId, year: year,
        activity_id: activity.activity_id, sub_activity_id: subActivity.sub_activity_id,
        source_file_id: payload.sourceFileId, final_file_id: payload.finalFileId,
        final_file_name: finalFile.getName(),
        target_folder_id: payload.targetFolderId || subActivity.folder_id || '',
        target_folder_name: '', target_folder_path: '',
        spreadsheet_file_id: '', spreadsheet_row_number: '',
        status: STATUS.FAILED, created_at: new Date().toISOString(),
        created_by: createdBy, error_message: error.message || String(error),
        metadata_json: JSON.stringify({ failedPayload: metadata })
      });
      // Rollback: trash file dari step 2.
      try { finalFile.setTrashed(true); } catch (cleanupErr) {
        console.warn('archiveStep_writeAndLog: rollback trash gagal: ' + cleanupErr.message);
      }
      CacheHelper.invalidate(year);
      throw error;
    }

    return {
      archiveId: payload.archiveId, status: STATUS.COMPLETED,
      finalFile: DriveService.fileToDto(result.finalFile),
      spreadsheet: result.writeResult, rekapSpreadsheet: result.rekapResult,
      metadata: result.metadata, warning: logWarning || ''
    };
  },

  adoptExistingArchives: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
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
    payload = payload || {};
    requireAuth_(payload);
    const ctx = _resolveArchiveContext_(payload);
    const activity = ctx.activity, subActivity = ctx.subActivity;

    const existingData = SpreadsheetService.listExistingArchiveRows(activity, subActivity);
    const errors = [];

    const incomingItem = String(payload.nomor_item_arsip || '').trim();
    if (incomingItem) {
      const normalizedItem = incomingItem.replace(/^0+/, '').padStart(2, '0');
      existingData.rows.forEach(function (row) {
        const existingItem = String(row.metadata.nomor_item_arsip || '').trim();
        const normalizedExisting = existingItem ? existingItem.replace(/^0+/, '').padStart(2, '0') : '';
        if (normalizedExisting === normalizedItem) {
          errors.push({ field: 'nomor_item_arsip', message: 'Nomor Item Arsip "' + incomingItem + '" sudah digunakan.' });
        }
      });
    }

    return { valid: errors.length === 0, errors: errors };
  },

  editMetadata: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const ctx = _resolveArchiveContext_(payload);
    const year = ctx.year, config = ctx.config, activity = ctx.activity, subActivity = ctx.subActivity;
    Validator.requireString(payload.archiveId, 'Archive ID');

    const fields = config.fields.filter(function (field) { return field.activity_id === payload.activityId && isTrue_(field.is_visible_in_form); });
    Validator.requireMetadata(payload.metadata, fields);

    const log = ConfigRepository.getArchiveLog(payload.archiveId);
    if (!log) throw new Error('Log arsip tidak ditemukan.');

    return withLock_(() => {
      // B4: jangan percaya nomor baris yang dibaca DI LUAR lock. Hapus/reorder arsip
      // lain oleh user lain bisa menggeser baris (decrementArchiveLogRows menyelaraskan
      // nomor baris di log), jadi muat ulang log DI DALAM lock dan pakai nilai segar
      // agar tidak menimpa baris milik arsip lain.
      const freshLog = ConfigRepository.getArchiveLog(payload.archiveId) || log;
      ArchiveController._validateUniqueMetadata(payload, activity, subActivity, Number(freshLog.spreadsheet_row_number));

      const metadata = ArchiveController._prepareMetadata(payload, activity, subActivity, freshLog.final_file_name);
      let finalFileDto = null;
      let finalFileName = freshLog.final_file_name;

      let fileUrl = '';
      // B5: simpan nama lama; jika penulisan baris/rekap gagal SETELAH file di-rename,
      // kembalikan nama file ke semula agar Drive tidak desync dengan sheet/log.
      let renamedFile = null;
      const previousFileName = freshLog.final_file_name;
      if (freshLog.final_file_id) {
        try {
          const file = DriveApp.getFileById(freshLog.final_file_id);
          fileUrl = file.getUrl();
          if (metadata.lokasi_simpan !== freshLog.final_file_name) {
            file.setName(metadata.lokasi_simpan);
            renamedFile = file;
            finalFileDto = DriveService.fileToDto(file);
            finalFileName = file.getName();
          }
        } catch (e) {
          console.warn('Failed to handle drive file: ' + e.message);
          metadata.lokasi_simpan = freshLog.final_file_name;
        }
      }
      metadata._lokasi_simpan_url = fileUrl;

      try {
        if (metadata._resolved_row_number) {
          SpreadsheetService.updateArchiveRow(activity, subActivity, parseInt(metadata._resolved_row_number, 10), metadata);
        } else if (freshLog.spreadsheet_row_number) {
          SpreadsheetService.updateArchiveRow(activity, subActivity, parseInt(freshLog.spreadsheet_row_number, 10), metadata);
        }
        SpreadsheetService.updateRekapSummary(activity, subActivity, metadata);
      } catch (sheetError) {
        if (renamedFile) {
          try { renamedFile.setName(previousFileName); } catch (restoreError) {
            console.warn('editMetadata: gagal mengembalikan nama file setelah error sheet: ' + restoreError.message);
          }
        }
        throw sheetError;
      }

      let newMetadataJson = freshLog.metadata_json;
      try {
        const draftStateObj = JSON.parse(freshLog.metadata_json);
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

      let editWarning = '';
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

  parseDocumentContent: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireString(payload.fileId, 'File ID');
    // IDOR: jangan biarkan client memparse (membaca isi) file sembarang. Batasi ke
    // file di dalam ruang kerja. Penting di USER_DEPLOYING (script baca sbg pemilik).
    const psYear = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    requireWithinWorkspace_(payload.fileId, psYear);

    // Cache hasil parse per fileId: buka file yang sama lagi = instan, tak konversi ulang.
    const parseCacheKey = 'parsecache_v26_' + cleanId_(payload.fileId);
    const parseLockKey = 'parselock_v26_' + cleanId_(payload.fileId);
    try {
      const cachedParse = CacheService.getScriptCache().get(parseCacheKey);
      if (cachedParse) { const r = JSON.parse(cachedParse); r.cached = true; return r; }
    } catch (e) { /* cache opsional */ }

    // Mutex lunak: bila konversi/OCR file ini sedang berjalan di eksekusi lain, tunggu
    // sebentar lalu pakai hasil cache-nya — cegah konversi+OCR ganda (boros kuota Drive).
    try {
      const _c = CacheService.getScriptCache();
      for (let w = 0; w < 4 && _c.get(parseLockKey); w++) {
        Utilities.sleep(1500);
        const again = _c.get(parseCacheKey);
        if (again) { const r = JSON.parse(again); r.cached = true; return r; }
      }
    } catch (e) { /* mutex opsional */ }

    const startTime = Date.now();
    const file = DriveApp.getFileById(cleanId_(payload.fileId));
    const mimeType = file.getMimeType();
    const fileName = file.getName();
    const fileSize = file.getSize();

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

    // Pasang sentinel mutex (TTL pendek; auto-kedaluwarsa bila eksekusi crash).
    try { CacheService.getScriptCache().put(parseLockKey, '1', 90); } catch (e) { /* sentinel opsional */ }

    let text = '';
    let pageCount = 0;
    if (mimeType === 'application/pdf') {
      pageCount = extractPdfPageCount_(file);
    }
    const tempDocId = '';

    // Try multiple extraction methods
    const extractionMethods = [];
    const extractionErrors = [];

    try {
      // Method 1: Direct text extraction (works for text/plain files)
      if (mimeType === 'text/plain') {
        text = file.getBlob().getDataAsString();
        extractionMethods.push('direct_text');
      }
    } catch (e) {
      console.warn('Method 1 failed: ' + e.message);
    }

    // Method 2: ekstrak ISI dokumen via konversi Drive (OCR untuk PDF scan, convert untuk Word).
    // Inilah yang membuat autofill membaca isi surat sungguhan, bukan sekadar nama file.
    if (!text && mimeType !== 'text/plain') {
      const conv = extractTextViaConversion_(file, mimeType);
      if (conv.text && conv.text.trim()) {
        text = conv.text;
        extractionMethods.push(conv.method);
      } else if (conv.method) {
        extractionErrors.push('konversi: ' + conv.method);
      }
    }

    // Fallback to filename if nothing extracted
    if (!text) {
      text = fileName;
      extractionMethods.push('filename_only');
    }



    // Cleanup temp doc
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (e) {}
    }

    // Fallback: If pageCount couldn't be extracted natively (e.g. compressed PDF)
    // estimate it from the form-feeds (\f) or \x0c in the converted text.
    if (pageCount === 0 && text) {
      const ffMatches = text.match(/\x0c/g);
      if (ffMatches && ffMatches.length > 0) {
        pageCount = ffMatches.length + 1;
      } else {
        // Rough estimation based on text length (~1500 chars per page on average)
        pageCount = Math.max(1, Math.ceil(text.length / 1500));
      }
    }


    // Run ParseEngine for scored multi-pass extraction
    const engineResult = ParseEngine.analyze(text, fileName, { activity: {}, subActivity: {} });
    const fields = engineResult.fields;

    // Use file modified date as fallback for date field
    if (!fields.tanggal) {
      try {
        const lastUpdated = file.getLastUpdated();
        const lastUpdatedStr = Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

    if (pageCount > 0) {
      fields.jumlah = {
        value: String(pageCount),
        score: 1.0,
        confidence: 'high',
        source: 'pdf_page_count'
      };
    }

    // Ringkasan diagnostik — TIDAK menyertakan teks dokumen mentah (rawText/ocrTextSample)
    // agar isi dokumen tidak bocor ke payload client maupun ke Stackdriver.
    const debugInfo = {
      errors: extractionErrors,
      ocrTextLength: text ? text.length : 0,
      extractionMethods: extractionMethods,
      structure: engineResult.structure,
      documentType: engineResult.documentType,
      fieldsFound: Object.keys(fields),
      fieldsMissing: ['nomor_surat','kode_klasifikasi','tanggal','uraian_informasi_item','klasifikasi_akses','dari','kepada','tanda_tangan','lampiran']
        .filter(function(k) { return !fields[k]; })
    };

    // Flatten complex fields for client compatibility
    if (fields.tanda_tangan && fields.tanda_tangan.value) {
      const tt = fields.tanda_tangan.value;
      const ttParts = [];
      if (tt.jabatan) ttParts.push(tt.jabatan);
      if (tt.nama) ttParts.push(tt.nama);
      fields.tanda_tangan.value = ttParts.join(', ') || ttParts.join('');
    }
    if (fields.dari && typeof fields.dari.value === 'object') {
      delete fields.dari; // skip if no string value
    }

    // Add page count (estimated separately from OCR)
    if (pageCount > 0) {
      fields.jumlah = { value: String(pageCount), score: 0.6, confidence: 'medium', source: 'page_estimation' };
    }

    const result = {
      fields: fields,
      rawTextLength: engineResult.rawTextLength || text.length,
      parseDuration: Date.now() - startTime,
      fieldCount: Object.keys(fields).length,
      totalFields: engineResult.totalFields || 9,
      documentType: engineResult.documentType || '',
      documentDirection: engineResult.documentDirection || 'masuk',
      extractionMethods: extractionMethods,
      structure: engineResult.structure || {},
      debug: debugInfo
    };

    // Simpan ke cache (6 jam). Lewati diam-diam bila terlalu besar untuk CacheService.
    try {
      CacheService.getScriptCache().put(parseCacheKey, JSON.stringify(result), 21600);
    } catch (e) { /* >100KB atau cache penuh: abaikan */ }

    try { CacheService.getScriptCache().remove(parseLockKey); } catch (e) { /* lepas mutex */ }
    return result;
  },

  bulkAddArchiveDocumentLinks: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    const items = payload.items || [];
    if (!items.length) return { successCount: 0, failCount: 0, errors: [] };
    
    let spreadsheet = null;
    let activity = null;
    let subActivity = null;
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    
    if (payload.activityId && payload.subActivityId) {
      const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      const config = CacheHelper.getConfig(year);
      activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');
      subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
      if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
    }

    const files = [];
    items.forEach(function(item) {
      let createdFile = null;
      try {
        requireSafeUrl_(item.url, 'URL dokumen' + (item.name ? ' ' + item.name : ''));
        if (item.parentFolderId) requireWithinWorkspace_(item.parentFolderId, payload.year);
        createdFile = DriveService.addArchiveDocumentLink(item);
        if (activity && subActivity) {
          const linkFile = { name: createdFile.name, url: item.url || createdFile.url };
          spreadsheet = SpreadsheetService.updateArchiveDocumentLink(activity, subActivity, item.name, linkFile);
        }
        files.push(createdFile);
        successCount++;
      } catch (e) {
        // Kompensasi: bila shortcut sudah dibuat tetapi rekap gagal, buang shortcut yatim.
        if (createdFile && createdFile.id) {
          try { DriveApp.getFileById(createdFile.id).setTrashed(true); } catch (cleanupError) {
            console.warn('bulkAddArchiveDocumentLinks: gagal membersihkan shortcut yatim: ' + cleanupError.message);
          }
        }
        failCount++;
        errors.push(item.name + ': ' + e.message);
      }
    });

    bumpVersion();
    auditAction_(actor, 'DOCLINK_BULK_ADDED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      message: 'Tambah tautan dokumen massal: ' + successCount + ' sukses, ' + failCount + ' gagal'
    });
    return { successCount: successCount, failCount: failCount, errors: errors, files: files };
  },
  addArchiveDocumentLink: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.parentFolderId, 'Folder induk');
    Validator.requireString(payload.name, 'Kategori dokumen');
    requireSafeUrl_(payload.url, 'Link Google Drive');
    requireWithinWorkspace_(payload.parentFolderId, payload.year);
    const file = DriveService.addArchiveDocumentLink(payload);
    let spreadsheet = null;
    if (payload.activityId && payload.subActivityId) {
      // Kompensasi: shortcut sudah dibuat di Drive. Bila penulisan rekap gagal,
      // buang shortcut agar tidak jadi tautan yatim tanpa entri rekap.
      try {
        const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
        const config = CacheHelper.getConfig(year);
        const activity = ConfigService.findActivity(config, payload.activityId);
        if (!activity) throw new Error('Kegiatan tidak ditemukan.');
        const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
        if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
        const linkFile = { name: file.name, url: payload.url || file.url };
        spreadsheet = SpreadsheetService.updateArchiveDocumentLink(activity, subActivity, payload.name, linkFile);
      } catch (linkError) {
        if (file && file.id) {
          try { DriveApp.getFileById(file.id).setTrashed(true); } catch (cleanupError) {
            console.warn('addArchiveDocumentLink: gagal membersihkan shortcut yatim: ' + cleanupError.message);
          }
        }
        throw linkError;
      }
    }
    bumpVersion();
    auditAction_(actor, 'DOCLINK_ADDED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      folderId: payload.parentFolderId, message: 'Menambah tautan dokumen: ' + payload.name
    });
    return Object.assign({}, file, { spreadsheet: spreadsheet });
  },
  getShortcutTargetInfo: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    if (payload.fileId) requireWithinWorkspace_(payload.fileId, payload.year);
    return DriveService.getShortcutTargetInfo(payload);
  },
  updateArchiveDocumentLink: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    requireSafeUrl_(payload.url, 'URL Drive');
    requireWithinWorkspace_(payload.fileId, payload.year);
    const result = DriveService.updateArchiveDocumentLink(payload);
    if (payload.activityId && payload.subActivityId && payload.categoryName) {
      const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
      const config = CacheHelper.getConfig(year);
      const activity = ConfigService.findActivity(config, payload.activityId);
      if (!activity) throw new Error('Kegiatan tidak ditemukan.');
      const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
      if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
      const linkFile = { name: result.name, url: payload.url || result.url };
      SpreadsheetService.updateArchiveDocumentLink(activity, subActivity, payload.categoryName, linkFile);
    }
    bumpVersion();
    auditAction_(actor, 'DOCLINK_UPDATED', {
      year: payload.year, activityId: payload.activityId, subActivityId: payload.subActivityId,
      folderId: payload.fileId, message: 'Memperbarui tautan dokumen: ' + (payload.categoryName || '-')
    });
    return result;
  },
  syncExistingPhysicalFiles: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const actor = auditActor_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    
    // We pass startTime to prevent 6-minute timeout
    const startTime = Date.now();
    const report = [];
    
    const result = wsSyncExistingFilesInFolder_(year, report, startTime);
    
    auditAction_(actor, 'FILES_SYNCED', { year: year, message: 'Sinkronisasi berkas fisik. Total tersinkronisasi: ' + result.totalSynced });
    
    return {
      success: true,
      year: year,
      totalSynced: result.totalSynced,
      timeLimitReached: result.timeLimitReached,
      report: report
    };
  },
  getArchiveLogByFileId: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    const log = ConfigRepository.getArchiveLogByFileId(payload.fileId);
    return log || null;
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
  const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
  Validator.requireString(payload.activityId, 'Activity ID');
  Validator.requireString(payload.subActivityId, 'Sub Activity ID');
  const config = CacheHelper.getConfig(year);
  const activity = ConfigService.findActivity(config, payload.activityId);
  if (!activity) throw new Error('Kegiatan tidak ditemukan.');
  const subActivity = ConfigService.findSubActivity(config, payload.activityId, payload.subActivityId);
  if (!subActivity) throw new Error('Sub-kegiatan tidak ditemukan.');
  return { year: year, config: config, activity: activity, subActivity: subActivity };
}

/** Ambang baca-penuh untuk hitung halaman; di atas ini pakai byte-range agar tak OOM. */
const PDF_PAGECOUNT_FULL_READ_MAX_BYTES = 25 * 1024 * 1024; // 25MB

function extractPdfPageCount_(file) {
  try {
    const size = (file.getSize && file.getSize()) || 0;

    // File kecil/menengah: baca penuh. Akurat — bisa pakai fallback hitung /Type /Page.
    if (!size || size <= PDF_PAGECOUNT_FULL_READ_MAX_BYTES) {
      const pdfText = file.getBlob().getDataAsString('ISO-8859-1');
      return countPagesFromPdfChunk_(pdfText, true);
    }

    // File besar (mis. surat scan/SRIKANDI ratusan MB): JANGAN muat seluruh blob —
    // getDataAsString akan OOM / timeout. Ambil hanya potongan awal & akhir via
    // byte-range. /Count di root /Pages (total halaman dalam satu angka) hampir
    // selalu tertangkap di sini tanpa membaca semua byte.
    const CHUNK = 3 * 1024 * 1024; // 3MB per potongan
    const head = fetchDriveBytesRange_(file.getId(), 0, CHUNK - 1);
    let count = countPagesFromPdfChunk_(head, false);
    if (count > 0) return count;

    const tailStart = Math.max(0, size - CHUNK);
    if (tailStart > 0) {
      const tail = fetchDriveBytesRange_(file.getId(), tailStart, size - 1);
      count = countPagesFromPdfChunk_(tail, false);
      if (count > 0) return count;
    }
  } catch (e) {
    console.warn('extractPdfPageCount_ failed: ' + e.message);
  }
  return 0;
}

/**
 * Hitung jumlah halaman dari sepotong teks mentah PDF.
 * @param {string} pdfText  Isi PDF (ISO-8859-1) — penuh atau sebagian.
 * @param {boolean} allowPageTypeCount  true hanya bila pdfText = SELURUH file,
 *   karena hitung /Type /Page perlu seluruh isi (kalau sebagian → undercount).
 * @return {number} jumlah halaman, atau 0 bila tak terdeteksi.
 */
function countPagesFromPdfChunk_(pdfText, allowPageTypeCount) {
  if (!pdfText) return 0;

  // Anchor ke dict page-tree: ambil /Count yang menempel pada /Type /Pages, supaya
  // tidak tertukar dengan /Count milik outline/bookmark. Urutan key bisa
  // "/Type /Pages ... /Count" atau "/Count ... /Type /Pages", jadi cek dua arah dalam
  // jendela pendek. Ambil nilai TERBESAR: root page-tree memuat total halaman, selalu
  // >= count tiap sub-tree (mis. pada PDF hasil merge/incremental update).
  let best = 0;
  // Jendela [^<>] = tidak boleh menyeberangi batas dictionary PDF (<< >>), supaya
  // /Count dan /Type /Pages yang ter-anchor benar-benar berada di dict yang sama —
  // mencegah /Count milik outline tersangkut ke /Pages di dict tetangga.
  const W = 250;
  const patterns = [
    new RegExp('/Type\\s*/Pages\\b[^<>]{0,' + W + '}?/Count\\s+(\\d+)', 'g'),
    new RegExp('/Count\\s+(\\d+)[^<>]{0,' + W + '}?/Type\\s*/Pages\\b', 'g')
  ];
  for (let p = 0; p < patterns.length; p++) {
    let m;
    while ((m = patterns[p].exec(pdfText))) {
      const n = parseInt(m[1], 10);
      if (n > 0 && n < 5000 && n > best) best = n;
    }
  }
  if (best > 0) return best;

  // Cadangan: /Count pertama mana pun. Bisa keliru ke outline, tapi lebih baik dari 0.
  const anyCount = pdfText.match(/\/Count\s+(\d+)/);
  if (anyCount && anyCount[1]) {
    const n = parseInt(anyCount[1], 10);
    if (n > 0 && n < 5000) return n;
  }

  // Cadangan terakhir: hitung object /Type /Page — hanya valid bila punya seluruh isi.
  if (allowPageTypeCount) {
    const pageMatch = pdfText.match(/\/Type\s*\/Page\b/g);
    if (pageMatch && pageMatch.length > 0) return pageMatch.length;
  }
  return 0;
}

/**
 * Unduh sebagian byte file Drive via Range header (tanpa memuat seluruh blob).
 * Pakai endpoint alt=media + token OAuth (cukup scope Drive yang sudah ada).
 * @return {string} potongan byte sebagai string ISO-8859-1, atau '' bila gagal.
 */
function fetchDriveBytesRange_(fileId, start, end) {
  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://www.googleapis.com/drive/v3/files/' + cleanId_(fileId) +
      '?alt=media&supportsAllDrives=true';
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, Range: 'bytes=' + start + '-' + end },
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code === 200 || code === 206) return resp.getBlob().getDataAsString('ISO-8859-1');
  } catch (e) {
    console.warn('fetchDriveBytesRange_ failed: ' + e.message);
  }
  return '';
}
