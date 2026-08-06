'use strict';

function wsSyncExistingFilesInFolder_(year, report, startTime) {
  const config = CacheHelper.getConfig(year);
  const existingKeys = ConfigRepository.getArchiveLogKeyMap(year);
  
  let totalSynced = 0;
  const allLogsToAppend = [];
  
  const HARD_BUDGET_MS = 4.5 * 60 * 1000; // 4.5 minutes
  let timeLimitReached = false;

  config.activities.forEach(function (activity) {
    if (timeLimitReached) return;

    const subActivities = config.subActivities.filter(function (sub) {
      return sub.activity_id === activity.activity_id;
    });

    subActivities.forEach(function (subActivity) {
      if (timeLimitReached) return;
      
      if (startTime && (Date.now() - startTime) > HARD_BUDGET_MS) {
        timeLimitReached = true;
        return;
      }

      let pageToken = null;
      do {
        if (startTime && (Date.now() - startTime) > HARD_BUDGET_MS) {
          timeLimitReached = true;
          break;
        }

        let result;
        try {
          result = Drive.Files.list({
            q: "'" + subActivity.folder_id + "' in parents and trashed = false",
            fields: "nextPageToken, files(id, name, webViewLink)",
            pageSize: 1000,
            pageToken: pageToken,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
          });
        } catch (e) {
          break;
        }
        
        const files = result.files || [];
        const metadataList = [];
        const currentFiles = [];

        for (let j = 0; j < files.length; j++) {
          const file = files[j];
          const fileId = file.id;
          if (existingKeys['file:' + fileId]) continue; // Already mapped!
          
          const fileName = file.name;
          
          // Hanya import file yg namanya sesuai format: "Nomor. (Tingkat) Uraian..." 
          const isFormatted = /^\d{1,4}\.\s*\([^)]+\)/.test(fileName);
          if (!isFormatted) continue;

          const metadata = MetadataService.parseExistingFileName(fileName, activity, subActivity);
          metadata.lokasi_simpan = fileName;
          metadata._lokasi_simpan_url = file.webViewLink;

          // Skip kalo hasil parse gak bermutu (uraian kosong atau cuma raw filename)
          const rawName = fileName.replace(/\.[a-z0-9]+$/i, '').trim();
          const uraian = metadata.uraian_informasi_item || metadata.uraian_informasi_item;
          if (!uraian || uraian.length < 3 || uraian === rawName) continue;


          metadataList.push(metadata);
          currentFiles.push({ id: fileId, name: fileName });
          existingKeys['file:' + fileId] = true;
        }

        if (metadataList.length > 0) {
          totalSynced += metadataList.length;
          // Bulk write to Spreadsheet — lock sempit: cegah interleave dgn finalizeArchive
          // pada detail/rekap sheet yang sama (race pemilihan baris). Reentrant per-eksekusi.
          const writeResults = withLock_(function () {
            const wr = SpreadsheetService.appendArchiveRowsBulk(activity, subActivity, metadataList);
            SpreadsheetService.updateRekapSummary(activity, subActivity, {}); // Update rekap once per sub activity
            return wr;
          });

          // Log them
          for (let i = 0; i < metadataList.length; i++) {
            const fileInfo = currentFiles[i];
            const writeResult = writeResults[i];
            const metadata = metadataList[i];
            const archiveId = 'SYNC-' + Utilities.getUuid().slice(0, 8).toUpperCase();
            
            allLogsToAppend.push({
              archive_id: archiveId,
              year: year,
              activity_id: activity.activity_id,
              sub_activity_id: subActivity.sub_activity_id,
              source_file_id: '',
              final_file_id: fileInfo.id,
              final_file_name: fileInfo.name,
              target_folder_id: subActivity.folder_id,
              target_folder_name: subActivity.sub_activity_name || subActivity.folder_path,
              target_folder_path: '',
              spreadsheet_file_id: writeResult.spreadsheetId,
              spreadsheet_row_number: writeResult.rowNumber,
              status: STATUS.COMPLETED,
              created_at: new Date().toISOString(),
              created_by: 'system',
              error_message: '',
              metadata_json: JSON.stringify({ importedFromExisting: true, metadata: metadata, finalFileName: fileInfo.name })
            });
          }
        }

        pageToken = result.nextPageToken;
      } while (pageToken);
    });
  });
  
  if (allLogsToAppend.length > 0) {
    ConfigRepository.appendArchiveLogsBulk(allLogsToAppend);
  }
  
  if (timeLimitReached) {
    wsPushReport_(report, 'warning', 'Waktu eksekusi hampir habis. Sebagian berkas mungkin belum tersinkronisasi. Silakan     jalankan Sinkronisasi Ruang Kerja ulang nanti.');
  } else if (totalSynced > 0) {
    wsPushReport_(report, 'created', 'Berhasil mensinkronisasi ' + totalSynced + ' file fisik yang sudah ada di Drive secara     bulk.');
  }
  
  return {
    totalSynced: totalSynced,
    timeLimitReached: timeLimitReached
  };
}

