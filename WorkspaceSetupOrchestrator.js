'use strict';

/** @private Workspace setup orchestration and year lifecycle. */
const WorkspaceSetupImpl_ = {
  parseLeadingNumber: function (value) {
    const match = String(value || '').trim().match(/^(\d+)/);
    return match ? String(Number(match[1])) : '';
  },

  initialize: function (payload) {
    const startTime = Date.now();
    const year = Number(payload.year || DEFAULT_YEAR);
    const rootFolderId = cleanId_(payload.workspaceRootFolderId || payload.workspaceRootFolderUrl);
    if (!rootFolderId) throw new Error('Pilih/isi link Folder Workspace terlebih dahulu.');

    const root = DriveApp.getFolderById(rootFolderId);
    const report = [];

    const systemFolderName = String(payload.systemFolderName || WORKSPACE_CONFIG.systemFolderName).trim();
    const customParentId = cleanId_(payload.systemFolderParentId);
    const systemSettings = {
      configFolderId: cleanId_(payload.configFolderId),
      systemFolderParentId: customParentId,
      systemFolderName: systemFolderName
    };
    let configFolderId = cleanId_(payload.configFolderId);
    const systemFolder = DriveService.resolveSystemFolder(systemSettings, root);
    if (configFolderId && systemFolder.getId() !== configFolderId) {
      configFolderId = '';
    }

    const inbox = wsGetOrCreateFolder_(systemFolder, WORKSPACE_CONFIG.inboxFolderName, report);
    wsGetOrCreateFolder_(inbox, 'PDF', report);
    wsGetOrCreateFolder_(inbox, 'DOCX', report);
    wsGetOrCreateFolder_(inbox, 'Lainnya', report);
    const templateFolder = wsGetOrCreateFolder_(systemFolder, WORKSPACE_CONFIG.templateFolderName, report);

    const daftarArsip = wsFindOrCreateChildFolder_(root, ['1. Daftar Arsip', 'Daftar Arsip'], '1. Daftar Arsip (Spreadsheet)    ', report);
    const naskahDinas = wsFindOrCreateChildFolder_(root, ['2. Naskah Dinas Latbang', 'Naskah Dinas LitBang', 'Naskah Dinas'],     '2. Naskah Dinas Latbang (Dokumen)', report);
    // Folder arsip harus tetap Restricted; ACL eksplisit Google Drive tetap berlaku.
    // Menonaktifkan link sharing tidak mengubah akses yang diberikan langsung ke akun tertentu.
    wsEnsureRestrictedSharing_(daftarArsip, report);
    wsEnsureRestrictedSharing_(naskahDinas, report);
    const persuratan = wsFindOrCreateChildFolder_(naskahDinas, ['1. Persuratan', 'Persuratan'], '1. Persuratan', report);

    const configSpreadsheet = wsGetOrCreateConfigSpreadsheet_(systemFolder);
    
    // Save settings early to fix "CONFIG_SPREADSHEET_ID belum diisi"
    ConfigService.saveSettings(Object.assign({}, ConfigService.getSettings() || {}, {
      configSpreadsheetId: configSpreadsheet.getId(),
      currentYear: year,
      workspaceRootFolderId: root.getId(),
      referenceRootFolderId: '',
      systemFolderParentId: customParentId || '',
      systemFolderName: systemFolderName,
      configFolderId: configFolderId || systemFolder.getId()
    }));
    
    // 1. Initialize the requested year
    this.initializeSingleYear_(configSpreadsheet, year, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder,     report);

    // 2. Scan and auto-import any other years that physically exist in Google Drive
    this.scanAndImportPhysicalYears_(configSpreadsheet, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder,     report, startTime);

    return {
      year: year, rootFolderId: root.getId(), rootFolderUrl: root.getUrl(),
      referenceRootFolderId: '', systemFolderParentId: customParentId || '',
      systemFolderName: systemFolderName, configFolderId: configFolderId || systemFolder.getId(),
      configSpreadsheetId: configSpreadsheet.getId(), configSpreadsheetUrl: configSpreadsheet.getUrl(),
      report: report
    };
  },

  initializeSingleYear_: function(ss, year, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report) {
    const arsipDiklat = wsFindOrCreateChildFolder_(daftarArsip, ['ARSIP DIKLAT ' + year, 'DIKLAT ' + year, String(year)],     'ARSIP DIKLAT ' + year, report);
    const tahunFolder = wsFindOrCreateChildFolder_(persuratan, ['Tahun ' + year, String(year)], 'Tahun ' + year, report);
    const dokumenFolder = wsFindOrCreateChildFolder_(naskahDinas, ['2. Dokumen', 'Dokumen'], '2. Dokumen', report);
    const dokumenTahunFolder = wsFindOrCreateChildFolder_(dokumenFolder, ['Tahun ' + year, String(year)], 'Tahun ' + year,     report);

    const activityRows = [];
    const subActivityRows = [];
    const existingSubMap = wsExistingSubActivityMap_(ss, year);
    
    let globalSortOrder = 1;
    Object.keys(existingSubMap).forEach(function(key) {
      const sub = existingSubMap[key];
      if (sub && sub.sort_order) {
        const so = Number(sub.sort_order);
        if (so >= globalSortOrder) globalSortOrder = so + 1;
      }
    });

    WORKSPACE_ACTIVITIES.forEach(function (activity) {
      const laciFolder = wsFindOrCreateChildFolder_(arsipDiklat, activity.laciCandidates, activity.laciFolderName, report);
      const targetFolder = wsFindOrCreateChildFolder_(tahunFolder, activity.targetCandidates, activity.targetFolderName,       report);
      const docTargetFolder = wsFindOrCreateChildFolder_(dokumenTahunFolder, activity.targetCandidates, activity.      targetFolderName, report);
      const spreadsheetContext = wsResolveActivitySpreadsheetContext_(laciFolder, activity, year, report);
      const spreadsheet = spreadsheetContext.defaultSpreadsheet;

      activityRows.push([year, activity.id, activity.label, activity.laciNo, activity.folderNo, spreadsheet.getId(),       targetFolder.getId(), laciFolder.getId(), 'TRUE', activity.sortOrder]);

      const existingSubFolders = wsListChildFolders_(targetFolder);
      const defaultSubActivities = wsDefaultSubActivityNames_(activity, year);
      if (!existingSubFolders.length && defaultSubActivities.length > 0) {
        defaultSubActivities.forEach(function (subName) {
          wsFindOrCreateChildFolder_(targetFolder, [subName], subName, report);
          wsFindOrCreateChildFolder_(docTargetFolder, [subName], subName, report);
        });
      }

      const subActivityEntries = wsBuildSubActivityEntries_(
        root,
        naskahDinas,
        persuratan,
        tahunFolder,
        targetFolder,
        activity
      );
      subActivityEntries.sort(function (a, b) {
        return compareSubActivitiesByLocalOrder_({
          activity_id: activity.id,
          sub_activity_id: a.subActivityId,
          sub_activity_name: a.folder.getName(),
          parent_folder_name: a.parentFolderName
        }, {
          activity_id: activity.id,
          sub_activity_id: b.subActivityId,
          sub_activity_name: b.folder.getName(),
          parent_folder_name: b.parentFolderName
        }, { activity_id: activity.id });
      });

      subActivityEntries.forEach(function (entry, index) {
        const subFolder = entry.folder;
        if (entry.parentFolderName) {
          const docParentGroup = wsFindOrCreateChildFolder_(docTargetFolder, [entry.parentFolderName], entry.          parentFolderName, report);
          wsFindOrCreateChildFolder_(docParentGroup, [subFolder.getName()], subFolder.getName(), report);
        } else {
          wsFindOrCreateChildFolder_(docTargetFolder, [subFolder.getName()], subFolder.getName(), report);
        }
        const noFolder = WorkspaceSetupService.parseLeadingNumber(subFolder.getName()) || activity.folderNo;
        const subSpreadsheet = wsPickSpreadsheetForSubActivity_(spreadsheetContext, activity, entry.groupName || subFolder.        getName());
        const existingSub = existingSubMap[activity.id + '|' + entry.subActivityId] || {};
        const formalArchiveName = existingSub.formal_archive_name ||
          wsSuggestFormalArchiveName_(activity, subFolder.getName());
        const targetSheetName = existingSub.target_sheet_name ||
          wsSuggestDetailSheetName_(activity, subFolder.getName(), subSpreadsheet);
        const mappingStatus = existingSub.mapping_status ||
          wsInferMappingStatus_(subFolder.getName(), formalArchiveName, targetSheetName);
        subActivityRows.push([
          year,
          activity.id,
          entry.subActivityId,
          subFolder.getName(),
          formalArchiveName,
          subFolder.getId(),
          entry.folderPath,
          noFolder,
          activity.defaultCode || DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI,
          activity.allowNonLetter ? 'TRUE' : 'FALSE',
          'TRUE',
          existingSub.sort_order ? Number(existingSub.sort_order) : (globalSortOrder++),
          targetSheetName,
          mappingStatus,
          existingSub.mapping_note || '',
          existingSub.rekap_row_number || '',
          '',
          '',
          entry.parentFolderId,
          entry.parentFolderName,
          entry.parentFolderPath,
          subSpreadsheet ? subSpreadsheet.getId() : '',
          '',
          index + 1
        ]);
      });
    });

    wsWriteConfig_(ss, year, root, null, daftarArsip, arsipDiklat, tahunFolder, inbox, templateFolder, activityRows,     subActivityRows);

    // Bersihkan log arsip yang sub-kegiatannya sudah tidak ada (karena pindah struktur folder)
    wsCleanupOrphanedArchiveLogs_(ss, year, subActivityRows);
    
    try {
      CacheHelper.invalidate(year);
      // NOTE: wsSyncExistingFilesInFolder_ is removed from here to prevent 6-minute timeout.
      // It must be called explicitly by the client after initWorkspace succeeds.
    } catch(e) {
      console.warn('Gagal invalidate cache otomatis: ' + e.message);
    }
  },

  scanAndImportPhysicalYears_: function(ss, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report,   startTime) {
    // Read years that are already registered in config_years sheet
    let registeredYears = [];
    try {
      const yearsSheet = ss.getSheetByName(CONFIG_SHEETS.YEARS);
      if (yearsSheet && yearsSheet.getLastRow() > 1) {
        const values = yearsSheet.getRange(2, 1, yearsSheet.getLastRow() - 1, 1).getValues();
        registeredYears = values.map(function(row) { return Number(row[0]); });
      }
    } catch(e) {
      console.error('Error reading registered years: ' + e.message);
    }

    const detectedYears = [];

    // Helper to scan a folder and collect 4-digit years
    const scanFolder = function(parentFolder) {
      if (!parentFolder) return;
      const items = listAllChildren_(
        "'" + parentFolder.getId() + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        'nextPageToken, files(name)'
      );
      for (let i = 0; i < items.length; i++) {
        const folderName = items[i].name;
        const match = folderName.match(/\b(202\d|203[0-5])\b/);
        if (match) {
          const detectedYear = Number(match[1]);
          if (detectedYears.indexOf(detectedYear) === -1) {
            detectedYears.push(detectedYear);
          }
        }
      }
    };

    // Scan both potential sources for physical years
    scanFolder(daftarArsip); // Scans for "ARSIP DIKLAT [Year]" under Daftar Arsip
    scanFolder(persuratan);  // Scans for "Tahun [Year]" under Persuratan

    // Process detected years.
    // Batas eksekusi GAS ~6 menit. Sisakan ~1 menit untuk finalisasi + return,
    // dan jangan mulai tahun baru kalau sisa waktu < estimasi durasi satu tahun
    // (+20% margin). Init bersifat idempoten, jadi tahun yang belum sempat
    // diimport bisa dilanjutkan dengan menjalankan ulang "Bangun Ruang Kerja".
    const HARD_BUDGET_MS = 5 * 60 * 1000;
    let lastYearDurationMs = 90 * 1000; // estimasi awal 1,5 menit per tahun
    let timeLimitReached = false;

    detectedYears.forEach((detectedYear) => {
      if (timeLimitReached) return;
      if (registeredYears.indexOf(detectedYear) !== -1) return;

      const elapsed = startTime ? (Date.now() - startTime) : 0;
      const remaining = HARD_BUDGET_MS - elapsed;
      if (remaining < lastYearDurationMs * 1.2) {
        wsPushReport_(report, 'warning', 'Sebagian tahun belum diimport otomatis karena mendekati batas waktu eksekusi.         Jalankan ulang "Bangun Ruang Kerja" untuk melanjutkan; struktur yang sudah dibuat tidak akan diulang.');
        timeLimitReached = true;
        return;
      }

      // Found a year folder in Drive that is NOT in config_years. Auto-import/initialize it!
      wsPushReport_(report, 'found', 'Mendeteksi folder tahun fisik ' + detectedYear + ' (sebagian/lengkap) di Drive.       Melakukan sinkronisasi otomatis...');
      const yearStart = Date.now();
      try {
        this.initializeSingleYear_(ss, detectedYear, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder,         report);
        registeredYears.push(detectedYear);
        lastYearDurationMs = Math.max(lastYearDurationMs, Date.now() - yearStart);
      } catch(err) {
        console.error('Failed to auto-import physical year ' + detectedYear + ': ' + err.message);
      }
    });
  }
};
