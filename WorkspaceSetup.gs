'use strict';

const WORKSPACE_CONFIG = {
  systemFolderName: '00. Sistem Portal',
  configSpreadsheetName: 'Portal Arsip PUSJARSKPP - Config Production',
  inboxFolderName: 'Inbox Dokumen Masuk',
  templateFolderName: 'Template Surat'
};

const WORKSPACE_ACTIVITIES = [
  {
    id: 'kepemimpinan',
    label: 'Kepemimpinan',
    laciNo: '1',
    folderNo: '1',
    sortOrder: 1,
    laciFolderName: 'LACI NO 1 (KEPEMIMPINAN)',
    targetFolderName: 'Folder 1 (Pelatihan Kepemimpinan)',
    laciCandidates: ['LACI NO 1', 'KEPEMIMPINAN'],
    targetCandidates: ['Folder 1', 'Kepemimpinan'],
    hasRekapSheet: true,
    defaultSubActivities: [
      'PKN Tk.II Angkatan X Tahun {year}',
      'PKA Angkatan 1 Tahun {year}',
      'PKA Angkatan 2 Tahun {year}',
      'PKP Angkatan 1 Tahun {year}',
      'PKP Angkatan 2 Tahun {year}'
    ]
  },
  {
    id: 'latsar_cpns',
    label: 'Latsar CPNS',
    laciNo: '2',
    folderNo: '2',
    sortOrder: 2,
    laciFolderName: 'LACI NO 2 (LATSAR CPNS)',
    targetFolderName: 'Folder 2 (Latsar CPNS)',
    laciCandidates: ['LACI NO 2', 'LATSAR'],
    targetCandidates: ['Folder 2', 'Latsar'],
    hasRekapSheet: true,
    defaultCode: '',
    defaultSubActivities: [
      '01. Latsar CPNS Angkatan I',
      '02. Latsar CPNS Angkatan II',
      '03. Latsar CPNS Angkatan III',
      '04. Latsar CPNS Angkatan IV',
      '05. Latsar CPNS Angkatan V',
      '06. Latsar CPNS Angkatan VI',
      '07. Latsar CPNS Angkatan VII',
      '08. Latsar CPNS Angkatan VIII',
      '09. Latsar CPNS Angkatan IX',
      '10. Latsar CPNS Angkatan X',
      '11. Latsar CPNS Angkatan XI',
      '12. Latsar CPNS Angkatan XII',
      '13. Latsar CPNS Kutai Timur (Kerjasama)',
      '14. Latsar CPNS Bengkayang (Kerjasama)'
    ]
  },
  {
    id: 'teknis',
    label: 'Teknis',
    laciNo: '3',
    folderNo: '3',
    sortOrder: 3,
    laciFolderName: 'LACI NO 3 (TEKNIS)',
    targetFolderName: 'Folder 3 (Pelatihan Teknis dan Lainnya)',
    laciCandidates: ['LACI NO 3', 'TEKNIS'],
    targetCandidates: ['Folder 3', 'Teknis'],
    hasRekapSheet: true,
    allowNonLetter: true,
    defaultSubActivities: [
      '01. Coaching bagi ASN',
      '02. Puskesmas Prima',
      '03. Desa Madani',
      '04. Sekolah Idaman',
      '10. Pelatihan Teknis Lainnya'
    ]
  },
  {
    id: 'lain_lain',
    label: 'Lain-lain',
    laciNo: '4',
    folderNo: '4',
    sortOrder: 4,
    laciFolderName: 'LACI NO 4 (LAIN-LAIN)',
    targetFolderName: 'Folder 4 (Lain-lain)',
    laciCandidates: ['LACI NO 4', 'LAIN'],
    targetCandidates: ['Folder 4', 'Lain'],
    hasRekapSheet: true,
    allowNonLetter: true,
    defaultSubActivities: [
      '01. Jabatan Fungsional',
      '02. Data Jabatan Fungsional',
      '03. Permohonan Fasilitator',
      '04. Kerjasama Pelatihan',
      '05. Evaluasi Pasca Pelatihan',
      '06. SK Tim Kerja Pusjar SKPP',
      '07. Workshop PKN-II'
    ]
  }
];

const WORKSPACE_DETAIL_HEADERS = ARCHIVE_FIELD_LABELS;

const WORKSPACE_REKAP_HEADERS = [
  'Nomor Berkas',
  'Kode Klasifikasi',
  'Uraian/ Informasi Arsip',
  'Kurun Waktu',
  'Jumlah',
  'No Filing Cabinet',
  'No Laci',
  'No Folder',
  'Keamanan & Akses Arsip',
  'Ket',
  'Data Fix Peserta',
  'Kumpulan Materi',
  'Laporan',
  'Sertifikat'
];

const WorkspaceSetupService = {
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

    const daftarArsip = wsFindOrCreateChildFolder_(root, ['1. Daftar Arsip', 'Daftar Arsip'], '1. Daftar Arsip (Spreadsheet)', report);
    const naskahDinas = wsFindOrCreateChildFolder_(root, ['2. Naskah Dinas Latbang', 'Naskah Dinas LitBang', 'Naskah Dinas'], '2. Naskah Dinas Latbang (Dokumen)', report);
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
    this.initializeSingleYear_(configSpreadsheet, year, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report);

    // 2. Scan and auto-import any other years that physically exist in Google Drive
    this.scanAndImportPhysicalYears_(configSpreadsheet, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report, startTime);

    return {
      year: year, rootFolderId: root.getId(), rootFolderUrl: root.getUrl(),
      referenceRootFolderId: '', systemFolderParentId: customParentId || '',
      systemFolderName: systemFolderName, configFolderId: configFolderId || systemFolder.getId(),
      configSpreadsheetId: configSpreadsheet.getId(), configSpreadsheetUrl: configSpreadsheet.getUrl(),
      report: report
    };
  },

  initializeSingleYear_: function(ss, year, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report) {
    const arsipDiklat = wsFindOrCreateChildFolder_(daftarArsip, ['ARSIP DIKLAT ' + year, 'DIKLAT ' + year, String(year)], 'ARSIP DIKLAT ' + year, report);
    const tahunFolder = wsFindOrCreateChildFolder_(persuratan, ['Tahun ' + year, String(year)], 'Tahun ' + year, report);
    const dokumenFolder = wsFindOrCreateChildFolder_(naskahDinas, ['2. Dokumen', 'Dokumen'], '2. Dokumen', report);
    const dokumenTahunFolder = wsFindOrCreateChildFolder_(dokumenFolder, ['Tahun ' + year, String(year)], 'Tahun ' + year, report);

    const activityRows = [];
    const subActivityRows = [];
    const existingSubMap = wsExistingSubActivityMap_(ss, year);

    WORKSPACE_ACTIVITIES.forEach(function (activity) {
      const laciFolder = wsFindOrCreateChildFolder_(arsipDiklat, activity.laciCandidates, activity.laciFolderName, report);
      const targetFolder = wsFindOrCreateChildFolder_(tahunFolder, activity.targetCandidates, activity.targetFolderName, report);
      const docTargetFolder = wsFindOrCreateChildFolder_(dokumenTahunFolder, activity.targetCandidates, activity.targetFolderName, report);
      const spreadsheetContext = wsResolveActivitySpreadsheetContext_(laciFolder, activity, year, report);
      const spreadsheet = spreadsheetContext.defaultSpreadsheet;

      activityRows.push([year, activity.id, activity.label, activity.laciNo, activity.folderNo, spreadsheet.getId(), targetFolder.getId(), laciFolder.getId(), 'TRUE', activity.sortOrder]);

      const existingSubFolders = wsListChildFolders_(targetFolder);
      const defaultSubActivities = wsDefaultSubActivityNames_(activity, year);
      if (!existingSubFolders.length && defaultSubActivities.length > 0) {
        defaultSubActivities.forEach(function (subName) {
          wsFindOrCreateChildFolder_(targetFolder, [subName], subName, report);
          wsFindOrCreateChildFolder_(docTargetFolder, [subName], subName, report);
        });
      }

      wsBuildSubActivityEntries_(root, naskahDinas, persuratan, tahunFolder, targetFolder, activity).forEach(function (entry, index) {
        const subFolder = entry.folder;
        if (entry.parentFolderName) {
          const docParentGroup = wsFindOrCreateChildFolder_(docTargetFolder, [entry.parentFolderName], entry.parentFolderName, report);
          wsFindOrCreateChildFolder_(docParentGroup, [subFolder.getName()], subFolder.getName(), report);
        } else {
          wsFindOrCreateChildFolder_(docTargetFolder, [subFolder.getName()], subFolder.getName(), report);
        }
        const noFolder = WorkspaceSetupService.parseLeadingNumber(subFolder.getName()) || activity.folderNo;
        const subSpreadsheet = wsPickSpreadsheetForSubActivity_(spreadsheetContext, activity, entry.groupName || subFolder.getName());
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
          index + 1,
          targetSheetName,
          mappingStatus,
          existingSub.mapping_note || '',
          existingSub.rekap_row_number || '',
          '',
          '',
          entry.parentFolderId,
          entry.parentFolderName,
          entry.parentFolderPath,
          subSpreadsheet ? subSpreadsheet.getId() : ''
        ]);
      });
    });

    wsWriteConfig_(ss, year, root, null, daftarArsip, arsipDiklat, tahunFolder, inbox, templateFolder, activityRows, subActivityRows);

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

  scanAndImportPhysicalYears_: function(ss, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report, startTime) {
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
      let pageToken = null;
      do {
        let res;
        try {
          res = Drive.Files.list({
            q: "'" + parentFolder.getId() + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            fields: "nextPageToken, files(name)",
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
          const folderName = items[i].name;
          const match = folderName.match(/\b(202\d|203[0-5])\b/);
          if (match) {
            const detectedYear = Number(match[1]);
            if (detectedYears.indexOf(detectedYear) === -1) {
              detectedYears.push(detectedYear);
            }
          }
        }
        pageToken = res.nextPageToken;
      } while (pageToken);
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
        wsPushReport_(report, 'warning', 'Sebagian tahun belum diimport otomatis karena mendekati batas waktu eksekusi. Jalankan ulang "Bangun Ruang Kerja" untuk melanjutkan; struktur yang sudah dibuat tidak akan diulang.');
        timeLimitReached = true;
        return;
      }

      // Found a year folder in Drive that is NOT in config_years. Auto-import/initialize it!
      wsPushReport_(report, 'found', 'Mendeteksi folder tahun fisik ' + detectedYear + ' (sebagian/lengkap) di Drive. Melakukan sinkronisasi otomatis...');
      const yearStart = Date.now();
      try {
        this.initializeSingleYear_(ss, detectedYear, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report);
        registeredYears.push(detectedYear);
        lastYearDurationMs = Math.max(lastYearDurationMs, Date.now() - yearStart);
      } catch(err) {
        console.error('Failed to auto-import physical year ' + detectedYear + ': ' + err.message);
      }
    });
  }
};

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
          // Bulk write to Spreadsheet
          const writeResults = SpreadsheetService.appendArchiveRowsBulk(activity, subActivity, metadataList);
          SpreadsheetService.updateRekapSummary(activity, subActivity, {}); // Update rekap once per sub activity
          
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
          totalSynced += metadataList.length;
        }

        pageToken = result.nextPageToken;
      } while (pageToken);
    });
  });
  
  if (allLogsToAppend.length > 0) {
    ConfigRepository.appendArchiveLogsBulk(allLogsToAppend);
  }
  
  if (timeLimitReached) {
    wsPushReport_(report, 'warning', 'Waktu eksekusi hampir habis. Sebagian berkas mungkin belum tersinkronisasi. Silakan jalankan Sinkronisasi Ruang Kerja ulang nanti.');
  } else if (totalSynced > 0) {
    wsPushReport_(report, 'created', 'Berhasil mensinkronisasi ' + totalSynced + ' file fisik yang sudah ada di Drive secara bulk.');
  }
  
  return {
    totalSynced: totalSynced,
    timeLimitReached: timeLimitReached
  };
}

function wsGetOrCreateConfigSpreadsheet_(folder) {
  const existing = wsGetFileByNameAndMime_(folder, WORKSPACE_CONFIG.configSpreadsheetName, MimeType.GOOGLE_SHEETS);
  if (existing) return SpreadsheetApp.openById(existing.getId());
  const ss = SpreadsheetApp.create(WORKSPACE_CONFIG.configSpreadsheetName);
  wsMoveFileToFolder_(ss.getId(), folder);
  return ss;
}

function wsEnsureArchiveSpreadsheet_(laciFolder, activity, year, report) {
  const laciNo = activity.laciNo || activity.laci_no || '';
  const spreadsheetName = 'Daftar Isi Berkas Arsip Laci No.' + laciNo + ' Th. ' + year + ' - Production';
  const companion = wsGetFileByNameAndMime_(laciFolder, spreadsheetName, MimeType.GOOGLE_SHEETS);
  if (companion) return SpreadsheetApp.openById(companion.getId());

  const native = wsFindFirstArchiveSheet_(laciFolder);
  if (native) return SpreadsheetApp.openById(native.getId());

  const office = wsFindFirstOfficeSpreadsheet_(laciFolder);
  if (office) {
    const converted = wsTryConvertOfficeSpreadsheet_(office, spreadsheetName, laciFolder);
    if (converted) return converted;
  }

  const ss = SpreadsheetApp.create(spreadsheetName);
  wsMoveFileToFolder_(ss.getId(), laciFolder);
  wsPrepareArchiveWorkbook_(ss, activity);
  wsPushReport_(report, 'created', 'Spreadsheet arsip dibuat: ' + spreadsheetName);
  return ss;
}

function wsResolveActivitySpreadsheetContext_(laciFolder, activity, year, report) {
  const context = {
    defaultSpreadsheet: null,
    byGroup: {}
  };

  if (activity.spreadsheetBySubActivity) {
    const groups = activity.spreadsheetGroups || ['PKN', 'PKA', 'PKP'];
    groups.forEach(function (groupName) {
      const groupFolder = wsFindOrCreateChildFolder_(laciFolder, [groupName], groupName, report);
      const groupSpreadsheet = wsEnsureArchiveSpreadsheet_(groupFolder, activity, year, report);
      const groupKey = wsLeadershipGroupKey_(groupName);
      if (groupKey) context.byGroup[groupKey] = groupSpreadsheet;
      if (!context.defaultSpreadsheet) context.defaultSpreadsheet = groupSpreadsheet;
    });
  }

  if (!context.defaultSpreadsheet) {
    context.defaultSpreadsheet = wsEnsureArchiveSpreadsheet_(laciFolder, activity, year, report);
  }

  return context;
}

function wsPickSpreadsheetForSubActivity_(context, activity, subActivityName) {
  if (!context) return null;
  const groupKey = wsActivityId_(activity) === 'kepemimpinan'
    ? wsLeadershipGroupKey_(subActivityName)
    : '';
  if (groupKey && context.byGroup && context.byGroup[groupKey]) return context.byGroup[groupKey];
  return context.defaultSpreadsheet || null;
}

function wsBuildSubActivityEntries_(root, naskahDinas, persuratan, tahunFolder, targetFolder, activity) {
  const rootParts = [naskahDinas.getName(), persuratan.getName(), tahunFolder.getName(), targetFolder.getName()];
  return wsBuildLeafSubActivityEntries_(root.getName(), rootParts, targetFolder, activity);
}

function wsBuildLeafSubActivityEntries_(rootName, rootParts, targetFolder, activity) {
  const entries = [];
  const baseParts = rootParts || [];

  function findGroupName_(ancestors) {
    if (wsActivityId_(activity) !== 'kepemimpinan') return '';
    for (let i = 0; i < ancestors.length; i++) {
      if (wsLeadershipGroupKey_(ancestors[i].getName())) return ancestors[i].getName();
    }
    return '';
  }

  function visit_(folder, ancestors) {
    const childFolders = wsListChildFolders_(folder);
    if (childFolders.length > 0) {
      childFolders.forEach(function (childFolder) {
        visit_(childFolder, ancestors.concat(folder));
      });
      return;
    }

    addLeaf_(folder, ancestors);
  }

  function addLeaf_(folder, ancestors) {
    const ancestorNames = ancestors.map(function (ancestor) { return ancestor.getName(); });
    const relativeParts = ancestorNames.concat(folder.getName());
    const parentFolder = ancestors.length ? ancestors[ancestors.length - 1] : null;
    entries.push({
      folder: folder,
      groupName: findGroupName_(ancestors),
      subActivityId: slug_(relativeParts.join('_')),
      folderPath: wsBuildPath_(rootName, baseParts.concat(relativeParts)),
      parentFolderId: parentFolder ? parentFolder.getId() : '',
      parentFolderName: parentFolder ? parentFolder.getName() : '',
      parentFolderPath: parentFolder ? wsBuildPath_(rootName, baseParts.concat(ancestorNames)) : ''
    });
  }

  wsListChildFolders_(targetFolder).forEach(function (folder) {
    visit_(folder, []);
  });

  return entries;
}

function wsResolveSpreadsheetForSubActivity_(activity, subActivityName, year, report, parentGroupName) {
  if (!activity || !activity.laci_folder_id) return null;
  const laciFolder = DriveApp.getFolderById(activity.laci_folder_id);
  if (!activity.spreadsheetBySubActivity) {
    return wsEnsureArchiveSpreadsheet_(laciFolder, activity, year || DEFAULT_YEAR, report);
  }

  const groupKey = wsLeadershipGroupKey_(parentGroupName) || wsLeadershipGroupKey_(subActivityName);
  if (!groupKey) return wsEnsureArchiveSpreadsheet_(laciFolder, activity, year || DEFAULT_YEAR, report);

  const groupFolderName = groupKey.toUpperCase();
  const groupFolder = wsFindOrCreateChildFolder_(laciFolder, [groupFolderName], groupFolderName, report);
  return wsEnsureArchiveSpreadsheet_(groupFolder, activity, year || DEFAULT_YEAR, report);
}

function wsExistingSubActivityMap_(ss, year) {
  const map = {};
  if (!ss || !ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES)) return map;
  readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
    .filter(function (row) { return Number(row.year) === Number(year); })
    .forEach(function (row) {
      if (row.activity_id && row.sub_activity_id) {
        map[row.activity_id + '|' + row.sub_activity_id] = row;
      }
    });
  return map;
}

function wsSuggestFormalArchiveName_(activity, folderName) {
  const name = wsStripLeadingIndex_(folderName);
  const activityId = wsActivityId_(activity);
  if (activityId === 'latsar_cpns') {
    return name
      .replace(/^Latsar\s+CPNS\b/i, 'Pelatihan Dasar CPNS')
      .replace(/^Latsar\b/i, 'Pelatihan Dasar CPNS')
      .trim();
  }
  if (activityId === 'kepemimpinan') {
    if (/^PKN\b/i.test(name)) return name.replace(/^PKN\b/i, 'Pelatihan Kepemimpinan Nasional').trim();
    if (/^PKA\b/i.test(name)) return name.replace(/^PKA\b/i, 'Pelatihan Kepemimpinan Administrator').trim();
    if (/^PKP\b/i.test(name)) return name.replace(/^PKP\b/i, 'Pelatihan Kepemimpinan Pengawas').trim();
  }
  return name || folderName;
}

function wsSuggestDetailSheetName_(activity, folderName, spreadsheet) {
  const candidate = wsPreferredDetailSheetCandidate_(activity, folderName);
  const existing = wsFindMatchingSheetName_(spreadsheet, folderName, candidate);
  return existing || candidate || wsStripLeadingIndex_(folderName) || folderName;
}

function wsPreferredDetailSheetCandidate_(activity, folderName) {
  const name = wsStripLeadingIndex_(folderName);
  const activityId = wsActivityId_(activity);
  if (activityId === 'latsar_cpns') {
    const angkatan = name.match(/Angkatan\s+([IVXLCDM]+|\d+)/i);
    if (angkatan) return 'Latsar Ak ' + angkatan[1].toUpperCase();
    const kutim = name.match(/Kutim\s+(\d+)/i);
    if (kutim) return 'Latsar Kutim ' + kutim[1];
    if (/Bengkayang/i.test(name)) return 'Latsar Bengkayang';
  }
  if (activityId === 'kepemimpinan') {
    return name.replace(/\s*Tahun\s+20\d{2}\b/i, '').trim();
  }
  return name;
}

function wsFindMatchingSheetName_(spreadsheet, folderName, candidate) {
  if (!spreadsheet || !spreadsheet.getSheets) return '';
  const normalizedFolder = wsNormalizeMappingText_(folderName);
  const normalizedCandidate = wsNormalizeMappingText_(candidate);
  const rekapName = wsNormalizeMappingText_(REKAP_SHEET_NAME);
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    const normalizedSheet = wsNormalizeMappingText_(sheetName);
    if (!normalizedSheet || normalizedSheet === rekapName) continue;
    if (normalizedCandidate && normalizedSheet === normalizedCandidate) return sheetName;
    if (normalizedFolder && normalizedSheet === normalizedFolder) return sheetName;
    if (normalizedCandidate && (normalizedSheet.indexOf(normalizedCandidate) >= 0 || normalizedCandidate.indexOf(normalizedSheet) >= 0)) {
      return sheetName;
    }
  }
  return '';
}

function wsInferMappingStatus_(folderName, formalArchiveName, targetSheetName) {
  const folderNorm = wsNormalizeMappingText_(folderName);
  const formalNorm = wsNormalizeMappingText_(formalArchiveName);
  const sheetNorm = wsNormalizeMappingText_(targetSheetName);
  if (!formalNorm || !sheetNorm) return 'PERLU_REVIEW';
  if (formalNorm === folderNorm || sheetNorm === folderNorm) return 'PERLU_REVIEW';
  return 'AUTO_MATCHED';
}

function wsStripLeadingIndex_(value) {
  return String(value || '').trim().replace(/^\d+\s*[\.\-\)]\s*/, '').trim();
}

function wsNormalizeMappingText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\btahun\s+20\d{2}\b/g, '')
    .replace(/\bcpns\b/g, '')
    .replace(/\bpelatihan\s+dasar\b/g, 'latsar')
    .replace(/\bangkatan\b/g, 'ak')
    .replace(/[^a-z0-9ivxlcdm]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function wsActivityId_(activity) {
  return String(activity.id || activity.activity_id || '');
}

function wsLeadershipGroupKey_(value) {
  const normalized = wsNormalize_(value);
  if (/\bpkn\b/.test(normalized)) return 'pkn';
  if (/\bpka\b/.test(normalized)) return 'pka';
  if (/\bpkp\b/.test(normalized)) return 'pkp';
  return '';
}

function wsDefaultSubActivityNames_(activity, year) {
  return (activity.defaultSubActivities || []).map(function (name) {
    return String(name).replace(/\{year\}/g, String(year || DEFAULT_YEAR));
  });
}

function wsTryConvertOfficeSpreadsheet_(file, targetName, folder) {
  try {
    const resource = {
      name: targetName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [folder.getId()]
    };
    const copied = Drive.Files.copy(resource, file.getId(), { supportsAllDrives: true });
    return SpreadsheetApp.openById(copied.id);
  } catch (error) {
    console.error('Office spreadsheet conversion failed; blank companion will be created. ' + error.message);
    return null;
  }
}

function wsPrepareArchiveWorkbook_(ss, activity) {
  const first = ss.getSheets()[0];
  first.setName(activity.hasRekapSheet ? 'Daftar Berkas Arsip Aktip' : 'Daftar Isi Berkas Arsip Aktip');
  if (activity.hasRekapSheet) {
    wsFormatRekapSheet_(first);
  } else {
    wsFormatDetailSheet_(first);
  }
}

function wsWriteConfig_(ss, year, root, referenceRoot, daftarArsip, arsipDiklat, tahunFolder, inbox, templateFolder, activityRows, subActivityRows) {
  // Ensure all config sheets exist
  [
    CONFIG_SHEETS.YEARS,
    CONFIG_SHEETS.ACTIVITIES,
    CONFIG_SHEETS.SUB_ACTIVITIES,
    CONFIG_SHEETS.METADATA_FIELDS
  ].forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // Ensure log sheets exist with headers
  wsEnsureSheetWithHeaders_(ss, CONFIG_SHEETS.ARCHIVE_LOG, [
    'archive_id', 'year', 'activity_id', 'sub_activity_id', 'source_file_id', 'final_file_id', 'final_file_name', 'target_folder_id', 'target_folder_name', 'target_folder_path', 'spreadsheet_file_id', 'spreadsheet_row_number', 'status', 'created_at', 'created_by', 'error_message', 'metadata_json'
  ]);
  wsEnsureSheetWithHeaders_(ss, CONFIG_SHEETS.ADMIN_AUDIT_LOG, [
    'created_at', 'actor', 'action', 'year', 'activity_id', 'sub_activity_id', 'folder_id', 'status', 'message'
  ]);

  // Upsert config_years (year is at index 0)
  const yearsHeaders = ['year', 'status', 'root_folder_id', 'reference_root_folder_id', 'daftar_arsip_folder_id', 'spreadsheet_year_folder_id', 'persuratan_year_folder_id', 'inbox_folder_id', 'template_folder_id', 'root_url'];
  const yearsNewRows = [[year, 'PROD_ACTIVE', root.getId(), referenceRoot ? referenceRoot.getId() : '', daftarArsip.getId(), arsipDiklat.getId(), tahunFolder.getId(), inbox.getId(), templateFolder.getId(), root.getUrl()]];
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.YEARS), yearsHeaders, year, 0, yearsNewRows);

  // Upsert config_activities (year is at index 0)
  const activitiesHeaders = ['year', 'activity_id', 'activity_name', 'laci_no', 'folder_no', 'spreadsheet_file_id', 'target_folder_id', 'laci_folder_id', 'is_active', 'sort_order'];
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.ACTIVITIES), activitiesHeaders, year, 0, activityRows);

  // Upsert config_sub_activities (year is at index 0)
  const subActivitiesHeaders = SUB_ACTIVITY_HEADERS;
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES), subActivitiesHeaders, year, 0, subActivityRows);

  // Upsert config_metadata_fields (year is at index 1)
  const metadataHeaders = ['field_id', 'year', 'activity_id', 'field_name', 'spreadsheet_column_name', 'is_required', 'default_value', 'input_type', 'sort_order', 'is_visible_in_form'];
  const metadataNewRows = wsBuildMetadataRows_(year).slice(1);
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.METADATA_FIELDS), metadataHeaders, year, 1, metadataNewRows);
}

function wsBuildMetadataRows_(year) {
  const rows = [['field_id', 'year', 'activity_id', 'field_name', 'spreadsheet_column_name', 'is_required', 'default_value', 'input_type', 'sort_order', 'is_visible_in_form']];
  const fields = [
    ['no_berkas', 'No Berkas', true, '', 'number'],
    ['nomor_item_arsip', 'Nomor Item Arsip', true, '', 'text'],
    ['nomor_surat', 'Nomor Surat', false, '', 'text'],
    ['kode_klasifikasi', 'Kode Klasifikasi', false, '', 'text'],
    ['uraian_informasi_item', 'Uraian Informasi Item', true, '', 'textarea'],
    ['tanggal', 'Tanggal', false, '', 'date'],
    ['tingkat_perkembangan', 'Tingkat Perkembangan', true, 'Asli', 'select'],
    ['jumlah', 'Jumlah', true, '1', 'number'],
    ['satuan', 'Satuan', false, 'Lembar', 'text'],
    ['no_filing_cabinet', 'No Filing Cabinet', false, '', 'number'],
    ['no_laci', 'No Laci', true, '', 'number'],
    ['no_folder', 'No Folder', true, '', 'number'],
    ['klasifikasi_akses', 'Klasifikasi Keamanan & Akses Arsip', false, 'Terbatas', 'select'],
    ['ket', 'Ket.', false, '', 'text'],
    ['lokasi_simpan', 'Lokasi Simpan', true, '', 'text']
  ];

  WORKSPACE_ACTIVITIES.forEach((activity) => {
    fields.forEach((field, index) => {
      rows.push([
        activity.id + '_' + field[0],
        year,
        activity.id,
        field[0],
        field[1],
        field[2] ? 'TRUE' : 'FALSE',
        field[0] === 'no_laci' ? activity.laciNo : field[3],
        field[4],
        index + 1,
        'TRUE'
      ]);
    });
  });
  return rows;
}

function wsFormatRekapSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(7);
  sheet.getRange('B1:O1').merge().setValue('Daftar Berkas Arsip Aktip');
  sheet.getRange('B1:O1')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('B4').setValue('Unit Pengolah :');
  sheet.getRange('D4').setValue('Kompartemen Latbang');
  sheet.getRange('B4:D4')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');

  sheet.getRange(5, 2, 2, WORKSPACE_REKAP_HEADERS.length).setValues([
    ['Nomor\nBerkas', 'Kode\nKlasifikasi', 'Uraian/\nInformasi\nArsip', 'Kurun\nWaktu', 'Jumlah', 'Lokasi', '', '', 'Keamanan\n& Akses\nArsip', 'Ket', 'Data Fix\nPeserta', 'Kumpulan\nMateri', 'Laporan', 'Sertifikat'],
    ['', '', '', '', '', 'No Filing\nCabinet', 'No Laci', 'No Folder', '', '', '', '', '', '']
  ]);
  sheet.getRange('B5:B6').merge();
  sheet.getRange('C5:C6').merge();
  sheet.getRange('D5:D6').merge();
  sheet.getRange('E5:E6').merge();
  sheet.getRange('F5:F6').merge();
  sheet.getRange('G5:I5').merge();
  sheet.getRange('J5:J6').merge();
  sheet.getRange('K5:K6').merge();
  sheet.getRange('L5:L6').merge();
  sheet.getRange('M5:M6').merge();
  sheet.getRange('N5:N6').merge();
  sheet.getRange('O5:O6').merge();
  sheet.getRange(7, 2, 1, WORKSPACE_REKAP_HEADERS.length)
    .setValues([WORKSPACE_REKAP_HEADERS.map(function (_, index) { return index + 1; })]);

  wsStyleRekapSheet_(sheet);
}

function wsStyleRekapSheet_(sheet) {
  const startRow = 5;
  const startCol = 2;
  const dataRows = 14;
  const width = WORKSPACE_REKAP_HEADERS.length;
  const lastTableRow = 7 + dataRows;

  sheet.getRange(startRow, startCol, 3, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setWrap(true);
  sheet.getRange(startRow, startCol, 2, width)
    .setFontWeight('normal')
    .setBackground('#bfbfbf');
  sheet.getRange(7, startCol, 1, width)
    .setFontSize(7)
    .setFontWeight('normal')
    .setBackground('#ffffff');
  sheet.getRange(8, startCol, dataRows, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 22);
  sheet.setRowHeight(4, 24);
  sheet.setRowHeights(5, 2, 30);
  sheet.setRowHeight(7, 18);
  sheet.setRowHeights(8, dataRows, 20);
  [50, 80, 280, 120, 80, 80, 70, 70, 90, 80, 90, 90, 90, 90]
    .forEach(function (widthPx, index) {
      sheet.setColumnWidth(startCol + index, widthPx);
    });

  wsWriteRekapNotes_(sheet, lastTableRow + 4);
}

function wsWriteRekapNotes_(sheet, startRow) {
  const notes = [
    ['Keterangan Petunjuk Pengisian:'],
    ['Kolom (1), diisi dengan nomor urut berkas;'],
    ['Kolom (2), diisi dengan kode klasifikasi Arsip;'],
    ['Kolom (3), diisi dengan uraian informasi dari berkas Arsip berdasarkan kegiatan dalam Klasifikasi Arsip;'],
    ['Kolom (4), diisi dengan masa/kurun waktu Arsip yang tercipta;'],
    ['Kolom (5), diisi dengan jumlah banyaknya Arsip dalam satuan yang sesuai dengan jenis Arsip;'],
    ['Kolom (6), diisi dengan nomor Filing Cabinet;'],
    ['Kolom (7), diisi dengan nomor laci pada Filing Cabinet;'],
    ['Kolom (8), diisi dengan nomor folder Arsip;'],
    ['Kolom (9), diisi dengan klasifikasi keamanan dan akses seperti terbuka, terbatas, dan rahasia;'],
    ['Kolom (10), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik, dan digital.']
  ];
  sheet.getRange(startRow, 3, notes.length, 1)
    .setValues(notes)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');
  sheet.getRange(startRow, 3).setFontWeight('normal');
}

function wsFormatDetailSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(8);
  sheet.getRange('B2:N2').merge().setValue('Daftar Isi Berkas Arsip Aktip');
  sheet.getRange('B2:N2')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('B4').setValue('Unit pengolah : Latbang');
  sheet.getRange('B4')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');

  wsWriteDetailHeader_(sheet);
  wsStyleDetailSheet_(sheet);
}

function wsWriteDetailHeader_(sheet) {
  sheet.getRange('B6:B7').merge().setValue('No\nBerkas');
  sheet.getRange('C6:C7').merge().setValue('Nomor Item\nArsip');
  sheet.getRange('D6:D7').merge().setValue('Kode\nKlasifikasi');
  sheet.getRange('E6:E7').merge().setValue('Uraian Informasi Item');
  sheet.getRange('F6:F7').merge().setValue('Tgl');
  sheet.getRange('G6:G7').merge().setValue('Tingkat\nPengembangan');
  sheet.getRange('H6:I6').merge().setValue('Jumlah');
  sheet.getRange('H7').setValue('Jumlah');
  sheet.getRange('I7').setValue('Satuan');
  sheet.getRange('J6:L6').merge().setValue('Lokasi');
  sheet.getRange('J7').setValue('No Filing\nCabinet');
  sheet.getRange('K7').setValue('No Laci');
  sheet.getRange('L7').setValue('No Folder');
  sheet.getRange('M6:M7').merge().setValue('Klasifikasi\nKeamanan &\nAkses Arsip');
  sheet.getRange('N6:N7').merge().setValue('Ket.\nLokasi\nSimpan');
  sheet.getRange('B8:N8').setValues([[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]]);
}

function wsStyleDetailSheet_(sheet) {
  const startCol = 2;
  const width = 13;
  const dataRows = 24;

  sheet.getRange(6, startCol, 3, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal')
    .setBackground('#b4c6e7')
    .setWrap(true);
  sheet.getRange(8, startCol, 1, width)
    .setFontSize(7);
  sheet.getRange(9, startCol, dataRows, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);
  const numberRows = [];
  for (let i = 1; i <= dataRows; i++) numberRows.push([i]);
  sheet.getRange(9, startCol, dataRows, 1)
    .setValues(numberRows)
    .setHorizontalAlignment('center');

  sheet.setRowHeight(2, 22);
  sheet.setRowHeight(4, 22);
  sheet.setRowHeights(6, 2, 31);
  sheet.setRowHeight(8, 18);
  sheet.setRowHeights(9, dataRows, 18);
  [55, 86, 78, 270, 108, 90, 42, 42, 84, 84, 84, 100, 72].forEach(function (widthPx, index) {
    sheet.setColumnWidth(startCol + index, widthPx);
  });

  wsWriteDetailNotes_(sheet, 34);
}

function wsWriteDetailNotes_(sheet, startRow) {
  const notes = [
    ['Keterangan Petunjuk Pengisian:'],
    ['Kolom (2), diisi dengan nomor item Arsip;'],
    ['Kolom (3), diisi dengan kode Klasifikasi Arsip;'],
    ['Kolom (4), diisi dengan uraian informasi Arsip dari setiap naskah dinas;'],
    ['Kolom (5), diisi dengan tanggal Arsip itu tercipta;'],
    ['Kolom (6), diisi dengan tingkat perkembangan Arsip (asli, copy, dan cetak);'],
    ['Kolom (7), diisi dengan jumlah Arsip dalam satuan naskah dinas;'],
    ['Kolom (8), diisi dengan nomor Filing Cabinet;'],
    ['Kolom (9), diisi dengan nomor laci pada Filing Cabinet;'],
    ['Kolom (10), diisi dengan nomor folder Arsip;'],
    ['Kolom (11), diisi dengan klasifikasi keamanan seperti terbuka, terbatas, dan rahasia.'],
    ['Kolom (12), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik, dan digital.']
  ];
  sheet.getRange(startRow, 3, notes.length, 1)
    .setValues(notes)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');
  sheet.getRange(startRow, 3).setFontWeight('normal');
}


function wsGetOrCreateFolder_(parent, name, report) {
  const folders = parent.getFoldersByName(name);
  while (folders.hasNext()) {
    const folder = folders.next();
    if (!folder.isTrashed()) {
      wsPushReport_(report, 'found', 'Folder ditemukan: ' + name);
      return folder;
    }
  }
  const folder = withRetry_(() => parent.createFolder(name));
  wsPushReport_(report, 'created', 'Folder dibuat: ' + name);
  return folder;
}

function wsFindOrCreateChildFolder_(parent, candidates, createName, report) {
  const found = wsFindChildFolder_(parent, candidates);
  if (found) {
    wsPushReport_(report, 'found', 'Folder ditemukan: ' + found.getName());
    return found;
  }
  const folder = withRetry_(() => parent.createFolder(createName));
  wsPushReport_(report, 'created', 'Folder dibuat: ' + createName);
  return folder;
}

function wsFindChildFolder_(parent, candidates) {
  let pageToken = null;
  do {
    let result;
    try {
      result = Drive.Files.list({
        q: "'" + parent.getId() + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        fields: "nextPageToken, files(id, name)",
        pageSize: 1000,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
    } catch (e) {
      break;
    }
    const items = result.files || [];
    
    // exact match first
    for (let i = 0; i < items.length; i++) {
      for (let j = 0; j < candidates.length; j++) {
        if (items[i].name === candidates[j]) return DriveApp.getFolderById(items[i].id);
      }
    }
    // fuzzy match second
    for (let i = 0; i < items.length; i++) {
      const normalized = wsNormalize_(items[i].name);
      for (let j = 0; j < candidates.length; j++) {
        if (normalized.indexOf(wsNormalize_(candidates[j])) >= 0) return DriveApp.getFolderById(items[i].id);
      }
    }
    pageToken = result.nextPageToken;
  } while (pageToken);
  
  return null;
}

function wsListChildFolders_(parent) {
  const folders = [];
  let pageToken = null;
  do {
    let result;
    try {
      result = Drive.Files.list({
        q: "'" + parent.getId() + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        fields: "nextPageToken, files(id, name)",
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
      folders.push({
        folder: DriveApp.getFolderById(items[i].id),
        name: items[i].name
      });
    }
    pageToken = result.nextPageToken;
  } while (pageToken);
  
  return folders.sort((a, b) => a.name.localeCompare(b.name, 'id', { numeric: true }))
                .map(item => item.folder);
}

function wsGetFileByNameAndMime_(folder, name, mimeType) {
  let q = "'" + folder.getId() + "' in parents and name = '" + name.replace(/'/g, "\\'") + "' and trashed = false";
  if (mimeType) {
    q += " and mimeType = '" + mimeType + "'";
  }
  
  let pageToken = null;
  do {
    let result;
    try {
      result = Drive.Files.list({
        q: q,
        fields: "nextPageToken, files(id)",
        pageSize: 10,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
    } catch (e) {
      break;
    }
    if (result.files && result.files.length > 0) return DriveApp.getFileById(result.files[0].id);
    pageToken = result.nextPageToken;
  } while (pageToken);
  return null;
}

function wsFindFirstArchiveSheet_(folder) {
  let pageToken = null;
  do {
    let result;
    try {
      result = Drive.Files.list({
        q: "'" + folder.getId() + "' in parents and mimeType = '" + MimeType.GOOGLE_SHEETS + "' and trashed = false",
        fields: "nextPageToken, files(id, name)",
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
      const name = wsNormalize_(items[i].name);
      if (name.indexOf('daftar') >= 0 && name.indexOf('arsip') >= 0) {
        return DriveApp.getFileById(items[i].id);
      }
    }
    pageToken = result.nextPageToken;
  } while (pageToken);
  return null;
}

function wsFindFirstOfficeSpreadsheet_(folder) {
  let pageToken = null;
  do {
    let result;
    try {
      result = Drive.Files.list({
        q: "'" + folder.getId() + "' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel' or name contains '.xls' or name contains '.xlsx') and trashed = false",
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
      const name = items[i].name;
      const normalizedName = wsNormalize_(name);
      const mimeType = items[i].mimeType;
      if (normalizedName.indexOf('daftar') >= 0 && normalizedName.indexOf('arsip') >= 0 &&
        (/\.xlsx?$/i.test(name) || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mimeType === 'application/vnd.ms-excel')) {
        return DriveApp.getFileById(items[i].id);
      }
    }
    pageToken = result.nextPageToken;
  } while (pageToken);
  return null;
}

function wsMoveFileToFolder_(fileId, folder) {
  const file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}

function wsNormalize_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function wsBuildPath_(rootName, parts) {
  return [rootName].concat(parts).join(' > ');
}

function wsPushReport_(report, status, label) {
  if (!report) return;
  report.push({
    status: status,
    label: label
  });
}

function wsEnsureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('normal').setBackground('#d9ead3');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function wsUpsertSheetData_(sheet, headers, year, yearColIndex, newRows) {
  let existingRows = [];
  if (sheet.getLastRow() > 1) {
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    existingRows = range.getValues();
  }
  
  // Filter out rows that belong to the current year
  const filteredRows = existingRows.filter(function (row) {
    return Number(row[yearColIndex]) !== Number(year);
  });
  
  // Combine
  const allRows = [headers].concat(filteredRows).concat(newRows).map(function (row) {
    return wsPadRow_(row, headers.length);
  });
  
  // Clear sheet and write
  sheet.clear();
  sheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, allRows[0].length).setFontWeight('normal').setBackground('#d9ead3');
  sheet.autoResizeColumns(1, allRows[0].length);
}

function wsPadRow_(row, width) {
  const padded = (row || []).slice(0, width);
  while (padded.length < width) padded.push('');
  return padded;
}

function wsCleanupOrphanedArchiveLogs_(ss, year, subActivityRows) {
  const logSheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
  if (!logSheet) return;
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return;

  const validFolderIds = subActivityRows.map(function(row) {
    return String(row[5] || '').trim(); // index 5 = folder_id in SUB_ACTIVITY_HEADERS
  });

  const logs = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues();
  const rowsToDelete = [];
  
  for (let i = 0; i < logs.length; i++) {
    const logYear = Number(logs[i][1]);
    if (logYear === year) {
      const targetFolderId = String(logs[i][7] || '').trim(); // index 7 = target_folder_id in ARCHIVE_LOG
      if (targetFolderId && validFolderIds.indexOf(targetFolderId) === -1) {
        rowsToDelete.push(i + 2);
      }
    }
  }

  if (rowsToDelete.length > 0) {
    // Delete from bottom to top so indices don't shift
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      logSheet.deleteRow(rowsToDelete[i]);
    }
  }
}

