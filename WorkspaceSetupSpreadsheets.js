'use strict';

function wsGetOrCreateConfigSpreadsheet_(folder) {
  const existing = wsGetFileByNameAndMime_(folder, WORKSPACE_CONFIG.configSpreadsheetName, MimeType.GOOGLE_SHEETS);
  if (existing) return openSpreadsheetById_(existing.getId());
  const ss = SpreadsheetApp.create(WORKSPACE_CONFIG.configSpreadsheetName);
  wsMoveFileToFolder_(ss.getId(), folder);
  return ss;
}

function wsEnsureArchiveSpreadsheet_(laciFolder, activity, year, report) {
  const laciNo = activity.laciNo || activity.laci_no || '';
  const spreadsheetName = 'Daftar Isi Berkas Arsip Laci No.' + laciNo + ' Th. ' + year + ' - Production';
  const companion = wsGetFileByNameAndMime_(laciFolder, spreadsheetName, MimeType.GOOGLE_SHEETS);
  if (companion) {
    // wsInstallRekapTriggerIfMissing_(companion.getId()); // Dimatikan sementara agar Init Workspace lebih cepat (Opt-in)
    return openSpreadsheetById_(companion.getId());
  }

  const native = wsFindFirstArchiveSheet_(laciFolder);
  if (native) {
    // wsInstallRekapTriggerIfMissing_(native.getId()); // Dimatikan sementara agar Init Workspace lebih cepat (Opt-in)
    return openSpreadsheetById_(native.getId());
  }

  const office = wsFindFirstOfficeSpreadsheet_(laciFolder);
  if (office) {
    const converted = wsTryConvertOfficeSpreadsheet_(office, spreadsheetName, laciFolder);
    if (converted) {
      // wsInstallRekapTriggerIfMissing_(converted.getId()); // Dimatikan sementara agar Init Workspace lebih cepat (Opt-in)
      return converted;
    }
  }

  const ss = SpreadsheetApp.create(spreadsheetName);
  wsMoveFileToFolder_(ss.getId(), laciFolder);
  wsPrepareArchiveWorkbook_(ss, activity);
  // wsInstallRekapTriggerIfMissing_(ss.getId()); // Dimatikan sementara agar Init Workspace lebih cepat (Opt-in)
  wsPushReport_(report, 'created', 'Spreadsheet arsip dibuat: ' + spreadsheetName);
  return ss;
}

function wsInstallRekapTriggerIfMissing_(spreadsheetId) {
  // DINONAKTIFKAN: Google Apps Script memiliki batas maksimal 20 installable trigger per user per script.
  // Karena jumlah sub-kegiatan bisa lebih dari 20, memasang onEdit trigger pada setiap spreadsheet
  // akan menyebabkan error "This script has too many triggers" dan menyebabkan inisialisasi crash (INTERNAL ERROR).
  // Sinkronisasi Kode Klasifikasi 2 arah dari Spreadsheet ke App dihentikan. Gunakan antarmuka App untuk mengubah kode.
}

function cleanupOldRekapTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onRekapSheetEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function cleanupOcrTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processOcrQueue') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }
  console.log('Deleted ' + deleted + ' old processOcrQueue triggers.');
}

function installTriggersForExistingSpreadsheets() {
  const year = ConfigService.getSettings().currentYear || new Date().getFullYear();
  const activities = ConfigRepository.getActivities(year);
  activities.forEach(function(act) {
    if (act.spreadsheet_file_id) {
      wsInstallRekapTriggerIfMissing_(act.spreadsheet_file_id);
    }
    const subActivities = ConfigRepository.getSubActivities(year, act.activity_id);
    subActivities.forEach(function(sub) {
      if (sub.spreadsheet_file_id && sub.spreadsheet_file_id !== act.spreadsheet_file_id) {
         wsInstallRekapTriggerIfMissing_(sub.spreadsheet_file_id);
      }
    });
  });
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
    if (normalizedCandidate && (normalizedSheet.indexOf(normalizedCandidate) >= 0 || normalizedCandidate.indexOf    (normalizedSheet) >= 0)) {
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
    return openSpreadsheetById_(copied.id);
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

  // Sembunyikan sheet template bawaan agar tidak mengganggu
  ['Template Detail Item', 'Template Detail Kegiatan', 'Template Kegiatan'].forEach(function(name) {
    const tSheet = ss.getSheetByName(name);
    if (tSheet) tSheet.hideSheet();
  });
}

