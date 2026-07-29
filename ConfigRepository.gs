'use strict';

const ConfigRepository = {
  getConfigSpreadsheet: function() {
    const settings = ConfigService.getSettings();
    if (!settings.configSpreadsheetId) {
      throw new Error('CONFIG_SPREADSHEET_ID belum diisi. Pilih Folder Ruang Kerja di Pengaturan lalu bangun ruang kerja.');
    }
    // Memo handle per-eksekusi (lihat openSpreadsheetById_ di SpreadsheetService.gs):
    // loadAll + tiap mutasi membuka config ss berulang dalam satu request.
    return openSpreadsheetById_(settings.configSpreadsheetId);
  },

  // Cari 1 baris di archive_log tanpa membaca SELURUH sheet: baca header + HANYA kolom
  // kunci. Mengembalikan { headers, lastRow, lastCol, keyCol, keyValues } atau null bila
  // sheet kosong / kolom kunci tak ada. keyValues[k] = baris sheet (k+2).
  _scanArchiveLogColumn_: function(sheet, keyHeaderName) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return null;
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normalizeHeader_);
    const keyCol = headers.indexOf(keyHeaderName);
    if (keyCol === -1) return null;
    const keyValues = sheet.getRange(2, keyCol + 1, lastRow - 1, 1).getDisplayValues();
    return { headers: headers, lastRow: lastRow, lastCol: lastCol, keyCol: keyCol, keyValues: keyValues };
  },

  loadAll: function(year) {
    const selectedYear = Number(year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const ss = this.getConfigSpreadsheet();
    const years = readSheetObjects_(ss, CONFIG_SHEETS.YEARS);
    const activities = readSheetObjects_(ss, CONFIG_SHEETS.ACTIVITIES)
      .filter(row => Number(row.year) === selectedYear && isTrue_(row.is_active))
      .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
    const subActivities = readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
      .filter(row => Number(row.year) === selectedYear && isTrue_(row.is_active))
      .sort(compareSubActivitiesByArchiveNumber_);
    const fields = readSheetObjects_(ss, CONFIG_SHEETS.METADATA_FIELDS)
      .filter(row => Number(row.year) === selectedYear)
      .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
    const activityById = {};
    activities.forEach(a => { activityById[a.activity_id] = a; });
    const subActivityById = {};
    subActivities.forEach(s => { subActivityById[s.sub_activity_id] = s; });
    const archiveLogAll = readSheetObjects_(ss, CONFIG_SHEETS.ARCHIVE_LOG)
      .filter(row => !row.year || Number(row.year) === selectedYear);

    const history = archiveLogAll
      .slice(-80)
      .reverse()
      .map(row => {
        const act = activityById[row.activity_id];
        const sub = subActivityById[row.sub_activity_id];
        return Object.assign({}, row, {
          activity_name: act ? act.activity_name : '',
          sub_activity_name: sub ? sub.sub_activity_name : ''
        });
      });

    return {
      selectedYear,
      years,
      activities,
      subActivities,
      fields,
      history,
      historyAll: archiveLogAll
    };
  },

  getSubActivitiesForSync: function (year, activityId) {
    const ss = this.getConfigSpreadsheet();
    return readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
      .filter(row => Number(row.year) === Number(year) && row.activity_id === activityId)
      .sort(compareSubActivitiesByArchiveNumber_);
  },

  // Kegiatan aktif untuk tahun tertentu (dipakai rekonsiliasi global dan maintenance).
  getActivities: function (year) {
    const ss = this.getConfigSpreadsheet();
    return readSheetObjects_(ss, CONFIG_SHEETS.ACTIVITIES)
      .filter(row => Number(row.year) === Number(year) && isTrue_(row.is_active))
      .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
  },

  // Sub-kegiatan untuk activity tertentu (termasuk yang nonaktif — caller boleh memfilter).
  getSubActivities: function (year, activityId) {
    return this.getSubActivitiesForSync(year, activityId);
  },

  getSubActivityById: function (year, activityId, subActivityId) {
    return this.getSubActivitiesForSync(year, activityId).find(function (sub) {
      return sub.sub_activity_id === subActivityId;
    }) || null;
  },

  getInactiveSubActivities: function(year) {
    const selectedYear = Number(year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const ss = this.getConfigSpreadsheet();
    return readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
      .filter(row => Number(row.year) === selectedYear && !isTrue_(row.is_active))
      .sort((a, b) => String(a.activity_id).localeCompare(String(b.activity_id)) || Number(a.sort_order || 0) - Number(b.sort_order || 0));
  },

  getAdminAuditLogs: function(limit) {
    const ss = this.getConfigSpreadsheet();
    return readRecentSheetObjects_(ss, CONFIG_SHEETS.ADMIN_AUDIT_LOG, Number(limit || 100)).reverse();
  },

  getArchiveLogKeyMap: function(year) {
    const selectedYear = Number(year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return {};
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return {};
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normalizeHeader_);
    const ssCol = headers.indexOf('spreadsheet_file_id');
    const rowCol = headers.indexOf('spreadsheet_row_number');
    const fileCol = headers.indexOf('final_file_id');
    const yearCol = headers.indexOf('year');
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    const map = {};
    data.forEach(function (r) {
      if (yearCol !== -1 && r[yearCol] && Number(r[yearCol]) !== selectedYear) return;
      const spreadsheetId = ssCol !== -1 ? cleanId_(r[ssCol] || '') : '';
      const rowNumber = rowCol !== -1 ? String(r[rowCol] || '').trim() : '';
      if (spreadsheetId && rowNumber) map[spreadsheetId + ':' + rowNumber] = true;
      const finalFileId = fileCol !== -1 ? cleanId_(r[fileCol] || '') : '';
      if (finalFileId) map['file:' + finalFileId] = true;
    });
    return map;
  },

  // Map cleanId(final_file_id) -> { archive_id, year } untuk seluruh log (skip DRAFT),
  // dibangun dalam SATU baca sheet. Hilangkan N+1 getArchiveLogByFileId saat cascade trash.
  // Last-wins (top-down) menyamai getArchiveLogByFileId yang bottom-up-first.
  getArchiveLogFileIdMap: function() {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return {};
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return {};
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normalizeHeader_);
    const fileCol = headers.indexOf('final_file_id');
    const archiveCol = headers.indexOf('archive_id');
    const yearCol = headers.indexOf('year');
    const statusCol = headers.indexOf('status');
    if (fileCol === -1 || archiveCol === -1) return {};
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    const map = {};
    data.forEach(function (r) {
      if (statusCol !== -1 && r[statusCol] === STATUS.DRAFT) return;
      const fid = cleanId_(r[fileCol] || '');
      if (!fid) return;
      map[fid] = { archive_id: r[archiveCol], year: yearCol !== -1 ? r[yearCol] : '' };
    });
    return map;
  },

  appendArchiveLog: function(row) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.ARCHIVE_LOG);
    ensureHeaders_(sheet, [
      'archive_id',
      'year',
      'activity_id',
      'sub_activity_id',
      'source_file_id',
      'final_file_id',
      'final_file_name',
      'target_folder_id',
      'target_folder_name',
      'target_folder_path',
      'spreadsheet_file_id',
      'spreadsheet_row_number',
      'status',
      'created_at',
      'created_by',
      'error_message',
      'metadata_json'
    ]);
    const values = {
      archive_id: row.archive_id,
      year: row.year,
      activity_id: row.activity_id,
      sub_activity_id: row.sub_activity_id,
      source_file_id: row.source_file_id,
      final_file_id: row.final_file_id,
      final_file_name: row.final_file_name,
      target_folder_id: row.target_folder_id || '',
      target_folder_name: row.target_folder_name || '',
      target_folder_path: row.target_folder_path || '',
      spreadsheet_file_id: row.spreadsheet_file_id,
      spreadsheet_row_number: row.spreadsheet_row_number,
      status: row.status,
      created_at: row.created_at,
      created_by: row.created_by,
      error_message: row.error_message || '',
      metadata_json: row.metadata_json || ''
    };
    const headers = getHeaders_(sheet);
    sheet.appendRow(headers.map(function (header) {
      return values[header] === undefined ? '' : values[header];
    }));
    });
  },

  appendArchiveLogsBulk: function(logs) {
    if (!logs || logs.length === 0) return;
    return withLock_(() => {
      const ss = this.getConfigSpreadsheet();
      const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.ARCHIVE_LOG);
      ensureHeaders_(sheet, [
        'archive_id', 'year', 'activity_id', 'sub_activity_id', 'source_file_id', 
        'final_file_id', 'final_file_name', 'target_folder_id', 'target_folder_name', 
        'target_folder_path', 'spreadsheet_file_id', 'spreadsheet_row_number', 
        'status', 'created_at', 'created_by', 'error_message', 'metadata_json'
      ]);
      const headers = getHeaders_(sheet);
      
      const newRows = logs.map(function(row) {
        const values = {
          archive_id: row.archive_id,
          year: row.year,
          activity_id: row.activity_id,
          sub_activity_id: row.sub_activity_id,
          source_file_id: row.source_file_id || '',
          final_file_id: row.final_file_id,
          final_file_name: row.final_file_name,
          target_folder_id: row.target_folder_id || '',
          target_folder_name: row.target_folder_name || '',
          target_folder_path: row.target_folder_path || '',
          spreadsheet_file_id: row.spreadsheet_file_id,
          spreadsheet_row_number: row.spreadsheet_row_number,
          status: row.status,
          created_at: row.created_at,
          created_by: row.created_by,
          error_message: row.error_message || '',
          metadata_json: row.metadata_json || ''
        };
        return headers.map(function (header) {
          return values[header] === undefined ? '' : values[header];
        });
      });
      
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    });
  },

  deleteDraftLog: function(archiveId) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return null;
    const ctx = this._scanArchiveLogColumn_(sheet, 'archive_id');
    if (!ctx) return null;
    const statusCol = ctx.headers.indexOf('status');
    const sourceFileCol = ctx.headers.indexOf('source_file_id');
    if (statusCol === -1) return null;

    for (let k = ctx.keyValues.length - 1; k >= 0; k--) {
      if (ctx.keyValues[k][0] === archiveId) {
        const rowNum = k + 2;
        const rowVals = sheet.getRange(rowNum, 1, 1, ctx.lastCol).getDisplayValues()[0];
        if (rowVals[statusCol] === STATUS.DRAFT) {
          const sourceFileId = sourceFileCol !== -1 ? rowVals[sourceFileCol] : '';
          sheet.deleteRow(rowNum);
          return sourceFileId || null;
        }
      }
    }
    return null;
    });
  },

  deleteArchiveLog: function(archiveId) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return false;
    const ctx = this._scanArchiveLogColumn_(sheet, 'archive_id');
    if (!ctx) return false;

    for (let k = ctx.keyValues.length - 1; k >= 0; k--) {
      if (ctx.keyValues[k][0] === archiveId) {
        sheet.deleteRow(k + 2);
        return true;
      }
    }
    return false;
    });
  },

  getArchiveLog: function(archiveId) {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return null;
    const ctx = this._scanArchiveLogColumn_(sheet, 'archive_id');
    if (!ctx) return null;

    for (let k = ctx.keyValues.length - 1; k >= 0; k--) {
      if (ctx.keyValues[k][0] === archiveId) {
        const rowVals = sheet.getRange(k + 2, 1, 1, ctx.lastCol).getDisplayValues()[0];
        return objectFromHeaders_(ctx.headers, rowVals);
      }
    }
    return null;
  },

  getArchiveLogByFileId: function(fileId) {
    if (!fileId) return null;
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return null;
    const ctx = this._scanArchiveLogColumn_(sheet, 'final_file_id');
    if (!ctx) return null;
    const statusCol = ctx.headers.indexOf('status');
    const target = cleanId_(fileId);

    for (let k = ctx.keyValues.length - 1; k >= 0; k--) {
      if (cleanId_(ctx.keyValues[k][0]) === target) {
        const rowVals = sheet.getRange(k + 2, 1, 1, ctx.lastCol).getDisplayValues()[0];
        if (statusCol !== -1 && rowVals[statusCol] === STATUS.DRAFT) continue;
        return objectFromHeaders_(ctx.headers, rowVals);
      }
    }
    return null;
  },

  updateArchiveLog: function(archiveId, updates) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
    if (!sheet) return false;
    const ctx = this._scanArchiveLogColumn_(sheet, 'archive_id');
    if (!ctx) return false;

    for (let k = ctx.keyValues.length - 1; k >= 0; k--) {
      if (ctx.keyValues[k][0] === archiveId) {
        updateConfigRow_(sheet, k + 2, updates);
        return true;
      }
    }
    return false;
    });
  },

  decrementArchiveLogRows: function(spreadsheetFileId, thresholdRow) {
    if (!spreadsheetFileId || !thresholdRow) return 0;
    return withLock_(() => {
      const ss = this.getConfigSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG_SHEETS.ARCHIVE_LOG);
      if (!sheet) return 0;
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2) return 0;
      const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normalizeHeader_);
      const ssCol = headers.indexOf('spreadsheet_file_id');
      const rowCol = headers.indexOf('spreadsheet_row_number');
      if (ssCol === -1 || rowCol === -1) return 0;

      // Baca HANYA dua kolom yang relevan (file_id & row_number), bukan seluruh sheet.
      const n = lastRow - 1;
      const ssVals = sheet.getRange(2, ssCol + 1, n, 1).getDisplayValues();
      const rowVals = sheet.getRange(2, rowCol + 1, n, 1).getDisplayValues();

      const targetSs = cleanId_(spreadsheetFileId);
      let updatedCount = 0;
      // Bangun kolom spreadsheet_row_number di memori lalu tulis SEKALI,
      // bukan satu setValue per baris (sebelumnya O(N) write per delete).
      const newColumn = [];
      for (let i = 0; i < n; i++) {
        let cell = rowVals[i][0];
        if (cleanId_(ssVals[i][0]) === targetSs) {
          const rowNum = parseInt(rowVals[i][0], 10);
          if (!isNaN(rowNum) && rowNum > thresholdRow) {
            cell = rowNum - 1;
            updatedCount++;
          }
        }
        newColumn.push([cell]);
      }
      if (updatedCount > 0) {
        sheet.getRange(2, rowCol + 1, newColumn.length, 1).setValues(newColumn);
      }
      return updatedCount;
    });
  },

  appendAdminAudit: function(row) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.ADMIN_AUDIT_LOG);
    ensureHeaders_(sheet, [
      'created_at',
      'actor',
      'action',
      'year',
      'activity_id',
      'sub_activity_id',
      'folder_id',
      'status',
      'message'
    ]);
    sheet.appendRow([
      row.created_at || new Date().toISOString(),
      row.actor || '',
      row.action || '',
      row.year || '',
      row.activity_id || '',
      row.sub_activity_id || '',
      row.folder_id || '',
      row.status || '',
      row.message || ''
    ]);
    });
  },

  addSubActivityRow: function(payload, folder) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.SUB_ACTIVITIES);
    ensureSubActivityHeaders_(sheet);
    const headers = getHeaders_(sheet);

    const allRowsForYear = readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
      .filter(row => Number(row.year) === Number(payload.year));
    const existingRows = allRowsForYear.filter(row => row.activity_id === payload.activityId);
    let subActivityId = payload.subActivityId || slug_(payload.subActivityName);
    // Cegah tabrakan sub_activity_id: dua nama berbeda bisa menghasilkan slug sama.
    // findConfigRow_/findSubActivity hanya cocokkan yang pertama, jadi ID dobel
    // bikin update/hapus/restore salah sasaran. Beri sufiks unik bila bentrok.
    const existingIds = {};
    existingRows.forEach(function (r) { existingIds[String(r.sub_activity_id || '')] = true; });
    if (existingIds[subActivityId]) {
      let suffix = 2;
      while (existingIds[subActivityId + '_' + suffix]) suffix++;
      subActivityId = subActivityId + '_' + suffix;
    }
    // Nomor final dihitung ulang secara global setelah mutation selesai.
    // Nilai provisional hanya mencegah field kosong sebelum rekonsiliasi berjalan.
    const newSortOrder = allRowsForYear.reduce(function (max, row) {
      const number = Number(row.sort_order) || 0;
      return number > max ? number : max;
    }, 0) + 1;
    const inferredLocalOrder = inferSubActivityLocalOrder_({
      activity_id: payload.activityId,
      sub_activity_name: payload.subActivityName,
      target_sheet_name: payload.targetSheetName,
      parent_folder_name: payload.parentFolderName
    }, { activity_id: payload.activityId });
    const maxLocalOrder = existingRows.reduce(function (max, row) {
      const number = Number(row.local_sort_order) || 0;
      return number > max ? number : max;
    }, 0);
    const localSortOrder = payload.localSortOrder ||
      (inferredLocalOrder < Number.MAX_SAFE_INTEGER ? inferredLocalOrder : maxLocalOrder + 1);
    const shiftedSubActivities = [];

    const values = {
      year: payload.year,
      activity_id: payload.activityId,
      sub_activity_id: subActivityId,
      sub_activity_name: payload.subActivityName,
      formal_archive_name: payload.formalArchiveName || payload.subActivityName,
      folder_id: folder.getId(),
      folder_path: payload.folderPath || folder.getName(),
      no_folder: payload.noFolder || '',
      default_kode_klasifikasi: payload.defaultKodeKlasifikasi || '',
      allow_non_letter_document: payload.allowNonLetterDocument ? 'TRUE' : 'FALSE',
      is_active: 'TRUE',
      sort_order: newSortOrder,
      target_sheet_name: payload.targetSheetName || payload.subActivityName,
      mapping_status: payload.mappingStatus || 'PERLU_REVIEW',
      mapping_note: payload.mappingNote || '',
      rekap_row_number: payload.rekapRowNumber || '',
      inactive_reason: '',
      inactive_at: '',
      parent_folder_id: payload.parentFolderId || '',
      parent_folder_name: payload.parentFolderName || '',
      parent_folder_path: payload.parentFolderPath || '',
      spreadsheet_file_id: payload.spreadsheetFileId ? cleanId_(payload.spreadsheetFileId) : '',
      metadata_locks: payload.metadataLocks || '',
      local_sort_order: localSortOrder
    };
    const row = headers.map(function (header) {
      return values[header] === undefined ? '' : values[header];
    });
    sheet.appendRow(row);
    const resultObj = objectFromHeaders_(headers, row);
    resultObj._shiftedSubActivities = shiftedSubActivities; // Include shifted metadata for cascading
    return resultObj;
    });
  },

  updateActivityMapping: function(payload) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ACTIVITIES);
    if (!sheet) throw new Error('Sheet config_activities tidak ditemukan.');
    const rowIndex = findConfigRow_(sheet, {
      year: payload.year,
      activity_id: payload.activityId
    });
    if (!rowIndex) throw new Error('Kegiatan tidak ditemukan di config.');

    updateConfigRow_(sheet, rowIndex, {
      spreadsheet_file_id: payload.spreadsheetFileId ? cleanId_(payload.spreadsheetFileId) : undefined,
      target_folder_id: payload.targetFolderId ? cleanId_(payload.targetFolderId) : undefined,
      laci_folder_id: payload.laciFolderId ? cleanId_(payload.laciFolderId) : undefined
    });
    return objectFromHeaders_(getHeaders_(sheet), sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
    });
  },

  updateSubActivityMapping: function(payload) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) throw new Error('Sheet config_sub_activities tidak ditemukan.');
    ensureSubActivityHeaders_(sheet);
    const rowIndex = findConfigRow_(sheet, {
      year: payload.year,
      activity_id: payload.activityId,
      sub_activity_id: payload.subActivityId
    });
    if (!rowIndex) throw new Error('Sub-kegiatan tidak ditemukan di config.');

    updateConfigRow_(sheet, rowIndex, {
      folder_id: payload.folderId ? cleanId_(payload.folderId) : undefined,
      folder_path: payload.folderPath,
      target_sheet_name: payload.targetSheetName,
      formal_archive_name: payload.formalArchiveName,
      no_folder: payload.noFolder,
      mapping_status: payload.mappingStatus,
      mapping_note: payload.mappingNote,
      rekap_row_number: payload.rekapRowNumber,
      sub_activity_name: payload.subActivityName,
      parent_folder_id: payload.parentFolderId,
      parent_folder_name: payload.parentFolderName,
      parent_folder_path: payload.parentFolderPath,
      spreadsheet_file_id: payload.spreadsheetFileId ? cleanId_(payload.spreadsheetFileId) : undefined,
      is_active: payload.isActive === undefined ? undefined : (payload.isActive ? 'TRUE' : 'FALSE'),
      inactive_reason: payload.inactiveReason,
      inactive_at: payload.inactiveAt,
      metadata_locks: payload.metadataLocks,
      default_kode_klasifikasi: payload.defaultKodeKlasifikasi,
      sort_order: payload.sortOrder,
      local_sort_order: payload.localSortOrder
    });
    return objectFromHeaders_(getHeaders_(sheet), sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
    });
  },

  getAllSubActivitiesForYear: function (year) {
    const ss = this.getConfigSpreadsheet();
    return readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
      .filter(function (row) { return Number(row.year) === Number(year); });
  },

  bulkUpdateSubActivityNumbering: function (year, assignments) {
    const rows = assignments || [];
    if (!rows.length) return 0;
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) return 0;
    ensureSubActivityHeaders_(sheet);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return 0;
    const headers = values[0].map(normalizeHeader_);
    const yearCol = headers.indexOf('year');
    const activityCol = headers.indexOf('activity_id');
    const subCol = headers.indexOf('sub_activity_id');
    const sortCol = headers.indexOf('sort_order');
    const localCol = headers.indexOf('local_sort_order');
    const rowCol = headers.indexOf('rekap_row_number');
    const byKey = {};
    rows.forEach(function (assignment) {
      byKey[String(assignment.activityId) + '|' + String(assignment.subActivityId)] = assignment;
    });

    let changed = 0;
    for (let index = 1; index < values.length; index++) {
      if (Number(values[index][yearCol]) !== Number(year)) continue;
      const key = String(values[index][activityCol]) + '|' + String(values[index][subCol]);
      const assignment = byKey[key];
      if (!assignment) continue;
      if (localCol !== -1 && assignment.localSortOrder !== undefined &&
          String(values[index][localCol]) !== String(assignment.localSortOrder)) {
        values[index][localCol] = assignment.localSortOrder;
        changed++;
      }
      if (sortCol !== -1 && assignment.globalNumber !== undefined &&
          String(values[index][sortCol]) !== String(assignment.globalNumber)) {
        values[index][sortCol] = assignment.globalNumber;
        changed++;
      }
      if (rowCol !== -1 && assignment.rekapRowNumber !== undefined &&
          String(values[index][rowCol]) !== String(assignment.rekapRowNumber)) {
        values[index][rowCol] = assignment.rekapRowNumber;
        changed++;
      }
    }
    if (changed) {
      sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
    }
    return changed;
  },

  deactivateSubActivity: function(year, activityId, subActivityId, reason) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) return 0;
    ensureSubActivityHeaders_(sheet);
    const rowIndex = findConfigRow_(sheet, {
      year: year,
      activity_id: activityId,
      sub_activity_id: subActivityId
    });
    if (!rowIndex) return 0;
    updateConfigRow_(sheet, rowIndex, {
      is_active: 'FALSE',
      inactive_reason: reason || SUB_ACTIVITY_INACTIVE_REASON.MANUAL,
      inactive_at: new Date().toISOString()
    });
    return 1;
    });
  },

  restoreSubActivity: function(year, activityId, subActivityId) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) return 0;
    ensureSubActivityHeaders_(sheet);
    const rowIndex = findConfigRow_(sheet, {
      year: year,
      activity_id: activityId,
      sub_activity_id: subActivityId
    });
    if (!rowIndex) return 0;
    updateConfigRow_(sheet, rowIndex, {
      is_active: 'TRUE',
      inactive_reason: '',
      inactive_at: ''
    });
    return 1;
    });
  },

  purgeSubActivity: function(year, activityId, subActivityId) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) return 0;
    ensureSubActivityHeaders_(sheet);
    const rowIndex = findConfigRow_(sheet, {
      year: year,
      activity_id: activityId,
      sub_activity_id: subActivityId
    });
    if (!rowIndex) return 0;
    sheet.deleteRow(rowIndex);
    return 1;
    });
  },

  deactivateSubActivityRows: function(year, activityId, folderIds, reason) {
    return withLock_(() => {
    if (!folderIds || !folderIds.length) return 0;
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) return 0;
    ensureSubActivityHeaders_(sheet);
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return 0;
    const headers = values[0].map(normalizeHeader_);
    const folderIdCol = headers.indexOf('folder_id');
    const yearCol = headers.indexOf('year');
    const activityCol = headers.indexOf('activity_id');
    if (folderIdCol === -1) return 0;

    let count = 0;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowYear = String(row[yearCol] || '');
      const rowActivity = String(row[activityCol] || '');
      const rowFolderId = String(row[folderIdCol] || '').trim();
      if (rowYear === String(year) && rowActivity === activityId && folderIds.indexOf(rowFolderId) !== -1) {
        updateConfigRow_(sheet, i + 1, {
          is_active: 'FALSE',
          inactive_reason: reason || SUB_ACTIVITY_INACTIVE_REASON.DRIVE_MISSING,
          inactive_at: new Date().toISOString()
        });
        count++;
      }
    }

    return count;
    });
  },

  purgeInactiveSubActivities: function(options) {
    return withLock_(() => {
    options = options || {};
    const reason = options.reason || SUB_ACTIVITY_INACTIVE_REASON.DRIVE_TRASHED;
    const retentionDays = Number(options.retentionDays || TRASHED_SUB_ACTIVITY_RETENTION_DAYS);
    const cutoffMs = new Date().getTime() - (retentionDays * 24 * 60 * 60 * 1000);

    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES);
    if (!sheet) return { removed: 0, retentionDays: retentionDays };
    ensureSubActivityHeaders_(sheet);

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return { removed: 0, retentionDays: retentionDays };
    const headers = values[0].map(normalizeHeader_);
    const activeCol = headers.indexOf('is_active');
    const reasonCol = headers.indexOf('inactive_reason');
    const inactiveAtCol = headers.indexOf('inactive_at');
    if (activeCol === -1 || reasonCol === -1 || inactiveAtCol === -1) {
      return { removed: 0, retentionDays: retentionDays };
    }

    let removed = 0;
    for (let i = values.length - 1; i >= 1; i--) {
      const row = values[i];
      const inactiveAtMs = Date.parse(row[inactiveAtCol]);
      if (!inactiveAtMs) continue;
      if (!isTrue_(row[activeCol]) && row[reasonCol] === reason && inactiveAtMs <= cutoffMs) {
        sheet.deleteRow(i + 1);
        removed++;
      }
    }

    return {
      removed: removed,
      retentionDays: retentionDays,
      reason: reason,
      cutoff: new Date(cutoffMs).toISOString()
    };
    });
  },
  getTemplateCategories: function () {
    const ss = this.getConfigSpreadsheet();
    return readSheetObjects_(ss, CONFIG_SHEETS.TEMPLATE_CATEGORIES)
      .sort(function (a, b) {
        return Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
          String(a.name || '').localeCompare(String(b.name || ''), 'id', { sensitivity: 'base' });
      });
  },

  saveTemplateCategory: function (payload) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.TEMPLATE_CATEGORIES);
    ensureHeaders_(sheet, ['category_id', 'name', 'color', 'sort_order', 'created_at', 'updated_at']);
    const color = normalizeHexColor_(payload.color, DEFAULT_TEMPLATE_CATEGORY_COLOR);
    const nowIso = new Date().toISOString();

    if (payload.categoryId) {
      const rowIndex = findConfigRow_(sheet, { category_id: payload.categoryId });
      if (rowIndex) {
        updateConfigRow_(sheet, rowIndex, {
          name: payload.name,
          color: color,
          updated_at: nowIso
        });
        return {
          category_id: payload.categoryId,
          name: payload.name,
          color: color
        };
      }
    }

    const allRows = readSheetObjects_(ss, CONFIG_SHEETS.TEMPLATE_CATEGORIES);
    const sortOrder = allRows.length + 1;
    const categoryId = payload.categoryId || Utilities.getUuid();
    sheet.appendRow([categoryId, payload.name, color, sortOrder, nowIso, nowIso]);
    return {
      category_id: categoryId,
      name: payload.name,
      color: color,
      sort_order: sortOrder
    };
    });
  },

  deleteTemplateCategory: function (categoryId) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEETS.TEMPLATE_CATEGORIES);
    if (!sheet) return false;
    const rowIndex = findConfigRow_(sheet, { category_id: categoryId });
    if (!rowIndex) return false;
    sheet.deleteRow(rowIndex);
    this._clearCategoryMapping_(categoryId);
    return true;
    });
  },

  getTemplateCategoryMap: function () {
    const ss = this.getConfigSpreadsheet();
    const rows = readSheetObjects_(ss, CONFIG_SHEETS.TEMPLATE_CATEGORY_MAP);
    const map = {};
    rows.forEach(function (row) { if (row.file_id) map[row.file_id] = row.category_id; });
    return map;
  },

  setTemplateCategory: function (fileId, categoryId) {
    return withLock_(() => {
    const ss = this.getConfigSpreadsheet();
    const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.TEMPLATE_CATEGORY_MAP);
    ensureHeaders_(sheet, ['file_id', 'category_id']);

    const existingRowIndex = findConfigRow_(sheet, { file_id: fileId });
    if (existingRowIndex) {
      if (categoryId) {
        updateConfigRow_(sheet, existingRowIndex, { category_id: categoryId });
        return { file_id: fileId, category_id: categoryId };
      }
      sheet.deleteRow(existingRowIndex);
      return { file_id: fileId, category_id: '' };
    }

    if (categoryId) {
      sheet.appendRow([fileId, categoryId]);
      return { file_id: fileId, category_id: categoryId };
    }
    return { file_id: fileId, category_id: '' };
    });
  },

  enrichTemplatesWithCategories: function (templates) {
    if (!templates || !templates.length) return templates || [];
    const mapping = this.getTemplateCategoryMap();
    const categories = this.getTemplateCategories();
    const catById = {};
    categories.forEach(function (c) { catById[c.category_id] = c; });
    return templates.map(function (t) {
      const catId = mapping[t.id] || '';
      const catName = catId && catById[catId] ? catById[catId].name : '';
      return Object.assign({}, t, { categoryId: catId, categoryName: catName });
    });
  },

  _clearCategoryMapping_: function (categoryId) {
    return withLock_(() => {
    const sheet = this.getConfigSpreadsheet().getSheetByName(CONFIG_SHEETS.TEMPLATE_CATEGORY_MAP);
    if (!sheet) return;
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return;
    const headers = values[0].map(normalizeHeader_);
    const catCol = headers.indexOf('category_id');
    if (catCol === -1) return;
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][catCol] === categoryId) sheet.deleteRow(i + 1);
    }
    });
  },

  // ── Tipe dokumen pendukung (dinamis, dikelola via sheet) ──
  // Header: key | label | form_label | aliases | sort_order | is_active
  DOCUMENT_TYPE_HEADERS: ['key', 'label', 'form_label', 'aliases', 'sort_order', 'is_active'],

  getDocumentTypes: function () {
    return this._readDocumentTypes_().filter(function (t) { return t.is_active; });
  },

  getRemovedDocumentTypes: function () {
    return this._readDocumentTypes_().filter(function (t) { return !t.is_active; });
  },

  _readDocumentTypes_: function () {
    const ss = this.getConfigSpreadsheet();
    const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.DOCUMENT_TYPES);
    ensureHeaders_(sheet, this.DOCUMENT_TYPE_HEADERS);
    let rows = readSheetObjects_(ss, CONFIG_SHEETS.DOCUMENT_TYPES);
    
    // Auto-seed missing ones
    let seededAny = false;
    const existingKeys = {};
    rows.forEach(r => { if (r.key) existingKeys[r.key] = true; });
    
    const missingDocs = REKAP_DOC_COLUMNS.filter(t => !existingKeys[t.key]);
    if (missingDocs.length > 0) {
      this._seedDocumentTypes_(sheet, missingDocs, rows.length);
      seededAny = true;
    }
    
    if (seededAny) {
      rows = readSheetObjects_(ss, CONFIG_SHEETS.DOCUMENT_TYPES);
    }

    return rows
      // Baris harus punya key DAN label (header kolom). Tanpa label = belum
      // lengkap → diabaikan supaya kolom tak dibuat dengan nama key mentah.
      .filter(function (r) { return String(r.key || '').trim() && String(r.label || '').trim(); })
      .map(function (r) {
        const key = String(r.key).trim();
        const label = String(r.label).trim();
        const formLabel = String(r.form_label || '').trim();
        const aliases = String(r.aliases || '')
          .split(/[;,]/)
          .map(function (s) { return s.trim(); })
          .filter(Boolean);
        return {
          key: key,
          label: label,
          formLabel: formLabel || undefined,
          match: [label].concat(aliases),
          sort_order: Number(r.sort_order || 0),
          is_active: isTrue_(r.is_active)
        };
      })
      .sort(function (a, b) { return a.sort_order - b.sort_order; });
  },

  _seedDocumentTypes_: function (sheet, docsToSeed, existingCount) {
    const toSeed = docsToSeed || REKAP_DOC_COLUMNS;
    const startIndex = existingCount || 0;
    const rows = toSeed.map(function (t, i) {
      const labelKey = normalizeHeaderKey_(t.label);
      // Simpan alias tambahan saja (selain yang sama dengan label).
      const extraAliases = (t.match || []).filter(function (m) {
        return normalizeHeaderKey_(m) !== labelKey;
      });
      return [t.key, t.label, t.formLabel || '', extraAliases.join('; '), startIndex + i + 1, 'TRUE'];
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, this.DOCUMENT_TYPE_HEADERS.length).setValues(rows);
    }
  }
};
