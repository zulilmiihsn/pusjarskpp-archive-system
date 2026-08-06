'use strict';

/**
 * SheetHelpers.gs
 * Private helper functions for SpreadsheetService.
 * Extracted from SpreadsheetService.gs for maintainability.
 * All functions use the trailing-underscore convention to signal internal use.
 */

// ── Akses tipe dokumen (dinamis dari sheet, memo per-eksekusi) ──
// GAS membuat scope global baru tiap request, jadi `let` modul = cache 1 request.
let _docTypesActiveMemo_ = null;
let _docTypesRemovedMemo_ = null;

function getRekapDocColumns_() {
  if (_docTypesActiveMemo_) return _docTypesActiveMemo_;
  try {
    _docTypesActiveMemo_ = ConfigRepository.getDocumentTypes();
  } catch (e) {
    console.warn('getRekapDocColumns_ fallback ke default: ' + e.message);
    _docTypesActiveMemo_ = null;
  }
  if (!_docTypesActiveMemo_ || !_docTypesActiveMemo_.length) _docTypesActiveMemo_ = REKAP_DOC_COLUMNS.slice();
  return _docTypesActiveMemo_;
}

function getRemovedRekapDocColumns_() {
  if (_docTypesRemovedMemo_) return _docTypesRemovedMemo_;
  try {
    _docTypesRemovedMemo_ = ConfigRepository.getRemovedDocumentTypes() || [];
  } catch (e) {
    console.warn('getRemovedRekapDocColumns_ gagal: ' + e.message);
    _docTypesRemovedMemo_ = [];
  }
  return _docTypesRemovedMemo_;
}

function resolveRekapDocumentCategory_(categoryName) {
  const normalized = normalizeHeaderKey_(categoryName);
  if (!normalized) return null;
  // Cocokkan terhadap setiap tipe dokumen via alias-nya (substring).
  // Pakai frasa penuh di `match` agar tidak bentrok antar-tipe
  // (mis. "Evaluasi Peserta" tidak nyangkut ke "Data Fix Peserta").
  const docTypes = getRekapDocColumns_();
  for (let i = 0; i < docTypes.length; i++) {
    const column = docTypes[i];
    const aliases = column.match || [];
    for (let j = 0; j < aliases.length; j++) {
      if (normalized.indexOf(normalizeHeaderKey_(aliases[j])) >= 0) return column;
    }
  }
  return null;
}

function ensureRekapDocumentColumns_(sheet) {
  const activeTypes = getRekapDocColumns_();
  const activeKeys = {};
  activeTypes.forEach(function (t) { activeKeys[t.key] = true; });

  // 1) Hapus kolom milik tipe yang dinonaktifkan (is_active=FALSE di sheet).
  //    Cari ulang tiap iterasi karena deleteColumn menggeser indeks.
  //    GUARD: hanya hapus kolom di AREA DOKUMEN (setelah kolom ringkasan
  //    terakhir) supaya kolom ringkasan tak ikut terhapus bila ada tipe yang
  //    namanya kebetulan menyerempet (mis. tipe bernama "Ket").
  const removedTypes = getRemovedRekapDocColumns_();
  if (removedTypes.length) {
    const summaryMap = getRekapHeaderMap_(sheet);
    let summaryMaxCol = REKAP_FALLBACK_START_COL;
    Object.keys(REKAP_SUMMARY_COLUMNS).forEach(function (k) {
      const c = findRekapHeaderColumnFromMap_(summaryMap, REKAP_SUMMARY_COLUMNS[k]);
      if (c) summaryMaxCol = Math.max(summaryMaxCol, c);
    });
    const activeLabelKeys = {};
    activeTypes.forEach(function (t) { activeLabelKeys[normalizeHeaderKey_(t.label)] = true; });
    removedTypes.forEach(function (removed) {
      if (activeKeys[removed.key]) return; // diaktifkan lagi → jangan hapus
      // Targetkan kolom HANYA via label (cara kolom dibuat), bukan alias —
      // mencegah alias generik menghapus kolom tipe aktif yang lain.
      if (activeLabelKeys[normalizeHeaderKey_(removed.label)]) return;
      const col = findRekapHeaderColumn_(sheet, [removed.label]);
      if (col && col > summaryMaxCol) sheet.deleteColumn(col);
    });
  }

  // 2) Tambah kolom untuk tipe aktif yang belum ada.
  const existingMap = getRekapHeaderMap_(sheet);
  let nextCol = Math.max(findLastRekapHeaderColumn_(sheet) + 1, REKAP_FALLBACK_START_COL + 10);
  let addedColsCount = 0;
  activeTypes.forEach(function (column) {
    if (findRekapHeaderColumnFromMap_(existingMap, column.match)) return;
    sheet.getRange(REKAP_HEADER_ROW, nextCol).setValue(column.label);
    sheet.getRange(REKAP_SUBHEADER_ROW, nextCol).setValue('');
    sheet.getRange(REKAP_NUMBERING_ROW, nextCol).setValue(nextCol - REKAP_FALLBACK_START_COL + 1);
    
    // Gabungkan (merge) baris 5 dan 6 secara vertikal
    sheet.getRange(REKAP_HEADER_ROW, nextCol, 2, 1).merge();

    sheet.getRange(REKAP_HEADER_ROW, nextCol, 3, 1)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontFamily('Bookman Old Style')
      .setFontSize(11)
      .setWrap(true);
    sheet.getRange(REKAP_HEADER_ROW, nextCol, 2, 1)
      .setFontWeight('normal')
      .setBackground('#bfbfbf');
    sheet.getRange(REKAP_NUMBERING_ROW, nextCol)
      .setBackground('#ffffff')
      .setFontSize(7)
      .setFontWeight('normal');
    
    sheet.setColumnWidth(nextCol, 90);

    existingMap[normalizeHeaderKey_(column.label)] = nextCol;
    nextCol++;
    addedColsCount++;
  });

  if (addedColsCount > 0) {
    const startCol = nextCol - addedColsCount;
    const noteRow = findNoteRow_(sheet);
    let dataRowsCount = 14;
    if (noteRow && noteRow > REKAP_DATA_START_ROW) {
      dataRowsCount = noteRow - REKAP_DATA_START_ROW;
    }
    
    sheet.getRange(REKAP_DATA_START_ROW, startCol, dataRowsCount, addedColsCount)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setFontFamily('Bookman Old Style')
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  // 3) Perluas merger title (baris 1) supaya tetap rata tengah ke seluruh tabel
  const finalLastCol = findLastRekapHeaderColumn_(sheet);
  if (finalLastCol >= REKAP_FALLBACK_START_COL) {
    const titleRange = sheet.getRange(1, REKAP_FALLBACK_START_COL, 1, finalLastCol - REKAP_FALLBACK_START_COL + 1);
    titleRange.breakApart();
    titleRange.mergeAcross();
  }
}

/**
 * Normalisasi nama sheet Rekap untuk pencocokan toleran:
 * lowercase, samakan "aktif"/"aktip", buang semua non-alfanumerik.
 * "Daftar Berkas Arsip Aktip" / "Daftar Berkas Arsip Aktif" → "daftarberkasarsipaktip".
 * Sheet Detail ("Daftar Isi Berkas Arsip ...") tetap beda karena ada "isi".
 */
function normalizeRekapSheetName_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/aktif/g, 'aktip')
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Cari sheet Rekap secara toleran: coba nama persis dulu, lalu cocokkan ternormalisasi.
 * Mengembalikan null jika workbook memang tak punya sheet Rekap (mis. laci single-detail).
 */
function findRekapSheet_(ss) {
  const sheet = ss.getSheetByName(REKAP_SHEET_NAME);
  if (sheet) return sheet;
  const target = normalizeRekapSheetName_(REKAP_SHEET_NAME);
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (normalizeRekapSheetName_(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

function findOrCreateRekapRow_(sheet, activity, subActivity) {
  const existingRow = findRekapRowForSubActivity_(sheet, subActivity);
  if (existingRow) {
    writeRekapIdentityMarker_(sheet, existingRow, subActivity);
    return existingRow;
  }
  const created = SpreadsheetService.appendRekapRowIfPresent(activity, subActivity);
  if (created && created.rowNumber) return created.rowNumber;
  throw new Error('Baris rekap untuk sub-kegiatan tidak dapat dibuat.');
}

function buildRekapIdentityMarker_(subActivity) {
  const subActivityId = encodeURIComponent(String(subActivity && subActivity.sub_activity_id || ''));
  const folderId = encodeURIComponent(String(subActivity && subActivity.folder_id || ''));
  return 'PORTAL_ARSIP_ROW_ID|' + subActivityId + '|' + folderId;
}

function parseRekapIdentityMarker_(note) {
  const prefix = 'PORTAL_ARSIP_ROW_ID|';
  const line = String(note || '').split(/\r?\n/).find(function (item) {
    return item.indexOf(prefix) === 0;
  });
  if (!line) return null;
  const parts = line.split('|');
  try {
    return {
      subActivityId: decodeURIComponent(parts[1] || ''),
      folderId: decodeURIComponent(parts[2] || '')
    };
  } catch (e) {
    return null;
  }
}

function writeRekapIdentityMarker_(sheet, rowIndex, subActivity) {
  if (!sheet || !rowIndex || !subActivity) return;
  const headerMap = getRekapHeaderMap_(sheet);
  const uraianCol = findRekapHeaderColumnFromMap_(headerMap, ['uraian informasi arsip', 'uraian informasi']) || 4;
  const cell = sheet.getRange(rowIndex, uraianCol);
  const prefix = 'PORTAL_ARSIP_ROW_ID|';
  const existingLines = String(cell.getNote() || '').split(/\r?\n/).filter(function (line) {
    return line && line.indexOf(prefix) !== 0;
  });
  existingLines.push(buildRekapIdentityMarker_(subActivity));
  cell.setNote(existingLines.join('\n'));
}

function buildRekapRowLookup_(sheet) {
  const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
  const headerMap = getRekapHeaderMap_(sheet);
  const nomorBerkasCol = findRekapHeaderColumnFromMap_(headerMap, ['nomor berkas']) || 2;
  const uraianCol = findRekapHeaderColumnFromMap_(headerMap, ['uraian informasi arsip', 'uraian informasi']) || 4;
  const noLaciCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noLaci) || 8;
  const noFolderCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noFolder) || 9;
  const width = Math.max(sheet.getLastColumn(), nomorBerkasCol, uraianCol, noFolderCol);
  if (noteRow <= REKAP_DATA_START_ROW) {
    return {
      startRow: REKAP_DATA_START_ROW,
      noteRow: noteRow,
      nomorBerkasCol: nomorBerkasCol,
      uraianCol: uraianCol,
      noLaciCol: noLaciCol,
      noFolderCol: noFolderCol,
      values: [],
      formulas: [],
      notes: []
    };
  }

  const range = sheet.getRange(REKAP_DATA_START_ROW, 1, noteRow - REKAP_DATA_START_ROW, width);
  const values = range.getDisplayValues();
  let formulas = [];
  let notes = [];
  try { formulas = range.getFormulas(); } catch (e) { formulas = values.map(function (row) { return row.map(function () { return ''; }); }); }
  try { notes = range.getNotes(); } catch (e) { notes = values.map(function (row) { return row.map(function () { return ''; }); }); }
  return {
    startRow: REKAP_DATA_START_ROW,
    noteRow: noteRow,
    nomorBerkasCol: nomorBerkasCol,
    uraianCol: uraianCol,
    noLaciCol: noLaciCol,
    noFolderCol: noFolderCol,
    values: values,
    formulas: formulas,
    notes: notes
  };
}

function getRekapArchiveNumberFromLookup_(lookup, rowIndex, fallback) {
  const index = Number(rowIndex) - lookup.startRow;
  if (index >= 0 && index < lookup.values.length) {
    const noFolder = String((lookup.values[index] || [])[lookup.noFolderCol - 1] || '').trim();
    if (noFolder) return noFolder;
  }
  return String(fallback || '').trim();
}

function getRekapArchiveNumber_(sheet, rowIndex, fallback) {
  const headerMap = getRekapHeaderMap_(sheet);
  const noFolderCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noFolder);
  if (noFolderCol) {
    const noFolder = String(sheet.getRange(rowIndex, noFolderCol).getDisplayValue() || '').trim();
    if (noFolder) return noFolder;
  }
  return String(fallback || '').trim();
}

function getNextAvailableRekapArchiveNumber_(lookup) {
  const used = {};
  (lookup.values || []).forEach(function (row) {
    const value = String((row || [])[lookup.noFolderCol - 1] || '').trim();
    if (/^\d+$/.test(value) && Number(value) > 0) used[Number(value)] = true;
  });
  let next = 1;
  while (used[next]) next++;
  return String(next);
}

function getDetailArchiveNumberForSubActivity_(ss, subActivity) {
  const detailSheet = findExistingDetailSheet_(ss, subActivity);
  if (!detailSheet) return '';
  const data = getDetailDataValues_(detailSheet);
  if (!data) return '';
  const noBerkasCol = data.columnMap.no_berkas;
  if (!noBerkasCol) return '';
  for (let index = 0; index < data.values.length; index++) {
    const value = String((data.values[index] || [])[noBerkasCol - 1] || '').trim();
    if (value) return value;
  }
  return '';
}

/**
 * Menyamakan Nomor Berkas pada semua baris rincian yang benar-benar berisi data.
 * No Folder rincian tidak disentuh karena nilainya adalah Nomor Item Arsip (01, 02, ...).
 */
function writeDetailArchiveNumber_(sheet, archiveNumber) {
  const data = getDetailDataValues_(sheet);
  if (!data || !data.columnMap.no_berkas) return { changedRows: 0, dataRows: 0 };

  const numberCol = data.columnMap.no_berkas;
  const target = Number(archiveNumber);
  const output = [];
  let changedRows = 0;
  let meaningfulRows = 0;

  data.values.forEach(function (row) {
    const metadata = {};
    DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
      const col = data.columnMap[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
      metadata[field] = row[col - 1];
    });
    const current = row[numberCol - 1];
    if (!hasMeaningfulDetailData_(metadata)) {
      output.push([current]);
      return;
    }
    meaningfulRows++;
    if (String(current || '').trim() !== String(target)) changedRows++;
    output.push([target]);
  });

  if (changedRows) {
    sheet.getRange(DETAIL_DATA_START_ROW, numberCol, output.length, 1).setValues(output);
  }
  return { changedRows: changedRows, dataRows: meaningfulRows };
}

function rekapLookupRowHasFolderId_(lookup, index, folderId) {
  if (!folderId) return false;
  const formula = String((lookup.formulas[index] || [])[lookup.noLaciCol - 1] || '');
  const note = parseRekapIdentityMarker_(String((lookup.notes[index] || [])[lookup.uraianCol - 1] || ''));
  return formula.indexOf(folderId) >= 0 || !!(note && note.folderId === folderId);
}

function findStableRekapRowsFromLookup_(lookup, subActivity) {
  const targetSubActivityId = String(subActivity && subActivity.sub_activity_id || '');
  const targetFolderId = String(subActivity && subActivity.folder_id || '');
  const markerRows = [];
  const folderRows = [];

  for (let i = 0; i < lookup.values.length; i++) {
    const note = parseRekapIdentityMarker_(String((lookup.notes[i] || [])[lookup.uraianCol - 1] || ''));
    if (targetSubActivityId && note && note.subActivityId === targetSubActivityId) markerRows.push(lookup.startRow + i);
    if (targetFolderId && rekapLookupRowHasFolderId_(lookup, i, targetFolderId)) folderRows.push(lookup.startRow + i);
  }

  return markerRows.length ? markerRows : folderRows;
}

/**
 * Mengembalikan baris yang terbukti mungkin milik sub-kegiatan.
 * Fuzzy substring dilarang: "Angkatan I" bukan "Angkatan III/VIII/XII".
 */
function findPossibleRekapRowsFromLookup_(lookup, subActivity) {
  if (!lookup || !lookup.values.length) return [];

  const targetName = normalizeComparableText_(getSubActivityFormalName_(subActivity));
  const folderName = normalizeComparableText_(subActivity && subActivity.sub_activity_name);
  const possibleRows = {};

  findStableRekapRowsFromLookup_(lookup, subActivity).forEach(function (rowNumber) {
    possibleRows[rowNumber] = true;
  });

  for (let i = 0; i < lookup.values.length; i++) {
    const uraian = normalizeComparableText_((lookup.values[i] || [])[lookup.uraianCol - 1]);
    const exactName = !!((targetName && uraian === targetName) ||
      (folderName && uraian === folderName));
    if (exactName) {
      possibleRows[lookup.startRow + i] = true;
    }
  }

  return Object.keys(possibleRows).map(Number).sort(function (a, b) { return a - b; });
}

function findRekapRowForSubActivityFromLookup_(lookup, subActivity) {
  if (!lookup || !lookup.values.length) return null;

  const configuredRow = Number(subActivity && subActivity.rekap_row_number);
  const configuredIndex = configuredRow - lookup.startRow;
  const targetName = normalizeComparableText_(getSubActivityFormalName_(subActivity));
  const folderName = normalizeComparableText_(subActivity && subActivity.sub_activity_name);
  const stableRows = findStableRekapRowsFromLookup_(lookup, subActivity);

  if (stableRows.length) {
    if (stableRows.indexOf(configuredRow) !== -1) return configuredRow;
    if (stableRows.length === 1) return stableRows[0];
    const exactStableRows = stableRows.filter(function (rowNumber) {
      const i = rowNumber - lookup.startRow;
      const uraian = normalizeComparableText_((lookup.values[i] || [])[lookup.uraianCol - 1]);
      return !!(targetName && uraian === targetName);
    });
    return exactStableRows.length === 1 ? exactStableRows[0] : null;
  }

  // rekap_row_number hanyalah hint posisi. Sort/insert dapat memindahkan baris.
  // Percaya hint hanya jika identitas pada baris tersebut masih cocok dan tidak
  // membawa marker milik sub-kegiatan lain.
  if (configuredIndex >= 0 && configuredIndex < lookup.values.length) {
    const configuredNote = parseRekapIdentityMarker_(
      String((lookup.notes[configuredIndex] || [])[lookup.uraianCol - 1] || '')
    );
    const targetSubActivityId = String(subActivity && subActivity.sub_activity_id || '');
    const noteConflicts = !!(configuredNote && targetSubActivityId &&
      configuredNote.subActivityId && configuredNote.subActivityId !== targetSubActivityId);
    const configuredUraian = normalizeComparableText_(
      (lookup.values[configuredIndex] || [])[lookup.uraianCol - 1]
    );
    if (!noteConflicts && ((targetName && configuredUraian === targetName) ||
        (folderName && configuredUraian === folderName))) {
      return configuredRow;
    }
  }

  // Hanya kecocokan PERSIS. Nama angka/Romawi tidak boleh diperlakukan sebagai
  // substring: I berbeda dari III, V berbeda dari VIII, X berbeda dari XII.
  const exactTargetRows = [];
  const exactFolderRows = [];
  for (let i = 0; i < lookup.values.length; i++) {
    const uraian = normalizeComparableText_(lookup.values[i][lookup.uraianCol - 1]);
    if (targetName && uraian === targetName) exactTargetRows.push(lookup.startRow + i);
    if (folderName && uraian === folderName) exactFolderRows.push(lookup.startRow + i);
  }
  if (exactTargetRows.length === 1) return exactTargetRows[0];
  if (exactFolderRows.length === 1) return exactFolderRows[0];

  // Tanpa identitas atau nama persis, baris dianggap benar-benar belum ada.
  return null;
}

function findRekapRowForSubActivity_(sheet, subActivity) {
  return findRekapRowForSubActivityFromLookup_(buildRekapRowLookup_(sheet), subActivity);
}

function rekapIdentityNeedsRepair_(lookup, rowIndex, activity, subActivity) {
  const i = rowIndex - lookup.startRow;
  if (i < 0 || i >= lookup.values.length) return true;
  const values = lookup.values[i] || [];
  const notes = lookup.notes[i] || [];
  const marker = parseRekapIdentityMarker_(String(notes[lookup.uraianCol - 1] || ''));
  const expectedId = String(subActivity && subActivity.sub_activity_id || '');
  const expectedFolderId = String(subActivity && subActivity.folder_id || '');
  const expectedName = normalizeComparableText_(getSubActivityFormalName_(subActivity));
  const actualName = normalizeComparableText_(values[lookup.uraianCol - 1]);
  const actualFolderNumber = normalizeComparableText_(values[lookup.noFolderCol - 1]);
  const expectedNumber = actualFolderNumber ||
    normalizeComparableText_(resolveSubActivityArchiveNumber_(subActivity, activity));
  const actualNumber = normalizeComparableText_(values[lookup.nomorBerkasCol - 1]);
  let locks = {};
  try { locks = subActivity && subActivity.metadata_locks ? JSON.parse(subActivity.metadata_locks) : {}; } catch (e) {}

  if (!marker || marker.subActivityId !== expectedId || marker.folderId !== expectedFolderId) return true;
  if (expectedName && actualName !== expectedName) return true;
  if (expectedNumber && actualNumber !== expectedNumber) return true;
  if (locks.noLaci !== false && expectedFolderId && !rekapLookupRowHasFolderId_(lookup, i, expectedFolderId)) return true;
  return false;
}

function buildRekapSummary_(ss, activity, subActivity, metadata) {
  const detailSheet = ensureDetailSheet_(ss, activity, subActivity);
  const detailSummary = summarizeDetailSheet_(detailSheet, subActivity);
  const fallbackDate = parseDateCell_(metadata.tanggal);
  const startDate = detailSummary.startDate || fallbackDate;
  const endDate = detailSummary.endDate || fallbackDate;

  return {
    nomorBerkas: resolveSubActivityArchiveNumber_(subActivity, activity),
    kodeKlasifikasi: subActivity.default_kode_klasifikasi || (typeof DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI !== 'undefined' ? DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI : ''),
    uraian: getSubActivityFormalName_(subActivity) || metadata.uraian_informasi_item || '',
    kurunWaktu: formatDateRange_(startDate, endDate),
    jumlah: detailSummary.sumLembar ? detailSummary.sumLembar + ' lembar' : '',
    noLaci: metadata.no_laci || activity.laci_no || '',
    noFolder: resolveSubActivityArchiveNumber_(subActivity, activity),
    akses: detailSummary.akses || metadata.klasifikasi_akses || 'Terbuka',
    ket: metadata.ket || ''
  };
}

function writeRekapSummary_(sheet, rowIndex, summary) {
  const headerMap = getRekapHeaderMap_(sheet);
  if (summary.nomorBerkas !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas, summary.nomorBerkas, false);
  if (summary.kodeKlasifikasi !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kodeKlasifikasi, summary.kodeKlasifikasi, true);
  if (summary.uraian !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.uraian, summary.uraian, true);
  if (summary.kurunWaktu !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu, summary.kurunWaktu, true);
  if (summary.jumlah !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.jumlah, summary.jumlah, true);
  if (summary.filingCabinet !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet, summary.filingCabinet, true);
  if (summary.noLaci !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noLaci, summary.noLaci, true);
  if (summary.noFolder !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noFolder, summary.noFolder, true);
  if (summary.akses !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, summary.akses, true);
  if (summary.ket !== undefined) setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.ket, summary.ket, false);
  sheet.getRange(rowIndex, REKAP_FALLBACK_START_COL, 1, Math.max(findLastRekapHeaderColumn_(sheet) - REKAP_FALLBACK_START_COL + 1, 1))
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle');
}

function writeRekapIdentity_(sheet, rowIndex, activity, subActivity, options) {
  const headerMap = getRekapHeaderMap_(sheet);
  const respectManualLocks = !!(options && options.respectManualLocks);
  let locks = {};
  try { locks = subActivity && subActivity.metadata_locks ? JSON.parse(subActivity.metadata_locks) : {}; } catch (e) {}
  const noFolder = getRekapArchiveNumber_(
    sheet,
    rowIndex,
    resolveSubActivityArchiveNumber_(subActivity, activity)
  );
  const identity = {
    nomorBerkas: noFolder,
    kodeKlasifikasi: (subActivity && subActivity.default_kode_klasifikasi) || (typeof DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI !== 'undefined' ? DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI : ''),
    uraian: getSubActivityFormalName_(subActivity),
    noLaci: activity && activity.laci_no ? activity.laci_no : '',
    noFolder: noFolder,
    akses: 'Terbuka'
  };

  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas, identity.nomorBerkas, true);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kodeKlasifikasi, identity.kodeKlasifikasi, true);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.uraian, identity.uraian, true);
  if (!respectManualLocks || locks.filingCabinet !== false) {
    setRekapFilingCabinetCell_(sheet, rowIndex, headerMap, activity, false);
  }
  if (!respectManualLocks || locks.noLaci !== false) {
    const noLaciCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noLaci);
    if (noLaciCol && subActivity && subActivity.folder_id) {
      const displayText = String(subActivity.sub_activity_name || '').replace(/"/g, '""').replace(/\n/g, '" & CHAR(10) & "');
      const hsep = typeof formulaSep_ === 'function' ? formulaSep_(sheet.getParent()) : ';';
      sheet.getRange(rowIndex, noLaciCol).setFormula(
        '=HYPERLINK("https://drive.google.com/drive/folders/' + subActivity.folder_id + '"' + hsep + ' "' + displayText + '")'
      );
    } else {
      setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noLaci, identity.noLaci, true);
    }
  }
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noFolder, identity.noFolder, false);
  if (!respectManualLocks || locks.akses !== false) {
    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, identity.akses, false);
  }
  writeRekapIdentityMarker_(sheet, rowIndex, subActivity);
  sheet.getRange(rowIndex, REKAP_FALLBACK_START_COL, 1, Math.max(findLastRekapHeaderColumn_(sheet) - REKAP_FALLBACK_START_COL + 1, 1))
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle');
}

// Sel "No. Filing Cabinet" di Rekap = hyperlink: teks "<laci_no>. Laci <nama kegiatan>",
// link ke folder Drive laci. (Menggantikan angka lemari statis '02' yang lama.)
function buildFilingCabinetFormula_(activity, sheet) {
  if (!activity || !activity.laci_no || !activity.laci_folder_id) return null;
  const text = (activity.laci_no + '. Laci ' + activity.activity_name).replace(/"/g, '""').replace(/\n/g, '" & CHAR(10) & "');
  const s = formulaSep_(sheet);
  return '=HYPERLINK("https://drive.google.com/drive/folders/' + activity.laci_folder_id + '"' + s + ' "' + text + '")';
}

function setRekapFilingCabinetCell_(sheet, rowIndex, headerMap, activity, overwrite) {
  const column = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet);
  if (!column) return;
  const cell = sheet.getRange(rowIndex, column);
  const formula = buildFilingCabinetFormula_(activity, sheet);
  if (formula) {
    if (!overwrite && (cell.getFormula() || cell.getDisplayValue())) return;
    cell.setFormula(formula);
    return;
  }
  // Tanpa folder id: tulis teks polos nama laci, jangan timpa formula/hyperlink yang sudah ada.
  if (cell.getFormula()) return;
  if (!overwrite && cell.getDisplayValue()) return;
  const text = activity && activity.laci_no ? activity.laci_no + '. Laci ' + activity.activity_name : '';
  if (text) cell.setValue(sanitizeCellValue_(text));
}

function setRekapSummaryCell_(sheet, rowIndex, headerMap, aliases, value, overwrite) {
  if (value === null || value === undefined || value === '') return;
  const column = findRekapHeaderColumnFromMap_(headerMap, aliases);
  if (!column) return;
  const cell = sheet.getRange(rowIndex, column);
  if (!overwrite && cell.getDisplayValue()) return;
  if (cell.getFormula()) return; // preserve formulas (gembok tertutup / auto mode)
  cell.setValue(sanitizeCellValue_(value));
}

function getSubActivityFormalName_(subActivity) {
  return (subActivity && (subActivity.formal_archive_name || subActivity.sub_activity_name)) || '';
}

function summarizeDetailSheet_(sheet, subActivity) {
  const data = getDetailDataValues_(sheet);
  if (!data) {
    return { sumLembar: 0, startDate: null, endDate: null, filingCabinet: '', noLaci: '', noFolder: '', akses: '', lastRow: null };
  }

  const values = data.values;
  const columnMap = data.columnMap;
  const dateCol = columnMap.tanggal || (DETAIL_FALLBACK_START_COL + 4);
  const jumlahCol = columnMap.jumlah || (DETAIL_FALLBACK_START_COL + 6);

  let sumLembar = 0;
  let startDate = null;
  let endDate = null;
  let lastRow = null;
  const aksesByKey = {};

  // SATU lintasan untuk semua agregat — objek metadata per baris dibangun
  // sekali saja (sebelumnya 4x). Hemat & lebih mudah dibaca.
  values.forEach(function (row) {
    const metadata = {};
    DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
      const col = columnMap[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
      const val = row[col - 1];
      metadata[field] = (val === null || val === undefined) ? '' : String(val).trim();
    });

    // Jumlah (lembar): cukup parse nilai yang ada di kolom jumlah
    const raw = row[jumlahCol - 1];
    const n = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/[^0-9.\-]/g, ''));
    if (!isNaN(n)) sumLembar += n;

    // Agregat lain pakai predikat "baris bermakna".
    if (!hasMeaningfulDetailData_(metadata)) return;
    lastRow = metadata;

    const dateValue = parseDateCell_(row[dateCol - 1]);
    if (dateValue) {
      if (!startDate || dateValue.getTime() < startDate.getTime()) startDate = dateValue;
      if (!endDate || dateValue.getTime() > endDate.getTime()) endDate = dateValue;
    }

    const akses = normalizeAccessSummaryValue_(metadata.klasifikasi_akses);
    if (akses) aksesByKey[normalizeLooseLabel_(akses)] = akses;
  });

  return {
    sumLembar: sumLembar,
    startDate: startDate,
    endDate: endDate,
    akses: formatAccessSummary_(aksesByKey),
    lastRow: lastRow
  };
}

function getDetailDataValues_(sheet) {
  const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
  if (noteRow <= DETAIL_DATA_START_ROW) return null;
  const width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
  const columnMap = getDetailColumnMap_(sheet, width);
  const rowCount = noteRow - DETAIL_DATA_START_ROW;
  const values = sheet.getRange(DETAIL_DATA_START_ROW, 1, rowCount, width).getValues();
  return { values: values, columnMap: columnMap };
}

function normalizeAccessSummaryValue_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const key = normalizeLooseLabel_(text);
  if (key.indexOf('terbuka') >= 0) return 'Terbuka';
  if (key.indexOf('biasa') >= 0) return 'Terbuka';
  if (key.indexOf('terbatas') >= 0) return 'Terbatas';
  if (key.indexOf('rahasia') >= 0) return 'Rahasia';
  return text.replace(/\s+/g, ' ');
}

function formatAccessSummary_(aksesByKey) {
  const priority = ['terbuka', 'biasa', 'terbatas', 'rahasia'];
  const values = [];
  priority.forEach(function (key) {
    if (aksesByKey[key]) values.push(aksesByKey[key]);
  });
  Object.keys(aksesByKey).sort().forEach(function (key) {
    if (priority.indexOf(key) === -1) values.push(aksesByKey[key]);
  });
  return values.join(' & ');
}

function parseDateCell_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    if (value.getFullYear() <= 1900) return null;
    return value;
  }
  const text = String(value).trim();
  if (!text) return null;
  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) {
    const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    if (d.getFullYear() <= 1900) return null;
    return d;
  }
  const parsed = new Date(text);
  if (isNaN(parsed.getTime()) || parsed.getFullYear() <= 1900) return null;
  return parsed;
}

function formatDateRange_(startDate, endDate) {
  if (!startDate && !endDate) return '';
  const start = startDate || endDate;
  const end = endDate || startDate;
  const startText = formatRekapDate_(start);
  const endText = formatRekapDate_(end);
  return startText === endText ? startText : startText + ' - ' + endText;
}

function formatRekapDate_(date) {
  if (!date) return '';
  const parsed = (date instanceof Date) ? date : parseDateCell_(date);
  if (!parsed || isNaN(parsed.getTime())) return String(date);
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return parsed.getDate() + ' ' + months[parsed.getMonth()] + ' ' + parsed.getFullYear();
}

function findRekapHeaderColumn_(sheet, aliases) {
  return findRekapHeaderColumnFromMap_(getRekapHeaderMap_(sheet), aliases);
}

function findRekapHeaderColumnFromMap_(headerMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizeHeaderKey_(aliases[i]);
    if (headerMap[key]) return headerMap[key];
  }
  return null;
}

function getRekapHeaderMap_(sheet) {
  const width = Math.max(sheet.getLastColumn(), REKAP_FALLBACK_START_COL + 13);
  const headerRows = sheet.getRange(REKAP_HEADER_ROW, 1, 2, width).getDisplayValues();
  const map = {};
  for (let c = 0; c < width; c++) {
    const primary = normalizeHeaderKey_(headerRows[0][c]);
    const secondary = normalizeHeaderKey_(headerRows[1][c]);
    const combined = normalizeHeaderKey_([headerRows[0][c], headerRows[1][c]].filter(Boolean).join(' '));
    if (primary) map[primary] = c + 1;
    if (secondary) map[secondary] = c + 1;
    if (combined) map[combined] = c + 1;
  }
  return map;
}

function findLastRekapHeaderColumn_(sheet) {
  const width = Math.max(sheet.getLastColumn(), REKAP_FALLBACK_START_COL + 9);
  const headerRows = sheet.getRange(REKAP_HEADER_ROW, 1, 2, width).getDisplayValues();
  for (let c = width - 1; c >= 0; c--) {
    if (String(headerRows[0][c] || headerRows[1][c] || '').trim()) return c + 1;
  }
  return REKAP_FALLBACK_START_COL + 9;
}

function normalizeHeaderKey_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function normalizeComparableText_(value) {
  return normalizeHeaderKey_(value).replace(/\btahun\s+20\d{2}\b/g, '').replace(/\s+/g, ' ').trim();
}

function getArchiveSpreadsheetId_(activity, subActivity) {
  const subSpreadsheetId = subActivity && subActivity.spreadsheet_file_id
    ? cleanId_(subActivity.spreadsheet_file_id)
    : '';
  const activitySpreadsheetId = activity && activity.spreadsheet_file_id
    ? cleanId_(activity.spreadsheet_file_id)
    : '';
  const spreadsheetId = subSpreadsheetId || activitySpreadsheetId;
  if (!spreadsheetId) throw new Error('Spreadsheet arsip belum dipetakan untuk kegiatan/sub-kegiatan ini.');
  return spreadsheetId;
}

function ensureDetailSheet_(ss, activity, subActivity) {
  const targetName = normalizeSheetName_(subActivity.target_sheet_name || subActivity.sub_activity_name || DEFAULT_DETAIL_SHEET_NAME);
  let sheet = ss.getSheetByName(targetName);
  if (sheet) return sheet;

  // Cek-lalu-bikin di atas tidak atomik (TOCTOU). Dua eksekusi yang jalan
  // bersamaan bisa sama-sama lolos cek getSheetByName(null) lalu sama-sama bikin
  // sheet bernama sama — yang kalah dapat error native "Sudah ada sheet dengan
  // nama tersebut" dan simpan gagal. Tangani race-nya: bersihkan sheet orphan
  // dari copyTo yang gagal, lalu pakai sheet yang sudah ada bikinan eksekusi lain.

  // Jangan pernah pakai sheet Rekap sebagai template Detail — bikin layout Detail berantakan.
  const template = ss.getSheetByName('Template Detail Item') ||
    ss.getSheetByName(DEFAULT_DETAIL_SHEET_NAME) ||
    null;

  if (template && template.getName() !== REKAP_SHEET_NAME) {
    const copied = template.copyTo(ss);
    try {
      sheet = copied.setName(targetName);
    } catch (e) {
      // Kalah race: setName gagal karena nama sudah dipakai. Buang salinan orphan
      // ("Copy of ...") dan pakai sheet yang sudah ada.
      try { ss.deleteSheet(copied); } catch (_) {}
      SpreadsheetApp.flush();
      const racedSheet = ss.getSheetByName(targetName);
      if (racedSheet) return racedSheet;
      throw e;
    }
    clearDetailRows_(sheet);
    const clearNoteRow = findNoteRow_(sheet) || DETAIL_NOTE_FALLBACK_ROW;
    const clearDataRows = Math.max(clearNoteRow - DETAIL_DATA_START_ROW, 24);
    sheet.getRange(DETAIL_DATA_START_ROW, getDetailStartColumn_(sheet) + DETAIL_ITEM_NUMBER_OFFSET, clearDataRows, 1)
      .setNumberFormat('00');
  } else {
    try {
      sheet = ss.insertSheet(targetName);
    } catch (e) {
      SpreadsheetApp.flush();
      const racedSheet = ss.getSheetByName(targetName);
      if (racedSheet) return racedSheet;
      throw e;
    }
    formatBasicDetailSheet_(sheet, activity, subActivity);
  }
  return sheet;
}

function findExistingDetailSheet_(ss, subActivity) {
  const candidates = [
    subActivity && subActivity.target_sheet_name,
    subActivity && subActivity.effective_target_sheet_name,
    subActivity && subActivity.sub_activity_name,
    DEFAULT_DETAIL_SHEET_NAME
  ].filter(Boolean);
  for (let i = 0; i < candidates.length; i++) {
    const sheet = ss.getSheetByName(normalizeSheetName_(candidates[i]));
    if (sheet) return sheet;
  }
  return null;
}

function clearDetailRows_(sheet) {
  const noteRow = findNoteRow_(sheet) || DETAIL_NOTE_FALLBACK_ROW;
  const maxRows = Math.max(noteRow - DETAIL_DATA_START_ROW, 1);
  sheet.getRange(
    DETAIL_DATA_START_ROW,
    getDetailStartColumn_(sheet),
    maxRows,
    DETAIL_FIELD_ORDER.length
  ).clearContent();
}

function formatBasicDetailSheet_(sheet, activity, subActivity) {
  sheet.clear();
  sheet.setFrozenRows(8);
  sheet.getRange('B2:N2').merge().setValue('Daftar Isi Berkas Arsip Aktip');
  sheet.getRange('B2:N2')
    .setFontFamily('Bookman Old Style')
    .setFontSize(10)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('B4').setValue('Unit pengolah : Latbang');
  sheet.getRange('B4')
    .setFontFamily('Bookman Old Style')
    .setFontSize(10)
    .setFontWeight('normal');

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

  // Format umum untuk header (B6:N8)
  sheet.getRange('B6:N8')
    .setBackground('#b4c6e7')
    .setFontWeight('normal')
    .setFontFamily('Bookman Old Style')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true);

  // Ukuran font spesifik: Baris judul (6-7) dan Baris nomor (8)
  sheet.getRange('B6:N7').setFontSize(10);
  sheet.getRange('B8:N8').setFontSize(7);

  const dataRows = 24;
  sheet.getRange(DETAIL_DATA_START_ROW, DETAIL_FALLBACK_START_COL, dataRows, DETAIL_FIELD_ORDER.length)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily('Bookman Old Style')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);
  const noBerkasRows = [];
  const itemRows = [];
  for (let i = 1; i <= dataRows; i++) {
    noBerkasRows.push(['']);
    itemRows.push(['']);
  }
  sheet.getRange(DETAIL_DATA_START_ROW, DETAIL_FALLBACK_START_COL, dataRows, 1)
    .setValues(noBerkasRows)
    .setHorizontalAlignment('center');
  sheet.getRange(DETAIL_DATA_START_ROW, DETAIL_FALLBACK_START_COL + DETAIL_ITEM_NUMBER_OFFSET, dataRows, 1)
    .setNumberFormat('00')
    .setValues(itemRows)
    .setHorizontalAlignment('center');

  sheet.setRowHeight(2, 22);
  sheet.setRowHeight(4, 22);
  sheet.setRowHeights(6, 2, 31);
  sheet.setRowHeight(8, 18);
  sheet.setRowHeights(DETAIL_DATA_START_ROW, dataRows, 18);
  [55, 86, 78, 270, 108, 90, 60, 84, 84, 84, 84, 100, 72].forEach(function (widthPx, index) {
    sheet.setColumnWidth(DETAIL_FALLBACK_START_COL + index, widthPx);
  });

  writeBasicDetailNotes_(sheet, DETAIL_NOTE_FALLBACK_ROW);
}

function writeBasicDetailNotes_(sheet, startRow) {
  const notes = [
    ['Keterangan Petunjuk Pengisian:'],
    ['Kolom (9), diisi dengan nomor laci pada Filing Cabinet;'],
    ['Kolom (10), diisi dengan nomor folder Arsip;'],
    ['Kolom (11), diisi dengan klasifikasi keamanan seperti terbuka, terbatas, dan rahasia.'],
    ['Kolom (12), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik, dan digital.']
  ];
  sheet.getRange(startRow, 3, notes.length, 1).setValues(notes);
  sheet.getRange(startRow, 3, notes.length, 1)
    .setFontFamily('Bookman Old Style')
    .setFontSize(10)
    .setFontWeight('normal');
}

function findWritableDetailRow_(sheet) {
  const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
  const startCol = getDetailStartColumn_(sheet);
  const blank = findFirstBlankInRange_(
    sheet,
    DETAIL_DATA_START_ROW,
    noteRow - 1,
    startCol + DETAIL_WRITABLE_CHECK_OFFSET
  );
  if (blank) return blank;
  sheet.insertRowBefore(noteRow);
  invalidateNoteRowCache_(sheet);
  return noteRow;
}

// Global execution cache for note rows to prevent repeated spreadsheet reads
const noteRowCache_ = {};

function invalidateNoteRowCache_(sheet) {
  const sheetId = sheet.getParent().getId() + ':' + sheet.getName();
  delete noteRowCache_[sheetId];
}

function findNoteRow_(sheet) {
  const sheetId = sheet.getParent().getId() + ':' + sheet.getName();
  if (noteRowCache_[sheetId] !== undefined) {
    return noteRowCache_[sheetId];
  }
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, lastRow, Math.min(sheet.getLastColumn(), 6)).getValues();
  for (let r = 0; r < values.length; r++) {
    const rowStr = values[r].join(' ');
    if (rowStr.indexOf('Keterangan Petunjuk Pengisian') >= 0) {
      noteRowCache_[sheetId] = r + 1;
      return r + 1;
    }
  }
  noteRowCache_[sheetId] = null;
  return null;
}

function findFirstBlankInRange_(sheet, startRow, endRow, col) {
  if (endRow < startRow) return null;
  const values = sheet.getRange(startRow, col, endRow - startRow + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0]).trim()) return startRow + i;
  }
  return null;
}

function getDetailStartColumn_(sheet) {
  const width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
  const headers = sheet.getRange(DETAIL_HEADER_ROW, 1, 1, width).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === 'no berkas') return i + 1;
  }
  return DETAIL_FALLBACK_START_COL;
}

function getDetailColumnMap_(sheet, width) {
  const top = sheet.getRange(DETAIL_HEADER_ROW, 1, 1, width).getDisplayValues()[0];
  const bottom = sheet.getRange(DETAIL_HEADER_ROW + 1, 1, 1, width).getDisplayValues()[0];
  const labels = top.map(function (value, index) {
    return normalizeLooseLabel_(String(value || '') + ' ' + String(bottom[index] || ''));
  });
  const aliases = {
    no_berkas: ['no berkas'],
    nomor_item_arsip: ['nomor item arsip'],
    kode_klasifikasi: ['kode klasifikasi'],
    uraian_informasi_item: ['uraian informasi item', 'uraian informasi berkas', 'uraian informasi arsip', 'uraian informasi'],
    tanggal: ['tgl', 'tanggal'],
    tingkat_perkembangan: ['tingkat perkembangan', 'tingkat pengembangan'],
    jumlah: ['jumlah'],
    no_filing_cabinet: ['no filing cabinet', 'no. filing cabinet', 'no fc', 'no. fc'],
    no_laci: ['no laci', 'no. laci'],
    no_folder: ['no folder', 'no. folder'],
    klasifikasi_akses: ['klasifikasi keamanan akses arsip', 'keamanan akses arsip', 'klasifikasi akses'],
    lokasi_simpan: ['ket lokasi simpan', 'lokasi simpan'],
    jumlah_satuan: ['satuan']
  };
  const map = {};
  Object.keys(aliases).forEach(function (field) {
    const col = findDetailColumnByAliases_(labels, aliases[field]);
    if (col) map[field] = col;
  });
  return map;
}

function findDetailColumnByAliases_(labels, aliases) {
  const normalizedAliases = aliases.map(normalizeLooseLabel_);
  for (let i = 0; i < labels.length; i++) {
    const label = labels[i];
    if (!label) continue;
    for (let j = 0; j < normalizedAliases.length; j++) {
      if (label.indexOf(normalizedAliases[j]) >= 0) return i + 1;
    }
  }
  return null;
}

function normalizeLooseLabel_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/&/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function hasMeaningfulDetailData_(metadata) {
  return !!(
    metadata.nomor_item_arsip ||
    metadata.kode_klasifikasi ||
    metadata.uraian_informasi_item ||
    metadata.tanggal ||
    metadata.jumlah ||
    metadata.lokasi_simpan
  );
}

// Paksa sel angka-murni di kolom kunci jadi tipe Number, biar native sort
// numerik (1,2,10) bukan lexicographic teks (1,10,2). Sel kosong/non-angka/
// formula dibiarkan. Sumber bug "hasil init ga urut": Nomor Berkas ditulis
// sebagai teks oleh sanitizeCellValue_.
function coerceColumnNumeric_(sheet, startRow, col, numRows) {
  if (numRows < 1) return;
  const rng = sheet.getRange(startRow, col, numRows, 1);
  const vals = rng.getValues();
  const formulas = rng.getFormulas();
  let changed = false;
  const out = vals.map(function (r, i) {
    const v = r[0];
    if (formulas[i][0]) return [v];          // jangan sentuh sel berformula
    if (typeof v === 'number') return [v];
    const s = String(v).trim();
    if (s !== '' && /^\d+$/.test(s)) { changed = true; return [Number(s)]; }
    return [v];
  });
  if (changed) rng.setValues(out);
}

function sortRekapSheetByNomorBerkas_(sheet) {
  try {
    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const lastDataRow = noteRow - 1;
    if (lastDataRow < REKAP_DATA_START_ROW) return;
    const lastCol = findLastRekapHeaderColumn_(sheet);
    const numRows = lastDataRow - REKAP_DATA_START_ROW + 1;
    const numCols = lastCol - REKAP_FALLBACK_START_COL + 1;
    coerceColumnNumeric_(sheet, REKAP_DATA_START_ROW, 2, numRows); // kolom B = Nomor Berkas
    const range = sheet.getRange(REKAP_DATA_START_ROW, REKAP_FALLBACK_START_COL, numRows, numCols);
    // Use native Sheets sort to preserve formulas and formatting
    range.sort({column: 2, ascending: true});
  } catch (e) {
    console.warn('sortRekapSheetByNomorBerkas_: ' + e.message);
  }
}

function colNumToLetter_(col) {
  let letter = '';
  while (col > 0) {
    col--;
    letter = String.fromCharCode(65 + (col % 26)) + letter;
    col = Math.floor(col / 26);
  }
  return letter;
}

function detailSheetFormulaRef_(detailSheet) {
  const name = detailSheet.getName();
  const escaped = name.replace(/'/g, "''");
  return "'" + escaped + "'";
}

function getDetailColumnLetter_(sheet, field, fallbackLetter) {
  try {
    const width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
    const colMap = getDetailColumnMap_(sheet, width);
    const col = colMap[field];
    if (col) return colNumToLetter_(col);
  } catch (e) {
    console.warn('getDetailColumnLetter_ error for ' + field + ': ' + e.message);
  }
  return fallbackLetter;
}

// Pemisah argumen formula beda per locale spreadsheet: id_ID (& semua locale
// desimal-koma) pakai ';', US/desimal-titik pakai ','.
// Project ini standalone, jadi Document Properties dapat null. Gunakan locale
// langsung dan memo per-eksekusi; jangan menulis sel spreadsheet sebagai probe.
function formulaSep_(ssOrSheet) {
  var ss;
  if (!ssOrSheet) return ';';
  if (typeof ssOrSheet.getParent === 'function') {
    try { ss = ssOrSheet.getParent(); } catch (e) { return ';'; }
  } else {
    ss = ssOrSheet;
  }
  if (!ss || typeof ss.getId !== 'function') return ';';

  // Memo per-eksekusi
  if (!formulaSep_._m) formulaSep_._m = {};
  var ssId;
  try { ssId = ss.getId(); } catch (e) { return ';'; }
  if (formulaSep_._m[ssId] !== undefined) return formulaSep_._m[ssId];

  var sep = ';';
  try {
    var locale = String(ss.getSpreadsheetLocale() || '').toLowerCase();
    sep = /^(id|de|fr|it|nl|pt|es|ru|tr|pl|cs|da|fi|el|hu|nb|nn|ro|sk|sl|sv|uk|vi)(_|-|$)/.test(locale)
      ? ';'
      : ',';
  } catch (e) {
    console.warn('formulaSep_ locale gagal: ' + e.message);
  }
  formulaSep_._m[ssId] = sep;
  return sep;
}

function buildKurunWaktuFormula_(detailSheet) {
  const ref = detailSheetFormulaRef_(detailSheet);
  const colLetter = getDetailColumnLetter_(detailSheet, 'tanggal', 'F');
  const rangeRef = ref + '!' + colLetter + '9:' + colLetter;
  const s = formulaSep_(detailSheet);
  return "=IF(COUNT(" + rangeRef + ")>0" + s + " TEXT(MIN(" + rangeRef + ")" + s + " \"d mmmm yyyy\") & \" - \" & TEXT(MAX(" + rangeRef + ")" + s + " \"d mmmm yyyy\")" + s + " \"\")";
}

// Jumlah = SUM kolom "jumlah" (lembar) tiap dokumen, bukan hitungan baris.
function buildJumlahFormula_(detailSheet) {
  const ref = detailSheetFormulaRef_(detailSheet);
  const jumlahLetter = getDetailColumnLetter_(detailSheet, 'jumlah', 'H');
  const sumRange = ref + '!' + jumlahLetter + '9:' + jumlahLetter;
  const s = formulaSep_(detailSheet);
  return "=IF(COUNT(" + sumRange + ")>0" + s + " SUM(" + sumRange + ") & \" lembar\"" + s + " \"\")";
}

function buildAksesFormula_(detailSheet) {
  const ref = detailSheetFormulaRef_(detailSheet);
  const colLetter = getDetailColumnLetter_(detailSheet, 'klasifikasi_akses', 'M');
  const rangeRef = ref + '!' + colLetter + '9:' + colLetter;
  const s = formulaSep_(detailSheet);
  return "=IFERROR(TEXTJOIN(\" & \"" + s + " TRUE" + s + " UNIQUE(FILTER(" + rangeRef + s + " " + rangeRef + "<>\"\")))" + s + " \"Terbuka\")";
}

function setRekapFormulaCell_(sheet, rowIndex, headerMap, aliases, formula) {
  const column = findRekapHeaderColumnFromMap_(headerMap, aliases);
  if (!column) return;
  const cell = sheet.getRange(rowIndex, column);
  cell.setFormula(formula);
}

function setRekapStaticCell_(sheet, rowIndex, headerMap, aliases, value) {
  if (value === null || value === undefined || value === '') return;
  const column = findRekapHeaderColumnFromMap_(headerMap, aliases);
  if (!column) return;
  const cell = sheet.getRange(rowIndex, column);
  cell.setValue(sanitizeCellValue_(value));
}

// Cari nomor baris detail berdasarkan URL/ID file di kolom lokasi_simpan.
// Dipakai setelah sort agar nomor baris yang dicatat ke log tidak basi.
// Mengembalikan 0 jika tidak ketemu.
function locateDetailRowByUrl_(sheet, url) {
  if (!url) return 0;
  const startCol = getDetailStartColumn_(sheet);
  const lokasiIdx = DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
  if (lokasiIdx < 0) return 0;
  const dataRows = Math.max(0, sheet.getLastRow() - DETAIL_DATA_START_ROW + 1);
  if (dataRows <= 0) return 0;
  // Pakai segmen ID Drive (>=25 char) bila ada, kalau tidak pakai URL utuh.
  const m = String(url).match(/[-\w]{25,}/);
  const needle = m ? m[0] : String(url);
  const rich = sheet.getRange(DETAIL_DATA_START_ROW, startCol + lokasiIdx, dataRows, 1).getRichTextValues();
  for (let i = 0; i < rich.length; i++) {
    const cell = rich[i][0];
    const cellUrl = cell ? (cell.getLinkUrl() || cell.getText() || '') : '';
    if (cellUrl.indexOf(needle) !== -1) return DETAIL_DATA_START_ROW + i;
  }
  return 0;
}

function sortDetailSheetByNomorItemArsip_(sheet) {
  try {
    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const lastDataRow = noteRow - 1;
    if (lastDataRow < DETAIL_DATA_START_ROW) return;
    const startCol = getDetailStartColumn_(sheet);
    const lastCol = Math.max(sheet.getLastColumn(), startCol + DETAIL_FIELD_ORDER.length - 1);
    const numRows = lastDataRow - DETAIL_DATA_START_ROW + 1;
    const numCols = lastCol - startCol + 1;
    const itemCol = startCol + DETAIL_ITEM_NUMBER_OFFSET;
    coerceColumnNumeric_(sheet, DETAIL_DATA_START_ROW, itemCol, numRows); // Nomor Item Arsip
    const colMap = getDetailColumnMap_(sheet, lastCol);
    if (colMap['jumlah']) coerceColumnNumeric_(sheet, DETAIL_DATA_START_ROW, colMap['jumlah'], numRows);

    const range = sheet.getRange(DETAIL_DATA_START_ROW, startCol, numRows, numCols);
    // Use native Sheets sort to preserve formulas and formatting
    range.sort({column: itemCol, ascending: true});
    sheet.getRange(DETAIL_DATA_START_ROW, itemCol, numRows, 1).setNumberFormat('00');
  } catch (e) {
    console.warn('sortDetailSheetByNomorItemArsip_: ' + e.message);
  }
}
