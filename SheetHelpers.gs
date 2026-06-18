'use strict';

/**
 * SheetHelpers.gs
 * Private helper functions for SpreadsheetService.
 * Extracted from SpreadsheetService.gs for maintainability.
 * All functions use the trailing-underscore convention to signal internal use.
 */

function resolveRekapDocumentCategory_(categoryName) {
  const normalized = normalizeHeaderKey_(categoryName);
  if (!normalized) return null;
  if (normalized.indexOf('data fix') >= 0 || normalized.indexOf('peserta') >= 0) return REKAP_DOC_COLUMNS[0];
  if (normalized.indexOf('materi') >= 0) return REKAP_DOC_COLUMNS[1];
  if (normalized.indexOf('laporan') >= 0) return REKAP_DOC_COLUMNS[2];
  if (normalized.indexOf('sertifikat') >= 0) return REKAP_DOC_COLUMNS[3];
  return null;
}

function ensureRekapDocumentColumns_(sheet) {
  const existingMap = getRekapHeaderMap_(sheet);
  let nextCol = Math.max(findLastRekapHeaderColumn_(sheet) + 1, REKAP_FALLBACK_START_COL + 10);
  REKAP_DOC_COLUMNS.forEach(function (column) {
    if (findRekapHeaderColumnFromMap_(existingMap, column.match)) return;
    sheet.getRange(REKAP_HEADER_ROW, nextCol).setValue(column.label);
    sheet.getRange(REKAP_SUBHEADER_ROW, nextCol).setValue('');
    sheet.getRange(REKAP_NUMBERING_ROW, nextCol).setValue(nextCol - REKAP_FALLBACK_START_COL + 1);
    sheet.getRange(REKAP_HEADER_ROW, nextCol, 3, 1)
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontFamily('Arial')
      .setFontSize(11)
      .setFontWeight('bold')
      .setBackground('#bfbfbf')
      .setWrap(true);
    sheet.getRange(REKAP_NUMBERING_ROW, nextCol)
      .setBackground('#ffffff')
      .setFontSize(7)
      .setFontWeight('bold');
    existingMap[normalizeHeaderKey_(column.label)] = nextCol;
    nextCol++;
  });
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
  let sheet = ss.getSheetByName(REKAP_SHEET_NAME);
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
  if (existingRow) return existingRow;
  const created = SpreadsheetService.appendRekapRowIfPresent(activity, subActivity);
  if (created && created.rowNumber) return created.rowNumber;
  throw new Error('Baris rekap untuk sub-kegiatan tidak dapat dibuat.');
}

function findRekapRowForSubActivity_(sheet, subActivity) {
  const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
  if (noteRow <= REKAP_DATA_START_ROW) return null;

  const configuredRow = Number(subActivity && subActivity.rekap_row_number);
  if (configuredRow >= REKAP_DATA_START_ROW && configuredRow < noteRow) return configuredRow;

  const headerMap = getRekapHeaderMap_(sheet);
  const nomorBerkasCol = findRekapHeaderColumnFromMap_(headerMap, ['nomor berkas']) || 2;
  const uraianCol = findRekapHeaderColumnFromMap_(headerMap, ['uraian informasi arsip', 'uraian informasi']) || 4;
  const width = Math.max(sheet.getLastColumn(), nomorBerkasCol, uraianCol);
  const values = sheet.getRange(REKAP_DATA_START_ROW, 1, noteRow - REKAP_DATA_START_ROW, width).getDisplayValues();
  const targetName = normalizeComparableText_(getSubActivityFormalName_(subActivity));
  const folderName = normalizeComparableText_(subActivity && subActivity.sub_activity_name);
  const targetSortOrder = normalizeComparableText_(subActivity && subActivity.sort_order);
  const targetNoBerkas = normalizeComparableText_(subActivity && subActivity.no_berkas);

  // Kecocokan PERSIS dulu, di SELURUH baris, urut prioritas. Substring TIDAK boleh
  // dipakai lebih awal: nama mirip ("Laporan" vs "Laporan Akhir", atau "1." vs "11."
  // setelah dinormalisasi) bisa membuat ringkasan tertulis ke baris sub-kegiatan
  // yang SALAH dan diam-diam merusak data sub lain.
  for (let i = 0; i < values.length; i++) {
    const uraian = normalizeComparableText_(values[i][uraianCol - 1]);
    if (targetName && uraian === targetName) return REKAP_DATA_START_ROW + i;
  }
  for (let i = 0; i < values.length; i++) {
    const uraian = normalizeComparableText_(values[i][uraianCol - 1]);
    if (folderName && uraian === folderName) return REKAP_DATA_START_ROW + i;
  }
  for (let i = 0; i < values.length; i++) {
    const nomorBerkas = normalizeComparableText_(values[i][nomorBerkasCol - 1]);
    if (targetNoBerkas && nomorBerkas === targetNoBerkas) return REKAP_DATA_START_ROW + i;
    if (targetSortOrder && nomorBerkas === targetSortOrder) return REKAP_DATA_START_ROW + i;
  }

  // Fallback substring HANYA bila tidak ada kecocokan persis DAN hasilnya unik
  // (tepat satu baris). Kalau ambigu, kembalikan null — lebih baik gagal lembut
  // daripada menebak dan merusak baris milik sub-kegiatan lain.
  if (targetName) {
    let fuzzyRow = -1;
    let fuzzyCount = 0;
    for (let i = 0; i < values.length; i++) {
      const uraian = normalizeComparableText_(values[i][uraianCol - 1]);
      if (uraian && (uraian.indexOf(targetName) >= 0 || targetName.indexOf(uraian) >= 0)) {
        fuzzyCount++;
        fuzzyRow = i;
      }
    }
    if (fuzzyCount === 1) return REKAP_DATA_START_ROW + fuzzyRow;
  }
  return null;
}

function buildRekapSummary_(ss, activity, subActivity, metadata) {
  const detailSheet = ensureDetailSheet_(ss, activity, subActivity);
  const detailSummary = summarizeDetailSheet_(detailSheet, subActivity);
  const fallbackDate = parseDateCell_(metadata.tanggal);
  const startDate = detailSummary.startDate || fallbackDate;
  const endDate = detailSummary.endDate || fallbackDate;

  return {
    nomorBerkas: subActivity.sort_order || '',
    kodeKlasifikasi: subActivity.default_kode_klasifikasi || '',
    uraian: getSubActivityFormalName_(subActivity) || metadata.uraian_informasi_berkas || '',
    kurunWaktu: formatDateRange_(startDate, endDate),
    jumlah: detailSummary.count ? detailSummary.count + ' dokumen' : '',
    filingCabinet: metadata.no_filing_cabinet || ConfigService.getSettings().defaultFilingCabinet || '02',
    noLaci: metadata.no_laci || activity.laci_no || '',
    noFolder: metadata.no_folder || subActivity.no_folder || activity.folder_no || '',
    akses: detailSummary.akses || metadata.klasifikasi_akses || 'Terbatas',
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

function writeRekapIdentity_(sheet, rowIndex, activity, subActivity) {
  const headerMap = getRekapHeaderMap_(sheet);
  const noFolder = subActivity && subActivity.no_folder
    ? subActivity.no_folder
    : (activity && activity.folder_no) || '';
  const identity = {
    nomorBerkas: subActivity && subActivity.sort_order ? subActivity.sort_order : '',
    kodeKlasifikasi: subActivity && subActivity.default_kode_klasifikasi ? subActivity.default_kode_klasifikasi : '',
    uraian: getSubActivityFormalName_(subActivity),
    filingCabinet: '02',
    noLaci: activity && activity.laci_no ? activity.laci_no : '',
    noFolder: noFolder,
    akses: 'Terbatas'
  };

  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas, identity.nomorBerkas, false);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kodeKlasifikasi, identity.kodeKlasifikasi, true);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.uraian, identity.uraian, true);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet, identity.filingCabinet, false);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noLaci, identity.noLaci, true);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noFolder, identity.noFolder, true);
  setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, identity.akses, false);
  sheet.getRange(rowIndex, REKAP_FALLBACK_START_COL, 1, Math.max(findLastRekapHeaderColumn_(sheet) - REKAP_FALLBACK_START_COL + 1, 1))
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle');
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
    return { count: 0, startDate: null, endDate: null, filingCabinet: '', noLaci: '', noFolder: '', akses: '', lastRow: null };
  }

  const values = data.values;
  const columnMap = data.columnMap;

  let lastRow = null;

  values.forEach(function (row) {
    const metadata = {};
    DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
      const col = columnMap[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
      const val = row[col - 1];
      metadata[field] = (val === null || val === undefined) ? '' : String(val).trim();
    });
    if (!hasMeaningfulDetailData_(metadata)) return;
    lastRow = metadata;
  });

  const dateRange = calculateDetailKurunWaktuDatesFromValues_(values, columnMap);

  return {
    count: calculateDetailJumlahFromValues_(values, columnMap),
    startDate: dateRange.startDate,
    endDate: dateRange.endDate,
    akses: calculateDetailAksesFromValues_(values, columnMap),
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

function calculateDetailJumlahFromValues_(values, columnMap) {
  let count = 0;
  values.forEach(function (row) {
    const metadata = {};
    DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
      const col = columnMap[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
      const val = row[col - 1];
      metadata[field] = (val === null || val === undefined) ? '' : String(val).trim();
    });
    if (hasMeaningfulDetailData_(metadata)) count++;
  });
  return count;
}

function calculateDetailKurunWaktuDatesFromValues_(values, columnMap) {
  let startDate = null;
  let endDate = null;
  values.forEach(function (row) {
    const metadata = {};
    DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
      const col = columnMap[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
      const val = row[col - 1];
      metadata[field] = (val === null || val === undefined) ? '' : String(val).trim();
    });
    if (!hasMeaningfulDetailData_(metadata)) return;
    const dateCol = columnMap.tanggal || (DETAIL_FALLBACK_START_COL + 4);
    const dateValue = parseDateCell_(row[dateCol - 1]);
    if (dateValue) {
      if (!startDate || dateValue.getTime() < startDate.getTime()) startDate = dateValue;
      if (!endDate || dateValue.getTime() > endDate.getTime()) endDate = dateValue;
    }
  });
  return { startDate: startDate, endDate: endDate };
}

function calculateDetailAksesFromValues_(values, columnMap) {
  const aksesByKey = {};
  values.forEach(function (row) {
    const metadata = {};
    DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
      const col = columnMap[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
      const val = row[col - 1];
      metadata[field] = (val === null || val === undefined) ? '' : String(val).trim();
    });
    if (!hasMeaningfulDetailData_(metadata)) return;
    const akses = normalizeAccessSummaryValue_(metadata.klasifikasi_akses);
    if (akses) aksesByKey[normalizeLooseLabel_(akses)] = akses;
  });
  return formatAccessSummary_(aksesByKey);
}

function normalizeAccessSummaryValue_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const key = normalizeLooseLabel_(text);
  if (key.indexOf('terbuka') >= 0) return 'Terbuka';
  if (key.indexOf('biasa') >= 0) return 'Biasa';
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
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  const text = String(value).trim();
  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
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

  // Jangan pernah pakai sheet Rekap sebagai template Detail — bikin layout Detail berantakan.
  const template = ss.getSheetByName('Template Detail Item') ||
    ss.getSheetByName(DEFAULT_DETAIL_SHEET_NAME) ||
    null;

  if (template && template.getName() !== REKAP_SHEET_NAME) {
    sheet = template.copyTo(ss).setName(targetName);
    clearDetailRows_(sheet);
    const clearNoteRow = findNoteRow_(sheet) || DETAIL_NOTE_FALLBACK_ROW;
    const clearDataRows = Math.max(clearNoteRow - DETAIL_DATA_START_ROW, 24);
    sheet.getRange(DETAIL_DATA_START_ROW, getDetailStartColumn_(sheet) + DETAIL_ITEM_NUMBER_OFFSET, clearDataRows, 1)
      .setNumberFormat('00');
  } else {
    sheet = ss.insertSheet(targetName);
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
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('B4').setValue('Unit pengolah : Latbang');
  sheet.getRange('B4')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold');

  sheet.getRange('B6:B7').merge().setValue('No\nBerkas');
  sheet.getRange('C6:C7').merge().setValue('Nomor Item\nArsip');
  sheet.getRange('D6:D7').merge().setValue('Kode\nKlasifikasi');
  sheet.getRange('E6:E7').merge().setValue('Uraian informasi Berkas');
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

  sheet.getRange('B6:M8')
    .setBackground('#b4c6e7')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true);

  const dataRows = 24;
  sheet.getRange(DETAIL_DATA_START_ROW, DETAIL_FALLBACK_START_COL, dataRows, DETAIL_FIELD_ORDER.length)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);
  const numberRows = [];
  for (let i = 1; i <= dataRows; i++) numberRows.push(['']);
  sheet.getRange(DETAIL_DATA_START_ROW, DETAIL_FALLBACK_START_COL, dataRows, 1)
    .setValues(numberRows)
    .setHorizontalAlignment('center');
  sheet.getRange(DETAIL_DATA_START_ROW, DETAIL_FALLBACK_START_COL + DETAIL_ITEM_NUMBER_OFFSET, dataRows, 1)
    .setNumberFormat('00')
    .clearContent();

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
    ['Kolom (9), diisi dengan nomor laci pada Filing Cabinet;'],
    ['Kolom (10), diisi dengan nomor folder Arsip;'],
    ['Kolom (11), diisi dengan klasifikasi keamanan seperti terbuka, terbatas, dan rahasia.'],
    ['Kolom (12), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik, dan digital.']
  ];
  sheet.getRange(startRow, 3, notes.length, 1).setValues(notes);
  sheet.getRange(startRow, 3, notes.length, 1)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold');
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

function nextNumberInColumn_(sheet, startRow, endRow, col) {
  if (endRow < startRow) return 1;
  const values = sheet.getRange(startRow, col, endRow - startRow + 1, 1).getDisplayValues();
  let max = 0;
  values.forEach(row => {
    const parsed = Number(row[0]);
    if (!isNaN(parsed) && parsed > max) max = parsed;
  });
  return max + 1;
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
    uraian_informasi_berkas: ['uraian informasi berkas', 'uraian informasi arsip', 'uraian informasi'],
    tanggal: ['tgl', 'tanggal'],
    tingkat_perkembangan: ['tingkat perkembangan', 'tingkat pengembangan'],
    jumlah: ['jumlah'],
    no_filing_cabinet: ['no filing cabinet'],
    no_laci: ['no laci'],
    no_folder: ['no folder'],
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
    metadata.uraian_informasi_berkas ||
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
  var letter = '';
  while (col > 0) {
    col--;
    letter = String.fromCharCode(65 + (col % 26)) + letter;
    col = Math.floor(col / 26);
  }
  return letter;
}

function detailSheetFormulaRef_(detailSheet) {
  var name = detailSheet.getName();
  var escaped = name.replace(/'/g, "''");
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
// desimal-koma) pakai ';', US/desimal-titik pakai ','. setFormula di locale
// ';' PARSE-ERROR kalau dikasih koma. Satu-satunya cara lintas-locale: tanya
// ke sheet-nya sendiri.
//
// Probe '=1/2' (tanpa argumen -> ga mungkin parse-error) ungkap pemisah
// desimal lewat display "0,5" vs "0.5". Aturan Sheets tak terkecuali:
// desimal-koma => argumen ';', desimal-titik => ','.
//
// Hasil di-cache di Document Properties, di-key pakai locale; auto re-probe
// kalau user ganti locale spreadsheet. Plus memo per-eksekusi.
var FORMULA_SEP_MEMO_ = null;
function formulaSep_(ss) {
  if (FORMULA_SEP_MEMO_) return FORMULA_SEP_MEMO_;
  try {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    var locale = '';
    try { locale = ss.getSpreadsheetLocale() || ''; } catch (e) {}
    var propKey = 'FORMULA_SEP::' + locale;
    var props = PropertiesService.getDocumentProperties();
    var cached = props.getProperty(propKey);
    if (cached === ',' || cached === ';') return (FORMULA_SEP_MEMO_ = cached);

    var sheet = ss.getSheets()[0];
    var cell = sheet.getRange(sheet.getMaxRows(), sheet.getMaxColumns());
    var prevFormula = cell.getFormula();
    var prevValue = prevFormula ? null : cell.getValue();
    cell.setFormula('=1/2');
    SpreadsheetApp.flush();
    var disp = String(cell.getDisplayValue());
    if (prevFormula) cell.setFormula(prevFormula);
    else cell.setValue(prevValue);

    FORMULA_SEP_MEMO_ = disp.indexOf(',') >= 0 ? ';' : ',';
    try { props.setProperty(propKey, FORMULA_SEP_MEMO_); } catch (e) {}
  } catch (e) {
    FORMULA_SEP_MEMO_ = ',';
  }
  return FORMULA_SEP_MEMO_;
}

function buildKurunWaktuFormula_(detailSheet) {
  const ref = detailSheetFormulaRef_(detailSheet);
  const colLetter = getDetailColumnLetter_(detailSheet, 'tanggal', 'F');
  const rangeRef = ref + '!' + colLetter + '9:' + colLetter;
  const s = formulaSep_(detailSheet.getParent());
  return "=IF(COUNTA(" + rangeRef + ")>0" + s + " TEXT(MIN(" + rangeRef + ")" + s + " \"d mmmm yyyy\") & \" - \" & TEXT(MAX(" + rangeRef + ")" + s + " \"d mmmm yyyy\")" + s + " \"\")";
}

function buildJumlahFormula_(detailSheet) {
  const ref = detailSheetFormulaRef_(detailSheet);
  const colLetter = getDetailColumnLetter_(detailSheet, 'nomor_item_arsip', 'C');
  const bLetter = getDetailColumnLetter_(detailSheet, 'no_berkas', 'B');
  const rangeRef = ref + '!' + colLetter + '9:' + colLetter;
  const bRange = ref + '!' + bLetter + '9:' + bLetter;
  const s = formulaSep_(detailSheet.getParent());
  return "=COUNTIFS(" + rangeRef + s + " \"<>\"" + s + " " +
         bRange + s + " \"<>Keterangan*\"" + s + " " + bRange + s + " \"<>Kolom*\"" + s + " " +
         rangeRef + s + " \"<>Keterangan*\"" + s + " " + rangeRef + s + " \"<>Kolom*\") & \" dokumen\"";
}

function buildAksesFormula_(detailSheet) {
  const ref = detailSheetFormulaRef_(detailSheet);
  const colLetter = getDetailColumnLetter_(detailSheet, 'klasifikasi_akses', 'M');
  const rangeRef = ref + '!' + colLetter + '9:' + colLetter;
  const s = formulaSep_(detailSheet.getParent());
  return "=IFERROR(TEXTJOIN(\" & \"" + s + " TRUE" + s + " UNIQUE(FILTER(" + rangeRef + s + " " + rangeRef + "<>\"\")))" + s + " \"Terbatas\")";
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

function sortDetailSheetByNoBerkas_(sheet) {
  try {
    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const lastDataRow = noteRow - 1;
    if (lastDataRow < DETAIL_DATA_START_ROW) return;
    const startCol = getDetailStartColumn_(sheet);
    const lastCol = Math.max(sheet.getLastColumn(), startCol + DETAIL_FIELD_ORDER.length - 1);
    const numRows = lastDataRow - DETAIL_DATA_START_ROW + 1;
    const numCols = lastCol - startCol + 1;
    coerceColumnNumeric_(sheet, DETAIL_DATA_START_ROW, startCol, numRows); // No Berkas
    const range = sheet.getRange(DETAIL_DATA_START_ROW, startCol, numRows, numCols);
    // Use native Sheets sort to preserve formulas and formatting
    range.sort({column: startCol, ascending: true});
    sheet.getRange(DETAIL_DATA_START_ROW, startCol + DETAIL_ITEM_NUMBER_OFFSET, numRows, 1).setNumberFormat('00');
  } catch (e) {
    console.warn('sortDetailSheetByNoBerkas_: ' + e.message);
  }
}
