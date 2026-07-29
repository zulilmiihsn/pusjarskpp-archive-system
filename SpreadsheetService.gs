'use strict';

// Konstanta layout sheet Detail & Rekap dipindah ke ConfigConstants.gs (dipakai lintas
// SpreadsheetService.gs + SheetHelpers.gs supaya tidak ada coupling const antar dua file).

function sanitizeCellValue_(value) {
  if (value === null || value === undefined) return '';
  const str = String(value);
  if (/^[=+\-@]/.test(str)) return "'" + str;
  return str;
}

function buildDetailRowValues_(metadata) {
  return DETAIL_FIELD_ORDER.map(function (key) {
    if (key === 'uraian_informasi_item') {
      const parts = [];
      const nomor = String(metadata.nomor_surat || '').trim();
      const uraian = String(metadata.uraian_informasi_item || '').trim();
      const kepada = String(metadata.kepada || '').trim();
      const dari = String(metadata.dari || '').trim();
      
      if (nomor) parts.push('No. Surat: ' + nomor);
      if (uraian) parts.push('Uraian: ' + uraian);
      if (kepada) parts.push('Kepada: ' + kepada);
      if (dari) parts.push('Dari: ' + dari);
      
      return sanitizeCellValue_(parts.join('\n'));
    }
    if (key === 'lokasi_simpan' && metadata._lokasi_simpan_url) {
      return sanitizeCellValue_(metadata._lokasi_simpan_url);
    }
    if (key === 'no_filing_cabinet' && metadata._no_filing_cabinet_path) {
      return sanitizeCellValue_(metadata._no_filing_cabinet_path);
    }
    if (key === 'no_filing_cabinet' && !metadata._no_filing_cabinet_path) {
      return ''; // Tanpa path/hyperlink, kosongin aja drpd nulis ID folder mentah
    }
    if (key === 'jumlah_satuan') {
      return sanitizeCellValue_(metadata.satuan || metadata.jumlah_satuan || '');
    }
    return sanitizeCellValue_(metadata[key]);
  });
}

// Bangun satu sel detail arsip: nilai + format (border SOLID, wrap, vertical middle),
// konversi tanggal -> formula =DATE() + format DATE, format '00' utk nomor item, dan
// link sel bila ada *_url. Dipakai append/bulk/update agar formatnya konsisten.
function buildDetailCellData_(fieldName, val, metadata) {
  const isItem = fieldName === 'nomor_item_arsip';
  const isNumericField = fieldName === 'jumlah' || isItem;
  const numericVal = Number(val);
  let userEnteredValue;
  let numberFormat;

  if (fieldName === 'tanggal' && val) {
    const d = parseDateCell_(val);
    if (d) {
      const serialDate = 25569 + (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000);
      userEnteredValue = { numberValue: serialDate };
      numberFormat = { type: 'DATE', pattern: 'd mmmm yyyy' };
    } else {
      userEnteredValue = { stringValue: String(val) };
    }
  } else if (isNumericField && !isNaN(numericVal) && String(val).trim() !== '') {
    userEnteredValue = { numberValue: numericVal };
  } else {
    userEnteredValue = { stringValue: String(val) };
  }

  const cellData = {
    userEnteredValue: userEnteredValue,
    userEnteredFormat: {
      wrapStrategy: 'WRAP',
      verticalAlignment: 'MIDDLE',
      borders: {
        top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
        bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
        left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
        right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } }
      }
    }
  };

  if (numberFormat) {
    cellData.userEnteredFormat.numberFormat = numberFormat;
  }
  if (isItem) {
    cellData.userEnteredFormat.numberFormat = { type: 'NUMBER', pattern: '00' };
  }

  if (fieldName === 'lokasi_simpan' && metadata._lokasi_simpan_url) {
    cellData.userEnteredFormat.textFormat = { link: { uri: metadata._lokasi_simpan_url } };
  }
  if (fieldName === 'no_filing_cabinet' && metadata._no_filing_cabinet_url) {
    cellData.userEnteredFormat.textFormat = { link: { uri: metadata._no_filing_cabinet_url } };
  }
  if (fieldName === 'no_laci' && metadata._no_laci_url) {
    cellData.userEnteredFormat.textFormat = { link: { uri: metadata._no_laci_url } };
  }
  if (fieldName === 'no_folder' && metadata._no_folder_url) {
    cellData.userEnteredFormat.textFormat = { link: { uri: metadata._no_folder_url } };
  }

  return cellData;
}

function extractMetadataFromRow_(rowValues) {
  const metadata = {};
  DETAIL_FIELD_ORDER.forEach(function (key, i) {
    let val = rowValues[i];
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    metadata[key] = String(val || '');
  });
  if (metadata.jumlah_satuan && !metadata.satuan) {
    metadata.satuan = metadata.jumlah_satuan;
  }
  return metadata;
}

// Memo handle Spreadsheet per-eksekusi. openById adalah RPC penuh, dan satu request
// kerap membuka spreadsheet yang sama berkali-kali (mis. getNextItemNumber ->
// appendArchiveRow -> updateRekapSummary -> log). State modul GAS direset antar-eksekusi
// sehingga memo ini aman & tak pernah basi lintas-request. Di-key per id agar bila id
// berubah dalam satu request (mis. ganti workspace) handle baru tetap dibuka.
const _ssHandleCache_ = {};
function openSpreadsheetById_(id) {
  const key = String(id);
  let ss = _ssHandleCache_[key];
  if (!ss) {
    ss = SpreadsheetApp.openById(id);
    _ssHandleCache_[key] = ss;
  }
  return ss;
}

// Penulis sel detail terpadu (dipakai append/bulk/update). Membangun nilai+format
// via batchUpdate updateCells untuk SEMUA baris sekaligus, lalu menerapkan RichText
// (teks display = nama folder, link = URL) pada kolom FC/laci/folder di TIAP baris —
// supaya jalur update/bulk tidak lagi kehilangan teks display seperti dulu (B6).
function writeDetailRows_(sheet, startRow, startCol, metadataList) {
  const rowsData = metadataList.map(function (meta) {
    const rowVals = buildDetailRowValues_(meta);
    const values = [];
    for (let c = 0; c < DETAIL_FIELD_ORDER.length; c++) {
      values.push(buildDetailCellData_(DETAIL_FIELD_ORDER[c], rowVals[c] || '', meta));
    }
    return { values: values };
  });

  Sheets.Spreadsheets.batchUpdate({
    requests: [{
      updateCells: {
        rows: rowsData,
        fields: "userEnteredValue,userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment,userEnteredFormat.borders,userEnteredFormat.numberFormat,userEnteredFormat.textFormat",
        range: {
          sheetId: sheet.getSheetId(),
          startRowIndex: startRow - 1,
          endRowIndex: startRow - 1 + metadataList.length,
          startColumnIndex: startCol - 1,
          endColumnIndex: startCol - 1 + DETAIL_FIELD_ORDER.length
        }
      }
    }]
  }, sheet.getParent().getId());

  metadataList.forEach(function (meta, i) {
    applyDetailRichTextLinks_(sheet, startRow + i, startCol, meta);
  });
}

// RichText (teks berbeda dari URL) untuk kolom hyperlink agar nama folder tampil,
// bukan ID mentah. Dipanggil per baris oleh writeDetailRows_.
function applyDetailRichTextLinks_(sheet, rowIndex, startCol, metadata) {
  if (!SpreadsheetApp.newRichTextValue) return;
  const LINK_FIELDS = [
    ['no_filing_cabinet', '_no_filing_cabinet_path', '_no_filing_cabinet_url'],
    ['no_laci', '_no_laci_path', '_no_laci_url'],
    ['no_folder', '_no_folder_path', '_no_folder_url']
  ];
  LINK_FIELDS.forEach(function (spec) {
    const idx = DETAIL_FIELD_ORDER.indexOf(spec[0]);
    const path = metadata[spec[1]];
    const url = metadata[spec[2]];
    if (idx >= 0 && path && url) {
      sheet.getRange(rowIndex, startCol + idx).setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText(path).setLinkUrl(url).build()
      );
    }
  });
}

const SpreadsheetService = {
  getNextItemNumber: function (activity, subActivity) {
    try {
      const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
      const sheet = findExistingDetailSheet_(ss, subActivity);
      if (!sheet) return '01';

      const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
      const dataStart = DETAIL_DATA_START_ROW;
      if (noteRow <= dataStart) return '01';
      const rowCount = noteRow - dataStart;

      const width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
      const headerColumns = getDetailColumnMap_(sheet, width);
      const itemCol = headerColumns.nomor_item_arsip || (DETAIL_FALLBACK_START_COL + DETAIL_ITEM_NUMBER_OFFSET);

      const values = sheet.getRange(dataStart, itemCol, rowCount, 1).getDisplayValues();
      let maxNum = 0;
      values.forEach(function (row) {
        const raw = String(row[0] || '').trim();
        if (raw) {
          const num = Number(raw);
          if (!isNaN(num) && num > maxNum) maxNum = num;
        }
      });
      return String(maxNum + 1).padStart(2, '0');
    } catch (e) {
      return '01';
    }
  },

  appendArchiveRow: function (activity, subActivity, metadata) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const rowIndex = findWritableDetailRow_(sheet);
    const startCol = getDetailStartColumn_(sheet);

    writeDetailRows_(sheet, rowIndex, startCol, [metadata]);
    sortDetailSheetByNomorItemArsip_(sheet);

    // Sort mengubah urutan baris; cari ulang posisi baris via URL file agar
    // spreadsheet_row_number yang dicatat tidak basi (edit/hapus bisa salah baris).
    let finalRow = rowIndex;
    if (metadata._lokasi_simpan_url) {
      const located = locateDetailRowByUrl_(sheet, metadata._lokasi_simpan_url);
      if (located) finalRow = located;
    }

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: finalRow
    };
  },

  appendArchiveRowsBulk: function (activity, subActivity, metadataList) {
    if (!metadataList || metadataList.length === 0) return [];
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const startCol = getDetailStartColumn_(sheet);
    const checkCol = startCol + DETAIL_WRITABLE_CHECK_OFFSET;
    
    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    let lastFilledRow = DETAIL_DATA_START_ROW - 1;
    
    const numCheckRows = noteRow - DETAIL_DATA_START_ROW;
    if (numCheckRows > 0) {
       const checkValues = sheet.getRange(DETAIL_DATA_START_ROW, checkCol, numCheckRows, 1).getValues();
       for (let i = numCheckRows - 1; i >= 0; i--) {
          if (String(checkValues[i][0]).trim() !== '') {
             lastFilledRow = DETAIL_DATA_START_ROW + i;
             break;
          }
       }
    }
    
    const rowIndex = lastFilledRow + 1;
    const blanksAvailable = noteRow - rowIndex;
    const numRows = metadataList.length;
    
    if (blanksAvailable < numRows) {
       sheet.insertRowsBefore(noteRow, numRows - blanksAvailable);
       invalidateNoteRowCache_(sheet);
    }

    writeDetailRows_(sheet, rowIndex, startCol, metadataList);
    sortDetailSheetByNomorItemArsip_(sheet);

    return metadataList.map(function(meta, i) {
      let finalRow = rowIndex + i;
      if (meta._lokasi_simpan_url) {
        const located = locateDetailRowByUrl_(sheet, meta._lokasi_simpan_url);
        if (located) finalRow = located;
      }
      return {
        spreadsheetId: ss.getId(),
        spreadsheetUrl: ss.getUrl(),
        sheetName: sheet.getName(),
        rowNumber: finalRow
      };
    });
  },

  getArchiveRow: function (activity, subActivity, rowIndex) {
    if (!rowIndex) return {};
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const startCol = getDetailStartColumn_(sheet);
    const rowValues = sheet.getRange(rowIndex, startCol, 1, DETAIL_FIELD_ORDER.length).getValues()[0];
    const metadata = extractMetadataFromRow_(rowValues);
    return metadata;
  },

  getArchiveRowByFileId: function (activity, subActivity, fileId, fallbackRowIndex) {
    if (!fileId && !fallbackRowIndex) return {};
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const startCol = getDetailStartColumn_(sheet);
    const lokasiSimpanIdx = DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
    let rowIndex = fallbackRowIndex;

    if (fileId && lokasiSimpanIdx >= 0) {
      const dataRows = Math.max(0, sheet.getLastRow() - DETAIL_DATA_START_ROW + 1);
      if (dataRows > 0) {
        const linkRange = sheet.getRange(DETAIL_DATA_START_ROW, startCol + lokasiSimpanIdx, dataRows, 1);
        const richTextValues = linkRange.getRichTextValues();
        for (let i = 0; i < richTextValues.length; i++) {
          const cellRichText = richTextValues[i][0];
          const url = cellRichText ? (cellRichText.getLinkUrl() || cellRichText.getText() || '') : '';
          if (url.indexOf(fileId) !== -1) {
            rowIndex = DETAIL_DATA_START_ROW + i;
            break;
          }
        }
      }
    }

    if (!rowIndex) return {};

    const rowValues = sheet.getRange(rowIndex, startCol, 1, DETAIL_FIELD_ORDER.length).getValues()[0];
    const metadata = extractMetadataFromRow_(rowValues);
    metadata._resolved_row_number = rowIndex;
    return metadata;
  },

  updateArchiveRow: function (activity, subActivity, rowIndex, metadata) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const startCol = getDetailStartColumn_(sheet);

    writeDetailRows_(sheet, rowIndex, startCol, [metadata]);
    sortDetailSheetByNomorItemArsip_(sheet);

    // Sort mengubah urutan baris; cari ulang posisi baris via URL file agar
    // spreadsheet_row_number yang dicatat tidak basi (edit/hapus bisa salah baris).
    let finalRow = rowIndex;
    if (metadata._lokasi_simpan_url) {
      const located = locateDetailRowByUrl_(sheet, metadata._lokasi_simpan_url);
      if (located) finalRow = located;
    }

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: finalRow
    };
  },

  ensureSubActivitySheet: function (activity, subActivity) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName()
    };
  },

  getDetailMetadataDefaults: function (activity, subActivity, forceCalculate) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));

    let locks = {};
    try {
      locks = subActivity && subActivity.metadata_locks ? JSON.parse(subActivity.metadata_locks) : {};
    } catch(e) { console.warn('getDetailMetadataDefaults: failed to parse metadata_locks: ' + e.message); }

    const needSummary = forceCalculate || (locks.kurunWaktu === false) || (locks.jumlah === false) || (locks.akses === false);

    const sheet = findExistingDetailSheet_(ss, subActivity);
    const summary = (sheet && needSummary)
      ? summarizeDetailSheet_(sheet, subActivity)
      : { count: 0, startDate: null, endDate: null, akses: '' };

    let rekapFC = '', rekapLaci = '', rekapFolder = '', rekapAkses = '';
    let rekapNomorBerkas = '', rekapKurunWaktu = '', rekapJumlah = '', rekapKet = '';
    let rekapSheet, rowIndex, docCols = [];
    try {
      rekapSheet = findRekapSheet_(ss);
      if (rekapSheet) {
        rowIndex = findRekapRowForSubActivity_(rekapSheet, subActivity);
        if (rowIndex) {
          const headerMap = getRekapHeaderMap_(rekapSheet);
          // Resolusi kolom URL untuk SEMUA tipe dokumen aktif (data-driven).
          docCols = getRekapDocColumns_().map(function (c) {
            return { key: c.key, col: findRekapHeaderColumnFromMap_(headerMap, c.match) };
          });
          const kodeCol       = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.kodeKlasifikasi);
          const fcCol         = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet);
          const laciCol       = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noLaci);
          const folderCol     = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noFolder);
          const aksesCol      = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.akses);
          const nomorBerkasCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas);
          const kurunWaktuCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu);
          const jumlahCol     = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.jumlah);
          const ketCol        = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.ket);

          // Batch-read entire rekap row in one call
          const docColNums = docCols.map(function (d) { return d.col; });
          const allCols = docColNums.concat([kodeCol, fcCol, laciCol, folderCol, aksesCol, nomorBerkasCol, kurunWaktuCol, jumlahCol, ketCol]).filter(Boolean);
          const maxCol = Math.max.apply(null, allCols);
          const rowValues = rekapSheet.getRange(rowIndex, 1, 1, maxCol).getDisplayValues()[0];
          
          const cell = function (col) { return col ? String(rowValues[col - 1] || '').trim() : ''; };

          rekapFC          = cell(fcCol);
          rekapLaci        = cell(laciCol);
          rekapFolder      = cell(folderCol);
          rekapAkses       = cell(aksesCol);
          rekapNomorBerkas = cell(nomorBerkasCol);
          rekapKurunWaktu  = cell(kurunWaktuCol);
          rekapJumlah      = cell(jumlahCol);
          rekapKet         = cell(ketCol);
        }
      }
    } catch (e) {
      console.warn('Failed to get rekap metadata: ' + e.message);
    }

    const showJsComputed = forceCalculate;
    const nextItemNum = SpreadsheetService.getNextItemNumber(activity, subActivity);
    const meta = {
      no_berkas: rekapNomorBerkas || resolveSubActivityArchiveNumber_(subActivity, activity),
      kode_klasifikasi: '',
      kurun_waktu: (showJsComputed ? '' : rekapKurunWaktu) || formatDateRange_(summary.startDate, summary.endDate),
      jumlah: (showJsComputed ? '' : rekapJumlah) || ((summary.sumLembar || 0) + ' lembar'),
      no_filing_cabinet: (showJsComputed ? '' : rekapFC) || (activity && activity.laci_no ? String(activity.laci_no).padStart(2, '0') + '. Laci ' + activity.activity_name : '') || '',
      _no_filing_cabinet_path: activity && activity.laci_no ? String(activity.laci_no).padStart(2, '0') + '. Laci ' + activity.activity_name : '',
      _no_filing_cabinet_url: (activity && activity.laci_folder_id) ? 'https://drive.google.com/drive/folders/' + activity.laci_folder_id : '',
      
      no_laci: (showJsComputed ? '' : rekapLaci) || (subActivity && subActivity.sub_activity_name) || '',
      _no_laci_path: (subActivity && subActivity.sub_activity_name) || '',
      _no_laci_url: (subActivity && subActivity.folder_id) ? 'https://drive.google.com/drive/folders/' + subActivity.folder_id : '',
      
      no_folder: (showJsComputed ? '' : rekapFolder) || (subActivity && subActivity.no_folder) || String(nextItemNum).padStart(2, '0'),
      klasifikasi_akses: (showJsComputed ? '' : rekapAkses) || summary.akses || 'Terbatas',
      ket: rekapKet || '',
      next_item_number: nextItemNum
    };

    if (rekapSheet && rowIndex) {
      // Batch-read formulas for URL columns in one call
      const urlCols = docCols.map(function (d) { return d.col; }).filter(Boolean);
      if (urlCols.length) {
        const urlMaxCol = Math.max.apply(null, urlCols);
        const formulas = rekapSheet.getRange(rowIndex, 1, 1, urlMaxCol).getFormulas()[0];
        const extractUrl = function (col) {
          if (!col) return '';
          const formula = formulas[col - 1] || '';
          const match = formula.match(/=HYPERLINK\(\s*["']([^"']+)["']/i);
          return match && match[1] ? match[1] : '';
        };
        // Emit per tipe dokumen: meta['url_' + key]
        docCols.forEach(function (d) {
          if (d.col) meta['url_' + d.key] = extractUrl(d.col);
        });
      }
    }

    return meta;
  },

  listExistingItemNumbers: function (activity, subActivity) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findExistingDetailSheet_(ss, subActivity);
    if (!sheet) return [];

    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    if (noteRow <= DETAIL_DATA_START_ROW) return [];

    const width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
    const headerColumns = getDetailColumnMap_(sheet, width);
    const itemCol = headerColumns.nomor_item_arsip || (DETAIL_FALLBACK_START_COL + 1);

    const rowCount = noteRow - DETAIL_DATA_START_ROW;
    const itemValues = sheet.getRange(DETAIL_DATA_START_ROW, itemCol, rowCount, 1).getDisplayValues();

    const pairs = [];
    for (let i = 0; i < rowCount; i++) {
      const item = String(itemValues[i][0] || '').trim();
      if (!item) continue;
      pairs.push({
        rowNumber: DETAIL_DATA_START_ROW + i,
        nomor_item_arsip: item
      });
    }
    return pairs;
  },

  listExistingArchiveRows: function (activity, subActivity) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findExistingDetailSheet_(ss, subActivity);
    if (!sheet) {
      return {
        spreadsheetId: ss.getId(),
        spreadsheetUrl: ss.getUrl(),
        sheetName: '',
        rows: [],
        missingSheet: true
      };
    }

    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const startRow = DETAIL_DATA_START_ROW;
    if (noteRow <= startRow) {
      return {
        spreadsheetId: ss.getId(),
        spreadsheetUrl: ss.getUrl(),
        sheetName: sheet.getName(),
        rows: []
      };
    }

    const width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
    const headerColumns = getDetailColumnMap_(sheet, width);
    const rowCount = noteRow - startRow;
    const values = sheet.getRange(startRow, 1, rowCount, width).getDisplayValues();
    const rows = [];

    values.forEach(function (row, index) {
      const metadata = {};
      DETAIL_FIELD_ORDER.forEach(function (field, fallbackIndex) {
        const col = headerColumns[field] || (DETAIL_FALLBACK_START_COL + fallbackIndex);
        metadata[field] = sanitizeCellValue_(row[col - 1] || '');
      });
      if (!hasMeaningfulDetailData_(metadata)) return;

      rows.push({
        rowNumber: startRow + index,
        metadata: metadata
      });
    });

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rows: rows
    };
  },

  appendRekapRowIfPresent: function (activity, subActivity) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    ensureRekapDocumentColumns_(sheet);
    const existingRow = findRekapRowForSubActivity_(sheet, subActivity);
    if (existingRow) {
      writeRekapIdentityMarker_(sheet, existingRow, subActivity);
      return { sheetName: sheet.getName(), rowNumber: existingRow, existing: true };
    }

    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const rowIndex = findFirstBlankInRange_(sheet, REKAP_DATA_START_ROW, noteRow - 1, 4) || noteRow;
    if (rowIndex === noteRow) {
      sheet.insertRowBefore(noteRow);
      invalidateNoteRowCache_(sheet);
    }

    const nomorBerkas = getNextAvailableRekapArchiveNumber_(buildRekapRowLookup_(sheet));
    
    const rawFcText = activity.laci_no + '. Laci ' + activity.activity_name;
    const fcText = rawFcText.replace(/"/g, '""');
    const rawLaciText = subActivity.sub_activity_name || '';
    const laciText = rawLaciText.replace(/"/g, '""');

    const row = [
      sanitizeCellValue_(nomorBerkas),
      sanitizeCellValue_(''),
      sanitizeCellValue_(getSubActivityFormalName_(subActivity) || ''),
      '',
      '',
      sanitizeCellValue_(rawFcText),
      sanitizeCellValue_(rawLaciText),
      sanitizeCellValue_(nomorBerkas),
      'Terbatas',
      '',
      '',
      '',
      '',
      ''
    ];
    sheet.getRange(rowIndex, 2, 1, row.length).setValues([row]);
    
    const hsep = typeof formulaSep_ === 'function' ? formulaSep_(sheet.getParent()) : ';';
    if (activity.laci_folder_id) {
      sheet.getRange(rowIndex, 7).setFormula('=HYPERLINK("https://drive.google.com/drive/folders/' + activity.laci_folder_id + '"' + hsep + ' "' + fcText.replace(/"/g, '""').replace(/\n/g, '" & CHAR(10) & "') + '")');
    }
    if (subActivity.folder_id) {
      sheet.getRange(rowIndex, 8).setFormula('=HYPERLINK("https://drive.google.com/drive/folders/' + subActivity.folder_id + '"' + hsep + ' "' + laciText.replace(/"/g, '""').replace(/\n/g, '" & CHAR(10) & "') + '")');
    }
    writeRekapIdentityMarker_(sheet, rowIndex, subActivity);
    
    sheet.getRange(rowIndex, 2, 1, row.length)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment('middle');

    return { sheetName: sheet.getName(), rowNumber: rowIndex };
  },

  updateRekapSummary: function (activity, subActivity, metadata) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    ensureRekapDocumentColumns_(sheet);
    const rowIndex = findOrCreateRekapRow_(sheet, activity, subActivity);
    const summary = buildRekapSummary_(ss, activity, subActivity, metadata || {});
    
    let locks = {};
    try {
      locks = subActivity && subActivity.metadata_locks ? JSON.parse(subActivity.metadata_locks) : {};
    } catch(e) { console.warn('updateRekapSummary: failed to parse metadata_locks: ' + e.message); }

    const detailSheet = ensureDetailSheet_(ss, activity, subActivity);
    const headerMap = getRekapHeaderMap_(sheet);
    const archiveNumber = getRekapArchiveNumber_(
      sheet,
      rowIndex,
      resolveSubActivityArchiveNumber_(subActivity, activity)
    );
    summary.nomorBerkas = archiveNumber;
    summary.noFolder = archiveNumber;

    setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas, archiveNumber);

    if (locks.kurunWaktu !== false) {
      setRekapFormulaCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu, buildKurunWaktuFormula_(detailSheet));
      delete summary.kurunWaktu;
    } else {
      setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu, summary.kurunWaktu);
      delete summary.kurunWaktu;
    }

    if (locks.jumlah !== false) {
      setRekapFormulaCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.jumlah, buildJumlahFormula_(detailSheet));
      delete summary.jumlah;
    } else {
      setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.jumlah, summary.jumlah);
      delete summary.jumlah;
    }

    if (locks.akses !== false) {
      setRekapFormulaCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, buildAksesFormula_(detailSheet));
      delete summary.akses;
    } else {
      setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, summary.akses);
      delete summary.akses;
    }

    // Filing cabinet: hyperlink nama laci -> folder Drive (mode manual dibiarkan apa adanya).
    if (locks.filingCabinet !== false) {
      setRekapFilingCabinetCell_(sheet, rowIndex, headerMap, activity, true);
    }
    delete summary.filingCabinet;

    writeRekapSummary_(sheet, rowIndex, summary);
    sortRekapSheetByNomorBerkas_(sheet);
    const resolvedRowIndex = findRekapRowForSubActivity_(sheet, subActivity) || rowIndex;

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: resolvedRowIndex
    };
  },

  deleteArchiveRowAndReorder: function(activity, subActivity, rowNumber, fileId) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    
    let targetRow = rowNumber;

    if (fileId) {
      const maxRows = sheet.getLastRow();
      let foundRow = -1;
      
      if (maxRows >= DETAIL_DATA_START_ROW) {
        const urlCol = getDetailStartColumn_(sheet) + DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
        if (targetRow >= DETAIL_DATA_START_ROW && targetRow <= maxRows) {
           const hintUrl = sheet.getRange(targetRow, urlCol).getFormula();
           if (hintUrl && hintUrl.indexOf(fileId) !== -1) {
             foundRow = targetRow;
           }
        }
        
        if (foundRow === -1) {
           const formulas = sheet.getRange(DETAIL_DATA_START_ROW, urlCol, maxRows - DETAIL_DATA_START_ROW + 1, 1).getFormulas();
           const values = sheet.getRange(DETAIL_DATA_START_ROW, urlCol, maxRows - DETAIL_DATA_START_ROW + 1, 1).getValues();
           for (let i = 0; i < formulas.length; i++) {
             const cellContent = (formulas[i][0] || values[i][0] || '').toString();
             if (cellContent.indexOf(fileId) !== -1) {
               foundRow = DETAIL_DATA_START_ROW + i;
               break;
             }
           }
        }
      }

      if (foundRow !== -1) {
        targetRow = foundRow;
      } else {
        return null;
      }
    }

    if (!targetRow || targetRow < DETAIL_DATA_START_ROW) return null;

    try {
      sheet.deleteRow(targetRow);
      invalidateNoteRowCache_(sheet);
    } catch (e) {
      return null;
    }

    // 2. Reorder nomor_item_arsip and no_berkas
    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const maxRows = Math.max(noteRow - DETAIL_DATA_START_ROW, 0);
    
    if (maxRows > 0) {
      const startCol = getDetailStartColumn_(sheet);
      const itemNumberCol = startCol + DETAIL_ITEM_NUMBER_OFFSET;
      const writableCol = startCol + DETAIL_WRITABLE_CHECK_OFFSET;
      
      const firstBlank = findFirstBlankInRange_(sheet, DETAIL_DATA_START_ROW, noteRow - 1, writableCol);
      const filledRowsCount = firstBlank ? (firstBlank - DETAIL_DATA_START_ROW) : maxRows;
      
      // The user requested that we do NOT reorder existing rows, 
      // because changing the Nomor Item Arsip will mismatch with their filename in Google Drive.
      // So we leave the filled rows completely intact (even if it creates gaps like 01, 03, 04).

      
      // Clear the numbers for any remaining blank rows up to the note row
      const blankRowsCount = maxRows - filledRowsCount;
      if (blankRowsCount > 0) {
        sheet.getRange(DETAIL_DATA_START_ROW + filledRowsCount, itemNumberCol, blankRowsCount, 1).clearContent();
        sheet.getRange(DETAIL_DATA_START_ROW + filledRowsCount, startCol, blankRowsCount, 1).clearContent();
      }
    }
    
    return targetRow;
  },

  updateRekapSubActivityIdentity: function (activity, previousSubActivity, nextSubActivity) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, nextSubActivity || previousSubActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    ensureRekapDocumentColumns_(sheet);
    const rowIndex = findRekapRowForSubActivity_(sheet, nextSubActivity) ||
      findRekapRowForSubActivity_(sheet, previousSubActivity);
    if (!rowIndex) return null;

    writeRekapIdentity_(sheet, rowIndex, activity, nextSubActivity || previousSubActivity);
    sortRekapSheetByNomorBerkas_(sheet);
    const resolvedRowIndex = findRekapRowForSubActivity_(sheet, nextSubActivity || previousSubActivity) || rowIndex;

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: resolvedRowIndex
    };
  },

  /**
   * Rekonsiliasi non-destruktif antara config sub-kegiatan dan baris rekap.
   * Tidak menghapus baris ambigu. Identitas yang bisa dibuktikan lewat marker,
   * folder_id, atau nama unik diperbaiki dan rekap_row_number disegarkan.
   */
  reconcileRekapSubActivities: function (activity, subActivities) {
    const report = {
      repairedIdentities: 0,
      repairedPointers: 0,
      createdRows: 0,
      duplicateStableRows: [],
      missingSubActivityIds: []
    };
    const rows = (subActivities || []).filter(function (sub) {
      return sub && sub.sub_activity_id;
    });
    if (!rows.length) return report;

    const grouped = {};
    rows.forEach(function (sub) {
      const spreadsheetId = String(sub.spreadsheet_file_id || activity.spreadsheet_file_id || '');
      if (!spreadsheetId) return;
      if (!grouped[spreadsheetId]) grouped[spreadsheetId] = [];
      grouped[spreadsheetId].push(sub);
    });

    Object.keys(grouped).forEach(function (spreadsheetId) {
      const ss = openSpreadsheetById_(spreadsheetId);
      const sheet = findRekapSheet_(ss);
      if (!sheet) return;
      ensureRekapDocumentColumns_(sheet);

      let lookup = buildRekapRowLookup_(sheet);
      grouped[spreadsheetId].forEach(function (sub) {
        const stableRows = findStableRekapRowsFromLookup_(lookup, sub);
        if (stableRows.length > 1) {
          report.duplicateStableRows.push({
            subActivityId: sub.sub_activity_id,
            rows: stableRows.slice()
          });
        }

        let rowIndex = findRekapRowForSubActivityFromLookup_(lookup, sub);
        if (!rowIndex) {
          const possibleRows = findPossibleRekapRowsFromLookup_(lookup, sub);
          if (possibleRows.length) {
            // Ada baris lama yang mungkin milik sub-kegiatan ini. Jangan menebak
            // atau membuat duplikat; laporkan agar diperiksa manual.
            report.duplicateStableRows.push({
              subActivityId: sub.sub_activity_id,
              rows: possibleRows,
              reason: 'ambiguous_candidates'
            });
            return;
          }

          // Tidak ada marker, folder, nama, atau nomor yang cocok. Baris benar-benar
          // hilang, jadi aman dibuat ulang lalu dihitung dari sheet rincian.
          const created = SpreadsheetService.appendRekapRowIfPresent(activity, sub);
          if (!created || !created.rowNumber) {
            report.missingSubActivityIds.push(sub.sub_activity_id);
            return;
          }
          report.createdRows++;
          const summaryResult = SpreadsheetService.updateRekapSummary(activity, sub, {});
          rowIndex = summaryResult && summaryResult.rowNumber
            ? summaryResult.rowNumber
            : created.rowNumber;
          lookup = buildRekapRowLookup_(sheet);
        }

        if (rekapIdentityNeedsRepair_(lookup, rowIndex, activity, sub)) {
          writeRekapIdentity_(sheet, rowIndex, activity, sub, { respectManualLocks: true });
          report.repairedIdentities++;
          lookup = buildRekapRowLookup_(sheet);
        }
      });

      if (report.repairedIdentities || report.createdRows) {
        sortRekapSheetByNomorBerkas_(sheet);
      }

      lookup = buildRekapRowLookup_(sheet);
      grouped[spreadsheetId].forEach(function (sub) {
        const actualRow = findRekapRowForSubActivityFromLookup_(lookup, sub);
        if (!actualRow) return;
        const rowMatches = Number(sub.rekap_row_number) === Number(actualRow);
        if (rowMatches) return;
        ConfigRepository.updateSubActivityMapping({
          year: sub.year || activity.year,
          activityId: sub.activity_id || activity.activity_id,
          subActivityId: sub.sub_activity_id,
          rekapRowNumber: actualRow
        });
        sub.rekap_row_number = actualRow;
        report.repairedPointers++;
      });
    });

    return report;
  },

  /**
   * Menyusun nomor arsip global lintas seluruh kegiatan pada satu tahun.
   * Urutan lokal dipulihkan dari detail/rekap hanya saat local_sort_order belum ada.
   */
  planGlobalArchiveNumbers: function (year) {
    const activities = ConfigRepository.getActivities(year);
    const activityById = {};
    activities.forEach(function (activity) {
      activityById[String(activity.activity_id || '')] = activity;
    });

    const contexts = {};
    const errors = [];
    function getContext_(spreadsheetId) {
      const id = String(spreadsheetId || '');
      if (!id) return null;
      if (contexts[id] !== undefined) return contexts[id];
      try {
        const ss = openSpreadsheetById_(id);
        const rekapSheet = findRekapSheet_(ss);
        contexts[id] = {
          ss: ss,
          rekapSheet: rekapSheet,
          rekapLookup: rekapSheet ? buildRekapRowLookup_(rekapSheet) : null
        };
      } catch (error) {
        contexts[id] = null;
        errors.push({ spreadsheetId: id, message: error.message });
      }
      return contexts[id];
    }

    const subActivities = ConfigRepository.getAllSubActivitiesForYear(year).map(function (source) {
      const subActivity = Object.assign({}, source);
      if (Number(subActivity.local_sort_order) > 0) return subActivity;

      const activity = activityById[String(subActivity.activity_id || '')];
      if (!activity) return subActivity;
      const spreadsheetId = subActivity.spreadsheet_file_id || activity.spreadsheet_file_id;
      const context = getContext_(spreadsheetId);
      if (!context) return subActivity;

      let detailNumber = '';
      let rekapNumber = '';
      let rekapFolderNumber = '';
      try {
        detailNumber = getDetailArchiveNumberForSubActivity_(context.ss, subActivity);
        if (context.rekapLookup) {
          const rowIndex = findRekapRowForSubActivityFromLookup_(context.rekapLookup, subActivity);
          if (rowIndex) {
            const offset = rowIndex - context.rekapLookup.startRow;
            const row = context.rekapLookup.values[offset] || [];
            rekapNumber = String(row[context.rekapLookup.nomorBerkasCol - 1] || '').trim();
            rekapFolderNumber = String(row[context.rekapLookup.noFolderCol - 1] || '').trim();
          }
        }
      } catch (error) {
        errors.push({
          spreadsheetId: String(spreadsheetId || ''),
          subActivityId: subActivity.sub_activity_id,
          message: error.message
        });
      }
      subActivity._observed_archive_number = detailNumber || rekapNumber || rekapFolderNumber || '';
      return subActivity;
    });

    const plan = buildGlobalArchiveNumberPlan_(activities, subActivities);
    plan.errors = errors;
    return plan;
  },

  /**
   * Menerapkan satu sumber nomor global ke Config, Rekap, dan semua baris Detail.
   * Rekap: Nomor Berkas = No Folder = nomor global.
   * Detail: Nomor Berkas = nomor global; No Folder tetap Nomor Item Arsip.
   */
  reconcileGlobalArchiveNumbers: function (year, options) {
    options = options || {};
    const applyChanges = options.apply !== false;
    const plan = this.planGlobalArchiveNumbers(year);
    const report = {
      year: Number(year),
      applied: applyChanges,
      totalActive: plan.totalActive,
      configCellsUpdated: 0,
      rekapRowsUpdated: 0,
      detailRowsUpdated: 0,
      createdRows: 0,
      repairedPointers: 0,
      inactiveRowsCleared: 0,
      ambiguousRows: [],
      errors: (plan.errors || []).slice(),
      assignments: plan.activeAssignments.map(function (assignment) {
        return {
          activityId: assignment.activityId,
          subActivityId: assignment.subActivityId,
          subActivityName: assignment.subActivity.sub_activity_name,
          localSortOrder: assignment.localSortOrder,
          globalNumber: assignment.globalNumber
        };
      })
    };
    if (!applyChanges) return report;

    const grouped = {};
    plan.assignments.forEach(function (assignment) {
      const spreadsheetId = String(
        assignment.subActivity.spreadsheet_file_id ||
        assignment.activity.spreadsheet_file_id ||
        ''
      );
      if (!spreadsheetId) {
        if (assignment.isActive) {
          report.errors.push({
            subActivityId: assignment.subActivityId,
            message: 'Spreadsheet arsip belum dipetakan.'
          });
        }
        return;
      }
      if (!grouped[spreadsheetId]) grouped[spreadsheetId] = [];
      grouped[spreadsheetId].push(assignment);
    });

    Object.keys(grouped).forEach(function (spreadsheetId) {
      let ss;
      let rekapSheet;
      try {
        ss = openSpreadsheetById_(spreadsheetId);
        rekapSheet = findRekapSheet_(ss);
        if (!rekapSheet) throw new Error('Sheet rekap tidak ditemukan.');
        ensureRekapDocumentColumns_(rekapSheet);
      } catch (error) {
        report.errors.push({ spreadsheetId: spreadsheetId, message: error.message });
        return;
      }

      let lookup = buildRekapRowLookup_(rekapSheet);
      const activeAssignments = grouped[spreadsheetId].filter(function (assignment) {
        return assignment.isActive;
      });
      const inactiveAssignments = grouped[spreadsheetId].filter(function (assignment) {
        return !assignment.isActive;
      });
      activeAssignments.forEach(function (assignment) {
        const activity = assignment.activity;
        const subActivity = Object.assign({}, assignment.subActivity, {
          sort_order: assignment.globalNumber,
          local_sort_order: assignment.localSortOrder
        });
        const stableRows = findStableRekapRowsFromLookup_(lookup, subActivity);
        if (stableRows.length > 1) {
          report.ambiguousRows.push({
            subActivityId: assignment.subActivityId,
            spreadsheetId: spreadsheetId,
            rows: stableRows.slice(),
            reason: 'duplicate_stable_identity'
          });
          return;
        }

        let rowIndex = findRekapRowForSubActivityFromLookup_(lookup, subActivity);
        let createdNow = false;
        if (!rowIndex) {
          const possibleRows = findPossibleRekapRowsFromLookup_(lookup, subActivity);
          if (possibleRows.length) {
            report.ambiguousRows.push({
              subActivityId: assignment.subActivityId,
              spreadsheetId: spreadsheetId,
              rows: possibleRows,
              reason: 'ambiguous_candidates'
            });
            return;
          }
          try {
            const created = SpreadsheetService.appendRekapRowIfPresent(activity, subActivity);
            if (!created || !created.rowNumber) throw new Error('Baris rekap tidak dapat dibuat.');
            rowIndex = created.rowNumber;
            createdNow = true;
            report.createdRows++;
            lookup = buildRekapRowLookup_(rekapSheet);
          } catch (error) {
            report.errors.push({
              spreadsheetId: spreadsheetId,
              subActivityId: assignment.subActivityId,
              message: error.message
            });
            return;
          }
        }

        try {
          const headerMap = getRekapHeaderMap_(rekapSheet);
          writeRekapIdentity_(rekapSheet, rowIndex, activity, subActivity, { respectManualLocks: true });
          setRekapStaticCell_(
            rekapSheet,
            rowIndex,
            headerMap,
            REKAP_SUMMARY_COLUMNS.nomorBerkas,
            assignment.globalNumber
          );
          setRekapStaticCell_(
            rekapSheet,
            rowIndex,
            headerMap,
            REKAP_SUMMARY_COLUMNS.noFolder,
            assignment.globalNumber
          );
          writeRekapIdentityMarker_(rekapSheet, rowIndex, subActivity);
          report.rekapRowsUpdated++;

          const detailSheet = ensureDetailSheet_(ss, activity, subActivity);
          const detailResult = writeDetailArchiveNumber_(detailSheet, assignment.globalNumber);
          report.detailRowsUpdated += detailResult.changedRows;
          if (createdNow) {
            SpreadsheetService.updateRekapSummary(activity, subActivity, {});
            lookup = buildRekapRowLookup_(rekapSheet);
          }
        } catch (error) {
          report.errors.push({
            spreadsheetId: spreadsheetId,
            subActivityId: assignment.subActivityId,
            message: error.message
          });
        }
      });

      lookup = buildRekapRowLookup_(rekapSheet);
      inactiveAssignments.forEach(function (assignment) {
        const rowIndex = findRekapRowForSubActivityFromLookup_(lookup, assignment.subActivity);
        if (!rowIndex) return;
        const headerMap = getRekapHeaderMap_(rekapSheet);
        const nomorBerkasCol = findRekapHeaderColumnFromMap_(
          headerMap,
          REKAP_SUMMARY_COLUMNS.nomorBerkas
        );
        const noFolderCol = findRekapHeaderColumnFromMap_(
          headerMap,
          REKAP_SUMMARY_COLUMNS.noFolder
        );
        let cleared = false;
        [nomorBerkasCol, noFolderCol].forEach(function (column) {
          if (!column) return;
          const cell = rekapSheet.getRange(rowIndex, column);
          if (!String(cell.getDisplayValue() || '').trim()) return;
          cell.clearContent();
          cleared = true;
        });
        if (cleared) report.inactiveRowsCleared++;
      });

      sortRekapSheetByNomorBerkas_(rekapSheet);
      lookup = buildRekapRowLookup_(rekapSheet);
      activeAssignments.forEach(function (assignment) {
        const subActivity = Object.assign({}, assignment.subActivity, {
          sort_order: assignment.globalNumber
        });
        const actualRow = findRekapRowForSubActivityFromLookup_(lookup, subActivity);
        if (!actualRow) return;
        assignment.rekapRowNumber = actualRow;
        if (Number(assignment.subActivity.rekap_row_number) !== Number(actualRow)) {
          report.repairedPointers++;
        }
      });
    });

    report.configCellsUpdated = ConfigRepository.bulkUpdateSubActivityNumbering(
      year,
      plan.assignments.map(function (assignment) {
        return {
          activityId: assignment.activityId,
          subActivityId: assignment.subActivityId,
          localSortOrder: assignment.localSortOrder,
          globalNumber: assignment.globalNumber,
          rekapRowNumber: assignment.rekapRowNumber
        };
      })
    );
    CacheHelper.invalidate(year);
    return report;
  },

  markRekapSubActivityInactive: function (activity, subActivity, reason) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    ensureRekapDocumentColumns_(sheet);
    const rowIndex = findRekapRowForSubActivity_(sheet, subActivity);
    if (!rowIndex) return null;

    const headerMap = getRekapHeaderMap_(sheet);
    const message = reason ? 'Nonaktif: ' + reason : 'Nonaktif';
    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.ket, message, true);

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    };
  },

  clearRekapSubActivityInactiveMark: function (activity, subActivity) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    const rowIndex = findRekapRowForSubActivity_(sheet, subActivity);
    if (!rowIndex) return null;

    const column = findRekapHeaderColumn_(sheet, REKAP_SUMMARY_COLUMNS.ket);
    if (column) {
      const cell = sheet.getRange(rowIndex, column);
      const value = String(cell.getDisplayValue() || '');
      if (value.toLowerCase().indexOf('nonaktif') === 0) cell.clearContent();
    }

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    };
  },

  updateArchiveDocumentLink: function (activity, subActivity, categoryName, file) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) throw new Error('Sheet rekap "' + REKAP_SHEET_NAME + '" tidak ditemukan.');

    const category = resolveRekapDocumentCategory_(categoryName);
    if (!category) throw new Error('Kategori dokumen tidak dikenal: ' + categoryName);

    ensureRekapDocumentColumns_(sheet);
    const column = findRekapHeaderColumn_(sheet, category.match);
    if (!column) throw new Error('Kolom "' + category.label + '" tidak ditemukan di sheet rekap.');

    const rowIndex = findOrCreateRekapRow_(sheet, activity, subActivity);
    const displayName = file && file.name ? file.name : category.label;
    const linkUrl = file && file.url ? file.url : '';
    const cell = sheet.getRange(rowIndex, column);

    if (linkUrl) {
      const hsep = typeof formulaSep_ === 'function' ? formulaSep_(ss) : ';';
      cell.setFormula('=HYPERLINK("' + linkUrl + '"' + hsep + ' "' + displayName.replace(/"/g, '""').replace(/\n/g, '" & CHAR(10) & "') + '")');
    } else {
      cell.setValue(sanitizeCellValue_(linkUrl || displayName));
    }
    cell.setWrap(true).setVerticalAlignment('middle');

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex,
      columnNumber: column,
      category: category.label
    };
  },

  clearArchiveDocumentLink: function (activity, subActivity, categoryName) {
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    const category = resolveRekapDocumentCategory_(categoryName);
    if (!category) return null;

    const column = findRekapHeaderColumn_(sheet, category.match);
    if (!column) return null;

    const rowIndex = findRekapRowForSubActivity_(sheet, subActivity);
    if (!rowIndex) return null;

    const cell = sheet.getRange(rowIndex, column);
    cell.clearContent();

    return {
      spreadsheetId: ss.getId(),
      rowNumber: rowIndex,
      columnNumber: column,
      category: category.label
    };
  },

  updateRekapDocumentMetadata: function (activity, subActivity, metadata, locks) {
    metadata = metadata || {};
    locks = locks || {};
    const summaryResult = SpreadsheetService.updateRekapSummary(activity, subActivity, {});
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) throw new Error('Sheet rekap "' + REKAP_SHEET_NAME + '" tidak ditemukan.');

    ensureRekapDocumentColumns_(sheet);
    const rowIndex = findOrCreateRekapRow_(sheet, activity, subActivity);
    const headerMap = getRekapHeaderMap_(sheet);
    const detailSheet = ensureDetailSheet_(ss, activity, subActivity);
    const summary = summarizeDetailSheet_(detailSheet, subActivity);
    const computedJumlah = summary.sumLembar ? summary.sumLembar + ' lembar' : '';

    const archiveNumber = getRekapArchiveNumber_(
      sheet,
      rowIndex,
      resolveSubActivityArchiveNumber_(subActivity, activity)
    );
    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas, archiveNumber, true);
    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kodeKlasifikasi, metadata.kodeKlasifikasi, true);

    if (locks.kurunWaktu === false) {
      setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu, metadata.kurunWaktu || formatDateRange_(summary.startDate, summary.endDate));
    } else {
      setRekapFormulaCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu, buildKurunWaktuFormula_(detailSheet));
    }

    if (locks.jumlah === false) {
      setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.jumlah, metadata.jumlah || computedJumlah);
    } else {
      setRekapFormulaCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.jumlah, buildJumlahFormula_(detailSheet));
    }

    if (locks.filingCabinet === false && metadata.noFilingCabinet !== undefined) {
      setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet, metadata.noFilingCabinet, true);
    } else if (locks.filingCabinet !== false) {
      setRekapFilingCabinetCell_(sheet, rowIndex, headerMap, activity, true);
    }

    if (locks.noLaci === false && metadata.noLaci !== undefined) {
      setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noLaci, metadata.noLaci, true);
    } else if (locks.noLaci !== false) {
      const rawLaciText = subActivity.sub_activity_name || '';
      const laciText = rawLaciText.replace(/"/g, '""');
      const laciCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noLaci);
      if (laciCol) {
        if (subActivity.folder_id) {
          const hsep = typeof formulaSep_ === 'function' ? formulaSep_(sheet.getParent()) : ';';
          sheet.getRange(rowIndex, laciCol).setFormula('=HYPERLINK("https://drive.google.com/drive/folders/' + subActivity.folder_id + '"' + hsep + ' "' + laciText.replace(/"/g, '""').replace(/\n/g, '" & CHAR(10) & "') + '")');
        } else {
          const laciCell = sheet.getRange(rowIndex, laciCol);
          if (!laciCell.getFormula()) laciCell.setValue(sanitizeCellValue_(rawLaciText));
        }
      }
    }

    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noFolder, archiveNumber, false);

    if (locks.akses === false) {
      setRekapStaticCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, metadata.klasifikasiAkses || summary.akses);
    } else {
      setRekapFormulaCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.akses, buildAksesFormula_(detailSheet));
    }

    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.ket, metadata.ket, true);

    return Object.assign({}, summaryResult || {}, {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    });
  },

  renameSubActivitySheet: function (activity, oldName, newName, subActivity) {
    if (!oldName || !newName || oldName === newName) return false;
    const ss = openSpreadsheetById_(getArchiveSpreadsheetId_(activity, subActivity));
    const safeOldName = normalizeSheetName_(oldName);
    const safeNewName = normalizeSheetName_(newName);

    const sheet = ss.getSheetByName(oldName) || ss.getSheetByName(safeOldName);
    if (!sheet) return false;

    if (ss.getSheetByName(newName) || ss.getSheetByName(safeNewName)) return false;

    try {
      sheet.setName(safeNewName);
      return true;
    } catch (e) {
      console.error('SpreadsheetService.renameSubActivitySheet failed: ' + e.message);
      return false;
    }
  }
};
// Private helper functions moved to SheetHelpers.gs
