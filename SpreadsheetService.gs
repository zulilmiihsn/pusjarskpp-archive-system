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
    if (key === 'lokasi_simpan' && metadata._lokasi_simpan_url) {
      return sanitizeCellValue_(metadata._lokasi_simpan_url);
    }
    if (key === 'jumlah_satuan') {
      return sanitizeCellValue_(metadata.satuan || metadata.jumlah_satuan || '');
    }
    return sanitizeCellValue_(metadata[key]);
  });
}

function extractMetadataFromRow_(rowValues) {
  var metadata = {};
  DETAIL_FIELD_ORDER.forEach(function (key, i) {
    var val = rowValues[i];
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

const SpreadsheetService = {
  getNextItemNumber: function (activity, subActivity) {
    try {
      const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
      var maxNum = 0;
      values.forEach(function (row) {
        var raw = String(row[0] || '').trim();
        if (raw) {
          var num = Number(raw);
          if (!isNaN(num) && num > maxNum) maxNum = num;
        }
      });
      return String(maxNum + 1).padStart(2, '0');
    } catch (e) {
      Logger.log('Error calculating next item number: ' + e);
      return '01';
    }
  },

  appendArchiveRow: function (activity, subActivity, metadata) {
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const rowIndex = findWritableDetailRow_(sheet);
    const rowValues = buildDetailRowValues_(metadata);
    const startCol = getDetailStartColumn_(sheet);

    sheet.getRange(rowIndex, startCol, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(rowIndex, startCol + DETAIL_ITEM_NUMBER_OFFSET, 1, 1).setNumberFormat('00');

    const lokasiSimpanIdx = DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
    if (lokasiSimpanIdx >= 0 && metadata._lokasi_simpan_url && SpreadsheetApp.newRichTextValue) {
      const cell = sheet.getRange(rowIndex, startCol + lokasiSimpanIdx);
      const url = metadata._lokasi_simpan_url;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(url)
        .setLinkUrl(url)
        .build();
      cell.setRichTextValue(richText);
    }
    sheet.getRange(rowIndex, startCol, 1, rowValues.length)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment('middle');
    sortDetailSheetByNoBerkas_(sheet);

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    };
  },

  appendArchiveRowsBulk: function (activity, subActivity, metadataList) {
    if (!metadataList || metadataList.length === 0) return [];
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const rowIndex = findWritableDetailRow_(sheet);
    const startCol = getDetailStartColumn_(sheet);

    const allRowValues = metadataList.map(function(meta) { return buildDetailRowValues_(meta); });

    sheet.getRange(rowIndex, startCol, allRowValues.length, DETAIL_FIELD_ORDER.length).setValues(allRowValues);
    sheet.getRange(rowIndex, startCol + DETAIL_ITEM_NUMBER_OFFSET, allRowValues.length, 1).setNumberFormat('00');

    const lokasiSimpanIdx = DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
    if (lokasiSimpanIdx >= 0 && SpreadsheetApp.newRichTextValue) {
      const richTextArray = [];
      for (let i = 0; i < metadataList.length; i++) {
        const meta = metadataList[i];
        const url = meta._lokasi_simpan_url;
        if (url) {
          richTextArray.push([SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build()]);
        } else {
          const val = allRowValues[i][lokasiSimpanIdx];
          richTextArray.push([SpreadsheetApp.newRichTextValue().setText(val || '').build()]);
        }
      }
      sheet.getRange(rowIndex, startCol + lokasiSimpanIdx, allRowValues.length, 1).setRichTextValues(richTextArray);
    }
    
    sheet.getRange(rowIndex, startCol, allRowValues.length, DETAIL_FIELD_ORDER.length)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment('middle');
    sortDetailSheetByNoBerkas_(sheet);

    return metadataList.map(function(meta, i) {
      return {
        spreadsheetId: ss.getId(),
        spreadsheetUrl: ss.getUrl(),
        sheetName: sheet.getName(),
        rowNumber: rowIndex + i
      };
    });
  },

  getArchiveRow: function (activity, subActivity, rowIndex) {
    if (!rowIndex) return {};
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const startCol = getDetailStartColumn_(sheet);
    const rowValues = sheet.getRange(rowIndex, startCol, 1, DETAIL_FIELD_ORDER.length).getValues()[0];
    const metadata = extractMetadataFromRow_(rowValues);
    return metadata;
  },

  getArchiveRowByFileId: function (activity, subActivity, fileId, fallbackRowIndex) {
    if (!fileId && !fallbackRowIndex) return {};
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    const rowValues = buildDetailRowValues_(metadata);
    const startCol = getDetailStartColumn_(sheet);

    sheet.getRange(rowIndex, startCol, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(rowIndex, startCol + DETAIL_ITEM_NUMBER_OFFSET, 1, 1).setNumberFormat('00');

    const lokasiSimpanIdx = DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
    if (lokasiSimpanIdx >= 0 && metadata._lokasi_simpan_url && SpreadsheetApp.newRichTextValue) {
      const cell = sheet.getRange(rowIndex, startCol + lokasiSimpanIdx);
      const url = metadata._lokasi_simpan_url;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(url)
        .setLinkUrl(url)
        .build();
      cell.setRichTextValue(richText);
    }
    sortDetailSheetByNoBerkas_(sheet);

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    };
  },

  ensureSubActivitySheet: function (activity, subActivity) {
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = ensureDetailSheet_(ss, activity, subActivity);
    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName()
    };
  },

  getDetailMetadataDefaults: function (activity, subActivity, forceCalculate) {
    var ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));

    let locks = {};
    try {
      locks = subActivity && subActivity.metadata_locks ? JSON.parse(subActivity.metadata_locks) : {};
    } catch(e) { console.warn('getDetailMetadataDefaults: failed to parse metadata_locks: ' + e.message); }

    var needSummary = forceCalculate || (locks.kurunWaktu === false) || (locks.jumlah === false) || (locks.akses === false);

    var sheet = findExistingDetailSheet_(ss, subActivity);
    var summary = (sheet && needSummary)
      ? summarizeDetailSheet_(sheet, subActivity)
      : { count: 0, startDate: null, endDate: null, akses: '' };

    var rekapKode = '', rekapFC = '', rekapLaci = '', rekapFolder = '', rekapAkses = '';
    var rekapNomorBerkas = '', rekapKurunWaktu = '', rekapJumlah = '', rekapKet = '';
    var rekapSheet, rowIndex, dfpCol, materiCol, laporanCol, sertifikatCol;
    try {
      rekapSheet = findRekapSheet_(ss);
      if (rekapSheet) {
        rowIndex = findRekapRowForSubActivity_(rekapSheet, subActivity);
        if (rowIndex) {
          var headerMap = getRekapHeaderMap_(rekapSheet);
          dfpCol        = findRekapHeaderColumnFromMap_(headerMap, ['data fix peserta']);
          materiCol     = findRekapHeaderColumnFromMap_(headerMap, ['kumpulan materi']);
          laporanCol    = findRekapHeaderColumnFromMap_(headerMap, ['laporan']);
          sertifikatCol = findRekapHeaderColumnFromMap_(headerMap, ['sertifikat']);
          var kodeCol       = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.kodeKlasifikasi);
          var fcCol         = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet);
          var laciCol       = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noLaci);
          var folderCol     = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.noFolder);
          var aksesCol      = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.akses);
          var nomorBerkasCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas);
          var kurunWaktuCol = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.kurunWaktu);
          var jumlahCol     = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.jumlah);
          var ketCol        = findRekapHeaderColumnFromMap_(headerMap, REKAP_SUMMARY_COLUMNS.ket);

          // Batch-read entire rekap row in one call
          var allCols = [dfpCol, materiCol, laporanCol, sertifikatCol, kodeCol, fcCol, laciCol, folderCol, aksesCol, nomorBerkasCol, kurunWaktuCol, jumlahCol, ketCol].filter(Boolean);
          var maxCol = Math.max.apply(null, allCols);
          var rowValues = rekapSheet.getRange(rowIndex, 1, 1, maxCol).getDisplayValues()[0];
          var cell = function (col) { return col ? String(rowValues[col - 1] || '').trim() : ''; };

          rekapKode        = cell(kodeCol);
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

    var showJsComputed = forceCalculate;
    var meta = {
      no_berkas: rekapNomorBerkas || (subActivity && subActivity.sort_order ? subActivity.sort_order : ''),
      kode_klasifikasi: rekapKode || (subActivity && subActivity.default_kode_klasifikasi) || '',
      kurun_waktu: (showJsComputed ? '' : rekapKurunWaktu) || formatDateRange_(summary.startDate, summary.endDate),
      jumlah: (showJsComputed ? '' : rekapJumlah) || ((summary.count || 0) + ' dokumen'),
      no_filing_cabinet: (showJsComputed ? '' : rekapFC) || ConfigService.getSettings().defaultFilingCabinet || '02',
      no_laci: (showJsComputed ? '' : rekapLaci) || (activity && activity.laci_no) || '',
      no_folder: (showJsComputed ? '' : rekapFolder) || (subActivity && subActivity.no_folder) || (activity && activity.folder_no) || '',
      klasifikasi_akses: (showJsComputed ? '' : rekapAkses) || summary.akses || 'Terbatas',
      ket: rekapKet || '',
      next_item_number: SpreadsheetService.getNextItemNumber(activity, subActivity)
    };

    if (rekapSheet && rowIndex) {
      // Batch-read formulas for URL columns in one call
      var urlCols = [dfpCol, materiCol, laporanCol, sertifikatCol].filter(Boolean);
      if (urlCols.length) {
        var urlMaxCol = Math.max.apply(null, urlCols);
        var formulas = rekapSheet.getRange(rowIndex, 1, 1, urlMaxCol).getFormulas()[0];
        var extractUrl = function (col) {
          if (!col) return '';
          var formula = formulas[col - 1] || '';
          var match = formula.match(/=HYPERLINK\(\s*["']([^"']+)["']/i);
          return match && match[1] ? match[1] : '';
        };
        if (dfpCol) meta.url_dfp = extractUrl(dfpCol);
        if (materiCol) meta.url_materi = extractUrl(materiCol);
        if (laporanCol) meta.url_laporan = extractUrl(laporanCol);
        if (sertifikatCol) meta.url_sertifikat = extractUrl(sertifikatCol);
      }
    }

    return meta;
  },

  listExistingItemUraianPairs: function (activity, subActivity) {
    var ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    var sheet = findExistingDetailSheet_(ss, subActivity);
    if (!sheet) return [];

    var noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    if (noteRow <= DETAIL_DATA_START_ROW) return [];

    var width = Math.max(sheet.getLastColumn(), DETAIL_FALLBACK_START_COL + DETAIL_FIELD_ORDER.length);
    var headerColumns = getDetailColumnMap_(sheet, width);
    var itemCol = headerColumns.nomor_item_arsip || (DETAIL_FALLBACK_START_COL + 1);
    var uraianCol = headerColumns.uraian_informasi_berkas || (DETAIL_FALLBACK_START_COL + 3);

    var rowCount = noteRow - DETAIL_DATA_START_ROW;
    var itemValues = sheet.getRange(DETAIL_DATA_START_ROW, itemCol, rowCount, 1).getDisplayValues();
    var uraianValues = sheet.getRange(DETAIL_DATA_START_ROW, uraianCol, rowCount, 1).getDisplayValues();

    var pairs = [];
    for (var i = 0; i < rowCount; i++) {
      var item = String(itemValues[i][0] || '').trim();
      var uraian = String(uraianValues[i][0] || '').trim();
      if (!item && !uraian) continue;
      pairs.push({
        rowNumber: DETAIL_DATA_START_ROW + i,
        nomor_item_arsip: item,
        uraian_informasi_berkas: uraian
      });
    }
    return pairs;
  },

  listExistingArchiveRows: function (activity, subActivity) {
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    ensureRekapDocumentColumns_(sheet);
    const noteRow = findNoteRow_(sheet) || sheet.getLastRow() + 1;
    const rowIndex = findFirstBlankInRange_(sheet, REKAP_DATA_START_ROW, noteRow - 1, 4) || noteRow;
    if (rowIndex === noteRow) {
      sheet.insertRowBefore(noteRow);
      invalidateNoteRowCache_(sheet);
    }

    const nomorBerkas = nextNumberInColumn_(sheet, REKAP_DATA_START_ROW, rowIndex - 1, 2);
    const row = [
      sanitizeCellValue_(nomorBerkas),
      sanitizeCellValue_(subActivity.default_kode_klasifikasi || ''),
      sanitizeCellValue_(getSubActivityFormalName_(subActivity) || ''),
      '',
      '',
      '02',
      sanitizeCellValue_(activity.laci_no || ''),
      sanitizeCellValue_(subActivity.no_folder || activity.folder_no || ''),
      'Terbatas',
      '',
      '',
      '',
      '',
      ''
    ];
    sheet.getRange(rowIndex, 2, 1, row.length).setValues([row]);
    sheet.getRange(rowIndex, 2, 1, row.length)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment('middle');

    return { sheetName: sheet.getName(), rowNumber: rowIndex };
  },

  updateRekapSummary: function (activity, subActivity, metadata) {
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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

    writeRekapSummary_(sheet, rowIndex, summary);
    sortRekapSheetByNomorBerkas_(sheet);

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    };
  },

  deleteArchiveRowAndReorder: function(activity, subActivity, rowNumber, fileId) {
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, nextSubActivity || previousSubActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) return null;

    ensureRekapDocumentColumns_(sheet);
    const rowIndex = findRekapRowForSubActivity_(sheet, nextSubActivity) ||
      findRekapRowForSubActivity_(sheet, previousSubActivity);
    if (!rowIndex) return null;

    writeRekapIdentity_(sheet, rowIndex, activity, nextSubActivity || previousSubActivity);

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sheet.getName(),
      rowNumber: rowIndex
    };
  },

  markRekapSubActivityInactive: function (activity, subActivity, reason) {
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
      const hsep = formulaSep_(ss);
      cell.setFormula('=HYPERLINK("' + linkUrl + '"' + hsep + ' "' + displayName.replace(/"/g, '""') + '")');
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
    const sheet = findRekapSheet_(ss);
    if (!sheet) throw new Error('Sheet rekap "' + REKAP_SHEET_NAME + '" tidak ditemukan.');

    ensureRekapDocumentColumns_(sheet);
    const rowIndex = findOrCreateRekapRow_(sheet, activity, subActivity);
    const headerMap = getRekapHeaderMap_(sheet);
    const detailSheet = ensureDetailSheet_(ss, activity, subActivity);
    const summary = summarizeDetailSheet_(detailSheet, subActivity);
    const computedJumlah = summary.count ? summary.count + ' dokumen' : '';

    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.nomorBerkas, metadata.nomorBerkas, true);
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

    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.filingCabinet, metadata.noFilingCabinet, true);
    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noLaci, metadata.noLaci, true);
    setRekapSummaryCell_(sheet, rowIndex, headerMap, REKAP_SUMMARY_COLUMNS.noFolder, metadata.noFolder, true);

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
    const ss = SpreadsheetApp.openById(getArchiveSpreadsheetId_(activity, subActivity));
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
