'use strict';

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
  // Lebar tulis = max(skema, lebar sheet sekarang) supaya kolom ekstra milik baris
  // tahun-lain (mis. skema lama lebih lebar) tidak ikut terpotong (B1).
  const width = Math.max(headers.length, sheet.getLastColumn());
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
    return wsPadRow_(row, width);
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


