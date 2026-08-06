'use strict';

/** @private Workspace year deletion and active-year lifecycle. */
const SettingsYearImpl_ = {
  deleteYear: function (payload) {
    payload = payload || {};
    const year = Number(payload.year);
    if (!year || isNaN(year)) {
      throw new Error('Tahun kerja tidak valid.');
    }

    const settings = ConfigService.getSettings();
    const ss = ConfigRepository.getConfigSpreadsheet();

    // 1. Find the year record in config_years
    const yearsSheet = ss.getSheetByName(CONFIG_SHEETS.YEARS);
    if (!yearsSheet) throw new Error('Sheet config_years tidak ditemukan.');
    const yearsValues = yearsSheet.getDataRange().getDisplayValues();
    const yearsHeaders = yearsValues[0].map(normalizeHeader_);
    const yearCol = yearsHeaders.indexOf('year');
    
    let spreadsheetYearFolderId = '';
    let persuratanYearFolderId = '';

    if (yearCol !== -1) {
      for (let i = 1; i < yearsValues.length; i++) {
        if (Number(yearsValues[i][yearCol]) === year) {
          const spreadsheetFolderCol = yearsHeaders.indexOf('spreadsheet_year_folder_id');
          const persuratanFolderCol = yearsHeaders.indexOf('persuratan_year_folder_id');
          if (spreadsheetFolderCol !== -1) spreadsheetYearFolderId = yearsValues[i][spreadsheetFolderCol];
          if (persuratanFolderCol !== -1) persuratanYearFolderId = yearsValues[i][persuratanFolderCol];
          break;
        }
      }
    }

    // 2. Move associated folders in Google Drive to Trash
    const trashReport = [];
    
    // Trash spreadsheet year folder (ARSIP DIKLAT {year})
    if (spreadsheetYearFolderId) {
      try {
        const folder = DriveApp.getFolderById(spreadsheetYearFolderId);
        folder.setTrashed(true);
        trashReport.push('Folder Arsip Diklat berhasil dibuang ke sampah.');
      } catch (e) {
        console.warn('Error trashing spreadsheet year folder: ' + e.message);
        trashReport.push('Gagal membuang Folder Arsip Diklat: ' + e.message);
      }
    }

    // Trash persuratan year folder (Tahun {year} under Persuratan) and try to find/trash dokumen year folder (Tahun {year} under Dokumen)
    if (persuratanYearFolderId) {
      try {
        const folder = DriveApp.getFolderById(persuratanYearFolderId);
        
        // Find document year folder under naskahDinas/2. Dokumen/Tahun {year}
        try {
          const parentFolders = folder.getParentFolders();
          if (parentFolders.hasNext()) {
            const persuratanFolder = parentFolders.next();
            const naskahDinasFolders = persuratanFolder.getParentFolders();
            if (naskahDinasFolders.hasNext()) {
              const naskahDinas = naskahDinasFolders.next();
              const docFolders = naskahDinas.getFoldersByName('2. Dokumen');
              if (docFolders.hasNext()) {
                const docFolder = docFolders.next();
                const docYearFolders = docFolder.getFoldersByName('Tahun ' + year);
                if (docYearFolders.hasNext()) {
                  docYearFolders.next().setTrashed(true);
                  trashReport.push('Folder Dokumen Tahun ' + year + ' berhasil dibuang ke sampah.');
                }
              }
            }
          }
        } catch (innerErr) {
          console.warn('Error locating and trashing doc year folder: ' + innerErr.message);
        }

        folder.setTrashed(true);
        trashReport.push('Folder Persuratan Tahun ' + year + ' berhasil dibuang ke sampah.');
      } catch (e) {
        console.warn('Error trashing persuratan year folder: ' + e.message);
        trashReport.push('Gagal membuang Folder Persuratan: ' + e.message);
      }
    }

    // 3. Delete config rows matching the year in all configuration sheets
    const sheetsToDelete = [
      { name: CONFIG_SHEETS.YEARS, col: 'year' },
      { name: CONFIG_SHEETS.ACTIVITIES, col: 'year' },
      { name: CONFIG_SHEETS.SUB_ACTIVITIES, col: 'year' },
      { name: CONFIG_SHEETS.METADATA_FIELDS, col: 'year' },
      { name: CONFIG_SHEETS.ARCHIVE_LOG, col: 'year' }
    ];

    sheetsToDelete.forEach(function (sheetConfig) {
      try {
        const sheet = ss.getSheetByName(sheetConfig.name);
        if (sheet) {
          const values = sheet.getDataRange().getValues();
          if (values.length >= 2) {
            const headers = values[0].map(normalizeHeader_);
            const colIdx = headers.indexOf(sheetConfig.col);
            if (colIdx !== -1) {
              let deleteCount = 0;
              for (let i = values.length - 1; i >= 1; i--) {
                if (Number(values[i][colIdx]) === year) {
                  sheet.deleteRow(i + 1);
                  deleteCount++;
                }
              }
            }
          }
        }
      } catch (err) {
        console.warn('Error deleting rows from ' + sheetConfig.name + ': ' + err.message);
      }
    });

    // 4. Invalidate Script Cache
    CacheHelper.invalidateAll();

    // 5. If the deleted year was the current active year, switch to another available year
    let nextActiveYear = Number(settings.currentYear);
    if (nextActiveYear === year) {
      nextActiveYear = 0;
      try {
        const remainingYearsValues = yearsSheet.getDataRange().getDisplayValues();
        const yearColIdx = remainingYearsValues[0].map(normalizeHeader_).indexOf('year');
        if (yearColIdx !== -1 && remainingYearsValues.length > 1) {
          // Select the first remaining year in the list
          nextActiveYear = Number(remainingYearsValues[1][yearColIdx]);
        }
      } catch (e) {
        console.warn('Error finding next active year: ' + e.message);
      }

      if (nextActiveYear && !isNaN(nextActiveYear)) {
        ConfigService.saveSettings({ currentYear: nextActiveYear });
      } else {
        // No remaining years, reset currentYear settings or fallback to DEFAULT_YEAR
        ConfigService.saveSettings({ currentYear: '' });
      }
    }

    return {
      success: true,
      year: year,
      nextActiveYear: nextActiveYear || null,
      trashReport: trashReport
    };
  },

};
