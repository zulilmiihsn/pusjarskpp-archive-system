/**
 * Script untuk memproses OCR secara background (latar belakang)
 * Dipicu oleh Time-driven Trigger setiap 15 menit.
 */

const OCR_TRIGGER_FUNCTION = 'processOcrQueue';
const OCR_MAX_EXECUTION_TIME_MS = 240000; // 4 minutes

function processOcrQueue() {
  const props = PropertiesService.getScriptProperties();
  const isRunning = props.getProperty('OCR_JOB_RUNNING');
  if (isRunning === 'true') {
    console.log('OCR Job is already running. Skipping this cycle.');
    return;
  }
  
  props.setProperty('OCR_JOB_RUNNING', 'true');
  const startTime = Date.now();
  try {
    const currentYear = new Date().getFullYear();
    const config = CacheHelper.getConfig(currentYear);
    if (!config || !config.subActivities) return;

    const spreadsheets = {};
    config.subActivities.forEach(function(sub) {
      if (sub.spreadsheet_file_id) spreadsheets[sub.spreadsheet_file_id] = true;
    });

    const ssIds = Object.keys(spreadsheets);
    
    // Load continuation state
    const stateStr = props.getProperty('OCR_JOB_STATE');
    let state = stateStr ? JSON.parse(stateStr) : { ssIndex: 0, sheetIndex: 0 };
    
    let timeLimitReached = false;

    for (let j = state.ssIndex; j < ssIds.length; j++) {
      if (timeLimitReached) break;
      
      try {
        const ss = SpreadsheetApp.openById(ssIds[j]);
        const sheets = ss.getSheets();
        
        let startSheetIdx = (j === state.ssIndex) ? state.sheetIndex : 0;
        
        for (let i = startSheetIdx; i < sheets.length; i++) {
          if (Date.now() - startTime > OCR_MAX_EXECUTION_TIME_MS) {
            // Save state and exit gracefully to avoid 6-minute timeout
            props.setProperty('OCR_JOB_STATE', JSON.stringify({ ssIndex: j, sheetIndex: i }));
            console.log('OCR Job reached time limit. Will resume from SS index ' + j + ' sheet index ' + i);
            timeLimitReached = true;
            break;
          }
          
          const sheet = sheets[i];
          const name = sheet.getName();
          if (name === REKAP_SHEET_NAME || name.indexOf('config_') === 0) continue;
          processSheetOcr_(sheet);
        }
      } catch (e) {
        console.warn('Gagal buka spreadsheet OCR ' + ssIds[j] + ': ' + e.message);
      }
    }
    
    // If we finished everything naturally, clear the state
    if (!timeLimitReached) {
      props.deleteProperty('OCR_JOB_STATE');
      console.log('OCR Job completed full cycle.');
    }
    
  } catch (e) {
    console.error('Error in processOcrQueue: ' + e.message);
  } finally {
    props.deleteProperty('OCR_JOB_RUNNING');
  }
}

function processSheetOcr_(sheet) {
  const startCol = getDetailStartColumn_(sheet);
  if (!startCol) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < DETAIL_DATA_START_ROW) return;

  const dataRange = sheet.getRange(DETAIL_DATA_START_ROW, startCol, lastRow - DETAIL_DATA_START_ROW + 1, DETAIL_FIELD_ORDER.length);
  const data = dataRange.getValues();

  let uraianIdx = DETAIL_FIELD_ORDER.indexOf('uraian_informasi_item');
  if (uraianIdx === -1) uraianIdx = DETAIL_FIELD_ORDER.indexOf('uraian_informasi_berkas'); // fallback
  const fileUrlIdx = DETAIL_FIELD_ORDER.indexOf('lokasi_simpan');
  const fileIdIdx = DETAIL_FIELD_ORDER.indexOf('file_id');
  
  if (uraianIdx === -1 || fileUrlIdx === -1) return;

  // Cek apakah ada baris yang butuh OCR
  const updates = [];
  
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const uraian = String(row[uraianIdx] || '').trim();
    // Kalau uraian kosong ATAU isinya cuma nama file (tanpa spasi/ciri khas kalimat), kita asumsikan belum di OCR
    if (uraian === '' || uraian.indexOf(' ') === -1) {
      
      // Butuh file_id untuk mengambil file
      let fileId = '';
      if (fileIdIdx !== -1 && row[fileIdIdx]) {
        fileId = row[fileIdIdx];
      } else {
        // Fallback: ekstrak dari URL
        const url = row[fileUrlIdx];
        if (url) {
          const match = url.match(/[-\w]{25,}/);
          if (match) fileId = match[0];
        }
      }

      if (fileId) {
        console.log('Menjalankan OCR untuk baris ' + (r + DETAIL_DATA_START_ROW) + ', fileId: ' + fileId);
        const parsedMeta = runOcrExtraction_(fileId);
        if (parsedMeta) {
          updates.push({
            rowIndex: r + DETAIL_DATA_START_ROW,
            meta: parsedMeta
          });
        }
      }
    }
  }

  // Jika ada update, lakukan batchUpdate
  if (updates.length > 0) {
    applyOcrUpdates_(sheet, startCol, updates, data);
  }
}

function runOcrExtraction_(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    const fileName = file.getName();
    
    let text = '';
    
    // Hanya lakukan OCR untuk gambar atau PDF
    if (mimeType !== 'application/pdf' && mimeType.indexOf('image/') !== 0) {
      return null; 
    }

    // OCR via Drive API v3
    const blob = file.getBlob();
    const resource = { 
      name: fileName + ' (Temp OCR)',
      mimeType: 'application/vnd.google-apps.document' 
    };
    const copy = Drive.Files.create(resource, blob, { ocrLanguage: 'id', supportsAllDrives: true });
    const tempDocId = copy.id;
    
    const url = 'https://www.googleapis.com/drive/v3/files/' + tempDocId + '/export?mimeType=text/plain';
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      text = response.getContentText();
    }
    
    // Hapus dokumen temp
    try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (e) {}

    if (!text || text.length < 100) return null;

    // Gunakan ParseEngine untuk ekstrak field
    const engineResult = ParseEngine.analyze(text, fileName, { activity: {}, subActivity: {} });
    return engineResult.fields;

  } catch (e) {
    console.warn('Gagal background OCR file ' + fileId + ': ' + e.message);
    return null;
  }
}

function applyOcrUpdates_(sheet, startCol, updates, oldData) {
  const requests = [];
  
  for (let i = 0; i < updates.length; i++) {
    const rowIndex = updates[i].rowIndex;
    const meta = updates[i].meta;
    const rowIdxInData = rowIndex - DETAIL_DATA_START_ROW;
    const oldRow = oldData[rowIdxInData];
    
    const rowData = { values: [] };
    
    for (let c = 0; c < DETAIL_FIELD_ORDER.length; c++) {
      const fieldName = DETAIL_FIELD_ORDER[c];
      const oldVal = oldRow[c] || '';
      let newVal = oldVal;
      
      // Jangan timpa data yang sudah ada nilainya, KECUALI uraian yg kosong/hanya nama file
      if (fieldName === 'uraian_informasi_item' || fieldName === 'uraian_informasi_berkas') {
        if (meta.uraian && meta.uraian.value && (!oldVal || oldVal.toString().indexOf(' ') === -1)) {
          newVal = meta.uraian.value;
        }
      } else if (fieldName === 'tanggal') {
        if (meta.tanggal && meta.tanggal.value && !oldVal) {
          newVal = meta.tanggal.value;
        }
      } else if (fieldName === 'no_berkas' || fieldName === 'item_number') {
        if (meta.no_surat && meta.no_surat.value && !oldVal) {
          newVal = meta.no_surat.value;
        }
      }

      rowData.values.push({
        userEnteredValue: { stringValue: String(newVal) }
      });
    }

    requests.push({
      updateCells: {
        rows: [rowData],
        fields: "userEnteredValue",
        range: {
          sheetId: sheet.getSheetId(),
          startRowIndex: rowIndex - 1,
          endRowIndex: rowIndex,
          startColumnIndex: startCol - 1,
          endColumnIndex: startCol - 1 + DETAIL_FIELD_ORDER.length
        }
      }
    });
  }

  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, sheet.getParent().getId());
    console.log('Applied ' + requests.length + ' OCR updates to sheet ' + sheet.getName());
  }
}

function installOcrTrigger() {
  // Hapus pemicu lama agar tidak duplikat
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === OCR_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Buat pemicu baru tiap 15 menit
  ScriptApp.newTrigger(OCR_TRIGGER_FUNCTION)
    .timeBased()
    .everyMinutes(15)
    .create();
    
  return 'Trigger OCR Background berhasil dipasang.';
}
