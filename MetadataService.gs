'use strict';

const MetadataService = {
  normalize: function (metadata, activity, subActivity, sourceName) {
    const normalized = Object.assign({}, metadata);
    
    const uraian = String(normalized.uraian_informasi_item || '').trim();
    normalized.uraian_informasi_item = uraian;
    
    if (normalized.nomor_item_arsip) {
      normalized.nomor_item_arsip = String(normalized.nomor_item_arsip).padStart(2, '0');
    }
    normalized.no_berkas = normalized.no_berkas || resolveSubActivityArchiveNumber_(subActivity, activity);

    normalized.no_laci = normalized.no_laci || '';
    // Paksa No Folder agar selalu sinkron dengan Nomor Item Arsip sesuai permintaan user
    normalized.no_folder = normalized.nomor_item_arsip || '';
    normalized.no_filing_cabinet = normalized.no_filing_cabinet || '';
    
    // Inject hyperlink metadata supaya SpreadsheetService tahu ini link
    if (activity) {
      normalized._no_filing_cabinet_path = activity.laci_no ? String(activity.laci_no).padStart(2, '0') + '. Laci ' + activity.activity_name : '';
      normalized._no_filing_cabinet_url = activity.laci_folder_id ? 'https://drive.google.com/drive/folders/' + activity.laci_folder_id : '';
    }
    if (subActivity) {
      normalized._no_laci_path = subActivity.sub_activity_name || '';
      normalized._no_laci_url = subActivity.folder_id ? 'https://drive.google.com/drive/folders/' + subActivity.folder_id : '';
    }

    normalized.satuan = normalized.satuan || 'Lembar';
    normalized.tingkat_perkembangan = normalized.tingkat_perkembangan || 'Asli';
    normalized.jumlah = metadata.jumlah !== undefined ? metadata.jumlah : 1;
    normalized.klasifikasi_akses = normalized.klasifikasi_akses || 'Terbatas';
    normalized.lokasi_simpan = normalized.lokasi_simpan || this.buildFinalFileName(normalized, sourceName);
    return normalized;
  },

  buildFinalFileName: buildFinalFileName_
};

function extractNomorSurat_(text) {
  const str = String(text || '');
  const patterns = [
    // 0. "Nomor : B-123/DL.01/2026" or "No. 123/ABC/2026" (most explicit, highest priority)
    /(?:No(?:mor)?\.?)\s*[:.]?\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    // 1. "No: 113/PDP" (explicit No: but without year)
    /(?:No(?:mor)?\.?)\s*[:.]\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+)/i,
    // 2. Letter prefix: "B-123/DL.01/2026", "SP-456/KP.02/2025"
    /\b([A-Z]{1,3}-\d{1,6}(?:\/[A-Z0-9.]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    // 3. Pure numeric segments with year: "123/DL.01/2026", "01/02/03/2026"
    /\b(\d{1,6}\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    // 4. Formal: "Nomor 45 Tahun 2025"
    /(?:No(?:mor)?\.?)\s*[:.]?\s*(\d{1,6})\s+Tahun\s+(\d{4})/i,
    // 5. Space/dash separated filename format with year
    /(?:No(?:mor)?)\s*[:.\-\s_]+\s*([A-Z0-9][A-Z0-9.\-\s_]+?[\-\s_][12]\d{3})\b/i,
    // 6. Strict Uppercase/Number without Year (e.g. 113/PDP, B-417/BKPSDM) - Case SENSITIVE
    new RegExp('\\b((?:[A-Z]{1,3}-\\d{1,6}|\\d{1,6})\\/[A-Z0-9.\\-]+(?:\\/[A-Z0-9.\\-]+)*)\\b')
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = str.match(patterns[i]);
    if (match) {
      if (i === 4) return match[1] + '/Tahun/' + match[2];
      if (i === 5) return match[1].replace(/[\s_]+/g, '/').trim();
      return match[1].replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function extractKodeKlasifikasi_(text) {
  const str = String(text || '');
  // Contextual: "Kode: KP.01.02" or "Klasifikasi: DL.01"
  const ctx = str.match(/(?:Kode|Klasifikasi)\s*[:.]?\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?)/i);
  if (ctx) return ctx[1].toUpperCase().trim();
  // Standard: "KP.01.02", "DL.01"
  const raw = str.match(/\b([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)\b/);
  return raw ? raw[1].trim() : '';
}

const _MONTH_NAMES = {
  januari: '01', februari: '02', maret: '03', april: '04',
  mei: '05', juni: '06', juli: '07', agustus: '08',
  september: '09', oktober: '10', november: '11', desember: '12'
};
const _MONTH_RE = 'Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember';

function extractDate_(text) {
  const str = String(text || '');
  const lines = str.split('\n');
  const headerLines = [];
  for (let i = 0; i < Math.min(lines.length, 30); i++) {
    const line = lines[i].trim();
    if (/^(?:yth|kepada|kpd|dengan\s+hormat|menindaklanjuti|sehubungan|merujuk|berdasarkan)\b/i.test(line)) {
      break;
    }
    headerLines.push(lines[i]);
  }
  const headerText = headerLines.join('\n');
  const date = extractDateFromText_(headerText);
  if (date) return date;
  return extractDateFromText_(str);
}

function extractDateFromText_(str) {
  // Priority 1: Contextual — near "tanggal", "ditetapkan", "ditandatangani"
  const ctxRe = new RegExp(
    '(?:tanggal|ditetapkan|ditandatangani)\\D{0,20}?' +
    '(?:' +
      '(\\d{1,2})\\s+(' + _MONTH_RE + ')\\s+(20\\d{2})' +
      '|(20\\d{2})[-/.](\\d{1,2})[-/.](\\d{1,2})' +
    ')', 'i');
  const ctx = str.match(ctxRe);
  if (ctx) {
    if (ctx[1] && ctx[2] && ctx[3]) {
      return [ctx[3], _MONTH_NAMES[ctx[2].toLowerCase()], pad2_(ctx[1])].join('-');
    }
    if (ctx[4] && ctx[5] && ctx[6]) {
      return [ctx[4], pad2_(ctx[5]), pad2_(ctx[6])].join('-');
    }
  }
  // Priority 2: ISO format anywhere — 2025-03-15, 2025/3/15
  const iso = str.match(/\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/);
  if (iso) return [iso[1], pad2_(iso[2]), pad2_(iso[3])].join('-');
  // Priority 3: Full Indonesian — "15 Maret 2025"
  const idRe = new RegExp('\\b(\\d{1,2})\\s+(' + _MONTH_RE + ')\\s+(20\\d{2})\\b', 'i');
  const id = str.match(idRe);
  if (id) return [id[3], _MONTH_NAMES[id[2].toLowerCase()], pad2_(id[1])].join('-');
  // Priority 4: Month + year only — "Maret 2025" → 1st of month
  const myRe = new RegExp('\\b(' + _MONTH_RE + ')\\s+(20\\d{2})\\b', 'i');
  const my = str.match(myRe);
  if (my) return [my[2], _MONTH_NAMES[my[1].toLowerCase()], '01'].join('-');
  // Priority 5: DD-MM-YYYY with dots/dashes (e.g., "15.03.2025")
  const dmy = str.match(/\b(\d{1,2})[.\-](\d{1,2})[.\-](20\d{2})\b/);
  if (dmy) return [dmy[3], pad2_(dmy[2]), pad2_(dmy[1])].join('-');
  return '';
}


function extractTingkatPerkembangan_(name) {
  const text = String(name || '').toLowerCase();
  // Hanya 2 tingkat: Asli (default/elektronik) dan Salinan (turunan/copy/cetak).
  if (text.indexOf('salinan') >= 0 || text.indexOf('copy') >= 0 || text.indexOf('cetak') >= 0) return 'Salinan';
  if (text.indexOf('asli') >= 0 || text.indexOf('srikandi') >= 0) return 'Asli';
  return '';
}

function extractKlasifikasiAkses_(text) {
  const str = String(text || '');
  const upper = str.toUpperCase();
  // Contextual patterns (higher confidence)
  const ctx = str.match(/(?:bersifat|klasifikasi\s*akses|tingkat\s*akses|sifat\s*dokumen)\s*[:.]?\s*(Rahasia|Terbatas|Biasa|Terbuka|Umum)/i);
  if (ctx) {
    const val = ctx[1].toLowerCase();
    if (val === 'rahasia') return 'Rahasia';
    if (val === 'terbatas') return 'Terbatas';
    return 'Terbuka';
  }
  // Keyword matching
  if (upper.indexOf('RAHASIA') >= 0) return 'Rahasia';
  if (upper.indexOf('TERBATAS') >= 0) return 'Terbatas';
  if (upper.indexOf('TERBUKA') >= 0 || upper.indexOf('UMUM') >= 0 || upper.indexOf('BIASA') >= 0) return 'Terbuka';
  return '';
}

function extractUraian_(text, sourceName, activity, subActivity) {
  const str = String(text || '');
  // Priority 1: "Perihal" or "Hal" — capture the subject line
  const perihal = str.match(/(?:Perihal|Hal)\s*[:.]?\s*(.+?)(?=\n(?:Ke(?:pada)?|Lampiran|Yth|Nomor|Tanggal|$)|\n\n|$)/is);
  if (perihal && perihal[1].trim().length > 3) return cleanUraian_(perihal[1]);
  // Fallback: simpler Perihal pattern
  const perihalSimple = str.match(/(?:Perihal|Hal)\s*[:.]?\s*(.+)/i);
  if (perihalSimple && perihalSimple[1].trim().length > 3) return cleanUraian_(perihalSimple[1]);
  // Priority 2: "Tentang" keyword (common in SK/SK)
  const tentang = str.match(/\b[Tt]entang\s+(.+?)(?=\n|$)/);
  if (tentang && tentang[1].trim().length > 3) return cleanUraian_(tentang[1]);
  // Priority 3: Clean filename
  const cleanedName = String(sourceName || '')
    .replace(/\.[a-z0-9]+$/i, '')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (cleanedName) return cleanUraian_(cleanedName);
  return ['Surat', activity.activity_name, subActivity.sub_activity_name].filter(Boolean).join(' - ');
}

function cleanUraian_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/^_+|_+$/g, '')
    .trim()
    .slice(0, MAX_URAIAN_LENGTH);
}

function buildFinalFileName_(metadata, sourceName) {
  const extMatch = String(sourceName || '').match(/\.[a-z0-9]+$/i);
  const ext = extMatch ? extMatch[0].toLowerCase() : '.pdf';
  const item = pad2_(metadata.nomor_item_arsip || '01');
  const tingkat = metadata.tingkat_perkembangan || 'Asli';
  
  const nomor = String(metadata.nomor_surat || '').trim();
  let uraianStr = String(metadata.uraian_informasi_item || '').trim();
  const kepada = String(metadata.kepada || '').trim();
  const dari = String(metadata.dari || '').trim();

  // Hitung panjang bagian tetap: nomor, kepada, dari, plus separator-separatornya
  // Format target: "[nomor] / [uraianStr]_Kepada [kepada]_Dari [dari]"
  let fixedLength = 0;
  if (nomor) {
    fixedLength += nomor.length;
  }
  if (kepada) {
    fixedLength += (nomor ? 3 : 0) + 8 + kepada.length; // " / " (jika no) + "_Kepada " + kepada
  }
  if (dari) {
    fixedLength += (nomor || kepada ? 1 : 0) + 5 + dari.length; // "_" + "Dari " + dari
  }

  // Budget untuk Uraian adalah total 170 dikurangi fixedLength
  const separatorLength = nomor ? 3 : 0; // " / "
  const maxUraianBudget = 170 - fixedLength - separatorLength;

  if (uraianStr.length > maxUraianBudget) {
    // Potong uraian agar total string pas 170. Sisakan minimal 15 karakter agar tidak kosong
    const limit = Math.max(15, maxUraianBudget - 3);
    uraianStr = uraianStr.slice(0, limit) + '...';
  }

  const baseParts = [];
  if (nomor) baseParts.push(nomor);
  if (uraianStr) baseParts.push(uraianStr);

  let assembled = baseParts.join(' / ');

  if (kepada) assembled += (assembled ? '_' : '') + 'Kepada ' + kepada;
  if (dari) assembled += (assembled ? '_' : '') + 'Dari ' + dari;

  if (!assembled) assembled = 'Dokumen Surat';

  const safeUraian = sanitizeFilePart_(assembled);

  return item + '. (' + tingkat + ') ' + safeUraian + ext;
}

// Pemetaan tingkat perkembangan dari skema LAMA (nama file arsip legacy) ke
// skema baru 2-nilai. Dipakai HANYA saat adopsi/scan file lama di init workspace.
//   Srikandi (digital)        -> Asli
//   Asli (fisik) / Copy / Cetak -> Salinan
// Kosong/asing -> '' (pemanggil isi default 'Asli').
function normalizeLegacyTingkat_(value) {
  const t = String(value || '').toLowerCase();
  if (t.indexOf('srikandi') >= 0) return 'Asli';
  if (t.indexOf('salinan') >= 0 || t.indexOf('copy') >= 0 || t.indexOf('cetak') >= 0 || t.indexOf('asli') >= 0) return 'Salinan';
  return '';
}

function parseExistingFileName_(fileName, defaultActivity, defaultSubActivity) {
  const meta = {
    nomor_item_arsip: '',
    no_berkas: resolveSubActivityArchiveNumber_(defaultSubActivity, defaultActivity),
    tingkat_perkembangan: '',
    nomor_surat: '',
    uraian_informasi_item: '',
    lokasi_simpan: fileName,
    kode_klasifikasi: '',
    klasifikasi_akses: 'Terbatas',
    jumlah: 1,
    satuan: 'Lembar',
    tanggal: '',
    no_filing_cabinet: defaultActivity ? defaultActivity.laci_no + '. Laci ' + defaultActivity.activity_name : '',
    _no_filing_cabinet_path: defaultActivity ? defaultActivity.laci_no + '. Laci ' + defaultActivity.activity_name : '',
    _no_filing_cabinet_url: (defaultActivity && defaultActivity.laci_folder_id) ? 'https://drive.google.com/drive/folders/' + defaultActivity.laci_folder_id : '',
    no_laci: defaultSubActivity ? (defaultSubActivity.sub_activity_name || '') : '',
    _no_laci_path: defaultSubActivity ? (defaultSubActivity.sub_activity_name || '') : '',
    _no_laci_url: (defaultSubActivity && defaultSubActivity.folder_id) ? 'https://drive.google.com/drive/folders/' + defaultSubActivity.folder_id : '',
    no_folder: '' // Akan diisi di bawah setelah nomor_item_arsip di-parse
  };

  const nameWithoutExt = fileName.replace(/\.[a-z0-9]+$/i, '').trim();
  
  const match = nameWithoutExt.match(/^(\d{1,4})\.\s*\(([^)]+)\)\s*(?:No:\s*([^]+?)_)?([^]+)$/);
  if (match) {
    meta.nomor_item_arsip = pad2_(match[1]);
    meta.tingkat_perkembangan = normalizeLegacyTingkat_(match[2].trim());
    if (match[3]) {
      meta.nomor_surat = match[3].trim();
    }
    meta.uraian_informasi_item = match[4].trim();
    // Paksa No Folder sinkron dengan Nomor Item Arsip
    meta.no_folder = meta.nomor_item_arsip;
  } else {
    meta.uraian_informasi_item = nameWithoutExt;
  }
  
  if (!meta.nomor_item_arsip) meta.nomor_item_arsip = '01';
  if (!meta.tingkat_perkembangan) meta.tingkat_perkembangan = 'Asli';
  
  // Fallback: Jika tidak ada nomor surat tapi ada teks uraian, coba ekstrak nomor surat dari uraian
  if (!meta.nomor_surat && meta.uraian_informasi_item) {
    const extNo = extractNomorSurat_(meta.uraian_informasi_item);
    // Guard (E3): jangan anggap token mirip tanggal/pecahan ("12/2025", "1/2")
    // sebagai nomor surat. Terima hanya bila ada huruf, atau punya >=2 pemisah,
    // atau memuat "Tahun".
    const looksLikeNomor = extNo && (/[A-Za-z]/.test(extNo) ||
      (extNo.match(/[\/\-]/g) || []).length >= 2 || /tahun/i.test(extNo));
    if (looksLikeNomor) {
      meta.nomor_surat = extNo;
      // Regex dinamis untuk menghapus nomor dari uraian. Escape dulu SEMUA
      // metakarakter regex (mis. '.'), lalu longgarkan HANYA pemisah - dan /.
      const escaped = extNo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const rawPattern = escaped.replace(/[-\/]/g, '[\\s\\_\\-\\/]+');
      const removeRe = new RegExp('(?:No(?:mor)?\\s*[:.\\-\\s_]+)?' + rawPattern, 'i');
      let cleanUraian = meta.uraian_informasi_item.replace(removeRe, '');
      // Bersihkan underscore dan spasi berlebih
      cleanUraian = cleanUraian.replace(/^_+|_+$/g, '').replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
      meta.uraian_informasi_item = cleanUraian;
    }
  }

  const dariMatch = meta.uraian_informasi_item.match(/(?:_|\b)Dari\s+(.+)$/i);
  if (dariMatch) {
    meta.dari = dariMatch[1].trim();
    meta.uraian_informasi_item = meta.uraian_informasi_item.substring(0, dariMatch.index).trim();
  }

  const kepMatch = meta.uraian_informasi_item.match(/(?:_|\b)Kepada\s+(.+)$/i);
  if (kepMatch) {
    meta.kepada = kepMatch[1].trim();
    meta.uraian_informasi_item = meta.uraian_informasi_item.substring(0, kepMatch.index).trim();
  }

  meta.uraian_informasi_item = meta.uraian_informasi_item.replace(/[\/_\s]+$/, '').replace(/\s+/g, ' ').trim();

  // Sinkronkan no_folder dengan nomor_item_arsip (seperti logic di form)
  meta.no_folder = meta.nomor_item_arsip;

  return meta;
}

MetadataService.parseExistingFileName = parseExistingFileName_;
