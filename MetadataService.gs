'use strict';

const MetadataService = {
  normalize: function (metadata, activity, subActivity, sourceName) {
    const normalized = Object.assign({}, metadata);
    
    const uraian = String(normalized.uraian_informasi_item || '').trim();
    const parts = [uraian];
    if (normalized.kepada && uraian.indexOf(normalized.kepada) === -1) parts.push(normalized.kepada);
    if (normalized.dari && uraian.indexOf(normalized.dari) === -1) parts.push(normalized.dari);
    
    // overwrite uraian_informasi_item so SpreadsheetService and buildFinalFileName uses it directly
    normalized.uraian_informasi_item = parts.filter(Boolean).join(' / ');
    
    if (normalized.nomor_item_arsip) {
      const parsedNum = Number(normalized.nomor_item_arsip);
      if (!isNaN(parsedNum)) {
        normalized.no_berkas = String(parsedNum);
      } else {
        normalized.no_berkas = String(normalized.nomor_item_arsip).trim();
      }
      normalized.nomor_item_arsip = String(normalized.nomor_item_arsip).padStart(2, '0');
    } else {
      normalized.no_berkas = normalized.no_berkas || (subActivity && subActivity.sort_order ? String(subActivity.sort_order) : '');
    }

    normalized.no_laci = normalized.no_laci || '';
    normalized.no_folder = normalized.no_folder || '';
    normalized.no_filing_cabinet = normalized.no_filing_cabinet || '';
    normalized.satuan = normalized.satuan || 'Lembar';
    normalized.tingkat_perkembangan = normalized.tingkat_perkembangan || 'Asli';
    normalized.jumlah = metadata.jumlah !== undefined ? metadata.jumlah : 1;
    normalized.klasifikasi_akses = normalized.klasifikasi_akses || 'Terbatas';
    normalized.lokasi_simpan = normalized.lokasi_simpan || this.buildFinalFileName(normalized, sourceName);
    return normalized;
  },

  createDraft: function (payload, activity, subActivity, fields) {
    const sourceName = payload.sourceFileName || '';
    const rawText = payload.rawText || sourceName;
    const text = [sourceName, rawText].join('\n');
    const nomorSurat = extractNomorSurat_(text);
    const tanggal = extractDate_(text) || payload.fileLastUpdatedStr || '';
    const tingkat = extractTingkatPerkembangan_(sourceName) || 'Asli';
    const kode = extractKodeKlasifikasi_(text) || '';
    const uraian = extractUraian_(text, sourceName, activity, subActivity);

    const itemVal = String(payload.nomorItemArsip || payload.noBerkas || '').replace(/^0+/, '');
    
    const base = {
      no_berkas: itemVal,
      nomor_item_arsip: itemVal ? itemVal.padStart(2, '0') : '',
      kode_klasifikasi: kode,
      // Komponen uraian: perihal terisi hasil ekstraksi; kepada/dari diisi user.
      perihal: uraian,
      kepada: '',
      dari: '',
      uraian_informasi_item: uraian,
      tanggal: tanggal,
      tingkat_perkembangan: tingkat,
      jumlah: 1,
      satuan: 'Lembar',
      no_filing_cabinet: payload.noFilingCabinet || '',
      no_laci: payload.noLaci || '',
      no_folder: payload.noFolder || '',
      klasifikasi_akses: 'Terbatas',
      ket: '',
      lokasi_simpan: ''
    };

    fields.forEach(field => {
      if ((base[field.field_name] === '' || base[field.field_name] === undefined) && field.default_value !== '') {
        base[field.field_name] = field.default_value;
      }
    });

    base.nomor_surat = nomorSurat;
    // Uraian = nomor surat/perihal/kepada/dari (bagian kosong dilewati).
    base.uraian_informasi_item = assembleUraian_(base) || uraian;
    base.lokasi_simpan = buildFinalFileName_(base, sourceName);

    return {
      metadata: base,
      confidence: {
        nomor_surat: nomorSurat ? 'medium' : 'low',
        tanggal: tanggal ? 'medium' : 'low',
        kode_klasifikasi: kode ? 'medium' : 'low',
        uraian: uraian ? 'medium' : 'low'
      },
      notes: [
        'Draft ini hasil bantuan parsing sederhana dari nama file/teks tempel.',
        'User tetap wajib review sebelum file disimpan dan spreadsheet diisi.'
      ]
    };
  },

  buildFinalFileName: buildFinalFileName_
};

function extractNomorSurat_(text) {
  const str = String(text || '');
  const patterns = [
    // "Nomor : B-123/DL.01/2026" or "No. 123/ABC/2026" (most explicit, highest priority)
    /(?:No(?:mor)?\.?)\s*[:.]?\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    // Letter prefix: "B-123/DL.01/2026", "SP-456/KP.02/2025"
    /\b([A-Z]{1,3}-\d{1,6}(?:\/[A-Z0-9.]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    // Pure numeric segments: "123/DL.01/2026", "01/02/03/2026"
    /\b(\d{1,6}\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    // Formal: "Nomor 45 Tahun 2025"
    /(?:No(?:mor)?\.?)\s*[:.]?\s*(\d{1,6})\s+Tahun\s+(\d{4})/i,
    // Space/dash separated filename format
    /(?:No(?:mor)?)\s*[:.\-\s_]+\s*([A-Z0-9][A-Z0-9.\-\s_]+?[\-\s_][12]\d{3})\b/i
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = str.match(patterns[i]);
    if (match) {
      if (i === 3) return match[1] + '/Tahun/' + match[2];
      if (i === 4) return match[1].replace(/[\s_]+/g, '/').trim();
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
    return 'Biasa';
  }
  // Keyword matching
  if (upper.indexOf('RAHASIA') >= 0) return 'Rahasia';
  if (upper.indexOf('TERBATAS') >= 0) return 'Terbatas';
  if (upper.indexOf('BIASA') >= 0 || upper.indexOf('TERBUKA') >= 0 || upper.indexOf('UMUM') >= 0) return 'Biasa';
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

// Rakit kolom Uraian dari komponen: nomor surat/perihal/kepada/dari.
// Bagian kosong dilewati supaya tak ada garis miring dobel.
function assembleUraian_(metadata) {
  metadata = metadata || {};
  return ['nomor_surat', 'perihal', 'kepada', 'dari']
    .map(function (k) { return String(metadata[k] || '').trim(); })
    .filter(Boolean)
    .join('/');
}

function buildFinalFileName_(metadata, sourceName) {
  const extMatch = String(sourceName || '').match(/\.[a-z0-9]+$/i);
  const ext = extMatch ? extMatch[0].toLowerCase() : '.pdf';
  const item = pad2_(metadata.nomor_item_arsip || '01');
  const tingkat = metadata.tingkat_perkembangan || 'Asli';
  const nomor = metadata.nomor_surat || '';
  const uraian = sanitizeFilePart_(metadata.uraian_informasi_item || 'Dokumen Surat');

  if (nomor) {
    return item + '. (' + tingkat + ') No: ' + nomor + '_' + uraian + ext;
  }
  return item + '. (' + tingkat + ') ' + uraian + ext;
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
    no_berkas: '',
    tingkat_perkembangan: '',
    nomor_surat: '',
    uraian_informasi_item: '',
    lokasi_simpan: fileName,
    kode_klasifikasi: typeof DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI !== 'undefined' ? DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI : 'PDP.07.1',
    klasifikasi_akses: 'Terbatas',
    jumlah: 1,
    satuan: 'Lembar',
    tanggal: ''
  };

  const nameWithoutExt = fileName.replace(/\.[a-z0-9]+$/i, '').trim();
  
  const match = nameWithoutExt.match(/^(\d{1,4})\.\s*\(([^)]+)\)\s*(?:No:\s*([^]+?)_)?([^]+)$/);
  if (match) {
    meta.nomor_item_arsip = pad2_(match[1]);
    meta.no_berkas = String(Number(match[1]));
    meta.tingkat_perkembangan = normalizeLegacyTingkat_(match[2].trim());
    if (match[3]) {
      meta.nomor_surat = match[3].trim();
    }
    meta.uraian_informasi_item = match[4].trim();
  } else {
    meta.uraian_informasi_item = nameWithoutExt;
  }
  
  if (!meta.nomor_item_arsip) meta.nomor_item_arsip = '01';
  if (!meta.no_berkas) meta.no_berkas = '1';
  if (!meta.tingkat_perkembangan) meta.tingkat_perkembangan = 'Asli';

  return meta;
}

MetadataService.parseExistingFileName = parseExistingFileName_;
