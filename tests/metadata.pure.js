// Mirrors PureFunctions.gs + MetadataService.gs extraction functions for Node.js unit testing.
// GAS .gs files share global scope at runtime but cannot be require()'d by Node.
// Keep in sync with: PureFunctions.gs, MetadataService.gs, ConfigHelpers.gs

function extractNomorSurat_(text) {
  var str = String(text || '');
  var patterns = [
    /(?:No(?:mor)?\.?)\s*[:.]?\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    /(?:No(?:mor)?\.?)\s*[:.]?\s*([A-Z0-9][A-Z0-9.\/\-]+\/[A-Z0-9.\/\-]+)/i,
    /\b([A-Z]{1,3}-\d{1,6}(?:\/[A-Z0-9.]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    /\b(\d{1,6}\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i,
    /(?:No(?:mor)?\.?)\s*[:.]?\s*(\d{1,6})\s+Tahun\s+(\d{4})/i,
    /(?:No(?:mor)?)\s*[:.\-\s_]+\s*([A-Z0-9][A-Z0-9.\-\s_]+?[\-\s_][12]\d{3})\b/i
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = str.match(patterns[i]);
    if (match) {
      if (i === 4) return match[1] + '/Tahun/' + match[2];
      if (i === 5) return match[1].replace(/[\s_]+/g, '/').trim();
      return match[1].replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function extractKodeKlasifikasi_(text) {
  var str = String(text || '');
  var ctx = str.match(/(?:Kode|Klasifikasi)\s*[:.]?\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}(?:\.\d{1,3})*)/i);
  if (ctx) return ctx[1].toUpperCase().trim();
  var raw = str.match(/\b([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)\b/);
  return raw ? raw[1].trim() : '';
}

var _MONTH_NAMES = {
  januari: '01', februari: '02', maret: '03', april: '04',
  mei: '05', juni: '06', juli: '07', agustus: '08',
  september: '09', oktober: '10', november: '11', desember: '12'
};
var _MONTH_RE = 'Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember';

function extractDate_(text) {
  var str = String(text || '');
  var ctxRe = new RegExp(
    '(?:tanggal|ditetapkan|ditandatangani)\\D{0,20}?' +
    '(?:' +
      '(\\d{1,2})\\s+(' + _MONTH_RE + ')\\s+(20\\d{2})' +
      '|(20\\d{2})[-/.](\\d{1,2})[-/.](\\d{1,2})' +
    ')', 'i');
  var ctx = str.match(ctxRe);
  if (ctx) {
    if (ctx[1] && ctx[2] && ctx[3]) {
      return [ctx[3], _MONTH_NAMES[ctx[2].toLowerCase()], pad2_(ctx[1])].join('-');
    }
    if (ctx[4] && ctx[5] && ctx[6]) {
      return [ctx[4], pad2_(ctx[5]), pad2_(ctx[6])].join('-');
    }
  }
  var iso = str.match(/\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/);
  if (iso) return [iso[1], pad2_(iso[2]), pad2_(iso[3])].join('-');
  var idRe = new RegExp('(?:[A-Za-z\\s]+,\\s*)?\\b(\\d{1,2})\\s+(' + _MONTH_RE + ')\\s+(20\\d{2})\\b', 'i');
  var id = str.match(idRe);
  if (id) return [id[3], _MONTH_NAMES[id[2].toLowerCase()], pad2_(id[1])].join('-');
  var myRe = new RegExp('\\b(' + _MONTH_RE + ')\\s+(20\\d{2})\\b', 'i');
  var my = str.match(myRe);
  if (my) return [my[2], _MONTH_NAMES[my[1].toLowerCase()], '01'].join('-');
  var dmy = str.match(/\b(\d{1,2})[.\-](\d{1,2})[.\-](20\d{2})\b/);
  if (dmy) return [dmy[3], pad2_(dmy[2]), pad2_(dmy[1])].join('-');
  return '';
}

function extractTingkatPerkembangan_(name) {
  const text = String(name || '').toLowerCase();
  if (text.indexOf('asli') >= 0) return 'Asli';
  if (text.indexOf('copy') >= 0) return 'Copy';
  if (text.indexOf('cetak') >= 0) return 'Cetak';
  if (text.indexOf('srikandi') >= 0) return 'Srikandi';
  return '';
}

function extractKlasifikasiAkses_(text) {
  var str = String(text || '');
  var upper = str.toUpperCase();
  var ctx = str.match(/(?:bersifat|klasifikasi\s*akses|tingkat\s*akses|sifat\s*dokumen)\s*[:.]?\s*(Rahasia|Terbatas|Biasa|Terbuka|Umum)/i);
  if (ctx) {
    var val = ctx[1].toLowerCase();
    if (val === 'rahasia') return 'Rahasia';
    if (val === 'terbatas') return 'Terbatas';
    return 'Biasa';
  }
  if (upper.indexOf('RAHASIA') >= 0) return 'Rahasia';
  if (upper.indexOf('TERBATAS') >= 0) return 'Terbatas';
  if (upper.indexOf('BIASA') >= 0 || upper.indexOf('TERBUKA') >= 0 || upper.indexOf('UMUM') >= 0) return 'Biasa';
  return '';
}

function extractUraian_(text, sourceName, activity, subActivity) {
  var str = String(text || '');
  var perihal = str.match(/(?:Perihal|Hal)\s*[:.]?\s*(.+?)(?=\n(?:Ke(?:pada)?|Lampiran|Yth|Nomor|Tanggal|$)|\n\n|$)/is);
  if (perihal && perihal[1].trim().length > 3) return cleanUraian_(perihal[1]);
  var perihalSimple = str.match(/(?:Perihal|Hal)\s*[:.]?\s*(.+)/i);
  if (perihalSimple && perihalSimple[1].trim().length > 3) return cleanUraian_(perihalSimple[1]);
  var tentang = str.match(/\b[Tt]entang\s+(.+?)(?=\n|$)/);
  if (tentang && tentang[1].trim().length > 3) return cleanUraian_(tentang[1]);
  var cleanedName = String(sourceName || '')
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
    .slice(0, 260);
}

function buildFinalFileName_(metadata, sourceName) {
  const extMatch = String(sourceName || '').match(/\.[a-z0-9]+$/i);
  const ext = extMatch ? extMatch[0].toLowerCase() : '.pdf';
  const item = pad2_(metadata.nomor_item_arsip || '01');
  const tingkat = metadata.tingkat_perkembangan || 'Srikandi';
  const nomor = metadata.nomor_surat || '';
  const uraian = sanitizeFilePart_(metadata.uraian_informasi_berkas || 'Dokumen Surat');

  if (nomor) {
    return item + '. (' + tingkat + ') No: ' + nomor + '_' + uraian + ext;
  }
  return item + '. (' + tingkat + ') ' + uraian + ext;
}

function sanitizeFilePart_(value) {
  return String(value || '')
    .replace(/[\\:*?"<>|]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 170);
}

function pad2_(value) {
  if (value === null || value === undefined) return '00';
  const text = String(value).trim();
  if (!text) return '00';
  return text.length === 1 ? '0' + text : text;
}

function isEmpty_(value) {
  return value === undefined || value === null || String(value).trim() === '';
}

function safeJsonParse_(str, fallback) {
  try { return JSON.parse(str); } catch (_) { return fallback; }
}

function cleanId_(value) {
  const text = String(value || '').trim();
  const match = text.match(/[-\w]{25,}/);
  return match ? match[0] : text;
}

function isTrue_(value) {
  return String(value).toUpperCase() === 'TRUE' || value === true;
}

function slug_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeSheetName_(value) {
  return String(value || 'Sheet')
    .replace(/[:\\/?*\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 90);
}

function parseExistingFileName_(fileName, defaultActivity, defaultSubActivity) {
  const meta = {
    nomor_item_arsip: '',
    no_berkas: '',
    tingkat_perkembangan: '',
    nomor_surat: '',
    uraian_informasi_berkas: '',
    lokasi_simpan: fileName,
    kode_klasifikasi: defaultSubActivity ? defaultSubActivity.default_kode_klasifikasi : '',
    klasifikasi_akses: 'Terbatas',
    jumlah: 1,
    satuan: 'Lembar',
    tanggal: '2026-06-08' // fixed for unit testing consistency
  };

  const nameWithoutExt = fileName.replace(/\.[a-z0-9]+$/i, '').trim();
  
  const match = nameWithoutExt.match(/^(\d+)\.\s*\(([^)]+)\)\s*(?:No:\s*([^]+?)_)?([^]+)$/);
  if (match) {
    meta.nomor_item_arsip = pad2_(match[1]);
    meta.no_berkas = String(Number(match[1]));
    meta.tingkat_perkembangan = match[2].trim();
    if (match[3]) {
      meta.nomor_surat = match[3].trim();
    }
    meta.uraian_informasi_berkas = match[4].trim();
  } else {
    // Try simple number prefix, e.g. "02. Surat Perintah" or "2. Surat Perintah"
    const simpleMatch = nameWithoutExt.match(/^(\d+)\.\s*(.+)$/);
    if (simpleMatch) {
      meta.nomor_item_arsip = pad2_(simpleMatch[1]);
      meta.no_berkas = String(Number(simpleMatch[1]));
      meta.uraian_informasi_berkas = simpleMatch[2].trim();
    } else {
      meta.uraian_informasi_berkas = nameWithoutExt;
    }
  }
  
  if (!meta.nomor_item_arsip) meta.nomor_item_arsip = '01';
  if (!meta.no_berkas) meta.no_berkas = '1';
  if (!meta.tingkat_perkembangan) meta.tingkat_perkembangan = 'Srikandi';

  return meta;
}

function escapeDriveQueryValue_(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

const DEFAULT_TEMPLATE_CATEGORY_COLOR = '#2563EB';

function normalizeHexColor_(value, fallback) {
  const candidate = String(value || '').trim().toUpperCase();
  if (/^#[0-9A-F]{6}$/.test(candidate)) return candidate;
  return String(fallback || DEFAULT_TEMPLATE_CATEGORY_COLOR || '#2563EB').trim().toUpperCase();
}

// ═══ ParseEngine testable functions (mirrored from ParseEngine.gs) ═══

var _PE_DOC_TYPE_KEYWORDS = {
  'Surat Keputusan': ['surat keputusan', 'keputusan kepala', 'keputusan direktur', 'keputusan ketua', 'menetapkan', 'mengingat', 'memutuskan', 'kesatu', 'kedua', 'ketiga'],
  'Surat Undangan': ['surat undangan', 'undangan', 'mengharapkan kehadiran', 'dimohon hadir', 'harap hadir', 'menghadiri'],
  'Surat Tugas': ['surat tugas', 'menugaskan', 'ditugaskan', 'bertugas', 'pelaksanaan tugas'],
  'Surat Edaran': ['surat edaran', 'edaran', 'untuk diketahui', 'untuk menjadi perhatian'],
  'Surat Perintah': ['surat perintah', 'memerintahkan', 'diperintahkan'],
  'Nota Dinas': ['nota dinas', 'memo dinas', 'catatan dinas'],
  'Berita Acara': ['berita acara', 'pada hari ini', 'bertanda tangan di bawah ini'],
  'Laporan': ['laporan', 'hasil laporan', 'laporan kegiatan', 'laporan bulanan', 'laporan tahunan'],
  'Surat Keterangan': ['surat keterangan', 'menerangkan bahwa', 'dengan ini menerangkan'],
  'Surat Permohonan': ['surat permohonan', 'memohon', 'permohonan', 'mengajukan permohonan'],
  'Surat Balasan': ['surat balasan', 'membalas surat', 'menanggapi surat', 'sehubungan dengan surat'],
  'Perjanjian': ['perjanjian', 'kesepakatan', 'pihak pertama', 'pihak kedua', 'pasal']
};

function classifyDocumentType_(text, fileName) {
  var combined = (String(text || '') + ' ' + String(fileName || '')).toLowerCase();
  var scores = {};
  var types = Object.keys(_PE_DOC_TYPE_KEYWORDS);
  for (var i = 0; i < types.length; i++) {
    var type = types[i];
    var keywords = _PE_DOC_TYPE_KEYWORDS[type];
    var score = 0;
    for (var j = 0; j < keywords.length; j++) {
      var count = 0;
      var pos = combined.indexOf(keywords[j]);
      while (pos >= 0) { count++; pos = combined.indexOf(keywords[j], pos + 1); }
      score += count * (j < 2 ? 3 : 1);
    }
    if (score > 0) scores[type] = score;
  }
  var bestType = '';
  var bestScore = 0;
  var keys = Object.keys(scores);
  for (var k = 0; k < keys.length; k++) {
    if (scores[keys[k]] > bestScore) {
      bestScore = scores[keys[k]];
      bestType = keys[k];
    }
  }
  return bestScore >= 3 ? bestType : '';
}

var _PE_OCR_FIXES = [
  [/[\u200B-\u200D\uFEFF]/g, ''],
  [/\r\n/g, '\n'],
  [/[ \t]{2,}/g, ' '],
  [/\u2018|\u2019/g, "'"],
  [/\u201C|\u201D/g, '"'],
  [/\u2013|\u2014/g, '-'],
  [/\u00A0/g, ' '],
  [/\u2026/g, '...'],
  [/[|]/g, 'I'],
  [/\u2022|\u25CF/g, '-']
];

function normalizeTextPE_(raw) {
  var text = String(raw || '');
  for (var i = 0; i < _PE_OCR_FIXES.length; i++) {
    text = text.replace(_PE_OCR_FIXES[i][0], _PE_OCR_FIXES[i][1]);
  }
  // Collapse 4+ consecutive newlines to 3
  text = text.replace(new RegExp('(?:' + String.fromCharCode(10) + '){4,}', 'g'), String.fromCharCode(10, 10, 10));
  return text;
}

module.exports = {
  extractNomorSurat_,
  extractKodeKlasifikasi_,
  extractDate_,
  extractTingkatPerkembangan_,
  extractKlasifikasiAkses_,
  extractUraian_,
  cleanUraian_,
  buildFinalFileName_,
  sanitizeFilePart_,
  pad2_,
  cleanId_,
  isTrue_,
  slug_,
  normalizeSheetName_,
  isEmpty_,
  safeJsonParse_,
  parseExistingFileName_,
  escapeDriveQueryValue_,
  normalizeHexColor_,
  classifyDocumentType_,
  normalizeTextPE_
};
