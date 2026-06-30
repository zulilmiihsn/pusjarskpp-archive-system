'use strict';

/**
 * ParseEngine — Advanced metadata extraction engine for Indonesian archival documents.
 *
 * Multi-pass architecture:
 *   Pass 1: Pre-process raw OCR text (normalize, clean artifacts)
 *   Pass 2: Analyze document structure (sections, zones)
 *   Pass 3: Classify document type (SK, Undangan, Tugas, etc.)
 *   Pass 4: Extract fields with scored pattern matching
 *   Pass 5: Rank, validate, and return best results
 *
 * Each extractor returns candidates with scores. Best candidate is chosen by:
 *   score = patternWeight * keywordBonus * positionBonus * formatSpecificity
 *
 * Integration: ArchiveController.parseDocumentContent calls ParseEngine.analyze()
 * instead of calling individual extractors directly.
 */
const ParseEngine = (function () {

  // ═══════════════════════════════════════════════════════════════
  //  CONSTANTS
  // ═══════════════════════════════════════════════════════════════

  const INDONESIAN_MONTHS = {
    januari: '01', feb: '02', februari: '02', mar: '03', maret: '03',
    apr: '04', april: '04', mei: '05', may: '05', mel: '05', me1: '05', jun: '06', juni: '06',
    jul: '07', juli: '07', agu: '08', agustus: '08', agst: '08', aug: '08', august: '08',
    sep: '09', sept: '09', september: '09', okt: '10', oktober: '10', oct: '10', october: '10',
    nov: '11', november: '11', des: '12', desember: '12', dec: '12', december: '12'
  };

  const MONTH_FULL = 'Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember|January|February|March|August|October|December';
  const MONTH_ABBR = 'Jan|Feb|Mar|Apr|Mei|May|Mel|Me1|Jun|Jul|Agu|Agst|Aug|Sep|Sept|Okt|Oct|Nov|Des|Dec';
  const MONTH_ALL = MONTH_FULL + '|' + MONTH_ABBR;

  // Common OCR character confusions
  const _NL = String.fromCharCode(10);
  const OCR_FIXES = [
    [/[\u200B-\u200D\uFEFF]/g, ''],          // zero-width chars
    [/\r\n/g, _NL],                            // normalize line endings
    [/[ \t]{2,}/g, ' '],                        // collapse horizontal whitespace
    [new RegExp('(?:' + _NL + '){4,}', 'g'), _NL + _NL + _NL],  // max 3 consecutive newlines
    [/\u2018|\u2019/g, "'"],                    // smart single quotes
    [/\u201C|\u201D/g, '"'],                    // smart double quotes
    [/\u2013|\u2014/g, '-'],                    // en/em dashes
    [/\u00A0/g, ' '],                           // non-breaking space
    [/\u2026/g, '...'],                         // ellipsis
    [/[|]/g, 'I'],                              // pipe → I
    [/\u2022|\u25CF/g, '-'],                    // bullets → dash
  ];

  // Known Indonesian government letter number patterns
  const NOMOR_FORMATS = [
    { name: 'SK format', re: /(?:SK|Surat\s+Keputusan)[\s/.:-]+(\d+\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.95 },
    { name: 'SP format', re: /(?:SP|Surat\s+Perintah)[\s/.:-]+([A-Z]?\d{1,6}[\-\/][A-Z0-9.\/\-]+\/[12]\d{3})/i, weight: 0.95 },
    { name: 'SE format', re: /(?:SE|Surat\s+Edaran)[\s/.:-]+(\d+\/[A-Z0-9.\/\-]+\/[12]\d{3})/i, weight: 0.95 },
    { name: 'explicit Nomor', re: /(?:No(?:mor)?)\s*[:.]\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.9 },
    { name: 'explicit No', re: /(?:No(?:mor)?)\s*[:.]?\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.85 },
    { name: 'letter prefix', re: /\b([A-Z]{1,4}[\-]\d{1,6}(?:\/[A-Z0-9.]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)\b/i, weight: 0.8 },
    { name: 'numeric segments', re: /\b(\d{1,6}\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)\b/i, weight: 0.7 },
    { name: 'Nomor Tahun', re: /(?:No(?:mor)?)\s*[:.]?\s*(\d{1,6})\s+Tahun\s+(20[12]\d)/i, weight: 0.75 },
    { name: 'slash year', re: /\b(\d{1,5}\/[A-Z]{1,4}[A-Z0-9.]*\/[12]\d{3})\b/i, weight: 0.6 }
  ];

  // Document type keywords
  const DOC_TYPE_KEYWORDS = {
    'Surat Keputusan': ['surat keputusan', 'keputusan kepala', 'keputusan direktur', 'keputusan ketua', 'keputusan rektor', 'menetapkan', 'mengingat', 'memutuskan', 'kesatu', 'kedua', 'ketiga'],
    'Surat Undangan': ['surat undangan', 'undangan', 'mengharapkan kehadiran', 'dimohon hadir', 'harap hadir', 'menghadiri'],
    'Surat Edaran': ['surat edaran', 'edaran', 'untuk diketahui', 'untuk menjadi perhatian'],
    'Surat Perintah': ['surat perintah', 'memerintahkan', 'diperintahkan'],
    'Nota Dinas': ['nota dinas'],
    'Berita Acara': ['berita acara', 'pada hari ini', 'bertanda tangan di bawah ini'],
    'Laporan': ['laporan', 'hasil laporan', 'laporan kegiatan', 'laporan bulanan', 'laporan tahunan'],
    'Surat Keterangan': ['surat keterangan', 'menerangkan bahwa', 'dengan ini menerangkan'],
    'Surat Permohonan': ['surat permohonan', 'memohon', 'permohonan', 'mengajukan permohonan'],
    'Surat Balasan': ['surat balasan', 'membalas surat', 'menanggapi surat', 'sehubungan dengan surat'],
    'Perjanjian': ['perjanjian', 'kesepakatan', 'pihak pertama', 'pihak kedua', 'pasal']
  };

  // Klasifikasi akses phrase patterns (ordered by specificity)
  const AKSES_PATTERNS = [
    { re: /(?:bersifat|sifat)\s*[:.]?\s*rahasia/i, value: 'Rahasia', score: 0.95 },
    { re: /(?:bersifat|sifat)\s*[:.]?\s*terbatas/i, value: 'Terbatas', score: 0.9 },
    { re: /(?:bersifat|sifat)\s*[:.]?\s*(?:biasa|umum|terbuka)/i, value: 'Biasa', score: 0.9 },
    { re: /klasifikasi\s*(?:akses)?\s*[:.]?\s*rahasia/i, value: 'Rahasia', score: 0.95 },
    { re: /klasifikasi\s*(?:akses)?\s*[:.]?\s*terbatas/i, value: 'Terbatas', score: 0.9 },
    { re: /klasifikasi\s*(?:akses)?\s*[:.]?\s*(?:biasa|umum|terbuka)/i, value: 'Biasa', score: 0.9 },
    { re: /tingkat\s*akses\s*[:.]?\s*rahasia/i, value: 'Rahasia', score: 0.9 },
    { re: /tingkat\s*akses\s*[:.]?\s*terbatas/i, value: 'Terbatas', score: 0.85 },
    { re: /tidak\s+untuk\s+(?:disebarluaskan|umum|publik)/i, value: 'Terbatas', score: 0.7 },
    { re: /untuk\s+kalangan\s+terbatas/i, value: 'Terbatas', score: 0.75 },
    { re: /dokumen\s+(?:negara|resmi)\s+rahasia/i, value: 'Rahasia', score: 0.8 },
    { re: /\bRAHASIA\b/, value: 'Rahasia', score: 0.5 },
    { re: /\bTERBATAS\b/, value: 'Terbatas', score: 0.45 },
    { re: /\bBIASA\b|\bTERBUKA\b|\bUMUM\b/, value: 'Biasa', score: 0.4 }
  ];

  // Kode klasifikasi patterns (contextual → raw)
  const KODE_PATTERNS = [
    { re: /(?:kode|kode\s+klasifikasi)\s*[:.]?\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?)/i, weight: 0.95 },
    { re: /klasifikasi\s*[:.]?\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?)/i, weight: 0.9 },
    { re: /\b([A-Z]{2,4}\.\d{2}\.\d{1,2})\b/, weight: 0.8 },
    { re: /\b([A-Z]{2,4}\.\d{2})\b/, weight: 0.65 },
    { re: /\b([A-Z]\.\d{2}(?:\.\d{1,2})?)\b/, weight: 0.5 }
  ];

  // Uraian section-boundary keywords (where Perihal value stops)
  const URAIAN_STOP = '\\b(?:Ke(?:pada)?|Lampiran|Yth|Nomor|Tanggal|Perihal|Hal|Lamp|Isi|Dengan)\\b';

  // Kamus Singkatan (Alias) untuk instansi & jabatan
  const INSTITUTION_ALIASES = [
    { re: /\bpusat\s+pembelajaran\s+dan\s+strategi\s+kebijakan\s+(?:pengembangan\s+kompetensi|pelayanan\s+publik)?\b/ig, replace: 'Pusjar SKPP' },
    { re: /\blembaga\s+administrasi\s+negara(?:\s+republik\s+indonesia)?\b/ig, replace: 'LAN RI' },
    { re: /\bkajian\s+manajemen\s+pemerintahan\b/ig, replace: 'KMP' },
    { re: /\bkajian\s+hukum\s+administrasi\s+negara\b/ig, replace: 'KHAN' },
    { re: /\bpengembangan\s+kompetensi\s+dan\s+pemetaan\s+(?:asn|aparatur\s+sipil\s+negara)\b/ig, replace: 'PKASN' },
    { re: /\bpusat\s+inovasi\s+administrasi\s+negara\b/ig, replace: 'PIAN' },
    { re: /\bpusat\s+pembinaan\s+analis\s+kebijakan\b/ig, replace: 'Pusbin AK' },
    { re: /\bpusat\s+data\s+dan\s+informasi\b/ig, replace: 'Pusdatin' },
    { re: /\bpusat\s+kajian\s+dan\s+pendidikan\s+dan\s+pelatihan\s+aparatur\b/ig, replace: 'PKP2A' },
    { re: /\bsekolah\s+tinggi\s+ilmu\s+administrasi\b/ig, replace: 'STIA' },
    { re: /\bpusat\s+pelatihan\s+dasar\s+dan\s+core\s+values\s+asn\b/ig, replace: 'Puslatsar' },
    { re: /\bsekretaris\s+jenderal\b/ig, replace: 'Sekjen' },
    { re: /\bsekretaris\s+daerah\b/ig, replace: 'Sekda' },
    { re: /\bpemerintah\s+provinsi\b/ig, replace: 'Pemprov' },
    { re: /\bpemerintah\s+kabupaten\b/ig, replace: 'Pemkab' },
    { re: /\bpemerintah\s+kota\b/ig, replace: 'Pemkot' },
    { re: /\bkementerian\s+pendayagunaan\s+aparatur\s+negara\s+dan\s+reformasi\s+birokrasi\b/ig, replace: 'KemenPAN-RB' },
    { re: /\bbadan\s+kepegawaian\s+negara\b/ig, replace: 'BKN' }
  ];

  function applyAliases_(text) {
    let result = String(text || '');
    for (let i = 0; i < INSTITUTION_ALIASES.length; i++) {
      result = result.replace(INSTITUTION_ALIASES[i].re, INSTITUTION_ALIASES[i].replace);
    }
    return result;
  }

  // ═══════════════════════════════════════════════════════════════
  //  TEXT PREPROCESSOR
  // ═══════════════════════════════════════════════════════════════

  function normalizeText_(raw) {
    let text = String(raw || '');
    for (let i = 0; i < OCR_FIXES.length; i++) {
      text = text.replace(OCR_FIXES[i][0], OCR_FIXES[i][1]);
    }
    return text;
  }

  function getLines_(text) {
    return String(text || '').split('\n');
  }

  function findKeywordLine_(text, keywords) {
    const lines = getLines_(text);
    for (let i = 0; i < lines.length; i++) {
      const lower = lines[i].toLowerCase();
      for (let j = 0; j < keywords.length; j++) {
        if (lower.indexOf(keywords[j]) >= 0) {
          return { lineIndex: i, line: lines[i], context: getContext_(lines, i, 3) };
        }
      }
    }
    return null;
  }

  function getContext_(lines, index, radius) {
    const start = Math.max(0, index - radius);
    const end = Math.min(lines.length - 1, index + radius);
    return lines.slice(start, end + 1).join('\n');
  }

  // ═══════════════════════════════════════════════════════════════
  //  DOCUMENT STRUCTURE ANALYZER
  // ═══════════════════════════════════════════════════════════════

  function analyzeStructure_(text) {
    const lines = getLines_(text);
    const totalLines = lines.length;
    const structure = {
      header: { start: 0, end: Math.min(totalLines, Math.ceil(totalLines * 0.3)), lines: [] },
      body: { start: 0, end: totalLines, lines: [] },
      signature: { start: 0, end: totalLines, lines: [] },
      footer: { start: 0, end: totalLines, lines: [] }
    };

    // Find signature block
    for (let i = totalLines - 1; i >= Math.floor(totalLines * 0.6); i--) {
      const lower = lines[i].toLowerCase();
      if (lower.indexOf('hormat') >= 0 || lower.indexOf('tanda tangan') >= 0 ||
          lower.indexOf('kepala') >= 0 || lower.indexOf('direktur') >= 0 ||
          lower.indexOf('ketua') >= 0 || lower.indexOf('sekretaris') >= 0 ||
          /ditetapkan\s+di/i.test(lines[i]) || /ditandatangani/i.test(lines[i])) {
        structure.signature = { start: i, end: totalLines, lines: lines.slice(i) };
        break;
      }
    }

    // Find header (first 35% or until body/salutation markers are encountered)
    let headerEnd = 0;
    for (let j = 0; j < Math.min(totalLines, Math.ceil(totalLines * 0.35)); j++) {
      const line = lines[j].trim();
      if (/^(?:yth|kepada|kpd|dengan\s+hormat|menindaklanjuti|sehubungan|merujuk|berdasarkan)\b/i.test(line)) {
        break;
      }
      if (/nomor|no\.|perihal|hal\s*:/i.test(line)) {
        headerEnd = j + 1;
      }
    }
    if (headerEnd === 0) {
      let salutationIndex = -1;
      for (let j = 0; j < Math.min(totalLines, Math.ceil(totalLines * 0.35)); j++) {
        if (/^(?:yth|kepada|kpd|dengan\s+hormat|menindaklanjuti|sehubungan|merujuk|berdasarkan)\b/i.test(lines[j].trim())) {
          salutationIndex = j;
          break;
        }
      }
      headerEnd = salutationIndex > 0 ? salutationIndex : Math.min(totalLines, 10);
    }
    structure.header.end = headerEnd;
    structure.header.lines = lines.slice(0, headerEnd);

    // Body is between header and signature. Bila blok tanda tangan TAK terdeteksi,
    // signature.start tetap 0 (default) — jangan pakai itu sebagai batas, sebab body
    // akan kolaps jadi kosong (slice(headerEnd, 0)). Saat tak ada tanda tangan,
    // body membentang sampai akhir dokumen.
    const bodyEnd = structure.signature.lines.length ? structure.signature.start : totalLines;
    structure.body.start = headerEnd;
    structure.body.end = bodyEnd;
    structure.body.lines = lines.slice(headerEnd, bodyEnd);

    // Footer is last 10% (tembusan, etc.)
    const footerStart = Math.max(structure.signature.start, Math.floor(totalLines * 0.9));
    structure.footer.start = footerStart;
    structure.footer.lines = lines.slice(footerStart);

    return structure;
  }

  // ═══════════════════════════════════════════════════════════════
  //  DOCUMENT TYPE CLASSIFIER
  // ═══════════════════════════════════════════════════════════════

  function classifyDocumentType_(text, fileName) {
    const combined = (String(text || '') + ' ' + String(fileName || '')).toLowerCase();
    const scores = {};
    const types = Object.keys(DOC_TYPE_KEYWORDS);

    for (let i = 0; i < types.length; i++) {
      const type = types[i];
      const keywords = DOC_TYPE_KEYWORDS[type];
      let score = 0;
      for (let j = 0; j < keywords.length; j++) {
        let count = 0;
        let pos = combined.indexOf(keywords[j]);
        while (pos >= 0) { count++; pos = combined.indexOf(keywords[j], pos + 1); }
        score += count * (j < 2 ? 3 : 1);  // first 2 keywords weighted more
      }
      if (score > 0) scores[type] = score;
    }

    let bestType = '';
    let bestScore = 0;
    const keys = Object.keys(scores);
    for (let k = 0; k < keys.length; k++) {
      if (scores[keys[k]] > bestScore) {
        bestScore = scores[keys[k]];
        bestType = keys[k];
      }
    }

    return bestScore >= 3 ? bestType : '';
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: NOMOR SURAT
  // ═══════════════════════════════════════════════════════════════

  function extractNomorSuratScored_(text, structure, fileName) {
    const candidates = [];
    const headerText = structure.header.lines.join('\n');

    // STRICT keyword-anchored patterns only
    const patterns = [
      // Nomor label on one line, actual number on the next line (often with templates or dates on same line as Nomor label)
      { re: /(?:No(?:mor)?)\s*[:.\s]+.*?\n\s*([A-Z0-9][A-Z0-9.\/\-]{3,}\/[A-Z0-9.\/\-]+)/i, score: 0.97 },
      // Explicit "Nomor:" or "No:" prefix with Year
      { re: /(?:No(?:mor)?)\s*[:.\s]+\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)*\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, score: 0.95 },
      // Explicit "Nomor:" or "No:" prefix without Year
      { re: /(?:No(?:mor)?)\s*[:.\s]+\s*([A-Z0-9][A-Z0-9.\/\-]+\/[A-Z0-9.\/\-]+)/i, score: 0.85 },
      // "Nomor XX Tahun YYYY" format
      { re: /(?:No(?:mor)?)\s*[:.\s]+\s*(\d{1,6})\s+Tahun\s+(20[12]\d)/i, score: 0.9 },
      // SK/SP/ST/SE prefix
      { re: /\b(?:SK|SP|ST|SE)[\s\/\-:.]+(\d+[\-\/][A-Z0-9.\/\-]+(?:\/[12]\d{3})?(?:\/[A-Z0-9.\-]+)?)/i, score: 0.88 },
      // Surat Keputusan/Perintah/Tugas/Edaran prefix
      { re: /(?:Surat\s+(?:Keputusan|Perintah|Tugas|Edaran))[\s\/\-:.]+([A-Z0-9][A-Z0-9.\/\-]+(?:\/[12]\d{3})?(?:\/[A-Z0-9.\-]+)?)/i, score: 0.88 },
      // Space/dash separated filename number (e.g. "No B 417 BKPSDM 800.2 12 2025")
      { re: /(?:No(?:mor)?)\s*[:.\-\s_]+\s*([A-Z0-9][A-Z0-9.\-\s_]+?[\-\s_][12]\d{3})\b/i, score: 0.85, isSpaceFormat: true },
      // Bare number on a line by itself in header (e.g. following a template placeholder line)
      { re: /(?:^|\n)\s*([A-Z0-9][A-Z0-9.\/\-]{3,}\/[A-Z0-9.\/\-]+)\s*(?:\n|$)/i, score: 0.82 }
    ];

    // Only search in header, first 30% of document, and fileName
    const searchZones = [
      { text: headerText, bonus: 1.3 },
      { text: text.substring(0, Math.floor(text.length * 0.3)), bonus: 1.0 },
      { text: String(fileName || ''), bonus: 1.5 }
    ];

    for (let z = 0; z < searchZones.length; z++) {
      const zone = searchZones[z];
      for (let i = 0; i < patterns.length; i++) {
        const p = patterns[i];
        const match = zone.text.match(p.re);
        if (match && match[1]) {
          let value = match[1].replace(/\s+/g, ' ').trim();
          // Validate: must have at least one slash, or be "XX Tahun YYYY", or be a space format
          if (value.indexOf('/') >= 0 || /\d+\s+Tahun/i.test(value) || p.isSpaceFormat) {
            if (p.isSpaceFormat) value = value.replace(/[\s_]+/g, '/');
            candidates.push({ value: value, score: p.score * zone.bonus, source: 'keyword_nomor_' + i, zone: z });
          }
        }
      }
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: KODE KLASIFIKASI
  // ═══════════════════════════════════════════════════════════════

  function extractKodeKlasifikasiScored_(text, structure) {
    const candidates = [];
    const headerText = structure.header.lines.join('\n');

    // STRICT keyword-anchored patterns only
    const patterns = [
      // Explicit "Kode:" or "Klasifikasi:" prefix
      { re: /(?:Kode|Klasifikasi)\s*[:.\s]+\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)/i, score: 0.95 },
      // "Kode Klasifikasi:" prefix
      { re: /Kode\s+Klasifikasi\s*[:.\s]+\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)/i, score: 0.98 },
      // Bare code in header (must have a dot for numeric to prevent matching random 3 digits like 595)
      { re: /\b([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)\b/i, score: 0.7 }
    ];

    // Only search in header
    const searchZones = [
      { text: headerText, bonus: 1.3 }
    ];

    for (let z = 0; z < searchZones.length; z++) {
      const zone = searchZones[z];
      for (let i = 0; i < patterns.length; i++) {
        const p = patterns[i];
        const match = zone.text.match(p.re);
        if (match && match[1]) {
          const value = match[1].toUpperCase().trim();
          if (/^[A-Z]{1,4}\.\d{2}/.test(value) || /^\d{3}/.test(value)) {
            candidates.push({ value: value, score: p.score * zone.bonus, source: 'keyword_kode_' + i, zone: z });
          }
        }
      }
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: TANGGAL
  // ═══════════════════════════════════════════════════════════════

  function extractTanggalScored_(text, structure) {
    const candidates = [];
    const headerText = structure.header.lines.join('\n');
    const sigText = structure.signature.lines.join('\n');
    const bodyText = structure.body.lines.join('\n');

    // STRICT keyword-anchored patterns only
    const patterns = [
      // "Tanggal DD MMMM YYYY" or "Ditetapkan pada DD MMMM YYYY"
      { re: new RegExp('(?:Tanggal|Ditetapkan|Ditandatangani|Berangka)[\\s:.,]+(?:[^\\n]{0,20}?)(\\d{1,2})\\s+(' + MONTH_ALL + ')\\s+(20[12]\\d)', 'i'), score: 0.95 },
      // Bare date with City prefix like "Samarinda,      // Standard ID format
      { re: new RegExp('(?:[A-Za-z\\s]+,\\s*)?(\\d{1,2})\\s+(' + MONTH_ALL + ')\\s+(20[12]\\d)', 'i'), score: 0.85, bareDate: true },
      // Contextual ISO format
      { re: /(?:Tanggal|Ditetapkan|Ditandatangani)\s*[:.,]?\s*(?:[^\n]{0,10}?)(20[12]\d)[\-\/.](\d{1,2})[\-\/.](\d{1,2})/i, score: 0.9 },
      // Dateline kota di AWAL BARIS: "Samarinda, 13 Desember 2025". Diizinkan di body
      // (bukan bareDate) karena pola "Kota, tanggal" di awal baris hampir pasti dateline
      // surat, bukan tanggal yang sekadar disebut di tengah kalimat isi.
      { re: new RegExp('(?:^|\\n)[ \\t]*[A-Za-z][A-Za-z. ]{1,24},[ \\t]*(\\d{1,2})\\s+(' + MONTH_ALL + ')\\s+(20[12]\\d)', 'i'), score: 0.8 }
    ];

    // Header & signature = zona utama dateline surat (bonus tinggi). Body = cadangan
    // bonus rendah supaya dateline asli selalu menang, TAPI tetap menerima tanggal
    // polos: pada surat yang tanda tangannya berupa GAMBAR (tak terdeteksi sebagai
    // blok signature), dateline jatuh ke body sebagai tanggal polos. match() ambil
    // kemunculan PERTAMA + getPositionBonus_ memihak posisi teratas, jadi dateline di
    // atas tetap diutamakan ketimbang tanggal yang disebut jauh di dalam isi.
    const searchZones = [
      { text: headerText, bonus: 1.8 },
      { text: sigText, bonus: 1.0 },
      { text: bodyText, bonus: 0.8 }
    ];

    for (let z = 0; z < searchZones.length; z++) {
      const zone = searchZones[z];
      for (let i = 0; i < patterns.length; i++) {
        const p = patterns[i];
        if (zone.keywordOnly && p.bareDate) continue;
        const match = zone.text.match(p.re);
        if (match) {
          let dateVal = null;
          if (match[1] && match[2] && match[3]) {
            if (match[1].length === 4) {
              dateVal = [match[1], pad2_(match[2]), pad2_(match[3])].join('-');
            } else {
              const m = resolveMonth_(match[2]);
              if (m) dateVal = [match[3], m, pad2_(match[1])].join('-');
            }
          }
          if (dateVal) {
            const indexInFullText = text.indexOf(match[0]);
            const posBonus = getPositionBonus_(text, indexInFullText);
            candidates.push({ value: dateVal, score: p.score * zone.bonus * posBonus, source: 'keyword_tanggal_' + i, zone: z });
          }
        }
      }
    }

    return rankCandidates_(candidates, true);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: URAIAN
  // ═══════════════════════════════════════════════════════════════

  function extractUraianScored_(text, structure, fileName, activity, subActivity) {
    const candidates = [];
    const upperHalfText = text.substring(0, Math.floor(text.length * 0.5));

    // STRICT keyword-anchored patterns only
    const patterns = [
      // Multi-line Perihal in header
      { re: new RegExp('(?:Perihal|Hal)\\s*[:.]?\\s*(.+?)(?=\\n' + URAIAN_STOP + '|\\n\\n|$)', 'is'), score: 0.95, zone: upperHalfText },
      // Simple Perihal/Hal on same line in header
      { re: /(?:Perihal|Hal)\s*[:.]?\s*(.+)/i, score: 0.85, zone: upperHalfText },
      // "Tentang" keyword anywhere
      { re: /\bTentang\s+(.+?)(?=\n|$)/i, score: 0.8, zone: text }
    ];

    for (let i = 0; i < patterns.length; i++) {
      const p = patterns[i];
      const match = p.zone.match(p.re);
      if (match && match[1]) {
        const value = cleanValue_(match[1]);
        if (value.length >= 5) {
          candidates.push({ value: value, score: p.score, source: 'keyword_uraian_' + i, zone: 0 });
        }
      }
    }

    return rankCandidates_(candidates, true);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: KLASIFIKASI AKSES
  // ═══════════════════════════════════════════════════════════════

  function extractKlasifikasiAksesScored_(text) {
    const candidates = [];

    for (let i = 0; i < AKSES_PATTERNS.length; i++) {
      const pat = AKSES_PATTERNS[i];
      if (pat.re.test(text)) {
        candidates.push({ value: pat.value, score: pat.score, source: 'akses_pattern_' + i, zone: 0 });
      }
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: DARI (Sender)
  // ═══════════════════════════════════════════════════════════════

  function extractDari_(text, structure) {
    const candidates = [];
    const headerText = structure.header.lines.join('\n');
    const sigText = structure.signature.lines.join('\n');

    // Khusus Surat Keluar: Jika KOP mengandung instansi internal, jadikan Pengirim standar
    if (LANRI_KALTIM_RE.test(headerText)) {
      candidates.push({ value: 'Pusjar SKPP LAN RI', score: 0.95, source: 'internal_kop_rule', zone: 0 });
    }

    // Look in header for institution name (KOP SURAT)
    // Scan up to 10 lines to bypass OCR garbage at the top
    const headerLines = structure.header.lines;
    for (let i = 0; i < Math.min(headerLines.length, 10); i++) {
      const line = headerLines[i].trim();
      
      // Syarat KOP:
      // 1. Cukup panjang (min 10 karakter)
      if (line.length < 10) continue;
      // 2. Bukan field metadata surat
      if (/^(NOMOR|NO\.|PERIHAL|HAL|LAMPIRAN|TANGGAL|YTH|KEPADA|SIFAT)/i.test(line)) continue;
      // 3. Bukan nomor seri / telp (kumpulan angka/simbol)
      if (/[\d/.\-]{7,}/.test(line)) continue;
      // 4. Bukan tanggal tempat pembuatan surat
      if (/(?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4}/i.test(line)) continue;
      // 5. Bukan info kontak / alamat jalan pendek (batas 40 agar gabungan OCR Kop+Jalan tidak ke-skip)
      if (line.length < 40 && (line.includes('@') || /www\.|http|telp|fax|email|jalan|jln/i.test(line))) continue;
      
      // Jika lolos semua, ini sangat mungkin nama instansi di KOP
      const isUpper = (line === line.toUpperCase());
      candidates.push({ value: line, score: (isUpper ? 0.6 : 0.55) - (i * 0.05), source: 'kop_line_' + i, zone: 0 });
    }

    // Look for "Dari:" pattern
    const dari = headerText.match(/(?:Dari|Pengirim|Asal)\s*[:.]?\s*(.+)/i);
    if (dari && dari[1]) {
      candidates.push({ value: cleanValue_(dari[1]), score: 0.85, source: 'dari_pattern', zone: 0 });
    }

    // Look in signature for jabatan + nama
    const sigPatterns = [
      /(?:Bupati|Gubernur|Walikota|Menteri|Panglima|Kapolri|Camat|Lurah|Kades|Kepala|Direktur|Ketua|Sekretaris|Rektor|Dekan|Wakil)\s+[\w\s]+/i,
    ];
    for (let s = 0; s < sigPatterns.length; s++) {
      const sigMatch = sigText.match(sigPatterns[s]);
      if (sigMatch) {
        candidates.push({ value: cleanValue_(sigMatch[0]), score: 0.5, source: 'signature_jabatan', zone: 2 });
      }
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: KEPADA (Recipient)
  // ═══════════════════════════════════════════════════════════════

  function extractKepada_(text) {
    const candidates = [];

    // "Kepada:" or "Yth." patterns
    const kepada = text.match(/(?:Kepada|Kpd)\s*[:.]?\s*(.+?)(?=\n|$)/i);
    if (kepada && kepada[1]) {
      candidates.push({ value: cleanValue_(kepada[1]), score: 0.85, source: 'kepada', zone: 0 });
    }

    const yth = text.match(/Yth\.?\s*[:.]?\s*(.+?)(?=\n|$)/i);
    if (yth && yth[1]) {
      candidates.push({ value: cleanValue_(yth[1]), score: 0.8, source: 'yth', zone: 0 });
    }

    // "Kepada Yth." combined
    const kepYth = text.match(/(?:Kepada|Kpd)\s+Yth\.?\s*[:.]?\s*(.+?)(?=\n|$)/i);
    if (kepYth && kepYth[1]) {
      candidates.push({ value: cleanValue_(kepYth[1]), score: 0.9, source: 'kepada_yth', zone: 0 });
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: TANDA TANGAN (Signer)
  // ═══════════════════════════════════════════════════════════════

  function extractTandaTangan_(text, structure) {
    const candidates = [];
    const sigLines = structure.signature.lines;
    if (sigLines.length === 0) return null;

    const result = { jabatan: '', nama: '', nip: '' };

    // Find jabatan (position) — typically before the name
    for (let i = 0; i < sigLines.length; i++) {
      const line = sigLines[i].trim();
      const jabRe = /(?:Kepala|Direktur|Ketua|Sekretaris|Rektor|Dekan|Wakil|Manager|Manajer|Camat|Lurah|Bupati|Walikota|Gubernur|Plt\.?|Pjs\.?)\s+[\w\s.]+/i;
      const jabMatch = line.match(jabRe);
      if (jabMatch) {
        result.jabatan = cleanValue_(jabMatch[0]);
        // Name is usually 2-5 lines after jabatan
        for (let j = i + 2; j < Math.min(i + 6, sigLines.length); j++) {
          const nameLine = sigLines[j].trim();
          // Name lines: have a name pattern, possibly with gelar
          if (nameLine.length > 2 && nameLine.length < 80 &&
              !/^(NIP|NI[Pp]\.?\s)/i.test(nameLine) &&
              !/^(ttd|tandatangan|tanda\s+tangan)/i.test(nameLine) &&
              /^[A-Z]/.test(nameLine)) {
            result.nama = cleanValue_(nameLine.replace(/[.,]+$/, ''));
            break;
          }
        }
        break;
      }
    }

    // Find NIP
    const nipMatch = text.match(/(?:NIP|NI[Pp]\.?)\s*[:.]?\s*(\d{10,})/i);
    if (nipMatch) {
      result.nip = nipMatch[1].trim();
    }

    if (result.jabatan || result.nama) return result;
    return null;
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: LAMPIRAN
  // ═══════════════════════════════════════════════════════════════

  function extractLampiran_(text, structure) {
    const headerText = structure.header.lines.join('\n');
    // "Lampiran: 5 lembar" or "Lamp: 3 berkas"
    const match = headerText.match(/(?:Lampiran|Lamp)\s*[:.]?\s*(\d+)\s*(?:lembar|berkas|dokumen|halaman|eksemplar|lembar|copy)?/i);
    if (match && match[1]) {
      return parseInt(match[1], 10);
    }
    // "Lampiran: 1 (satu) berkas"
    const match2 = headerText.match(/(?:Lampiran|Lamp)\s*[:.]?\s*(\d+)/i);
    if (match2 && match2[1]) {
      return parseInt(match2[1], 10);
    }
    return null;
  }

  // ═══════════════════════════════════════════════════════════════
  //  HELPER FUNCTIONS
  // ═══════════════════════════════════════════════════════════════

  // Deteksi arah surat: jika KOP/header memuat instansi "LAN RI Kalimantan Timur"
  // berarti surat KELUAR (dikeluarkan oleh kita). Bila tidak ada → surat MASUK.
  // Toleransi spasi antar-kata & case agar tahan variasi OCR.
  const LANRI_KALTIM_RE = /LAN\s*RI\s+Kalimantan\s+Timur/i;
  function detectDirection_(headerText, fullText) {
    const zone = String(headerText || '') + '\n' + String(fullText || '').substring(0, 600);
    return LANRI_KALTIM_RE.test(zone) ? 'keluar' : 'masuk';
  }

  function resolveMonth_(name) {
    return INDONESIAN_MONTHS[String(name || '').toLowerCase().replace(/\.$/, '')] || null;
  }

  function pad2_(value) {
    if (value === null || value === undefined) return '00';
    const text = String(value).trim();
    if (!text) return '00';
    return text.length === 1 ? '0' + text : text;
  }

  function cleanValue_(value) {
    return String(value || '')
      .replace(/\s+/g, ' ')
      .replace(/^_+|_+$/g, '')
      .replace(/\s*[.:\-]+$/, '')
      .trim()
      .slice(0, 260);
  }

  function getPositionBonus_(text, index) {
    if (index < 0) return 1.0;
    const totalLen = text.length;
    if (totalLen === 0) return 1.0;
    const ratio = index / totalLen;
    // Bonus for being in first 30% of text (typical for header fields)
    if (ratio < 0.1) return 1.4;
    if (ratio < 0.3) return 1.2;
    if (ratio < 0.5) return 1.0;
    return 0.9;
  }

  function rankCandidates_(candidates, deduplicate) {
    if (!candidates || candidates.length === 0) return null;

    // Sort by score descending
    candidates.sort(function (a, b) { return b.score - a.score; });

    // Deduplicate by value if requested
    if (deduplicate) {
      const seen = {};
      const unique = [];
      for (let i = 0; i < candidates.length; i++) {
        const key = candidates[i].value.toLowerCase();
        if (!seen[key]) {
          seen[key] = true;
          unique.push(candidates[i]);
        }
      }
      candidates = unique;
    }

    const best = candidates[0];
    return {
      value: best.value,
      score: Math.round(best.score * 100) / 100,
      confidence: best.score >= 0.7 ? 'high' : best.score >= 0.4 ? 'medium' : 'low',
      source: best.source,
      candidateCount: candidates.length
    };
  }

  // ═══════════════════════════════════════════════════════════════
  //  MAIN ANALYZE ENTRY POINT
  // ═══════════════════════════════════════════════════════════════

  /**
   * Main entry point for document analysis.
   * @param {string} rawText - Raw OCR text from document
   * @param {string} fileName - Original file name
   * @param {object} context - { activity: {}, subActivity: {} }
   * @returns {object} Analysis result with fields, scores, and metadata
   */
  function analyze(rawText, fileName, context) {
    const startTime = Date.now();
    context = context || {};

    // Pass 1: Pre-process
    const text = normalizeText_(rawText);

    // Pass 2: Analyze structure
    const structure = analyzeStructure_(text);

    // Pass 2.5: Trim body to maximize speed & accuracy
    const topLines = structure.header.lines.concat(structure.body.lines.slice(0, 7));
    const topText = topLines.join('\n');
    const bottomText = structure.signature.lines.join('\n');
    const optimizedText = topText + '\n\n' + bottomText;

    // Pass 2.6: Deteksi arah surat (masuk/keluar) dari KOP header
    const headerTextForDir = structure.header.lines.join('\n');
    let direction = detectDirection_(headerTextForDir, text);

    // Pass 3: Classify document type
    const docType = classifyDocumentType_(topText, fileName);

    // Pass 4: Extract all fields with scoring
    const nomorSurat = extractNomorSuratScored_(optimizedText, structure, fileName);
    let kodeKlasifikasi = extractKodeKlasifikasiScored_(optimizedText, structure);

    // Pass 4b: Untuk surat KELUAR, kode klasifikasi tertanam di dalam Nomor Surat
    // sebagai segmen yang diawali "PDP" (kode klasifikasi khusus surat keluar kantor
    // ini). Contoh: 273/P.3/PDP.07.1
    //   -> Nomor Surat tetap FULL (273/P.3/PDP.07.1)
    //   -> Kode Klasifikasi disalin cukup segmen PDP-nya (PDP.07.1)
    // Nomor Surat TIDAK dipotong. Keberadaan segmen PDP itu sendiri = penanda pasti
    // surat keluar (surat masuk dari kantor lain tak akan memakai kode PDP), jadi
    // arah surat dipaksa 'keluar' agar kolom kode klasifikasi ikut terisi.
    if (nomorSurat && nomorSurat.value) {
      const pdpMatch = nomorSurat.value.match(/\b(PDP\.\d{1,3}(?:\.\d{1,3})*)\b/i);
      if (pdpMatch) {
        kodeKlasifikasi = { value: pdpMatch[1].toUpperCase(), score: 0.98, confidence: 'high', source: 'pdp_from_nomor', candidateCount: 1 };
        direction = 'keluar';
      }
    }

    const tanggal = extractTanggalScored_(optimizedText, structure);
    const uraian = extractUraianScored_(topText, structure, fileName, context.activity, context.subActivity);
    const klasifikasiAkses = extractKlasifikasiAksesScored_(topText);
    const dari = extractDari_(topText, structure);
    const kepada = extractKepada_(topText);
    const tandaTangan = extractTandaTangan_(optimizedText, structure);
    const lampiran = extractLampiran_(topText, structure);
    
    const tingkatPerkembangan = (function() {
      const pat1 = /dokumen\s+ini\s+telah\s+ditandatangani\s+secara\s+elektronik\s+menggunakan\s+sertifikat\s+elektronik/i;
      const pat2 = /ditandatangani\s+secara\s+elektronik\s+menggunakan\s+sertifikat/i;
      if (pat1.test(optimizedText) || pat2.test(optimizedText)) {
        return { value: 'Asli', score: 0.99, confidence: 'high', source: 'electronic_signature_detection' };
      }
      return { value: 'Salinan', score: 0.99, confidence: 'high', source: 'default_salinan' };
    })();

    // Pass 5: Build result
    const fields = {};
    if (nomorSurat) fields.nomor_surat = nomorSurat;
    // Kode klasifikasi hanya relevan untuk surat KELUAR. Surat MASUK: biarkan kosong
    // (jangan autofill) sesuai aturan kearsipan di kantor ini.
    if (kodeKlasifikasi && direction === 'keluar') fields.kode_klasifikasi = kodeKlasifikasi;
    if (tanggal) fields.tanggal = tanggal;
    let uraianFallback = String(fileName || '')
      .replace(/\.[a-z0-9]+$/i, '')
      .replace(/[_-]+/g, ' ')
      .replace(/^\d{14,}\.?\s*(?:\([^)]+\)\s*)?/, '')
      .replace(/^\d{1,3}\.?\s*(?:\([^)]+\)\s*)?/, '');
      
    if (nomorSurat && nomorSurat.value) {
      const escNomor = nomorSurat.value.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&').replace(/\//g, '[\\s\\/\\-_]+');
      uraianFallback = uraianFallback.replace(new RegExp('(?:No(?:mor)?\\s*[:.]?\\s*)?' + escNomor, 'i'), '');
    }
    
    uraianFallback = uraianFallback.replace(/\s+/g, ' ').trim();

    if (docType === 'Nota Dinas') {
      if (uraian && uraian.value && !/^nota\s+dinas\b/i.test(uraian.value)) {
        uraian.value = 'Nota Dinas ' + uraian.value;
      }
      if (uraianFallback && !/^nota\s+dinas\b/i.test(uraianFallback)) {
        uraianFallback = 'Nota Dinas ' + uraianFallback;
      }
    }

    if (uraian && uraian.value) {
      fields.uraian_informasi_item = uraian;
    } else if (uraianFallback) {
      fields.uraian_informasi_item = { value: uraianFallback, score: 0.4, confidence: 'low', source: 'filename_fallback' };
    }
    if (klasifikasiAkses) fields.klasifikasi_akses = klasifikasiAkses;
    if (docType === 'surat_keluar') {
      fields.dari = { value: 'Pusjar SKPP LAN RI', score: 0.99, confidence: 'high', source: 'hardcoded_surat_keluar' };
    } else if (dari) {
      if (dari.value) dari.value = applyAliases_(dari.value);
      fields.dari = dari;
    }
    
    if (kepada) {
      if (kepada.value) kepada.value = applyAliases_(kepada.value);
      fields.kepada = kepada;
    }
    if (tandaTangan) fields.tanda_tangan = { value: tandaTangan, score: 0.6, confidence: 'medium' };
    if (lampiran !== null) fields.lampiran = { value: String(lampiran), score: 0.8, confidence: 'high' };
    fields.tingkat_perkembangan = tingkatPerkembangan;

    const fieldCount = Object.keys(fields).length;

    return {
      fields: fields,
      rawTextLength: text.length,
      parseDuration: Date.now() - startTime,
      fieldCount: fieldCount,
      totalFields: 9,
      documentType: docType,
      documentDirection: direction,
      structure: {
        headerLines: structure.header.lines.length,
        bodyLines: structure.body.lines.length,
        signatureLines: structure.signature.lines.length,
        totalLines: getLines_(text).length
      }
    };
  }

  // ═══════════════════════════════════════════════════════════════
  //  PUBLIC API
  // ═══════════════════════════════════════════════════════════════

  return {
    analyze: analyze,
    normalizeText: normalizeText_,
    analyzeStructure: analyzeStructure_,
    classifyDocumentType: classifyDocumentType_,
    detectDirection: detectDirection_
  };

})();
