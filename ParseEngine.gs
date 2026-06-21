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
var ParseEngine = (function () {

  // ═══════════════════════════════════════════════════════════════
  //  CONSTANTS
  // ═══════════════════════════════════════════════════════════════

  var INDONESIAN_MONTHS = {
    januari: '01', feb: '02', februari: '02', mar: '03', maret: '03',
    apr: '04', april: '04', mei: '05', jun: '06', juni: '06',
    jul: '07', juli: '07', agu: '08', agustus: '08', agst: '08',
    sep: '09', sept: '09', september: '09', okt: '10', oktober: '10',
    nov: '11', november: '11', des: '12', desember: '12'
  };

  var MONTH_FULL = 'Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember';
  var MONTH_ABBR = 'Jan|Feb|Mar|Apr|Mei|Jun|Jul|Agu|Agst|Sep|Sept|Okt|Nov|Des';
  var MONTH_ALL = MONTH_FULL + '|' + MONTH_ABBR;

  // Common OCR character confusions
  var _NL = String.fromCharCode(10);
  var OCR_FIXES = [
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
  var NOMOR_FORMATS = [
    { name: 'SK format', re: /(?:SK|Surat\s+Keputusan)[\s/.:-]+(\d+\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.95 },
    { name: 'SP format', re: /(?:SP|Surat\s+Perintah)[\s/.:-]+([A-Z]?\d{1,6}[\-\/][A-Z0-9.\/\-]+\/[12]\d{3})/i, weight: 0.95 },
    { name: 'ST format', re: /(?:ST|Surat\s+Tugas)[\s/.:-]+(\d+\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.95 },
    { name: 'SE format', re: /(?:SE|Surat\s+Edaran)[\s/.:-]+(\d+\/[A-Z0-9.\/\-]+\/[12]\d{3})/i, weight: 0.95 },
    { name: 'explicit Nomor', re: /(?:No(?:mor)?)\s*[:.]\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.9 },
    { name: 'explicit No', re: /(?:No(?:mor)?)\s*[:.]?\s*([A-Z0-9][A-Z0-9.\/\-]+(?:\/[A-Z0-9.\-]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)/i, weight: 0.85 },
    { name: 'letter prefix', re: /\b([A-Z]{1,4}[\-]\d{1,6}(?:\/[A-Z0-9.]+)+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)\b/i, weight: 0.8 },
    { name: 'numeric segments', re: /\b(\d{1,6}\/[A-Z0-9.\/\-]+\/[12]\d{3}(?:\/[A-Z0-9.\-]+)?)\b/i, weight: 0.7 },
    { name: 'Nomor Tahun', re: /(?:No(?:mor)?)\s*[:.]?\s*(\d{1,6})\s+Tahun\s+(20[12]\d)/i, weight: 0.75 },
    { name: 'slash year', re: /\b(\d{1,5}\/[A-Z]{1,4}[A-Z0-9.]*\/[12]\d{3})\b/i, weight: 0.6 }
  ];

  // Document type keywords
  var DOC_TYPE_KEYWORDS = {
    'Surat Keputusan': ['surat keputusan', 'keputusan kepala', 'keputusan direktur', 'keputusan ketua', 'keputusan rektor', 'menetapkan', 'mengingat', 'memutuskan', 'kesatu', 'kedua', 'ketiga'],
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

  // Klasifikasi akses phrase patterns (ordered by specificity)
  var AKSES_PATTERNS = [
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
  var KODE_PATTERNS = [
    { re: /(?:kode|kode\s+klasifikasi)\s*[:.]?\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?)/i, weight: 0.95 },
    { re: /klasifikasi\s*[:.]?\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?)/i, weight: 0.9 },
    { re: /\b([A-Z]{2,4}\.\d{2}\.\d{1,2})\b/, weight: 0.8 },
    { re: /\b([A-Z]{2,4}\.\d{2})\b/, weight: 0.65 },
    { re: /\b([A-Z]\.\d{2}(?:\.\d{1,2})?)\b/, weight: 0.5 }
  ];

  // Uraian section-boundary keywords (where Perihal value stops)
  var URAIAN_STOP = '\\b(?:Ke(?:pada)?|Lampiran|Yth|Nomor|Tanggal|Perihal|Hal|Lamp|Isi|Dengan)\\b';

  // ═══════════════════════════════════════════════════════════════
  //  TEXT PREPROCESSOR
  // ═══════════════════════════════════════════════════════════════

  function normalizeText_(raw) {
    var text = String(raw || '');
    for (var i = 0; i < OCR_FIXES.length; i++) {
      text = text.replace(OCR_FIXES[i][0], OCR_FIXES[i][1]);
    }
    return text;
  }

  function getLines_(text) {
    return String(text || '').split('\n');
  }

  function findKeywordLine_(text, keywords) {
    var lines = getLines_(text);
    for (var i = 0; i < lines.length; i++) {
      var lower = lines[i].toLowerCase();
      for (var j = 0; j < keywords.length; j++) {
        if (lower.indexOf(keywords[j]) >= 0) {
          return { lineIndex: i, line: lines[i], context: getContext_(lines, i, 3) };
        }
      }
    }
    return null;
  }

  function getContext_(lines, index, radius) {
    var start = Math.max(0, index - radius);
    var end = Math.min(lines.length - 1, index + radius);
    return lines.slice(start, end + 1).join('\n');
  }

  // ═══════════════════════════════════════════════════════════════
  //  DOCUMENT STRUCTURE ANALYZER
  // ═══════════════════════════════════════════════════════════════

  function analyzeStructure_(text) {
    var lines = getLines_(text);
    var totalLines = lines.length;
    var structure = {
      header: { start: 0, end: Math.min(totalLines, Math.ceil(totalLines * 0.3)), lines: [] },
      body: { start: 0, end: totalLines, lines: [] },
      signature: { start: 0, end: totalLines, lines: [] },
      footer: { start: 0, end: totalLines, lines: [] }
    };

    // Find signature block
    for (var i = totalLines - 1; i >= Math.floor(totalLines * 0.6); i--) {
      var lower = lines[i].toLowerCase();
      if (lower.indexOf('hormat') >= 0 || lower.indexOf('tanda tangan') >= 0 ||
          lower.indexOf('kepala') >= 0 || lower.indexOf('direktur') >= 0 ||
          lower.indexOf('ketua') >= 0 || lower.indexOf('sekretaris') >= 0 ||
          /ditetapkan\s+di/i.test(lines[i]) || /ditandatangani/i.test(lines[i])) {
        structure.signature = { start: i, end: totalLines, lines: lines.slice(i) };
        break;
      }
    }

    // Find header (first 35% or until body/salutation markers are encountered)
    var headerEnd = 0;
    for (var j = 0; j < Math.min(totalLines, Math.ceil(totalLines * 0.35)); j++) {
      var line = lines[j].trim();
      if (/^(?:yth|kepada|kpd|dengan\s+hormat|menindaklanjuti|sehubungan|merujuk|berdasarkan)\b/i.test(line)) {
        break;
      }
      if (/nomor|no\.|perihal|hal\s*:/i.test(line)) {
        headerEnd = j + 1;
      }
    }
    if (headerEnd === 0) {
      var salutationIndex = -1;
      for (var j = 0; j < Math.min(totalLines, Math.ceil(totalLines * 0.35)); j++) {
        if (/^(?:yth|kepada|kpd|dengan\s+hormat|menindaklanjuti|sehubungan|merujuk|berdasarkan)\b/i.test(lines[j].trim())) {
          salutationIndex = j;
          break;
        }
      }
      headerEnd = salutationIndex > 0 ? salutationIndex : Math.min(totalLines, 10);
    }
    structure.header.end = headerEnd;
    structure.header.lines = lines.slice(0, headerEnd);

    // Body is between header and signature
    structure.body.start = headerEnd;
    structure.body.end = structure.signature.start;
    structure.body.lines = lines.slice(headerEnd, structure.signature.start);

    // Footer is last 10% (tembusan, etc.)
    var footerStart = Math.max(structure.signature.start, Math.floor(totalLines * 0.9));
    structure.footer.start = footerStart;
    structure.footer.lines = lines.slice(footerStart);

    return structure;
  }

  // ═══════════════════════════════════════════════════════════════
  //  DOCUMENT TYPE CLASSIFIER
  // ═══════════════════════════════════════════════════════════════

  function classifyDocumentType_(text, fileName) {
    var combined = (String(text || '') + ' ' + String(fileName || '')).toLowerCase();
    var scores = {};
    var types = Object.keys(DOC_TYPE_KEYWORDS);

    for (var i = 0; i < types.length; i++) {
      var type = types[i];
      var keywords = DOC_TYPE_KEYWORDS[type];
      var score = 0;
      for (var j = 0; j < keywords.length; j++) {
        var count = 0;
        var pos = combined.indexOf(keywords[j]);
        while (pos >= 0) { count++; pos = combined.indexOf(keywords[j], pos + 1); }
        score += count * (j < 2 ? 3 : 1);  // first 2 keywords weighted more
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

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: NOMOR SURAT
  // ═══════════════════════════════════════════════════════════════

  function extractNomorSuratScored_(text, structure, fileName) {
    var candidates = [];
    var headerText = structure.header.lines.join('\n');

    // STRICT keyword-anchored patterns only
    var patterns = [
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
    var searchZones = [
      { text: headerText, bonus: 1.3 },
      { text: text.substring(0, Math.floor(text.length * 0.3)), bonus: 1.0 },
      { text: String(fileName || ''), bonus: 1.5 }
    ];

    for (var z = 0; z < searchZones.length; z++) {
      var zone = searchZones[z];
      for (var i = 0; i < patterns.length; i++) {
        var p = patterns[i];
        var match = zone.text.match(p.re);
        if (match && match[1]) {
          var value = match[1].replace(/\s+/g, ' ').trim();
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
    var candidates = [];
    var headerText = structure.header.lines.join('\n');

    // STRICT keyword-anchored patterns only
    var patterns = [
      // Explicit "Kode:" or "Klasifikasi:" prefix
      { re: /(?:Kode|Klasifikasi)\s*[:.\s]+\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)/i, score: 0.95 },
      // "Kode Klasifikasi:" prefix
      { re: /Kode\s+Klasifikasi\s*[:.\s]+\s*([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)/i, score: 0.98 },
      // Bare code in header (must have a dot for numeric to prevent matching random 3 digits like 595)
      { re: /\b([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)\b/i, score: 0.7 }
    ];

    // Only search in header
    var searchZones = [
      { text: headerText, bonus: 1.3 }
    ];

    for (var z = 0; z < searchZones.length; z++) {
      var zone = searchZones[z];
      for (var i = 0; i < patterns.length; i++) {
        var p = patterns[i];
        var match = zone.text.match(p.re);
        if (match && match[1]) {
          var value = match[1].toUpperCase().trim();
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
    var candidates = [];
    var headerText = structure.header.lines.join('\n');
    var sigText = structure.signature.lines.join('\n');

    // STRICT keyword-anchored patterns only
    var patterns = [
      // "Tanggal DD MMMM YYYY" or "Ditetapkan pada DD MMMM YYYY"
      { re: new RegExp('(?:Tanggal|Ditetapkan|Ditandatangani|Berangka)[\\s:.,]+(?:[^\\n]{0,20}?)(\\d{1,2})\\s+(' + MONTH_ALL + ')\\s+(20[12]\\d)', 'i'), score: 0.95 },
      // Bare date with City prefix like "Samarinda,      // Standard ID format
      { re: new RegExp('(?:[A-Za-z\\s]+,\\s*)?(\\d{1,2})\\s+(' + MONTH_ALL + ')\\s+(20[12]\\d)', 'i'), score: 0.85 },
      // Contextual ISO format
      { re: /(?:Tanggal|Ditetapkan|Ditandatangani)\s*[:.,]?\s*(?:[^\n]{0,10}?)(20[12]\d)[\-\/.](\d{1,2})[\-\/.](\d{1,2})/i, score: 0.9 }
    ];

    // Only search in header and signature
    var searchZones = [
      { text: headerText, bonus: 1.8 },
      { text: sigText, bonus: 1.0 }
    ];

    for (var z = 0; z < searchZones.length; z++) {
      var zone = searchZones[z];
      for (var i = 0; i < patterns.length; i++) {
        var p = patterns[i];
        var match = zone.text.match(p.re);
        if (match) {
          var dateVal = null;
          if (match[1] && match[2] && match[3]) {
            var m = resolveMonth_(match[2]);
            if (m) dateVal = [match[3], m, pad2_(match[1])].join('-');
          } else if (match[4] && match[5] && match[6]) {
            dateVal = [match[4], pad2_(match[5]), pad2_(match[6])].join('-');
          }
          if (dateVal) {
            var indexInFullText = text.indexOf(match[0]);
            var posBonus = getPositionBonus_(text, indexInFullText);
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
    var candidates = [];
    var upperHalfText = text.substring(0, Math.floor(text.length * 0.5));

    // STRICT keyword-anchored patterns only
    var patterns = [
      // Multi-line Perihal in header
      { re: new RegExp('(?:Perihal|Hal)\\s*[:.]?\\s*(.+?)(?=\\n' + URAIAN_STOP + '|\\n\\n|$)', 'is'), score: 0.95, zone: upperHalfText },
      // Simple Perihal/Hal on same line in header
      { re: /(?:Perihal|Hal)\s*[:.]?\\s*(.+)/i, score: 0.85, zone: upperHalfText },
      // "Tentang" keyword anywhere
      { re: /\bTentang\s+(.+?)(?=\n|$)/i, score: 0.8, zone: text }
    ];

    for (var i = 0; i < patterns.length; i++) {
      var p = patterns[i];
      var match = p.zone.match(p.re);
      if (match && match[1]) {
        var value = cleanValue_(match[1]);
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
    var candidates = [];

    for (var i = 0; i < AKSES_PATTERNS.length; i++) {
      var pat = AKSES_PATTERNS[i];
      if (pat.re.test(text)) {
        candidates.push({ value: pat.value, score: pat.score, source: 'akses_pattern_' + i, zone: 0 });
      }
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: PENGIRIM (Sender)
  // ═══════════════════════════════════════════════════════════════

  function extractPengirim_(text, structure) {
    var candidates = [];
    var headerText = structure.header.lines.join('\n');
    var sigText = structure.signature.lines.join('\n');

    // Look in header for institution name (KOP SURAT)
    // Typically the first 2-5 lines contain the institution
    var headerLines = structure.header.lines;
    for (var i = 0; i < Math.min(headerLines.length, 5); i++) {
      var line = headerLines[i].trim();
      // Institution lines are usually uppercase and > 5 chars
      if (line.length > 5 && line === line.toUpperCase() && !/^(NOMOR|NO\.|PERIHAL|HAL|LAMPIRAN|TANGGAL)/i.test(line)) {
        candidates.push({ value: line, score: 0.6 - (i * 0.1), source: 'kop_line_' + i, zone: 0 });
      }
    }

    // Look for "Dari:" pattern
    var dari = headerText.match(/(?:Dari|Pengirim|Asal)\s*[:.]?\s*(.+)/i);
    if (dari && dari[1]) {
      candidates.push({ value: cleanValue_(dari[1]), score: 0.85, source: 'dari_pattern', zone: 0 });
    }

    // Look in signature for jabatan + nama
    var sigPatterns = [
      /(?:Kepala|Direktur|Ketua|Sekretaris|Rektor|Dekan|Wakil)\s+[\w\s]+/i,
    ];
    for (var s = 0; s < sigPatterns.length; s++) {
      var sigMatch = sigText.match(sigPatterns[s]);
      if (sigMatch) {
        candidates.push({ value: cleanValue_(sigMatch[0]), score: 0.5, source: 'signature_jabatan', zone: 2 });
      }
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: PENERIMA (Recipient)
  // ═══════════════════════════════════════════════════════════════

  function extractPenerima_(text) {
    var candidates = [];

    // "Kepada:" or "Yth." patterns
    var kepada = text.match(/(?:Kepada|Kpd)\s*[:.]?\s*(.+?)(?=\n|$)/i);
    if (kepada && kepada[1]) {
      candidates.push({ value: cleanValue_(kepada[1]), score: 0.85, source: 'kepada', zone: 0 });
    }

    var yth = text.match(/Yth\.?\s+(.+?)(?=\n|$)/i);
    if (yth && yth[1]) {
      candidates.push({ value: cleanValue_(yth[1]), score: 0.8, source: 'yth', zone: 0 });
    }

    // "Kepada Yth." combined
    var kepYth = text.match(/(?:Kepada|Kpd)\s+Yth\.?\s+(.+?)(?=\n|$)/i);
    if (kepYth && kepYth[1]) {
      candidates.push({ value: cleanValue_(kepYth[1]), score: 0.9, source: 'kepada_yth', zone: 0 });
    }

    return rankCandidates_(candidates);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SCORED EXTRACTOR: TANDA TANGAN (Signer)
  // ═══════════════════════════════════════════════════════════════

  function extractTandaTangan_(text, structure) {
    var candidates = [];
    var sigLines = structure.signature.lines;
    if (sigLines.length === 0) return null;

    var result = { jabatan: '', nama: '', nip: '' };

    // Find jabatan (position) — typically before the name
    for (var i = 0; i < sigLines.length; i++) {
      var line = sigLines[i].trim();
      var jabRe = /(?:Kepala|Direktur|Ketua|Sekretaris|Rektor|Dekan|Wakil|Manager|Manajer|Camat|Lurah|Bupati|Walikota|Gubernur|Plt\.?|Pjs\.?)\s+[\w\s.]+/i;
      var jabMatch = line.match(jabRe);
      if (jabMatch) {
        result.jabatan = cleanValue_(jabMatch[0]);
        // Name is usually 2-5 lines after jabatan
        for (var j = i + 2; j < Math.min(i + 6, sigLines.length); j++) {
          var nameLine = sigLines[j].trim();
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
    var nipMatch = text.match(/(?:NIP|NI[Pp]\.?)\s*[:.]?\s*(\d{10,})/i);
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
    var headerText = structure.header.lines.join('\n');
    // "Lampiran: 5 lembar" or "Lamp: 3 berkas"
    var match = headerText.match(/(?:Lampiran|Lamp)\s*[:.]?\s*(\d+)\s*(?:lembar|berkas|dokumen|halaman|eksemplar|lembar|copy)?/i);
    if (match && match[1]) {
      return parseInt(match[1], 10);
    }
    // "Lampiran: 1 (satu) berkas"
    var match2 = headerText.match(/(?:Lampiran|Lamp)\s*[:.]?\s*(\d+)/i);
    if (match2 && match2[1]) {
      return parseInt(match2[1], 10);
    }
    return null;
  }

  // ═══════════════════════════════════════════════════════════════
  //  HELPER FUNCTIONS
  // ═══════════════════════════════════════════════════════════════

  function resolveMonth_(name) {
    return INDONESIAN_MONTHS[String(name || '').toLowerCase().replace(/\.$/, '')] || null;
  }

  function pad2_(value) {
    if (value === null || value === undefined) return '00';
    var text = String(value).trim();
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
    var totalLen = text.length;
    if (totalLen === 0) return 1.0;
    var ratio = index / totalLen;
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
      var seen = {};
      var unique = [];
      for (var i = 0; i < candidates.length; i++) {
        var key = candidates[i].value.toLowerCase();
        if (!seen[key]) {
          seen[key] = true;
          unique.push(candidates[i]);
        }
      }
      candidates = unique;
    }

    var best = candidates[0];
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
    var startTime = Date.now();
    context = context || {};

    // Pass 1: Pre-process
    var text = normalizeText_(rawText);

    // Pass 2: Analyze structure
    var structure = analyzeStructure_(text);

    // Pass 2.5: Trim body to maximize speed & accuracy
    var topLines = structure.header.lines.concat(structure.body.lines.slice(0, 7));
    var topText = topLines.join('\n');
    var bottomText = structure.signature.lines.join('\n');
    var optimizedText = topText + '\n\n' + bottomText;

    // Pass 3: Classify document type
    var docType = classifyDocumentType_(topText, fileName);

    // Pass 4: Extract all fields with scoring
    var nomorSurat = extractNomorSuratScored_(optimizedText, structure, fileName);
    var kodeKlasifikasi = extractKodeKlasifikasiScored_(optimizedText, structure);

    // Pass 4b: Extract kode from nomor if embedded and split them
    if (nomorSurat && nomorSurat.value) {
      var kodeMatch = nomorSurat.value.match(/(?:^|[A-Z]{1,3}[\/-]|[\/-])([A-Z]{1,4}\.\d{2}(?:\.\d{1,2})?|\d{3}\.\d{1,3}(?:\.\d{1,3})*)(?:[\/-]|$)/i);
      if (kodeMatch) {
        if (!kodeKlasifikasi || kodeKlasifikasi.score < 0.95) {
          kodeKlasifikasi = { value: kodeMatch[1].toUpperCase(), score: 0.95, confidence: 'high', source: 'extracted_from_nomor', candidateCount: 1 };
        }
        // Strip kode klasifikasi from nomor surat
        var escapedKode = kodeMatch[1].replace(/\./g, '\\.');
        nomorSurat.value = nomorSurat.value.replace(new RegExp(escapedKode, 'i'), '')
          .replace(/[\/-][\/-]+/g, '/')
          .replace(/([A-Z]{1,3}-)\//i, '$1')
          .replace(/^[\\/-]|[\\/-]$/g, '');
        if (/^[A-Z]{1,3}[\/-]?$/i.test(nomorSurat.value)) nomorSurat.value = '';
      }
    }

    var tanggal = extractTanggalScored_(optimizedText, structure);
    var uraian = extractUraianScored_(topText, structure, fileName, context.activity, context.subActivity);
    var klasifikasiAkses = extractKlasifikasiAksesScored_(topText);
    var pengirim = extractPengirim_(topText, structure);
    var penerima = extractPenerima_(topText);
    var tandaTangan = extractTandaTangan_(optimizedText, structure);
    var lampiran = extractLampiran_(topText, structure);
    
    var tingkatPerkembangan = (function() {
      var pat1 = /dokumen\s+ini\s+telah\s+ditandatangani\s+secara\s+elektronik\s+menggunakan\s+sertifikat\s+elektronik/i;
      var pat2 = /ditandatangani\s+secara\s+elektronik\s+menggunakan\s+sertifikat/i;
      if (pat1.test(optimizedText) || pat2.test(optimizedText)) {
        return { value: 'Asli', score: 0.99, confidence: 'high', source: 'electronic_signature_detection' };
      }
      return { value: 'Salinan', score: 0.99, confidence: 'high', source: 'default_salinan' };
    })();

    // Pass 5: Build result
    var fields = {};
    if (nomorSurat) fields.nomor_surat = nomorSurat;
    if (kodeKlasifikasi) fields.kode_klasifikasi = kodeKlasifikasi;
    if (tanggal) fields.tanggal = tanggal;
    var uraianFallback = String(fileName || '')
      .replace(/\.[a-z0-9]+$/i, '')
      .replace(/[_-]+/g, ' ')
      .replace(/^\d{14,}\.?\s*(?:\([^)]+\)\s*)?/, '')
      .replace(/^\d{1,3}\.?\s*(?:\([^)]+\)\s*)?/, '');
      
    if (nomorSurat && nomorSurat.value) {
      var escNomor = nomorSurat.value.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&').replace(/\//g, '[\\s\\/\\-_]+');
      uraianFallback = uraianFallback.replace(new RegExp('(?:No(?:mor)?\\s*[:.]?\\s*)?' + escNomor, 'i'), '');
    }
    
    uraianFallback = uraianFallback.replace(/\s+/g, ' ').trim();

    if (uraian && uraian.value) {
      fields.uraian_informasi_item = uraian;
    } else if (uraianFallback) {
      fields.uraian_informasi_item = { value: uraianFallback, score: 0.4, confidence: 'low', source: 'filename_fallback' };
    }
    if (klasifikasiAkses) fields.klasifikasi_akses = klasifikasiAkses;
    if (pengirim) fields.pengirim = pengirim;
    if (penerima) fields.penerima = penerima;
    if (tandaTangan) fields.tanda_tangan = { value: tandaTangan, score: 0.6, confidence: 'medium' };
    if (lampiran !== null) fields.lampiran = { value: String(lampiran), score: 0.8, confidence: 'high' };
    fields.tingkat_perkembangan = tingkatPerkembangan;

    var fieldCount = Object.keys(fields).length;

    return {
      fields: fields,
      rawTextLength: text.length,
      parseDuration: Date.now() - startTime,
      fieldCount: fieldCount,
      totalFields: 9,
      documentType: docType,
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
    classifyDocumentType: classifyDocumentType_
  };

})();
