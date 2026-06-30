const assert = require('assert');
const {
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
} = require('./metadata.pure.js');

console.log('Running unit tests for Portal Arsip pure helpers...\n');

let passedCount = 0;
let failedCount = 0;

function test(name, fn) {
  try {
    fn();
    console.log(`  ✓ ${name}`);
    passedCount++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(err);
    failedCount++;
  }
}

// 1. extractNomorSurat_
test('extractNomorSurat_ - standard No format', () => {
  const text = 'Surat Keputusan No. 123/ABC/2026 tentang sesuatu';
  assert.strictEqual(extractNomorSurat_(text), '123/ABC/2026');
});

test('extractNomorSurat_ - Nomor format', () => {
  const text = 'Nomor: 456/DEF-G/2026';
  assert.strictEqual(extractNomorSurat_(text), '456/DEF-G/2026');
});

test('extractNomorSurat_ - B- prefix format', () => {
  const text = 'Surat Undangan B-789/XYZ/2026/Latbang';
  assert.strictEqual(extractNomorSurat_(text), 'B-789/XYZ/2026/Latbang');
});

test('extractNomorSurat_ - pure slash and digits format', () => {
  const text = 'No 01/02/03/2026';
  assert.strictEqual(extractNomorSurat_(text), '01/02/03/2026');
});

// 2. extractKodeKlasifikasi_
test('extractKodeKlasifikasi_ - standard classification code', () => {
  const text = 'Klasifikasi KP.01.02 tentang kepegawaian';
  assert.strictEqual(extractKodeKlasifikasi_(text), 'KP.01.02');
});

test('extractKodeKlasifikasi_ - 2 digits code', () => {
  const text = 'Klasifikasi HM.02 tentang humas';
  assert.strictEqual(extractKodeKlasifikasi_(text), 'HM.02');
});

// 3. extractDate_
test('extractDate_ - ISO format', () => {
  const text = 'Surat bertanggal 2026-05-30.';
  assert.strictEqual(extractDate_(text), '2026-05-30');
});

test('extractDate_ - Indonesian format', () => {
  const text = 'Jakarta, 30 Mei 2026';
  assert.strictEqual(extractDate_(text), '2026-05-30');
});

test('extractDate_ - Indonesian single digit day', () => {
  const text = 'Surat dibuat pada 9 Januari 2026';
  assert.strictEqual(extractDate_(text), '2026-01-09');
});

// 4. extractTingkatPerkembangan_
test('extractTingkatPerkembangan_ - asli', () => {
  assert.strictEqual(extractTingkatPerkembangan_('Surat_Asli_Undangan.pdf'), 'Asli');
});

test('extractTingkatPerkembangan_ - copy menjadi Salinan', () => {
  assert.strictEqual(extractTingkatPerkembangan_('Copy_of_Undangan.pdf'), 'Salinan');
});

test('extractTingkatPerkembangan_ - default/none', () => {
  assert.strictEqual(extractTingkatPerkembangan_('Undangan.pdf'), '');
});

// 4b. extractKlasifikasiAkses_
test('extractKlasifikasiAkses_ - rahasia', () => {
  assert.strictEqual(extractKlasifikasiAkses_('Dokumen ini bersifat RAHASIA'), 'Rahasia');
});

test('extractKlasifikasiAkses_ - terbatas', () => {
  assert.strictEqual(extractKlasifikasiAkses_('Klasifikasi: Terbatas'), 'Terbatas');
});

test('extractKlasifikasiAkses_ - biasa', () => {
  assert.strictEqual(extractKlasifikasiAkses_('Untuk kalangan BIASA / umum'), 'Biasa');
});

test('extractKlasifikasiAkses_ - terbuka', () => {
  assert.strictEqual(extractKlasifikasiAkses_('Dokumen TERBUKA'), 'Biasa');
});

test('extractKlasifikasiAkses_ - no match returns empty', () => {
  assert.strictEqual(extractKlasifikasiAkses_('Surat Undangan Rapat'), '');
});

test('extractKlasifikasiAkses_ - empty string returns empty', () => {
  assert.strictEqual(extractKlasifikasiAkses_(''), '');
});

// 5. extractUraian_
test('extractUraian_ - Perihal match', () => {
  const text = 'Nomor: 123\nPerihal: Rapat Evaluasi Portal Arsip\nTanggal: 30 Mei';
  const sourceName = 'surat.pdf';
  const activity = { activity_name: 'Kegiatan A' };
  const subActivity = { sub_activity_name: 'Sub A' };
  assert.strictEqual(extractUraian_(text, sourceName, activity, subActivity), 'Rapat Evaluasi Portal Arsip');
});

test('extractUraian_ - filename fallback', () => {
  const text = 'Random text without perihal';
  const sourceName = 'Surat-Undangan_Rapat.pdf';
  const activity = { activity_name: 'Kegiatan A' };
  const subActivity = { sub_activity_name: 'Sub A' };
  assert.strictEqual(extractUraian_(text, sourceName, activity, subActivity), 'Surat Undangan Rapat');
});

// 6. buildFinalFileName_
test('buildFinalFileName_ - with nomor_surat', () => {
  const metadata = {
    nomor_item_arsip: '2',
    tingkat_perkembangan: 'Asli',
    nomor_item_arsip: '15',
    tingkat_perkembangan: 'Asli',
    uraian_informasi_item: 'Undangan Rapat/Kepada Sekretaris Daerah/Dari Panitia Pusat'
  };
  const sourceName = 'original.pdf';
  assert.strictEqual(
    buildFinalFileName_(metadata, null),
    '15. (Asli) Undangan Rapat/Kepada Sekretaris Daerah/Dari Panitia Pusat.pdf'
  );
});

test('buildFinalFileName_ - without nomor_surat', () => {
  const metadata = {
    nomor_item_arsip: '15',
    tingkat_perkembangan: 'Copy',
    uraian_informasi_item: 'Laporan Bulanan'
  };
  const sourceName = 'original.pdf';
  assert.strictEqual(
    buildFinalFileName_(metadata, sourceName),
    '15. (Copy) Laporan Bulanan.pdf'
  );
});

// 7. cleanId_
test('cleanId_ - extraction from URL', () => {
  const url = 'https://docs.google.com/spreadsheets/d/1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef/edit';
  assert.strictEqual(cleanId_(url), '1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef');
});

test('cleanId_ - raw ID stays same', () => {
  const rawId = '1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef';
  assert.strictEqual(cleanId_(rawId), rawId);
});

// 8. slug_
test('slug_ - normal conversion', () => {
  assert.strictEqual(slug_('Angkatan 1 - Latsar CPNS'), 'angkatan_1_latsar_cpns');
});

test('slug_ - special characters', () => {
  assert.strictEqual(slug_('Unit/Bagian & Kepegawaian!'), 'unit_bagian_kepegawaian');
});

// 9. normalizeSheetName_
test('normalizeSheetName_ - remove invalid characters', () => {
  assert.strictEqual(normalizeSheetName_('Daftar [Arsip] : Aktif / Inaktif?'), 'Daftar Arsip Aktif Inaktif');
});

test('normalizeSheetName_ - character length limit', () => {
  const longName = 'A'.repeat(120);
  assert.strictEqual(normalizeSheetName_(longName).length, 90);
});

// 10. Edge cases
test('extractNomorSurat_ - empty string returns empty', () => {
  assert.strictEqual(extractNomorSurat_(''), '');
});

test('extractNomorSurat_ - no match returns empty', () => {
  assert.strictEqual(extractNomorSurat_('Plain text without number'), '');
});

test('extractKodeKlasifikasi_ - empty string returns empty', () => {
  assert.strictEqual(extractKodeKlasifikasi_(''), '');
});

test('extractKodeKlasifikasi_ - no match returns empty', () => {
  assert.strictEqual(extractKodeKlasifikasi_('No code here'), '');
});

test('extractDate_ - empty string returns empty', () => {
  assert.strictEqual(extractDate_(''), '');
});

test('extractDate_ - slash separated date', () => {
  assert.strictEqual(extractDate_('2026/05/30'), '2026-05-30');
});

test('extractTingkatPerkembangan_ - cetak menjadi Salinan', () => {
  assert.strictEqual(extractTingkatPerkembangan_('Cetak_Surat.pdf'), 'Salinan');
});

test('extractTingkatPerkembangan_ - empty string returns empty', () => {
  assert.strictEqual(extractTingkatPerkembangan_(''), '');
});

test('extractUraian_ - Hal match', () => {
  const text = 'Hal: Undangan Rapat Evaluasi';
  const sourceName = 'surat.pdf';
  assert.strictEqual(extractUraian_(text, sourceName, { activity_name: 'A' }, { sub_activity_name: 'B' }), 'Undangan Rapat Evaluasi');
});

test('extractUraian_ - empty everything returns fallback', () => {
  assert.strictEqual(extractUraian_('', '', { activity_name: 'A' }, { sub_activity_name: 'B' }), 'Surat - A - B');
});

test('cleanUraian_ - trims and limits to 260 chars', () => {
  const long = 'x'.repeat(300);
  assert.strictEqual(cleanUraian_(long).length, 260);
});

test('cleanUraian_ - removes leading/trailing underscores', () => {
  assert.strictEqual(cleanUraian_('___hello___'), 'hello');
});

test('buildFinalFileName_ - handles 0 as nomor_item_arsip', () => {
  const metadata = {
    nomor_item_arsip: '0',
    tingkat_perkembangan: 'Asli',
    uraian_informasi_item: 'Test'
  };
  assert.strictEqual(buildFinalFileName_(metadata, 'file.pdf'), '00. (Asli) Test.pdf');
});

test('buildFinalFileName_ - no source extension defaults to pdf', () => {
  const metadata = {
    nomor_item_arsip: '1',
    tingkat_perkembangan: 'Copy',
    uraian_informasi_item: 'Doc'
  };
  assert.strictEqual(buildFinalFileName_(metadata, 'file'), '01. (Copy) Doc.pdf');
});

test('pad2_ - zero value', () => {
  assert.strictEqual(pad2_(0), '00');
});

test('pad2_ - null value', () => {
  assert.strictEqual(pad2_(null), '00');
});

test('pad2_ - single digit', () => {
  assert.strictEqual(pad2_('5'), '05');
});

test('pad2_ - double digit stays same', () => {
  assert.strictEqual(pad2_('12'), '12');
});

test('cleanId_ - URL with long path', () => {
  const url = 'https://drive.google.com/open?id=1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef';
  assert.strictEqual(cleanId_(url), '1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef');
});

test('cleanId_ - empty string', () => {
  assert.strictEqual(cleanId_(''), '');
});

test('isTrue_ - TRUE is true', () => {
  assert.strictEqual(isTrue_('TRUE'), true);
});

test('isTrue_ - FALSE is false', () => {
  assert.strictEqual(isTrue_('FALSE'), false);
});

test('isTrue_ - lowercase true is true', () => {
  assert.strictEqual(isTrue_('true'), true);
});

test('isTrue_ - boolean true is true', () => {
  assert.strictEqual(isTrue_(true), true);
});

test('slug_ - leading/trailing special chars stripped', () => {
  assert.strictEqual(slug_('__hello__'), 'hello');
});

test('slug_ - empty string returns empty', () => {
  assert.strictEqual(slug_(''), '');
});

test('normalizeSheetName_ - empty string defaults', () => {
  assert.strictEqual(normalizeSheetName_(''), 'Sheet');
});

test('normalizeSheetName_ - colons and brackets removed', () => {
  assert.strictEqual(normalizeSheetName_('Test [Sheet]: 2026'), 'Test Sheet 2026');
});

// 11. parseExistingFileName_
test('parseExistingFileName_ - standard compliant format with No', () => {
  const result = parseExistingFileName_('02. (Asli) No: B-123/2026_Undangan Rapat.pdf');
  assert.strictEqual(result.nomor_item_arsip, '02');
  assert.strictEqual(result.tingkat_perkembangan, 'Salinan'); // legacy Asli -> Salinan
  assert.strictEqual(result.nomor_surat, 'B-123/2026');
  assert.strictEqual(result.uraian_informasi_item, 'Undangan Rapat');
});

test('parseExistingFileName_ - standard compliant format without No', () => {
  const result = parseExistingFileName_('15. (Copy) Laporan Bulanan.docx');
  assert.strictEqual(result.nomor_item_arsip, '15');
  assert.strictEqual(result.tingkat_perkembangan, 'Salinan'); // legacy Copy -> Salinan
  assert.strictEqual(result.nomor_surat, '');
  assert.strictEqual(result.uraian_informasi_item, 'Laporan Bulanan');
});

test('parseExistingFileName_ - simple number prefix format without parenthesis', () => {
  const result = parseExistingFileName_('03. Surat Tugas Panitia.pdf');
  assert.strictEqual(result.nomor_item_arsip, '03');
  assert.strictEqual(result.tingkat_perkembangan, 'Asli'); // default fallback
  assert.strictEqual(result.nomor_surat, '');
  assert.strictEqual(result.uraian_informasi_item, 'Surat Tugas Panitia');
});

test('parseExistingFileName_ - simple number prefix format single digit', () => {
  const result = parseExistingFileName_('5. Kuitansi Pembayaran.docx');
  assert.strictEqual(result.nomor_item_arsip, '05');
  assert.strictEqual(result.tingkat_perkembangan, 'Asli');
  assert.strictEqual(result.nomor_surat, '');
  assert.strictEqual(result.uraian_informasi_item, 'Kuitansi Pembayaran');
});

test('parseExistingFileName_ - non-numbered format fallback', () => {
  const result = parseExistingFileName_('Dokumen Kegiatan Diklat.pdf');
  assert.strictEqual(result.nomor_item_arsip, '01');
  assert.strictEqual(result.tingkat_perkembangan, 'Asli');
  assert.strictEqual(result.nomor_surat, '');
  assert.strictEqual(result.uraian_informasi_item, 'Dokumen Kegiatan Diklat');
});

test('parseExistingFileName_ - legacy Srikandi menjadi Asli', () => {
  const result = parseExistingFileName_('06. (Srikandi) No: 85/P.3/PDP.07.1_Surat Pemanggilan.pdf');
  assert.strictEqual(result.tingkat_perkembangan, 'Asli');
});

// 12. isEmpty_
test('isEmpty_ - null is empty', () => {
  assert.strictEqual(isEmpty_(null), true);
});

test('isEmpty_ - undefined is empty', () => {
  assert.strictEqual(isEmpty_(undefined), true);
});

test('isEmpty_ - empty string is empty', () => {
  assert.strictEqual(isEmpty_(''), true);
});

test('isEmpty_ - whitespace-only string is empty', () => {
  assert.strictEqual(isEmpty_('   '), true);
});

test('isEmpty_ - non-empty string is not empty', () => {
  assert.strictEqual(isEmpty_('hello'), false);
});

test('isEmpty_ - zero is not empty', () => {
  assert.strictEqual(isEmpty_(0), false);
});

test('isEmpty_ - false is not empty', () => {
  assert.strictEqual(isEmpty_(false), false);
});

// 13. safeJsonParse_
test('safeJsonParse_ - valid JSON object', () => {
  assert.deepStrictEqual(safeJsonParse_('{"a":1}'), { a: 1 });
});

test('safeJsonParse_ - valid JSON array', () => {
  assert.deepStrictEqual(safeJsonParse_('[1,2,3]'), [1, 2, 3]);
});

test('safeJsonParse_ - invalid JSON returns fallback', () => {
  assert.strictEqual(safeJsonParse_('not json', 'fallback'), 'fallback');
});

test('safeJsonParse_ - invalid JSON returns undefined fallback', () => {
  assert.strictEqual(safeJsonParse_('bad'), undefined);
});

test('safeJsonParse_ - null string returns null', () => {
  assert.strictEqual(safeJsonParse_('null'), null);
});

test('safeJsonParse_ - empty object', () => {
  assert.deepStrictEqual(safeJsonParse_('{}'), {});
});

// 14. escapeDriveQueryValue_
test('escapeDriveQueryValue_ - plain string unchanged', () => {
  assert.strictEqual(escapeDriveQueryValue_('hello'), 'hello');
});

test('escapeDriveQueryValue_ - single quotes escaped', () => {
  assert.strictEqual(escapeDriveQueryValue_("it's"), "it\\'s");
});

test('escapeDriveQueryValue_ - backslashes escaped', () => {
  assert.strictEqual(escapeDriveQueryValue_('path\\to'), 'path\\\\to');
});

test('escapeDriveQueryValue_ - both backslash and quote', () => {
  assert.strictEqual(escapeDriveQueryValue_("it\\'s"), "it\\\\\\'s");
});

test('escapeDriveQueryValue_ - empty string', () => {
  assert.strictEqual(escapeDriveQueryValue_(''), '');
});

test('escapeDriveQueryValue_ - null/undefined returns empty', () => {
  assert.strictEqual(escapeDriveQueryValue_(null), '');
  assert.strictEqual(escapeDriveQueryValue_(undefined), '');
});

// 15. normalizeHexColor_
test('normalizeHexColor_ - valid hex passes through', () => {
  assert.strictEqual(normalizeHexColor_('#ff0000'), '#FF0000');
});

test('normalizeHexColor_ - lowercase hex normalized to upper', () => {
  assert.strictEqual(normalizeHexColor_('#abcdef'), '#ABCDEF');
});

test('normalizeHexColor_ - invalid hex returns default', () => {
  assert.strictEqual(normalizeHexColor_('not-a-color'), '#2563EB');
});

test('normalizeHexColor_ - empty string returns default', () => {
  assert.strictEqual(normalizeHexColor_(''), '#2563EB');
});

test('normalizeHexColor_ - invalid with fallback uses fallback', () => {
  assert.strictEqual(normalizeHexColor_('bad', '#112233'), '#112233');
});

test('normalizeHexColor_ - shorthand hex invalid returns default', () => {
  assert.strictEqual(normalizeHexColor_('#fff'), '#2563EB');
});

test('normalizeHexColor_ - mixed case already valid', () => {
  assert.strictEqual(normalizeHexColor_('#aBcDeF'), '#ABCDEF');
});

// 16. Additional edge cases for existing functions
test('cleanId_ - special characters stripped from URL', () => {
  const url = 'https://drive.google.com/file/d/1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef/view?usp=sharing';
  assert.strictEqual(cleanId_(url), '1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890abcdef');
});

test('cleanId_ - short ID returned as-is', () => {
  assert.strictEqual(cleanId_('short'), 'short');
});

test('isTrue_ - number 1 is false', () => {
  assert.strictEqual(isTrue_(1), false);
});

test('isTrue_ - empty string is false', () => {
  assert.strictEqual(isTrue_(''), false);
});

test('pad2_ - undefined returns 00', () => {
  assert.strictEqual(pad2_(undefined), '00');
});

test('pad2_ - triple digit stays same', () => {
  assert.strictEqual(pad2_('123'), '123');
});

test('slug_ - numbers preserved', () => {
  assert.strictEqual(slug_('Kegiatan 2026'), 'kegiatan_2026');
});

test('sanitizeFilePart_ - removes forbidden characters', () => {
  assert.strictEqual(sanitizeFilePart_('file:name*with?bad"chars<>|here'), 'filenamewithbadcharshere');
});

test('sanitizeFilePart_ - collapses multiple spaces', () => {
  assert.strictEqual(sanitizeFilePart_('too   many   spaces'), 'too many spaces');
});

// 17. Smarter extractNomorSurat_ patterns
test('extractNomorSurat_ - Nomor XX Tahun YYYY format', () => {
  assert.strictEqual(extractNomorSurat_('Nomor 45 Tahun 2025'), '45/Tahun/2025');
});

test('extractNomorSurat_ - No. XX Tahun YYYY format', () => {
  assert.strictEqual(extractNomorSurat_('No. 12 Tahun 2026'), '12/Tahun/2026');
});

test('extractNomorSurat_ - SP- prefix Surat Perintah', () => {
  assert.strictEqual(extractNomorSurat_('SP-123/SDM.01/2025'), 'SP-123/SDM.01/2025');
});

test('extractNomorSurat_ - R- prefix with trailing segment', () => {
  assert.strictEqual(extractNomorSurat_('R-456/KP.02/2026/Kepegawaian'), 'R-456/KP.02/2026/Kepegawaian');
});

test('extractNomorSurat_ - ignores 4-digit numbers that are not years', () => {
  assert.strictEqual(extractNomorSurat_('Kode pos 1234 dan angka 5678'), '');
});

test('extractNomorSurat_ - space separated from filename', () => {
  assert.strictEqual(extractNomorSurat_('No B 417 BKPSDM 800.2 12 2025 Permohonan'), 'B/417/BKPSDM/800.2/12/2025');
});

// 18. Smarter extractKodeKlasifikasi_ patterns
test('extractKodeKlasifikasi_ - contextual "Kode:" prefix', () => {
  assert.strictEqual(extractKodeKlasifikasi_('Kode: KP.01.02 tentang kepegawaian'), 'KP.01.02');
});

test('extractKodeKlasifikasi_ - contextual "Klasifikasi:" prefix', () => {
  assert.strictEqual(extractKodeKlasifikasi_('Klasifikasi: DL.01 tentang pelatihan'), 'DL.01');
});

test('extractKodeKlasifikasi_ - single letter code K.01', () => {
  assert.strictEqual(extractKodeKlasifikasi_('Kode K.01.03 tentang umum'), 'K.01.03');
});

test('extractKodeKlasifikasi_ - 3-level code KP.01.02', () => {
  assert.strictEqual(extractKodeKlasifikasi_('Berkas dengan kode KP.01.02 tentang kepegawaian'), 'KP.01.02');
});

test('extractKodeKlasifikasi_ - ignores bare 3 digit numbers', () => {
  assert.strictEqual(extractKodeKlasifikasi_('Tahun 595 yang lalu'), '');
});

test('extractKodeKlasifikasi_ - accepts 3 digit with dots', () => {
  assert.strictEqual(extractKodeKlasifikasi_('Surat 800.2.2 perihal'), '800.2.2');
});

// 19. Smarter extractDate_ patterns
test('extractDate_ - contextual "ditetapkan" near Indonesian date', () => {
  assert.strictEqual(extractDate_('Ditetapkan di Jakarta pada tanggal 15 Maret 2025'), '2025-03-15');
});

test('extractDate_ - contextual "tanggal" near ISO date', () => {
  assert.strictEqual(extractDate_('tanggal: 2025-06-30'), '2025-06-30');
});

test('extractDate_ - month + year only defaults to 1st', () => {
  assert.strictEqual(extractDate_('Dokumen ini berlaku mulai Maret 2025'), '2025-03-01');
});

test('extractDate_ - DD.MM.YYYY dot-separated format', () => {
  assert.strictEqual(extractDate_('Surat bertanggal 15.03.2025 dikirim'), '2025-03-15');
});

test('extractDate_ - DD-MM-YYYY dash format', () => {
  assert.strictEqual(extractDate_('Ditandatangani 30-05-2026'), '2026-05-30');
});

test('extractDate_ - contextual "ditandatangani" near date', () => {
  assert.strictEqual(extractDate_('Ditandatangani pada 10 Januari 2026'), '2026-01-10');
});

// 20. Smarter extractKlasifikasiAkses_ patterns
test('extractKlasifikasiAkses_ - contextual "bersifat Rahasia"', () => {
  assert.strictEqual(extractKlasifikasiAkses_('Dokumen ini bersifat Rahasia'), 'Rahasia');
});

test('extractKlasifikasiAkses_ - contextual "klasifikasi akses: Biasa"', () => {
  assert.strictEqual(extractKlasifikasiAkses_('klasifikasi akses: Biasa'), 'Biasa');
});

test('extractKlasifikasiAkses_ - contextual "sifat dokumen: Terbatas"', () => {
  assert.strictEqual(extractKlasifikasiAkses_('sifat dokumen: Terbatas'), 'Terbatas');
});

test('extractKlasifikasiAkses_ - contextual Terbuka maps to Biasa', () => {
  assert.strictEqual(extractKlasifikasiAkses_('bersifat Terbuka untuk umum'), 'Biasa');
});

// 21. Smarter extractUraian_ patterns
test('extractUraian_ - Tentang keyword', () => {
  const text = 'KEPUTUSAN\nTentang\nPenetapan Penerima Beasiswa\nMenimbang';
  assert.strictEqual(extractUraian_(text, 'sk.pdf', { activity_name: 'A' }, { sub_activity_name: 'B' }), 'Penetapan Penerima Beasiswa');
});

test('extractUraian_ - multi-line Perihal stops at next section', () => {
  const text = 'Perihal: Undangan Rapat Koordinasi\nKepada:\nYth. Kepala Bagian';
  assert.strictEqual(extractUraian_(text, 'surat.pdf', { activity_name: 'A' }, { sub_activity_name: 'B' }), 'Undangan Rapat Koordinasi');
});

test('extractUraian_ - Perihal stops at double newline', () => {
  const text = 'Perihal: Evaluasi Kinerja Tahunan\n\nIsi surat dimulai di sini';
  assert.strictEqual(extractUraian_(text, 'surat.pdf', { activity_name: 'A' }, { sub_activity_name: 'B' }), 'Evaluasi Kinerja Tahunan');
});

// 22. ParseEngine: Document Type Classification
test('classifyDocumentType_ - Surat Keputusan', () => {
  const text = 'SURAT KEPUTUSAN\nNomor: 123/2025\nMenimbang:\nMengingat:\nMemutuskan:\nMenetapkan:';
  assert.strictEqual(classifyDocumentType_(text, 'SK.pdf'), 'Surat Keputusan');
});

test('classifyDocumentType_ - Surat Undangan', () => {
  const text = 'SURAT UNDANGAN\nNomor: 456/2025\nDengan ini mengharapkan kehadiran Bapak/Ibu';
  assert.strictEqual(classifyDocumentType_(text, 'undangan.pdf'), 'Surat Undangan');
});

test('classifyDocumentType_ - Berita Acara', () => {
  const text = 'BERITA ACARA\nPada hari ini Senin tanggal 15 Maret 2025, kami yang bertanda tangan di bawah ini';
  assert.strictEqual(classifyDocumentType_(text, 'BA.pdf'), 'Berita Acara');
});

test('classifyDocumentType_ - unknown type returns empty', () => {
  const text = 'Dokumen biasa tanpa indikator tipe khusus';
  assert.strictEqual(classifyDocumentType_(text, 'dokumen.pdf'), '');
});

test('classifyDocumentType_ - detects from filename', () => {
  const text = 'Nomor: 123/2025';
  assert.strictEqual(classifyDocumentType_(text, 'Surat_Undangan_Rapat.pdf'), 'Surat Undangan');
});

test('classifyDocumentType_ - Perjanjian', () => {
  const text = 'PERJANJIAN KERJA SAMA\nPIHAK PERTAMA:\nPIHAK KEDUA:\nPasal 1\nPasal 2';
  assert.strictEqual(classifyDocumentType_(text, 'PKS.pdf'), 'Perjanjian');
});

// 23. ParseEngine: OCR Text Normalization
test('normalizeTextPE_ - collapses multiple spaces', () => {
  assert.strictEqual(normalizeTextPE_('hello    world'), 'hello world');
});

test('normalizeTextPE_ - normalizes CRLF to LF', () => {
  assert.strictEqual(normalizeTextPE_('line1\r\nline2'), 'line1\nline2');
});

test('normalizeTextPE_ - replaces smart quotes', () => {
  assert.strictEqual(normalizeTextPE_('\u201CHello\u201D'), '"Hello"');
});

test('normalizeTextPE_ - replaces em dash', () => {
  assert.strictEqual(normalizeTextPE_('word\u2014word'), 'word-word');
});

test('normalizeTextPE_ - removes zero-width chars', () => {
  assert.strictEqual(normalizeTextPE_('hello\u200Bworld'), 'helloworld');
});

test('normalizeTextPE_ - replaces bullet with dash', () => {
  assert.strictEqual(normalizeTextPE_('\u2022 item one'), '- item one');
});

test('normalizeTextPE_ - replaces pipe with I', () => {
  assert.strictEqual(normalizeTextPE_('No|123'), 'NoI123');
});

test('normalizeTextPE_ - empty string stays empty', () => {
  assert.strictEqual(normalizeTextPE_(''), '');
});

test('normalizeTextPE_ - replaces NBSP with space', () => {
  assert.strictEqual(normalizeTextPE_('hello\u00A0world'), 'hello world');
});

console.log(`\nTests finished: ${passedCount} passed, ${failedCount} failed.`);
if (failedCount > 0) {
  process.exit(1);
}

