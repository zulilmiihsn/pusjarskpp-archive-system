const assert = require('assert');
const { loadGasScripts, resetMocks, Properties, mockFiles } = require('./gas-mocks.js');

console.log('Running Integration Tests with GAS Mocks...\n');

let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    resetMocks();
    fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(err);
    failed++;
  }
}

// 1. Load the scripts into global scope
try {
  loadGasScripts();
  console.log('GAS Scripts loaded successfully.');
} catch (e) {
  console.error('Failed to load GAS scripts:', e);
  process.exit(1);
}

// --- Auth guard regression tests (hasil audit hardening 2026-06) ---

// Buat sesi valid langsung di Script Properties (meniru output login AuthService).
function makeSession(role) {
  const id = 'test-session-' + role;
  Properties['sess_' + id] = JSON.stringify({
    accountId: 'acc-' + role,
    username: role === 'admin' ? 'admin' : 'user',
    role: role,
    displayName: role,
    expiresAt: Date.now() + 60 * 60 * 1000
  });
  return id;
}

test('requireAuth_ - tolak tanpa _sessionId', () => {
  assert.throws(() => requireAuth_({}), /Sesi login tidak ditemukan/);
});

test('requireAuth_ - tolak sesi yang sudah kedaluwarsa', () => {
  Properties['sess_expired'] = JSON.stringify({ role: 'admin', expiresAt: Date.now() - 1000 });
  assert.throws(() => requireAuth_({ _sessionId: 'expired' }), /berakhir/);
});

test('requireAuth_ - lolos untuk sesi valid & kembalikan role', () => {
  const sid = makeSession('user');
  const user = requireAuth_({ _sessionId: sid });
  assert.strictEqual(user.role, 'user');
});

test('requireAdmin_ - tolak sesi non-admin', () => {
  const sid = makeSession('user');
  assert.throws(() => requireAdmin_({ _sessionId: sid }), /Hanya admin/);
});

test('requireAdmin_ - lolos untuk sesi admin', () => {
  const sid = makeSession('admin');
  const user = requireAdmin_({ _sessionId: sid });
  assert.strictEqual(user.role, 'admin');
});

test('B1 - deleteArchive wajib admin (guard terpasang, tak menyentuh data tanpa auth)', () => {
  // Tanpa sesi: harus throw di guard SEBELUM operasi apa pun.
  assert.throws(() => ArchiveController.deleteArchive({ archiveId: 'ARC-1' }), /Sesi login|admin/i);
  // Sesi user biasa: tetap ditolak (hapus = admin-only).
  const sid = makeSession('user');
  assert.throws(() => ArchiveController.deleteArchive({ _sessionId: sid, archiveId: 'ARC-1' }), /Hanya admin/);
});

test('B1 - deleteSubActivity & trashArchiveFolder admin-only', () => {
  const sid = makeSession('user');
  assert.throws(() => SubActivityController.deleteSubActivity({ _sessionId: sid, activityId: 'a', subActivityId: 's' }), /Hanya admin/);
  assert.throws(() => DriveController.trashArchiveFolder({ _sessionId: sid, folderId: 'f' }), /Hanya admin/);
});

test('B6 - requireAdminIfWorkspaceSecured_ fail-closed saat baca akun gagal', () => {
  // Workspace dianggap terpasang (configSpreadsheetId ada) tapi spreadsheet tak bisa dibuka
  // -> listAccounts melempar -> harus DITOLAK (throw), bukan diloloskan (return null).
  Properties['CONFIG_SPREADSHEET_ID'] = 'ss-yang-tidak-ada';
  const sid = makeSession('admin');
  assert.throws(() => requireAdminIfWorkspaceSecured_({ _sessionId: sid }), /Akses ditolak/);
});

// --- ParseEngine: deteksi arah surat masuk/keluar (autofill) ---

test('ParseEngine.detectDirection - KOP Lembaga Administrasi Negara = keluar', () => {
  assert.strictEqual(ParseEngine.detectDirection('PEMERINTAH\nLembaga Administrasi Negara\nNomor: 1', ''), 'keluar');
});

test('ParseEngine.detectDirection - toleran spasi/case OCR', () => {
  assert.strictEqual(ParseEngine.detectDirection('lembaga  administrasi   negara', ''), 'keluar');
});

test('ParseEngine.detectDirection - tanpa KOP itu = masuk', () => {
  assert.strictEqual(ParseEngine.detectDirection('Dinas Pendidikan Provinsi\nNomor: 2', ''), 'masuk');
});

test('ParseEngine.analyze - surat MASUK tidak mengisi kode_klasifikasi', () => {
  const txt = 'Dinas X\nNomor: 100/AB.02/2025\nKode: KP.01.02\nPerihal: Undangan rapat\n\nIsi surat...';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.strictEqual(r.documentDirection, 'masuk');
  assert.ok(!r.fields.kode_klasifikasi, 'kode_klasifikasi harus kosong untuk surat masuk');
});

test('ParseEngine.analyze - surat KELUAR boleh mengisi kode_klasifikasi', () => {
  const txt = 'Lembaga Administrasi Negara\nNomor: 100/AB.02/2025\nKode: KP.01.02\nPerihal: Undangan rapat\n\nIsi surat...';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.strictEqual(r.documentDirection, 'keluar');
  assert.ok(r.fields.kode_klasifikasi && r.fields.kode_klasifikasi.value, 'kode_klasifikasi harus terisi untuk surat keluar');
});

test('ParseEngine.analyze - segmen PDP di nomor: nomor tetap FULL, kode = segmen PDP', () => {
  const txt = 'Dinas X\nNomor: 273/P.3/PDP.07.1\nPerihal: Surat pemanggilan\n\nIsi surat...';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.strictEqual(r.documentDirection, 'keluar', 'segmen PDP harus memaksa arah keluar');
  assert.strictEqual(r.fields.nomor_surat.value, '273/P.3/PDP.07.1', 'Nomor Surat harus tetap full, tidak dipotong');
  assert.ok(r.fields.kode_klasifikasi, 'kode_klasifikasi harus terisi');
  assert.strictEqual(r.fields.kode_klasifikasi.value, 'PDP.07.1', 'kode_klasifikasi hanya segmen PDP');
});

test('ParseEngine.analyze - tanggal dateline header terbaca', () => {
  const txt = 'Lembaga Administrasi Negara\nNomor: 273/P.3/PDP.07.1\nSamarinda, 13 Desember 2025\nPerihal: Undangan\n\nIsi\n\nKepala\nBudi';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.ok(r.fields.tanggal, 'tanggal harus terisi');
  assert.strictEqual(r.fields.tanggal.value, '2025-12-13');
});

test('ParseEngine.analyze - tanggal berkata-kunci di body terbaca', () => {
  const txt = 'Dinas X\nNomor: 100/AB.02/2025\nPerihal: Rapat\n\nPada tanggal 17 Maret 2026 akan diadakan rapat.\n\nHormat kami,\nKepala';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.ok(r.fields.tanggal, 'tanggal di body (berkata-kunci) harus terbaca');
  assert.strictEqual(r.fields.tanggal.value, '2026-03-17');
});

test('ParseEngine.analyze - dateline kota di body walau tanpa blok tanda tangan', () => {
  const txt = 'Pemkot\nNomor: 5/X/2026\nPerihal: Edaran\nYth. Bapak\nSamarinda, 1 April 2026\n\nIsi surat panjang sekali di sini ya.';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.ok(r.fields.tanggal, 'dateline kota harus terbaca walau body tak punya tanda tangan');
  assert.strictEqual(r.fields.tanggal.value, '2026-04-01');
});

test('ParseEngine.analyze - dateline header menang atas tanggal lain di isi', () => {
  const txt = 'Pemkot\nNomor: 9/IX/2026\nSamarinda, 5 Februari 2026\nPerihal: Lap\nYth Bapak\n\nKegiatan berlangsung 20 Januari 2026 hingga selesai dengan lancar.';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.ok(r.fields.tanggal, 'tanggal harus terisi');
  assert.strictEqual(r.fields.tanggal.value, '2026-02-05', 'dateline (5 Feb) harus menang, bukan tanggal di isi (20 Jan)');
});

test('ParseEngine.analyze - tanggal polos di body (tanda tangan gambar) tetap terbaca', () => {
  // Tak ada blok tanda tangan teks; dateline polos jatuh ke body.
  const txt = 'Dinas X\nNomor: 100/AB.02/2025\nPerihal: Undangan\nYth. Bapak\n\nIsi undangan rapat di sini.\n\n13 Desember 2025';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.ok(r.fields.tanggal, 'tanggal polos di body harus terbaca');
  assert.strictEqual(r.fields.tanggal.value, '2025-12-13');
});

// --- Ekstraksi isi dokumen via konversi (OCR/convert) untuk autofill ---

function fakeFile(id, sizeBytes) {
  return { getId: () => id, getSize: () => sizeBytes || 1000, getName: () => 'surat.pdf' };
}

test('extractTextViaConversion_ - PDF: OCR id, baca teks, hapus doc sementara', () => {
  global._driveMock.copyResult = { id: 'tmp-doc-1' };
  global._driveMock.exportText = 'LAN RI Kalimantan Timur\nNomor: 100/AB.02/2025';
  const out = extractTextViaConversion_(fakeFile('file-1'), 'application/pdf');
  assert.strictEqual(out.method, 'ocr_id');
  assert.ok(out.text.indexOf('Kalimantan Timur') >= 0, 'harus mengembalikan teks hasil export');
  assert.deepStrictEqual(global._driveMock.removed, ['tmp-doc-1'], 'doc sementara wajib dihapus');
});

test('extractTextViaConversion_ - Word: convert tanpa OCR', () => {
  global._driveMock.copyResult = { id: 'tmp-doc-2' };
  global._driveMock.exportText = 'Isi dokumen word';
  const out = extractTextViaConversion_(fakeFile('file-2'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  assert.strictEqual(out.method, 'convert');
  assert.ok(out.text.indexOf('word') >= 0);
});

test('extractTextViaConversion_ - file terlalu besar dilewati (jaga kuota)', () => {
  const out = extractTextViaConversion_(fakeFile('file-3', 30 * 1024 * 1024), 'application/pdf');
  assert.strictEqual(out.method, 'skipped_too_large');
  assert.strictEqual(out.text, '');
});

// --- Hitung halaman PDF (jumlah lembar) ---

test('countPagesFromPdfChunk_ - ambil /Count milik /Pages, bukan outline', () => {
  // Outline /Count 99 muncul DULUAN; page-tree /Count 12 yang benar.
  const pdf = '/Type /Outlines /Count 99 /First 5 0 R\n<< /Type /Pages /Kids [3 0 R] /Count 12 >>';
  assert.strictEqual(countPagesFromPdfChunk_(pdf, true), 12);
});

test('countPagesFromPdfChunk_ - urutan key /Count sebelum /Type /Pages', () => {
  const pdf = '<< /Count 7 /Kids [1 0 R] /Type /Pages >>';
  assert.strictEqual(countPagesFromPdfChunk_(pdf, true), 7);
});

test('countPagesFromPdfChunk_ - PDF merge: ambil count terbesar (root), bukan sub-tree', () => {
  const pdf = '<< /Type /Pages /Count 10 >> ... << /Type /Pages /Count 50 /Kids [..] >>';
  assert.strictEqual(countPagesFromPdfChunk_(pdf, true), 50);
});

test('countPagesFromPdfChunk_ - fallback hitung /Type /Page hanya bila isi penuh', () => {
  const pdf = '/Type /Page x /Type /Page y /Type /Page z';
  assert.strictEqual(countPagesFromPdfChunk_(pdf, true), 3, 'isi penuh: boleh hitung /Type /Page');
  assert.strictEqual(countPagesFromPdfChunk_(pdf, false), 0, 'potongan: jangan hitung /Type /Page (undercount)');
});

test('countPagesFromPdfChunk_ - /Pages tak ketemu, fallback ke /Count apa pun', () => {
  const pdf = 'no page tree here /Count 4 somewhere';
  assert.strictEqual(countPagesFromPdfChunk_(pdf, false), 4);
});

console.log(`\nIntegration Tests Finished: ${passed} passed, ${failed} failed.`);
if (failed > 0) {
  process.exit(1);
}
