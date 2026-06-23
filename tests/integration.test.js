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

test('ParseEngine.detectDirection - KOP LAN RI Kalimantan Timur = keluar', () => {
  assert.strictEqual(ParseEngine.detectDirection('PEMERINTAH\nLAN RI Kalimantan Timur\nNomor: 1', ''), 'keluar');
});

test('ParseEngine.detectDirection - toleran spasi/case OCR', () => {
  assert.strictEqual(ParseEngine.detectDirection('lan  ri   kalimantan timur', ''), 'keluar');
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
  const txt = 'LAN RI Kalimantan Timur\nNomor: 100/AB.02/2025\nKode: KP.01.02\nPerihal: Undangan rapat\n\nIsi surat...';
  const r = ParseEngine.analyze(txt, 'surat.pdf', {});
  assert.strictEqual(r.documentDirection, 'keluar');
  assert.ok(r.fields.kode_klasifikasi && r.fields.kode_klasifikasi.value, 'kode_klasifikasi harus terisi untuk surat keluar');
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

console.log(`\nIntegration Tests Finished: ${passed} passed, ${failed} failed.`);
if (failed > 0) {
  process.exit(1);
}
