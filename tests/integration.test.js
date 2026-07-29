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

test('sha256Hex_ - kompatibel dengan SHA-256 standar', () => {
  assert.strictEqual(
    sha256Hex_('abc'),
    'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'
  );
});

test('workspace archive sharing - pemegang link hanya Viewer', () => {
  const sharingCalls = [];
  const folder = {
    getName: () => '1. Daftar Arsip (Spreadsheet)',
    getSharingAccess: () => DriveApp.Access.PRIVATE,
    getSharingPermission: () => null,
    setSharing: (access, permission) => sharingCalls.push({ access, permission })
  };
  const report = [];

  wsEnsureAnyoneWithLinkViewer_(folder, report);

  assert.deepStrictEqual(sharingCalls, [{
    access: DriveApp.Access.ANYONE_WITH_LINK,
    permission: DriveApp.Permission.VIEW
  }]);
  assert.ok(report.some(item => /Anyone with link \(Viewer\)/.test(item.label)));
});

test('pbkdf2Like_ - hasil iterasi kompatibel dengan hash lama', () => {
  assert.strictEqual(
    pbkdf2Like_('pass', 'Admin', 'abc', 3),
    'dd70a2b952cafc7a7448224eaf98fb0ae66a0ab82077bf5080a86bace074f460abc'
  );
});

test('session helper - simpan durable dan cache baca cepat', () => {
  const session = { username: 'admin', role: 'admin', expiresAt: Date.now() + 60000 };
  saveSession_('cached-session', session);
  assert.ok(Properties['sess_cached-session']);
  assert.strictEqual(loadSession_('cached-session').username, 'admin');
  deleteSession_('cached-session');
  assert.strictEqual(loadSession_('cached-session'), null);
});

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

test('requireAuth_ - tidak mengikat sesi portal ke email Google', () => {
  Properties.sess_cross_google = JSON.stringify({
    accountId: 'acc-admin',
    username: 'admin',
    role: 'admin',
    displayName: 'Administrator',
    activeEmail: 'akun-lama@example.net',
    expiresAt: Date.now() + 60000
  });
  assert.strictEqual(requireAuth_({ _sessionId: 'cross_google' }).username, 'admin');
});

test('akses role ditolak bersifat nonfatal dan menyebut identitas portal', () => {
  const sid = makeSession('user');
  const response = wrapApi(() => requireAdmin_({ _sessionId: sid }));
  assert.strictEqual(response.success, false);
  assert.strictEqual(response.errorCode, 'ACCESS_DENIED');
  assert.match(response.error, /username portal "user"/);
  assert.ok(Properties['sess_' + sid], 'sesi tidak boleh dihapus saat akses ditolak');
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

test('deleteArchive - role user boleh menghapus melalui app, guest ditolak', () => {
  // Tanpa sesi: harus throw di guard SEBELUM operasi apa pun.
  assert.throws(() => ArchiveController.deleteArchive({ archiveId: 'ARC-1' }), /Sesi login/i);

  const originalDeleteCore = ArchiveController._deleteArchiveCore_;
  ArchiveController._deleteArchiveCore_ = () => ({ success: true });
  const sid = makeSession('user');
  try {
    assert.deepStrictEqual(
      ArchiveController.deleteArchive({ _sessionId: sid, archiveId: 'ARC-1', year: 2026 }),
      { success: true }
    );
  } finally {
    ArchiveController._deleteArchiveCore_ = originalDeleteCore;
  }

  const guestSid = makeSession('guest');
  assert.throws(
    () => ArchiveController.deleteArchive({ _sessionId: guestSid, archiveId: 'ARC-1' }),
    /Hanya petugas atau admin/
  );
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

function makeRekapLookup(rows) {
  return {
    startRow: 8,
    noteRow: 8 + rows.length,
    nomorBerkasCol: 2,
    uraianCol: 4,
    noLaciCol: 8,
    noFolderCol: 9,
    values: rows.map(row => row.values),
    formulas: rows.map(row => row.formulas || Array(9).fill('')),
    notes: rows.map(row => row.notes || Array(9).fill(''))
  };
}

test('rekap lookup - folder_id menang atas rekap_row_number basi dan nama duplikat', () => {
  const duplicateFormalName = 'Pelatihan Kepemimpinan Pengawas Angkatan 3 Tahun 2026';
  const lookup = makeRekapLookup([
    {
      values: ['', '7', '', duplicateFormalName, '', '', '', 'PKP Angkatan 1 Tahun 2026'],
      formulas: ['', '', '', '', '', '', '', '=HYPERLINK("https://drive.google.com/drive/folders/folder-1"; "PKP Angkatan 1")']
    },
    {
      values: ['', '7', '', duplicateFormalName, '', '', '', 'PKP Angkatan 3 Tahun 2026'],
      formulas: ['', '', '', '', '', '', '', '=HYPERLINK("https://drive.google.com/drive/folders/folder-3"; "PKP Angkatan 3")']
    }
  ]);
  const row = findRekapRowForSubActivityFromLookup_(lookup, {
    sub_activity_id: 'pkp_angkatan_3',
    sub_activity_name: 'PKP Angkatan 3 Tahun 2026',
    formal_archive_name: duplicateFormalName,
    folder_id: 'folder-3',
    sort_order: '7',
    rekap_row_number: '8'
  });
  assert.strictEqual(row, 9, 'hint baris 8 tidak boleh menimpa baris folder-1');
});

test('rekap lookup - marker stabil tetap menemukan baris setelah rename/sort', () => {
  const notes = Array(8).fill('');
  notes[3] = buildRekapIdentityMarker_({
    sub_activity_id: 'stable-sub',
    folder_id: 'stable-folder'
  });
  const lookup = makeRekapLookup([
    { values: ['', '1', '', 'Baris lain', '', '', '', 'Lain'], notes: Array(8).fill('') },
    { values: ['', '5', '', 'Nama lama sebelum rename', '', '', '', 'Manual'], notes: notes }
  ]);
  const row = findRekapRowForSubActivityFromLookup_(lookup, {
    sub_activity_id: 'stable-sub',
    sub_activity_name: 'Nama baru',
    formal_archive_name: 'Nama Formal Baru',
    folder_id: 'stable-folder',
    sort_order: '5',
    rekap_row_number: '8'
  });
  assert.strictEqual(row, 9);
});

test('rekap lookup - nomor berkas duplikat tidak boleh dipakai untuk menebak', () => {
  const lookup = makeRekapLookup([
    { values: ['', '7', '', 'Sub A', '', '', '', 'A'] },
    { values: ['', '7', '', 'Sub B', '', '', '', 'B'] }
  ]);
  const row = findRekapRowForSubActivityFromLookup_(lookup, {
    sub_activity_id: 'missing-sub',
    sub_activity_name: 'Tidak Ditemukan',
    formal_archive_name: 'Tidak Ditemukan',
    sort_order: '7'
  });
  assert.strictEqual(row, null);
});

test('rekap reconcile - baris benar-benar hilang aman dibuat ulang', () => {
  const lookup = makeRekapLookup([
    { values: ['', '3', '', 'Pelatihan Kepemimpinan Nasional Tk.II Angkatan X Tahun 2026', '', '', '', 'PKN Tk.II Angkatan X'] },
    { values: ['', '5', '', 'Pelatihan Kepemimpinan Pengawas Angkatan 2 Tahun 2026', '', '', '', 'PKP Angkatan 2'] }
  ]);
  const possibleRows = findPossibleRekapRowsFromLookup_(lookup, {
    sub_activity_id: 'pkp_angkatan_1',
    sub_activity_name: 'PKP Angkatan 1 Tahun 2026',
    formal_archive_name: 'Pelatihan Kepemimpinan Pengawas Angkatan 1 Tahun 2026',
    folder_id: 'folder-pkp-1',
    sort_order: '4'
  });
  assert.deepStrictEqual(possibleRows, []);
});

test('rekap lookup - angka Romawi tidak boleh cocok sebagai substring', () => {
  const existingNumerals = ['III', 'IV', 'VIII', 'IX', 'XII'];
  const lookup = makeRekapLookup(existingNumerals.map(function (numeral) {
    return {
      values: [
        '',
        '',
        '',
        'Pelatihan Dasar CPNS Angkatan ' + numeral,
        '',
        '',
        '',
        'Latsar CPNS Angkatan ' + numeral
      ]
    };
  }));
  ['I', 'II', 'V', 'VI', 'VII', 'X', 'XI'].forEach(function (numeral) {
    const subActivity = {
      sub_activity_id: 'latsar_' + numeral.toLowerCase(),
      sub_activity_name: 'Latsar CPNS Angkatan ' + numeral,
      formal_archive_name: 'Pelatihan Dasar CPNS Angkatan ' + numeral
    };
    assert.strictEqual(
      findRekapRowForSubActivityFromLookup_(lookup, subActivity),
      null,
      'Angkatan ' + numeral + ' tidak boleh memakai baris angka Romawi lain'
    );
    assert.deepStrictEqual(
      findPossibleRekapRowsFromLookup_(lookup, subActivity),
      [],
      'Angkatan ' + numeral + ' harus dianggap benar-benar belum punya baris'
    );
  });

  assert.strictEqual(findRekapRowForSubActivityFromLookup_(lookup, {
    sub_activity_id: 'latsar_iii',
    sub_activity_name: 'Latsar CPNS Angkatan III',
    formal_archive_name: 'Pelatihan Dasar CPNS Angkatan III'
  }), 8, 'nama persis Angkatan III tetap harus ditemukan');
});

test('formulaSep_ - standalone web app memakai locale tanpa Document Properties atau probe sel', () => {
  formulaSep_._m = {};
  const idSpreadsheet = {
    getId: function () { return 'ss-id-locale'; },
    getSpreadsheetLocale: function () { return 'id_ID'; },
    getSheets: function () { throw new Error('probe sel tidak boleh dijalankan'); }
  };
  const usSpreadsheet = {
    getId: function () { return 'ss-us-locale'; },
    getSpreadsheetLocale: function () { return 'en_US'; },
    getSheets: function () { throw new Error('probe sel tidak boleh dijalankan'); }
  };
  assert.strictEqual(formulaSep_(idSpreadsheet), ';');
  assert.strictEqual(formulaSep_(usSpreadsheet), ',');
  assert.strictEqual(formulaSep_({
    getParent: function () { return idSpreadsheet; }
  }), ';');
});

test('rekap reconcile - nomor berkas tidak boleh dipakai sebagai identitas baris', () => {
  const lookup = makeRekapLookup([
    { values: ['', '7', '', 'Sub A', '', '', '', 'A'] },
    { values: ['', '7', '', 'Sub B', '', '', '', 'B'] }
  ]);
  const possibleRows = findPossibleRekapRowsFromLookup_(lookup, {
    sub_activity_id: 'missing-sub',
    sub_activity_name: 'Tidak Ditemukan',
    formal_archive_name: 'Tidak Ditemukan',
    sort_order: '7'
  });
  assert.deepStrictEqual(possibleRows, []);
  assert.strictEqual(findRekapRowForSubActivityFromLookup_(lookup, {
    sub_activity_id: 'missing-sub',
    sub_activity_name: 'Tidak Ditemukan',
    formal_archive_name: 'Tidak Ditemukan',
    sort_order: '7'
  }), null);
});

test('nomor berkas sub-kegiatan - no_folder config tidak boleh dipakai sebagai nomor rekap', () => {
  const sub = {
    sub_activity_id: 'pkp_angkatan_1',
    no_folder: '1',
    sort_order: '7'
  };
  assert.strictEqual(resolveSubActivityArchiveNumber_(sub), '7');
});

test('urutan sub-kegiatan - mengikuti sort_order secara numerik', () => {
  const rows = [
    { no_folder: '10', sort_order: '1' },
    { no_folder: '2', sort_order: '9' },
    { no_folder: '1', sort_order: '8' }
  ].sort(compareSubActivitiesByArchiveNumber_);
  assert.deepStrictEqual(rows.map(row => row.sort_order), ['1', '8', '9']);
});

test('nomor arsip global - Latsar mulai tepat setelah jumlah Kepemimpinan aktif', () => {
  const activities = [
    { year: 2026, activity_id: 'kepemimpinan', is_active: 'TRUE', sort_order: 1 },
    { year: 2026, activity_id: 'latsar_cpns', is_active: 'TRUE', sort_order: 2 }
  ];
  const leadership = Array.from({ length: 7 }, function (_, index) {
    return {
      year: 2026,
      activity_id: 'kepemimpinan',
      sub_activity_id: 'p' + (index + 1),
      sub_activity_name: 'Program ' + (index + 1),
      local_sort_order: index + 1,
      is_active: 'TRUE'
    };
  });
  const latsar = [
    ['xii', 'Latsar CPNS Angkatan XII'],
    ['iii', 'Latsar CPNS Angkatan III'],
    ['ix', 'Latsar CPNS Angkatan IX'],
    ['i', 'Latsar CPNS Angkatan I'],
    ['vi', 'Latsar CPNS Angkatan VI'],
    ['kutim2', 'Latsar CPNS Kutim Angkatan 2'],
    ['xi', 'Latsar CPNS Angkatan XI'],
    ['iv', 'Latsar CPNS Angkatan IV'],
    ['ii', 'Latsar CPNS Angkatan II'],
    ['viii', 'Latsar CPNS Angkatan VIII'],
    ['v', 'Latsar CPNS Angkatan V'],
    ['x', 'Latsar CPNS Angkatan X'],
    ['vii', 'Latsar CPNS Angkatan VII'],
    ['kutim1', 'Latsar CPNS Kutim Angkatan 1']
  ].map(function (row) {
    return {
      year: 2026,
      activity_id: 'latsar_cpns',
      sub_activity_id: row[0],
      sub_activity_name: row[1],
      is_active: 'TRUE'
    };
  });

  const plan = buildGlobalArchiveNumberPlan_(activities, leadership.concat(latsar));
  const byId = {};
  plan.activeAssignments.forEach(function (assignment) {
    byId[assignment.subActivityId] = assignment.globalNumber;
  });
  assert.strictEqual(byId.i, 8);
  assert.strictEqual(byId.ii, 9);
  assert.strictEqual(byId.iii, 10);
  assert.strictEqual(byId.xii, 19);
  assert.strictEqual(byId.kutim1, 20);
  assert.strictEqual(byId.kutim2, 21);
});

test('nomor arsip global - hapus merapatkan dan restore menggeser kembali', () => {
  const activities = [
    { year: 2026, activity_id: 'kepemimpinan', is_active: 'TRUE', sort_order: 1 },
    { year: 2026, activity_id: 'latsar_cpns', is_active: 'TRUE', sort_order: 2 }
  ];
  const rows = [
    { year: 2026, activity_id: 'kepemimpinan', sub_activity_id: 'p1', local_sort_order: 1, is_active: 'TRUE' },
    { year: 2026, activity_id: 'kepemimpinan', sub_activity_id: 'p2', local_sort_order: 2, is_active: 'FALSE' },
    { year: 2026, activity_id: 'kepemimpinan', sub_activity_id: 'p3', local_sort_order: 3, is_active: 'TRUE' },
    { year: 2026, activity_id: 'latsar_cpns', sub_activity_id: 'l1', sub_activity_name: 'Latsar CPNS Angkatan I', is_active: 'TRUE' }
  ];
  let plan = buildGlobalArchiveNumberPlan_(activities, rows);
  assert.deepStrictEqual(
    plan.activeAssignments.map(function (assignment) { return assignment.subActivityId + ':' + assignment.globalNumber; }),
    ['p1:1', 'p3:2', 'l1:3']
  );

  rows[1].is_active = 'TRUE';
  plan = buildGlobalArchiveNumberPlan_(activities, rows);
  assert.deepStrictEqual(
    plan.activeAssignments.map(function (assignment) { return assignment.subActivityId + ':' + assignment.globalNumber; }),
    ['p1:1', 'p2:2', 'p3:3', 'l1:4']
  );
});

test('rekap reconcile - nomor berkas berbeda dari No Folder harus diperbaiki', () => {
  const notes = Array(8).fill('');
  notes[3] = buildRekapIdentityMarker_({
    sub_activity_id: 'pkp_angkatan_1',
    folder_id: 'folder-pkp-1'
  });
  const formulas = Array(8).fill('');
  formulas[7] = '=HYPERLINK("https://drive.google.com/drive/folders/folder-pkp-1"; "PKP Angkatan 1")';
  const lookup = makeRekapLookup([{
    values: ['', '7', '', 'Pelatihan Kepemimpinan Pengawas Angkatan 1', '', '', '', 'PKP Angkatan 1', '4'],
    formulas: formulas,
    notes: notes
  }]);
  assert.strictEqual(rekapIdentityNeedsRepair_(lookup, 8, {}, {
    sub_activity_id: 'pkp_angkatan_1',
    folder_id: 'folder-pkp-1',
    formal_archive_name: 'Pelatihan Kepemimpinan Pengawas Angkatan 1',
    no_folder: '1',
    sort_order: '7'
  }), true);
});

test('rekap append - memilih nomor kosong terkecil dari kolom No Folder', () => {
  const lookup = makeRekapLookup([
    { values: ['', '1', '', 'A', '', '', '', 'A', '1'] },
    { values: ['', '2', '', 'B', '', '', '', 'B', '2'] },
    { values: ['', '3', '', 'C', '', '', '', 'C', '3'] },
    { values: ['', '5', '', 'D', '', '', '', 'D', '5'] }
  ]);
  assert.strictEqual(getNextAvailableRekapArchiveNumber_(lookup), '4');
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
