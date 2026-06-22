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

// Test cases
test('forceResetAdmin - initializes workspace auth bypass correctly', () => {
  // Set fake script ID to pretend workspace is initialized
  Properties['CONFIG_SPREADSHEET_ID'] = 'mock-config-ss-id';
  
  // Create a mock sheet for config
  const fakeSheet = {
    getDataRange: () => ({ getDisplayValues: () => [['account_id']] }),
    getRange: () => ({ clearContent: () => {}, setValues: () => {}, setFontWeight: () => {} }),
    appendRow: (row) => {
      // row should contain admin credentials
      assert.strictEqual(row[1], 'admin'); // username
      assert.strictEqual(row[3], 'admin'); // role
    }
  };
  
  mockFiles['mock-config-ss-id'] = {
    getId: () => 'mock-config-ss-id',
    getOwner: () => ({ getEmail: () => 'admin@example.com' }), // must match Session mock
    getSheetByName: (name) => {
      if (name === 'config_accounts') return fakeSheet;
      return null;
    },
    insertSheet: () => fakeSheet
  };

  const result = WorkspaceController.forceResetAdmin();
  assert.ok(result.password, 'Should return a generated password');
});

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

console.log(`\nIntegration Tests Finished: ${passed} passed, ${failed} failed.`);
if (failed > 0) {
  process.exit(1);
}
