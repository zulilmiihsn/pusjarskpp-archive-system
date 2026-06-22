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

console.log(`\nIntegration Tests Finished: ${passed} passed, ${failed} failed.`);
if (failed > 0) {
  process.exit(1);
}
