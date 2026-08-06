'use strict';

function runSecurityHardeningRegressionTests_() {
  const passed = [];
  function check(name, condition) {
    if (!condition) throw new Error('TEST FAILED: ' + name);
    passed.push(name);
  }
  function throws(name, fn) {
    let didThrow = false;
    try { fn(); } catch (_) { didThrow = true; }
    check(name, didThrow);
  }

  const guest = redactBootstrapForGuest_({
    configured: true,
    settings: { currentYear: 2026, configSpreadsheetId: 'secret-config-id' },
    years: [{ year: 2026, root_folder_id: 'secret-root-id' }],
    activities: [{
      activity_id: 'a1',
      activity_name: 'Kegiatan',
      target_folder_id: 'secret-folder-id',
      spreadsheet_file_id: 'secret-sheet-id',
      subActivities: [{ sub_activity_id: 's1', sub_activity_name: 'Sub', folder_id: 'secret-sub-folder' }]
    }],
    history: [{ archive_id: 'secret-history' }],
    documentTypes: [],
    progress: { total: 1, completed: 1, failed: 0, byActivity: {} }
  });
  check('guest settings redacted', Object.keys(guest.settings).join(',') === 'currentYear');
  check('guest years redacted', Object.keys(guest.years[0]).join(',') === 'year');
  check('guest history empty', guest.history.length === 0 && guest.historyMeta.total === 0);
  check('guest activity Drive IDs removed', !guest.activities[0].target_folder_id && !guest.activities[0].spreadsheet_file_id);
  check('guest subactivity Drive ID removed', guest.activities[0].subActivities[0].folder_id === '');

  const validUploadUrl = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&upload_id=test';
  check('Google resumable URL accepted', assertResumableSession_(validUploadUrl) === validUploadUrl);
  throws('foreign resumable host rejected', function () { assertResumableSession_('https://example.com/upload'); });
  throws('lookalike resumable path rejected', function () { assertResumableSession_('https://www.googleapis.com/upload/drive/v3/files.evil'); });

  const generated = generatePassword_();
  check('generated password length', generated.length >= 16);
  check('server password minimum', ACCOUNT_PASSWORD_MIN_LENGTH >= 12);

  return { ok: true, passed: passed.length, tests: passed };
}
