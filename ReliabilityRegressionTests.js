'use strict';

/**
 * Regression suite non-destruktif. Runner dibuat private agar tidak menjadi RPC
 * publik. Saat pengujian editor, nama fungsi boleh sementara dilepas akhiran `_`.
 */
function runReliabilityRegressionTests_() {
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

  const policy = validateAuthorizationPolicy_();
  check('authorization policy valid', policy.ok);
  check('guest bootstrap allowed', isRoleAllowedByPolicy_('getBootstrap', 'guest', true));
  check('guest mutation denied', !isRoleAllowedByPolicy_('finalizeArchive', 'guest', true));
  check('user archive allowed', isRoleAllowedByPolicy_('finalizeArchive', 'user', true));
  check('user admin endpoint denied', !isRoleAllowedByPolicy_('resetWorkspace', 'user', true));
  check('admin endpoint allowed', isRoleAllowedByPolicy_('resetWorkspace', 'admin', true));
  check('bootstrap open before secured', isRoleAllowedByPolicy_('initializeWorkspace', 'guest', false));
  check('bootstrap closed after secured', !isRoleAllowedByPolicy_('initializeWorkspace', 'guest', true));
  check('unknown endpoint denied', !isRoleAllowedByPolicy_('doesNotExist', 'admin', true));

  const redacted = redactLogMetadata_({
    username: 'tester',
    password: 'secret',
    _sessionId: 'session-secret',
    nested: { authorization: 'Bearer secret', ok: true }
  });
  check('password redacted', redacted.password === '[REDACTED]');
  check('session redacted', redacted._sessionId === '[REDACTED]');
  check('nested authorization redacted', redacted.nested.authorization === '[REDACTED]');
  check('safe metadata preserved', redacted.username === 'tester' && redacted.nested.ok === true);

  check('Drive ID sanitized from client error', sanitizeError_('Gagal 1AbCdEfGhIjKlMnOpQrStUvWxYz123456789') .indexOf('[ID]') >= 0);
  check('URL sanitized from client error', sanitizeError_('Gagal https://example.com/private').indexOf('[URL]') >= 0);
  check('access error code', getErrorCode_(accessDeniedError_('TEST', 'Akses ditolak.')) === 'ACCESS_DENIED');

  const context = beginApiRequestContext_();
  check('request ID generated', /^[A-Za-z0-9_-]{20,100}$/.test(context.requestId));
  updateApiRequestActor_({ username: 'tester', role: 'user' });
  check('request actor attached', getApiRequestContext_().actor === 'tester' && getApiRequestContext_().role === 'user');
  finishApiRequestContext_();

  let resumableToken = '';
  try {
    setRequestPortalUser_({ accountId: 'regression-owner-a', username: 'regression-a', role: 'user' });
    resumableToken = registerResumableSession_(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&upload_id=regression',
      1024
    );
    check('resumable session owner accepted', loadResumableSession_(resumableToken, 1024).accountId === 'regression-owner-a');
    setRequestPortalUser_({ accountId: 'regression-owner-b', username: 'regression-b', role: 'user' });
    throws('resumable session foreign owner rejected', function () { loadResumableSession_(resumableToken, 1024); });
    setRequestPortalUser_({ accountId: 'regression-owner-a', username: 'regression-a', role: 'user' });
    throws('resumable changed size rejected', function () { loadResumableSession_(resumableToken, 2048); });
  } finally {
    if (resumableToken) deleteResumableSession_(resumableToken);
    resetRequestPortalUser_();
    finishApiRequestContext_();
  }

  const lockResult = withLock_(function () { return 'locked-ok'; });
  check('repository lock helper operational', lockResult === 'locked-ok');

  const security = runSecurityHardeningRegressionTests_();
  check('security regression suite passed', security && security.ok);

  const health = MonitoringService.check();
  check('health snapshot returned', health && Array.isArray(health.checks) && health.checks.length > 0);
  console.log('REGRESSION_HEALTH status=' + health.status + ' failedChecks=' +
    health.checks.filter(function (item) { return !item.ok; }).map(function (item) {
      return item.name + ':' + item.detail;
    }).join(','));
  const configCheck = health.checks.filter(function (item) {
    return item.name === 'config_spreadsheet_property';
  })[0];
  if (configCheck && configCheck.ok) {
    check('configured environment has no critical failure', health.status !== 'CRITICAL');
  } else {
    check('uninitialized environment reported explicitly', health.status === 'CRITICAL');
  }

  return {
    ok: true,
    passed: passed.length,
    tests: passed,
    policyEndpoints: policy.endpointCount,
    health: health
  };
}
