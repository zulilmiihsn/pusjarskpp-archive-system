'use strict';

const API_SLOW_REQUEST_THRESHOLD_MS_ = 3000;
const LOG_METADATA_MAX_DEPTH_ = 4;
const LOG_METADATA_MAX_KEYS_ = 60;
let _apiRequestContext_ = null;

function newRequestId_() {
  return Utilities.getUuid();
}

function inferApiEndpoint_() {
  try {
    const stack = String(new Error().stack || '');
    const known = typeof ENDPOINT_ACCESS_POLICY_ === 'object' ? ENDPOINT_ACCESS_POLICY_ : {};
    const lines = stack.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const match = lines[i].match(/\bat\s+([A-Za-z0-9_]+)\b/);
      if (match && known[match[1]]) return match[1];
    }
  } catch (_) {}
  return 'unknown';
}

function beginApiRequestContext_() {
  _apiRequestContext_ = {
    requestId: newRequestId_(),
    endpoint: inferApiEndpoint_(),
    startedAt: Date.now(),
    actor: 'guest',
    role: 'guest'
  };
  return _apiRequestContext_;
}

function getApiRequestContext_() {
  return _apiRequestContext_;
}

function updateApiRequestActor_(user) {
  if (!_apiRequestContext_ || !user) return;
  _apiRequestContext_.actor = String(user.displayName || user.username || 'guest');
  _apiRequestContext_.role = String(user.role || 'guest').toLowerCase();
}

function finishApiRequestContext_() {
  _apiRequestContext_ = null;
}

function redactLogMetadata_(value, depth, state) {
  depth = Number(depth || 0);
  state = state || { keys: 0 };
  if (value === null || value === undefined) return value;
  if (depth > LOG_METADATA_MAX_DEPTH_) return '[MAX_DEPTH]';
  if (typeof value === 'string') return value.length > 1000 ? value.slice(0, 1000) + '…' : value;
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  if (value instanceof Date) return value.toISOString();
  if (value instanceof Error) {
    return {
      name: String(value.name || 'Error'),
      message: sanitizeError_(String(value.message || value)),
      errorCode: String(value.errorCode || '')
    };
  }
  if (Array.isArray(value)) {
    return value.slice(0, 20).map(function (item) {
      return redactLogMetadata_(item, depth + 1, state);
    });
  }
  if (typeof value !== 'object') return String(value);
  const output = {};
  Object.keys(value).slice(0, LOG_METADATA_MAX_KEYS_).forEach(function (key) {
    state.keys++;
    if (state.keys > LOG_METADATA_MAX_KEYS_) return;
    if (/(password|passphrase|secret|token|authorization|cookie|sessionurl|sessionid|_sessionid|hash)/i.test(key)) {
      output[key] = '[REDACTED]';
    } else {
      output[key] = redactLogMetadata_(value[key], depth + 1, state);
    }
  });
  return output;
}

function logSlowApiRequest_(context) {
  if (!context) return;
  const durationMs = Date.now() - Number(context.startedAt || Date.now());
  if (durationMs < API_SLOW_REQUEST_THRESHOLD_MS_) return;
  SystemLogger.warn('API_SLOW_REQUEST', 'Endpoint melampaui ambang waktu.', {
    requestId: context.requestId,
    endpoint: context.endpoint,
    durationMs: durationMs,
    portalUsername: context.actor,
    portalRole: context.role,
    outcome: 'SUCCESS'
  });
}

function logFailedApiRequest_(context, error, errorCode) {
  context = context || {};
  SystemLogger.error('API_REQUEST_FAILED', sanitizeError_(error && error.message ? error.message : error), {
    requestId: context.requestId || newRequestId_(),
    endpoint: context.endpoint || 'unknown',
    durationMs: Date.now() - Number(context.startedAt || Date.now()),
    portalUsername: context.actor || 'guest',
    portalRole: context.role || 'guest',
    outcome: 'FAILED',
    errorCode: errorCode || getErrorCode_(error)
  });
}
