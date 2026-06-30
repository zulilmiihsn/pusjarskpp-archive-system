'use strict';

/**
 * Extract Google Drive file/folder ID from a full URL or raw ID.
 * @param {string} value
 * @return {string}
 */
function cleanId_(value) {
  const text = String(value || '').trim();
  const match = text.match(/[-\w]{25,}/);
  return match ? match[0] : text;
}

/**
 * Parse boolean from spreadsheet text or JS value.
 * @param {*} value
 * @return {boolean}
 */
function isTrue_(value) {
  return String(value).toUpperCase() === 'TRUE' || value === true;
}

function slug_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function pad2_(value) {
  if (value === null || value === undefined) return '00';
  const text = String(value).trim();
  if (!text) return '00';
  return text.length === 1 ? '0' + text : text;
}

function sanitizeFilePart_(value) {
  return String(value || '')
    .replace(/[\\:*?"<>|]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_FILE_PART_LENGTH);
}

function normalizeSheetName_(value) {
  return String(value || 'Sheet')
    .replace(/[:\\/?*\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_SHEET_NAME_LENGTH);
}

/* ── Higher-level helpers built from primitives ── */

function normalizeHexColor_(value, fallback) {
  const candidate = String(value || '').trim().toUpperCase();
  if (/^#[0-9A-F]{6}$/.test(candidate)) return candidate;
  return String(fallback || DEFAULT_TEMPLATE_CATEGORY_COLOR || '#2563EB').trim().toUpperCase();
}

/**
 * Run action with exponential backoff retry for transient Google API errors.
 * Retries only on rate-limit, quota, timeout, or server errors.
 * @param {function(): *} action
 * @param {number} maxRetries
 * @return {*}
 */
function withRetry_(action, maxRetries = RETRY_MAX_ATTEMPTS) {
  let attempt = 0;
  while (attempt < maxRetries) {
    try {
      return action();
    } catch (error) {
      const errStr = String(error).toLowerCase();
      if (errStr.includes('limit') || errStr.includes('quota') || errStr.includes('timeout') || errStr.includes('server')) {
        attempt++;
        if (attempt >= maxRetries) throw error;
        Utilities.sleep((Math.pow(2, attempt) * 1000) + Math.round(Math.random() * 1000));
      } else {
        throw error;
      }
    }
  }
}

/**
 * Run action inside LockService for concurrency safety.
 * @param {function(): *} action
 * @param {number} lockTimeoutMs
 * @return {*}
 */
function withLock_(action, lockTimeoutMs = LOCK_TIMEOUT_MS) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(lockTimeoutMs);
    locked = true;
    return action();
  } catch (e) {
    if (locked) throw e;
    throw new Error('Sistem sedang sibuk. Silakan coba beberapa saat lagi.');
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}

function invalidateTemplatesCache_(year) {
  try {
    const y = year || ConfigService.getSettings().currentYear || DEFAULT_YEAR;
    CacheService.getScriptCache().remove('tpl_data_' + y);
  } catch (_) {}
}
