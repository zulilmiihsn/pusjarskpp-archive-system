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

/**
 * Nomor urut sub-kegiatan disimpan di sort_order.
 * no_folder pada config adalah nomor folder tingkat kegiatan, bukan kolom
 * "No Folder" per baris pada sheet Rekap.
 */
function resolveSubActivityArchiveNumber_(subActivity, activity) {
  const sortOrder = String(subActivity && subActivity.sort_order || '').trim();
  if (sortOrder) return sortOrder;
  return String(activity && activity.folder_no || '').trim();
}

function compareSubActivitiesByArchiveNumber_(a, b) {
  const archiveCompare = resolveSubActivityArchiveNumber_(a).localeCompare(
    resolveSubActivityArchiveNumber_(b),
    'id',
    { numeric: true, sensitivity: 'base' }
  );
  return archiveCompare || Number(a && a.sort_order || 0) - Number(b && b.sort_order || 0);
}

function romanNumeralToInt_(value) {
  const text = String(value || '').toUpperCase().trim();
  if (!text || !/^[IVXLCDM]+$/.test(text)) return null;
  const values = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0;
  let previous = 0;
  for (let index = text.length - 1; index >= 0; index--) {
    const current = values[text.charAt(index)] || 0;
    if (current < previous) total -= current;
    else {
      total += current;
      previous = current;
    }
  }
  return total > 0 ? total : null;
}

function extractSubActivityOrdinal_(value) {
  const text = String(value || '').trim();
  if (!text) return null;
  const match = text.match(/(?:angkatan|(?:^|\s)ak)\s*[-.]?\s*(\d+|[IVXLCDM]+)\b/i) ||
    text.match(/^\s*(\d+|[IVXLCDM]+)\s*[.)-]?\s+/i);
  if (!match) return null;
  if (/^\d+$/.test(match[1])) return Number(match[1]);
  return romanNumeralToInt_(match[1]);
}

function inferSubActivityLocalOrder_(subActivity, activity) {
  const explicit = Number(subActivity && subActivity.local_sort_order);
  if (explicit > 0) return explicit;

  const activityId = String(
    (activity && activity.activity_id) ||
    (subActivity && subActivity.activity_id) ||
    ''
  ).toLowerCase();
  const label = [
    subActivity && subActivity.sub_activity_name,
    subActivity && subActivity.target_sheet_name,
    subActivity && subActivity.formal_archive_name,
    subActivity && subActivity.parent_folder_name
  ].filter(Boolean).join(' ');
  const ordinal = extractSubActivityOrdinal_(label);

  if (activityId === 'latsar_cpns') {
    if (ordinal) {
      if (/\b(kutim|kutai\s+timur)\b/i.test(label)) return 1000 + ordinal;
      if (/\b(bengkayang|berau|kabupaten|kerja\s*sama|kerjasama)\b/i.test(label)) return 2000 + ordinal;
      return ordinal;
    }
  }

  const observed = Number(subActivity && subActivity._observed_archive_number);
  if (observed > 0) return observed;
  const existingGlobal = Number(subActivity && subActivity.sort_order);
  if (existingGlobal > 0) return existingGlobal;
  return Number.MAX_SAFE_INTEGER;
}

function compareSubActivitiesByLocalOrder_(a, b, activity) {
  const rankA = inferSubActivityLocalOrder_(a, activity);
  const rankB = inferSubActivityLocalOrder_(b, activity);
  if (rankA !== rankB) return rankA - rankB;
  const nameA = String(a && (a.sub_activity_name || a.formal_archive_name) || '');
  const nameB = String(b && (b.sub_activity_name || b.formal_archive_name) || '');
  return nameA.localeCompare(nameB, 'id', { numeric: true, sensitivity: 'base' }) ||
    String(a && a.sub_activity_id || '').localeCompare(String(b && b.sub_activity_id || ''));
}

/**
 * Membuat satu urutan global 1..N untuk seluruh sub-kegiatan aktif pada satu tahun.
 * Sub-kegiatan nonaktif tetap ikut urutan lokal agar restore kembali ke posisi lama,
 * tetapi tidak memperoleh nomor global.
 */
function buildGlobalArchiveNumberPlan_(activities, subActivities) {
  const activeActivities = (activities || []).filter(function (activity) {
    return activity && (activity.is_active === undefined || isTrue_(activity.is_active));
  }).slice().sort(function (a, b) {
    return Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
      String(a.activity_id || '').localeCompare(String(b.activity_id || ''));
  });
  const activityById = {};
  activeActivities.forEach(function (activity) {
    activityById[String(activity.activity_id || '')] = activity;
  });

  let globalNumber = 1;
  const assignments = [];
  activeActivities.forEach(function (activity) {
    const rows = (subActivities || []).filter(function (subActivity) {
      return String(subActivity && subActivity.activity_id || '') === String(activity.activity_id || '');
    }).slice().sort(function (a, b) {
      return compareSubActivitiesByLocalOrder_(a, b, activity);
    });

    rows.forEach(function (subActivity, index) {
      const active = isTrue_(subActivity.is_active);
      assignments.push({
        year: subActivity.year || activity.year,
        activityId: activity.activity_id,
        subActivityId: subActivity.sub_activity_id,
        localSortOrder: index + 1,
        globalNumber: active ? globalNumber++ : undefined,
        isActive: active,
        subActivity: subActivity,
        activity: activity
      });
    });
  });

  return {
    assignments: assignments,
    activeAssignments: assignments.filter(function (assignment) { return assignment.isActive; }),
    totalActive: globalNumber - 1,
    activityById: activityById
  };
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
