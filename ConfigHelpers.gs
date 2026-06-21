'use strict';

function readSheetObjects_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(normalizeHeader_);
  return values.slice(1)
    .filter(row => row.some(value => value !== ''))
    .map(row => objectFromHeaders_(headers, row));
}

function readRecentSheetObjects_(ss, sheetName, limit) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normalizeHeader_);
  const rowCount = Math.min(Math.max(Number(limit) || 80, 1), lastRow - 1);
  const startRow = Math.max(2, lastRow - rowCount + 1);
  return sheet.getRange(startRow, 1, rowCount, lastCol).getDisplayValues()
    .filter(row => row.some(value => value !== ''))
    .map(row => objectFromHeaders_(headers, row));
}

function getHeaders_(sheet) {
  const width = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(normalizeHeader_);
}

function objectFromHeaders_(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    if (header) obj[header] = row[index];
  });
  return obj;
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('normal').setBackground('#d9ead3');
    return;
  }

  const existing = getHeaders_(sheet);
  const missing = headers.filter(header => existing.indexOf(header) === -1);
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, 1, 1, existing.length + missing.length).setFontWeight('normal').setBackground('#d9ead3');
  }
}

function ensureSubActivityHeaders_(sheet) {
  ensureHeaders_(sheet, SUB_ACTIVITY_HEADERS);
}

function findConfigRow_(sheet, criteria) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return 0;
  const headers = values[0].map(normalizeHeader_);
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = objectFromHeaders_(headers, values[rowIndex]);
    const matches = Object.keys(criteria).every(key => String(row[key]) === String(criteria[key]));
    if (matches) return rowIndex + 1;
  }
  return 0;
}

// Batch update: 1 read + 1 write, bukan satu setValue per field.
// Hanya dipakai pada sheet config/log (tanpa formula), jadi read-modify-write
// seluruh baris aman — tidak ada formula yang ter-overwrite.
function updateConfigRow_(sheet, rowIndex, updates) {
  const headers = getHeaders_(sheet);
  const width = headers.length;
  if (width < 1) return;
  const range = sheet.getRange(rowIndex, 1, 1, width);
  const rowValues = range.getValues()[0];
  let dirty = false;
  Object.keys(updates).forEach(key => {
    if (updates[key] === undefined) return;
    const colIndex = headers.indexOf(key);
    if (colIndex >= 0) {
      rowValues[colIndex] = updates[key];
      dirty = true;
    }
  });
  if (dirty) range.setValues([rowValues]);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function normalizeHeader_(value) {
  return String(value || '').trim();
}

function generateSalt_(length) {
  const len = length || 16;
  // Pakai entropi UUID platform, bukan Math.random (yang bukan CSPRNG).
  let pool = '';
  while (pool.length < len) {
    pool += Utilities.getUuid().replace(/-/g, '');
  }
  return pool.slice(0, len);
}

function pbkdf2Like_(password, username, salt, iterations) {
  const activeSalt = salt || (username || '').toLowerCase() + 'PortalArsip2024';
  let key = activeSalt + (username || '').toLowerCase() + (password || '');
  for (let i = 0; i < iterations; i++) {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key, Utilities.Charset.UTF_8);
    key = digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('') + activeSalt;
  }
  return key;
}

function hashPasswordV1_(password, username) {
  const salt = generateSalt_(16);
  const key = pbkdf2Like_(password, username, salt, HASH_ITERATIONS);
  return HASH_PREFIX_V1 + salt + '$' + key;
}

/**
 * Hash password with v2 scheme (50,000 iterations). All new hashes use this.
 * @param {string} password
 * @param {string} username
 * @return {string}
 */
function hashPasswordV2_(password, username) {
  const salt = generateSalt_(16);
  const key = pbkdf2Like_(password, username, salt, HASH_ITERATIONS_V2);
  return HASH_PREFIX_V2 + salt + '$' + key;
}

function timingSafeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

function verifyPassword_(password, username, storedHash) {
  // v2 scheme: 50,000 iterations
  if (storedHash.indexOf(HASH_PREFIX_V2) === 0) {
    const raw = storedHash.slice(HASH_PREFIX_V2.length);
    const parts = raw.split('$');
    if (parts.length === 2) {
      const computed = pbkdf2Like_(password, username, parts[0], HASH_ITERATIONS_V2);
      return timingSafeEqual_(computed, parts[1]);
    }
    return false;
  }
  // v1 scheme: 800 iterations (legacy)
  const prefix = HASH_PREFIX_V1;
  if (storedHash.indexOf(prefix) === 0) {
    const raw = storedHash.slice(prefix.length);
    const parts = raw.split('$');
    if (parts.length === 2) {
      const computed = pbkdf2Like_(password, username, parts[0], HASH_ITERATIONS);
      return timingSafeEqual_(computed, parts[1]);
    }
    const legacy = pbkdf2Like_(password, username, null, 10000);
    return timingSafeEqual_(prefix + legacy, storedHash);
  }
  const legacy = pbkdf2Like_(password, username, null, 500);
  return timingSafeEqual_(legacy, storedHash);
}

/**
 * Check if a stored hash uses an outdated scheme and should be rehashed.
 * @param {string} storedHash
 * @return {boolean}
 */
function needsRehash_(storedHash) {
  return storedHash.indexOf(HASH_PREFIX_V2) !== 0;
}

function generatePassword_(length) {
  const len = length || 14;
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*';
  // Sumber indeks dari entropi UUID, bukan Math.random.
  let pool = '';
  while (pool.length < len * 2) pool += Utilities.getUuid().replace(/-/g, '');
  let result = '';
  for (let i = 0; i < len; i++) {
    const idx = parseInt(pool.substr(i * 2, 2), 16) % chars.length;
    result += chars.charAt(idx);
  }
  return result;
}
