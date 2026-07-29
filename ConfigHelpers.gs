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

// SHA-256 murni JavaScript. Utilities.computeDigest adalah panggilan service GAS;
// memanggilnya 5.000 kali membuat login dapat menggantung puluhan detik. Hasil
// fungsi ini identik (UTF-8 -> SHA-256 hex), tetapi seluruh iterasi berjalan di V8.
function sha256Hex_(message) {
  const bytes = [];
  const input = String(message === null || message === undefined ? '' : message);
  for (let i = 0; i < input.length; i++) {
    let code = input.charCodeAt(i);
    if (code >= 0xD800 && code <= 0xDBFF && i + 1 < input.length) {
      const next = input.charCodeAt(i + 1);
      if (next >= 0xDC00 && next <= 0xDFFF) {
        code = 0x10000 + ((code - 0xD800) << 10) + (next - 0xDC00);
        i++;
      }
    }
    if (code < 0x80) {
      bytes.push(code);
    } else if (code < 0x800) {
      bytes.push(0xC0 | (code >>> 6), 0x80 | (code & 0x3F));
    } else if (code < 0x10000) {
      bytes.push(0xE0 | (code >>> 12), 0x80 | ((code >>> 6) & 0x3F), 0x80 | (code & 0x3F));
    } else {
      bytes.push(0xF0 | (code >>> 18), 0x80 | ((code >>> 12) & 0x3F), 0x80 | ((code >>> 6) & 0x3F), 0x80 | (code & 0x3F));
    }
  }

  const bitLength = bytes.length * 8;
  bytes.push(0x80);
  while (bytes.length % 64 !== 56) bytes.push(0);
  const high = Math.floor(bitLength / 0x100000000);
  const low = bitLength >>> 0;
  bytes.push(
    (high >>> 24) & 0xFF, (high >>> 16) & 0xFF, (high >>> 8) & 0xFF, high & 0xFF,
    (low >>> 24) & 0xFF, (low >>> 16) & 0xFF, (low >>> 8) & 0xFF, low & 0xFF
  );

  const k = [
    0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
    0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
    0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
    0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
    0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
    0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
    0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
    0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2
  ];
  const h = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
  const rotr = function (value, amount) { return (value >>> amount) | (value << (32 - amount)); };
  const w = new Array(64);

  for (let offset = 0; offset < bytes.length; offset += 64) {
    for (let i = 0; i < 16; i++) {
      const p = offset + i * 4;
      w[i] = ((bytes[p] << 24) | (bytes[p + 1] << 16) | (bytes[p + 2] << 8) | bytes[p + 3]) | 0;
    }
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
    }
    let a = h[0], b = h[1], c = h[2], d = h[3], e = h[4], f = h[5], g = h[6], hh = h[7];
    for (let i = 0; i < 64; i++) {
      const s1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const t1 = (hh + s1 + ch + k[i] + w[i]) | 0;
      const s0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (s0 + maj) | 0;
      hh = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    h[0] = (h[0] + a) | 0; h[1] = (h[1] + b) | 0; h[2] = (h[2] + c) | 0; h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0; h[5] = (h[5] + f) | 0; h[6] = (h[6] + g) | 0; h[7] = (h[7] + hh) | 0;
  }
  return h.map(function (value) { return ('00000000' + (value >>> 0).toString(16)).slice(-8); }).join('');
}

function pbkdf2Like_(password, username, salt, iterations) {
  const activeSalt = salt || (username || '').toLowerCase() + 'PortalArsip2024';
  let key = activeSalt + (username || '').toLowerCase() + (password || '');
  for (let i = 0; i < iterations; i++) {
    key = sha256Hex_(key) + activeSalt;
  }
  return key;
}

/**
 * Hash password with v2 scheme (50,000 iterations). Kept for verifying hashes
 * created before the v3 migration — jangan dipakai lagi untuk hash baru.
 * @param {string} password
 * @param {string} username
 * @return {string}
 */
function hashPasswordV2_(password, username) {
  const salt = generateSalt_(16);
  const key = pbkdf2Like_(password, username, salt, HASH_ITERATIONS_V2);
  return HASH_PREFIX_V2 + salt + '$' + key;
}

/**
 * Hash password with v3 scheme (5,000 iterations). All new hashes use this.
 * @param {string} password
 * @param {string} username
 * @return {string}
 */
function hashPasswordV3_(password, username) {
  const salt = generateSalt_(16);
  const key = pbkdf2Like_(password, username, salt, HASH_ITERATIONS_V3);
  return HASH_PREFIX_V3 + salt + '$' + key;
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
  // v3 scheme: 5,000 iterations
  if (storedHash.indexOf(HASH_PREFIX_V3) === 0) {
    const raw = storedHash.slice(HASH_PREFIX_V3.length);
    const parts = raw.split('$');
    if (parts.length === 2) {
      const computed = pbkdf2Like_(password, username, parts[0], HASH_ITERATIONS_V3);
      return timingSafeEqual_(computed, parts[1]);
    }
    return false;
  }
  // v2 scheme: 50,000 iterations (legacy — masih diverifikasi, gak dipakai buat hash baru)
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

function generatePassword_(length) {
  const len = length || 6;
  // Charset tanpa simbol & karakter ambigu (0/O/1/I/l) agar password 6-karakter
  // mudah dibaca & diketik saat login.
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
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
