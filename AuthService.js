'use strict';

const LOGIN_RATE_LIMIT_WINDOW_MS = 60000;
const LOGIN_RATE_LIMIT_MAX_ATTEMPTS = 10;

// Konteks request hanya diisi dari sesi login internal portal. Email Google tidak
// pernah dipakai untuk autentikasi/otorisasi; pembacaan Session hanya untuk log diagnosis.
let _requestPortalUser_ = null;

function resetRequestPortalUser_() { _requestPortalUser_ = null; }
function setRequestPortalUser_(user) {
  _requestPortalUser_ = user || null;
  if (typeof updateApiRequestActor_ === 'function') updateApiRequestActor_(_requestPortalUser_);
}
function getRequestPortalUser_() { return _requestPortalUser_; }

function detectedGoogleEmailForDiagnostics_() {
  try { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); }
  catch (_) { return ''; }
}

function loginRateLimitKey_(username) {
  return 'login_attempts_' + String(username || '').trim().toLowerCase();
}

function checkLoginRateLimit_(username) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('Layanan login sedang sibuk. Coba lagi beberapa saat.');
  try {
    const props = PropertiesService.getScriptProperties();
    const key = loginRateLimitKey_(username);
    const now = Date.now();
    let entries = [];
    const raw = props.getProperty(key);
    if (raw) {
      try { entries = JSON.parse(raw); } catch (_) { entries = []; }
    }
    entries = entries.filter(function (t) { return now - Number(t) < LOGIN_RATE_LIMIT_WINDOW_MS; });
    if (entries.length >= LOGIN_RATE_LIMIT_MAX_ATTEMPTS) {
      throw new Error('Terlalu banyak percobaan login. Coba lagi dalam 1 menit.');
    }
    entries.push(now);
    props.setProperty(key, JSON.stringify(entries));
  } finally {
    lock.releaseLock();
  }
}

function clearLoginRateLimit_(username) {
  try { PropertiesService.getScriptProperties().deleteProperty(loginRateLimitKey_(username)); } catch (_) {}
}

function cleanupExpiredSessions_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const keys = props.getKeys();
    const now = Date.now();
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      if (key.indexOf('sess_') === 0) {
        try {
          const data = props.getProperty(key);
          if (data) {
            const session = JSON.parse(data);
            if (session.expiresAt && now > session.expiresAt) {
              props.deleteProperty(key);
              try { CacheService.getScriptCache().remove(key); } catch (_) {}
            }
          }
        } catch (e) {
          props.deleteProperty(key); // Hapus jika format korup
        }
      }
    }
  } catch (err) {
    console.error('cleanupExpiredSessions_ error:', err.message);
  }
}

function saveSession_(sessionId, session) {
  const key = 'sess_' + sessionId;
  const serialized = JSON.stringify(session);
  PropertiesService.getScriptProperties().setProperty(key, serialized);
  try { CacheService.getScriptCache().put(key, serialized, SESSION_CACHE_TTL_SECONDS); } catch (_) {}
}

function loadSession_(sessionId) {
  if (!sessionId) return null;
  const key = 'sess_' + sessionId;
  let serialized = null;
  try { serialized = CacheService.getScriptCache().get(key); } catch (_) {}
  if (!serialized) {
    serialized = PropertiesService.getScriptProperties().getProperty(key);
    if (serialized) {
      try { CacheService.getScriptCache().put(key, serialized, SESSION_CACHE_TTL_SECONDS); } catch (_) {}
    }
  }
  if (!serialized) return null;
  try { return JSON.parse(serialized); } catch (_) { deleteSession_(sessionId); return null; }
}

function deleteSession_(sessionId) {
  if (!sessionId) return;
  const key = 'sess_' + sessionId;
  PropertiesService.getScriptProperties().deleteProperty(key);
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
}


function deleteSessionsForAccount_(accountId) {
  if (!accountId) return 0;
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  let deleted = 0;
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (key.indexOf('sess_') !== 0) continue;
    try {
      const raw = props.getProperty(key);
      const session = raw ? JSON.parse(raw) : null;
      if (session && session.accountId === accountId) {
        const sessionId = key.slice(5);
        deleteSession_(sessionId);
        deleted++;
      }
    } catch (_) {
      props.deleteProperty(key);
      try { CacheService.getScriptCache().remove(key); } catch (_) {}
    }
  }
  return deleted;
}

/**
 * Session-based authentication. Script Properties menyimpan sesi durable 2 hari;
 * CacheService 6 jam menjadi lapisan baca cepat untuk validasi setiap request.
 */
const AuthService = {
  login: function (payload) {
    const perfStartedAt = Date.now();
    const loginPerf = {
      event: 'LOGIN_PERF',
      portalUsername: '',
      outcome: 'STARTED',
      accountRows: 0,
      hashVersion: 'unknown',
      rateLimitMs: 0,
      openSpreadsheetMs: 0,
      readAccountsMs: 0,
      verifyPasswordMs: 0,
      saveSessionMs: 0,
      auditMs: 0,
      totalMs: 0
    };
    payload = payload || {};
    Validator.requireString(payload.username, 'Username');
    Validator.requireString(payload.password, 'Kata Sandi');
    loginPerf.portalUsername = String(payload.username || '');
    let phaseStartedAt = Date.now();
    checkLoginRateLimit_(payload.username);
    loginPerf.rateLimitMs = Date.now() - phaseStartedAt;

    const ssId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID);
    if (!ssId) throw new Error('Ruang Kerja belum dikonfigurasi.');

    phaseStartedAt = Date.now();
    const ss = openSpreadsheetById_(ssId);
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ACCOUNTS);
    loginPerf.openSpreadsheetMs = Date.now() - phaseStartedAt;
    if (!sheet) throw new Error('Username atau password salah.');

    // Satu batch read. Tidak ada getRange/openById di dalam loop akun.
    phaseStartedAt = Date.now();
    const values = sheet.getDataRange().getValues();
    loginPerf.readAccountsMs = Date.now() - phaseStartedAt;
    loginPerf.accountRows = Math.max(values.length - 1, 0);
    if (values.length < 2) throw new Error('Username atau password salah.');

    const headers = values[0].map(String);
    const idx = {
      username: headers.indexOf('username'),
      password: headers.indexOf('password_hash'),
      role: headers.indexOf('role'),
      displayName: headers.indexOf('display_name'),
      isActive: headers.indexOf('is_active'),
      id: headers.indexOf('account_id')
    };
    if (idx.username === -1 || idx.password === -1) throw new Error('Username atau password salah.');

    let found = null;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[idx.username] || '').toLowerCase() !== String(payload.username).toLowerCase()) continue;
      if (String(row[idx.isActive] || '').toUpperCase() !== 'TRUE') continue;

      const storedHash = String(row[idx.password] || '');
      loginPerf.hashVersion = storedHash.indexOf(HASH_PREFIX_V3) === 0
        ? 'v3'
        : (storedHash.indexOf(HASH_PREFIX_V2) === 0 ? 'v2' : (storedHash.indexOf(HASH_PREFIX_V1) === 0 ? 'v1' : 'legacy'));
      phaseStartedAt = Date.now();
      const passwordValid = verifyPassword_(payload.password, row[idx.username], storedHash);
      loginPerf.verifyPasswordMs += Date.now() - phaseStartedAt;
      if (passwordValid) {
        found = {
          rowIndex: i,
          accountId: row[idx.id] || '',
          username: row[idx.username],
          role: idx.role >= 0 ? (row[idx.role] || 'user') : 'user',
          displayName: idx.displayName >= 0 ? (row[idx.displayName] || row[idx.username]) : row[idx.username]
        };
        break;
      }
    }

    if (!found) {
      phaseStartedAt = Date.now();
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: payload.username || 'unknown',
        action: 'LOGIN_FAILED',
        status: STATUS.FAILED,
        message: 'Login gagal untuk username: ' + (payload.username || 'unknown')
      });
      loginPerf.auditMs = Date.now() - phaseStartedAt;
      loginPerf.outcome = 'INVALID_CREDENTIALS';
      loginPerf.totalMs = Date.now() - perfStartedAt;
      console.info('LOGIN_PERF ' + JSON.stringify(loginPerf));
      throw new Error('Username atau password salah.');
    }

    clearLoginRateLimit_(found.username);

    const sessionId = Utilities.getUuid();
    const now = new Date();
    const session = {
      sessionId: sessionId,
      accountId: found.accountId,
      username: found.username,
      role: found.role,
      displayName: found.displayName,
      loggedInAt: now.toISOString(),
      expiresAt: now.getTime() + SESSION_TTL_MS
    };
    phaseStartedAt = Date.now();
    saveSession_(sessionId, session);
    loginPerf.saveSessionMs = Date.now() - phaseStartedAt;

    phaseStartedAt = Date.now();
    try {
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: found.displayName || found.username,
        action: 'LOGIN_SUCCESS',
        status: 'SUCCESS',
        message: 'Login berhasil: ' + found.username
      });
    } catch (e) { console.error('audit LOGIN_SUCCESS gagal: ' + e.message); }
    loginPerf.auditMs = Date.now() - phaseStartedAt;
    loginPerf.outcome = 'SUCCESS';
    loginPerf.totalMs = Date.now() - perfStartedAt;
    console.info('LOGIN_PERF ' + JSON.stringify(loginPerf));

    return { sessionId: sessionId, accountId: session.accountId, username: session.username, role: session.role, displayName: session.displayName, loggedInAt: session.loggedInAt };
  },

  logout: function (payload) {
    const sessionId = payload && payload._sessionId;
    if (sessionId) {
      let actorName = '';
      try {
        const s = loadSession_(sessionId);
        if (s) actorName = s.displayName || s.username || '';
      } catch (e) { /* ignore */ }
      deleteSession_(sessionId);
      try {
        ConfigRepository.appendAdminAudit({
          created_at: new Date().toISOString(),
          actor: actorName || 'unknown',
          action: 'LOGOUT',
          status: 'SUCCESS',
          message: 'Logout'
        });
      } catch (e) { console.error('audit LOGOUT gagal: ' + e.message); }
    }
    return { loggedOut: true };
  },

  getCurrentUser: function (payload) {
    const sessionId = payload && payload._sessionId;
    if (sessionId) {
      try {
        const session = loadSession_(sessionId);
        if (session) {
          if (session.expiresAt && Date.now() < session.expiresAt) {
            slideSession_(sessionId, session);
            const user = { accountId: session.accountId, username: session.username, role: session.role, displayName: session.displayName, loggedInAt: session.loggedInAt };
            setRequestPortalUser_(user);
            return user;
          }
          deleteSession_(sessionId);
        }
      } catch (_) {
        console.warn('getCurrentUser parse error');
      }
    }
    return { role: 'guest', username: '', displayName: 'Tamu' };
  },

  saveDefaultAdmin: function () {
    const ssId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID);
    if (!ssId) return { created: false, reason: 'no_workspace' };

    const ss = openSpreadsheetById_(ssId);
    let sheet = ss.getSheetByName(CONFIG_SHEETS.ACCOUNTS);
    if (!sheet) sheet = ss.insertSheet(CONFIG_SHEETS.ACCOUNTS);

    const existing = sheet.getDataRange().getDisplayValues();
    if (existing.length > 1) {
      const hasAdmin = existing.slice(1).some(function (row) { return String(row[1] || '').trim() === 'admin'; });
      if (hasAdmin) return { created: false, reason: 'exists' };
      sheet.getRange(2, 1, existing.length - 1, existing[0].length).clearContent();
    }

    if (!sheet.getRange(1, 1).getValue()) {
      sheet.getRange(1, 1, 1, ACCOUNT_HEADERS.length).setValues([ACCOUNT_HEADERS]);
      sheet.getRange(1, 1, 1, ACCOUNT_HEADERS.length).setFontWeight('normal');
    }

    const password = generatePassword_();
    const hash = hashPasswordV3_(password, 'admin');
    sheet.appendRow([Utilities.getUuid(), 'admin', hash, 'admin', 'Administrator', 'TRUE', new Date().toISOString(), new Date().toISOString()]);
    
    // Auto-login the new admin so subsequent sync operations don't fail
    const sessionId = Utilities.getUuid();
    const now = new Date();
    const session = {
      sessionId: sessionId,
      accountId: 'admin',
      username: 'admin',
      role: 'admin',
      displayName: 'Administrator',
      loggedInAt: now.toISOString(),
      expiresAt: now.getTime() + SESSION_TTL_MS
    };
    saveSession_(sessionId, session);

    try {
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: 'Sistem',
        action: 'ADMIN_INITIALIZED',
        status: 'SUCCESS',
        message: 'Membuat akun admin default'
      });
    } catch (e) { console.error('audit ADMIN_INITIALIZED gagal: ' + e.message); }
    return { created: true, defaultPassword: password, sessionId: sessionId };
  },
};

/**
 * Validate session and return user object. Throws if session is invalid or expired.
 * @param {object} payload - Must contain _sessionId.
 * @return {{accountId: string, username: string, role: string, displayName: string}}
 */
function requireAuth_(payload) {
  const sessionId = payload && payload._sessionId;
  if (!sessionId) throw new Error('Sesi login tidak ditemukan. Silakan login terlebih dahulu.');
  let session = null;
  try {
    session = loadSession_(sessionId);
  } catch (_) {
    session = null;
  }
  if (session && session.expiresAt && Date.now() < session.expiresAt) {
    slideSession_(sessionId, session);
    const user = {
      accountId: session.accountId,
      username: session.username,
      role: session.role,
      displayName: session.displayName
    };
    setRequestPortalUser_(user);
    return user;
  }
  throw new Error('Sesi login telah berakhir. Silakan login kembali.');
}

/**
 * Sliding session: perpanjang masa berlaku selama user masih aktif. Hanya menulis
 * ulang ketika sudah melewati separuh TTL agar tidak boros tulis Script Properties.
 */
function slideSession_(sessionId, session) {
  try {
    const remaining = session.expiresAt - Date.now();
    if (remaining < SESSION_TTL_MS / 2) {
      session.expiresAt = Date.now() + SESSION_TTL_MS;
      saveSession_(sessionId, session);
    }
  } catch (_) {}
}

/**
 * Validate session and require admin role. Throws if not admin.
 * @param {object} payload - Must contain _sessionId.
 * @return {{accountId: string, username: string, role: string, displayName: string}}
 */
function requireAdmin_(payload) {
  const user = requireAuth_(payload);
  if (user.role !== 'admin') throw accessDeniedError_('PORTAL_ROLE_ADMIN', 'Akses ditolak. Hanya admin yang dapat melakukan tindakan ini.');
  return user;
}

/**
 * Penghapusan arsip melalui portal boleh dilakukan petugas (user) dan admin.
 * Role guest tetap read-only. Permission Viewer di Drive tidak berpengaruh karena
 * mutasi dijalankan web app sebagai akun deployer.
 * @param {object} payload
 * @return {{accountId: string, username: string, role: string, displayName: string}}
 */
function requireArchiveDeletionRole_(payload) {
  const user = requireAuth_(payload);
  if (user.role !== 'admin' && user.role !== 'user') {
    throw accessDeniedError_(
      'PORTAL_ROLE_ARCHIVE_DELETE',
      'Akses ditolak. Hanya petugas atau admin yang dapat menghapus arsip.'
    );
  }
  return user;
}
