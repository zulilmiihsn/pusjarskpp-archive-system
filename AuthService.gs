'use strict';

const LOGIN_RATE_LIMIT_WINDOW_MS = 60000;
const LOGIN_RATE_LIMIT_MAX_ATTEMPTS = 10;

function checkLoginRateLimit_(username) {
  const props = PropertiesService.getScriptProperties();
  const key = 'login_attempts_' + String(username || '').toLowerCase();
  const raw = props.getProperty(key);
  const now = Date.now();
  if (raw) {
    try {
      const entries = JSON.parse(raw).filter(function (t) { return now - t < LOGIN_RATE_LIMIT_WINDOW_MS; });
      if (entries.length >= LOGIN_RATE_LIMIT_MAX_ATTEMPTS) {
        throw new Error('Terlalu banyak percobaan login. Coba lagi dalam 1 menit.');
      }
      entries.push(now);
      props.setProperty(key, JSON.stringify(entries));
    } catch (e) {
      if (e.message.indexOf('Terlalu banyak') >= 0) throw e;
      props.setProperty(key, JSON.stringify([now]));
    }
  } else {
    props.setProperty(key, JSON.stringify([now]));
  }
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

/**
 * Session-based authentication using Script Properties.
 * Login creates a UUID session (2-day TTL) stored under `sess_{id}`.
 */
const AuthService = {
  login: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.username, 'Username');
    Validator.requireString(payload.password, 'Kata Sandi');
    checkLoginRateLimit_(payload.username);

    const ssId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID);
    if (!ssId) throw new Error('Ruang Kerja belum dikonfigurasi.');

    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(CONFIG_SHEETS.ACCOUNTS);
    if (!sheet) throw new Error('Username atau password salah.');

    const values = sheet.getDataRange().getDisplayValues();
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

      if (verifyPassword_(payload.password, row[idx.username], String(row[idx.password] || ''))) {
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
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: payload.username || 'unknown',
        action: 'LOGIN_FAILED',
        status: STATUS.FAILED,
        message: 'Login gagal untuk username: ' + (payload.username || 'unknown')
      });
      throw new Error('Username atau password salah.');
    }

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
    PropertiesService.getScriptProperties().setProperty('sess_' + sessionId, JSON.stringify(session));
    
    // Garbage collection ringan setiap kali ada login sukses
    cleanupExpiredSessions_();

    try {
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: found.displayName || found.username,
        action: 'LOGIN_SUCCESS',
        status: 'SUCCESS',
        message: 'Login berhasil: ' + found.username
      });
    } catch (e) { console.error('audit LOGIN_SUCCESS gagal: ' + e.message); }

    return { sessionId: sessionId, accountId: session.accountId, username: session.username, role: session.role, displayName: session.displayName, loggedInAt: session.loggedInAt };
  },

  logout: function (payload) {
    const sessionId = payload && payload._sessionId;
    if (sessionId) {
      let actorName = '';
      try {
        const data = PropertiesService.getScriptProperties().getProperty('sess_' + sessionId);
        if (data) { const s = JSON.parse(data); actorName = s.displayName || s.username || ''; }
      } catch (e) { /* ignore */ }
      PropertiesService.getScriptProperties().deleteProperty('sess_' + sessionId);
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
        const data = PropertiesService.getScriptProperties().getProperty('sess_' + sessionId);
        if (data) {
          const session = JSON.parse(data);
          if (session.expiresAt && Date.now() < session.expiresAt) {
            slideSession_(sessionId, session);
            return { accountId: session.accountId, username: session.username, role: session.role, displayName: session.displayName, loggedInAt: session.loggedInAt };
          }
          PropertiesService.getScriptProperties().deleteProperty('sess_' + sessionId);
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

    const ss = SpreadsheetApp.openById(ssId);
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
    const hash = hashPasswordV2_(password, 'admin');
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
    PropertiesService.getScriptProperties().setProperty('sess_' + sessionId, JSON.stringify(session));

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

  getUserEmail: function () {
    try {
      return Session.getActiveUser().getEmail() || '';
    } catch (error) {
      console.error('Session.getActiveUser().getEmail() failed: ' + error.message);
      return '';
    }
  }
};

/**
 * Validate session and return user object. Throws if session is invalid or expired.
 * @param {object} payload - Must contain _sessionId.
 * @return {{accountId: string, username: string, role: string, displayName: string}}
 */
function requireAuth_(payload) {
  const sessionId = payload && payload._sessionId;
  if (!sessionId) throw new Error('Sesi login tidak ditemukan. Silakan login terlebih dahulu.');
  try {
    const data = PropertiesService.getScriptProperties().getProperty('sess_' + sessionId);
    if (data) {
      const session = JSON.parse(data);
      if (session.expiresAt && Date.now() < session.expiresAt) {
        slideSession_(sessionId, session);
        return {
          accountId: session.accountId,
          username: session.username,
          role: session.role,
          displayName: session.displayName
        };
      }
    }
  } catch (_) {}
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
      PropertiesService.getScriptProperties().setProperty('sess_' + sessionId, JSON.stringify(session));
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
  if (user.role !== 'admin') throw new Error('Akses ditolak. Hanya admin yang dapat melakukan tindakan ini.');
  return user;
}
