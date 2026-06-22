'use strict';

/**
 * Tulis satu baris audit. Tidak pernah melempar error — audit yang gagal tidak
 * boleh menggagalkan aksi utama.
 * @param {{displayName?:string,username?:string}|null} actor
 * @param {string} action
 * @param {{year?:*,activityId?:string,subActivityId?:string,folderId?:string,status?:string,message?:string}} [opts]
 */
function auditAction_(actor, action, opts) {
  opts = opts || {};
  try {
    ConfigRepository.appendAdminAudit({
      created_at: new Date().toISOString(),
      actor: (actor && (actor.displayName || actor.username)) || 'Sistem',
      action: action,
      year: opts.year || '',
      activity_id: opts.activityId || '',
      sub_activity_id: opts.subActivityId || '',
      folder_id: opts.folderId || '',
      status: opts.status || 'SUCCESS',
      message: opts.message || ''
    });
  } catch (error) {
    console.error('audit gagal (' + action + '): ' + error.message);
  }
}

/**
 * Ambil user aktif dari sesi pada payload (untuk endpoint tanpa requireAuth_).
 * @param {object} payload
 * @return {object|null}
 */
function auditActor_(payload) {
  try {
    const user = AuthService.getCurrentUser(payload);
    return (user && user.role !== 'guest') ? user : null;
  } catch (error) {
    return null;
  }
}

/**
 * Validasi URL yang dikirim client. Hanya izinkan skema http(s) untuk mencegah
 * stored-XSS lewat skema berbahaya (javascript:, data:, vbscript:).
 * @param {string} value
 * @param {string} [fieldName]
 * @return {string} URL yang sudah dipangkas
 */
function requireSafeUrl_(value, fieldName) {
  const str = Validator.requireString(value, fieldName || 'URL');
  if (!/^https?:\/\//i.test(str)) {
    throw new Error((fieldName || 'URL') + ' harus diawali http:// atau https://.');
  }
  return str;
}

/**
 * Validasi nama yang akan dipakai sebagai nama folder Drive DAN nama sheet.
 * Menolak karakter terlarang Drive (\ / : * ? " < > |) plus karakter terlarang
 * nama sheet Google ([ ]). Mencegah operasi gagal di tengah jalan (folder ter-rename
 * tapi rename sheet melempar error → desync).
 * @param {string} name
 * @param {string} [fieldName]
 * @return {string} nama yang sudah dipangkas
 */
function sanitizeNameForSheetAndDrive_(name, fieldName) {
  const trimmed = String(name || '').trim();
  if (!trimmed) throw new Error((fieldName || 'Nama') + ' wajib diisi.');
  if (/[\\/:*?"<>|[\]]/.test(trimmed)) {
    throw new Error((fieldName || 'Nama') + ' mengandung karakter terlarang ( \\ / : * ? " < > | [ ] ).');
  }
  return trimmed;
}

/**
 * Strip sensitive fields (password hash) from an account row before returning to client.
 * Internal callers (hasActiveAdminAccount_, deleteAccount) use ConfigService.listAccounts
 * directly and never expose the raw row to the client.
 * @param {object} account
 * @return {object}
 */
function stripAccountSecrets_(account) {
  const safe = Object.assign({}, account);
  delete safe.password_hash;
  return safe;
}

function requireAdminIfWorkspaceSecured_(payload) {
  const settings = ConfigService.getSettings();
  if (!settings.configSpreadsheetId) return null;
  let hasAdmin;
  try {
    hasAdmin = ConfigService.listAccounts().some(function (account) {
      return String(account.role || '').trim().toLowerCase() === 'admin';
    });
  } catch (error) {
    // FAIL CLOSED: kegagalan membaca daftar akun TIDAK boleh diperlakukan sebagai
    // "boleh". Kalau dibiarkan return null, error transient bisa membuka jalan
    // mengubah configSpreadsheetId tanpa otorisasi.
    console.warn('Admin account check failed, denying for safety: ' + error.message);
    throw new Error('Akses ditolak. Verifikasi akun admin gagal, coba lagi.');
  }
  if (!hasAdmin) return null; // benar-benar belum ada admin (bootstrap awal)
  return requireAdmin_(payload);
}

function hasActiveAdminAccount_() {
  try {
    return ConfigService.listAccounts().some(function (account) {
      return String(account.role || '').trim().toLowerCase() === 'admin';
    });
  } catch (error) {
    console.warn('Admin account check failed: ' + error.message);
    return false;
  }
}
