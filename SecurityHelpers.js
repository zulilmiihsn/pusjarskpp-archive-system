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

/**
 * Kumpulkan ID folder "root" workspace yang diizinkan dari Script Properties + config.
 * Dipakai validasi scope agar endpoint yang menerima folderId/fileId mentah dari
 * client tidak bisa dipakai mengakses item DI LUAR ruang kerja (cegah IDOR) — penting
 * terutama saat executeAs=USER_DEPLOYING (script jalan sebagai pemilik, bisa baca
 * semua file pemilik).
 * @param {*} [year]
 * @return {Object<string,boolean>} himpunan ID root (key = folderId)
 */
function _collectWorkspaceRootIds_(year) {
  const roots = {};
  const add = function (id) { const c = cleanId_(id); if (c) roots[c] = true; };
  try {
    const s = ConfigService.getSettings();
    add(s.workspaceRootFolderId);
    add(s.referenceRootFolderId);
    add(s.systemFolderParentId);
    add(s.configFolderId);
    if (!year) year = s.currentYear;
  } catch (e) { /* abaikan; lanjut dari config */ }
  try {
    const cfg = CacheHelper.getConfig(year);
    (cfg.years || []).forEach(function (y) {
      add(y.root_folder_id);
      add(y.reference_root_folder_id);
      add(y.daftar_arsip_folder_id);
      add(y.spreadsheet_year_folder_id);
      add(y.persuratan_year_folder_id);
      add(y.inbox_folder_id);
      add(y.template_folder_id);
    });
  } catch (e) { /* abaikan */ }
  return roots;
}

/**
 * Cek apakah `id` atau salah satu ancestornya (sampai kedalaman maxDepth) ada dalam
 * himpunan root yang diizinkan. Pakai Drive.Files.get(fields:parents) — satu REST call
 * per node, lebih murah dari DriveApp.getParents().
 * @param {string} id
 * @param {Object<string,boolean>} roots
 * @param {number} maxDepth
 * @return {boolean}
 */
function _isWithinRoots_(id, roots, maxDepth) {
  const visited = {};
  let frontier = [id];
  let depth = 0;
  while (frontier.length && depth <= maxDepth) {
    const next = [];
    for (let i = 0; i < frontier.length; i++) {
      const nodeId = frontier[i];
      if (roots[nodeId]) return true;
      if (visited[nodeId]) continue;
      visited[nodeId] = true;
      let parents = [];
      try {
        const meta = Drive.Files.get(nodeId, { fields: 'parents', supportsAllDrives: true });
        parents = meta.parents || [];
      } catch (e) {
        parents = [];
      }
      for (let j = 0; j < parents.length; j++) {
        if (roots[parents[j]]) return true;
        if (!visited[parents[j]]) next.push(parents[j]);
      }
    }
    frontier = next;
    depth++;
  }
  return false;
}

/**
 * Pastikan item Drive (folder/file) berada DI DALAM ruang kerja sebelum dioperasikan.
 * Fail-closed untuk item yang ada root-nya tapi tak ketemu (lempar Akses ditolak).
 * Bila TIDAK ada root yang bisa ditentukan sama sekali (workspace belum dikonfigurasi
 * penuh), scope-check dilewati dengan peringatan agar aplikasi tidak mati total —
 * ancaman terbatas (hanya user kantor terautentikasi).
 * @param {string} itemId
 * @param {*} [year]
 * @param {string[]} [allowedRootIds] override root (mis. khusus Inbox)
 */
function requireWithinWorkspace_(itemId, year, allowedRootIds) {
  const id = cleanId_(itemId);
  if (!id) throw new Error('Akses ditolak. ID item tidak valid.');
  let roots;
  if (allowedRootIds && allowedRootIds.length) {
    roots = {};
    allowedRootIds.forEach(function (r) { const c = cleanId_(r); if (c) roots[c] = true; });
  } else {
    roots = _collectWorkspaceRootIds_(year);
  }
  if (!Object.keys(roots).length) {
    // Fail-CLOSED bila workspace SUDAH terkonfigurasi tapi root tak terkumpul (mis. config
    // sengaja dirusak agar scope-check dilewati). Skip hanya saat benar-benar pra-init.
    let configured = false;
    try { configured = !!ConfigService.getSettings().configSpreadsheetId; } catch (e) { configured = false; }
    if (configured) throw new Error('Akses ditolak. Ruang kerja tidak dapat diverifikasi.');
    console.warn('requireWithinWorkspace_: pra-init, scope-check dilewati untuk ' + id);
    return;
  }
  if (!_isWithinRoots_(id, roots, 12)) {
    throw new Error('Akses ditolak. Item berada di luar ruang kerja.');
  }
}
