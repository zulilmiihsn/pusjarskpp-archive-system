'use strict';

const ACCOUNT_PASSWORD_MIN_LENGTH = 12;

const AccountController = {
  getAdminAuditLogs: function (payload) { requireAdmin_(payload); return SettingsController.getAdminAuditLogs(payload); },
  listAccounts: function (payload) {
    requireAdmin_(payload);
    return paginate_(ConfigService.listAccounts().map(stripAccountSecrets_), payload);
  },
  saveAccount: function (payload) {
    payload = payload || {};
    const adminUser = requireAdmin_(payload);
    Validator.requireString(payload.username, 'Username');
    const username = String(payload.username || '').trim();
    const role = String(payload.role || 'user').toLowerCase();
    if (role !== 'admin' && role !== 'user') throw new Error('Role akun tidak valid.');
    const allAccounts = ConfigService.listAccounts();
    const existing = payload.accountId
      ? allAccounts.find(function (a) { return a.account_id === payload.accountId; })
      : null;
    if (payload.accountId && !existing) throw new Error('Akun tidak ditemukan.');
    const duplicate = allAccounts.find(function (a) {
      return a.account_id !== payload.accountId && String(a.username || '').toLowerCase() === username.toLowerCase();
    });
    if (duplicate) throw new Error('Username sudah digunakan akun lain.');
    const password = String(payload.passwordHash || '');
    if (!existing && !password) throw new Error('Kata sandi wajib diisi untuk akun baru.');
    if (password && password.length < ACCOUNT_PASSWORD_MIN_LENGTH) throw new Error('Kata sandi minimal ' + ACCOUNT_PASSWORD_MIN_LENGTH + ' karakter.');
    if (existing && String(existing.username || '').toLowerCase() !== username.toLowerCase() && !password) {
      throw new Error('Kata sandi baru wajib diisi saat username diubah.');
    }
    const activeAdmins = allAccounts.filter(function (a) { return String(a.role || '').toLowerCase() === 'admin'; });
    if (existing && String(existing.role || '').toLowerCase() === 'admin' && role !== 'admin' && activeAdmins.length <= 1) {
      throw new Error('Admin terakhir tidak boleh diturunkan rolenya.');
    }
    payload.username = username;
    payload.role = role;
    const isUpdate = !!existing;
    const result = ConfigService.saveAccount(payload);
    if (isUpdate) deleteSessionsForAccount_(existing.account_id);
    bumpVersion();
    auditAction_(adminUser, isUpdate ? 'ACCOUNT_UPDATED' : 'ACCOUNT_CREATED', {
      message: (isUpdate ? 'Update akun: ' : 'Buat akun: ') + (payload.displayName || username)
    });
    return stripAccountSecrets_(result);
  },
  deleteAccount: function (payload) {
    payload = payload || {};
    const adminUser = requireAdmin_(payload);
    Validator.requireString(payload.accountId, 'Account ID');
    const allAccounts = ConfigService.listAccounts();
    const target = allAccounts.find(function (a) { return a.account_id === payload.accountId; });
    if (!target) return false;
    if (target.account_id === adminUser.accountId) throw new Error('Akun yang sedang digunakan tidak boleh dinonaktifkan.');
    const activeAdmins = allAccounts.filter(function (a) { return String(a.role || '').toLowerCase() === 'admin'; });
    if (String(target.role || '').toLowerCase() === 'admin' && activeAdmins.length <= 1) {
      throw new Error('Admin terakhir tidak boleh dinonaktifkan.');
    }
    const result = ConfigService.deleteAccount(payload.accountId) > 0;
    if (result) {
      deleteSessionsForAccount_(payload.accountId);
      bumpVersion();
      auditAction_(adminUser, 'ACCOUNT_DELETED', {
        message: 'Nonaktifkan akun: ' + (target.displayName || target.username || payload.accountId)
      });
    }
    return result;
  },
  login: function (payload) { return AuthService.login(payload); },
  logout: function (payload) { const r = AuthService.logout(payload); bumpVersion(); return r; },
  getCurrentUser: function (payload) { return AuthService.getCurrentUser(payload); },
  getHistory: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    return paginate_(CacheHelper.getConfig(year).history || [], payload);
  }
};

// Paginasi seragam untuk endpoint daftar (akun, riwayat). page mulai 1, pageSize default 50.
function paginate_(all, payload) {
  const page = payload && payload.page ? Math.max(1, payload.page) : 1;
  const pageSize = payload && payload.pageSize ? Math.max(1, payload.pageSize) : 50;
  const total = all.length;
  const start = (page - 1) * pageSize;
  return {
    items: all.slice(start, start + pageSize),
    total: total,
    page: page,
    totalPages: Math.ceil(total / pageSize)
  };
}
