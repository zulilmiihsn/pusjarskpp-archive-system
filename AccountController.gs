'use strict';

const AccountController = {
  getAdminAuditLogs: function (payload) { requireAdmin_(payload); return SettingsController.getAdminAuditLogs(payload); },
  listAccounts: function (payload) { 
    requireAdmin_(payload); 
    const all = ConfigService.listAccounts().map(stripAccountSecrets_);
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
  },
  saveAccount: function (payload) {
    payload = payload || {};
    const adminUser = requireAdmin_(payload);
    Validator.requireString(payload.username, 'Username');
    const currentUser = adminUser;
    const isUpdate = !!payload.accountId;
    const result = ConfigService.saveAccount(payload);
    bumpVersion();
    ConfigRepository.appendAdminAudit({
      created_at: new Date().toISOString(),
      actor: currentUser.displayName || currentUser.username || '',
      action: isUpdate ? 'ACCOUNT_UPDATED' : 'ACCOUNT_CREATED',
      status: 'SUCCESS',
      message: (isUpdate ? 'Update akun: ' : 'Buat akun: ') + (payload.displayName || payload.username)
    });
    return result;
  },
  deleteAccount: function (payload) {
    payload = payload || {};
    const adminUser = requireAdmin_(payload);
    Validator.requireString(payload.accountId, 'Account ID');
    const currentUser = adminUser;
    const allAccounts = ConfigService.listAccounts();
    const target = allAccounts.find(function (a) { return a.account_id === payload.accountId; });
    const result = ConfigService.deleteAccount(payload.accountId) > 0;
    bumpVersion();
    if (result) {
      ConfigRepository.appendAdminAudit({
        created_at: new Date().toISOString(),
        actor: currentUser.displayName || currentUser.username || '',
        action: 'ACCOUNT_DELETED',
        status: 'SUCCESS',
        message: 'Nonaktifkan akun: ' + (target ? (target.displayName || target.username) : payload.accountId)
      });
    }
    return result;
  },
  login: function (payload) { return AuthService.login(payload); },
  logout: function (payload) { const r = AuthService.logout(payload); bumpVersion(); return r; },
  getCurrentUser: function (payload) { return AuthService.getCurrentUser(payload); },
  saveDefaultAdmin: function (payload) { requireAdmin_(payload); const r = AuthService.saveDefaultAdmin(); bumpVersion(); return r; },
  getUserEmail: function () { return AuthService.getUserEmail(); },
  getHistory: function (payload) { 
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const all = CacheHelper.getConfig(year).history || [];
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
};
