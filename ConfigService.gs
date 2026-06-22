'use strict';

function AccountService_exec_(method) {
  const ssId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID);
  if (!ssId) throw new Error('Spreadsheet konfigurasi belum dikonfigurasi.');
  return method(SpreadsheetApp.openById(ssId));
}

const ConfigService = {
  findActivity: function (config, activityId) {
    return config.activities.find(function (a) { return a.activity_id === activityId; });
  },

  findSubActivity: function (config, activityId, subActivityId) {
    return config.subActivities.find(function (sub) {
      return sub.activity_id === activityId && sub.sub_activity_id === subActivityId;
    });
  },

  getYearConfig: function (config, year) {
    const row = config.years.find(function (item) { return Number(item.year) === Number(year); });
    if (!row) throw new Error('Konfigurasi tahun ' + year + ' tidak ditemukan.');
    return row;
  },

  getSettings: function () {
    const props = PropertiesService.getScriptProperties();
    return {
      configSpreadsheetId: props.getProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID) || '',
      currentYear: Number(props.getProperty(PROP_KEYS.CURRENT_YEAR) || DEFAULT_YEAR),
      workspaceRootFolderId: props.getProperty(PROP_KEYS.WORKSPACE_ROOT_FOLDER_ID) || '',
      referenceRootFolderId: props.getProperty(PROP_KEYS.REFERENCE_ROOT_FOLDER_ID) || '',
      systemFolderParentId: props.getProperty(PROP_KEYS.SYSTEM_FOLDER_PARENT_ID) || '',
      systemFolderName: props.getProperty(PROP_KEYS.SYSTEM_FOLDER_NAME) || '00. Sistem Portal',
      configFolderId: props.getProperty(PROP_KEYS.CONFIG_FOLDER_ID) || ''
    };
  },

  saveSettings: function (payload) {
    const props = PropertiesService.getScriptProperties();
    if (payload.configSpreadsheetId !== undefined) {
      props.setProperty(PROP_KEYS.CONFIG_SPREADSHEET_ID, cleanId_(payload.configSpreadsheetId));
    }
    if (payload.currentYear !== undefined) {
      props.setProperty(PROP_KEYS.CURRENT_YEAR, String(payload.currentYear));
    }
    if (payload.workspaceRootFolderId !== undefined) {
      props.setProperty(PROP_KEYS.WORKSPACE_ROOT_FOLDER_ID, cleanId_(payload.workspaceRootFolderId));
    }
    if (payload.referenceRootFolderId !== undefined) {
      props.setProperty(PROP_KEYS.REFERENCE_ROOT_FOLDER_ID, cleanId_(payload.referenceRootFolderId));
    }
    if (payload.systemFolderParentId !== undefined) {
      props.setProperty(PROP_KEYS.SYSTEM_FOLDER_PARENT_ID, cleanId_(payload.systemFolderParentId));
    }
    if (payload.systemFolderName !== undefined) {
      props.setProperty(PROP_KEYS.SYSTEM_FOLDER_NAME, String(payload.systemFolderName).trim());
    }
    if (payload.configFolderId !== undefined) {
      props.setProperty(PROP_KEYS.CONFIG_FOLDER_ID, cleanId_(payload.configFolderId));
    }
    return this.getSettings();
  },

  listAccounts: function () {
    return AccountService_exec_(function (ss) {
      return readSheetObjects_(ss, CONFIG_SHEETS.ACCOUNTS)
        .filter(function (row) { return isTrue_(row.is_active); })
        .sort(function (a, b) { return String(a.account_id).localeCompare(String(b.account_id)); });
    });
  },

  saveAccount: function (payload) {
    return AccountService_exec_(function (ss) {
      const sheet = getOrCreateSheet_(ss, CONFIG_SHEETS.ACCOUNTS);
      ensureHeaders_(sheet, ACCOUNT_HEADERS);
      const existing = readSheetObjects_(ss, CONFIG_SHEETS.ACCOUNTS);
      const match = existing.find(function (row) { return row.account_id === payload.accountId; });
      if (match) {
        const rowIndex = findConfigRow_(sheet, { account_id: payload.accountId });
        if (rowIndex) {
          const updates = {
            username: payload.username,
            role: payload.role,
            display_name: payload.displayName,
            updated_at: new Date().toISOString()
          };
          if (payload.passwordHash) {
            updates.password_hash = hashPasswordV2_(payload.passwordHash, payload.username);
          }
          updateConfigRow_(sheet, rowIndex, updates);
          return objectFromHeaders_(getHeaders_(sheet), sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
        }
      }
      const accountId = payload.accountId || Utilities.getUuid();
      sheet.appendRow([accountId, payload.username || '', payload.passwordHash ? hashPasswordV2_(payload.passwordHash, payload.username) : '', payload.role || 'user', payload.displayName || '', 'TRUE', new Date().toISOString(), new Date().toISOString()]);
      return { account_id: accountId };
    });
  },

  deleteAccount: function (accountId) {
    return AccountService_exec_(function (ss) {
      const sheet = ss.getSheetByName(CONFIG_SHEETS.ACCOUNTS);
      if (!sheet) return 0;
      const rowIndex = findConfigRow_(sheet, { account_id: accountId });
      if (!rowIndex) return 0;
      updateConfigRow_(sheet, rowIndex, { is_active: 'FALSE', updated_at: new Date().toISOString() });
      return 1;
    });
  }
};
