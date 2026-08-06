'use strict';

const SettingsController = {
  getBootstrap: function () {
    return SettingsBootstrapImpl_.getBootstrap.apply(this, arguments);
  },

  saveSettings: function (payload) {
    return SettingsBootstrapImpl_.saveSettings.apply(this, arguments);
  },

  resetWorkspace: function () {
    return SettingsWorkspaceImpl_.resetWorkspace.apply(this, arguments);
  },

  initializeWorkspace: function (payload) {
    return SettingsWorkspaceImpl_.initializeWorkspace.apply(this, arguments);
  },

  ensureArchiveMaintenanceTrigger: function () {
    return SettingsWorkspaceImpl_.ensureArchiveMaintenanceTrigger.apply(this, arguments);
  },

  ensureDocumentTypesSyncTrigger: function () {
    return SettingsWorkspaceImpl_.ensureDocumentTypesSyncTrigger.apply(this, arguments);
  },

  syncDocumentTypeColumns: function (year) {
    return SettingsWorkspaceImpl_.syncDocumentTypeColumns.apply(this, arguments);
  },

  buildProgress: function (history, activities) {
    return SettingsMappingImpl_.buildProgress.apply(this, arguments);
  },

  updateSubActivityMapping: function (payload) {
    return SettingsMappingImpl_.updateSubActivityMapping.apply(this, arguments);
  },

  renameDriveItem: function (payload) {
    return SettingsMappingImpl_.renameDriveItem.apply(this, arguments);
  },

  listDriveFolders: function (payload) {
    return SettingsMappingImpl_.listDriveFolders.apply(this, arguments);
  },

  getTemplatesData: function (payload) {
    return SettingsTemplateImpl_.getTemplatesData.apply(this, arguments);
  },

  getAdminAuditLogs: function (payload) {
    return SettingsTemplateImpl_.getAdminAuditLogs.apply(this, arguments);
  },

  getTemplateCategories: function () {
    return SettingsTemplateImpl_.getTemplateCategories.apply(this, arguments);
  },

  saveTemplateCategory: function (payload) {
    return SettingsTemplateImpl_.saveTemplateCategory.apply(this, arguments);
  },

  renameTemplateCategory: function (payload) {
    return SettingsTemplateImpl_.renameTemplateCategory.apply(this, arguments);
  },

  deleteTemplateCategory: function (payload) {
    return SettingsTemplateImpl_.deleteTemplateCategory.apply(this, arguments);
  },

  deleteYear: function (payload) {
    return SettingsYearImpl_.deleteYear.apply(this, arguments);
  },

  setTemplateCategory: function (payload) {
    return SettingsTemplateImpl_.setTemplateCategory.apply(this, arguments);
  }
};
