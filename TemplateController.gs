'use strict';

const TemplateController = {
  getTemplates: function (payload) { return SettingsController.getTemplates(payload); },
  getTemplatesData: function (payload) { return SettingsController.getTemplatesData(payload); },
  uploadTemplate: function (payload) { 
    requireAuth_(payload); 
    const r = DriveService.uploadTemplateFile(payload); 
    invalidateTemplatesCache_(payload.year); 
    bumpVersion(); 
    return r; 
  },
  deleteTemplate: function (payload) {
    payload = payload || {};
    requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    invalidateTemplatesCache_(payload.year);
    bumpVersion(); 
    return DriveService.trashTemplateFile(payload.fileId);
  },
  getTemplateCategories: function () { return SettingsController.getTemplateCategories(); },
  saveTemplateCategory: function (payload) {
    requireAuth_(payload || {});
    const r = SettingsController.saveTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year); 
    bumpVersion(); 
    return r; 
  },
  deleteTemplateCategory: function (payload) {
    requireAdmin_(payload || {});
    const r = SettingsController.deleteTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year); 
    bumpVersion(); 
    return r; 
  },
  renameTemplateCategory: function (payload) {
    requireAuth_(payload || {});
    const r = SettingsController.renameTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year); 
    bumpVersion(); 
    return r; 
  },
  setTemplateCategory: function (payload) {
    requireAuth_(payload || {});
    const r = SettingsController.setTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year); 
    bumpVersion(); 
    return r; 
  }
};
