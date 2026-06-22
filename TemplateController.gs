'use strict';

const TemplateController = {
  getTemplates: function (payload) { return SettingsController.getTemplates(payload); },
  getTemplatesData: function (payload) { return SettingsController.getTemplatesData(payload); },
  uploadTemplate: function (payload) {
    const actor = requireAuth_(payload);
    const r = DriveService.uploadTemplateFile(payload);
    invalidateTemplatesCache_(payload.year);
    bumpVersion();
    auditAction_(actor, 'TEMPLATE_UPLOADED', { year: payload && payload.year, message: 'Mengunggah template: ' + ((payload && payload.name) || '-') });
    return r;
  },
  deleteTemplate: function (payload) {
    payload = payload || {};
    const actor = requireAuth_(payload);
    Validator.requireId(payload.fileId, 'File ID');
    invalidateTemplatesCache_(payload.year);
    bumpVersion();
    const r = DriveService.trashTemplateFile(payload.fileId);
    auditAction_(actor, 'TEMPLATE_DELETED', { year: payload.year, folderId: payload.fileId, message: 'Menghapus template' });
    return r;
  },
  getTemplateCategories: function () { return SettingsController.getTemplateCategories(); },
  saveTemplateCategory: function (payload) {
    const actor = requireAuth_(payload || {});
    const r = SettingsController.saveTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year);
    bumpVersion();
    auditAction_(actor, 'TEMPLATE_CAT_SAVED', { year: payload && payload.year, message: 'Menyimpan kategori template: ' + ((payload && payload.name) || '-') });
    return r;
  },
  deleteTemplateCategory: function (payload) {
    const actor = requireAdmin_(payload || {});
    const r = SettingsController.deleteTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year);
    bumpVersion();
    auditAction_(actor, 'TEMPLATE_CAT_DELETED', { year: payload && payload.year, message: 'Menghapus kategori template: ' + ((payload && (payload.categoryId || payload.name)) || '-') });
    return r;
  },
  renameTemplateCategory: function (payload) {
    const actor = requireAuth_(payload || {});
    const r = SettingsController.renameTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year);
    bumpVersion();
    auditAction_(actor, 'TEMPLATE_CAT_RENAMED', { year: payload && payload.year, message: 'Mengganti nama kategori template menjadi: ' + ((payload && payload.name) || '-') });
    return r;
  },
  setTemplateCategory: function (payload) {
    const actor = requireAuth_(payload || {});
    const r = SettingsController.setTemplateCategory(payload);
    invalidateTemplatesCache_(payload && payload.year);
    bumpVersion();
    auditAction_(actor, 'TEMPLATE_CAT_SET', { year: payload && payload.year, folderId: payload && payload.fileId, message: 'Menetapkan kategori template untuk file' });
    return r;
  }
};
