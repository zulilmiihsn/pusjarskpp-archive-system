'use strict';

/** @private Template data and category operations. */
const SettingsTemplateImpl_ = {
  getTemplatesData: function (payload) {
    payload = payload || {};
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const cacheKey = 'tpl_data_' + year;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.categories) return parsed;
      }
    } catch (_) {}
    const driveData = DriveService.listTemplatesByCategory(year);
    const categoryMetaRows = ConfigRepository.getTemplateCategories();
    const catMetaById = {};
    categoryMetaRows.forEach(function (row) {
      if (row.category_id) catMetaById[row.category_id] = row;
    });

    const result = {
      categories: (driveData.categories || []).map(function (category) {
        const meta = catMetaById[category.id] || {};
        return Object.assign({}, category, {
          color: normalizeHexColor_(meta.color, DEFAULT_TEMPLATE_CATEGORY_COLOR),
          sort_order: meta.sort_order || ''
        });
      }).sort(function (a, b) {
        return Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
          String(a.name || '').localeCompare(String(b.name || ''), 'id', { sensitivity: 'base' });
      }),
      uncategorized: driveData.uncategorized || []
    };
    try {
      const serialized = JSON.stringify(result);
      if (serialized.length < 95000) {
        CacheService.getScriptCache().put(cacheKey, serialized, 600);
      }
    } catch (_) {}
    return result;
  },

  getAdminAuditLogs: function (payload) {
    payload = payload || {};
    return ConfigRepository.getAdminAuditLogs(Number(payload.limit || 100));
  },

  getTemplateCategories: function () {
    return ConfigRepository.getTemplateCategories();
  },

  saveTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.name, 'Nama kategori');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    const folder = DriveService.createTemplateCategory(year, payload.name);
    const meta = ConfigRepository.saveTemplateCategory({
      categoryId: folder.id,
      name: folder.name,
      color: payload.color
    });
    return {
      id: folder.id,
      name: folder.name,
      color: meta.color
    };
  },

  renameTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.categoryId, 'Category ID');
    Validator.requireString(payload.name, 'Nama baru');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    requireWithinTemplateWorkspace_(payload.categoryId, year);
    const folder = DriveService.renameTemplateCategoryFolder(payload.categoryId, payload.name, year);
    const meta = ConfigRepository.saveTemplateCategory({
      categoryId: folder.id,
      name: folder.name,
      color: payload.color
    });
    return {
      id: folder.id,
      name: folder.name,
      color: meta.color
    };
  },

  deleteTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.categoryId, 'Category ID');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    requireWithinTemplateWorkspace_(payload.categoryId, year);
    const result = DriveService.deleteTemplateCategoryFolder(payload.categoryId, year);
    ConfigRepository.deleteTemplateCategory(payload.categoryId);
    return result;
  },

  setTemplateCategory: function (payload) {
    payload = payload || {};
    Validator.requireString(payload.fileId, 'File template');
    const year = Validator.requireYear(payload.year || ConfigService.getSettings().currentYear || DEFAULT_YEAR);
    requireWithinTemplateWorkspace_(payload.fileId, year);
    if (payload.categoryId) requireWithinTemplateWorkspace_(payload.categoryId, year);
    return ConfigRepository.setTemplateCategory(payload.fileId, payload.categoryId || '');
  }
};
