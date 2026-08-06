'use strict';

function inferSubActivityMappingStatus_(sub) {
  if (!sub) return 'PERLU_REVIEW';
  const folderName = String(sub.sub_activity_name || '').trim();
  const formalName = String(sub.formal_archive_name || '').trim();
  const sheetName = String(sub.target_sheet_name || '').trim();
  if (!formalName || !sheetName) return 'PERLU_REVIEW';
  if (normalizeSettingsMappingText_(formalName) === normalizeSettingsMappingText_(folderName)) return 'PERLU_REVIEW';
  if (normalizeSettingsMappingText_(sheetName) === normalizeSettingsMappingText_(folderName)) return 'PERLU_REVIEW';
  return 'AUTO_MATCHED';
}

function normalizeSettingsMappingText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/^\d+\s*[\.\-\)]\s*/, '')
    .replace(/\btahun\s+20\d{2}\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

