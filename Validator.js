'use strict';

const Validator = {
  requireString: function (value, fieldName) {
    if (value === undefined || value === null || typeof value !== 'string' || value.trim() === '') {
      throw new Error((fieldName || 'Input') + ' harus berupa teks dan tidak boleh kosong.');
    }
    if (value.trim().length > MAX_INPUT_LENGTH) {
      throw new Error((fieldName || 'Input') + ' maksimal ' + MAX_INPUT_LENGTH + ' karakter.');
    }
    return value.trim();
  },

  requireLongString: function (value, fieldName) {
    if (value === undefined || value === null || typeof value !== 'string' || value.trim() === '') {
      throw new Error((fieldName || 'Input') + ' harus berupa teks dan tidak boleh kosong.');
    }
    return value.trim();
  },


  requireId: function (value, fieldName) {
    const str = this.requireString(value, fieldName);
    const cleaned = cleanId_(str);
    if (!cleaned) {
      throw new Error((fieldName || 'ID') + ' tidak valid.');
    }
    if (cleaned.length > MAX_ID_LENGTH) {
      throw new Error((fieldName || 'ID') + ' tidak valid (terlalu panjang).');
    }
    return cleaned;
  },

  requireYear: function (value) {
    const year = Number(value);
    if (isNaN(year) || year < 1900 || year > 2100) {
      throw new Error('Tahun tidak valid.');
    }
    return year;
  },

  requireMetadata: function (metadata, fields) {
    if (!metadata || typeof metadata !== 'object') {
      throw new Error('Metadata tidak valid.');
    }
    if (fields && Array.isArray(fields)) {
      fields.forEach(function (field) {
        if (isTrue_(field.is_required)) {
          const val = metadata[field.field_name];
          if (val === undefined || val === null || String(val).trim() === '') {
            throw new Error('Kolom "' + (field.field_label || field.field_name) + '" wajib diisi.');
          }
        }
      });
    }
    return metadata;
  }
};
