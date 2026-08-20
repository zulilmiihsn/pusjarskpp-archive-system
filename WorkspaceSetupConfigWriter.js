'use strict';

function wsWriteConfig_(ss, year, root, referenceRoot, daftarArsip, arsipDiklat, tahunFolder, inbox, templateFolder, activityRows, subActivityRows) {
  // Ensure all config sheets exist
  [
    CONFIG_SHEETS.YEARS,
    CONFIG_SHEETS.ACTIVITIES,
    CONFIG_SHEETS.SUB_ACTIVITIES,
    CONFIG_SHEETS.METADATA_FIELDS
  ].forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // Ensure log sheets exist with headers
  wsEnsureSheetWithHeaders_(ss, CONFIG_SHEETS.ARCHIVE_LOG, [
    'archive_id', 'year', 'activity_id', 'sub_activity_id', 'source_file_id', 'final_file_id', 'final_file_name',     'target_folder_id', 'target_folder_name', 'target_folder_path', 'spreadsheet_file_id', 'spreadsheet_row_number',     'status', 'created_at', 'created_by', 'error_message', 'metadata_json'
  ]);
  wsEnsureSheetWithHeaders_(ss, CONFIG_SHEETS.ADMIN_AUDIT_LOG, [
    'created_at', 'actor', 'action', 'year', 'activity_id', 'sub_activity_id', 'folder_id', 'status', 'message'
  ]);

  // Upsert config_years (year is at index 0)
  const yearsHeaders = ['year', 'status', 'root_folder_id', 'reference_root_folder_id', 'daftar_arsip_folder_id',   'spreadsheet_year_folder_id', 'persuratan_year_folder_id', 'inbox_folder_id', 'template_folder_id', 'root_url'];
  const yearsNewRows = [[year, 'PROD_ACTIVE', root.getId(), referenceRoot ? referenceRoot.getId() : '', daftarArsip.getId(),   arsipDiklat.getId(), tahunFolder.getId(), inbox.getId(), templateFolder.getId(), root.getUrl()]];
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.YEARS), yearsHeaders, year, 0, yearsNewRows);

  // Upsert config_activities (year is at index 0)
  const activitiesHeaders = ['year', 'activity_id', 'activity_name', 'laci_no', 'folder_no', 'spreadsheet_file_id',   'target_folder_id', 'laci_folder_id', 'is_active', 'sort_order'];
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.ACTIVITIES), activitiesHeaders, year, 0, activityRows);

  // Upsert config_sub_activities (year is at index 0)
  const subActivitiesHeaders = SUB_ACTIVITY_HEADERS;
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.SUB_ACTIVITIES), subActivitiesHeaders, year, 0, subActivityRows);

  // Upsert config_metadata_fields (year is at index 1)
  const metadataHeaders = ['field_id', 'year', 'activity_id', 'field_name', 'spreadsheet_column_name', 'is_required',   'default_value', 'input_type', 'sort_order', 'is_visible_in_form'];
  const metadataNewRows = wsBuildMetadataRows_(year).slice(1);
  wsUpsertSheetData_(ss.getSheetByName(CONFIG_SHEETS.METADATA_FIELDS), metadataHeaders, year, 1, metadataNewRows);
}

function wsBuildMetadataRows_(year) {
  const rows = [['field_id', 'year', 'activity_id', 'field_name', 'spreadsheet_column_name', 'is_required', 'default_value',   'input_type', 'sort_order', 'is_visible_in_form']];
  const fields = [
    ['no_berkas', 'No Berkas', true, '', 'number'],
    ['nomor_item_arsip', 'Nomor Item Arsip', true, '', 'text'],
    ['nomor_surat', 'Nomor Surat', true, '', 'text'],
    ['jenis_naskah', 'Jenis Naskah', false, 'Naskah Masuk', 'select'],
    ['kode_klasifikasi', 'Kode Klasifikasi', false, '', 'text'],
    ['uraian_informasi_item', 'Uraian Informasi Item', true, '', 'textarea'],
    ['tanggal', 'Tanggal', false, '', 'date'],
    ['tingkat_perkembangan', 'Tingkat Perkembangan', true, 'Asli', 'select'],
    ['jumlah', 'Jumlah', true, '1', 'number'],
    ['satuan', 'Satuan', true, 'Lembar', 'text'],
    ['no_filing_cabinet', 'No Filing Cabinet', true, '', 'number'],
    ['no_laci', 'No Laci', true, '', 'number'],
    ['no_folder', 'No Folder', true, '', 'number'],
    ['klasifikasi_akses', 'Klasifikasi Keamanan & Akses Arsip', true, 'Terbuka', 'select'],
    ['ket', 'Ket.', false, '', 'text'],
    ['lokasi_simpan', 'Lokasi Simpan', true, '', 'text']
  ];

  WORKSPACE_ACTIVITIES.forEach((activity) => {
    fields.forEach((field, index) => {
      rows.push([
        activity.id + '_' + field[0],
        year,
        activity.id,
        field[0],
        field[1],
        field[2] ? 'TRUE' : 'FALSE',
        field[0] === 'no_laci' ? activity.laciNo : field[3],
        field[4],
        index + 1,
        'TRUE'
      ]);
    });
  });
  return rows;
}

