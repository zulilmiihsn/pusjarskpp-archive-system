'use strict';

/** Keys used in ScriptProperties for workspace settings. */
const PROP_KEYS = {
  CONFIG_SPREADSHEET_ID: 'CONFIG_SPREADSHEET_ID',
  CURRENT_YEAR: 'CURRENT_YEAR',
  WORKSPACE_ROOT_FOLDER_ID: 'WORKSPACE_ROOT_FOLDER_ID',
  REFERENCE_ROOT_FOLDER_ID: 'REFERENCE_ROOT_FOLDER_ID',
  SYSTEM_FOLDER_PARENT_ID: 'SYSTEM_FOLDER_PARENT_ID',
  SYSTEM_FOLDER_NAME: 'SYSTEM_FOLDER_NAME',
  CONFIG_FOLDER_ID: 'CONFIG_FOLDER_ID',
  DEFAULT_FILING_CABINET: 'DEFAULT_FILING_CABINET',
  SCHEMA_VERSION: 'SCHEMA_VERSION'
};

const DEFAULT_YEAR = new Date().getFullYear();
const CURRENT_SCHEMA_VERSION = 1;
const TRASHED_SUB_ACTIVITY_RETENTION_DAYS = 30;
const SESSION_TTL_MS = 2 * 24 * 60 * 60 * 1000;
// CacheService mempercepat validasi request. Script Properties tetap menjadi
// penyimpanan durable selama 2 hari; cache 6 jam hanya lapisan baca cepat.
const SESSION_CACHE_TTL_SECONDS = 6 * 60 * 60;
const DEFAULT_SUB_ACTIVITY_KODE_KLASIFIKASI = 'PDP.07.1';

const HASH_ITERATIONS = 800;
const HASH_ITERATIONS_V2 = 50000;
const HASH_ITERATIONS_V3 = 5000;
const HASH_PREFIX_V1 = 'v1:';
const HASH_PREFIX_V2 = 'v2:';
const HASH_PREFIX_V3 = 'v3:';
const SESSION_KEY = 'portal_arsip_session';

/* ── Cross-cutting tuneables (magic numbers, once) ── */
const LOCK_TIMEOUT_MS       = 15000;   // LockService max wait (original default)
const RETRY_MAX_ATTEMPTS    = 3;
const MAX_FILENAME_LENGTH   = 255;
const MAX_INPUT_LENGTH      = 1000;
const MAX_ID_LENGTH         = 200;
const MAX_URAIAN_LENGTH     = 260;
const MAX_FILE_PART_LENGTH  = 170;
const MAX_SHEET_NAME_LENGTH = 90;

const CONFIG_SHEETS = {
  YEARS: 'config_years',
  ACTIVITIES: 'config_activities',
  SUB_ACTIVITIES: 'config_sub_activities',
  METADATA_FIELDS: 'config_metadata_fields',
  ARCHIVE_LOG: 'archive_log',
  ADMIN_AUDIT_LOG: 'admin_audit_log',
  ACCOUNTS: 'config_accounts',
  TEMPLATE_CATEGORIES: 'config_template_categories',
  TEMPLATE_CATEGORY_MAP: 'config_template_category_map',
  DOCUMENT_TYPES: 'config_document_types'
};

const ARCHIVE_FIELD_KEYS = [
  'no_berkas',
  'nomor_item_arsip',
  'kode_klasifikasi',
  'uraian_informasi_item',
  'tanggal',
  'tingkat_perkembangan',
  'jumlah',
  'satuan',
  'no_filing_cabinet',
  'no_laci',
  'no_folder',
  'klasifikasi_akses',
  'ket',
  'lokasi_simpan'
];

const ARCHIVE_FIELD_LABELS = [
  'No Berkas',
  'Nomor Item Arsip',
  'Kode Klasifikasi',
  'Uraian Informasi Item',
  'Tgl',
  'Tingkat Perkembangan',
  'Jumlah',
  'Satuan',
  'No Lemari Arsip',
  'No Laci',
  'No Folder',
  'Klasifikasi Keamanan & Akses Arsip',
  'Ket.',
  'Lokasi Simpan'
];

const STATUS = {
  DRAFT: 'DRAFT',
  COMPLETED: 'COMPLETED',
  FAILED: 'FAILED'
};

const DEFAULT_TEMPLATE_CATEGORY_COLOR = '#2563EB';

const SUB_ACTIVITY_INACTIVE_REASON = {
  MANUAL: 'manual',
  DRIVE_MISSING: 'drive_missing',
  DRIVE_TRASHED: 'drive_trashed'
};

const SUB_ACTIVITY_HEADERS = [
  'year',
  'activity_id',
  'sub_activity_id',
  'sub_activity_name',
  'formal_archive_name',
  'folder_id',
  'folder_path',
  'no_folder',
  'default_kode_klasifikasi',
  'allow_non_letter_document',
  'is_active',
  'sort_order',
  'target_sheet_name',
  'mapping_status',
  'mapping_note',
  'rekap_row_number',
  'inactive_reason',
  'inactive_at',
  'parent_folder_id',
  'parent_folder_name',
  'parent_folder_path',
  'spreadsheet_file_id',
  'metadata_locks',
  'local_sort_order'
];

const ACCOUNT_HEADERS = ['account_id', 'username', 'password_hash', 'role', 'display_name', 'is_active', 'created_at', 'updated_at'];

/* ── Layout sheet Detail & Rekap (dipakai SpreadsheetService.gs + SheetHelpers.gs) ── */

const DETAIL_FIELD_ORDER = [
  'no_berkas',
  'nomor_item_arsip',
  'kode_klasifikasi',
  'uraian_informasi_item',
  'tanggal',
  'tingkat_perkembangan',
  'jumlah',
  'jumlah_satuan',
  'no_filing_cabinet',
  'no_laci',
  'no_folder',
  'klasifikasi_akses',
  'lokasi_simpan'
];

/** Default name for detail sheets. */
const DEFAULT_DETAIL_SHEET_NAME = 'Daftar Isi Berkas Arsip Aktip';
const REKAP_SHEET_NAME = 'Daftar Berkas Arsip Aktip';
const REKAP_DATA_START_ROW = 8;
const REKAP_HEADER_ROW = 5;
const REKAP_SUBHEADER_ROW = 6;
const REKAP_NUMBERING_ROW = 7;
const REKAP_FALLBACK_START_COL = 2;
const REKAP_DOC_COLUMNS = [
  { key: 'data_fix_peserta', label: 'Data Fix Peserta', match: ['data fix peserta'] },
  { key: 'kumpulan_materi', label: 'Kumpulan Materi', match: ['kumpulan materi'] },
  { key: 'laporan', label: 'Laporan', match: ['laporan'] },
  { key: 'sertifikat', label: 'Sertifikat', match: ['sertifikat'] },
  { key: 'evaluasi_penyelenggaraan', label: 'Evaluasi Penyelenggaraan', match: ['evaluasi penyelenggaraan'] },
  { key: 'evaluasi_peserta', label: 'Evaluasi Peserta', match: ['evaluasi peserta'] },
  { key: 'dokumentasi_rapat', label: 'Dokumentasi Rapat', match: ['dokumentasi rapat'] },
  { key: 'berita_acara', label: 'Berita Acara', match: ['berita acara'] },
  { key: 'piagam_penghargaan', label: 'Piagam Penghargaan', match: ['piagam penghargaan'] },
  { key: 'video_pengajar', label: 'Video Pengajar/Fasilitator', match: ['video pengajar', 'video fasilitator', 'video pengajar/fasilitator'] }
];
const REKAP_SUMMARY_COLUMNS = {
  nomorBerkas: ['nomor berkas'],
  kodeKlasifikasi: ['kode klasifikasi'],
  uraian: ['uraian informasi arsip', 'uraian informasi'],
  kurunWaktu: ['kurun waktu'],
  jumlah: ['jumlah'],
  filingCabinet: ['no filing cabinet'],
  noLaci: ['no laci'],
  noFolder: ['no folder'],
  akses: ['keamanan akses arsip', 'ket klasifikasi keamanan akses arsip', 'klasifikasi keamanan akses arsip'],
  ket: ['ket']
};

/** Template header row number. */
const DETAIL_HEADER_ROW = 6;
/** First row where data starts. */
const DETAIL_DATA_START_ROW = 9;
/** Fallback column to start writing from (B). */
const DETAIL_FALLBACK_START_COL = 2;
/** Offset within row for item-number column. */
const DETAIL_ITEM_NUMBER_OFFSET = 1;
/** Column offset used for writable-blank-row detection. */
const DETAIL_WRITABLE_CHECK_OFFSET = 3;
/** Fallback row number for instruction notes. */
const DETAIL_NOTE_FALLBACK_ROW = 34;
