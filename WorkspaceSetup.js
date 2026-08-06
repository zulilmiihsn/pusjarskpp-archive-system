'use strict';

const WORKSPACE_CONFIG = {
  systemFolderName: '00. Sistem Portal',
  configSpreadsheetName: 'Portal Arsip PUSJARSKPP - Config Production',
  inboxFolderName: 'Inbox Dokumen Masuk',
  templateFolderName: 'Template Surat'
};

const WORKSPACE_ACTIVITIES = [
  {
    id: 'kepemimpinan',
    label: 'Kepemimpinan',
    laciNo: '1',
    folderNo: '1',
    sortOrder: 1,
    laciFolderName: 'LACI NO 1 (KEPEMIMPINAN)',
    targetFolderName: 'Folder 1 (Pelatihan Kepemimpinan)',
    laciCandidates: ['LACI NO 1', 'KEPEMIMPINAN'],
    targetCandidates: ['Folder 1', 'Kepemimpinan'],
    hasRekapSheet: true,
    defaultSubActivities: [
      'PKN Tk.II Angkatan X Tahun {year}',
      'PKA Angkatan 1 Tahun {year}',
      'PKA Angkatan 2 Tahun {year}',
      'PKP Angkatan 1 Tahun {year}',
      'PKP Angkatan 2 Tahun {year}'
    ]
  },
  {
    id: 'latsar_cpns',
    label: 'Latsar CPNS',
    laciNo: '2',
    folderNo: '2',
    sortOrder: 2,
    laciFolderName: 'LACI NO 2 (LATSAR CPNS)',
    targetFolderName: 'Folder 2 (Latsar CPNS)',
    laciCandidates: ['LACI NO 2', 'LATSAR'],
    targetCandidates: ['Folder 2', 'Latsar'],
    hasRekapSheet: true,
    defaultCode: '',
    defaultSubActivities: [
      '01. Latsar CPNS Angkatan I',
      '02. Latsar CPNS Angkatan II',
      '03. Latsar CPNS Angkatan III',
      '04. Latsar CPNS Angkatan IV',
      '05. Latsar CPNS Angkatan V',
      '06. Latsar CPNS Angkatan VI',
      '07. Latsar CPNS Angkatan VII',
      '08. Latsar CPNS Angkatan VIII',
      '09. Latsar CPNS Angkatan IX',
      '10. Latsar CPNS Angkatan X',
      '11. Latsar CPNS Angkatan XI',
      '12. Latsar CPNS Angkatan XII',
      '13. Latsar CPNS Kutai Timur (Kerjasama)',
      '14. Latsar CPNS Bengkayang (Kerjasama)'
    ]
  },
  {
    id: 'teknis',
    label: 'Teknis',
    laciNo: '3',
    folderNo: '3',
    sortOrder: 3,
    laciFolderName: 'LACI NO 3 (TEKNIS)',
    targetFolderName: 'Folder 3 (Pelatihan Teknis dan Lainnya)',
    laciCandidates: ['LACI NO 3', 'TEKNIS'],
    targetCandidates: ['Folder 3', 'Teknis'],
    hasRekapSheet: true,
    allowNonLetter: true,
    defaultSubActivities: [
      '01. Coaching bagi ASN',
      '02. Puskesmas Prima',
      '03. Desa Madani',
      '04. Sekolah Idaman',
      '10. Pelatihan Teknis Lainnya'
    ]
  },
  {
    id: 'lain_lain',
    label: 'Lain-lain',
    laciNo: '4',
    folderNo: '4',
    sortOrder: 4,
    laciFolderName: 'LACI NO 4 (LAIN-LAIN)',
    targetFolderName: 'Folder 4 (Lain-lain)',
    laciCandidates: ['LACI NO 4', 'LAIN-LAIN'],
    targetCandidates: ['Folder 4', 'Lain-lain'],
    hasRekapSheet: true,
    allowNonLetter: true,
    defaultSubActivities: [
      '01. Jabatan Fungsional',
      '02. Data Jabatan Fungsional',
      '03. Permohonan Fasilitator',
      '04. Kerjasama Pelatihan',
      '05. Evaluasi Pasca Pelatihan',
      '06. SK Tim Kerja Pusjar SKPP',
      '07. Workshop PKN-II'
    ]
  }
];



const WORKSPACE_REKAP_HEADERS = [
  'Nomor Berkas',
  'Kode Klasifikasi',
  'Uraian/ Informasi Arsip',
  'Kurun Waktu',
  'Jumlah',
  'No Filing Cabinet',
  'No Laci',
  'No Folder',
  'Keamanan & Akses Arsip',
  'Ket',
  'Data Fix Peserta',
  'Kumpulan Materi',
  'Laporan',
  'Sertifikat'
];

const WorkspaceSetupService = {
  parseLeadingNumber: function (value) {
    return WorkspaceSetupImpl_.parseLeadingNumber.apply(this, arguments);
  },

  initialize: function (payload) {
    return WorkspaceSetupImpl_.initialize.apply(this, arguments);
  },

  initializeSingleYear_: function (ss, year, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report) {
    return WorkspaceSetupImpl_.initializeSingleYear_.apply(this, arguments);
  },

  scanAndImportPhysicalYears_: function (ss, root, daftarArsip, naskahDinas, persuratan, inbox, templateFolder, report,   startTime) {
    return WorkspaceSetupImpl_.scanAndImportPhysicalYears_.apply(this, arguments);
  }
};
