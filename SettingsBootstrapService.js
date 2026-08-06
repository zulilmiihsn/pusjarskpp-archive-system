'use strict';

/** @private Bootstrap and workspace settings operations. */
const SettingsBootstrapImpl_ = {
  getBootstrap: function() {
    ensureSchemaMigrations_();
    const perfStartedAt = Date.now();
    const bootstrapPerf = {
      event: 'BOOTSTRAP_PERF',
      outcome: 'STARTED',
      settingsMs: 0,
      configMs: 0,
      adminCheckMs: 0,
      mapActivitiesMs: 0,
      triggerCheckMs: 0,
      inactiveCountMs: 0,
      totalMs: 0
    };
    let phaseStartedAt = Date.now();
    const settings = ConfigService.getSettings();
    bootstrapPerf.settingsMs = Date.now() - phaseStartedAt;
    if (!settings.configSpreadsheetId) {
      bootstrapPerf.outcome = 'NOT_CONFIGURED';
      bootstrapPerf.totalMs = Date.now() - perfStartedAt;
      console.info('BOOTSTRAP_PERF ' + JSON.stringify(bootstrapPerf));
      return {
        configured: false,
        settings: settings,
        message: 'Aplikasi belum punya ruang kerja. Pilih Folder Ruang Kerja di Pengaturan.'
      };
    }

    phaseStartedAt = Date.now();
    const config = CacheHelper.getConfig(settings.currentYear);
    bootstrapPerf.configMs = Date.now() - phaseStartedAt;
    if (!config || !config.activities || config.activities.length === 0) {
      bootstrapPerf.outcome = 'EMPTY_CONFIG';
      bootstrapPerf.totalMs = Date.now() - perfStartedAt;
      console.info('BOOTSTRAP_PERF ' + JSON.stringify(bootstrapPerf));
      return {
        configured: false,
        settings: settings,
        message: 'Inisialisasi sebelumnya terputus atau gagal. Silakan ulangi Inisialisasi Ruang Kerja.'
      };
    }

    // Cek READ-ONLY apakah admin ada. JANGAN buat admin di sini: getBootstrap adalah READ
    // yang dipanggil tiap load halaman. Membuat admin sbg efek samping = akun + sesi admin
    // siluman tanpa password pernah ditampilkan ke user (sumber "tiba-tiba login admin").
    // Pembuatan admin HANYA di initializeWorkspace yang menampilkan password.
    phaseStartedAt = Date.now();
    try {
      if (!hasActiveAdminAccount_()) {
        bootstrapPerf.adminCheckMs = Date.now() - phaseStartedAt;
        bootstrapPerf.outcome = 'ADMIN_NOT_FOUND';
        bootstrapPerf.totalMs = Date.now() - perfStartedAt;
        console.info('BOOTSTRAP_PERF ' + JSON.stringify(bootstrapPerf));
        return {
          configured: false,
          settings: settings,
          message: 'Akun admin belum ada. Silakan ulangi Inisialisasi Ruang Kerja untuk membuat admin & kata sandi.'
        };
      }
    } catch (e) {
      // Bila pengecekan gagal, jangan blok app — lanjutkan.
    }
    bootstrapPerf.adminCheckMs = Date.now() - phaseStartedAt;

    phaseStartedAt = Date.now();
    const activities = config.activities.map((activity) => {
      const subs = config.subActivities
        .filter(sub => sub.activity_id === activity.activity_id)
        .map((sub) => {
          const effectiveFormalName = sub.formal_archive_name || sub.sub_activity_name || '';
          const effectiveSheetName = sub.target_sheet_name || sub.sub_activity_name || '';
          const mappingStatus = sub.mapping_status || inferSubActivityMappingStatus_(sub);
          return Object.assign({}, sub, {
            effective_formal_archive_name: effectiveFormalName,
            effective_target_sheet_name: effectiveSheetName,
            effective_mapping_status: mappingStatus,
            folder: DriveService.folderDtoFromConfig(sub.folder_id, sub.folder_path || sub.sub_activity_name),
            spreadsheetFile: DriveService.fileDtoFromConfig(
              sub.spreadsheet_file_id || activity.spreadsheet_file_id,
              effectiveSheetName || activity.activity_name,
              MimeType.GOOGLE_SHEETS
            )
          });
        });
      const fields = config.fields.filter(
        field => field.activity_id === activity.activity_id && isTrue_(field.is_visible_in_form)
      );
      return Object.assign({}, activity, {
        subActivities: subs,
        fields: fields,
        laciFolder: DriveService.folderDtoFromConfig(activity.laci_folder_id, activity.activity_name || activity.laci_no),
        targetFolder: DriveService.folderDtoFromConfig(activity.target_folder_id, activity.activity_name),
        spreadsheetFile: DriveService.fileDtoFromConfig(
          activity.spreadsheet_file_id,
          activity.activity_name ? 'Daftar Arsip - ' + activity.activity_name : activity.spreadsheet_file_id,
          MimeType.GOOGLE_SHEETS
        )
      });
    });
    bootstrapPerf.mapActivitiesMs = Date.now() - phaseStartedAt;

    let triggerInstalled = false;
    phaseStartedAt = Date.now();
    try {
      const handlerName = 'runArchiveMaintenance';
      const triggers = ScriptApp.getProjectTriggers();
      triggerInstalled = triggers.some(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === handlerName);
    } catch (e) {
      console.warn('Error checking triggers: ' + e.message);
    }
    bootstrapPerf.triggerCheckMs = Date.now() - phaseStartedAt;

    let trashedCount = 0;
    let totalSubActivitiesCount = 0;
    phaseStartedAt = Date.now();
    try {
      const ss = ConfigRepository.getConfigSpreadsheet();
      const allSubs = readSheetObjects_(ss, CONFIG_SHEETS.SUB_ACTIVITIES)
        .filter(row => Number(row.year) === Number(settings.currentYear));
      totalSubActivitiesCount = allSubs.length;
      trashedCount = allSubs.filter(row => !isTrue_(row.is_active) && row.inactive_reason === 'DRIVE_TRASHED').length;
    } catch (e) {
      console.warn('Error loading inactive sub-activities: ' + e.message);
    }
    bootstrapPerf.inactiveCountMs = Date.now() - phaseStartedAt;

    const result = {
      configured: true,
      settings: settings,
      selectedYear: config.selectedYear,
      years: config.years,
      activities: activities,
      documentTypes: getRekapDocColumns_().map(function (c) {
        return { key: c.key, label: c.formLabel || c.label };
      }),
      history: config.history.slice(0, 50),
      historyMeta: {
        total: (config.historyAll || config.history).length,
        page: 1,
        totalPages: Math.ceil((config.historyAll || config.history).length / 50)
      },
      progress: this.buildProgress(config.historyAll || config.history, activities),
      maintenance: {
        triggerInstalled: triggerInstalled,
        trashedCount: trashedCount,
        totalCount: totalSubActivitiesCount
      }
    };
    bootstrapPerf.outcome = 'SUCCESS';
    bootstrapPerf.totalMs = Date.now() - perfStartedAt;
    console.info('BOOTSTRAP_PERF ' + JSON.stringify(bootstrapPerf));
    return result;
  },

  saveSettings: function (payload) {
    const settings = ConfigService.saveSettings(payload || {});
    if (settings.configSpreadsheetId) {
      ConfigRepository.getConfigSpreadsheet();
    }
    CacheHelper.invalidateAll();
    return this.getBootstrap();
  },

};
