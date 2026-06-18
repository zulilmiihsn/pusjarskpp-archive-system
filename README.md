# Portal Arsip PUSJARSKPP

Dokumentasi dan implementasi aplikasi pengelolaan arsip persuratan PUSJARSKPP berbasis Google Apps Script, Google Drive, dan Google Sheets.

Fokus production yang sudah disepakati:

- Persuratan tahun 2026 dan siap dikembangkan multi-tahun.
- Output utama di folder `1. Arsip Latbang`.
- Empat kategori kegiatan: Kepemimpinan, Latsar CPNS, Teknis, dan Lain-lain.
- Sub-menu mengikuti struktur kegiatan/angkatan/batch.
- Metadata arsip mengikuti spreadsheet laci masing-masing.
- Template surat tersedia sebagai menu app.
- Lampiran dokumen esensial disiapkan sebagai modul lanjutan setelah aturan pemilihannya final.

Dokumen pra-development ada di:

- [Product Brief](docs/pre-development/01-product-brief.md)
- [Menu and User Flow](docs/pre-development/03-menu-and-flow.md)
- [Google Apps Script Architecture](docs/pre-development/04-google-apps-script-architecture.md)
- [Data Model and Metadata Schema](docs/pre-development/05-data-model-and-metadata-schema.md)
- [Client Validation Checklist](docs/pre-development/06-client-validation-checklist.md)
- [Development Backlog](docs/pre-development/07-development-backlog.md)
- [Production Scope](docs/pre-development/12-production-scope.md)

Implementasi production Google Apps Script ada di:

- `Code.gs`, `AppController.gs`, `Config.gs`, `DriveService.gs`, `SpreadsheetService.gs`, `MetadataService.gs`, dan `WorkspaceSetup.gs`
- `Index.html`, `ClientRouter.html`, `ClientState.html`, `ClientApi.html`, `ClientDashboard.html`, `ClientActivityDetail.html`, `ClientArchiveFolder.html`, `ClientProcess.html`, `ClientHistory.html`, `ClientTemplates.html`, `ClientAccounts.html`, `ClientSettings.html`, `ClientFolderPicker.html`, `ClientLogin.html`, `ClientUtils.html`, `ClientAssets.html`, dan `Styles.html`
- `appsscript.json` sebagai manifest Apps Script

Catatan penting:

- Web app bisa membangun workspace dari folder `1. Arsip Latbang` yang dipilih user di Pengaturan.
- Web app bekerja dari konfigurasi spreadsheet, supaya tahun dan sub-kegiatan bisa ditambah tanpa mengubah kode utama.
- Folder `scripts/` disiapkan untuk tooling tambahan; saat ini verifikasi project ada di `scripts/verify-project.js`.

## Deployment dan Verifikasi

Project ini memakai `clasp` untuk push ke Google Apps Script.

```powershell
cmd /c npm test
.\node_modules\.bin\clasp.cmd status
.\node_modules\.bin\clasp.cmd push
```

Manifest production saat ini disetel untuk akses internal domain dan eksekusi sebagai user yang sedang mengakses:

```json
{
  "webapp": {
    "executeAs": "USER_ACCESSING",
    "access": "DOMAIN"
  }
}
```

Dengan konfigurasi ini, setiap user harus punya permission Google Drive/Sheets yang relevan. Ini lebih aman untuk arsip internal dibanding menjalankan semua aksi memakai otoritas deployer.

## Safety Guard

- Finalisasi arsip memakai lock agar nomor item tidak bentrok saat beberapa user menyimpan bersamaan.
- Nama file final dibuat unik di folder tujuan jika sudah ada file dengan nama sama.
- Jika simpan gagal, app mencoba menulis log `FAILED` dengan informasi file/spreadsheet yang sempat dibuat.
- Menonaktifkan sub-kegiatan hanya mengubah config `is_active` menjadi `FALSE`; folder Drive dan sheet arsip tidak dipindahkan ke Bin.
- Aksi `Hapus` tersedia terpisah. Aksi ini memindahkan folder Drive ke Bin dan menyimpan `inactive_reason = drive_trashed`, sehingga folder yang direstore dari Drive bisa muncul lagi setelah `Sync Drive`.
- Config sub-kegiatan yang dihapus ke Bin menyimpan `inactive_at`. Trigger harian `cleanupTrashedSubActivities` akan menghapus baris config `drive_trashed` setelah 30 hari.
- Tombol maintenance di Pengaturan dapat memasang trigger cleanup harian dan menjalankan cleanup manual.
- Halaman `Nonaktif` menampilkan sub-kegiatan nonaktif, dengan aksi restore dan purge config manual.
- Halaman `Audit Log` menampilkan aksi admin seperti nonaktifkan, hapus, restore, purge, install trigger, dan cleanup.
