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

- Backend `.gs`: dispatcher `Code.gs`; controller (`ArchiveController`, `SubActivityController`, `SettingsController`, `AccountController`, `DriveController`, `WorkspaceController`, `TemplateController`); service/helper (`DriveService`, `SpreadsheetService`, `SheetHelpers`, `ConfigRepository`, `ConfigService`, `ConfigConstants`, `ConfigHelpers`, `MetadataService`, `ParseEngine`, `AuthService`, `CacheHelper`, `SecurityHelpers`, `PureFunctions`, `Validator`, `SystemLogger`, `VersionService`, `WorkspaceSetup`)
- Frontend HTML (di-include lewat `Index.html`): `ClientRouter`, `ClientState`, `ClientApi`, `ClientDashboard`, `ClientActivityDetail`, `EditSubActivity`, `ClientArchiveFolder`, `ClientProcess`, `ClientDocumentProcess`, `ClientHistory`, `ClientTemplates`, `ClientAccounts`, `ClientSettings`, `ClientFolderPicker`, `ClientLogin`, `ClientUtils`, `ClientAssets`/`ClientAssetsHeavy`
- Styling: `StylesBase`, `StylesLayout`, `StylesComponents`, `StylesForms`, `StylesTables`
- `appsscript.json` sebagai manifest Apps Script

Catatan penting:

- Web app bisa membangun workspace dari folder `1. Arsip Latbang` yang dipilih user di Pengaturan.
- Web app bekerja dari konfigurasi spreadsheet, supaya tahun dan sub-kegiatan bisa ditambah tanpa mengubah kode utama.
- Folder `scripts/` berisi generator aset dan verifikasi project.

## Deployment dan Verifikasi

Project ini memakai `clasp` untuk push ke Google Apps Script.

```powershell
cmd /c npm test
.\node_modules\.bin\clasp.cmd status
.\node_modules\.bin\clasp.cmd push
```

Manifest production disetel untuk akun Google lintas domain dan eksekusi sebagai pemilik deploy:

```json
{
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE"
  }
}
```

Setiap operator kantor men-deploy web app-nya sendiri memakai akun masing-masing, sehingga semua aksi Drive/Sheets berjalan dengan otoritas pemilik deploy atas workspace-nya sendiri. Model ini dipilih agar tiap kantor mengelola arsipnya secara mandiri tanpa berbagi satu deployer pusat.

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
