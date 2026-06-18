# Production Scope

Arahnya sekarang production-ready app untuk workflow persuratan Latbang 2026 dan multi-tahun berikutnya.

## Production Folder

App production bekerja di:

```text
1. Arsip Latbang
|-- 1. Daftar Arsip (Spreadsheet)
|   `-- ARSIP DIKLAT 2026
|       |-- LACI NO 1
|       |-- LACI NO 2
|       |-- LACI NO 3
|       `-- LACI NO 4
`-- 2. Naskah Dinas Latbang (Dokumen)
    `-- 1. Persuratan
        `-- Tahun 2026
            |-- Folder 1 Kepemimpinan
            |-- Folder 2 Latsar CPNS
            |-- Folder 3 Teknis
            `-- Folder 4 Lain-lain
```

Folder `4. Kegiatan Pelatihan dan Pengembangan` tetap menjadi referensi/sumber kegiatan dan contoh dokumen, bukan folder output app.

## Modul Production

1. Dashboard progress arsip.
2. Proses surat baru.
3. Review metadata.
4. Generate nama file final.
5. Simpan file ke folder Persuratan production.
6. Catat baris ke spreadsheet Laci terkait.
7. Riwayat arsip dari `archive_log`.
8. Template surat, dengan tombol buka dan download Word jika template berupa Google Docs.
9. Pengaturan config, tahun kerja, dan tambah sub-kegiatan.

## Scope Yang Tetap Dikeluarkan

App tidak masuk ke folder top-level:

- `Notulensi`
- `SK/Surat Tugas`

Jika ada subfolder di bawah jalur `Persuratan`, production setup dapat memetakannya sebagai sub-kegiatan karena secara lokasi ia sudah berada di scope Persuratan.

## Konfigurasi Jangka Panjang

Tahun, kegiatan, sub-kegiatan, folder tujuan, spreadsheet tujuan, inbox, dan folder template dikendalikan dari spreadsheet:

```text
Portal Arsip Latbang - Config Production
```

Untuk tahun berikutnya, baris baru dapat ditambahkan ke:

- `config_years`
- `config_activities`
- `config_sub_activities`
- `config_metadata_fields`

Dengan pola ini, kode web app tidak perlu diubah hanya untuk menambah tahun atau sub-kegiatan.

## Workspace Selection

Production root tidak di-hardcode di aplikasi. User memilih workspace dengan memasukkan link atau ID folder:

```text
1. Arsip Latbang
```

Dari folder itu, app membangun:

- `00. Portal Arsip App Config`
- `Portal Arsip Latbang - Config Production`
- `03. Inbox Dokumen Masuk`
- `99. Template Surat`
- mapping Laci 1-4
- mapping folder Persuratan tahun aktif
- mapping sub-kegiatan yang ada di bawah Folder 1-4
