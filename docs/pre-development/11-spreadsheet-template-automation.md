# Spreadsheet Template Automation

Dokumen ini merancang cara Portal Arsip Latbang membuat dan mengendalikan spreadsheet daftar arsip dengan format seperti template existing client.

## Sumber Referensi

Folder:

```text
1. Arsip Latbang
> 1. Daftar Arsip (Spreadsheet)
> ARSIP DIKLAT 2026
> LACI NO 2 (LATSAR CPNS)
```

File:

```text
Daftar Isi Berkas Arsip Laci No.2 Th. 2026.xlsx
```

File ID:

```text
1swxoBSVtNwMvHaMUWF2VxHjGpH96aG4o
```

Catatan teknis:

```text
File existing masih format Office .xlsx.
Google Sheets API tidak bisa memanipulasi grid/formatnya langsung selama masih Office file.
Untuk automation, file template harus dikonversi/copy menjadi native Google Sheets.
```

## Kesimpulan

Bisa dibuat otomatis dan dikendalikan oleh aplikasi.

Ada dua pendekatan:

1. Recommended: master template native Google Sheets.
2. Alternative: generate spreadsheet dari nol via Apps Script.

Pendekatan paling aman adalah membuat master template native Google Sheets yang sudah memiliki format sesuai contoh, lalu aplikasi melakukan copy dan mengisi datanya.

## Jenis Sheet Dalam Latsar

### 1. Sheet Rekap Berkas

Judul:

```text
Daftar Berkas Arsip Aktip
```

Struktur kolom:

```text
Nomor Berkas
Kode Klasifikasi
Uraian / Informasi Arsip
Kurun Waktu
Jumlah
Lokasi
  - No Filing Cabinet
  - No Laci
  - No Folder
Keamanan & Akses Arsip
Ket
```

Isi contoh:

```text
1 | PDP.07 | Pelatihan Dasar CPNS Angkatan I | 27 November 2025 - April 2026 | 27 dokumen | 02 | 01 | 2 | Terbatas
2 | PDP.07 | Pelatihan Dasar CPNS Angkatan II | 27 November 2025 - April 2026 | 18 dokumen | 02 | 01 | 0 | Terbatas
```

Fungsi:

```text
Merekap kelompok berkas per angkatan/kegiatan.
Untuk Latsar, satu baris bisa mewakili satu angkatan.
```

### 2. Sheet Detail Item Arsip

Judul:

```text
Daftar Isi Berkas Arsip Aktip
```

Struktur kolom:

```text
No Berkas
Nomor Item Arsip
Kode Klasifikasi
Uraian informasi Berkas
Tgl
Tingkat Perkembangan
Jumlah
Satuan
Lokasi
  - No Filing Cabinet
  - No Laci
  - No Folder
Klasifikasi Keamanan & Akses Arsip
Ket.
Lokasi Simpan
```

Fungsi:

```text
Mencatat setiap dokumen/surat di dalam angkatan/sub-kegiatan.
```

## Elemen Format Yang Perlu Direplikasi

### Header

```text
Judul sheet di tengah:
Daftar Berkas Arsip Aktip
atau
Daftar Isi Berkas Arsip Aktip
```

### Unit Pengolah

```text
Unit Pengolah : Kompartemen Latbang
```

atau pada detail sheet:

```text
Unit pengolah : Latbang
```

### Header Tabel

Format:

```text
- Background abu-abu
- Border hitam
- Text center
- Text wrap
- Header Lokasi merge di atas tiga subkolom
- Row nomor kolom kecil: 1, 2, 3, dst
```

### Area Data

Format:

```text
- Border hitam
- Nomor urut otomatis
- Text wrap untuk uraian panjang
- Tanggal bisa format tanggal
- Jumlah bisa text seperti "27 dokumen" atau angka + satuan
```

### Petunjuk Pengisian

Bagian bawah sheet:

```text
Keterangan Petunjuk Pengisian:
Kolom (1), diisi dengan nomor urut berkas;
Kolom (2), diisi dengan kode klasifikasi Arsip;
Kolom (3), diisi dengan uraian informasi dari berkas Arsip berdasarkan kegiatan dalam Klasifikasi Arsip;
Kolom (4), diisi dengan masa/kurun waktu Arsip yang tercipta;
Kolom (5), diisi dengan jumlah banyaknya Arsip dalam satuan yang sesuai dengan jenis Arsip;
Kolom (6), diisi dengan nomor Filing Cabinet;
Kolom (7), diisi dengan nomor laci pada Filing Cabinet;
Kolom (8), diisi dengan nomor folder Arsip;
Kolom (9), diisi dengan klasifikasi keamanan dan akses seperti terbuka, terbatas, dan rahasia;
Kolom (10), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik, dan digital;
```

## Recommended Automation Design

### Master Template

Buat folder internal:

```text
Portal Arsip - Templates
```

Isi:

```text
Template Daftar Berkas Arsip Aktip
Template Daftar Isi Berkas Arsip Aktip
Template Laci Kepemimpinan
Template Laci Latsar
Template Laci Teknis
Template Laci Lain-lain
```

Semua template disimpan sebagai native Google Sheets, bukan `.xlsx`.

### Create Year Flow

```text
Admin pilih Setup Tahun Baru
-> pilih tahun, misalnya 2027
-> app buat folder ARSIP DIKLAT 2027
-> app copy template Laci 1-4
-> app buat folder Persuratan Tahun 2027
-> app buat Folder 1-4
-> app isi daftar awal sub-kegiatan dari config
```

### Add Sub-Activity Flow

Contoh Latsar:

```text
Admin klik + Tambah Angkatan
-> isi nomor angkatan / nama kerjasama
-> app buat subfolder Drive
-> app tambah baris di Sheet Rekap Berkas
-> app duplicate template Sheet Detail Item Arsip
-> nama sheet mengikuti angkatan/sub-kegiatan
```

Contoh Teknis:

```text
Admin klik + Tambah Pelatihan Teknis
-> isi nama pelatihan
-> app buat subfolder Drive
-> app tambah/duplicate sheet detail
-> app set default No Laci = 3
```

## Apps Script Capability

Apps Script bisa melakukan:

```text
- Membuat spreadsheet baru
- Meng-copy spreadsheet template
- Meng-copy sheet template
- Rename sheet
- Merge cells
- Set border
- Set background
- Set font
- Set column width
- Set row height
- Set wrap text
- Append row data
- Insert row before petunjuk pengisian
- Update formula/nomor otomatis
- Export ke .xlsx jika dibutuhkan
```

Apps Script tidak ideal untuk:

```text
- Memanipulasi file .xlsx existing secara langsung tanpa konversi
- Menjaga formatting kompleks jika file masih Office compatibility mode
```

## Recommendation

Untuk project ini:

```text
1. Buat native Google Sheets master template.
2. Template meniru format existing client.
3. App hanya copy template dan isi data.
4. Jika user butuh download Excel, export hasil native Google Sheet ke .xlsx.
```

Ini lebih stabil daripada membuat ulang seluruh formatting dari nol setiap kali.

## Data Mapping

### Sheet Rekap Berkas

```text
nomor_berkas -> Nomor Berkas
kode_klasifikasi -> Kode Klasifikasi
uraian_informasi_arsip -> Uraian / Informasi Arsip
kurun_waktu -> Kurun Waktu
jumlah -> Jumlah
no_filing_cabinet -> No Filing Cabinet
no_laci -> No Laci
no_folder -> No Folder
klasifikasi_akses -> Keamanan & Akses Arsip
ket -> Ket
```

### Sheet Detail Item

```text
no_berkas -> No Berkas
nomor_item_arsip -> Nomor Item Arsip
kode_klasifikasi -> Kode Klasifikasi
uraian_informasi_berkas -> Uraian informasi Berkas
tanggal -> Tgl
tingkat_perkembangan -> Tingkat Perkembangan
jumlah -> Jumlah
satuan -> Satuan
no_filing_cabinet -> No Filing Cabinet
no_laci -> No Laci
no_folder -> No Folder
klasifikasi_akses -> Klasifikasi Keamanan & Akses Arsip
ket -> Ket.
lokasi_simpan -> Lokasi Simpan
```

