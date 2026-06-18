# Visual Map Pre-Development Portal Arsip

Dokumen ini merangkum rancangan pre-development dalam bentuk visual agar mudah dipakai untuk diskusi dan validasi ke client.

## 1. Peta Scope Sistem

```mermaid
flowchart TD
    A["Google Drive Client"] --> B["1. Arsip Latbang"]
    A --> C["4. Kegiatan Pelatihan dan Pengembangan"]

    C --> C1["Sumber kegiatan / bahan kerja"]
    C1 --> C2["Kegiatan Latbang 2026"]
    C2 --> C3["Dokumen mentah / contoh / bahan yang perlu diarsipkan"]

    B --> B1["1. Daftar Arsip (Spreadsheet)"]
    B --> B2["2. Naskah Dinas Latbang (Dokumen)"]

    B1 --> B1A["ARSIP DIKLAT 2026"]
    B1A --> L1["Laci 1: Kepemimpinan"]
    B1A --> L2["Laci 2: Latsar CPNS"]
    B1A --> L3["Laci 3: Teknis"]
    B1A --> L4["Laci 4: Lain-lain"]

    B2 --> B2A["1. Persuratan"]
    B2A --> B2B["Tahun 2026"]
    B2B --> F1["Folder 1: Pelatihan Kepemimpinan"]
    B2B --> F2["Folder 2: Latsar CPNS"]
    B2B --> F3["Folder 3: Pelatihan Teknis dan Lainnya"]
    B2B --> F4["Folder 4: Lain-lain"]

    B2 --> OUT1["2. Notulensi: out of MVP"]
    B2 --> OUT2["3. SK/Surat Tugas: out of MVP"]
```

## 2. Hubungan Folder dan Output Kerja

```mermaid
flowchart LR
    A["Dokumen dari WhatsApp / divisi / folder kegiatan"] --> B["Dipilih sebagai dokumen persuratan"]
    B --> C["Dibaca dan diambil metadata"]
    C --> D["Dikategorikan ke kegiatan"]

    D --> K1["Kepemimpinan"]
    D --> K2["Latsar CPNS"]
    D --> K3["Teknis"]
    D --> K4["Lain-lain"]

    K1 --> S1["Catat ke Spreadsheet Laci 1"]
    K2 --> S2["Catat ke Spreadsheet Laci 2"]
    K3 --> S3["Catat ke Spreadsheet Laci 3"]
    K4 --> S4["Catat ke Spreadsheet Laci 4"]

    K1 --> F1["Simpan ke Folder 1 Persuratan 2026"]
    K2 --> F2["Simpan ke Folder 2 Persuratan 2026"]
    K3 --> F3["Simpan ke Folder 3 Persuratan 2026"]
    K4 --> F4["Simpan ke Folder 4 Persuratan 2026"]
```

## 3. Flow Kerja User di Aplikasi

```mermaid
flowchart TD
    A["Dashboard"] --> B["Pilih Kegiatan"]
    B --> C1["Kepemimpinan"]
    B --> C2["Latsar CPNS"]
    B --> C3["Teknis"]
    B --> C4["Lain-lain"]

    C1 --> D1["Pilih PKN / PKA / PKP"]
    C2 --> D2["Pilih Angkatan / Gelombang"]
    C3 --> D3["Pilih Program / Batch"]
    C4 --> D4["Pilih Tema / Subtema"]

    D1 --> E["Pilih atau upload dokumen"]
    D2 --> E
    D3 --> E
    D4 --> E

    E --> F["Ekstrak metadata"]
    F --> G["Tampilkan draft metadata"]
    G --> H{"User validasi?"}
    H -->|Revisi| G
    H -->|Setuju| I["Generate nama file final"]
    I --> J["Simpan / pindahkan ke folder tujuan"]
    J --> K["Tambahkan baris ke spreadsheet daftar arsip"]
    K --> L["Status: selesai"]
```

## 4. Struktur Menu MVP

```mermaid
mindmap
  root((Portal Arsip Latbang))
    Dashboard
      Antrian dokumen
      Progress arsip
      Riwayat proses
    Kepemimpinan
      PKN
      PKA
      PKP
      Lihat daftar arsip
      Proses dokumen baru
    Latsar CPNS
      Angkatan I-XII
      Kutim 1
      Kutim 2
      Tahap dokumen
      Proses dokumen baru
    Teknis
      Coaching Batch 1
      Coaching Batch 2
      Desa Madani
      Puskesmas Prima
      Proses dokumen baru
    Lain-lain
      Jabatan Fungsional
      Data Pejabat Fungsional
      Fasilitator
      Kerja Sama / Katalog
      Proses dokumen baru
    Pengaturan
      Folder tujuan
      Spreadsheet tujuan
      Schema metadata
      Pola nama file
```

## 5. Relasi Config Metadata ke Spreadsheet

```mermaid
flowchart TD
    A["config_activity"] --> A1["Kepemimpinan = Laci 1"]
    A --> A2["Latsar CPNS = Laci 2"]
    A --> A3["Teknis = Laci 3"]
    A --> A4["Lain-lain = Laci 4"]

    B["config_sub_activity"] --> B1["PKN / PKA / PKP"]
    B --> B2["Angkatan / Gelombang"]
    B --> B3["Program / Batch"]
    B --> B4["Tema / Subtema"]

    C["config_schema_fields"] --> C1["Field umum"]
    C --> C2["Field bantu per kegiatan"]
    C --> C3["Required / optional"]
    C --> C4["Default value"]

    A1 --> D["Form metadata dinamis"]
    A2 --> D
    A3 --> D
    A4 --> D
    B1 --> D
    B2 --> D
    B3 --> D
    B4 --> D
    C1 --> D
    C2 --> D

    D --> E["Baris spreadsheet daftar arsip"]
    E --> F["No Berkas"]
    E --> G["Nomor Item Arsip"]
    E --> H["Kode Klasifikasi"]
    E --> I["Uraian informasi Berkas"]
    E --> J["Tgl"]
    E --> K["Tingkat Perkembangan"]
    E --> L["Jumlah"]
    E --> M["Lokasi"]
    E --> N["Klasifikasi Akses"]
    E --> O["Lokasi Simpan"]
```

## 6. Screen Map MVP

```mermaid
flowchart TD
    A["Login / akses internal"] --> B["Dashboard"]
    B --> C["Pilih Kegiatan"]
    C --> D["Daftar Sub-kegiatan"]
    D --> E["Antrian Dokumen"]
    E --> F["Detail Dokumen"]
    F --> G["Draft Metadata"]
    G --> H["Preview Nama File"]
    H --> I["Konfirmasi Simpan"]
    I --> J["Hasil Proses"]

    B --> K["Riwayat Arsip"]
    K --> L["Detail Riwayat"]

    B --> M["Pengaturan"]
    M --> N["Config Kegiatan"]
    M --> O["Config Sub-kegiatan"]
    M --> P["Config Metadata"]
    M --> Q["Config Folder dan Spreadsheet"]
```

## 7. Batas MVP Yang Perlu Ditegaskan ke Client

```mermaid
flowchart LR
    A["MVP"] --> B["Persuratan 2026"]
    B --> C["Daftar Arsip Spreadsheet 2026"]
    B --> D["Naskah Dinas Latbang / Persuratan / Tahun 2026"]
    B --> E["4 kegiatan utama"]

    F["Belum MVP"] --> G["Notulensi"]
    F --> H["SK / Surat Tugas sebagai folder utama"]
    F --> I["Lampiran dokumen lengkap"]
    F --> J["Template surat Word"]
```
