# MVP Scope

## Scope Utama

MVP hanya menangani persuratan tahun 2026.

Folder yang menjadi fokus:

1. `1. Arsip Latbang`
   - `1. Daftar Arsip (Spreadsheet)`
   - `2. Naskah Dinas Latbang (Dokumen) > 1. Persuratan > Tahun 2026`

2. `4. Kegiatan Pelatihan dan Pengembangan`
   - Dipakai sebagai sumber bahan kerja atau referensi kegiatan tahun 2026.

## Kategori Utama

Aplikasi dibagi menjadi empat menu utama:

1. Pelatihan Kepemimpinan
2. Latsar CPNS
3. Pelatihan Teknis
4. Lain-lain

Empat menu ini berhubungan langsung dengan empat laci arsip:

| Menu | Laci | Output Spreadsheet |
| --- | --- | --- |
| Pelatihan Kepemimpinan | Laci No. 1 | Daftar Isi Berkas Arsip Laci No.1 Th. 2026 |
| Latsar CPNS | Laci No. 2 | Daftar Isi Berkas Arsip Laci No.2 Th. 2026 |
| Pelatihan Teknis | Laci No. 3 | Daftar Isi Berkas Arsip Laci No.3 Th. 2026 |
| Lain-lain | Laci No. 4 | Daftar Isi Berkas Arsip Laci No.4 Th. 2026 |

## Feature MVP

### 1. Dashboard

Menampilkan empat kegiatan utama dan status pekerjaan:

- Total dokumen draft.
- Perlu review.
- Sudah final.
- Gagal diproses.

### 2. Sub-Menu Kegiatan

Setiap menu punya daftar sub-kegiatan yang mengikuti struktur Drive.

Contoh:

- Kepemimpinan:
  - PKN Tk. II Angkatan X Tahun 2026
  - PKA Angkatan 1 Tahun 2026
  - PKA Angkatan 2 Tahun 2026
  - PKP Angkatan 1 Tahun 2026
  - PKP Angkatan 2 Tahun 2026

- Latsar CPNS:
  - Angkatan I sampai XII
  - Latsar CPNS Kutai Timur
  - Latsar CPNS Bengkayang

- Teknis:
  - Coaching bagi ASN Batch 1
  - Coaching bagi ASN Batch 2
  - Desa Madani
  - Puskesmas Prima
  - Pelatihan Teknis Lainnya

- Lain-lain:
  - Jabatan Fungsional
  - Data Jabatan Fungsional
  - Permohonan Fasilitator
  - Kerjasama Pelatihan
  - Evaluasi Pasca Pelatihan
  - Workshop PKN-II

### 3. Upload/Pilih Surat

User dapat:

- Upload PDF/DOCX.
- Pilih file dari Google Drive.
- Memproses satu file atau batch kecil.

### 4. Ekstraksi Metadata

App mencoba membaca:

- Nomor surat.
- Tanggal surat.
- Perihal.
- Asal/instansi pengirim.
- Tujuan.
- Tingkat perkembangan: Srikandi, Asli, Copy, Cetak.
- Kode klasifikasi jika bisa ditebak atau diisi manual.

### 5. Review Metadata

User wajib melihat dan mengoreksi hasil sebelum final.

Tidak ada perubahan permanen sebelum user menyetujui.

### 6. Generate Nama File Final

Pola awal:

```text
{nomor_item}. ({tingkat_perkembangan}) No: {nomor_surat}_{ringkasan_perihal}.pdf
```

Contoh:

```text
01. (Srikandi) No: B-417/BKPSDM/800.2/12/2025_ Permohonan Fasilitasi Pelatihan Struktural,Teknis, dan Fungsional.pdf
```

### 7. Simpan Output

Setelah review:

- File final disimpan ke folder `Naskah Dinas Latbang > Persuratan > Tahun 2026` sesuai kegiatan/sub-kegiatan.
- Metadata ditulis ke spreadsheet laci yang sesuai.
- Log proses disimpan di sheet internal.

## Out of Scope Untuk Sekarang

- Lampiran dokumen non-surat.
- Perwakilan peserta.
- Template surat Word.
- Template dokumen.
- Arsip tahun selain 2026.
- Notulensi.
- SK/Surat Tugas top-level.

