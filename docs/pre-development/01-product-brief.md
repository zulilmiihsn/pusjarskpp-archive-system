# Product Brief

## Project

Portal Arsip Latbang adalah aplikasi internal untuk membantu penanggung jawab Latbang mengelola arsip persuratan tahun 2026.

Aplikasi dirancang untuk berjalan di ekosistem Google:

- Google Apps Script sebagai backend dan web app.
- Google Drive sebagai sumber dan tempat arsip dokumen.
- Google Sheets sebagai database ringan, konfigurasi, antrean kerja, dan output daftar arsip.
- Google Drive API dan Google Sheets API untuk baca/tulis file dan spreadsheet.

## Problem

Client kewalahan karena proses saat ini manual:

1. Mengumpulkan surat dari WhatsApp atau berbagai divisi.
2. Membaca isi surat satu per satu.
3. Menentukan surat masuk kegiatan mana.
4. Rename file sesuai pola arsip.
5. Memindahkan atau menyimpan ke folder arsip yang benar.
6. Mencatat metadata ke spreadsheet daftar arsip.

Masalah utama bukan hanya jumlah file, tetapi banyaknya keputusan kecil berulang: klasifikasi kegiatan, penentuan folder, penamaan file, dan pencatatan metadata.

## Goal MVP

Membuat asisten persuratan yang membantu user:

- Memilih jenis kegiatan dan sub-kegiatan.
- Mengunggah atau memilih dokumen surat.
- Mengekstrak metadata utama dari dokumen.
- Menyarankan nama file final.
- Menyarankan folder tujuan.
- Menyiapkan data untuk spreadsheet daftar arsip aktif.
- Meminta user review sebelum final.

## Non-Goal MVP

Untuk tahap awal, fitur berikut belum dikerjakan sampai requirement divalidasi:

- Lampiran dokumen kegiatan awal/tengah/akhir.
- Pemilihan perwakilan peserta.
- Template surat Word.
- Template dokumen Excel.
- Notulensi.
- SK/Surat Tugas top-level.

## Product Principle

Aplikasi mengikuti mental model client:

1. Pilih kegiatan.
2. Pilih sub-kegiatan/angkatan/batch.
3. Proses surat.
4. Review hasil.
5. Simpan ke folder arsip dan spreadsheet.

Menu kegiatan bukan sekadar untuk melihat file, tetapi menjadi konteks kerja agar aplikasi tahu folder tujuan, spreadsheet tujuan, schema metadata, dan aturan nama file.

