# Development Backlog

## Phase 0 - Preparation

- [ ] Validasi scope MVP ke client.
- [ ] Validasi struktur menu dan sub-menu.
- [ ] Validasi metadata tiap laci.
- [ ] Validasi pola nama file final.
- [ ] Siapkan Google Sheet internal `Portal Arsip - App Config`.
- [ ] Isi `config_activity`.
- [ ] Isi `config_sub_activity`.
- [ ] Isi `config_schema_fields`.
- [ ] Isi `config_naming_templates`.

## Phase 1 - Prototype UI

- [ ] Dashboard 4 kegiatan.
- [ ] Halaman detail kegiatan.
- [ ] Halaman sub-kegiatan/angkatan.
- [ ] Form upload/pilih dokumen.
- [ ] Halaman review metadata dummy.
- [ ] Preview nama file final.
- [ ] Preview row spreadsheet.

Goal phase 1: client bisa klik flow tanpa integrasi penuh.

## Phase 2 - Google Drive Read Integration

- [ ] Read config dari Google Sheets.
- [ ] List folder kegiatan.
- [ ] List folder output arsip.
- [ ] Pilih file dari Google Drive.
- [ ] Ambil metadata file Drive.

## Phase 3 - Document Text Extraction

- [ ] Extract text dari Google Docs.
- [ ] Extract text dari DOCX jika tersedia via conversion/export.
- [ ] Extract text dari PDF digital.
- [ ] Simpan hasil ekstraksi ke queue.
- [ ] Tampilkan preview teks.

## Phase 4 - Metadata Parser

- [ ] Parse nomor surat.
- [ ] Parse tanggal surat.
- [ ] Parse perihal.
- [ ] Parse pengirim/instansi.
- [ ] Parse tujuan.
- [ ] Hitung confidence.
- [ ] Field dengan confidence rendah wajib review manual.

## Phase 5 - Review and Naming

- [ ] Form metadata dinamis berdasarkan schema.
- [ ] Generate nomor item arsip.
- [ ] Generate nama file final.
- [ ] Sanitasi nama file.
- [ ] Deteksi duplikasi nama.
- [ ] Preview row spreadsheet.

## Phase 6 - Final Write

- [ ] Copy file sumber ke folder output.
- [ ] Rename file copy menjadi nama final.
- [ ] Write row ke spreadsheet arsip.
- [ ] Update queue status.
- [ ] Write process log.
- [ ] Tampilkan link file final dan spreadsheet.

## Phase 7 - Hardening

- [ ] Error handling.
- [ ] Retry untuk Drive/Sheets API.
- [ ] Role/access check.
- [ ] Audit log detail.
- [ ] Backup/export config.
- [ ] Manual rollback guidance.

## Backlog Setelah Meeting Client

- [ ] Modul lampiran dokumen kegiatan.
- [ ] Modul perwakilan peserta.
- [ ] Modul template surat Word.
- [ ] Modul template dokumen.
- [ ] OCR untuk scan/foto.
- [ ] AI fallback untuk metadata sulit.

