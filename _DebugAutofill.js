'use strict';

/**
 * TEMPORARY debug helper — hapus setelah masalah tanggal selesai.
 *
 * Cara pakai (di editor Apps Script, akun yang sama dgn /dev):
 *   1. Ambil fileId surat yang tanggalnya gagal terbaca.
 *      (buka file di Drive → URL .../d/<FILE_ID>/view)
 *   2. Jalankan: debugAutofillText('<FILE_ID>')
 *   3. Lihat View > Logs (Ctrl+Enter), copy bagian "TANGGAL", "YEAR HITS",
 *      dan "HEADER" lalu tempel ke chat.
 */
function debugAutofillText(fileId) {
  const id = String(fileId || '').replace(/[^a-zA-Z0-9_\-]/g, '');
  const file = DriveApp.getFileById(id);
  const mime = file.getMimeType();
  const conv = extractTextViaConversion_(file, mime);
  const text = conv.text || '';
  const res = ParseEngine.analyze(text, file.getName(), {});
  const norm = ParseEngine.normalizeText(text);
  const struct = ParseEngine.analyzeStructure(norm);

  // Semua potongan teks di sekitar 4-digit tahun (20xx) — untuk lihat format tanggal asli.
  const yearHits = (text.match(/[\s\S]{0,40}20[12]\d[\s\S]{0,8}/g) || []).slice(0, 25);

  Logger.log('=== META ===');
  Logger.log('name=' + file.getName() + ' | mime=' + mime + ' | size=' + file.getSize());
  Logger.log('method=' + conv.method + ' | textLen=' + text.length);
  Logger.log('=== TANGGAL (hasil parser) ===');
  Logger.log(JSON.stringify(res.fields.tanggal));
  Logger.log('=== HEADER LINES ===');
  Logger.log(struct.header.lines.join('\n'));
  Logger.log('=== BODY (10 baris pertama) ===');
  Logger.log(struct.body.lines.slice(0, 10).join('\n'));
  Logger.log('=== SIGNATURE LINES ===');
  Logger.log(struct.signature.lines.join('\n'));
  Logger.log('=== YEAR HITS (sekitar tiap 20xx) ===');
  Logger.log(yearHits.join('\n----\n'));

  return {
    method: conv.method,
    textLen: text.length,
    tanggal: res.fields.tanggal || null,
    headerLines: struct.header.lines,
    bodyFirst10: struct.body.lines.slice(0, 10),
    signatureLines: struct.signature.lines,
    yearHits: yearHits
  };
}
