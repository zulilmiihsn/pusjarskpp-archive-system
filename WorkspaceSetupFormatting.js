'use strict';

function wsFormatRekapSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(7);
  sheet.getRange('B1:O1').merge().setValue('Daftar Berkas Arsip Aktip');
  sheet.getRange('B1:O1')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('B4').setValue('Unit Pengolah :');
  sheet.getRange('D4').setValue('Kompartemen Latbang');
  sheet.getRange('B4:D4')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');

  sheet.getRange(5, 2, 2, WORKSPACE_REKAP_HEADERS.length).setValues([
    ['Nomor\nBerkas', 'Kode\nKlasifikasi', 'Uraian/\nInformasi\nArsip', 'Kurun\nWaktu', 'Jumlah', 'Lokasi', '', '',     'Keamanan\n& Akses\nArsip', 'Ket', 'Data Fix\nPeserta', 'Kumpulan\nMateri', 'Laporan', 'Sertifikat'],
    ['', '', '', '', '', 'No Filing\nCabinet', 'No Laci', 'No Folder', '', '', '', '', '', '']
  ]);
  sheet.getRange('B5:B6').merge();
  sheet.getRange('C5:C6').merge();
  sheet.getRange('D5:D6').merge();
  sheet.getRange('E5:E6').merge();
  sheet.getRange('F5:F6').merge();
  sheet.getRange('G5:I5').merge();
  sheet.getRange('J5:J6').merge();
  sheet.getRange('K5:K6').merge();
  sheet.getRange('L5:L6').merge();
  sheet.getRange('M5:M6').merge();
  sheet.getRange('N5:N6').merge();
  sheet.getRange('O5:O6').merge();
  sheet.getRange(7, 2, 1, WORKSPACE_REKAP_HEADERS.length)
    .setValues([WORKSPACE_REKAP_HEADERS.map(function (_, index) { return index + 1; })]);

  wsStyleRekapSheet_(sheet);
}

function wsStyleRekapSheet_(sheet) {
  const startRow = 5;
  const startCol = 2;
  const dataRows = 14;
  const width = WORKSPACE_REKAP_HEADERS.length;
  const lastTableRow = 7 + dataRows;

  sheet.getRange(startRow, startCol, 3, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setWrap(true);
  sheet.getRange(startRow, startCol, 2, width)
    .setFontWeight('normal')
    .setBackground('#bfbfbf');
  sheet.getRange(7, startCol, 1, width)
    .setFontSize(7)
    .setFontWeight('normal')
    .setBackground('#ffffff');
  sheet.getRange(8, startCol, dataRows, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 22);
  sheet.setRowHeight(4, 24);
  sheet.setRowHeights(5, 2, 30);
  sheet.setRowHeight(7, 18);
  sheet.setRowHeights(8, dataRows, 20);
  [50, 80, 280, 120, 80, 80, 70, 70, 90, 80, 90, 90, 90, 90]
    .forEach(function (widthPx, index) {
      sheet.setColumnWidth(startCol + index, widthPx);
    });

  wsWriteRekapNotes_(sheet, lastTableRow + 2);
}

function wsWriteRekapNotes_(sheet, startRow) {
  const notes = [
    ['Keterangan Petunjuk Pengisian:'],
    ['Kolom (1), diisi dengan nomor urut berkas;'],
    ['Kolom (2), diisi dengan kode klasifikasi Arsip;'],
    ['Kolom (3), diisi dengan uraian informasi dari berkas Arsip berdasarkan kegiatan dalam Klasifikasi Arsip;'],
    ['Kolom (4), diisi dengan masa/kurun waktu Arsip yang tercipta;'],
    ['Kolom (5), diisi dengan jumlah banyaknya Arsip dalam satuan yang sesuai dengan jenis Arsip;'],
    ['Kolom (6), diisi dengan nomor Filing Cabinet;'],
    ['Kolom (7), diisi dengan nomor laci pada Filing Cabinet;'],
    ['Kolom (8), diisi dengan nomor folder Arsip;'],
    ['Kolom (9), diisi dengan klasifikasi keamanan dan akses seperti terbuka, terbatas, dan rahasia;'],
    ['Kolom (10), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik,     dan digital.']
  ];
  sheet.getRange(startRow, 3, notes.length, 1)
    .setValues(notes)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');
  sheet.getRange(startRow, 3).setFontWeight('normal');
}

function wsFormatDetailSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(8);
  sheet.getRange('B2:N2').merge().setValue('Daftar Isi Berkas Arsip Aktip');
  sheet.getRange('B2:N2')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('B4').setValue('Unit pengolah : Latbang');
  sheet.getRange('B4')
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');

  wsWriteDetailHeader_(sheet);
  wsStyleDetailSheet_(sheet);
}

function wsWriteDetailHeader_(sheet) {
  sheet.getRange('B6:B7').merge().setValue('No\nBerkas');
  sheet.getRange('C6:C7').merge().setValue('Nomor Item\nArsip');
  sheet.getRange('D6:D7').merge().setValue('Kode\nKlasifikasi');
  sheet.getRange('E6:E7').merge().setValue('Uraian Informasi Item');
  sheet.getRange('F6:F7').merge().setValue('Tgl');
  sheet.getRange('G6:G7').merge().setValue('Tingkat\nPengembangan');
  sheet.getRange('H6:I6').merge().setValue('Jumlah');
  sheet.getRange('H7').setValue('Jumlah');
  sheet.getRange('I7').setValue('Satuan');
  sheet.getRange('J6:L6').merge().setValue('Lokasi');
  sheet.getRange('J7').setValue('No Filing\nCabinet');
  sheet.getRange('K7').setValue('No Laci');
  sheet.getRange('L7').setValue('No Folder');
  sheet.getRange('M6:M7').merge().setValue('Klasifikasi\nKeamanan &\nAkses Arsip');
  sheet.getRange('N6:N7').merge().setValue('Ket.\nLokasi\nSimpan');
  sheet.getRange('B8:N8').setValues([[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]]);
}

function wsStyleDetailSheet_(sheet) {
  const startCol = 2;
  const width = 13;
  const dataRows = 24;

  // Format umum untuk header (Baris 6-8)
  sheet.getRange(6, startCol, 3, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Bookman Old Style')
    .setFontWeight('normal')
    .setBackground('#b4c6e7')
    .setWrap(true);

  // Ukuran font spesifik: Baris judul (6-7) dan Baris nomor (8)
  sheet.getRange(6, startCol, 2, width).setFontSize(11);
  sheet.getRange(8, startCol, 1, width).setFontSize(7);

  // Format data baris bawah
  sheet.getRange(9, startCol, dataRows, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);
  const itemRows = [];
  for (let i = 1; i <= dataRows; i++) itemRows.push([i]);
  sheet.getRange(9, startCol + 1, dataRows, 1)
    .setValues(itemRows)
    .setHorizontalAlignment('center');

  sheet.setRowHeight(2, 22);
  sheet.setRowHeight(4, 22);
  sheet.setRowHeights(6, 2, 31);
  sheet.setRowHeight(8, 18);
  sheet.setRowHeights(9, dataRows, 18);
  [55, 86, 78, 270, 108, 90, 42, 42, 84, 84, 84, 100, 72].forEach(function (widthPx, index) {
    sheet.setColumnWidth(startCol + index, widthPx);
  });

  wsWriteDetailNotes_(sheet, 33);
}

function wsWriteDetailNotes_(sheet, startRow) {
  const notes = [
    ['Keterangan Petunjuk Pengisian:'],
    ['Kolom (2), diisi dengan nomor item Arsip;'],
    ['Kolom (3), diisi dengan kode Klasifikasi Arsip;'],
    ['Kolom (4), diisi dengan uraian informasi Arsip dari setiap naskah dinas;'],
    ['Kolom (5), diisi dengan tanggal Arsip itu tercipta;'],
    ['Kolom (6), diisi dengan tingkat perkembangan Arsip (asli, copy, dan cetak);'],
    ['Kolom (7), diisi dengan jumlah Arsip dalam satuan naskah dinas;'],
    ['Kolom (8), diisi dengan nomor Filing Cabinet;'],
    ['Kolom (9), diisi dengan nomor laci pada Filing Cabinet;'],
    ['Kolom (10), diisi dengan nomor folder Arsip;'],
    ['Kolom (11), diisi dengan klasifikasi keamanan seperti terbuka, terbatas, dan rahasia.'],
    ['Kolom (12), diisi dengan keterangan spesifik dari jenis Arsip, seperti tekstual, kartografi, audio visual, elektronik,     dan digital.']
  ];
  sheet.getRange(startRow, 3, notes.length, 1)
    .setValues(notes)
    .setFontFamily('Bookman Old Style')
    .setFontSize(11)
    .setFontWeight('normal');
  sheet.getRange(startRow, 3).setFontWeight('normal');
}


