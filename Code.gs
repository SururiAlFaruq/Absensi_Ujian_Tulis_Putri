/**
 * Sistem Absensi Pengawas Ujian Awal Sanah (SANTRI PUTRI) - API Backend untuk GitHub Pages
 * Madrasah Diniyah Ar Rohman Tegalrejo - T.A. 2026/2027
 * Author: Sururi Al Faruq
 */

function doGet(e) {
  initDatabase();
  const output = getInitialData();
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let result = '';

    if (body.action === 'addSantriBaru') {
      result = addSantriBaru(body.payload);
    } else if (body.action === 'saveAbsensiRuang') {
      result = saveAbsensiRuang(body.payload);
    } else {
      result = getInitialData();
    }

    return ContentService.createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Jadwal Ujian Putri (30 Nov - 5 Des 2026)
  const jadwalHarian = [
    ['HARI-1', 'Senin, 30 Nov 2026', 'Sesi 1: 19.45 - 20.30 | Sesi 2: 20.45 - 21.30', 'Lughoh Arobiyah & Khot Pegon', 'Durusul Lughoh & Khot Pegon', 'Matan Jurumiyah & Khulashoh NY 1', 'Mutamimah & Khulashoh NY 2', 'Imrithy & Khulashoh NY 3', 'Alfiyah & Waroqot', 'Alfiyah & Tashilut Turoqot'],
    ['HARI-2', 'Selasa, 1 Des 2026', 'Sesi 1: 19.45 - 20.30 | Sesi 2: 20.45 - 21.30', 'Mutholaah & Sholat', 'Mutholaah & Fasholatan', "Aqidatul Awam & Mabadi' Fiqh", 'Aqidah Islamiyah & Safinatun Naja', 'Tijan Durori & Taqrib', 'Kifayatul Awam & Fathul Qorib', 'Matan Assanusi & Fathul Qorib'],
    ['HARI-3', 'Rabu, 2 Des 2026', 'Sesi 1: 19.45 - 20.30 | Sesi 2: 20.45 - 21.30', 'Mitro Sejati & -', 'Sabrowi & Alala', 'Amtsilah T & -', "Qowaidul I'lal & Akhlak Lil Banat", "Ta'lim Muta'alim & -", 'Balaghoh & -', 'Balaghoh & Ilmu Tafsir'],
    ['HARI-4', "Jum'at, 4 Des 2026", 'Sesi 1: 19.45 - 20.30 | Sesi 2: 20.45 - 21.30', "Al Qur'an & -", 'Fiqh Haid 1 & Hadist', 'Fiqh Haid 2 & Syifau Saqim', 'Fiqh Haid 3 & Arbain Nawawi', "I'anatun Nisa' & Mustholah Hadis Abuya", "I'anatun Nisa' & Riyadus Sholihin", 'Risalatul Mahid & Riyadus Sholihin'],
    ['HARI-5', 'Sabtu, 5 Des 2026', 'Sesi 1: 19.45 - 20.30 | Sesi 2: 20.45 - 21.30', 'Tajwid & -', 'Tajwid & Mahfudzot', 'Hidayatus Sibyan & -', 'Tuhfatul Athfal & -', 'Maqsud & -', 'Faroid & -', 'Mantiq & -']
  ];

  let sheetJadwal = ss.getSheetByName('Jadwal_Ujian');
  if (!sheetJadwal) {
    sheetJadwal = ss.insertSheet('Jadwal_Ujian');
    sheetJadwal.appendRow([
      'ID_Jadwal', 'Hari_Tanggal', 'Waktu',
      "I'dad", '1 Wustho', '2 Wustho', '3 Wustho', '4 Wustho', '1 Ulya', '2 Ulya'
    ]);
    sheetJadwal.getRange('A1:J1').setFontWeight('bold').setBackground('#0B1E3B').setFontColor('#FCD34D');
    sheetJadwal.getRange(2, 1, jadwalHarian.length, 10).setValues(jadwalHarian);
  }

  // 2. Master_Pengawas (8 Ustadzah Pengawas Putri)
  let sheetPengawas = ss.getSheetByName('Master_Pengawas');
  const daftarPengawasDefault = [
    ['Usth Amalia Hilda A'],
    ['Usth Fina Afifatu R'],
    ['Usth Kayla Zhafrina F'],
    ['Usth Siti Nur Fatimah'],
    ['Usth Afifah Wulandari'],
    ['Usth Atsilla Imtiyazul A'],
    ['Usth Cahyaning Ayu A'],
    ['Usth Diva Aulia Puji A']
  ];

  if (!sheetPengawas) {
    sheetPengawas = ss.insertSheet('Master_Pengawas');
    sheetPengawas.appendRow(['Nama_Ustadzah']);
    sheetPengawas.getRange('A1').setFontWeight('bold').setBackground('#0B1E3B').setFontColor('#FCD34D');
    sheetPengawas.getRange(2, 1, daftarPengawasDefault.length, 1).setValues(daftarPengawasDefault);
  }

  // 3. Sheet1 (Data 127 Santri Putri Ruang 1 - 5 jika masih kosong)
  const sheetSantri = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  if (sheetSantri.getLastRow() < 3) {
    const dataPutriPerRuang = {
      'Ruang 1': [
        ['Shofiyatuz Zukhruf I.H', '2 Ulya'], ['Dzakiyah Khiyarotul U', '1 Ulya'], ['Lathifah Hidayatul M', '1 Ulya'],
        ['Lutfyana Nur Azizah', '1 Ulya'], ['Maisaroh Kumala Sari', '4 Wustho'], ['Maulidia Nesti Hartono', '4 Wustho'],
        ['Nadia Tama Junita', '4 Wustho'], ['Naurasya Pranarendi H', '4 Wustho'], ["Maghfirotun Nisa'", '3 Wustho'],
        ['Maharani Dwi Rahmawati', '3 Wustho'], ['Quinn Assyifa Haya Wibowo', '3 Wustho'], ['Salma Yasa Assyfa', '3 Wustho'],
        ['Salsabila Priana Putri', '3 Wustho'], ['Abelia Mecca Hariani', '2 Wustho'], ['Alqusthonu Ashfiya', '2 Wustho'],
        ['Arina Septiya', '2 Wustho'], ['Azkiya Nur Syifa Az-Zahra', '2 Wustho'], ['Aira Salsabila Azzahra', '1 Wustho'],
        ['Aisyah Nur Janah Al Fatonah', '1 Wustho'], ['Anisa Rohayati', '1 Wustho'], ['Asizeh Koilatul Hasanah', '1 Wustho'],
        ['Nadya Ayu Salsabilla', '1 Wustho'], ['Natasya Eka Safitri', '1 Wustho'], ['Zahida Qolbi Nadhifa', '1 Wustho'],
        ['Zahra Haya Maulida', '1 Wustho']
      ],
      'Ruang 2': [
        ['Salma Khoirul Izzah', '2 Ulya'], ['Salwa Khoirul Izzah', '2 Ulya'], ['Mauliana Nur Fadila', '1 Ulya'],
        ['Nadia Nurul Nur Azifah', '1 Ulya'], ['Kaela Silfi Oktafia', '4 Wustho'], ['Kharisma Hasna Putri F', '4 Wustho'],
        ['Kholifatun Anisa N A', '4 Wustho'], ["Lutfiana Luluk Mar'atun", '4 Wustho'], ['Hasna Kismatur Rohmawati', '3 Wustho'],
        ['Intan Permata Sari Jagad F', '3 Wustho'], ['Janitra Navis Famanjha', '3 Wustho'], ['Khansa Hasna Adilla', '3 Wustho'],
        ['Luthfia Nuha Setiana', '3 Wustho'], ['Dennazkia Latiefa Hidaya', '2 Wustho'], ['Febi Figlia Asyifatus Zahra', '2 Wustho'],
        ['Fitri Widyastuti', '2 Wustho'], ['Kholisatun Aldila Husnul K', '2 Wustho'], ['Niken Wulandari', '1 Wustho'],
        ['Nur Fasilah Bahar', '1 Wustho'], ['Nur Padilah', '1 Wustho'], ['Qamra Khairunnisa Haziqah', '1 Wustho'],
        ['Saskia Ramadhani', '1 Wustho'], ['Siti Khomsiatun', '1 Wustho'], ['Afifah Ainul Mardhiah', "I'dad"],
        ['Amara Sheza Hayfa Ramadhani', "I'dad"], ['Nurhidayah', "I'dad"]
      ],
      'Ruang 3': [
        ['Rara Dina Rofiah', '2 Ulya'], ['Safina Ulfiah Zahra', '2 Ulya'], ['Noor Balqisya R', '1 Ulya'],
        ['Noyana Keysha', '1 Ulya'], ['Anggun Nur Khasanah', '4 Wustho'], ['Anti Izzatul', '4 Wustho'],
        ['Astrid Larasati Afanda N H', '4