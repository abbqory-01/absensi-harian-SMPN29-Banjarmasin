/**
 * Aplikasi Absensi Harian SMP Negeri 29 Banjarmasin
 * Backend Script & Database Handler
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Absensi Harian - SMP Negeri 29 Banjarmasin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDbSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan. Pastikan script ini terikat dengan Google Sheets.");
  }
  return ss;
}

function setupDatabaseScript() {
  var ss = getDbSpreadsheet();
  
  // 1. Setup Sheet Master Siswa
  var sheetSiswa = ss.getSheetByName("Master_Siswa");
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet("Master_Siswa");
  } else {
    sheetSiswa.clear();
  }
  
  sheetSiswa.getRange("A1:C1").setValues([["No", "Nama Siswa", "Kelas"]]);
  sheetSiswa.getRange("A1:C1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  
  // Master Data Siswa SMPN 29 Banjarmasin (Lengkap Sesuai Data)
  var rawDataSiswa = [
    // VII A
    ["Adelia Putri", "VII A"], ["Ahmad Humaidi", "VII A"], ["Ahmad Ihsan", "VII A"], ["Ahmad Reza", "VII A"],
    ["Ahmad Zazuli Ramadhan", "VII A"], ["Alika Putri Hairunnisa", "VII A"], ["Anggi Mutiara Ramadhan Syarief", "VII A"],
    ["Billqis Azzahra", "VII A"], ["Faiz Rizkian Fauzi", "VII A"], ["Gilang Prasetyo", "VII A"],
    ["Gusti Sofia Azkiya", "VII A"], ["Hasna Arifah", "VII A"], ["Johan Arifin", "VII A"], ["Julia Salsabella", "VII A"],
    ["Lidia Pertiwi", "VII A"], ["M. Sultan Hawari Haikal", "VII A"], ["Maura Kasafa", "VII A"],
    ["Muhammad Aidil", "VII A"], ["Muhammad Farid", "VII A"], ["Muhammad Hafizi", "VII A"],
    ["Muhammad Irfan", "VII A"], ["Muhammad Ramadhan", "VII A"], ["Muhammad Rifki", "VII A"],
    ["Muhammad Wildan Muharam", "VII A"], ["Muhlis Padillah", "VII A"], ["Nada Fitria", "VII A"],
    ["Nor Azizah Lastari", "VII A"], ["Putri Saira Asykia", "VII A"], ["Rapi Padilah", "VII A"],
    ["Raudah", "VII A"], ["Rizkia Fitriani", "VII A"], ["Siti Tarawiah", "VII A"],
    
    // VII B
    ["Adiba Husna", "VII B"], ["Ahmad", "VII B"], ["Ahmad Risky", "VII B"], ["Akhmad Humaidi", "VII B"],
    ["Alya Anwar", "VII B"], ["Aqila Azzahra", "VII B"], ["Bunga Husnul Khotimah", "VII B"],
    ["Fariz Ahsan Zuhdi", "VII B"], ["Hafizha Adzkia", "VII B"], ["Hairil", "VII B"], ["Huzaifah", "VII B"],
    ["Khalifah", "VII B"], ["M. Rani", "VII B"], ["M. Syaddat Ridho", "VII B"], ["Mahmud", "VII B"],
    ["Mahrisa Azkia", "VII B"], ["Mawaddah", "VII B"], ["Muhammad Andri Wirdana", "VII B"],
    ["Muhamad Airlangga Pratama Putra", "VII B"], ["Muhammad Galip", "VII B"], ["Muhammad Haikal Amin", "VII B"],
    ["Muhammad Ramadhan", "VII B"], ["Muhammad Rifqi", "VII B"], ["Natasya Putri", "VII B"],
    ["Naufal Wijaya", "VII B"], ["Norsabila", "VII B"], ["Raidatul Aisyah", "VII B"], ["Resmawati", "VII B"],
    ["Rizqa An-Nuri", "VII B"], ["Saipul Azhar", "VII B"], ["Salsabila Zahra", "VII B"], ["Taufik Hidayat", "VII B"],
    
    // VII C
    ["Ahmad Abizar", "VII C"], ["Ahmad Ramadhan Al-Karimi", "VII C"], ["Ahmad Zaki", "VII C"], ["Alika", "VII C"],
    ["Amira Faliha", "VII C"], ["Ana Fakhrina", "VII C"], ["Armadani", "VII C"], ["Assyifa Azza Auryn", "VII C"],
    ["Dahlia", "VII C"], ["Fathur Rohman", "VII C"], ["Halika", "VII C"], ["Hamdan", "VII C"], ["Haykal", "VII C"],
    ["Iin Jujro Atiah", "VII C"], ["Laili Munawarah", "VII C"], ["M. Rizqy Aulia", "VII C"],
    ["Mahdani Aji Syaputra", "VII C"], ["Maimunah", "VII C"], ["Mutmainah", "VII C"],
    ["Muhammad Afrizal Iqbal", "VII C"], ["Muhammad Erfan", "VII C"], ["Muhammad Hafizhi", "VII C"],
    ["Muhammad Hanif", "VII C"], ["Muhammad Nazza Alkhalifi", "VII C"], ["Muhammad Raqib", "VII C"],
    ["Muhammad Zen", "VII C"], ["Nazma Amania", "VII C"], ["Nur Liana", "VII C"], ["Nurrahim", "VII C"],
    ["Raisha Sabrina", "VII C"], ["Rima Melati", "VII C"], ["Shafa", "VII C"],
    
    // VIII A
    ["Aditya Ernanando Seputra", "VIII A"], ["Ahmad Muzakki", "VIII A"], ["Ahmad Nabil Afrianor", "VIII A"],
    ["Alisa", "VIII A"], ["Amelia Putri", "VIII A"], ["Amelia Safitri", "VIII A"], ["Asma Salsabila", "VIII A"],
    ["Bunga", "VIII A"], ["Fadillah", "VIII A"], ["Kevin Ahmad Riansyah", "VIII A"], ["Khairannor", "VIII A"],
    ["M. Adriyani", "VIII A"], ["M. Harun", "VIII A"], ["Marwaa Aulia Siswadi", "VIII A"], ["Mina Zaskia", "VIII A"],
    ["Muhammad Aufa", "VIII A"], ["Muhammad Haikal", "VIII A"], ["Muhammad Mujakir", "VIII A"],
    ["Muhammad Rehan", "VIII A"], ["Muhammad Syamil Muhtadi", "VIII A"], ["Nor Apika Putri", "VIII A"],
    ["Nurma", "VIII A"], ["Putri Senja Talia", "VIII A"], ["Rara Dewi", "VIII A"], ["Saira Ramadhani", "VIII A"],
    ["Salsabila", "VIII A"], ["Selvia", "VIII A"], ["Silvia Putri Jasmin Maulani", "VIII A"], ["Syafa'atul Marwah", "VIII A"],
    
    // VIII B
    ["Ahmad Arif Rahman", "VIII B"], ["Ahmad Nabil", "VIII B"], ["Akhmad Aulia", "VIII B"], ["Alya Humairo", "VIII B"],
    ["Arjuna Rifqi Al Pakih", "VIII B"], ["Bimau Lidan", "VIII B"], ["Faizah Humaira", "VIII B"],
    ["Hafisah Syahfitri", "VIII B"], ["M. Azka Firdaus", "VIII B"], ["Madinatussafa", "VIII B"],
    ["Miftahul Ghina", "VIII B"], ["Muh. Wahyu Hidayat Noriski", "VIII B"], ["Muhammad Bagir", "VIII B"],
    ["Muhammad Hanafi", "VIII B"], ["Muhammad Ibnu Zamil", "VIII B"], ["Muhammad Naufal", "VIII B"],
    ["Muhammad Restu", "VIII B"], ["Nadila", "VIII B"], ["Nor Hasyifa", "VIII B"], ["Nur Syafa'ah", "VIII B"],
    ["Putri Amanda", "VIII B"], ["Raesa Aziza", "VIII B"], ["Raudhatul Jannah", "VIII B"],
    ["Rio Abdul Rahman", "VIII B"], ["Rizky Ramadhan", "VIII B"], ["Shelia Ahmadiana", "VIII B"],
    ["Siti Humairah", "VIII B"], ["Siti Yuanna Hanifah", "VIII B"], ["Yulia Sabila", "VIII B"],
    
    // VIII C
    ["Ahmad Aufar Ridhoni", "VIII C"], ["Alia Safitri", "VIII C"], ["Alifurrahman", "VIII C"], ["Ardiansyah", "VIII C"],
    ["Asmah", "VIII C"], ["Fahmi", "VIII C"], ["Ikma Ariyanti", "VIII C"], ["Khumairah", "VIII C"],
    ["Mahmudah", "VIII C"], ["Marsaa Aulia Siswadi", "VIII C"], ["Maysa Amira", "VIII C"],
    ["Muhammad Baihaki", "VIII C"], ["Muhammad Fahri", "VIII C"], ["Muhammad Haidir Ali", "VIII C"],
    ["Muhammad Ibnu Akbari", "VIII C"], ["Muhammad Jaid Jidan", "VIII C"], ["Muhammad Rizki Abdillah", "VIII C"],
    ["Nazwa Anggraini Putri", "VIII C"], ["Nor Aradiba Husna", "VIII C"], ["Nur Latifah Zahra", "VIII C"],
    ["Raisya Khumairo", "VIII C"], ["Rehanisa Humairoh", "VIII C"], ["Rizkia Aditia", "VIII C"],
    ["Rizkya Azzahra", "VIII C"], ["Risty Alya Zahra", "VIII C"], ["Saina Ramadhani", "VIII C"],
    ["Selvina Nayla", "VIII C"], ["Sri Kumala Sari", "VIII C"], ["Tasyama", "VIII C"],
    
    // IX A
    ["Abdul Khair", "IX A"], ["Ahmad Zaini", "IX A"], ["Aisyah", "IX A"], ["Ajmal Syarif", "IX A"],
    ["Ambariah", "IX A"], ["Amelia", "IX A"], ["Aprilia", "IX A"], ["Azkia Zulfa", "IX A"],
    ["Irfan Rosadyi", "IX A"], ["Ivan", "IX A"], ["Lola Amelia", "IX A"], ["Marisa", "IX A"],
    ["Muhammad Ashfiansyah", "IX A"], ["Muhammad Jumri", "IX A"], ["Muhammad Saufi Rahman", "IX A"],
    ["Muhammad Syaipul", "IX A"], ["Nabila Apriani", "IX A"], ["Norjanah", "IX A"],
    ["Putri Hajahrah Mulya", "IX A"], ["Raihanah Padilah", "IX A"], ["Rokayah", "IX A"],
    ["Salma Zakiya", "IX A"], ["Salsabilla Azhara", "IX A"], ["Syawali Munawar", "IX A"],
    ["Wardah Salsabila", "IX A"], ["Zahra", "IX A"],
    
    // IX B
    ["Ahmad Dailami", "IX B"], ["Ahmad Muzakir", "IX B"], ["Aisya Asfia", "IX B"], ["Amanda Azzahra", "IX B"],
    ["Amrah", "IX B"], ["Anisa", "IX B"], ["Cahaya Maberurah", "IX B"], ["Febri Monalisa", "IX B"],
    ["Karmila", "IX B"], ["Kayla Zahra Nurqonitah", "IX B"], ["Khairul Ma'ani", "IX B"], ["Mispuah", "IX B"],
    ["Muhammad Alfiyan", "IX B"], ["Muhammad Maulidan", "IX B"], ["Muhammad Nabil", "IX B"],
    ["Muhammad Rangga Rabiullah", "IX B"], ["Muhammad Syahid", "IX B"], ["Nur Madinah", "IX B"],
    ["Nur Syifa Sabila", "IX B"], ["Raudatul Asiah", "IX B"], ["Risky Anwary", "IX B"],
    ["Safa Marwah", "IX B"], ["Samsir Yudi", "IX B"], ["Siti Rahmadina", "IX B"], ["Winey Gifni", "IX B"],
    
    // IX C
    ["Ahmad Fairuz", "IX C"], ["Ahmad Robbi Ramadani", "IX C"], ["Aliya Armadania", "IX C"],
    ["Ameliya Kasih", "IX C"], ["Dawiyah", "IX C"], ["Ervina Puspita Sari", "IX C"], ["Fitrhia Ramadhana", "IX C"],
    ["Friskilla Putri", "IX C"], ["Ilyas", "IX C"], ["Isnawati", "IX C"], ["M. Adittia", "IX C"],
    ["Muhammad Fikri", "IX C"], ["Muhammad Hunaifi", "IX C"], ["Muhammad Hulwani Faisal", "IX C"],
    ["Muhammad Rijali", "IX C"], ["Muz'zalifah", "IX C"], ["Nazwa Azzahra", "IX C"], ["Norhasifa", "IX C"],
    ["Nur Rosidah", "IX C"], ["Nurmainah Azzahra", "IX C"], ["Risqia Azza", "IX C"], ["Safa", "IX C"],
    ["Safwati", "IX C"], ["Sami Yusuf", "IX C"], ["Vita", "IX C"]
  ];
  
  var formattedSiswa = [];
  for (var i = 0; i < rawDataSiswa.length; i++) {
    formattedSiswa.push([i + 1, rawDataSiswa[i][0], rawDataSiswa[i][1]]);
  }
  
  sheetSiswa.getRange(2, 1, formattedSiswa.length, 3).setValues(formattedSiswa);
  
  var sheetRekap = ss.getSheetByName("Rekap_Absensi");
  if (!sheetRekap) {
    sheetRekap = ss.insertSheet("Rekap_Absensi");
  }
  
  if (sheetRekap.getLastRow() === 0) {
    var headers = ["No", "Waktu Input", "Hari / Tanggal", "Nama Pengawas", "Kelas", "Nama Siswa", "Pilihan Status", "Keterangan"];
    sheetRekap.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheetRekap.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
  }
  
  // Hapus Sheet default jika ada
  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  
  return { status: "success", message: "Database dan Master Data Siswa (258 Siswa) berhasil dikonfigurasi!" };
}

function getMasterData() {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName("Master_Siswa");
  if (!sheet) {
    setupDatabaseScript();
    sheet = ss.getSheetByName("Master_Siswa");
  }
  
  var data = sheet.getDataRange().getValues();
  var siswaList = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      siswaList.push({
        no: data[i][0],
        nama: data[i][1],
        kelas: data[i][2]
      });
    }
  }
  
  var listKelas = ["VII A", "VII B", "VII C", "VIII A", "VIII B", "VIII C", "IX A", "IX B", "IX C"];
  
  return {
    siswa: siswaList,
    kelas: listKelas
  };
}

function saveAttendanceData(payload) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName("Rekap_Absensi");
    if (!sheet) {
      setupDatabaseScript();
      sheet = ss.getSheetByName("Rekap_Absensi");
    }
    
    // Sinkronisasi: Hapus data lama untuk Kelas dan Hari/Tanggal yang sama agar TIDAK DUPLIKAT
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      // Hapus dari bawah ke atas agar index baris tidak bergeser saat dihapus
      for (var r = data.length - 1; r >= 0; r--) {
        var rowHariTanggal = data[r][2];
        if (rowHariTanggal instanceof Date) {
          rowHariTanggal = Utilities.formatDate(rowHariTanggal, "Asia/Makassar", "EEEE, dd MMMM yyyy");
        }
        var rowKelas = data[r][4];
        if (rowKelas === payload.kelas && rowHariTanggal === payload.hariTanggal) {
          sheet.deleteRow(r + 2); // +2 karena indeks 0-based dan baris 1 adalah header
        }
      }
    }
    
    var timestamp = Utilities.formatDate(new Date(), "Asia/Makassar", "dd/MM/yyyy HH:mm:ss");
    var currentLastRow = sheet.getLastRow();
    var rowsToInsert = [];
    var startNo = currentLastRow > 1 ? Number(sheet.getRange(currentLastRow, 1).getValue()) + 1 : 1;
    
    // Payload contains: { hariTanggal, namaPengawas, kelas, mode, records: [{namaSiswa, status, keterangan}] }
    for (var i = 0; i < payload.records.length; i++) {
      var item = payload.records[i];
      var namaDanKelas = item.namaSiswa ? (item.namaSiswa + " (" + payload.kelas + ")") : ("Seluruh Siswa " + payload.kelas);
      
      rowsToInsert.push([
        startNo++,
        timestamp,
        payload.hariTanggal,
        payload.namaPengawas,
        payload.kelas,
        namaDanKelas,
        item.status,
        item.keterangan || "-"
      ]);
    }
    
    if (rowsToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 8).setValues(rowsToInsert);
    }
    
    return { status: "success", count: rowsToInsert.length, message: "Absensi " + payload.kelas + " berhasil disinkronkan!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function deleteAllAttendanceRecords() {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName("Rekap_Absensi");
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return { status: "success", message: "Seluruh data rekapan absensi berhasil dihapus!" };
    }
    return { status: "error", message: "Sheet Rekap_Absensi tidak ditemukan" };
  } catch(err) {
    return { status: "error", message: err.toString() };
  }
}

function getAttendanceHistory() {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName("Rekap_Absensi");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var dateFormatted = data[i][2];
      if (dateFormatted instanceof Date) {
        dateFormatted = Utilities.formatDate(dateFormatted, "Asia/Makassar", "EEEE, dd MMMM yyyy");
      }
      
      result.push({
        rowId: i + 1,
        no: data[i][0],
        waktuInput: data[i][1],
        hariTanggal: dateFormatted,
        namaPengawas: data[i][3],
        kelas: data[i][4],
        namaSiswaKelas: data[i][5],
        status: data[i][6],
        keterangan: data[i][7]
      });
    }
    
    // Return reverse chronological
    return result.reverse();
  } catch (e) {
    return [];
  }
}

function deleteAttendanceRecord(rowId) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName("Rekap_Absensi");
    if (sheet) {
      sheet.deleteRow(rowId);
      return { status: "success", message: "Data absensi berhasil dihapus!" };
    }
    return { status: "error", message: "Sheet tidak ditemukan" };
  } catch(err) {
    return { status: "error", message: err.toString() };
  }
}
