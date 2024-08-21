function createTimeDrivenTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    Logger.log("⏱️ Trigger sebelumnya berhasil dihapus.");
  });

  ScriptApp.newTrigger('processLargeData')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log("⏰ Trigger baru dibuat untuk berjalan setiap 10 menit.");
}

function processLargeData() {
  try {
    var excelFileId = '1WiUtAl7ygUH1Yif9v9xXaVrgVcb3m11M'; // ID file Excel yang akan dikonversi
    var folderId = '1wQM6SrnYdT4BwOoVnNqFEmxvpEokOsex'; // ID folder tempat file berada
    Logger.log("📂 Folder ID: " + folderId);
    var folder = DriveApp.getFolderById(folderId);
    var excelFile = DriveApp.getFileById(excelFileId);
    Logger.log("📄 File Excel ID: " + excelFileId);

    var excelBlob = excelFile.getBlob();

    Logger.log("⏳ Mengonversi file Excel ke Google Sheets...");
    var convertedSpreadsheet = Drive.Files.insert({
      title: 'Converted Sheet',
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: folderId }]
    }, excelBlob, {
      convert: true
    });

    var convertedSpreadsheetId = convertedSpreadsheet.id;
    Logger.log("✅ File berhasil dikonversi dengan ID: " + convertedSpreadsheetId);

    var spreadsheet = SpreadsheetApp.openById(convertedSpreadsheetId);
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    Logger.log("📊 Data dari sheet berhasil diambil. Jumlah baris: " + data.length);

    var monthYearMap = {};

    var header = data[0];

    for (var i = 1; i < data.length; i++) { // Mulai dari 1 untuk melewati header
      var postDate = new Date(data[i][0]); // Post Date
      var postMonth = postDate.toLocaleString('default', { month: 'long' });
      var postYear = postDate.getFullYear();
      var fileName = 'Hasil Produksi ' + postMonth + ' ' + postYear;

      // Simpan data berdasarkan bulan dan tahun
      if (!monthYearMap[fileName]) {
        monthYearMap[fileName] = [];
      }
      monthYearMap[fileName].push(data[i]);
    }

    // Proses setiap bulan yang ditemukan
    for (var fileName in monthYearMap) {
      var dataForMonth = monthYearMap[fileName];
      var existingFiles = folder.getFilesByName(fileName);

      if (existingFiles.hasNext()) {
        var existingFile = existingFiles.next();
        var oldFileUrl = existingFile.getUrl();
        Logger.log("📁 File lama ditemukan: " + fileName);
        Logger.log("🔗 Link file lama: " + oldFileUrl);

        var existingSpreadsheet = SpreadsheetApp.openById(existingFile.getId());
        var existingSheet = existingSpreadsheet.getSheets()[0];
        existingSheet.clear(); // Hapus semua data lama
        Logger.log("🗑️ Konten lama berhasil dihapus.");

        // Tambahkan header dan data baru
        var range = existingSheet.getRange(1, 1, dataForMonth.length + 1, header.length);
        range.setValues([header].concat(dataForMonth));
        Logger.log("✅ Data baru berhasil disalin ke file lama.");

        DriveApp.getFileById(convertedSpreadsheetId).setTrashed(true);
        Logger.log("🗑️ File sementara berhasil dihapus.");
      } else {
        var file = DriveApp.getFileById(convertedSpreadsheetId);
        file.setName(fileName);
        var newSpreadsheet = SpreadsheetApp.openById(file.getId());
        var newSheet = newSpreadsheet.getSheets()[0];
        newSheet.clear(); // Hapus data lama

        // Tambahkan header dan data baru
        var range = newSheet.getRange(1, 1, dataForMonth.length + 1, header.length);
        range.setValues([header].concat(dataForMonth));
        Logger.log("✅ File baru berhasil dibuat dengan nama: " + fileName);
        Logger.log("🔗 Link file baru: " + file.getUrl());
      }
    }

  } catch (e) {
    Logger.log("❌ Error: " + e.toString());
  }
}
