// ==========================================
// BACKEND SISTEM KASIR PRO V6.9 (DRIVE INTEGRATED)
// DENGAN FITUR CONSUMER VIEW & ORDER ONLINE
// ==========================================

var S_USERS = 'Akun';
var S_PROD = 'Produk';
var S_TRX = 'Transaksi';
var S_PURCH = 'Pembelian';
var S_OPNAME = 'Opname';
var S_SETUP = 'Setup';
var S_LOG = 'Log_Sistem'; 
var S_SHIFT = 'Shift'; 
var S_BEBAN = 'Beban';
var S_ORDERS = 'Orders_Online'; // Sheet baru untuk pesanan online

// Format Rupiah
function formatRp(angka) {
  return Number(angka || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function doGet(e) {
  // --- ENDPOINT UNTUK PUBLIC (KATALOG PRODUK) ---
  if (e && e.parameter.action == 'getPublicProducts') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProd = ss.getSheetByName(S_PROD);
    if (!sheetProd) return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Sheet produk tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);
    
    var prodData = sheetProd.getDataRange().getValues();
    var products = [];
    
    // Cari index kolom Jual_Online (dinamis)
    var headers = prodData[0];
    var jualOnlineIdx = headers.indexOf('Jual Online'); // atau 'Jual_Online'
    if (jualOnlineIdx === -1) jualOnlineIdx = headers.indexOf('Jual_Online');
    
    // Looping data produk (mulai baris 2 karena baris 1 header)
    for (var i = 1; i < prodData.length; i++) {
      if (!prodData[i][0]) continue;
      // Filter hanya yang Jual_Online bernilai TRUE (case-insensitive)
      var jualOnline = (jualOnlineIdx !== -1 && prodData[i][jualOnlineIdx]) ? String(prodData[i][jualOnlineIdx]).toUpperCase() : 'FALSE';
      if (jualOnline !== 'TRUE') continue;
      
      products.push({
        id: prodData[i][0], 
        nama: String(prodData[i][1]), 
        kategori: String(prodData[i][2]),
        hargaJual: Number(prodData[i][4]) || 0, 
        image: prodData[i].length > 6 ? String(prodData[i][6]) : ''
      });
    }
    
    var output = ContentService.createTextOutput(JSON.stringify({
      success: true, 
      products: products
    }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
  
  // --- ENDPOINT UNTUK KASIR (POLLING PESANAN) ---
  if (e && e.parameter.action == 'getPendingOrders') {
    try {
      var token = e.parameter.token;
      if (!token) throw new Error("Token diperlukan");
      var orders = getPendingOrders(token);
      var output = ContentService.createTextOutput(JSON.stringify(orders));
      output.setMimeType(ContentService.MimeType.JSON);
      return output;
    } catch (err) {
      var output = ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }));
      output.setMimeType(ContentService.MimeType.JSON);
      return output;
    }
  }
  
  // --- ENDPOINT UNTUK UPDATE STATUS PESANAN (KASIR) ---
  if (e && e.parameter.action == 'updateOrderStatus') {
    try {
      var token = e.parameter.token;
      var orderId = e.parameter.orderId;
      var newStatus = e.parameter.newStatus;
      if (!token || !orderId || !newStatus) throw new Error("Parameter tidak lengkap");
      var result = updateOrderStatus(token, orderId, newStatus);
      var output = ContentService.createTextOutput(JSON.stringify(result));
      output.setMimeType(ContentService.MimeType.JSON);
      return output;
    } catch (err) {
      var output = ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }));
      output.setMimeType(ContentService.MimeType.JSON);
      return output;
    }
  }
  
  // --- DEFAULT: TAMPILAN HTML UNTUK KASIR/ADMIN ---
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Kasir Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// DO POST untuk menerima pesanan baru (guest checkout)
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action == 'createOrder') {
      var result = createOrder(data.payload);
      var output = ContentService.createTextOutput(JSON.stringify(result));
      output.setMimeType(ContentService.MimeType.JSON);
      return output;
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Aksi tidak dikenal" })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi Perekam Log
function addSystemLog(user, aksi, keterangan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(S_LOG);
    if(sheet) {
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, 5).setValues([['LOG' + new Date().getTime(), new Date(), user, aksi, keterangan]]);
    }
  } catch(e) {}
}

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [S_USERS, S_PROD, S_PURCH, S_OPNAME, S_SETUP, S_LOG, S_SHIFT, S_BEBAN, S_ORDERS];
  var headers = {
    'Akun': ['Username', 'Password', 'Nama', 'Role', 'Token'],
    'Produk': ['ID', 'Nama', 'Kategori', 'Harga Beli', 'Harga Jual', 'Stok', 'Image URL', 'Jual Online'],
    'Transaksi': ['ID Transaksi', 'Tanggal', 'Item', 'Total Harga', 'Modal', 'Laba', 'Metode Bayar', 'Uang Diterima', 'Kembalian', 'No WA', 'Kasir', 'Status', 'Cart JSON', 'PDF Link'],
    'Pembelian': ['ID', 'Tanggal', 'Kasir', 'IDProduk', 'NamaProduk', 'Qty', 'Harga Beli', 'TotalBiaya', 'Supplier'],
    'Opname': ['ID', 'Tanggal', 'Kasir', 'IDProduk', 'StokSistem', 'StokFisik', 'Keterangan'],
    'Setup': ['Key', 'Value'],
    'Log_Sistem': ['ID Log', 'Waktu', 'User', 'Aksi', 'Keterangan'],
    'Shift': ['ID Shift', 'Kasir', 'Waktu Buka', 'Waktu Tutup', 'Modal Awal', 'Total Tunai', 'Kas Fisik', 'Selisih', 'Status'],
    'Beban': ['ID', 'Tanggal', 'Kasir', 'Kategori', 'Keterangan', 'Nominal'],
    'Orders_Online': ['Order ID', 'Nama Pemesan', 'Nomor WA', 'Waktu Pengambilan', 'Detail Item (JSON)', 'Total Harga', 'Status', 'Tanggal Dibuat', 'Diproses Oleh']
  };

  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
      if (name === S_USERS) sheet.getRange(2, 1, 1, 5).setValues([['admin', 'admin123', 'Administrator', 'Admin', '']]);
      if (name === S_SETUP) {
        sheet.getRange(2, 1, 7, 2).setValues([
          ['Nama Toko', 'Toko POS Pro'],
          ['Alamat', 'Jl. Contoh Alamat No. 123'],
          ['Telepon', '08123456789'],
          ['Metode HPP', 'AVERAGE'],
          ['Fonnte Token', ''],
          ['Midtrans Server Key', ''],
          ['Midtrans Env', 'Sandbox']
        ]);
      }
    } else {
      // Untuk sheet yang sudah ada, pastikan kolom 'Jual Online' ada di Produk
      if (name === S_PROD) {
        var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (existingHeaders.indexOf('Jual Online') === -1 && existingHeaders.indexOf('Jual_Online') === -1) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Jual Online');
          // Isi default FALSE untuk semua produk yang sudah ada
          var lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            sheet.getRange(2, sheet.getLastColumn(), lastRow - 1, 1).setValue('FALSE');
          }
        }
      }
      if (name === S_TRX || name === S_LOG || name === S_PROD) {
        // update header jika perlu (tidak merusak data)
        var currentCols = sheet.getLastColumn();
        var targetCols = headers[name].length;
        if (currentCols < targetCols) {
          sheet.getRange(1, currentCols + 1, 1, targetCols - currentCols).setValues([headers[name].slice(currentCols)]);
        }
      }
    }
  });
  return "Database Berhasil Disiapkan! Silakan Refresh/Deploy ulang web app.";
}

function checkAuth(token) {
  if (!token) throw new Error("UNAUTHORIZED");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_USERS);
  if(!sheet) throw new Error("Tab/Sheet 'Akun' tidak ditemukan! Harap jalankan fungsi setupDatabase() di Editor Apps Script.");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === token) {
      return { username: data[i][0], nama: data[i][2], role: data[i][3], rowIndex: i + 1 };
    }
  }
  throw new Error("UNAUTHORIZED");
}

function loginApp(user, pass) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_USERS);
    if(!sheet) return { success: false, message: 'Database belum disiapkan! Buka Editor Apps Script, jalankan fungsi setupDatabase() terlebih dahulu.' };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == user && data[i][1] == pass) {
        var token = Utilities.getUuid();
        sheet.getRange(i + 1, 5).setValue(token);
        addSystemLog(data[i][2], 'LOGIN', 'Berhasil masuk ke sistem');
        return { success: true, token: token };
      }
    }
    return { success: false, message: 'Username atau Password salah!' };
  } catch (e) {
    return { success: false, message: 'Error Server: ' + e.message };
  }
}

function logoutApp(token) {
  try {
    var user = checkAuth(token);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_USERS).getRange(user.rowIndex, 5).setValue('');
    addSystemLog(user.nama, 'LOGOUT', 'Keluar dari sistem');
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

function getInitialData(token) {
  try {
    // 1. Tangani Mode Guest (Token Null / Kosong)
    var isGuest = (!token || token === 'null' || token === '');
    var user = isGuest ? { nama: 'Guest', role: 'Guest' } : checkAuth(token);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2. Load Data Produk (Bisa diakses Guest & Admin)
    var sheetProd = ss.getSheetByName(S_PROD);
    var prodData = sheetProd.getDataRange().getValues();
    
    // PERBAIKAN 1: Bersihkan header dari spasi tersembunyi & seragamkan huruf kecil
    var headersProd = prodData[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var jualOnlineIdx = headersProd.indexOf('jual online');
    if (jualOnlineIdx === -1) jualOnlineIdx = headersProd.indexOf('jual_online');
    
    var products = [];
    for (var i = 1; i < prodData.length; i++) {
      if (!prodData[i][0]) continue; 
      
      // PERBAIKAN 2: Pastikan nilai kolom dibersihkan dari spasi (trim) sebelum dicek
      var isJualOnline = false;
      if (jualOnlineIdx !== -1) {
          var valString = String(prodData[i][jualOnlineIdx]).trim().toUpperCase();
          isJualOnline = (valString === 'TRUE');
      }
      
      // Jika diakses oleh Guest, HANYA kirim produk yang boleh dijual online
      if (isGuest && !isJualOnline) continue;

      products.push({
        id: prodData[i][0], 
        nama: String(prodData[i][1]), 
        kategori: String(prodData[i][2]),
        hargaBeli: isGuest ? 0 : (Number(prodData[i][3]) || 0), 
        hargaJual: Number(prodData[i][4]) || 0, 
        stok: Number(prodData[i][5]) || 0,
        image: prodData[i].length > 6 ? String(prodData[i][6]) : '',
        jualOnline: isJualOnline
      });
    }

    // 3. Load Data Setup
    var sheetSetup = ss.getSheetByName(S_SETUP);
    var setupData = sheetSetup.getDataRange().getValues();
    var setup = {};
    
    for (var l = 1; l < setupData.length; l++) {
      if (!setupData[l][0]) continue;
      var key = setupData[l][0];
      var value = setupData[l][1];
      
      // PERBAIKAN: Jika Google Sheets membaca isinya sebagai Waktu/Tanggal (Date Object), 
      // kita paksa ubah menjadi teks (String) misal "08:00" agar sistem HTML tidak Crash!
      if (value instanceof Date) {
        var hh = ('0' + value.getHours()).slice(-2);
        var mm = ('0' + value.getMinutes()).slice(-2);
        value = hh + ':' + mm;
      }
      
      // Keamanan: Jangan kirim API Key / Token rahasia ke frontend Guest
      if (isGuest && (key.indexOf('Token') > -1 || key.indexOf('Key') > -1)) {
         continue;
      }
      setup[key] = value;
    }

    // ====================================================
    // 4. STOP POINT UNTUK GUEST
    // Jika ini Guest, kirim struktur data kosong untuk tabel rahasia
    // agar frontend tidak error (misal error map() is not a function)
    // ====================================================
    if (isGuest) {
      return { 
        success: true, 
        products: products, 
        transactions: [], 
        purchases: [], 
        setup: setup, 
        logs: [], 
        currentShift: null, 
        beban: [], 
        user: user 
      };
    }

    // ====================================================
    // 5. LOAD DATA KHUSUS KASIR & ADMIN (Token Valid)
    // ====================================================
    var sheetTrx = ss.getSheetByName(S_TRX);
    var sheetPurch = ss.getSheetByName(S_PURCH);
    var sheetLog = ss.getSheetByName(S_LOG);

    var trxData = sheetTrx.getDataRange().getValues();
    var transactions = [];
    var maxTrx = Math.min(trxData.length, 501);
    for (var j = 1; j < maxTrx; j++) {
      if (!trxData[j][0]) continue;
      transactions.push({
        id: trxData[j][0], 
        tanggal: new Date(trxData[j][1]).getTime(), 
        item: String(trxData[j][2]), 
        total: Number(trxData[j][3]), 
        modal: Number(trxData[j][4]), 
        laba: Number(trxData[j][5]), 
        metode: String(trxData[j][6]), 
        bayar: Number(trxData[j][7]), 
        kembali: Number(trxData[j][8]), 
        wa: String(trxData[j][9]), 
        kasir: String(trxData[j][10]), 
        status: String(trxData[j][11]), 
        cartJson: String(trxData[j][12] || '[]'),
        pdfLink: trxData[j].length > 13 ? String(trxData[j][13] || '') : ''
      });
    }

    var purData = sheetPurch.getDataRange().getValues();
    var purchases = [];
    var maxPur = Math.min(purData.length, 201);
    for (var k = 1; k < maxPur; k++) {
      if (!purData[k][0]) continue;
      purchases.push({
        id: purData[k][0], tanggal: new Date(purData[k][1]).getTime(), kasir: String(purData[k][2]),
        idProduk: purData[k][3], namaProduk: String(purData[k][4]), qty: Number(purData[k][5]),
        hargaBeli: Number(purData[k][6]), totalBiaya: Number(purData[k][7]), supplier: String(purData[k][8])
      });
    }

    var logs = [];
    if(sheetLog && user.role === 'Admin') {
       var logData = sheetLog.getDataRange().getValues();
       var maxLog = Math.min(logData.length, 201);
       for(var m = 1; m < maxLog; m++) {
           if(!logData[m][0]) continue;
           logs.push({
               id: logData[m][0], waktu: new Date(logData[m][1]).getTime(),
               user: logData[m][2], aksi: logData[m][3], keterangan: logData[m][4]
           });
       }
    }

    var sheetShift = ss.getSheetByName(S_SHIFT);
    var activeShift = null;
    if(sheetShift) {
        var shiftData = sheetShift.getDataRange().getValues();
        for(var s=1; s<shiftData.length; s++) {
            if(shiftData[s][1] === user.nama && shiftData[s][8] === 'OPEN') {
                activeShift = { id: shiftData[s][0], modalAwal: shiftData[s][4], waktuBuka: new Date(shiftData[s][2]).getTime() };
                break;
            }
        }
    }

    var sheetBeban = ss.getSheetByName(S_BEBAN);
    var bebanList = [];
    if(sheetBeban) {
        var bebData = sheetBeban.getDataRange().getValues();
        for(var n = 1; n < bebData.length; n++) {
            if(!bebData[n][0]) continue;
            bebanList.push({
                id: bebData[n][0], tanggal: new Date(bebData[n][1]).getTime(),
                kasir: bebData[n][2], kategori: bebData[n][3],
                keterangan: bebData[n][4], nominal: Number(bebData[n][5])
            });
        }
    }

    return { 
        success: true, 
        products: products, 
        transactions: transactions, 
        purchases: purchases, 
        setup: setup, 
        logs: logs, 
        currentShift: activeShift, 
        beban: bebanList, 
        user: { nama: user.nama, role: user.role } 
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// KONEKSI API FONNTE & DRIVE UPLOADER
// ==========================================

function uploadImageToDrive(base64Data, fileName) {
  try {
    var folderId = '1jfeBYQykqBZIsYu4LxNwdfPQvrcYYOvH';
    var folder = DriveApp.getFolderById(folderId);
    
    var base64Str = base64Data.split(',')[1];
    var mimeType = base64Data.split(';')[0].split(':')[1];
    
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Str), mimeType, fileName);
    var file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("Sharing diblokir oleh sistem, tapi file berhasil dibuat.");
    }
    
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
  } catch (e) {
    return ''; 
  }
}

function sendFonnteWA(tokenFonnte, targetWA, messageText) {
  if (!tokenFonnte || !targetWA || String(targetWA).length < 8) return "Gagal: Token/No WA Kosong";
  
  var phone = String(targetWA).replace(/\D/g, '');
  if (phone.indexOf('0') === 0) phone = '62' + phone.substring(1);

  var options = {
    'method' : 'post',
    'headers' : { 'Authorization': String(tokenFonnte).trim() },
    'payload' : { 'target': phone, 'message': messageText },
    'muteHttpExceptions': true
  };

  try {
    var respon = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
    return respon.getContentText();
  } catch (e) {
    return "Error Server: " + e.message;
  }
}

// ==========================================
// TRANSAKSI & PRODUK
// ==========================================

function saveTransaction(token, txData, cartPayload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTrx = ss.getSheetByName(S_TRX);
    var sheetProd = ss.getSheetByName(S_PROD);
    var sheetSetup = ss.getSheetByName(S_SETUP);

    var setupRaw = sheetSetup.getDataRange().getValues();
    var setupObj = {};
    for(var s=1; s<setupRaw.length; s++) { if(setupRaw[s][0]) setupObj[setupRaw[s][0]] = setupRaw[s][1]; }

    var statusTransaksi = (txData.metode === 'Kasbon') ? 'HUTANG' : 'SUKSES';

    var rowData = [
      txData.id, new Date(), txData.itemDesc, Number(txData.total) || 0, Number(txData.modal) || 0, Number(txData.laba) || 0, 
      txData.metode, Number(txData.bayar) || 0, Number(txData.kembali) || 0, txData.wa || '', user.nama, 
      statusTransaksi, JSON.stringify(cartPayload), ''
    ];

    sheetTrx.insertRowAfter(1);
    sheetTrx.getRange(2, 1, 1, rowData.length).setValues([rowData]);

    var prodData = sheetProd.getDataRange().getValues();
    for (var i = 0; i < cartPayload.length; i++) {
      var item = cartPayload[i];
      for (var j = 1; j < prodData.length; j++) {
        if (String(prodData[j][0]).trim() === String(item.id).trim()) {
          var currentStock = Number(prodData[j][5]) || 0;
          sheetProd.getRange(j + 1, 6).setValue(currentStock - Number(item.qty));
          break;
        }
      }
    }
    SpreadsheetApp.flush(); 

    addSystemLog(user.nama, 'TRANSAKSI', 'Transaksi baru ' + txData.id + ' senilai Rp ' + formatRp(txData.total) + ' (' + statusTransaksi + ')');

    // TRIGGER WHATSAPP
    var statusWA = "WA Tidak Dikirim";
    if (txData.wa && setupObj['Fonnte Token']) {
      try {
          var tokoName = setupObj['Nama Toko'] || 'Toko Kami';
          var waMsg = "*STRUK PEMBELIAN - " + tokoName + "*\n\n";
          waMsg += "ID: " + txData.id + "\n";
          waMsg += "Tanggal: " + Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm") + "\n";
          waMsg += "Pelanggan: " + (txData.pelanggan || '-') + "\n";
          waMsg += "----------------------------\n";
          
          cartPayload.forEach(function(cartItem) {
            waMsg += cartItem.nama + "\n";
            waMsg += cartItem.qty + " x Rp " + formatRp(cartItem.hargaJual) + " = Rp " + formatRp(cartItem.qty * cartItem.hargaJual) + "\n";
          });
          
          waMsg += "----------------------------\n";
          if(Number(txData.diskon) > 0) waMsg += "Diskon: Rp " + formatRp(txData.diskon) + "\n";
          waMsg += "*TOTAL: Rp " + formatRp(txData.total) + "*\n";
          waMsg += "Bayar (" + txData.metode + "): Rp " + formatRp(txData.bayar) + "\n";
          if(statusTransaksi === 'HUTANG') waMsg += "\n*STATUS: BELUM LUNAS (KASBON)*\n";
          waMsg += "\nTerima Kasih Atas Kunjungan Anda!";

          statusWA = sendFonnteWA(setupObj['Fonnte Token'], txData.wa, waMsg);
      } catch(waErr) { statusWA = "Error WA: " + waErr.message; }
    }

    return { success: true, pdfLink: '', wa_respon: statusWA };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function processReturn(token, txId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    if (user.role !== 'Admin') throw new Error("Akses ditolak");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTrx = ss.getSheetByName(S_TRX);
    var sheetProd = ss.getSheetByName(S_PROD);
    var trxData = sheetTrx.getDataRange().getValues();
    var found = false;
    var cartJson = '[]';
    for (var i = 1; i < trxData.length; i++) {
      if (String(trxData[i][0]).trim() === String(txId).trim() && String(trxData[i][11]).trim() !== 'RETUR') {
        sheetTrx.getRange(i + 1, 12).setValue('RETUR');
        cartJson = trxData[i][12];
        found = true;
        break;
      }
    }
    if (!found) throw new Error("Transaksi tidak ditemukan atau sudah diretur");
    var cartPayload = [];
    try { cartPayload = JSON.parse(cartJson); } catch (e) {}
    var prodData = sheetProd.getDataRange().getValues();
    for (var k = 0; k < cartPayload.length; k++) {
      var item = cartPayload[k];
      for (var j = 1; j < prodData.length; j++) {
        if (String(prodData[j][0]).trim() === String(item.id).trim()) {
          var currentStock = Number(prodData[j][5]) || 0;
          sheetProd.getRange(j + 1, 6).setValue(currentStock + Number(item.qty));
          break;
        }
      }
    }
    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'RETUR', 'Membatalkan/Retur transaksi ID ' + txId);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveProduct(token, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    if (user.role !== 'Admin') throw new Error("Akses ditolak");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_PROD);
    
    // Cari index kolom Jual Online secara akurat (ubah ke huruf kecil agar tidak sensitif kapital)
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var cleanHeaders = headers.map(function(h) { return String(h).trim().toLowerCase(); });
    
    var jualOnlineIdx = cleanHeaders.indexOf('jual online');
    if (jualOnlineIdx === -1) jualOnlineIdx = cleanHeaders.indexOf('jual_online');
    
    // Proses Gambar
    var imageUrl = '';
    if (data.image && data.image.indexOf('data:image') === 0) {
       var safeFileName = (data.id || data.nama).replace(/[^a-zA-Z0-9]/g, '_');
       imageUrl = uploadImageToDrive(data.image, 'PROD_' + safeFileName + '.jpg');
    } else if (data.image) {
       imageUrl = data.image;
    }
    
    // Pastikan Nilainya Tipe Data Boolean Asli (Bukan Teks String)
    var jualOnlineVal = (data.jualOnline === true || data.jualOnline === 'true') ? true : false;

    if (data.isNew) {
      var newId = data.id && data.id.trim() !== '' ? data.id : ('BRG' + new Date().getTime());
      
      // Pengecekan ID ganda
      var prodDataCheck = sheet.getDataRange().getValues();
      for(var x = 1; x < prodDataCheck.length; x++) {
          if(String(prodDataCheck[x][0]).trim() === String(newId).trim()) throw new Error("ID / Barcode tersebut sudah digunakan!");
      }

      // Siapkan baris kosong sesuai jumlah kolom header untuk menghindari eror dimensi
      var rowData = new Array(headers.length);
      for(var c = 0; c < rowData.length; c++) rowData[c] = ""; 
      
      // Isi data wajib (Kolom A s/d G)
      rowData[0] = newId;
      rowData[1] = data.nama;
      rowData[2] = data.kategori;
      rowData[3] = data.hargaBeli;
      rowData[4] = data.hargaJual;
      rowData[5] = data.stok;
      rowData[6] = imageUrl;
      
      // Isi data Jual Online berdasarkan indeks yang ditemukan (Kolom H dsb)
      if (jualOnlineIdx !== -1) {
          rowData[jualOnlineIdx] = jualOnlineVal;
      } else {
          rowData[7] = jualOnlineVal; // Fallback ke kolom H jika header tidak ketemu
      }

      // Sisipkan di baris kedua (teratas setelah header)
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
      
      SpreadsheetApp.flush();
      addSystemLog(user.nama, 'PRODUK', 'Menambahkan barang baru: ' + data.nama);
      return { success: true, id: newId, imageUrl: imageUrl };
      
    } else {
      
      // EDIT BARANG LAMA
      var prodData = sheet.getDataRange().getValues();
      for (var i = 1; i < prodData.length; i++) {
        if (String(prodData[i][0]).trim() === String(data.id).trim()) {
          
          // Update data reguler sekaligus (Nama s/d ImageUrl -> Kolom B sampai G)
          // Menggunakan 1 kali setValues jauh lebih ringan & cepat dari pada loop setValue
          var updateRow = [data.nama, data.kategori, data.hargaBeli, data.hargaJual, data.stok, imageUrl];
          sheet.getRange(i + 1, 2, 1, 6).setValues([updateRow]);
          
          // Update status Jual Online secara terpisah agar aman jika posisinya tergeser
          if (jualOnlineIdx !== -1) {
              sheet.getRange(i + 1, jualOnlineIdx + 1).setValue(jualOnlineVal);
          } else {
              sheet.getRange(i + 1, 8).setValue(jualOnlineVal); // Fallback kolom H
          }

          SpreadsheetApp.flush();
          addSystemLog(user.nama, 'PRODUK', 'Mengedit barang: ' + data.nama);
          return { success: true, imageUrl: imageUrl };
        }
      }
      throw new Error("Produk tidak ditemukan");
    }
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveBulkProducts(token, bulkData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    if (user.role !== 'Admin') throw new Error("Akses ditolak");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_PROD);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var jualOnlineIdx = headers.indexOf('Jual Online');
    if (jualOnlineIdx === -1) jualOnlineIdx = headers.indexOf('Jual_Online');
    var hasJualOnline = (jualOnlineIdx !== -1);
    
    var existingData = sheet.getDataRange().getValues();
    var existingMap = {};
    for (var i = 1; i < existingData.length; i++) {
      existingMap[String(existingData[i][0]).trim()] = i + 1;
    }

    var newRowsToInsert = [];

    for (var j = 0; j < bulkData.length; j++) {
      var item = bulkData[j];
      var id = String(item.id).trim();
      if (!id) id = 'BRG' + new Date().getTime() + j;
      
      var jualOnlineVal = (item.jualOnline === true || item.jualOnline === 'true') ? 'TRUE' : 'FALSE';

      if (existingMap[id]) {
        var rowIdx = existingMap[id];
        var currentImageUrl = existingData[rowIdx-1][6] || '';
        // Update kolom 2-7 dan jika ada kolom Jual Online
        sheet.getRange(rowIdx, 2, 1, 6).setValues([[item.nama, item.kategori, item.hargaBeli, item.hargaJual, item.stok, currentImageUrl]]);
        if (hasJualOnline) {
          sheet.getRange(rowIdx, jualOnlineIdx + 1).setValue(jualOnlineVal);
        }
      } else {
        var newRow = [id, item.nama, item.kategori, item.hargaBeli, item.hargaJual, item.stok, ''];
        if (hasJualOnline) newRow.push(jualOnlineVal);
        newRowsToInsert.push(newRow);
      }
    }

    if (newRowsToInsert.length > 0) {
      sheet.insertRowsAfter(1, newRowsToInsert.length);
      var numCols = hasJualOnline ? 8 : 7;
      sheet.getRange(2, 1, newRowsToInsert.length, numCols).setValues(newRowsToInsert);
    }

    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'PRODUK', 'Melakukan Upload Masal sebanyak ' + bulkData.length + ' item');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function savePembelian(token, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPurch = ss.getSheetByName(S_PURCH);
    var sheetProd = ss.getSheetByName(S_PROD);
    var sheetSetup = ss.getSheetByName(S_SETUP);

    var metodeHPP = 'AVERAGE';
    var setupData = sheetSetup.getDataRange().getValues();
    for (var s = 1; s < setupData.length; s++) {
      if (setupData[s][0] == 'Metode HPP') { metodeHPP = setupData[s][1]; break; }
    }

    var newId = 'BELI' + new Date().getTime();
    var rowData = [
      newId, new Date(), user.nama, data.idProduk, data.namaProduk,
      data.qty, data.hargaBeli, data.totalBiaya, data.supplier
    ];
    
    sheetPurch.insertRowAfter(1);
    sheetPurch.getRange(2, 1, 1, rowData.length).setValues([rowData]);

    var prodData = sheetProd.getDataRange().getValues();
    for (var i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]).trim() === String(data.idProduk).trim()) {
        var currentStock = Number(prodData[i][5]) || 0;
        var currentPrice = Number(prodData[i][3]) || 0;
        var newPrice = data.hargaBeli; 

        if (metodeHPP === 'AVERAGE') {
          var totalStock = currentStock + data.qty;
          if (totalStock > 0) newPrice = Math.round(((currentStock * currentPrice) + (data.qty * data.hargaBeli)) / totalStock);
        } else if (metodeHPP === 'FIFO') {
          newPrice = data.hargaBeli; 
        }

        sheetProd.getRange(i + 1, 4).setValue(newPrice);
        sheetProd.getRange(i + 1, 6).setValue(currentStock + Number(data.qty));
        break;
      }
    }
    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'RESTOCK', 'Membeli ' + data.namaProduk + ' (Qty: ' + data.qty + ') Total Rp ' + formatRp(data.totalBiaya));
    return { success: true, newId: newId, newHargaBeli: data.hargaBeli };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveOpname(token, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOpname = ss.getSheetByName(S_OPNAME);
    var sheetProd = ss.getSheetByName(S_PROD);

    var newId = 'OPN' + new Date().getTime();
    var rowData = [
      newId, new Date(), user.nama, data.idProduk,
      data.stokSistem, data.stokFisik, data.keterangan
    ];
    
    sheetOpname.insertRowAfter(1);
    sheetOpname.getRange(2, 1, 1, rowData.length).setValues([rowData]);

    var prodData = sheetProd.getDataRange().getValues();
    for (var i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]).trim() === String(data.idProduk).trim()) {
        sheetProd.getRange(i + 1, 6).setValue(Number(data.stokFisik));
        break;
      }
    }
    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'OPNAME', 'Opname produk ' + data.idProduk + ' dari stok ' + data.stokSistem + ' menjadi ' + data.stokFisik);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveSetup(token, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    if (user.role !== 'Admin') throw new Error("Akses ditolak");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_SETUP);
    var setupData = sheet.getDataRange().getValues();

    var keys = Object.keys(data);
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      var value = data[key];
      var found = false;

      for (var i = 1; i < setupData.length; i++) {
        if (setupData[i][0] == key) {
          sheet.getRange(i + 1, 2).setValue(value);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.insertRowAfter(1);
        sheet.getRange(2, 1, 1, 2).setValues([[key, value]]);
      }
    }
    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'PENGATURAN', 'Mengubah pengaturan sistem/toko');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function cancelRestock(token, id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    if (user.role !== 'Admin') throw new Error("Akses ditolak.");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPurch = ss.getSheetByName(S_PURCH);
    var sheetProd = ss.getSheetByName(S_PROD);
    var sheetSetup = ss.getSheetByName(S_SETUP);

    var metodeHPP = 'AVERAGE';
    var setupData = sheetSetup.getDataRange().getValues();
    for (var s = 1; s < setupData.length; s++) {
      if (setupData[s][0] == 'Metode HPP') { metodeHPP = setupData[s][1]; break; }
    }

    var purchData = sheetPurch.getDataRange().getValues();
    var rowIndex = -1, idProduk = '', qtyBatal = 0, hargaBeliBatal = 0;

    for (var i = 1; i < purchData.length; i++) {
      if (String(purchData[i][0]).trim() === String(id).trim()) {
        rowIndex = i + 1; idProduk = purchData[i][3]; qtyBatal = Number(purchData[i][5])||0; hargaBeliBatal = Number(purchData[i][6])||0;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("Data Pembelian tidak ditemukan.");

    sheetPurch.deleteRow(rowIndex);

    var prodData = sheetProd.getDataRange().getValues();
    var newPriceReturn = 0;

    for (var j = 1; j < prodData.length; j++) {
      if (String(prodData[j][0]).trim() === String(idProduk).trim()) {
        var currentStock = Number(prodData[j][5]) || 0;
        var currentPrice = Number(prodData[j][3]) || 0;
        var newStock = currentStock - qtyBatal;
        var newPrice = currentPrice; 

        if (metodeHPP === 'AVERAGE' && newStock > 0) {
          var totalValueSaatIni = currentStock * currentPrice;
          var valueYangDibatalkan = qtyBatal * hargaBeliBatal;
          var newValue = totalValueSaatIni - valueYangDibatalkan;
          if (newValue < 0) newValue = 0;
          newPrice = Math.round(newValue / newStock);
        } 
        sheetProd.getRange(j + 1, 4).setValue(newPrice);
        sheetProd.getRange(j + 1, 6).setValue(newStock);
        newPriceReturn = newPrice;
        break;
      }
    }
    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'BATAL RESTOCK', 'Membatalkan restock ID ' + id);
    return { success: true, newHpp: newPriceReturn };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function lunasiKasbon(token, txId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    if (user.role !== 'Admin') throw new Error("Akses ditolak.");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTx = ss.getSheetByName(S_TRX);
    var sheetSetup = ss.getSheetByName(S_SETUP);
    var txData = sheetTx.getDataRange().getValues();
    
    var rowIndex = -1; var totalBayar = 0;
    var waNum = ''; var tglTrx = '';

    for (var i = 1; i < txData.length; i++) {
      if (String(txData[i][0]).trim() === String(txId).trim()) {
        rowIndex = i + 1; 
        totalBayar = Number(txData[i][3]) || 0; 
        tglTrx = txData[i][1];
        waNum = String(txData[i][9] || '').trim();
        break;
      }
    }

    if (rowIndex === -1) throw new Error("ID Transaksi Kasbon tidak ditemukan.");

    sheetTx.getRange(rowIndex, 8).setValue(totalBayar);
    sheetTx.getRange(rowIndex, 9).setValue(0);
    sheetTx.getRange(rowIndex, 12).setValue('SUKSES');
    SpreadsheetApp.flush(); 

    addSystemLog(user.nama, 'PELUNASAN', 'Pelunasan Kasbon ID ' + txId + ' (Rp ' + formatRp(totalBayar) + ')');

    var statusWA = "Tersimpan. WA tidak dikirim (Nomor/Fonnte kosong)";
    try {
        var setupData = sheetSetup.getDataRange().getValues();
        var setupObj = {};
        for(var k=1; k<setupData.length; k++) { if(setupData[k][0]) setupObj[setupData[k][0]] = setupData[k][1]; }
        
        if (waNum && setupObj['Fonnte Token']) {
            var dateStr = Utilities.formatDate(new Date(tglTrx), "Asia/Jakarta", "dd/MM/yyyy HH:mm");
            var tokoName = setupObj['Nama Toko'] || 'Toko Kami';
            
            var waMsg = "*BUKTI PELUNASAN KASBON - " + tokoName + "*\n\n";
            waMsg += "ID Transaksi: " + txId + "\n";
            waMsg += "Waktu Pembelian: " + dateStr + "\n";
            waMsg += "Waktu Pelunasan: " + Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm") + "\n";
            waMsg += "----------------------------\n";
            waMsg += "*TOTAL LUNAS: Rp " + formatRp(totalBayar) + "*\n\n";
            waMsg += "Terima kasih, tagihan hutang Anda telah kami terima dan dinyatakan *LUNAS*.";
            
            var fonnteRes = sendFonnteWA(setupObj['Fonnte Token'], waNum, waMsg);
            statusWA = "WA: " + fonnteRes;
        }
    } catch(err) {
        statusWA = "Tersimpan. Tapi WA Error: " + err.message;
    }

    return { success: true, wa_respon: statusWA };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function openShift(token, modalAwal) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetShift = ss.getSheetByName(S_SHIFT);
    
    var shiftId = 'SHF' + new Date().getTime();
    var rowData = [shiftId, user.nama, new Date(), '', Number(modalAwal), 0, 0, 0, 'OPEN'];
    sheetShift.insertRowAfter(1);
    sheetShift.getRange(2, 1, 1, rowData.length).setValues([rowData]);
    SpreadsheetApp.flush();
    
    addSystemLog(user.nama, 'SHIFT BUKA', 'Buka kasir dengan modal Rp ' + formatRp(modalAwal));
    return { success: true, shift: { id: shiftId, modalAwal: modalAwal, waktuBuka: new Date().getTime() } };
  } catch(e) { return { success: false, error: e.message }; }
  finally { lock.releaseLock(); }
}

function closeShift(token, kasFisik) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetShift = ss.getSheetByName(S_SHIFT);
    var sheetTrx = ss.getSheetByName(S_TRX);
    
    var shiftData = sheetShift.getDataRange().getValues();
    var rowIndex = -1; var activeShift = null;
    for(var s=1; s<shiftData.length; s++) {
        if(shiftData[s][1] === user.nama && shiftData[s][8] === 'OPEN') {
            rowIndex = s + 1;
            activeShift = { id: shiftData[s][0], modalAwal: Number(shiftData[s][4]), waktuBuka: new Date(shiftData[s][2]).getTime() };
            break;
        }
    }
    if(rowIndex === -1) throw new Error("Tidak ada shift aktif.");

    var trxData = sheetTrx.getDataRange().getValues();
    var totalTunai = 0;
    for(var i=1; i<trxData.length; i++) {
        var tWaktu = new Date(trxData[i][1]).getTime();
        if(trxData[i][10] === user.nama && trxData[i][6] === 'Tunai' && trxData[i][11] === 'SUKSES' && tWaktu >= activeShift.waktuBuka) {
            totalTunai += Number(trxData[i][3]); 
        }
    }

    var selisih = Number(kasFisik) - (activeShift.modalAwal + totalTunai);
    
    sheetShift.getRange(rowIndex, 4).setValue(new Date()); 
    sheetShift.getRange(rowIndex, 6).setValue(totalTunai);
    sheetShift.getRange(rowIndex, 7).setValue(Number(kasFisik));
    sheetShift.getRange(rowIndex, 8).setValue(selisih);
    sheetShift.getRange(rowIndex, 9).setValue('CLOSED');
    SpreadsheetApp.flush();

    addSystemLog(user.nama, 'SHIFT TUTUP', 'Tutup kasir. Selisih: Rp ' + formatRp(selisih));
    return { success: true, summary: { totalTunai: totalTunai, selisih: selisih } };
  } catch(e) { return { success: false, error: e.message }; }
  finally { lock.releaseLock(); }
}

function saveBulkPembelian(token, restockArr, supplier) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPurch = ss.getSheetByName(S_PURCH);
    var sheetProd = ss.getSheetByName(S_PROD);
    var sheetSetup = ss.getSheetByName(S_SETUP);

    var metodeHPP = 'AVERAGE';
    var setupData = sheetSetup.getDataRange().getValues();
    for (var s = 1; s < setupData.length; s++) {
      if (setupData[s][0] == 'Metode HPP') { metodeHPP = setupData[s][1]; break; }
    }

    var newPurchases = [];
    var totalBiayaSemua = 0;
    var waktu = new Date();

    var rowsToInsert = [];
    for(var i=0; i<restockArr.length; i++) {
       var item = restockArr[i];
       var newId = 'BELI' + waktu.getTime() + i;
       var totalBiayaItem = item.qty * item.hargaBeli;
       totalBiayaSemua += totalBiayaItem;
       rowsToInsert.push([newId, waktu, user.nama, item.idProduk, item.namaProduk, item.qty, item.hargaBeli, totalBiayaItem, supplier]);
       newPurchases.unshift({ id: newId, idProduk: item.idProduk, namaProduk: item.namaProduk, qty: item.qty, hargaBeli: item.hargaBeli, totalBiaya: totalBiayaItem, supplier: supplier, tanggal: waktu.getTime(), kasir: user.nama });
    }

    if(rowsToInsert.length > 0) {
       sheetPurch.insertRowsAfter(1, rowsToInsert.length);
       sheetPurch.getRange(2, 1, rowsToInsert.length, 9).setValues(rowsToInsert);
    }

    var prodData = sheetProd.getDataRange().getValues();
    for(var j=0; j<restockArr.length; j++) {
       var item = restockArr[j];
       for (var k = 1; k < prodData.length; k++) {
          if (String(prodData[k][0]).trim() === String(item.idProduk).trim()) {
            var currentStock = Number(prodData[k][5]) || 0;
            var currentPrice = Number(prodData[k][3]) || 0;
            var newPrice = item.hargaBeli; 

            if (metodeHPP === 'AVERAGE') {
              var totalStock = currentStock + item.qty;
              if (totalStock > 0) newPrice = Math.round(((currentStock * currentPrice) + (item.qty * item.hargaBeli)) / totalStock);
            } else if (metodeHPP === 'FIFO') {
              newPrice = item.hargaBeli; 
            }

            sheetProd.getRange(k + 1, 4).setValue(newPrice);
            sheetProd.getRange(k + 1, 6).setValue(currentStock + Number(item.qty));
            break;
          }
       }
    }
    
    SpreadsheetApp.flush();
    addSystemLog(user.nama, 'RESTOCK MASAL', 'Restock ' + restockArr.length + ' item barang, Total Biaya Rp ' + formatRp(totalBiayaSemua));
    return { success: true, newPurchases: newPurchases };
  } catch (e) { return { success: false, error: e.message }; }
  finally { lock.releaseLock(); }
}

function saveBeban(token, data) {
  var lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000);
      var user = checkAuth(token);
      if (user.role !== 'Admin') throw new Error("Akses ditolak.");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(S_BEBAN);
      var newId = 'EXP' + new Date().getTime();
      var tgl = new Date();
      
      var rowData = [newId, tgl, user.nama, data.kategori, data.keterangan, data.nominal];
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
      SpreadsheetApp.flush();
      
      addSystemLog(user.nama, 'PENGELUARAN', 'Input beban ' + data.kategori + ': ' + data.keterangan + ' (Rp ' + formatRp(data.nominal) + ')');
      return { success: true, newId: newId, kasir: user.nama, tanggal: tgl.getTime() };
  } catch(e) { return { success: false, error: e.message }; }
  finally { lock.releaseLock(); }
}

// ==========================================
// INTEGRASI MIDTRANS QRIS (PERBAIKAN)
// ==========================================

function generateQRIS(token, payload) {
  try {
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetup = ss.getSheetByName(S_SETUP);
    var setupRaw = sheetSetup.getDataRange().getValues();
    var setupObj = {};
    for (var s = 1; s < setupRaw.length; s++) { 
      if (setupRaw[s][0]) setupObj[setupRaw[s][0]] = setupRaw[s][1]; 
    }

    var serverKey = setupObj['Midtrans Server Key'] || '';
    var env = setupObj['Midtrans Env'] || 'Sandbox';

    if (!serverKey) throw new Error("Midtrans Server Key belum diatur di menu Pengaturan!");

    var baseUrl = env === 'Production' 
      ? 'https://api.midtrans.com/v2/charge' 
      : 'https://api.sandbox.midtrans.com/v2/charge';

    // Order ID unik: dari payload.orderId + timestamp
    var orderIdMidtrans = payload.orderId + "-" + new Date().getTime();
    
    var data = {
      "payment_type": "qris",
      "transaction_details": {
        "order_id": orderIdMidtrans,
        "gross_amount": Math.round(payload.amount)
      },
      "qris": {
        "acquirer": "gopay"
      }
    };

    var options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Basic ' + Utilities.base64Encode(serverKey + ':'),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      'payload': JSON.stringify(data),
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(baseUrl, options);
    var result = JSON.parse(response.getContentText());

    if (result.status_code !== '201' && result.status_code !== '200') {
      throw new Error("Midtrans Error: " + (result.status_message || response.getContentText()));
    }

    var qrUrl = '';
    if (result.actions && result.actions.length > 0) {
      for (var i = 0; i < result.actions.length; i++) {
        if (result.actions[i].name === 'generate-qr-code') {
          qrUrl = result.actions[i].url;
          break;
        }
      }
    }

    if (!qrUrl) throw new Error("QRIS URL tidak ditemukan di respons Midtrans.");

    return { success: true, qrUrl: qrUrl, orderIdMidtrans: orderIdMidtrans };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function checkQRISStatus(token, orderIdMidtrans) {
  try {
    var user = checkAuth(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetup = ss.getSheetByName(S_SETUP);
    var setupRaw = sheetSetup.getDataRange().getValues();
    var setupObj = {};
    for (var s = 1; s < setupRaw.length; s++) { 
      if (setupRaw[s][0]) setupObj[setupRaw[s][0]] = setupRaw[s][1]; 
    }

    var serverKey = setupObj['Midtrans Server Key'] || '';
    var env = setupObj['Midtrans Env'] || 'Sandbox';

    if (!serverKey) throw new Error("Midtrans Server Key kosong!");

    var baseUrl = env === 'Production' 
      ? 'https://api.midtrans.com/v2/' + orderIdMidtrans + '/status' 
      : 'https://api.sandbox.midtrans.com/v2/' + orderIdMidtrans + '/status';

    var options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Basic ' + Utilities.base64Encode(serverKey + ':'),
        'Accept': 'application/json'
      },
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(baseUrl, options);
    var result = JSON.parse(response.getContentText());

    var isPaid = false;
    var status = result.transaction_status || 'unknown';
    if (result.status_code === '200' && (status === 'settlement' || status === 'capture')) {
      isPaid = true;
    }

    return { success: true, status: status, isPaid: isPaid };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// FITUR BARU: CONSUMER VIEW & ORDER ONLINE
// ==========================================

/**
 * Membuat pesanan baru dari consumer (guest checkout)
 * Payload: { namaPemesan, noWa, waktuPengambilan, items, totalHarga }
 * items: array of { id, nama, qty, hargaJual }
 */
function createOrder(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOrders = ss.getSheetByName(S_ORDERS);
    if (!sheetOrders) throw new Error("Sheet Orders_Online tidak ditemukan, jalankan setupDatabase()");
    
    var orderId = "ORD" + new Date().getTime();
    var now = new Date();
    var itemsJson = JSON.stringify(payload.items || []);
    var totalHarga = Number(payload.totalHarga) || 0;
    var status = "Menunggu Pembayaran";
    
    // Simpan ke sheet
    var rowData = [
      orderId,
      payload.namaPemesan || "Guest",
      payload.noWa || "",
      payload.waktuPengambilan || "",
      itemsJson,
      totalHarga,
      status,
      now,
      "" // Diproses Oleh (diisi kasir nanti)
    ];
    sheetOrders.insertRowAfter(1);
    sheetOrders.getRange(2, 1, 1, rowData.length).setValues([rowData]);
    
    addSystemLog("System", "ORDER_ONLINE", "Pesanan baru: " + orderId + " dari " + payload.namaPemesan);
    
    return { success: true, orderId: orderId, message: "Pesanan berhasil dibuat. Silakan lakukan pembayaran di kasir." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Mendapatkan daftar pesanan yang statusnya 'Menunggu Pembayaran' atau 'Diproses' untuk hari ini.
 * Hanya untuk kasir yang sudah login (token required).
 */
function getPendingOrders(token) {
  try {
    var user = checkAuth(token); // validasi kasir/admin
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOrders = ss.getSheetByName(S_ORDERS);
    if (!sheetOrders) return { success: false, error: "Sheet Orders_Online tidak ditemukan" };
    
    var data = sheetOrders.getDataRange().getValues();
    if (data.length < 2) return { success: true, orders: [] };
    
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var todayTimestamp = today.getTime();
    
    var pendingOrders = [];
    for (var i = 1; i < data.length; i++) {
      var orderStatus = String(data[i][6] || "").toLowerCase();
      if (orderStatus !== "menunggu pembayaran" && orderStatus !== "diproses") continue;
      
      var createdDate = data[i][7];
      if (createdDate instanceof Date) {
        var createdTimestamp = createdDate.setHours(0, 0, 0, 0);
        if (createdTimestamp !== todayTimestamp) continue;
      } else {
        // jika format bukan Date, lewati
        continue;
      }
      
      pendingOrders.push({
        orderId: data[i][0],
        customerName: data[i][1],
        waNumber: data[i][2],
        pickupTime: data[i][3],
        items: data[i][4], // JSON string
        totalPrice: Number(data[i][5]),
        status: data[i][6],
        createdAt: data[i][7] ? new Date(data[i][7]).getTime() : null,
        processedBy: data[i][8] || ""
      });
    }
    
    return { success: true, orders: pendingOrders };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Mengupdate status pesanan online (dipanggil oleh kasir setelah memproses pembayaran)
 */
function updateOrderStatus(token, orderId, newStatus) {
  try {
    var user = checkAuth(token);
    if (user.role !== 'Admin' && user.role !== 'Kasir') throw new Error("Akses ditolak, hanya kasir/admin");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOrders = ss.getSheetByName(S_ORDERS);
    if (!sheetOrders) throw new Error("Sheet Orders_Online tidak ditemukan");
    
    var data = sheetOrders.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        sheetOrders.getRange(i + 1, 7).setValue(newStatus); // kolom status
        if (newStatus === "Diproses" || newStatus === "Selesai") {
          sheetOrders.getRange(i + 1, 9).setValue(user.nama); // diproses oleh
        }
        found = true;
        break;
      }
    }
    if (!found) throw new Error("Order ID tidak ditemukan");
    
    addSystemLog(user.nama, "UPDATE_ORDER", "Order " + orderId + " status diubah menjadi " + newStatus);
    return { success: true, message: "Status order berhasil diupdate" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =========================================================================
// FITUR PESANAN ONLINE (CUSTOMER VIEW) - SESUAI KOLOM TERBARU
// =========================================================================

// 1. Fungsi untuk menerima pesanan dari Customer
function submitOnlineOrder(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOnline = ss.getSheetByName('Orders_Online');
    
    if (!sheetOnline) throw new Error("Sheet 'Orders_Online' tidak ditemukan.");

    // Buat Order ID unik
    var orderId = 'ONL-' + Math.floor(Math.random() * 10000) + '-' + new Date().getTime().toString().slice(-4);
    var tanggal = new Date();
    
    // Susunan kolom: A(ID), B(Nama), C(WA), D(Waktu), E(JSON), F(Total), G(Status), H(Tgl), I(Kasir)
    sheetOnline.appendRow([
      orderId,                      // A: Order ID
      payload.nama,                 // B: Nama Pemesan
      payload.wa,                   // C: Nomor WA
      payload.waktuAmbil,           // D: Waktu Pengambilan
      JSON.stringify(payload.cart), // E: Detail Item (JSON)
      payload.total,                // F: Total Harga
      'PENDING',                    // G: Status
      tanggal,                      // H: Tanggal Dibuat
      ''                            // I: Diproses Oleh (Masih kosong karena baru pesan)
    ]);

    return { success: true, orderId: orderId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 2. Fungsi untuk Kasir menarik data pesanan yang baru masuk (Polling)
function getPendingOrders() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOnline = ss.getSheetByName('Orders_Online');
    if (!sheetOnline) return { success: true, data: [] };

    var data = sheetOnline.getDataRange().getValues();
    var pendingOrders = [];

    // Looping dari baris kedua (lewati header)
    for (var i = 1; i < data.length; i++) {
      // Cek Kolom G (Indeks 6) apakah statusnya 'PENDING'
      if (data[i][6] === 'PENDING') { 
        pendingOrders.push({
          orderId: data[i][0],                      // A
          nama: data[i][1],                         // B
          wa: data[i][2],                           // C
          waktuAmbil: data[i][3],                   // D
          cart: JSON.parse(data[i][4] || '[]'),     // E
          total: Number(data[i][5]),                // F
          tanggal: new Date(data[i][7]).getTime()   // H
        });
      }
    }

    // Urutkan dari yang paling lama pesan ke paling baru
    pendingOrders.sort(function(a, b) { return a.tanggal - b.tanggal; });

    return { success: true, data: pendingOrders };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 3. Fungsi untuk Kasir memproses pesanan (Ubah status & potong stok)
function processOnlineOrder(token, orderId) {
  try {
    var user = checkAuth(token); // Validasi token kasir
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOnline = ss.getSheetByName('Orders_Online');
    
    var data = sheetOnline.getDataRange().getValues();
    var rowIndex = -1;
    var orderData = null;

    // Cari baris pesanan
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === orderId && data[i][6] === 'PENDING') {
        rowIndex = i + 1; // +1 karena getRange base 1
        orderData = {
          nama: data[i][1],
          wa: data[i][2],
          total: Number(data[i][5]),
          cart: JSON.parse(data[i][4] || '[]')
        };
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Pesanan tidak ditemukan atau sudah diproses.");

    // 1. Ubah status di sheet Orders_Online
    sheetOnline.getRange(rowIndex, 7).setValue('SELESAI'); // Kolom G: Status
    sheetOnline.getRange(rowIndex, 9).setValue(user.nama); // Kolom I: Diproses Oleh

    // 2. Format cart agar sesuai dengan struktur sheet Transaksi
    var itemDescArr = [];
    var modalTotal = 0;
    var cartPayload = [];
    
    var sheetProd = ss.getSheetByName(S_PROD);
    var prodData = sheetProd.getDataRange().getValues();

    for (var c = 0; c < orderData.cart.length; c++) {
      var item = orderData.cart[c];
      itemDescArr.push(item.nama + '(' + item.qty + ')');
      
      // Cari Harga Beli & Potong Stok
      var hpp = 0;
      for (var p = 1; p < prodData.length; p++) {
        if (prodData[p][0] === item.id) {
          hpp = Number(prodData[p][3]) || 0;
          var stokSekarang = Number(prodData[p][5]) || 0;
          sheetProd.getRange(p + 1, 6).setValue(stokSekarang - item.qty);
          break;
        }
      }
      
      modalTotal += (hpp * item.qty);
      cartPayload.push({
        id: item.id,
        nama: item.nama,
        qty: item.qty,
        hargaJual: item.hargaJual
      });
    }

    // 3. Masukkan ke sheet Transaksi
    var sheetTrx = ss.getSheetByName(S_TRX);
    var idTrx = orderId; // Agar ID tetap 'ONL-xxx' sehingga terbaca sebagai Online
    var laba = orderData.total - modalTotal;
    
    sheetTrx.appendRow([
      idTrx,
      new Date(),
      itemDescArr.join(', '),
      orderData.total,
      modalTotal,
      laba,
      'Tunai',       // Asumsi bayar tunai saat ambil
      orderData.total,
      0,             
      orderData.wa,
      user.nama,     
      'SUKSES',
      JSON.stringify(cartPayload)
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 4. Fungsi untuk Kasir membatalkan pesanan secara manual
function cancelOnlineOrder(token, orderId) {
  try {
    var user = checkAuth(token); // Validasi token
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOnline = ss.getSheetByName('Orders_Online');
    
    var data = sheetOnline.getDataRange().getValues();
    var rowIndex = -1;
    
    // Cari baris pesanan
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === orderId && data[i][6] === 'PENDING') {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Pesanan tidak ditemukan.");

    // Ubah status menjadi BATAL agar hilang dari antrean
    sheetOnline.getRange(rowIndex, 7).setValue('BATAL'); 
    sheetOnline.getRange(rowIndex, 9).setValue(user.nama + ' (Batal)'); // Log siapa yang membatalkan
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 5. Fungsi Otomatisasi (Sapu Bersih Pesanan Kadaluarsa)
// Fungsi ini tidak dipanggil dari HTML, melainkan dipanggil oleh Trigger Google Apps Script
function autoCancelOldOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetOnline = ss.getSheetByName('Orders_Online');
  if (!sheetOnline) return;

  var data = sheetOnline.getDataRange().getValues();
  var now = new Date();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][6] === 'PENDING') {
      var orderDate = new Date(data[i][7]);
      
      // Jika pesanan dibuat di hari yang berbeda dengan hari ini (sudah lewat tengah malam)
      if (orderDate.getDate() !== now.getDate() || orderDate.getMonth() !== now.getMonth()) {
         sheetOnline.getRange(i + 1, 7).setValue('BATAL OTOMATIS');
         sheetOnline.getRange(i + 1, 9).setValue('Sistem (Lewat Hari)');
      }
    }
  }
}
