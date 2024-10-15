var scriptSet = PropertiesService.getScriptProperties();

//HUBUNGKAN DENGAN TELEGRAM DAN GOOGLE SHEET
var token = '7512137752:AAFN-sHGnFzSI8xCdbrteQPStLgovImsCp4'; // Isi dengan token bot Telegram 
var sheetID = '1PW1M6FgHx-OOfiEb1rnjzrmo_yq_LbpOvjwE5DAw2Sk'; // Isi dengan SheetID Google Sheet
var sheetName = 'Sheet1'; // Isi dengan nama Sheet
var webAppURL = 'https://script.google.com/macros/s/AKfycbyCgL0v646ccYXwb6Hd4x_NTBIhDA0CHUOUF9P6uS2qJS8OlhMleBuwlTCMo9tVD1GsMA/exec'; // Isi dengan Web URL Google Script setelah deploy

//SETTING DATA APA SAJA YANG AKAN DIINPUT
var dataInput = /\/SITE_ID:\s*(.*)?\n\s*SITE_NAME:\s*(.*)?\n\s*METRO_HOSTNAME:\s*(.*)?\n\s*METRO_IP:\s*(.*)?\n\s*NE1_HOSTNAME:\s*(.*)?\n\s*NE1_IP:\s*(.*)?\n\s*FRAME:\s*(.*)?\n\s*SLOT:\s*(.*)?\n\s*PORT:\s*(.*)?\n\s*ONU:\s*(.*)?\n\s*IP_ONT:\s*(.*)?\n\s*IP_CEK:\s*(.*)?\n\s*STO:\s*(.*)?\n\s*CEK:\s*(.*)?\n\s*NAMA_FEEDER:\s*(.*)?\n\s*CORE_FEEDER:\s*(.*)?\n\s*NAMA_DISTRIBUSI:\s*(.*)?\n\s*CORE_DISTRIBUSI:\s*(.*)?\n\s*ODP:\s*(.*)?\n\s*TL:\s*(.*)?/gmi;

var validasiData = /:\s*(.*)?/g;  // Mengizinkan nilai kosong setelah tanda ":"


//PESAN JIKA FORMAT DATA YANG DIKIRIM SALAH
var errorMessage = "Format Salah!";

function tulis(dataInput) {
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  
  // Cek apakah SSID sudah ada
  Logger.log("Memeriksa keberadaan SSID: " + dataInput[0]);
  var isExist = isDataAvail(dataInput[0]); // Memeriksa apakah SSID sudah ada di sheet
  
  if (isExist) {
    // Jika SSID sudah ada, berikan notifikasi
    return "SSID " + dataInput[0] + " sudah ada dalam data, tidak disimpan lagi!";
  } else {
    // Jika SSID belum ada, simpan data baru
    var lastRow = sheet.getLastRow() + 1; // Mendapatkan baris terakhir + 1 untuk input baru

    // Ambil nomor urut terakhir dari sheet 
    var lastRow = sheet.getLastRow();
    var lastNumber = 0;
      if (lastRow > 1) { // Lebih dari satu karena baris pertama mungkin header
        var lastNumberCell = sheet.getRange(lastRow, 1).getValue(); // Ambil nilai di kolom 1 (Nomor Urut)
      if (lastNumberCell) {
        lastNumber = parseInt(lastNumberCell, 10); // Ambil nomor urut terakhir dan convert ke integer
      }
        }
var newNumber = lastNumber + 1;


    sheet.getRange(lastRow + 1, 1).setValue(newNumber); // Nomor urut disimpan di kolom A
    var simpanData = [newNumber].concat(dataInput); // Gabungkan nomor urut dengan data input
    sheet.getRange(lastRow, 2).setValue(dataInput[0]); // Kolom B untuk SSID
    sheet.getRange(lastRow, 3).setValue(dataInput[1]); // Kolom C untuk SITE NAME
    sheet.getRange(lastRow, 4).setValue(dataInput[2] || ''); // Kolom D untuk METRO HOSTNAME, tetap kosong jika tidak ada input
    sheet.getRange(lastRow, 5).setValue(dataInput[3] || ''); // Kolom E untuk METRO IP
    sheet.getRange(lastRow, 6).setValue(dataInput[4] || ''); // Kolom F untuk NE1 HOSTNAME 
    sheet.getRange(lastRow, 7).setValue(dataInput[5] || ''); // Kolom G untuk NE1 IP
    sheet.getRange(lastRow, 8).setValue(dataInput[6] || ''); // Kolom H untuk FRAME
    sheet.getRange(lastRow, 9).setValue(dataInput[7] || ''); // Kolom I untuk SLOT
    sheet.getRange(lastRow, 10).setValue(dataInput[8] || ''); // Kolom J untuk PORT
    sheet.getRange(lastRow, 11).setValue(dataInput[9] || ''); // Kolom K untuk ONU
    sheet.getRange(lastRow, 12).setValue(dataInput[10] || ''); // Kolom L untuk IP ONT
    sheet.getRange(lastRow, 13).setValue(dataInput[11] || ''); // Kolom M untuk IP CEK
    sheet.getRange(lastRow, 14).setValue(dataInput[12] || ''); // Kolom N untuk STO
    sheet.getRange(lastRow, 15).setValue(dataInput[13] || ''); // Kolom O untuk CEK
    sheet.getRange(lastRow, 16).setValue(dataInput[14] || ''); // Kolom P untuk NAMA FEEDER
    sheet.getRange(lastRow, 17).setValue(dataInput[15] || ''); // Kolom Q untuk CORE FEEDER
    sheet.getRange(lastRow, 18).setValue(dataInput[16] || ''); // Kolom R untuk NAMA DISTRIBUSI
    sheet.getRange(lastRow, 19).setValue(dataInput[17] || ''); // Kolom S untuk CORE DISTRIBUSI
    sheet.getRange(lastRow, 20).setValue(dataInput[18] || ''); // Kolom T untuk ODP
    sheet.getRange(lastRow, 21).setValue(dataInput[19] || ''); // Kolom U untuk TL
    return "Data dengan SSID " + dataInput[0] + " berhasil disimpan!";
  }
}



function breakData(update) {
  var ret = errorMessage;
  var msg = update.message;
  var str = msg.text;

  var match = dataInput.exec(str); // Cocokkan input dengan regex

  // Memastikan ada kecocokan dan memiliki 11 kelompok hasil (SSID + 6 kolom lainnya)
  if (match) {
    var simpan = [];
    
    // Ambil hasil match mulai dari index 1 hingga 11 [berapa kolom di excel]
    for (var i = 1; i <= 21; i++) {
      simpan.push(match[i] ? match[i].trim() : ''); // Jika null atau undefined, gunakan string kosong
    }

    ret = tulis(simpan); // Panggil fungsi tulis untuk menyimpan hasilnya
  } else {
    ret = "Format Salah! Pastikan Anda mengikuti format yang benar.";
  }

  return ret;
}


function escapeHtml(text) {
  var map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;',
  };
  return text.replace(/[&<>"']/g, function(m) {
    return map[m];
  });
}

function doGet(e) {
  return HtmlService.createHtmlOutput("Hey there! send POST request instead!");
}

function doPost(e) {
  if (e.postData.type === "application/json") {
    var update = JSON.parse(e.postData.contents);
    var bot = new Bot(token, update);
    var bus = new CommandBus();
    
    bus.on(/\/help/i, function() {
      this.replyToSender("<b>/finput -> Menampilkan Format</b>\n <b>/cari -> Mencari Data (/cari 222->SSID)</b>\n<b>/fupdate -> Menampilkan Format Update Data</b>");
    });

    bus.on(/\/test/i, function() {
      this.replyToSender("<b>Aman Maseh</b>");
    });

    bus.on(/\/finput/i, function() {
      this.replyToSender("<b>/SITE_ID:</b>\n<b>SITE_NAME:</b>\n<b>METRO_HOSTNAME:</b>\n<b>METRO_IP:</b>\n<b>NE1_HOSTNAME:</b>\n<b>NE1_IP:</b>\n<b>FRAME:</b>\n<b>SLOT:</b>\n<b>PORT:</b>\n<b>ONU:</b>\n<b>IP_ONT:</b>\n<b>IP_CEK:</b>\n<b>STO:</b>\n<b>CEK:</b>\n<b>NAMA_FEEDER:</b>\n<b>CORE_FEEDER:</b>\n<b>NAMA_DISTRIBUSI:</b>\n<b>CORE_DISTRIBUSI:</b>\n<b>ODP:</b>\n<b>TL:</b>\n");
    });

    // Menambahkan command /cari untuk mencari SSID
    bus.on(/\/cari (\S+)/i, function(ssid) { // Memproses command /cari diikuti SSID
      cari(update);
    });

    bus.on(/\/fupdate/i, function() {
      this.replyToSender("/update\n<b>SITE_ID:</b>\n<b>SITE_NAME:</b>\n<b>METRO_HOSTNAME:</b>\n<b>METRO_IP:</b>\n<b>NE1_HOSTNAME:</b>\n<b>NE1_IP:</b>\n<b>FRAME:</b>\n<b>SLOT:</b>\n<b>PORT:</b>\n<b>ONU:</b>\n<b>IP_ONT:</b>\n<b>IP_CEK:</b>\n<b>STO:</b>\n<b>CEK:</b>\n<b>NAMA_FEEDER:</b>\n<b>CORE_FEEDER:</b>\n<b>NAMA_DISTRIBUSI:</b>\n<b>CORE_DISTRIBUSI:</b>\n<b>ODP:</b>\n<b>TL:</b>\n");
    });

// Menambahkan command /update untuk memperbarui data
bus.on(/\/update\s*SITE_ID:\s*(.*)?\n\s*SITE_NAME:\s*(.*)?\n\s*METRO_HOSTNAME:\s*(.*)?\n\s*METRO_IP:\s*(.*)?\n\s*NE1_HOSTNAME:\s*(.*)?\n\s*NE1_IP:\s*(.*)?\n\s*FRAME:\s*(.*)?\n\s*SLOT:\s*(.*)?\n\s*PORT:\s*(.*)?\n\s*ONU:\s*(.*)?\n\s*IP_ONT:\s*(.*)?\n\s*IP_CEK:\s*(.*)?\n\s*STO:\s*(.*)?\n\s*CEK:\s*(.*)?\n\s*NAMA_FEEDER:\s*(.*)?\n\s*CORE_FEEDER:\s*(.*)?\n\s*NAMA_DISTRIBUSI:\s*(.*)?\n\s*CORE_DISTRIBUSI:\s*(.*)?\n\s*ODP:\s*(.*)?\n\s*TL:\s*(.*)?/gmi, function(
  site_id,
  site_name,
  metro_hostname,
  metro_ip,
  ne1_hostname,
  ne1_ip,
  frame,
  slot,
  port,
  onu,
  ip_ont,
  ip_cek,
  sto,
  cek,
  nama_feeder,
  core_feeder,
  nama_distribusi,
  core_distribusi,
  odp,
  tl
) {

  var dataLama = ambilData(site_id); // Ambil data lama berdasarkan SSID
  if (dataLama === "SSID tidak ditemukan!") {
    this.replyToSender(dataLama);
    return;
  }

  var konfirmasi = "Data Lama:\n" + dataLama + "\n\nData Baru:\n" +
    "<b>SITE_ID:</b> " + site_id + "\n" +
    "<b>SITE_NAME:</b> " + site_name + "\n" +
    "<b>METRO_HOSTNAME:</b> " + metro_hostname + "\n" +
    "<b>METRO_IP:</b> " + metro_ip + "\n" +
    "<b>NE1_HOSTNAME:</b> " + ne1_hostname + "\n" +
    "<b>NE1_IP:</b> " + ne1_ip + "\n" +
    "<b>FRAME:</b> " + frame + "\n" +
    "<b>SLOT:</b> " + slot + "\n" +
    "<b>PORT:</b> " + port + "\n" +
    "<b>ONU:</b> " + onu + "\n" +
    "<b>IP_ONT:</b> " + ip_ont + "\n" +
    "<b>IP_CEK:</b> " + ip_cek + "\n" +
    "<b>STO:</b> " + sto + "\n" +
    "<b>CEK:</b> " + cek + "\n" +
    "<b>NAMA_FEEDER:</b> " + nama_feeder + "\n" +
    "<b>CORE_FEEDER:</b> " + core_feeder + "\n" +
    "<b>NAMA_DISTRIBUSI:</b> " + nama_distribusi + "\n" +
    "<b>CORE_DISTRIBUSI:</b> " + core_distribusi + "\n" +
    "<b>ODP:</b> " + odp + "\n" +
    "<b>TL:</b> " + tl + "\n" +
    "Apakah Anda ingin menyimpan data yang ingin diupdate? /y untuk Ya, /t untuk Tidak.";

  this.replyToSender(konfirmasi);

  // Simpan data sementara untuk proses update
  scriptSet.setProperty('update_site_id', site_id);
  scriptSet.setProperty('update_site_name', site_name);
  scriptSet.setProperty('update_metro_hostname', metro_hostname);
  scriptSet.setProperty('update_metro_ip', metro_ip);
  scriptSet.setProperty('update_ne1_hostname', ne1_hostname);
  scriptSet.setProperty('update_ne1_ip', ne1_ip);
  scriptSet.setProperty('update_frame', frame);
  scriptSet.setProperty('update_slot', slot);
  scriptSet.setProperty('update_port', port);
  scriptSet.setProperty('update_onu', onu);
  scriptSet.setProperty('update_ip_ont', ip_ont);
  scriptSet.setProperty('update_ip_cek', ip_cek);
  scriptSet.setProperty('update_sto', sto);
  scriptSet.setProperty('update_cek', cek);
  scriptSet.setProperty('update_nama_feeder', nama_feeder);
  scriptSet.setProperty('update_core_feeder', core_feeder);
  scriptSet.setProperty('update_nama_distribusi', nama_distribusi);
  scriptSet.setProperty('update_core_distribusi', core_distribusi);
  scriptSet.setProperty('update_odp', odp);
  scriptSet.setProperty('update_tl', tl);

});

// Menambahkan handler untuk konfirmasi update
bus.on(/\/y/i, function() {
  var site_id = scriptSet.getProperty('update_site_id');
  var site_name = scriptSet.getProperty('update_site_name');
  var metro_hostname = scriptSet.getProperty('update_metro_hostname');
  var metro_ip = scriptSet.getProperty('update_metro_ip');
  var ne1_hostname = scriptSet.getProperty('update_ne1_hostname');
  var ne1_ip = scriptSet.getProperty('update_ne1_ip');
  var frame = scriptSet.getProperty('update_frame');
  var slot = scriptSet.getProperty('update_slot');
  var port = scriptSet.getProperty('update_port');
  var onu = scriptSet.getProperty('update_onu');
  var ip_ont = scriptSet.getProperty('update_ip_ont');
  var ip_cek = scriptSet.getProperty('update_ip_cek');
  var sto = scriptSet.getProperty('update_sto');
  var cek = scriptSet.getProperty('update_cek');
  var nama_feeder = scriptSet.getProperty('update_nama_feeder');
  var core_feeder = scriptSet.getProperty('update_core_feeder');
  var nama_distribusi = scriptSet.getProperty('update_nama_distribusi');
  var core_distribusi = scriptSet.getProperty('update_core_distribusi');
  var odp = scriptSet.getProperty('update_odp');
  var tl = scriptSet.getProperty('update_tl');

  
  if (site_id) {
    // Panggil fungsi update
    var result = updateData(site_id, site_name, metro_hostname, metro_ip, ne1_hostname, ne1_ip, frame,slot, port, onu, ip_ont, ip_cek, sto, cek, nama_feeder, core_feeder, nama_distribusi, core_distribusi, odp, tl);
    this.replyToSender(result);
    
    // Hapus data sementara setelah proses selesai
  scriptSet.deleteProperty('update_site_id');
  scriptSet.deleteProperty('update_site_name');
  scriptSet.deleteProperty('update_metro_hostname');
  scriptSet.deleteProperty('update_metro_ip');
  scriptSet.deleteProperty('update_ne1_hostname');
  scriptSet.deleteProperty('update_ne1_ip');
  scriptSet.deleteProperty('update_frame');
  scriptSet.deleteProperty('update_slot');
  scriptSet.deleteProperty('update_port');
  scriptSet.deleteProperty('update_onu');
  scriptSet.deleteProperty('update_ip_ont');
  scriptSet.deleteProperty('update_ip_cek');
  scriptSet.deleteProperty('update_sto');
  scriptSet.deleteProperty('update_cek');
  scriptSet.deleteProperty('update_nama_feeder');
  scriptSet.deleteProperty('update_core_feeder');
  scriptSet.deleteProperty('update_nama_distribusi');
  scriptSet.deleteProperty('update_core_distribusi');
  scriptSet.deleteProperty('update_odp');
  scriptSet.deleteProperty('update_tl');
  } else {
    this.replyToSender("Tidak ada data yang dapat diperbarui.");
  }
});

bus.on(/\/t/i, function() {
  this.replyToSender("Pembaruan dibatalkan.");
});

    bus.on(validasiData, function() {
      var rtext = breakData(update);
      this.replyToSender(rtext);
    });

    bot.register(bus);

    if (update) {
      bot.proses();
    }
  }
}

function setWebHook() {
  var bot = new Bot(token, {});
  var result = bot.request('setWebHook', {
    url: webAppURL
  });
  Logger.log(ScriptApp.getService().getUrl());
  Logger.log(result);
}

function Bot(token, update) {
  this.token = token;
  this.update = update;
  this.handlers = [];
}

Bot.prototype.register = function(handler) {
  this.handlers.push(handler);
}

Bot.prototype.proses = function() {
  for (var i in this.handlers) {
    var event = this.handlers[i];
    var result = event.condition(this);
    if (result) {
      return event.handle(this);
    }
  }
}

Bot.prototype.request = function(method, data) {
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(data)
  };

  var response = UrlFetchApp.fetch('https://api.telegram.org/bot' + this.token + '/' + method, options);

  if (response.getResponseCode() === 200) {
    return JSON.parse(response.getContentText());
  }
  return false;
}

Bot.prototype.replyToSender = function(text) {
  return this.request('sendMessage', {
    'chat_id': this.update.message.chat.id,
    'parse_mode': 'HTML',
    'text': text,
    'reply_to_message_id': this.update.message.message_id // Menambahkan reply ke pesan asli
  });
};


function CommandBus() {
  this.command = [];
}

CommandBus.prototype.on = function(regexp, callback) {
  this.command.push({ 'regexp': regexp, 'callback': callback });
}

CommandBus.prototype.condition = function(bot) {
  return bot.update.message.text.charAt(0) === '/';
}

CommandBus.prototype.handle = function(bot) {
  for (var i in this.command) {
    var cmd = this.command[i];
    var tokens = cmd.regexp.exec(bot.update.message.text);
    if (tokens != null) {
      return cmd.callback.apply(bot, tokens.slice(1));
    }
  }
  return bot.replyToSender(errorMessage);
}

function cari(update) {
  var msg = update.message;
  var text = msg.text;
  var match = text.match(/\/cari (\S+)/); // Regex untuk mencari SSID atau nama_feeder yang diinput setelah perintah /cari

  if (match && match.length > 1) {
    var id = match[1]; // SSID atau nama_feeder yang dicari dari pesan
    if (isDataAvail(id)) {
      var dataGabungan = ambilDataCari(id);
      var bot = new Bot(token, update); // Buat objek bot baru di sini
      bot.replyToSender(dataGabungan);
    } else {
      var bot = new Bot(token, update); // Buat objek bot baru di sini
      bot.replyToSender("SSID atau Nama Feeder tidak ditemukan!");
    }
  } else {
    var bot = new Bot(token, update); // Buat objek bot baru di sini
    bot.replyToSender("Format Salah! Gunakan /cari <SSID atau Nama Feeder>.");
  }
}


function ambilData(id) {
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  var dataRange = sheet.getRange("B2:U" + sheet.getLastRow()); // Ambil data dari kolom b sampai u
  var rows = dataRange.getValues();
  var data = [];

  for (var row = 0; row < rows.length; row++) {
    if (rows[row][0] == id) {  // Kolom B dianggap menyimpan SSID
      var info = 
        '<b>SITE_ID:</b> ' + rows[row][0] + '\n' +  // Kolom B: SSID
        '<b>SITE_NAME:</b> ' + rows[row][1] + '\n' + // Kolom C: feeder_name
        '<b>METRO_HOSTNAME:</b> ' + rows[row][2] + '\n' + // Kolom D: feeder_core
        '<b>METRO_IP:</b> ' + rows[row][3] + '\n' + // Kolom E: feeder_capacity
        '<b>NE1_HOSTNAME:</b> ' + rows[row][4] + '\n' + // Kolom F: sumber_data
        '<b>NE1_IP:</b> ' + rows[row][5] + '\n' + // Kolom G: distribusi_cable
        '<b>FRAME:</b> ' + rows[row][6] + '\n'+   // Kolom H: distribusi_core
        '<b>SLOT:</b> ' + rows[row][7] + '\n' +  // Kolom I: SSID
        '<b>PORT:</b> ' + rows[row][8] + '\n' + // Kolom J: feeder_name
        '<b>ONU:</b> ' + rows[row][9] + '\n' + // Kolom K: feeder_core
        '<b>IP_ONT:</b> ' + rows[row][10] + '\n' + // Kolom L: feeder_capacity
        '<b>IP_CEK:</b> ' + rows[row][11] + '\n' + // Kolom M: sumber_data
        '<b>STO:</b> ' + rows[row][12] + '\n' + // Kolom N: distribusi_cable
        '<b>CEK:</b> ' + rows[row][13] + '\n'+   // Kolom O: distribusi_core
        '<b>NAMA_FEEDER:</b> ' + rows[row][14] + '\n' + // Kolom P: feeder_name
        '<b>CORE_FEEDER:</b> ' + rows[row][15] + '\n' + // Kolom Q: feeder_core
        '<b>NAMA_DISTRIBUSI:</b> ' + rows[row][16] + '\n' + // Kolom R: feeder_capacity
        '<b>CORE_DISTRIBUSI:</b> ' + rows[row][17] + '\n' + // Kolom S: sumber_data
        '<b>ODP:</b> ' + rows[row][18] + '\n' + // Kolom T: distribusi_cable
        '<b>TL:</b> ' + rows[row][19] + '\n';   // Kolom U: distribusi_core
      data.push(info);
    }
  }

  if (data.length === 0) {
    return "SSID tidak ditemukan!";
  }

  var dataGabungan = data.join('\n');
  return dataGabungan;
}
 // ===============================================================================================
function ambilDataCari(id) {
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  var dataRange = sheet.getRange("B2:U" + sheet.getLastRow()); // Ambil data dari kolom B sampai U
  var rows = dataRange.getValues();
  var data = [];

  for (var row = 0; row < rows.length; row++) {
    if (rows[row][0] == id || rows[row][14] == id) {  // Kolom B untuk SSID dan Kolom P untuk nama_feeder
      var info = 
        '<b>SITE_ID:</b> ' + rows[row][0] + '\n' +  // Kolom B: SSID
        '<b>SITE_NAME:</b> ' + rows[row][1] + '\n' + // Kolom C: feeder_name
        '<b>NAMA_FEEDER:</b> ' + rows[row][14] + '\n' + // Kolom P: nama_feeder
        '<b>CORE_FEEDER:</b> ' + rows[row][15] + '\n' + // Kolom Q: feeder_core
        '<b>NAMA_DISTRIBUSI:</b> ' + rows[row][16] + '\n' + // Kolom R: feeder_capacity
        '<b>CORE_DISTRIBUSI:</b> ' + rows[row][17] + '\n'; // Kolom S: core distribusi
      data.push(info);
    }
  }

  if (data.length === 0) {
    return "SSID atau Nama Feeder tidak ditemukan!";
  }

  var dataGabungan = data.join('\n');
  return dataGabungan;
}

function isDataAvail(id) {
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  var dataRangeSSID = sheet.getRange("B2:B" + sheet.getLastRow()); // Ambil hanya kolom B (SSID)
  var dataRangeNamaFeeder = sheet.getRange("P2:P" + sheet.getLastRow()); // Ambil data dari kolom P (nama_feeder)
  var rowsSSID = dataRangeSSID.getValues();
  var rowsNamaFeeder = dataRangeNamaFeeder.getValues();

  for (var row = 0; row < rowsSSID.length; row++) {
    if (rowsSSID[row][0] == id || rowsNamaFeeder[row][0] == id) {  // Kolom B untuk SSID dan kolom P untuk nama_feeder
      return true;
    }
  }
  return false;
}


// Modifikasi fungsi updateData
// Modifikasi fungsi updateData
function updateData(site_id, site_name, metro_hostname, metro_ip, ne1_hostname, ne1_ip, frame, slot, port, onu, ip_ont, ip_cek, sto, cek, nama_feeder, core_feeder, nama_distribusi, core_distribusi, odp, tl) {
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  var dataRange = sheet.getRange("B2:B" + sheet.getLastRow()); // Ambil hanya kolom A (SSID)
  var rows = dataRange.getValues();

  for (var row = 0; row < rows.length; row++) {
    if (rows[row][0] == site_id) {  // Kolom A dianggap menyimpan SSID
      sheet.getRange(row + 2, 3).setValue(site_name); // Update feeder_name di kolom B
      sheet.getRange(row + 2, 4).setValue(metro_hostname); // Update feeder_core di kolom C
      sheet.getRange(row + 2, 5).setValue(metro_ip); // Update feeder_capacity di kolom D
      sheet.getRange(row + 2, 6).setValue(ne1_hostname); // Update sumber_data di kolom E
      sheet.getRange(row + 2, 7).setValue(ne1_ip); // Update distribusi_cable di kolom F
      sheet.getRange(row + 2, 8).setValue(frame); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 9).setValue(slot); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 10).setValue(port); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 11).setValue(onu); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 12).setValue(ip_ont); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 13).setValue(ip_cek); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 14).setValue(sto); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 15).setValue(cek); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 16).setValue(nama_feeder); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 17).setValue(core_feeder); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 18).setValue(nama_distribusi); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 19).setValue(core_distribusi); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 20).setValue(odp); // Update distribusi_core di kolom G
      sheet.getRange(row + 2, 21).setValue(tl); // Update distribusi_core di kolom G
      return "Data dengan SSID " + site_id + " berhasil diupdate!";
    }
  }

  return "SSID tidak ditemukan untuk diupdate!";
}

