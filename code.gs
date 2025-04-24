// Proje zaman dilimini kontrol edin
var TIMEZONE = Session.getScriptTimeZone(); // "Europe/Istanbul" olarak ayarlı olduğundan emin olun

// E-tablo ID'si (burayı kendi e-tablo ID'nizle değiştirin)
var SPREADSHEET_ID = "1MBSeN8Hpj7PgiqGjcDQ_XIQIjE17Iuu5-eN87FnjyJA";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Etkinlik Takvimi');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Etkinlik Takvimi')
      .addItem('Sayfaları Oluştur', 'setupSheet')
      .addToUi();
}

function setupSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Kullanıcılar sayfası
  var usersSheet = ss.getSheetByName("Kullanıcılar");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Kullanıcılar");
    usersSheet.appendRow(["Kullanıcı Adı", "Pin Kodu", "Rol"]);
    usersSheet.appendRow(["admin", "123456", "Admin"]); // Örnek admin kullanıcısı
    usersSheet.appendRow(["kullanici1", "123456", "Kullanıcı"]); // Örnek kullanıcı
  }

  // Etkinlikler sayfası
  var eventsSheet = ss.getSheetByName("Etkinlikler");
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet("Etkinlikler");
    eventsSheet.appendRow(["Tarih", "Kullanici", "Etkinlik"]);
  }

   // Loglar sayfası (yeni)
  var logsSheet = ss.getSheetByName("Loglar");
  if (!logsSheet) {
    logsSheet = ss.insertSheet("Loglar");
    logsSheet.appendRow(["Tarih", "Kullanıcı", "Eylem", "Detay"]);
  }

  SpreadsheetApp.getUi().alert('Sayfalar oluşturuldu/kontrol edildi.');
}

function logAction(user, action, detail) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Loglar");
  if (!sheet) {
    // Log sayfası yoksa oluşturmaya çalışma (önce setupSheet çalıştırılmalı)
    Logger.log("Log sayfası bulunamadı!");
    return;
  }
  var now = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([now, user, action, detail]);
}


function getUsers() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Kullanıcılar");
  if (!sheet) {
    return { success: false, message: "Kullanıcılar sayfası bulunamadı!" };
  }
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    // Başlık satırını atla
    if (i === 0) continue;
    users.push({
      username: data[i][0],
      pin: data[i][1],
      role: data[i][2] || "Kullanıcı" // Varsayılan rol
    });
  }
  return { success: true, users: users };
}

function getEvents() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Etkinlikler");
  if (!sheet) {
    return { success: false, message: "Etkinlikler sayfası bulunamadı!" };
  }
  var data = sheet.getDataRange().getValues();
  var events = [];
  for (var i = 1; i < data.length; i++) {
    var storedDate = (typeof data[i][0] === "string")
                      ? data[i][0]
                      : Utilities.formatDate(new Date(data[i][0]), TIMEZONE, "yyyy-MM-dd");
    events.push({
      date: storedDate,
      user: data[i][1],
      event: data[i][2]
    });
  }
  return { success: true, events: events };
}

function addUser(username, pin, role) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Kullanıcılar");
  if (!sheet) {
    return { success: false, message: "Kullanıcılar sayfası bulunamadı!" };
  }
  // Kullanıcı adının benzersizliğini kontrol et
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return { success: false, message: "Bu kullanıcı adı zaten kullanımda." };
    }
  }
  sheet.appendRow([username, pin, role]);
  logAction("Onurboni", "Kullanıcı Ekleme", "Kullanıcı: " + username + ", Rol: " + role);
  return { success: true, message: "Kullanıcı eklendi." };
}

function deleteUser(username) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Kullanıcılar");
  if (!sheet) {
    return { success: false, message: "Kullanıcılar sayfası bulunamadı!" };
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      sheet.deleteRow(i + 1);
      logAction("Onurboni", "Kullanıcı Silme", "Kullanıcı: " + username);
      return { success: true, message: "Kullanıcı silindi." };
    }
  }
  return { success: false, message: "Kullanıcı bulunamadı." };
}

function resetPin(username) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Kullanıcılar");
  if (!sheet) {
    return { success: false, message: "Kullanıcılar sayfası bulunamadı!" };
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      sheet.getRange(i + 1, 2).setValue("123456"); // Pini sıfırla
      logAction("Onurboni", "Pin Sıfırlama", "Kullanıcı: " + username);
      return { success: true, message: "Pin sıfırlandı." };
    }
  }
  return { success: false, message: "Kullanıcı bulunamadı." };
}

function changePin(username, newPin) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Kullanıcılar");
  if (!sheet) {
    return { success: false, message: "Kullanıcılar sayfası bulunamadı!" };
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      sheet.getRange(i + 1, 2).setValue(newPin); // Pini değiştir
      logAction(username, "Pin Değiştirme", "Yeni Pin: " + newPin);
      return { success: true, message: "Pin değiştirildi." };
    }
  }
  return { success: false, message: "Kullanıcı bulunamadı." };
}

function addEvents(dates, username) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Etkinlikler");
  if (!sheet) {
    return { success: false, message: "Etkinlikler sayfası bulunamadı!" };
  }
  var data = sheet.getDataRange().getValues();
  var existingCount = data.reduce(function(count, row, idx) {
    if (idx > 0 && row[1] === username) return count + 1;
    return count;
  }, 0);
  var MAX = 14;
  var remaining = MAX - existingCount;
  var results = [];
  var added = 0;

  dates.forEach(function(date) {
    if (added >= remaining) {
      results.push({ date: date, success: false, message: "Maksimum " + MAX + " güne ulaştınız." });
      return;
    }
    var conflict = data.some(function(row, idx) {
      if (idx === 0) return false;
      var stored = (typeof row[0] === "string")
                   ? row[0]
                   : Utilities.formatDate(new Date(row[0]), TIMEZONE, "yyyy-MM-dd");
      return stored === date && row[1] === username;
    });
    if (conflict) {
      results.push({ date: date, success: false, message: "Bu gün için zaten bir etkinlik var." });
    } else {
      var etkinlikAdi = username + " yıllık izin";
      sheet.appendRow([date, username, etkinlikAdi]);
      results.push({ date: date, success: true, message: "Etkinlik eklendi.", event: { date: date, user: username, event: etkinlikAdi } });
      added++;
    }
  });
  logAction(username, "Etkinlik Ekleme", dates.join(", "));
  return results;
}

function deleteEvent(date, username) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Etkinlikler");
  if (!sheet) return { success: false, message: "Etkinlik bulunamadı." };

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var storedDate = (typeof data[i][0] === "string")
                     ? data[i][0]
                     : Utilities.formatDate(new Date(data[i][0]), TIMEZONE, "yyyy-MM-dd");
    if (storedDate === date && data[i][1] === username) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return { success: false, message: "Silme yetkiniz yok." };

  sheet.deleteRow(rowIndex);
  logAction(username, "Tekil Etkinlik Silme", date);
  return { success: true, message: "Etkinlik silindi." };
}

function deleteEvents(dates, username) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Etkinlikler");
  if (!sheet) return dates.map(function(date) {
    return { date: date, success: false, message: "Etkinlik bulunamadı." };
  });

  var data = sheet.getDataRange().getValues();
  var results = [];

  dates.forEach(function(date) {
    var wasDeleted = false;
    for (var i = 1; i < data.length; i++) {
      var storedDate = (typeof data[i][0] === "string")
                       ? data[i][0]
                       : Utilities.formatDate(new Date(data[i][0]), TIMEZONE, "yyyy-MM-dd");
      if (storedDate === date && data[i][1] === username) {
        sheet.deleteRow(i + 1);
        wasDeleted = true;
        break;
      }
    }
    results.push({ date: date, success: wasDeleted, message: wasDeleted ? "Silindi" : "Silinecek etkinlik bulunamadı" });
  });
  logAction(username, "Çoklu Etkinlik Silme", dates.join(", "));
  return results;
}
