/**
 * Seren Kuru Temizleme — Google Apps Script backend
 *
 * ÖNEMLİ DEĞİŞİKLİK (depolama sınırı düzeltmesi):
 * Eskiden tüm {cx,tx,cb,cfg} verisi SerenData!A1 tek hücresine yazılıyordu.
 * Google Sheets'te bir hücre en fazla 50.000 karakter alır; bu yüzden ~304
 * müşteriden sonra kayıt başarısız oluyordu. Artık veri, SerenData sayfasının
 * A sütununda ~40.000 karakterlik parçalara bölünerek saklanıyor (A1, A2, A3…)
 * ve okurken birleştiriliyor. Böylece binlerce müşteri sorunsuz kaydedilir.
 *
 * Login, oturum (Sessions) ve kullanıcı (Users) mantığı AYNEN korunmuştur.
 */

var DATA_SHEET = "SerenData";
var CHUNK_SIZE = 40000; // hücre başına güvenli karakter sayısı (limit 50.000)

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAuthSheets_(ss);
  ensureAdminSeed_(ss);

  var token = getTokenFromRequest_(e);
  var session = verifySession_(ss, token);
  if (!session) {
    return jsonOut_({ ok: false, error: 'Yetkisiz' });
  }

  var data = readData_(ss);
  return jsonOut_(data); // eski davranış: okurken direkt cx/tx/cb/cfg döner
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAuthSheets_(ss);
  ensureAdminSeed_(ss);

  var bodyText = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
  var body;
  try {
    body = JSON.parse(bodyText || '{}');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Geçersiz JSON' });
  }

  // LOGIN: { action:"login", phone, password }
  if (body && body.action === 'login') {
    var phone = normalizePhone_(body.phone);
    var password = String(body.password || '');

    if (!phone || password.length < 4) return jsonOut_({ ok: false, error: 'Telefon veya şifre hatalı' });

    var user = findUserByPhone_(ss, phone);
    if (!user) return jsonOut_({ ok: false, error: 'Telefon veya şifre hatalı' });

    var hash = hashPassword_(user.salt, password);
    if (hash !== user.passwordHash) return jsonOut_({ ok: false, error: 'Telefon veya şifre hatalı' });

    var token = createSession_(ss, phone, user.role);
    return jsonOut_({ ok: true, token: token, role: user.role, phone: phone, fullName: user.fullName || '' });
  }

  // Diğer tüm işlemler token ister
  var token = getTokenFromRequest_(e);
  var session = verifySession_(ss, token);
  if (!session) return jsonOut_({ ok: false, error: 'Yetkisiz' });

  // ADMIN action'ları (kullanıcı yönetimi altyapısı)
  if (body && body.action && String(body.action).indexOf('admin_') === 0) {
    if (session.role !== 'ADMIN') return jsonOut_({ ok: false, error: 'ADMIN yetkisi gerekli' });

    if (body.action === 'admin_listUsers') {
      return jsonOut_({ ok: true, users: listUsers_(ss) });
    }

    if (body.action === 'admin_addUser') {
      var phone2 = normalizePhone_(body.phone);
      var fullName2 = String(body.fullName || '');
      var role2 = String(body.role || 'STAFF');
      var password2 = String(body.password || '');

      if (!phone2 || password2.length < 4) return jsonOut_({ ok: false, error: 'Geçersiz veri' });
      if (role2 !== 'ADMIN' && role2 !== 'STAFF') return jsonOut_({ ok: false, error: 'Rol geçersiz' });
      if (findUserByPhone_(ss, phone2)) return jsonOut_({ ok: false, error: 'Bu telefon zaten kayıtlı' });

      createUser_(ss, phone2, fullName2, role2, password2);
      return jsonOut_({ ok: true });
    }

    if (body.action === 'admin_updateUser') {
      var phoneU = normalizePhone_(body.phone);
      if (!phoneU) return jsonOut_({ ok: false, error: 'Telefon gerekli' });
      var targetU = findUserByPhone_(ss, phoneU);
      if (!targetU) return jsonOut_({ ok: false, error: 'Kullanıcı bulunamadı' });

      var newFullName = (typeof body.fullName !== 'undefined') ? String(body.fullName) : targetU.fullName;
      var newRole = (typeof body.role !== 'undefined') ? String(body.role) : targetU.role;
      if (newRole !== 'ADMIN' && newRole !== 'STAFF') return jsonOut_({ ok: false, error: 'Rol geçersiz' });

      // Son admin'in rolü düşürülemesin
      if (targetU.role === 'ADMIN' && newRole !== 'ADMIN' && countAdmins_(ss) <= 1) {
        return jsonOut_({ ok: false, error: 'Son admin rolü değiştirilemez' });
      }

      updateUserFields_(ss, phoneU, newFullName, newRole);
      if (body.password && String(body.password).length >= 4) {
        updatePassword_(ss, phoneU, String(body.password));
      }
      return jsonOut_({ ok: true });
    }

    if (body.action === 'admin_deleteUser') {
      var phone3 = normalizePhone_(body.phone);
      if (!phone3) return jsonOut_({ ok: false, error: 'Telefon gerekli' });

      var target = findUserByPhone_(ss, phone3);
      if (!target) return jsonOut_({ ok: false, error: 'Kullanıcı bulunamadı' });

      if (target.role === 'ADMIN') {
        var adminCount = countAdmins_(ss);
        if (adminCount <= 1) return jsonOut_({ ok: false, error: 'Son admin silinemez' });
        if (phone3 === session.phone) return jsonOut_({ ok: false, error: 'Kendi admin hesabını silemezsin' });
      }

      deleteUserByPhone_(ss, phone3);
      return jsonOut_({ ok: true });
    }

    if (body.action === 'admin_changeOwnPassword') {
      var newPassword = String(body.newPassword || '');
      if (newPassword.length < 4) return jsonOut_({ ok: false, error: 'Şifre çok kısa' });

      updatePassword_(ss, session.phone, newPassword);
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: 'Bilinmeyen admin action' });
  }

  // Kayıt akışı: HTML saveAll {cx,tx,cb,cfg} gönderiyor
  // STAFF cfg değiştiremez: eski cfg'yi geri bas
  var oldData = readData_(ss);
  if (session.role === 'STAFF') {
    body.cfg = oldData.cfg;
  }

  // Veriyi parçalı olarak sakla (50.000 karakter hücre sınırını aşmaz)
  writeData_(ss, JSON.stringify(body));

  // İnsan tarafından okunabilir aynalama (opsiyonel görünüm)
  writeReadableSheets(ss, body);

  return jsonOut_({ ok: true });
}

// ================= PARÇALI VERİ DEPOLAMA =================

function readData_(ss) {
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) return { cx: [], tx: [], cb: [], cfg: null };

  var last = sheet.getLastRow();
  if (last < 1) return { cx: [], tx: [], cb: [], cfg: null };

  var vals = sheet.getRange(1, 1, last, 1).getValues();
  var txt = '';
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] !== null && typeof vals[i][0] !== 'undefined') txt += String(vals[i][0]);
  }

  if (!txt) return { cx: [], tx: [], cb: [], cfg: null };
  try {
    return JSON.parse(txt);
  } catch (err) {
    return { cx: [], tx: [], cb: [], cfg: null };
  }
}

function writeData_(ss, jsonText) {
  var sheet = ss.getSheetByName(DATA_SHEET) || ss.insertSheet(DATA_SHEET);

  // Eski içeriği temizle (yeni veri daha az parça içerebilir)
  var last = sheet.getLastRow();
  if (last > 0) sheet.getRange(1, 1, last, 1).clearContent();

  jsonText = jsonText || '';
  var chunks = [];
  for (var i = 0; i < jsonText.length; i += CHUNK_SIZE) {
    chunks.push([jsonText.substring(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks.push(['']);

  sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
}

// İnsan tarafından okunabilir aynalama — TEK setValues çağrısıyla (hızlı, timeout yok)
function writeReadableSheets(ss, data) {
  try {
    var cs = ss.getSheetByName("Musteriler") || ss.insertSheet("Musteriler");
    cs.clearContents();
    var crows = [["ID", "Ad Soyad", "Telefon", "Adres", "Kayit Tarihi"]];
    (data.cx || []).forEach(function (c) {
      crows.push([c.id, c.name, c.phone, c.address || "", c.createdAt || ""]);
    });
    cs.getRange(1, 1, crows.length, 5).setValues(crows);

    var ts = ss.getSheetByName("Islemler") || ss.insertSheet("Islemler");
    ts.clearContents();
    var trows = [["Islem ID", "Musteri ID", "Durum", "Tarih", "Hazir Tarihi", "Toplam", "Odenen", "Kalan"]];
    (data.tx || []).forEach(function (t) {
      var total = 0, paid = 0;
      (t.items || []).forEach(function (i) { total += (i.qty || 0) * (i.unitPrice || 0) * (1 - (i.discount || 0) / 100); });
      (t.payments || []).forEach(function (p) { paid += (p.amount || 0); });
      trows.push([t.id, t.customerId, t.status, t.createdAt || "", t.readyDate || "", total, paid, total - paid]);
    });
    ts.getRange(1, 1, trows.length, 8).setValues(trows);
  } catch (e) {}
}

// ================= AUTH SHEETS / SEED =================

function ensureAuthSheets_(ss) {
  var users = ss.getSheetByName("Users") || ss.insertSheet("Users");
  if (users.getLastRow() < 1) users.appendRow(["phone", "fullName", "role", "salt", "passwordHash"]);

  var sessions = ss.getSheetByName("Sessions") || ss.insertSheet("Sessions");
  if (sessions.getLastRow() < 1) sessions.appendRow(["token", "phone", "role", "expiresAt"]);
}

function ensureAdminSeed_(ss) {
  var phoneAdmin = normalizePhone_('5318221952');
  if (!phoneAdmin) return;

  var existing = findUserByPhone_(ss, phoneAdmin);
  if (existing) return;

  createUser_(ss, phoneAdmin, 'Admin', 'ADMIN', 'admin1');
}

function createUser_(ss, phone, fullName, role, passwordPlain) {
  var users = ss.getSheetByName("Users");
  var salt = Utilities.getUuid();
  var hash = hashPassword_(salt, passwordPlain);
  users.appendRow([phone, fullName || '', role, salt, hash]);
}

function findUserByPhone_(ss, phone) {
  var users = ss.getSheetByName("Users");
  var lastRow = users.getLastRow();
  if (lastRow < 2) return null;

  var values = users.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[0]) === String(phone)) {
      return { phone: row[0], fullName: row[1], role: row[2], salt: row[3], passwordHash: row[4] };
    }
  }
  return null;
}

function listUsers_(ss) {
  var users = ss.getSheetByName("Users");
  var lastRow = users.getLastRow();
  if (lastRow < 2) return [];

  var values = users.getRange(2, 1, lastRow - 1, 5).getValues();
  return values.map(function (r) {
    return { phone: r[0], fullName: r[1], role: r[2] };
  });
}

function deleteUserByPhone_(ss, phone) {
  var users = ss.getSheetByName("Users");
  var lastRow = users.getLastRow();
  if (lastRow < 2) return;

  var values = users.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(phone)) {
      users.deleteRow(i + 2);
      return;
    }
  }
}

function updateUserFields_(ss, phone, fullName, role) {
  var users = ss.getSheetByName("Users");
  var lastRow = users.getLastRow();
  if (lastRow < 2) return;

  var values = users.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(phone)) {
      // columns: phone(1), fullName(2), role(3)
      users.getRange(i + 2, 2).setValue(fullName || '');
      users.getRange(i + 2, 3).setValue(role);
      return;
    }
  }
}

function updatePassword_(ss, phone, newPasswordPlain) {
  var users = ss.getSheetByName("Users");
  var lastRow = users.getLastRow();
  if (lastRow < 2) return;

  var values = users.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(phone)) {
      var newSalt = Utilities.getUuid();
      var newHash = hashPassword_(newSalt, newPasswordPlain);
      // columns: phone(1), fullName(2), role(3), salt(4), passwordHash(5)
      users.getRange(i + 2, 4).setValue(newSalt);
      users.getRange(i + 2, 5).setValue(newHash);
      return;
    }
  }
}

function countAdmins_(ss) {
  var users = ss.getSheetByName("Users");
  var lastRow = users.getLastRow();
  if (lastRow < 2) return 0;

  var values = users.getRange(2, 3, lastRow - 1, 1).getValues(); // role col=3
  var count = 0;
  for (var i = 0; i < values.length; i++) if (String(values[i][0]) === 'ADMIN') count++;
  return count;
}

// ================= SESSIONS =================

function createSession_(ss, phone, role) {
  var sessions = ss.getSheetByName("Sessions");
  var token = Utilities.getUuid();
  var expires = new Date(Date.now() + 1000 * 60 * 60 * 24 * 14); // 14 gün
  sessions.appendRow([token, phone, role, expires.toISOString()]);
  return token;
}

function verifySession_(ss, token) {
  if (!token) return null;

  var sessions = ss.getSheetByName("Sessions");
  var lastRow = sessions.getLastRow();
  if (lastRow < 2) return null;

  var values = sessions.getRange(2, 1, lastRow - 1, 4).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[0]) === String(token)) {
      var expiresAt = row[3] ? new Date(String(row[3])) : null;
      if (!expiresAt || expiresAt.getTime() < Date.now()) return null;
      return { token: row[0], phone: row[1], role: row[2] };
    }
  }
  return null;
}

// ================= UTIL =================

function normalizePhone_(phone) {
  if (phone === null || typeof phone === 'undefined') return '';
  return String(phone).replace(/\s+/g, '').replace(/[^\d+]/g, '');
}

function hashPassword_(salt, passwordPlain) {
  var text = String(salt) + '|' + String(passwordPlain);
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return bytesToHex_(bytes);
}

function bytesToHex_(bytes) {
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return hex;
}

function getTokenFromRequest_(e) {
  return (e && e.parameter && e.parameter.token) ? String(e.parameter.token) : '';
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
