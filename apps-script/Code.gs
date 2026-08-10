/**
 * OSTİM Teknik Üniversitesi — Bilimsel Performans Değerlendirme
 * Google Apps Script arka ucu (ücretsiz, sunucusuz veritabanı olarak Google E-Tablo kullanır).
 *
 * Tüm birimler (Mühendislik Fakültesi, Mimarlık ve Tasarım Fakültesi, İİBF,
 * Yüksek Meslek Okulu, Bilişim Teknolojileri MYO) TEK bu dosyayla ve TEK bir
 * Google E-Tablo ile çalışır. Her birimin kendi şifresi ve bildirim e-postası
 * ayrı ayrı "Settings" sayfasında tutulur; ayrıca tüm birimleri birden gören
 * bir "merkezi yönetici" şifresi de vardır.
 *
 * KURULUM ADIMLARI (README.md dosyasında da anlatılıyor):
 *   1. sheets.google.com → Boş bir e-tablo oluşturun.
 *   2. Üst menüden Uzantılar (Extensions) > Apps Script.
 *   3. Açılan editördeki örnek kodu silip bu dosyanın TAMAMINI yapıştırın.
 *   4. Üstteki fonksiyon listesinden "setup" fonksiyonunu seçip ▶ Çalıştır (Run) butonuna basın.
 *   5. Dağıt (Deploy) > Yeni Dağıtım (New deployment) > Tür: Web Uygulaması (Web app).
 *        - Yürütme kişisi (Execute as): Ben (Me)
 *        - Erişimi olanlar (Who has access): Herkes (Anyone)
 *   6. Çıkan Web App URL'sini config.js dosyasındaki APPS_SCRIPT_URL değişkenine yapıştırın.
 *   7. Şifreleri/e-postaları değiştirmek için e-tablodaki "Settings" sayfasını düzenleyin
 *      (kod değişikliği veya yeniden dağıtım gerekmez).
 */

const SUB_SHEET = 'Submissions';
const SET_SHEET = 'Settings';

// Birim kimlikleri — app.js içindeki UNITS dizisiyle BİREBİR aynı olmalı.
const UNIT_IDS = ['muhendislik', 'mtf', 'iibf', 'myo', 'btmyo'];
const UNIT_LABELS = {
  muhendislik: 'Mühendislik Fakültesi',
  mtf: 'Mimarlık ve Tasarım Fakültesi',
  iibf: 'İktisadi ve İdari Bilimler Fakültesi (İİBF)',
  myo: 'Yüksek Meslek Okulu',
  btmyo: 'Bilişim Teknolojileri MYO',
};
const UNIT_DEFAULT_EMAILS = {
  muhendislik: 'mf.dekanlik@ostimteknik.edu.tr',
  mtf: 'sare.sahil@ostimteknik.edu.tr',
  iibf: 'ihsan.alp@ostimteknik.edu.tr',
  myo: 'zeynep.aydemir@ostimteknik.edu.tr',
  btmyo: 'btmyo.bilgi@ostimteknik.edu.tr',
};
const UNIT_DEFAULT_PASSWORDS = {
  muhendislik: 'Muh2025!',
  mtf: 'Mtf2025!',
  iibf: 'Iibf2025!',
  myo: 'Myo2025!',
  btmyo: 'Btmyo2025!',
};

// Birim, en sona (index 18) eklendi ki eski sütun sıraları bozulmasın.
const SUB_HEADERS = [
  'Key', 'AdSoyad', 'Fakulte', 'Bolum', 'Unvan', 'IdariGorev', 'Donem',
  'ValuesJSON', 'Sec1', 'Sec2', 'Sec3', 'Sec4', 'Toplam', 'Rektorluk',
  'SubmittedAt', 'UpdatedAt', 'Email', 'EvidenceJSON', 'Birim',
];

/** İlk kurulumda (veya yeni birim eklendiğinde) çalıştırın — var olan ayarları korur, eksikleri tamamlar. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sub = ss.getSheetByName(SUB_SHEET);
  if (!sub) sub = ss.insertSheet(SUB_SHEET);
  if (sub.getLastRow() === 0) sub.appendRow(SUB_HEADERS);

  let set = ss.getSheetByName(SET_SHEET);
  if (!set) set = ss.insertSheet(SET_SHEET);
  if (set.getLastRow() === 0) set.appendRow(['Anahtar', 'Değer']);

  if (!getSetting('AdminSifre')) setSetting('AdminSifre', 'Ostim2025!'); // merkezi yönetici (tüm birimleri görür)
  UNIT_IDS.forEach((id) => {
    if (!getSetting('AdminSifre_' + id)) setSetting('AdminSifre_' + id, UNIT_DEFAULT_PASSWORDS[id]);
    if (!getSetting('BildirimEposta_' + id)) setSetting('BildirimEposta_' + id, UNIT_DEFAULT_EMAILS[id]);
    if (getSetting('Esik_' + id) === null) setSetting('Esik_' + id, '');
    if (getSetting('MinSec1Puan_' + id) === null) setSetting('MinSec1Puan_' + id, '');
    if (getSetting('MinSec1Adet_' + id) === null) setSetting('MinSec1Adet_' + id, '');
    if (getSetting('CapIIIIV_' + id) === null) setSetting('CapIIIIV_' + id, '');
  });
  Logger.log('Kurulum tamam. "Deploy > New deployment" adımına geçebilirsiniz.');
}

function getSetting(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SET_SHEET);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return data[i][1];
  }
  return null;
}

function setSetting(key, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SET_SHEET);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

/**
 * Şifreyi kontrol edip rol bilgisini döner:
 *  - Merkezi yönetici şifresiyle eşleşirse: { role: 'central', birim: null }
 *  - Bir birim şifresiyle eşleşirse:        { role: 'unit', birim: '<id>' }
 *  - Eşleşme yoksa: null
 */
function resolveRole(password) {
  if (password === undefined || password === null || password === '') return null;
  const central = getSetting('AdminSifre');
  if (central && String(password) === String(central)) return { role: 'central', birim: null };
  for (let i = 0; i < UNIT_IDS.length; i++) {
    const id = UNIT_IDS[i];
    const pw = getSetting('AdminSifre_' + id);
    if (pw && String(password) === String(pw)) return { role: 'unit', birim: id };
  }
  return null;
}

function findRow(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) return i + 1; // 1-tabanlı satır no
  }
  return -1;
}

function rowToRecord(row) {
  let values = {};
  try { values = JSON.parse(row[7] || '{}'); } catch (e) {}
  let evidence = {};
  try { evidence = JSON.parse(row[17] || '{}'); } catch (e) {}
  return {
    key: row[0],
    info: {
      adSoyad: row[1], fakulte: row[2], bolum: row[3], unvan: row[4], idariGorev: row[5],
      donem: row[6], email: row[16] || '', birim: row[18] || 'muhendislik',
    },
    values: values,
    evidence: evidence,
    calc: { bySection: { 1: Number(row[8]) || 0, 2: Number(row[9]) || 0, 3: Number(row[10]) || 0, 4: Number(row[11]) || 0 }, grand: Number(row[12]) || 0 },
    rektorluk: row[13],
    submittedAt: row[14],
    updatedAt: row[15],
  };
}

/** Form gönderildiğinde öğretim elemanına ve ilgili birimin dekanlığına/müdürlüğüne bilgilendirme e-postası gönderir. */
function sendNotificationEmails(key, record) {
  const info = record.info || {};
  const calc = record.calc || { bySection: {}, grand: 0 };
  const birim = info.birim || 'muhendislik';
  const notifyTo = getSetting('BildirimEposta_' + birim) || UNIT_DEFAULT_EMAILS[birim] || UNIT_DEFAULT_EMAILS.muhendislik;
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const bodyLines = [
    'Birim: ' + (UNIT_LABELS[birim] || info.fakulte),
    'Ad Soyad: ' + info.adSoyad,
    'Unvan: ' + info.unvan,
    'Bölüm: ' + info.bolum,
    'Form Dönemi: ' + info.donem,
    '',
    'I. Bilimsel Araştırma: ' + (calc.bySection[1] || 0),
    'II. Proje: ' + (calc.bySection[2] || 0),
    'III. Hizmet Faaliyetleri: ' + (calc.bySection[3] || 0),
    'IV. Tanıtım: ' + (calc.bySection[4] || 0),
    'GENEL TOPLAM: ' + (calc.grand || 0),
    '',
    'Tüm kayıtları görüntülemek için: ' + sheetUrl,
  ];
  const body = bodyLines.join('\n');

  try {
    MailApp.sendEmail({
      to: notifyTo,
      subject: 'Yeni Performans Değerlendirme Formu — ' + info.adSoyad + ' (' + info.donem + ')',
      body: body,
    });
  } catch (e) {
    Logger.log('Birim e-postası gönderilemedi: ' + e);
  }

  if (info.email) {
    try {
      MailApp.sendEmail({
        to: info.email,
        subject: 'OSTİM Teknik Üniversitesi — Performans Değerlendirme Formunuz Kaydedildi',
        body: 'Sayın ' + info.adSoyad + ',\n\n' + info.donem + ' dönemi için gönderdiğiniz performans değerlendirme formunun bir kopyası aşağıdadır. Bu e-postayı kayıtlarınız için saklayabilirsiniz.\n\n' + body,
      });
    } catch (e) {
      Logger.log('Öğretim elemanı e-postası gönderilemedi: ' + e);
    }
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ error: 'invalid_json' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sub = ss.getSheetByName(SUB_SHEET);
  if (!sub) { setup(); sub = ss.getSheetByName(SUB_SHEET); }

  const action = body.action;

  if (action === 'submit') {
    const r = body.record;
    const birim = (r.info && r.info.birim) || 'muhendislik';
    const row = [
      body.key, r.info.adSoyad, r.info.fakulte, r.info.bolum, r.info.unvan, r.info.idariGorev, r.info.donem,
      JSON.stringify(r.values || {}),
      (r.calc && r.calc.bySection && r.calc.bySection[1]) || 0,
      (r.calc && r.calc.bySection && r.calc.bySection[2]) || 0,
      (r.calc && r.calc.bySection && r.calc.bySection[3]) || 0,
      (r.calc && r.calc.bySection && r.calc.bySection[4]) || 0,
      (r.calc && r.calc.grand) || 0,
      r.rektorluk || '',
      r.submittedAt || new Date().toISOString(),
      new Date().toISOString(),
      r.info.email || '',
      JSON.stringify(r.evidence || {}),
      birim,
    ];
    const rowIdx = findRow(sub, body.key);
    if (rowIdx > 0) sub.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else sub.appendRow(row);
    sendNotificationEmails(body.key, r);
    return jsonOut({ ok: true });
  }

  if (action === 'getOne') {
    const rowIdx = findRow(sub, body.key);
    if (rowIdx < 1) return jsonOut({ ok: true, record: null });
    const row = sub.getRange(rowIdx, 1, 1, SUB_HEADERS.length).getValues()[0];
    return jsonOut({ ok: true, record: rowToRecord(row) });
  }

  if (action === 'checkPassword') {
    const roleInfo = resolveRole(body.password);
    return jsonOut(roleInfo ? { ok: true, role: roleInfo.role, birim: roleInfo.birim } : { ok: false });
  }

  if (action === 'list') {
    const roleInfo = resolveRole(body.password);
    if (!roleInfo) return jsonOut({ error: 'unauthorized' });
    const data = sub.getDataRange().getValues();
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rec = rowToRecord(data[i]);
      if (roleInfo.role === 'unit' && rec.info.birim !== roleInfo.birim) continue;
      items.push(rec);
    }
    const relevantUnits = roleInfo.role === 'central' ? UNIT_IDS : [roleInfo.birim];
    const thresholds = {};
    const policy = {};
    relevantUnits.forEach((id) => {
      thresholds[id] = getSetting('Esik_' + id);
      policy[id] = {
        minSec1Puan: getSetting('MinSec1Puan_' + id),
        minSec1Adet: getSetting('MinSec1Adet_' + id),
        capIIIIVPercent: getSetting('CapIIIIV_' + id),
      };
    });
    return jsonOut({ ok: true, role: roleInfo.role, birim: roleInfo.birim, items: items, thresholds: thresholds, policy: policy });
  }

  if (action === 'delete') {
    const roleInfo = resolveRole(body.password);
    if (!roleInfo) return jsonOut({ error: 'unauthorized' });
    const rowIdx = findRow(sub, body.key);
    if (rowIdx > 0) {
      if (roleInfo.role === 'unit') {
        const row = sub.getRange(rowIdx, 1, 1, SUB_HEADERS.length).getValues()[0];
        if (rowToRecord(row).info.birim !== roleInfo.birim) return jsonOut({ error: 'forbidden' });
      }
      sub.deleteRow(rowIdx);
    }
    return jsonOut({ ok: true });
  }

  if (action === 'setThreshold') {
    const roleInfo = resolveRole(body.password);
    if (!roleInfo) return jsonOut({ error: 'unauthorized' });
    const targetBirim = roleInfo.role === 'central' ? body.birim : roleInfo.birim;
    if (!targetBirim || UNIT_IDS.indexOf(targetBirim) === -1) return jsonOut({ error: 'invalid_birim' });
    setSetting('Esik_' + targetBirim, body.value);
    return jsonOut({ ok: true });
  }

  // Eşik + asgari şart (I. Bölüm) + III+IV tavanını tek seferde günceller.
  if (action === 'setUnitSettings') {
    const roleInfo = resolveRole(body.password);
    if (!roleInfo) return jsonOut({ error: 'unauthorized' });
    const targetBirim = roleInfo.role === 'central' ? body.birim : roleInfo.birim;
    if (!targetBirim || UNIT_IDS.indexOf(targetBirim) === -1) return jsonOut({ error: 'invalid_birim' });
    setSetting('Esik_' + targetBirim, body.threshold);
    setSetting('MinSec1Puan_' + targetBirim, body.minSec1Puan);
    setSetting('MinSec1Adet_' + targetBirim, body.minSec1Adet);
    setSetting('CapIIIIV_' + targetBirim, body.capIIIIVPercent);
    return jsonOut({ ok: true });
  }

  if (action === 'setRektorluk') {
    const roleInfo = resolveRole(body.password);
    if (!roleInfo) return jsonOut({ error: 'unauthorized' });
    const rowIdx = findRow(sub, body.key);
    if (rowIdx > 0) {
      if (roleInfo.role === 'unit') {
        const row = sub.getRange(rowIdx, 1, 1, SUB_HEADERS.length).getValues()[0];
        if (rowToRecord(row).info.birim !== roleInfo.birim) return jsonOut({ error: 'forbidden' });
      }
      sub.getRange(rowIdx, 14).setValue(body.value);
    }
    return jsonOut({ ok: true });
  }

  return jsonOut({ error: 'unknown_action' });
}

function doGet(e) {
  return jsonOut({ ok: true, message: 'OSTİM Performans API çalışıyor. İstekler POST ile gönderilmelidir.' });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
