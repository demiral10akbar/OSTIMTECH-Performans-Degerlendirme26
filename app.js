/* =========================================================================
   OSTİM Teknik Üniversitesi — Bilimsel Performans Değerlendirme Uygulaması
   Tüm kriterler, puanlamalar ve formüller ekte verilen resmi Excel formundan
   (Müh. Fak. sayfası) birebir alınmıştır.
   ========================================================================= */

/* --------------------------- Sabitler --------------------------------- */

// Yerel mod (Apps Script bağlanmadığında) için basit, istemci-taraflı şifre.
// GÜVENLİ DEĞİLDİR — yalnızca deneme/demo amaçlıdır. Gerçek kullanımda
// config.js içine Apps Script URL'nizi girin; asıl şifre kontrolü o zaman
// Google E-Tablo'daki "Settings" sayfasında, sunucu tarafında yapılır.
const LOCAL_ADMIN_PASSWORD = "Ostim2025!";

// Yazar sırası çarpanları (I10:M10 hücrelerinden — tüm form için sabittir)
const AUTHOR_MULT = [1, 0.8, 0.6, 0.4, 0.226];
const AUTHOR_LABELS = ["1. Yazar", "2. Yazar", "3. Yazar", "4. Yazar", "5. ve üstü"];

const CRITERIA = window.CRITERIA_DATA.items;
const SECTIONS = window.CRITERIA_DATA.sections; // {1:{title,weight}, ...}

const UYARI = {
  alan: "Alan Puanı: Bilimsel araştırma, proje, hizmet ve tanıtım gibi alanların her birinden alınan toplam puanı temsil eder.",
  turuncu: "Turuncu (çoklu yazarlı) maddeler: Üniversitenin Akademik Atama ve Yükseltme yönergesi uyarınca çoklu yazarlı makale/yayınlardan alınan bölünmüş puanı temsil eder. Tek yazarlı × 1; İki yazarlı × 0,8; Üç yazarlı × 0,6; Dört yazarlı × 0,4; Beş ve üstü yazarlı × 0,226.",
  yazarSira: "Yazar Sırası: Akademisyenin ilgili yayında aldığı yazar sırasını (1., 2., 3. ... ) temsil eder.",
};

/* --------------------------- Arka uç / depolama ------------------------
   İki mod:
   - "cloud": config.js içinde APPS_SCRIPT_URL tanımlıysa, tüm veriler
     Google E-Tablo'ya bağlı Apps Script Web App üzerinden okunur/yazılır.
     Bu, farklı bilgisayarlardan girilen formların otomatik olarak
     birleşmesini sağlar (GitHub Pages gibi statik barındırmalarda gerekli
     olan gerçek "sunucu" işlevini görür).
   - "local": APPS_SCRIPT_URL boşsa, veriler yalnızca o an açık olan
     tarayıcının localStorage'ında tutulur. Bu modda öğretim elemanları
     arasında otomatik paylaşım YOKTUR; bunun yerine "kayıt dosyası indir /
     içe aktar" akışı kullanılır.
   ------------------------------------------------------------------------ */

const STORAGE_MODE = (window.APPS_SCRIPT_URL && window.APPS_SCRIPT_URL.trim()) ? "cloud" : "local";

async function apiCall(action, payload) {
  try {
    const res = await fetch(window.APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // CORS ön-uçuşunu (preflight) önlemek için
      body: JSON.stringify(Object.assign({ action: action }, payload || {})),
    });
    return await res.json();
  } catch (e) {
    console.error("Apps Script isteği başarısız:", action, e);
    return { error: "network" };
  }
}

/* ---- Yerel mod depolama yardımcıları (localStorage) ---- */
const LS_PREFIX = "ostimperf::";
function lsKey(key, shared) { return LS_PREFIX + (shared ? "shared" : "priv") + "::" + key; }
function localGet(key, shared) {
  const v = localStorage.getItem(lsKey(key, shared));
  return v === null ? null : v;
}
function localSet(key, value, shared) { localStorage.setItem(lsKey(key, shared), value); return true; }
function localDelete(key, shared) { localStorage.removeItem(lsKey(key, shared)); return true; }
function localList(prefix, shared) {
  const full = lsKey(prefix || "", shared);
  const out = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.indexOf(full) === 0) out.push(k.slice((LS_PREFIX + (shared ? "shared" : "priv") + "::").length));
  }
  return out;
}

/* ---- Yardımcılar: yedek dosyası indir / içe aktar ---- */
function downloadJSON(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 1)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}
function buildBackupPayload(list) {
  return { type: "ostim-performans-backup", generatedAt: new Date().toISOString(), submissions: list };
}

/* --------------------------- Word belgesi (ıslak imza formu) ------------
   OSTİM'in resmi "Islak İmzalı Teslim Edilecek Word Formu" şablonunu
   HTML→.doc dönüşümüyle üretir (Word bunu doğrudan açar). Kütüphane
   gerektirmez, tamamen tarayıcıda çalışır.
   ------------------------------------------------------------------------ */

function escHtml(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function nl2html(s) {
  return escHtml(s).split(/\r?\n/).map((line) => "<p style='margin:0 0 4px 0'>" + (line || "&nbsp;") + "</p>").join("");
}

function buildWordHtml(record) {
  const info = record.info || {};
  const values = record.values || {};
  const evidence = record.evidence || {};
  const unitId = info.birim || "muhendislik";
  const calc = (record.calc && record.calc.meetsMinimums !== undefined) ? record.calc : computeAll(values, unitId);

  // --- Kriter tablosu satırları ---
  let rowsHtml = "";
  [1, 2, 3, 4].forEach((sNo) => {
    const meta = SECTIONS[sNo];
    rowsHtml += "<tr><td colspan='2' style='background:#DCE6F1;font-weight:bold'>" + sNo + "</td>" +
      "<td style='background:#DCE6F1;font-weight:bold'>" + escHtml(meta.title) + "</td>" +
      "<td style='background:#DCE6F1;font-weight:bold;text-align:center'>" + fmt(calc.bySection[sNo], 1) + "</td></tr>";
    const items = CRITERIA.filter((i) => i.section === sNo);
    let lastGroup = null;
    items.forEach((item) => {
      if (item.group !== lastGroup) {
        lastGroup = item.group;
        rowsHtml += "<tr><td></td><td style='font-weight:bold'>" + item.group + "</td>" +
          "<td style='font-weight:bold' colspan='2'>" + escHtml(item.groupTitle) + "</td></tr>";
      }
      const v = values[item.id];
      let countDisplay = "";
      if (item.author_row) {
        const vv = v || {};
        const parts = ["a1", "a2", "a3", "a4", "a5"].map((k) => [k, Number(vv[k]) || 0]).filter(([, n]) => n > 0);
        const total = parts.reduce((s, [, n]) => s + n, 0);
        countDisplay = total > 0 ? (total + " (" + parts.map(([k, n]) => k + "=" + n).join(", ") + ")") : "";
      } else {
        countDisplay = v ? String(v) : "";
      }
      rowsHtml += "<tr data-item='" + item.id + "'><td></td><td>" + item.id + "</td><td>" + escHtml(item.text) + "</td>" +
        "<td style='text-align:center'>" + escHtml(countDisplay) + "</td></tr>";
    });
  });

  // --- Kanıt / ayrıntılı döküm bölümü ---
  let evidenceHtml = "";
  let lastGroup2 = null;
  CRITERIA.forEach((item) => {
    const ev = evidence[item.id];
    const hasEv = ev && ((ev.text && ev.text.trim()) || (ev.images && ev.images.length));
    const v = values[item.id];
    const declared = item.author_row
      ? ["a1", "a2", "a3", "a4", "a5"].some((k) => Number((v || {})[k]) > 0)
      : Number(v) > 0;
    if (!declared && !hasEv) return;
    if (item.group !== lastGroup2) {
      lastGroup2 = item.group;
      evidenceHtml += "<p style='margin-top:16px'><b>" + item.group + ". " + escHtml(item.groupTitle) + "</b></p>";
    }
    evidenceHtml += "<div data-evidence-for='" + item.id + "'>";
    evidenceHtml += "<p style='margin:10px 0 2px 0'><b>" + item.id + ":</b> " + escHtml(item.text) + "</p>";
    if (ev && ev.text && ev.text.trim()) evidenceHtml += "<div class='ev-text'>" + nl2html(ev.text) + "</div>";
    if (ev && ev.images && ev.images.length) {
      ev.images.forEach((img) => {
        evidenceHtml += "<p><img src='" + img.dataUrl + "' style='max-width:600px;border:1px solid #999' /></p>";
      });
    }
    if (!hasEv) evidenceHtml += "<p style='color:#999'><i>(Kanıt girilmedi)</i></p>";
    evidenceHtml += "</div>";
  });

  return `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>90</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
  body{font-family:Calibri,Arial,sans-serif;font-size:11pt;}
  table{border-collapse:collapse;width:100%;margin:10px 0;}
  td,th{border:1px solid #999;padding:4px 6px;vertical-align:top;font-size:10pt;}
  .info-table td{border:1px solid #999;padding:5px 8px;}
  h1{font-size:14pt;text-align:center;}
  .warn{font-size:9.5pt;}
</style>
</head>
<body>
<h1>OSTİM Teknik Üniversitesi<br/>Yıllık Bilimsel Performans Değerlendirme Formu</h1>
<p class='warn'><u>Uyarı:</u> Bu form, beyan edilen faaliyetlerin ayrıntılı açıklamasının sunulması adına hazırlanmıştır. En altta bulunan kısma faaliyetlerin ayrıntılı dökümü eklenmiştir. Bu formun <u>ıslak imzalı matbu hali</u> ve <u>imzalı taranmış hali</u> ilgili Dekanlığa/Müdürlüğe/Başkanlığa gerekli son tarihe kadar teslim edilecektir.</p>

<span id='ostim-meta' style='display:none'>${escHtml(JSON.stringify({ birim: unitId, email: info.email || "" }))}</span>
<table class='info-table'>
<tr data-field='adSoyad'><td style='width:22%'><b>Adı Soyadı</b></td><td>${escHtml(info.adSoyad)}</td></tr>
<tr data-field='email'><td><b>E-posta</b></td><td>${escHtml(info.email || "")}</td></tr>
<tr data-field='fakulte'><td><b>Fakültesi</b></td><td>${escHtml(info.fakulte)}</td></tr>
<tr data-field='bolum'><td><b>Bölümü</b></td><td>${escHtml(info.bolum)}</td></tr>
<tr data-field='unvan'><td><b>Unvanı</b></td><td>${escHtml(info.unvan)}</td></tr>
<tr data-field='idariGorev'><td><b>İdari Görevi</b></td><td>${escHtml(info.idariGorev || "YOK")}</td></tr>
<tr data-field='donem'><td><b>Form Dönemi</b></td><td>${escHtml(info.donem)}</td></tr>
</table>

<table>
<tr><th colspan='2'>Sıra</th><th>Kriterler</th><th>Faaliyet Adedi</th></tr>
${rowsHtml}
</table>
<p style='text-align:right'><b>Genel Toplam: ${fmt(calc.grand, 1)}</b></p>
${calc.meetsMinimums !== undefined ? `<table style='margin:6px 0 16px'><tr><td style='width:55%'><b>Değerlendirme Toplamı</b> (III+IV tavanı uygulanmış)</td><td style='text-align:center'>${fmt(calc.adjustedGrand, 1)}${calc.capped ? " (tavan uygulandı)" : ""}</td></tr><tr><td><b>I. Bölüm Asgari Şart</b></td><td style='text-align:center'>${calc.meetsMinimums ? "Karşılıyor" : "Karşılamıyor"}</td></tr></table>` : ""}

<p><b><u>YUKARIDAKİ BEYANIN AYRINTILI DÖKÜMÜ:</u></b></p>
${evidenceHtml || "<p><i>Hiçbir madde için kanıt/ayrıntı girilmedi.</i></p>"}

<p style='margin-top:50px'>Öğretim Üyesi<br/>İsim-soyisim-unvan-imza</p>
<p>${escHtml(info.unvan)} ${escHtml(info.adSoyad)}</p>
<p style='margin-top:40px;font-size:8.5pt;color:#888'>Bu uygulama Prof. Dr. Demiral AKBAR tarafından geliştirilmiştir · 08.2026</p>
</body></html>`;
}

function downloadWordDoc(record) {
  const html = buildWordHtml(record);
  const blob = new Blob(["\ufeff", html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "OSTIM-Performans-" + slugify(record.info.adSoyad) + ".doc";
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

/* --------------------------- Word belgesinden içe aktarma ----------------
   "İmza İçin Word Belgesi" ile indirilen .doc dosyasını (veya kişi bunu
   Microsoft Word'de gerçek .docx olarak kaydettiyse onu da, best-effort)
   geri okuyup forma yükler. .doc dosyalarımız aslında HTML olduğundan
   DOMParser ile doğrudan, gerçek .docx'ler için mammoth.js (CDN, yalnızca
   ihtiyaç halinde yüklenir) ile önce HTML'e çevrilip aynı mantıkla okunur.
   ------------------------------------------------------------------------ */

const CRITERIA_BY_ID = {};
CRITERIA.forEach((i) => { CRITERIA_BY_ID[i.id] = i; });

function parseAuthorBreakdown(str) {
  const out = {};
  const re = /a([1-5])\s*=\s*(\d+)/g;
  let m;
  while ((m = re.exec(str))) out["a" + m[1]] = String(m[2]);
  return out;
}

function parseImportedWordHtml(htmlText) {
  const doc = new DOMParser().parseFromString(htmlText, "text/html");
  const result = { info: {}, values: {}, evidence: {} };

  const metaEl = doc.getElementById("ostim-meta");
  let meta = null;
  if (metaEl) { try { meta = JSON.parse(metaEl.textContent); } catch (e) {} }

  const labelMap = {
    "Adı Soyadı": "adSoyad", "E-posta": "email", "Fakültesi": "fakulte",
    "Bölümü": "bolum", "Unvanı": "unvan", "İdari Görevi": "idariGorev", "Form Dönemi": "donem",
  };
  doc.querySelectorAll("tr").forEach((tr) => {
    const tds = tr.querySelectorAll("td");
    if (tds.length !== 2) return;
    const field = tr.getAttribute("data-field") || labelMap[tds[0].textContent.trim()];
    if (!field) return;
    let val = tds[1].textContent.trim();
    if (field === "idariGorev" && val === "YOK") val = "";
    result.info[field] = val;
  });
  if (meta) {
    if (meta.birim) result.info.birim = meta.birim;
    if (meta.email) result.info.email = meta.email;
  }
  if (!result.info.birim) {
    const unit = UNITS.find((u) => u.label === result.info.fakulte);
    result.info.birim = unit ? unit.id : "muhendislik";
  }
  result.info.fakulte = result.info.fakulte || unitById(result.info.birim).label;

  function readCountCell(item, cellText) {
    if (!cellText) return null;
    if (item.author_row) {
      const breakdown = parseAuthorBreakdown(cellText);
      return Object.keys(breakdown).length ? breakdown : null;
    }
    const numMatch = cellText.match(/-?\d+(?:[.,]\d+)?/);
    return numMatch ? numMatch[0].replace(",", ".") : null;
  }

  // Birincil yol: data-item etiketiyle (kendi ürettiğimiz .doc'ta her zaman var)
  doc.querySelectorAll("tr[data-item]").forEach((tr) => {
    const item = CRITERIA_BY_ID[tr.getAttribute("data-item")];
    if (!item) return;
    const tds = tr.querySelectorAll("td");
    if (tds.length < 4) return;
    const val = readCountCell(item, tds[3].textContent.trim());
    if (val !== null) result.values[item.id] = val;
  });

  // Yedek yol: gerçek .docx olarak yeniden kaydedilmiş olabilir (özel etiketler kaybolur) —
  // ikinci sütundaki görünür ID metniyle eşleştirmeyi dener.
  if (Object.keys(result.values).length === 0) {
    doc.querySelectorAll("table tr").forEach((tr) => {
      const tds = tr.querySelectorAll("td");
      if (tds.length < 4) return;
      const item = CRITERIA_BY_ID[tds[1].textContent.trim()];
      if (!item) return;
      const val = readCountCell(item, tds[3].textContent.trim());
      if (val !== null) result.values[item.id] = val;
    });
  }

  doc.querySelectorAll("[data-evidence-for]").forEach((div) => {
    const id = div.getAttribute("data-evidence-for");
    const textDiv = div.querySelector(".ev-text");
    const text = textDiv ? [...textDiv.querySelectorAll("p")].map((p) => p.textContent).join("\n").trim() : "";
    const images = [...div.querySelectorAll("img")].map((img, idx) => ({ name: "gorsel-" + (idx + 1) + ".png", dataUrl: img.getAttribute("src") || "" })).filter((im) => im.dataUrl);
    if (text || images.length) result.evidence[id] = { text, images };
  });

  if (Object.keys(result.values).length === 0 && !result.info.adSoyad) {
    throw new Error("Bu belgede tanınabilir bir OSTİM formu bulunamadı.");
  }
  return result;
}

function loadMammoth() {
  return new Promise((resolve, reject) => {
    if (window.mammoth) return resolve(window.mammoth);
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/mammoth@1.7.0/mammoth.browser.min.js";
    s.onload = () => resolve(window.mammoth);
    s.onerror = () => reject(new Error("Word dönüştürücü (mammoth.js) yüklenemedi. İnternet bağlantınızı kontrol edin."));
    document.head.appendChild(s);
  });
}

async function parseImportedFile(file) {
  const name = (file.name || "").toLowerCase();
  if (name.endsWith(".docx")) {
    const buf = await file.arrayBuffer();
    const mammoth = await loadMammoth();
    const res = await mammoth.convertToHtml({ arrayBuffer: buf });
    return parseImportedWordHtml(res.value);
  }
  // .doc (bizim ürettiğimiz HTML tabanlı dosya) veya .json değilse düz metin olarak dene
  const text = await file.text();
  return parseImportedWordHtml(text);
}

function slugify(str) {
  return (str || "")
    .toLocaleLowerCase("tr-TR")
    .replace(/ç/g, "c").replace(/ğ/g, "g").replace(/ı/g, "i").replace(/ö/g, "o")
    .replace(/ş/g, "s").replace(/ü/g, "u").replace(/İ/g, "i")
    .replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
}

/* --------------------------- Uygulama durumu ---------------------------- */

const state = {
  view: "landing",
  facultyInfo: null,       // {adSoyad, fakulte, bolum, unvan, idariGorev, donem}
  submissionKey: null,     // storage key for the in-progress submission
  values: {},              // itemId -> number OR {a1..a5}
  evidence: {},            // itemId -> { text, images: [{name, dataUrl}] }
  editingExisting: false,
  adminAuthed: false,
  adminPassword: null,     // in-memory only, sent with each cloud request
  adminRole: "central",    // "central" | "unit"
  adminBirim: null,        // birim id when adminRole === "unit"
  adminFilterBirim: "all", // central görünümünde filtre ("all" veya birim id)
  adminSubmissions: [],    // cached list for dashboard
  adminThresholds: {},     // birim id -> eşik değeri
  adminPolicy: {},         // birim id -> {minSec1Puan, minSec1Adet, capIIIIVPercent}
  adminView: "dashboard",  // dashboard | detail
  adminDetailKey: null,
  lastSubmittedRecord: null,
};

const root = document.getElementById("app-root");

/* --------------------------- Birimler (fakülte / MYO) ------------------- */
// Not: Yalnızca Mühendislik Fakültesi (F sütunu) ve İİBF (E sütunu) için
// kaynak Excel'de ayrı puanlama sütunu vardı. MTF, YMO ve Bilişim
// Teknolojileri MYO için ayrı bir puan tablosu sağlanmadığından, kriterler
// aynı olduğu için bu üç birim GEÇİCİ olarak Mühendislik Fakültesi puan
// tablosunu kullanır. Gerçek puan tabloları sağlandığında `scoreKey` alanı
// güncellenmelidir.
const UNITS = [
  { id: "muhendislik", label: "Mühendislik Fakültesi", scoreKey: "puan", email: "mf.dekanlik@ostimteknik.edu.tr" },
  { id: "mtf", label: "Mimarlık ve Tasarım Fakültesi", scoreKey: "puan", email: "sare.sahil@ostimteknik.edu.tr" },
  { id: "iibf", label: "İktisadi ve İdari Bilimler Fakültesi (İİBF)", scoreKey: "iibfPuan", email: "ihsan.alp@ostimteknik.edu.tr" },
  { id: "myo", label: "Yüksek Meslek Okulu", scoreKey: "puan", email: "zeynep.aydemir@ostimteknik.edu.tr" },
  { id: "btmyo", label: "Bilişim Teknolojileri MYO", scoreKey: "puan", email: "btmyo.bilgi@ostimteknik.edu.tr" },
];
function unitById(id) { return UNITS.find((u) => u.id === id) || UNITS[0]; }
function birimHintText(birimId) {
  const u = unitById(birimId);
  if (birimId === "muhendislik" || birimId === "iibf") {
    return "Bu birim için resmi puanlama tablosu kullanılıyor. Bildirimler: " + u.email;
  }
  return "Bu birim için ayrı bir puan tablosu henüz sağlanmadığından, kriterler aynı olduğu için Mühendislik Fakültesi puan tablosu geçici olarak kullanılıyor. Bildirimler: " + u.email;
}

/* --------------------------- Hesaplama mantığı -------------------------- */

function computeItemScore(item, val, unitId) {
  const unit = unitById(unitId || "muhendislik");
  const puan = item[unit.scoreKey] || 0;
  if (item.author_row) {
    const v = val || {};
    let total = 0;
    for (let k = 0; k < 5; k++) {
      const cnt = Number(v["a" + (k + 1)]) || 0;
      if (cnt > 0) total += cnt * AUTHOR_MULT[k] * puan;
    }
    return total;
  }
  const cnt = Number(val) || 0;
  return cnt * puan;
}

function computeAll(values, unitId) {
  const bySection = { 1: 0, 2: 0, 3: 0, 4: 0 };
  const byItem = {};
  CRITERIA.forEach((item) => {
    const s = computeItemScore(item, values[item.id], unitId);
    byItem[item.id] = s;
    bySection[item.section] = (bySection[item.section] || 0) + s;
  });
  const grand = bySection[1] + bySection[2] + bySection[3] + bySection[4];
  return { bySection, byItem, grand };
}

function sectionItemCount(values, sectionNo) {
  let total = 0;
  CRITERIA.forEach((item) => {
    if (item.section !== sectionNo) return;
    const v = values[item.id];
    if (item.author_row) {
      const vv = v || {};
      total += ["a1", "a2", "a3", "a4", "a5"].reduce((s, k) => s + (Number(vv[k]) || 0), 0);
    } else {
      total += Number(v) || 0;
    }
  });
  return total;
}

function numOrNull(v) { return (v === undefined || v === null || v === "") ? null : Number(v); }

/**
 * Ham hesaplamanın üzerine, komisyonun belirleyeceği iki politika mekanizmasını uygular:
 *  A) I. Bölüm (Bilimsel Araştırma) için asgari puan ve/veya asgari faaliyet adedi şartı.
 *  B) III+IV (Hizmet+Tanıtım) toplamının, I+II toplamına oranla bir TAVAN ile sınırlanması.
 * Her iki mekanizma da `policy` boş/eksik geldiğinde devre dışıdır (mevcut davranış korunur).
 */
function computeEvaluation(values, unitId, policy) {
  const calc = computeAll(values, unitId);
  policy = policy || {};

  const capPct = numOrNull(policy.capIIIIVPercent);
  const rawIIIIV = calc.bySection[3] + calc.bySection[4];
  let adjIIIIV = rawIIIIV;
  let capped = false;
  if (capPct !== null && capPct >= 0) {
    const capValue = (capPct / 100) * (calc.bySection[1] + calc.bySection[2]);
    if (rawIIIIV > capValue) { adjIIIIV = capValue; capped = true; }
  }
  const adjustedGrand = calc.bySection[1] + calc.bySection[2] + adjIIIIV;

  const sec1Count = sectionItemCount(values, 1);
  const minPuan = numOrNull(policy.minSec1Puan);
  const minAdet = numOrNull(policy.minSec1Adet);
  const puanConfigured = minPuan !== null;
  const adetConfigured = minAdet !== null;
  const meetsPuan = !puanConfigured || calc.bySection[1] >= minPuan;
  const meetsAdet = !adetConfigured || sec1Count >= minAdet;

  // Asgari şart mantığı VEYA'dır: puan ve adet şartlarından İKİSİ de tanımlıysa,
  // ikisinden birinin karşılanması yeterlidir (örn. tek ama nitelikli bir yayınla
  // yüksek puan alan biri, "en az 2 adet" şartını karşılamasa da geçerli sayılır).
  // Yalnızca biri tanımlıysa, o tek şart geçerlidir. Hiçbiri tanımlı değilse sınırsızdır.
  let meetsMinimums;
  if (!puanConfigured && !adetConfigured) meetsMinimums = true;
  else if (puanConfigured && adetConfigured) meetsMinimums = meetsPuan || meetsAdet;
  else meetsMinimums = puanConfigured ? meetsPuan : meetsAdet;

  return Object.assign({}, calc, {
    rawGrand: calc.grand,
    adjustedGrand, capped, rawIIIIV, adjIIIIV, capPct,
    sec1Count, minPuan, minAdet, meetsPuan, meetsAdet,
    meetsMinimums,
  });
}

function fmt(n, dec) {
  const d = dec === undefined ? 2 : dec;
  return Number(n || 0).toLocaleString("tr-TR", { minimumFractionDigits: 0, maximumFractionDigits: d });
}

/* --------------------------- Render: shell ------------------------------ */

function setView(v) {
  state.view = v;
  render();
  window.scrollTo({ top: 0, behavior: "instant" in window ? "instant" : "auto" });
}

function el(tag, attrs, children) {
  const e = document.createElement(tag);
  if (attrs) {
    Object.keys(attrs).forEach((k) => {
      const v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === "class") e.className = v;
      else if (k === "html") e.innerHTML = v;
      else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v);
    });
  }
  (children || []).forEach((c) => {
    if (c === null || c === undefined) return;
    if (typeof c === "string" || typeof c === "number") e.appendChild(document.createTextNode(String(c)));
    else e.appendChild(c);
  });
  return e;
}

function toast(msg) {
  const t = el("div", { class: "toast" }, [msg]);
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2600);
}

function renderStorageBanner() {
  if (STORAGE_MODE === "cloud") return null;
  return el("div", {
    style: "background:var(--orange-soft);color:#7a3e12;padding:9px 24px;font-size:12.5px;text-align:center;border-bottom:1px solid #E3B98D",
  }, [
    "⚠ Apps Script bağlı değil — veriler yalnızca bu tarayıcıda saklanıyor, öğretim elemanları arasında otomatik paylaşılmıyor. Kurulum için README.md dosyasına bakın.",
  ]);
}

function render() {
  root.innerHTML = "";
  root.appendChild(renderTopbar());
  const banner = renderStorageBanner();
  if (banner) root.appendChild(banner);
  const main = el("main", {}, [renderView()]);
  root.appendChild(main);
  const footer = el("div", { class: "footer-note" }, [
    el("div", {}, ["OSTİM Teknik Üniversitesi · Yıllık Bilimsel Performans Değerlendirme Formu — bu araç, resmi Excel formundaki kriter ve formülleri esas alır."]),
    el("div", { style: "margin-top:4px;color:var(--ink-soft)" }, ["Bu uygulama Prof. Dr. Demiral AKBAR tarafından geliştirilmiştir · 08.2026"]),
  ]);
  root.appendChild(footer);
}

function renderTopbar() {
  const bar = el("div", { class: "topbar" }, [
    el("div", { class: "brand" }, [
      el("div", { class: "brand-mark" }, ["OT"]),
      el("div", { class: "brand-text" }, [
        el("div", { class: "t1" }, ["OSTİM Teknik Üniversitesi"]),
        el("div", { class: "t2" }, ["Bilimsel Performans Değerlendirme"]),
      ]),
    ]),
    el("div", { class: "topbar-actions" }, [
      el("button", {
        class: "pill-btn" + (state.view.startsWith("faculty") || state.view === "landing" ? " active" : ""),
        onclick: () => setView("landing"),
      }, ["Ana Sayfa"]),
      el("button", {
        class: "pill-btn" + (state.view.startsWith("admin") ? " active" : ""),
        onclick: () => setView(state.adminAuthed ? "admin-dashboard" : "admin-login"),
      }, ["Yönetici Paneli"]),
    ]),
  ]);
  return bar;
}

function renderView() {
  switch (state.view) {
    case "landing": return viewLanding();
    case "faculty-info": return viewFacultyInfo();
    case "faculty-form": return viewFacultyForm();
    case "faculty-done": return viewFacultyDone();
    case "admin-login": return viewAdminLogin();
    case "admin-dashboard": return viewAdminDashboard();
    case "admin-detail": return viewAdminDetail();
    default: return viewLanding();
  }
}

/* --------------------------- Landing ------------------------------------ */

function viewLanding() {
  return el("div", { class: "view" }, [
    el("div", { class: "landing-hero" }, [
      el("div", { class: "eyebrow" }, ["Form Dönemi Değerlendirmesi"]),
      el("h1", {}, ["Yıllık Bilimsel Performans", el("br"), "Değerlendirme Formu"]),
      el("p", {}, [
        "Bilimsel Araştırma, Proje, Hizmet Faaliyetleri ve Tanıtım kriterlerini içeren resmi OSTİM Teknik Üniversitesi formunun dijital sürümü. Öğretim elemanları verilerini girer, yönetim tüm sonuçları toplu olarak inceler.",
      ]),
    ]),
    el("div", { class: "role-cards" }, [
      el("div", { class: "role-card" }, [
        el("div", { class: "num" }, ["01 — ÖĞRETİM ELEMANI"]),
        el("h3", {}, ["Formumu Doldur"]),
        el("p", {}, ["Kimlik bilgilerinizi girin, ardından I–IV. bölümlerdeki faaliyet sayılarınızı işleyerek puanınızı canlı olarak görün."]),
        el("button", { onclick: () => { state.editingExisting = false; setView("faculty-info"); } }, ["Forma Başla →"]),
      ]),
      el("div", { class: "role-card admin-card" }, [
        el("div", { class: "num" }, ["02 — YÖNETİCİ"]),
        el("h3", {}, ["Sonuçları Görüntüle"]),
        el("p", {}, ["Şifre ile giriş yapın; tüm öğretim üyelerinin puanlarını, istatistikleri ve kriterlerin üzerinde performans gösterenleri inceleyin."]),
        el("button", { onclick: () => setView(state.adminAuthed ? "admin-dashboard" : "admin-login") }, ["Panele Git →"]),
      ]),
    ]),
  ]);
}

/* --------------------------- Faculty: bilgi formu ------------------------ */

function viewFacultyInfo() {
  const f = state.facultyInfo || {
    adSoyad: "", birim: "muhendislik", bolum: "", unvan: "Prof. Dr.",
    idariGorev: "", donem: "", email: "",
  };
  const errBox = el("div", { class: "error-msg hidden" }, []);

  const inpAd = el("input", { type: "text", value: f.adSoyad, placeholder: "Örn. Ayşe YILMAZ" });
  const inpEmail = el("input", { type: "email", value: f.email || "", placeholder: "Örn. ayse.yilmaz@ostimteknik.edu.tr" });
  const inpBolum = el("input", { type: "text", value: f.bolum, placeholder: "Örn. Makine Mühendisliği" });
  const selUnvan = el("select", {}, ["Prof. Dr.", "Doç. Dr.", "Dr. Öğr. Üyesi", "Öğr. Gör.", "Öğr. Gör. Dr.", "Arş. Gör.", "Arş. Gör. Dr."].map((u) =>
    el("option", { value: u, selected: u === f.unvan ? "selected" : null }, [u])
  ));
  const selBirim = el("select", {}, UNITS.map((u) =>
    el("option", { value: u.id, selected: u.id === (f.birim || "muhendislik") ? "selected" : null }, [u.label])
  ));
  const birimHint = el("div", { class: "hint" }, [birimHintText(f.birim || "muhendislik")]);
  selBirim.addEventListener("change", () => { birimHint.textContent = birimHintText(selBirim.value); });

  const inpIdari = el("input", { type: "text", value: f.idariGorev, placeholder: "Varsa (örn. Bölüm Başkanı)" });
  const inpDonem = el("input", { type: "text", value: f.donem, placeholder: "Örn. 1 Temmuz 2024 - 30 Haziran 2025" });

  const submit = () => {
    const adSoyad = inpAd.value.trim();
    const donem = inpDonem.value.trim();
    const bolum = inpBolum.value.trim();
    const email = inpEmail.value.trim();
    if (!adSoyad || !donem || !bolum || !email) {
      errBox.textContent = "Lütfen Ad Soyad, Bölüm, E-posta ve Form Dönemi alanlarını doldurun.";
      errBox.classList.remove("hidden");
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      errBox.textContent = "Lütfen geçerli bir e-posta adresi girin.";
      errBox.classList.remove("hidden");
      return;
    }
    const birim = selBirim.value;
    state.facultyInfo = {
      adSoyad, birim, fakulte: unitById(birim).label, bolum, unvan: selUnvan.value,
      idariGorev: inpIdari.value.trim(), donem, email,
    };
    const key = "submission:" + birim + "__" + slugify(adSoyad) + "__" + slugify(donem);
    state.submissionKey = key;
    loadExistingOrBlank(key).then(() => setView("faculty-form"));
  };

  const importInput = el("input", { type: "file", accept: "application/json,.json,.doc,.docx", style: "display:none" });
  const importMsg = el("div", { class: "ok-msg hidden" }, []);
  function applyImportedRecord(rec, key) {
    const birim = rec.info.birim || "muhendislik";
    state.facultyInfo = Object.assign({ birim, fakulte: unitById(birim).label }, rec.info);
    state.values = rec.values || {};
    state.evidence = {};
    Object.keys(rec.evidence || {}).forEach((id) => {
      const ev = rec.evidence[id] || {};
      state.evidence[id] = { text: ev.text || "", images: ev.images || [] };
    });
    state.submissionKey = key || ("submission:" + birim + "__" + slugify(rec.info.adSoyad) + "__" + slugify(rec.info.donem));
    state.editingExisting = true;
    setView("faculty-form");
  }
  importInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    importMsg.classList.add("hidden");
    try {
      if (name.endsWith(".json")) {
        const text = await file.text();
        const parsed = JSON.parse(text);
        const entry = parsed.submissions ? parsed.submissions[0] : (parsed.info ? { key: null, record: parsed } : null);
        if (!entry || !entry.record || !entry.record.info) throw new Error("Geçersiz dosya");
        applyImportedRecord(entry.record, entry.key);
      } else {
        toast("Word belgesi okunuyor…");
        const rec = await parseImportedFile(file);
        if (!rec.info.adSoyad) throw new Error("Ad Soyad bulunamadı");
        applyImportedRecord(rec, null);
        toast("Word belgesinden " + Object.keys(rec.values).length + " madde geri yüklendi.");
      }
    } catch (err) {
      importMsg.classList.remove("ok-msg"); importMsg.classList.add("error-msg");
      importMsg.textContent = "Dosya okunamadı (" + (err && err.message ? err.message : "bilinmeyen hata") + "). Lütfen daha önce bu uygulamadan indirdiğiniz bir .json veya .doc/.docx dosyası seçtiğinizden emin olun.";
      importMsg.classList.remove("hidden");
    }
  });

  return el("div", { class: "view view-narrow" }, [
    el("div", { class: "card" }, [
      el("h2", {}, ["Öğretim Elemanı Bilgileri"]),
      el("div", { class: "sub" }, ["Formu doldurmadan önce kimlik bilgilerinizi girin. Aynı ad, birim ve dönem ile daha önce kaydettiyseniz kaldığınız yerden devam edebilirsiniz."]),
      el("div", { class: "field-grid" }, [
        el("div", { class: "field full" }, [el("label", {}, ["Adı Soyadı *"]), inpAd]),
        el("div", { class: "field full" }, [el("label", {}, ["E-posta Adresiniz *"]), inpEmail, el("div", { class: "hint" }, ["Gönderdiğiniz formun bir kopyası bu adrese, birim dekanlığına/müdürlüğüne ise bilgi olarak iletilir."])]),
        el("div", { class: "field full" }, [el("label", {}, ["Fakülte / Birim *"]), selBirim, birimHint]),
        el("div", { class: "field" }, [el("label", {}, ["Bölüm *"]), inpBolum]),
        el("div", { class: "field" }, [el("label", {}, ["Unvan"]), selUnvan]),
        el("div", { class: "field" }, [el("label", {}, ["İdari Görevi"]), inpIdari]),
        el("div", { class: "field full" }, [el("label", {}, ["Form Dönemi *"]), inpDonem]),
      ]),
      errBox,
      el("div", { class: "form-actions" }, [
        el("button", { class: "btn btn-ghost", onclick: () => setView("landing") }, ["Vazgeç"]),
        el("button", { class: "btn btn-primary", onclick: submit }, ["Devam Et →"]),
      ]),
    ]),
    el("div", { class: "card" }, [
      el("h2", { style: "font-size:16px" }, ["Daha önce indirdiğiniz bir taslağınız mı var?"]),
      el("div", { class: "sub" }, ["“İlerlemeyi Yedekle (.json)” ile indirdiğiniz bir dosyayı, ya da “İmza İçin Word Belgesi” ile indirdiğiniz .doc/.docx dosyasını yükleyerek girdiğiniz faaliyet bilgileriyle kaldığınız yerden devam edebilirsiniz. (Word'den içe aktarımda, dosyayı Word'de aşırı değiştirmediyseniz en iyi sonucu alırsınız; kanıt görselleri Word belgesinde varsa onlar da geri yüklenir.)"]),
      importMsg,
      el("button", { class: "btn btn-ghost", onclick: () => importInput.click() }, ["📤 Taslak veya Word Belgesi Yükle (.json / .doc / .docx)"]),
      importInput,
    ]),
  ]);
}

async function getRecord(key) {
  if (STORAGE_MODE === "cloud") {
    const r = await apiCall("getOne", { key });
    return (r && r.ok && r.record) ? r.record : null;
  }
  const raw = localGet(key, true);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

async function putRecord(key, record) {
  if (STORAGE_MODE === "cloud") {
    const r = await apiCall("submit", { key, record });
    return !!(r && r.ok);
  }
  return localSet(key, JSON.stringify(record), true);
}

function collectEvidence(includeImages) {
  const out = {};
  Object.keys(state.evidence).forEach((id) => {
    const e = state.evidence[id];
    if (!e) return;
    const hasText = e.text && e.text.trim();
    const hasImages = e.images && e.images.length;
    if (!hasText && !hasImages) return;
    out[id] = { text: e.text || "" };
    if (includeImages && hasImages) out[id].images = e.images;
  });
  return out;
}

async function loadExistingOrBlank(key) {
  const rec = await getRecord(key);
  if (rec) {
    state.values = rec.values || {};
    state.evidence = {};
    if (rec.evidence) {
      Object.keys(rec.evidence).forEach((id) => {
        const e = rec.evidence[id] || {};
        state.evidence[id] = { text: e.text || "", images: e.images || [] };
      });
    }
    if (rec.info) state.facultyInfo = rec.info;
    state.editingExisting = true;
    return;
  }
  state.values = {};
  state.evidence = {};
  state.editingExisting = false;
}

/* --------------------------- Faculty: kriter formu ----------------------- */

let openSections = { 1: true, 2: false, 3: false, 4: false };

function viewFacultyForm() {
  const unitId = state.facultyInfo.birim || "muhendislik";
  const calc = computeAll(state.values, unitId);

  const legend = el("div", { class: "legend" }, [
    el("div", { class: "legend-item" }, [el("div", { class: "legend-swatch", style: "background:var(--navy)" }), el("div", {}, [el("b", {}, ["Alan Puanı — "]), UYARI.alan.replace("Alan Puanı: ", "")])]),
    el("div", { class: "legend-item" }, [el("div", { class: "legend-swatch", style: "background:var(--orange)" }), el("div", {}, [el("b", {}, ["Turuncu maddeler — "]), UYARI.turuncu.replace(/^Turuncu[^:]*: /, "")])]),
    el("div", { class: "legend-item" }, [el("div", { class: "legend-swatch", style: "background:var(--gold)" }), el("div", {}, [el("b", {}, ["Yazar Sırası — "]), UYARI.yazarSira.replace("Yazar Sırası: ", "")])]),
  ]);

  const sectionsEl = [1, 2, 3, 4].map((sNo) => renderSectionBlock(sNo, calc, unitId));

  const errBox = el("div", { class: "error-msg hidden" }, []);
  const okBox = el("div", { class: "ok-msg hidden" }, []);

  const saveDraft = async (silent) => {
    // Varsa mevcut submittedAt / rektörlük değerini koru
    const existing = await getRecord(state.submissionKey);
    const record = {
      info: state.facultyInfo,
      values: state.values,
      evidence: collectEvidence(STORAGE_MODE !== "cloud"), // görseller yalnızca yerel modda kaydedilir
      calc: computeAll(state.values, unitId),
      updatedAt: new Date().toISOString(),
      submittedAt: (existing && existing.submittedAt) || new Date().toISOString(),
      rektorluk: existing ? existing.rektorluk : undefined,
    };
    const ok = await putRecord(state.submissionKey, record);
    if (STORAGE_MODE === "cloud" && !ok) {
      errBox.textContent = "Sunucuya kaydedilemedi. İnternet bağlantınızı kontrol edip tekrar deneyin (Apps Script URL yanlış olabilir).";
      errBox.classList.remove("hidden");
      return null;
    }
    if (!silent) {
      okBox.textContent = "Kaydedildi.";
      okBox.classList.remove("hidden");
      setTimeout(() => okBox.classList.add("hidden"), 2000);
    }
    return record;
  };

  const submitFinal = async () => {
    const record = await saveDraft(true);
    if (!record) return; // hata oluştu, errBox zaten gösterildi
    state.lastSubmittedRecord = record;
    toast(STORAGE_MODE === "cloud" ? "Form sunucuya kaydedildi. Teşekkürler!" : "Form bu tarayıcıya kaydedildi.");
    setView("faculty-done");
  };

  const summary = el("div", { class: "summary-panel" }, [
    el("h3", {}, [state.facultyInfo.adSoyad]),
    el("div", { style: "font-size:12px;color:#AAB6C0" }, [state.facultyInfo.unvan + " · " + state.facultyInfo.bolum]),
    el("div", { class: "grand-label" }, ["Genel Toplam Puan"]),
    el("div", { class: "grand" }, [fmt(calc.grand, 2)]),
    el("div", { class: "summary-rows" }, [1, 2, 3, 4].map((sNo) =>
      el("div", { class: "summary-row" }, [
        el("span", { class: "l" }, [SECTIONS[sNo].title]),
        el("span", { class: "v" }, [fmt(calc.bySection[sNo], 1)]),
      ])
    )),
    el("div", { class: "submit-note" }, ["Değerleri girdikçe puan otomatik güncellenir. Formu tamamladığınızda “Formu Kaydet ve Gönder” butonuna basın."]),
    el("div", { style: "display:flex;flex-direction:column;gap:8px;margin-top:16px" }, [
      el("button", { class: "btn btn-gold", onclick: submitFinal }, ["Formu Kaydet ve Gönder"]),
      el("button", { class: "btn btn-ghost", style: "background:transparent;color:#EFE6D2;border-color:rgba(239,230,210,.3)", onclick: () => saveDraft(false) }, ["Taslak Olarak Kaydet"]),
      el("button", {
        class: "btn btn-ghost", style: "background:transparent;color:#EFE6D2;border-color:rgba(239,230,210,.3)",
        onclick: () => {
          const payload = buildBackupPayload([{ key: state.submissionKey, record: { info: state.facultyInfo, values: state.values, evidence: collectEvidence(true), calc: computeAll(state.values, unitId), submittedAt: state.lastSubmittedRecord ? state.lastSubmittedRecord.submittedAt : new Date().toISOString() } }]);
          downloadJSON("ostim-taslak-" + slugify(state.facultyInfo.adSoyad) + ".json", payload);
          toast("Taslak indirildi — daha sonra bu dosyayı yükleyerek kaldığınız yerden devam edebilirsiniz.");
        },
      }, ["💾 İlerlemeyi Yedekle (.json)"]),
      el("button", {
        class: "btn btn-ghost", style: "background:transparent;color:#EFE6D2;border-color:rgba(239,230,210,.3)",
        onclick: () => downloadWordDoc({ info: state.facultyInfo, values: state.values, evidence: collectEvidence(true), calc: computeAll(state.values, unitId) }),
      }, ["📄 İmza İçin Word Belgesi"]),
    ]),
  ]);

  return el("div", { class: "view" }, [
    el("div", { class: "detail-header" }, [
      el("div", {}, [
        el("h2", { style: "font-family:var(--display);font-size:24px;color:var(--navy);margin:0" }, [state.editingExisting ? "Formu Düzenle" : "Faaliyet Bilgilerini Girin"]),
        el("div", { class: "chip-row" }, [
          el("span", { class: "chip" }, [state.facultyInfo.fakulte]),
          el("span", { class: "chip" }, [state.facultyInfo.bolum]),
          el("span", { class: "chip" }, [state.facultyInfo.donem]),
        ]),
      ]),
      el("button", { class: "btn btn-ghost", onclick: () => setView("faculty-info") }, ["← Bilgileri Düzenle"]),
    ]),
    legend,
    el("div", { class: "form-layout" }, [
      el("div", {}, sectionsEl),
      summary,
    ]),
    errBox, okBox,
  ]);
}

function renderSectionBlock(sNo, calc, unitId) {
  const meta = SECTIONS[sNo];
  const items = CRITERIA.filter((i) => i.section === sNo);
  const groups = [];
  const seen = {};
  items.forEach((i) => { if (!seen[i.group]) { seen[i.group] = true; groups.push(i.group); } });

  const head = el("div", { class: "section-head" + (openSections[sNo] ? " open" : ""), onclick: () => { openSections[sNo] = !openSections[sNo]; render(); } }, [
    el("div", { class: "s-left" }, [
      el("span", { class: "s-no" }, ["Bölüm " + sNo]),
      el("h3", {}, [meta.title]),
      el("span", { class: "s-weight" }, ["(ağırlık %" + String(meta.weight).replace(".", ",") + ")"]),
    ]),
    el("div", { style: "display:flex;align-items:center;gap:14px" }, [
      el("span", { class: "s-total" }, [fmt(calc.bySection[sNo], 1) + " puan"]),
      el("span", { class: "chev" }, ["▸"]),
    ]),
  ]);

  const body = el("div", { class: "section-body" + (openSections[sNo] ? " open" : "") },
    groups.map((g) => renderGroup(g, items.filter((i) => i.group === g), calc, unitId))
  );

  return el("div", { class: "section-block" }, [head, body]);
}

function renderGroup(groupKey, items, calc, unitId) {
  const title = items[0].groupTitle || groupKey;
  return el("div", { class: "group-block" }, [
    el("div", { class: "group-title" }, [groupKey + " — " + title]),
    ...items.map((item) => renderCritRow(item, calc, unitId)),
  ]);
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function ensureEvidence(id) {
  if (!state.evidence[id]) state.evidence[id] = { text: "", images: [] };
  return state.evidence[id];
}

function buildEvidencePanel(item) {
  const ev = ensureEvidence(item.id);
  const textarea = el("textarea", {
    rows: "3",
    placeholder: "Bu madde için kanıt / ayrıntı metni (örn. tam kaynakça, DOI, sözleşme no, açıklama)",
    oninput: (e) => { ev.text = e.target.value; },
  }, [ev.text || ""]);

  const thumbWrap = el("div", { class: "evidence-thumbs" }, []);
  function renderThumbs() {
    thumbWrap.innerHTML = "";
    ev.images.forEach((img, idx) => {
      thumbWrap.appendChild(el("div", { class: "evidence-thumb" }, [
        el("img", { src: img.dataUrl, alt: img.name }),
        el("button", { class: "evidence-thumb-remove", onclick: () => { ev.images.splice(idx, 1); renderThumbs(); } }, ["×"]),
      ]));
    });
  }
  renderThumbs();

  const fileInput = el("input", { type: "file", accept: "image/*", multiple: "multiple", style: "display:none" });
  fileInput.addEventListener("change", async (e) => {
    const files = [...e.target.files];
    for (const f of files) {
      try { ev.images.push({ name: f.name, dataUrl: await fileToDataUrl(f) }); } catch (err) {}
    }
    renderThumbs();
    fileInput.value = "";
  });
  const addImgBtn = el("button", { class: "icon-btn", style: "font-size:11px", onclick: () => fileInput.click() }, ["📎 Görsel Ekle"]);

  return el("div", { class: "evidence-panel" }, [
    el("div", { class: "evidence-label" }, ["Kanıt / Ayrıntı (isteğe bağlı — Word belgesi oluşturulurken bu maddenin altına eklenir)"]),
    textarea,
    el("div", { style: "display:flex;gap:8px;align-items:center;margin-top:6px" }, [addImgBtn, fileInput]),
    thumbWrap,
  ]);
}

function buildEvidenceToggle(item) {
  const ev = ensureEvidence(item.id);
  const hasContent = !!((ev.text && ev.text.trim()) || (ev.images && ev.images.length));
  const panel = buildEvidencePanel(item);
  const btn = el("button", { class: "evidence-toggle-btn" + (hasContent ? " has-content" : "") }, [hasContent ? "✓ Kanıt" : "+ Kanıt"]);
  btn.addEventListener("click", () => {
    panel.classList.toggle("open");
    const nowHas = !!((ev.text && ev.text.trim()) || (ev.images && ev.images.length));
    btn.textContent = panel.classList.contains("open") ? "− Kanıt" : (nowHas ? "✓ Kanıt" : "+ Kanıt");
  });
  return { btn, panel };
}

function renderCritRow(item, calc, unitId) {
  const { btn: evBtn, panel: evPanel } = buildEvidenceToggle(item);
  const puanValue = item[unitById(unitId).scoreKey] || 0;

  if (item.author_row) {
    const v = state.values[item.id] || {};
    const cells = AUTHOR_LABELS.map((lab, idx) => {
      const key = "a" + (idx + 1);
      const inp = el("input", {
        type: "number", min: "0", step: "1", value: v[key] || "",
        placeholder: "0",
        oninput: (e) => {
          const val = e.target.value;
          if (!state.values[item.id]) state.values[item.id] = {};
          state.values[item.id][key] = val;
          rerenderScoresOnly(unitId);
        },
      });
      return el("div", { class: "author-cell" }, [el("label", {}, [lab]), inp]);
    });
    const rowEl = el("div", { class: "crit-row author-row", "data-item": item.id }, [
      el("div", { class: "crit-tag" }, [item.id]),
      el("div", {}, [
        el("div", { class: "crit-text" }, [item.text]),
        el("div", { class: "crit-meta" }, ["Puanlama: " + puanValue + " · çok yazarlı — yazar sırasına göre giriniz"]),
        el("div", { class: "author-grid", style: "margin-top:8px" }, cells),
      ]),
      el("div", { class: "crit-input-wrap", style: "align-self:flex-start;flex-direction:column;align-items:flex-end;gap:6px" }, [
        el("span", { class: "crit-score", id: "score-" + item.id }, [fmt(computeItemScore(item, state.values[item.id], unitId), 1)]),
        evBtn,
      ]),
    ]);
    return el("div", {}, [rowEl, evPanel]);
  }

  const val = state.values[item.id] || "";
  const isRating = !!item.maxValue;
  const inp = el("input", {
    type: "number", min: "0", step: isRating ? "0.1" : "1",
    max: isRating ? String(item.maxValue) : null,
    value: val, placeholder: "0",
    oninput: (e) => {
      let v = e.target.value;
      if (isRating && Number(v) > item.maxValue) v = String(item.maxValue);
      state.values[item.id] = v;
      rerenderScoresOnly(unitId);
    },
  });
  const rowEl = el("div", { class: "crit-row", "data-item": item.id }, [
    el("div", { class: "crit-tag" }, [item.id]),
    el("div", {}, [
      el("div", { class: "crit-text" }, [item.text]),
      el("div", { class: "crit-meta" }, [
        (isRating ? "Puan (0-" + item.maxValue + ")" : "Puanlama: " + puanValue) + (item.note ? " · " + item.note : ""),
      ]),
    ]),
    el("div", { class: "crit-input-wrap" }, [
      inp,
      el("span", { class: "crit-score", id: "score-" + item.id }, [fmt(computeItemScore(item, val, unitId), 1)]),
      evBtn,
    ]),
  ]);
  return el("div", {}, [rowEl, evPanel]);
}

// Lightweight partial re-render: update only score labels + summary + section totals,
// without rebuilding every input (keeps focus while typing).
function rerenderScoresOnly(unitId) {
  const calc = computeAll(state.values, unitId);
  CRITERIA.forEach((item) => {
    const elScore = document.getElementById("score-" + item.id);
    if (elScore) elScore.textContent = fmt(computeItemScore(item, state.values[item.id], unitId), 1);
  });
  const grandEl = document.querySelector(".summary-panel .grand");
  if (grandEl) grandEl.textContent = fmt(calc.grand, 2);
  const rows = document.querySelectorAll(".summary-row");
  const order = [1, 2, 3, 4];
  rows.forEach((row, idx) => {
    const v = row.querySelector(".v");
    if (v && order[idx]) v.textContent = fmt(calc.bySection[order[idx]], 1);
  });
  document.querySelectorAll(".section-head").forEach((head, idx) => {
    const sNo = idx + 1;
    const totalEl = head.querySelector(".s-total");
    if (totalEl) totalEl.textContent = fmt(calc.bySection[sNo], 1) + " puan";
  });
}

/* --------------------------- Faculty: bitiş ------------------------------ */

function viewFacultyDone() {
  const cloudNotice = STORAGE_MODE === "cloud" ? el("div", { class: "legend-item", style: "background:var(--green-soft);margin:18px 0;text-align:left" }, [
    el("div", { class: "legend-swatch", style: "background:var(--green)" }),
    el("div", {}, [
      el("b", {}, ["Bir kopyası size (" + state.facultyInfo.email + "), bilgi olarak da Mühendislik Fakültesi Dekanlığına (mf.dekanlik@ostimteknik.edu.tr) gönderildi. "]),
      "E-postayı birkaç dakika içinde alamazsanız spam/gereksiz klasörünü kontrol edin.",
    ]),
  ]) : null;

  const localNotice = STORAGE_MODE !== "cloud" ? el("div", { class: "legend-item", style: "background:var(--orange-soft);margin:18px 0;text-align:left" }, [
    el("div", { class: "legend-swatch", style: "background:var(--orange)" }),
    el("div", {}, [
      el("b", {}, ["Bu form yalnızca bu tarayıcıya kaydedildi. "]),
      "Bu bir sunucuya bağlı değil, bu yüzden yöneticiye otomatik ulaşmaz. Aşağıdaki “Kayıt Dosyasını İndir” butonuyla bir .json dosyası indirip yöneticiye (e-posta ile) iletin; yönetici bu dosyayı panelden içe aktaracaktır.",
    ]),
  ]) : null;

  const downloadBtn = STORAGE_MODE !== "cloud" ? el("button", {
    class: "btn btn-gold",
    onclick: () => {
      const payload = buildBackupPayload([{ key: state.submissionKey, record: state.lastSubmittedRecord }]);
      downloadJSON("ostim-form-" + slugify(state.facultyInfo.adSoyad) + ".json", payload);
    },
  }, ["Kayıt Dosyasını İndir (.json)"]) : null;

  const backupBtn = el("button", {
    class: "btn btn-ghost",
    onclick: () => {
      const payload = buildBackupPayload([{ key: state.submissionKey, record: {
        info: state.facultyInfo, values: state.values, evidence: collectEvidence(true),
        calc: computeAll(state.values, state.facultyInfo.birim),
        submittedAt: state.lastSubmittedRecord ? state.lastSubmittedRecord.submittedAt : new Date().toISOString(),
      } }]);
      downloadJSON("ostim-taslak-" + slugify(state.facultyInfo.adSoyad) + ".json", payload);
      toast("Yedek indirildi.");
    },
  }, ["💾 Yedek Olarak İndir (.json)"]);

  const wordBtn = el("button", {
    class: "btn btn-primary",
    onclick: () => downloadWordDoc({ info: state.facultyInfo, values: state.values, evidence: collectEvidence(true), calc: computeAll(state.values, state.facultyInfo.birim) }),
  }, ["📄 İmza İçin Word Belgesi İndir"]);

  return el("div", { class: "view view-narrow" }, [
    el("div", { class: "card", style: "text-align:center;padding:48px 30px" }, [
      el("div", { style: "font-size:38px;margin-bottom:10px" }, ["✓"]),
      el("h2", {}, ["Form Kaydedildi"]),
      el("div", { class: "sub" }, [state.facultyInfo.adSoyad + " için " + state.facultyInfo.donem + " dönemi değerlendirmesi kaydedildi. Dilerseniz formu tekrar açıp güncelleyebilirsiniz."]),
      cloudNotice,
      localNotice,
      el("div", { class: "form-actions", style: "justify-content:center;flex-wrap:wrap" }, [
        wordBtn,
        backupBtn,
        downloadBtn,
        el("button", { class: "btn btn-ghost", onclick: () => setView("faculty-form") }, ["Formu Aç / Düzenle"]),
        el("button", { class: "btn btn-ghost", onclick: () => setView("landing") }, ["Ana Sayfaya Dön"]),
      ].filter(Boolean)),
    ]),
  ]);
}

/* --------------------------- Admin: giriş -------------------------------- */

function viewAdminLogin() {
  const errBox = el("div", { class: "error-msg hidden" }, []);
  const modeNote = el("div", { class: "hint", style: "margin-bottom:14px" }, [
    STORAGE_MODE === "cloud"
      ? "Merkezi yönetici şifresi tüm birimleri görür; her birimin kendi şifresi yalnızca kendi birimini görür. Şifreler Google E-Tablo'daki “Settings” sayfasında saklanır."
      : "⚠ Apps Script bağlı değil (yerel mod). Bu modda yalnızca tek bir (merkezi) yönetici şifresi geçerlidir: config.js/app.js içindeki LOCAL_ADMIN_PASSWORD.",
  ]);
  const pw = el("input", { type: "password", placeholder: "Yönetici şifresi" });
  const tryLogin = async () => {
    if (STORAGE_MODE === "cloud") {
      state.adminPassword = pw.value;
      const res = await loadAdminData();
      if (!res.ok) {
        state.adminPassword = null;
        errBox.textContent = res.error === "network"
          ? "Sunucuya ulaşılamadı. Apps Script URL'sini ve internet bağlantınızı kontrol edin."
          : "Şifre hatalı. Lütfen tekrar deneyin.";
        errBox.classList.remove("hidden");
        return;
      }
      state.adminAuthed = true;
      setView("admin-dashboard");
    } else {
      if (pw.value !== LOCAL_ADMIN_PASSWORD) {
        errBox.textContent = "Şifre hatalı. Lütfen tekrar deneyin.";
        errBox.classList.remove("hidden");
        return;
      }
      state.adminAuthed = true;
      state.adminPassword = pw.value;
      await loadAdminData();
      setView("admin-dashboard");
    }
  };
  pw.addEventListener("keydown", (e) => { if (e.key === "Enter") tryLogin(); });
  return el("div", { class: "view" }, [
    el("div", { class: "card admin-login-box" }, [
      el("h2", {}, ["Yönetici Girişi"]),
      el("div", { class: "sub" }, ["Tüm öğretim elemanı sonuçlarını görüntülemek ve istatistikleri incelemek için şifrenizi girin."]),
      modeNote,
      el("div", { class: "field" }, [el("label", {}, ["Şifre"]), pw]),
      errBox,
      el("div", { class: "form-actions" }, [
        el("button", { class: "btn btn-primary", onclick: tryLogin }, ["Giriş Yap"]),
      ]),
    ]),
  ]);
}

/* --------------------------- Admin: veri yükleme -------------------------- */

async function loadAdminData() {
  let subs = [];
  let thresholds = {};
  let policy = {};
  let role = "central";
  let birim = null;
  if (STORAGE_MODE === "cloud") {
    const r = await apiCall("list", { password: state.adminPassword });
    if (!r || !r.ok) return { ok: false, error: (r && r.error) || "network" };
    role = r.role;
    birim = r.birim || null;
    thresholds = r.thresholds || {};
    policy = r.policy || {};
    subs = r.items.map((it) => {
      const b = (it.info && it.info.birim) || "muhendislik";
      return {
        key: it.key, info: it.info || {}, values: it.values || {}, evidence: it.evidence || {},
        calc: computeEvaluation(it.values || {}, b, policy[b]), rektorluk: it.rektorluk, submittedAt: it.submittedAt,
      };
    });
  } else {
    const keys = localList("submission:", true);
    for (const k of keys) {
      const raw = localGet(k, true);
      if (!raw) continue;
      try {
        const parsed = JSON.parse(raw);
        const b = (parsed.info && parsed.info.birim) || "muhendislik";
        subs.push({ key: k, info: parsed.info || {}, values: parsed.values || {}, evidence: parsed.evidence || {}, calc: computeEvaluation(parsed.values || {}, b, null), rektorluk: parsed.rektorluk, submittedAt: parsed.submittedAt });
      } catch (e) {}
    }
    UNITS.forEach((u) => {
      const thr = localGet("settings:threshold:" + u.id, true);
      thresholds[u.id] = thr || "";
      const polRaw = localGet("settings:policy:" + u.id, true);
      try { policy[u.id] = polRaw ? JSON.parse(polRaw) : {}; } catch (e) { policy[u.id] = {}; }
    });
    // yerel modda kayıtları politika ile yeniden hesapla (yukarıda policy henüz yokken hesaplanmıştı)
    subs.forEach((s) => {
      const b = s.info.birim || "muhendislik";
      s.calc = computeEvaluation(s.values, b, policy[b]);
    });
  }
  subs.sort((a, b) => b.calc.adjustedGrand - a.calc.adjustedGrand);
  state.adminSubmissions = subs;
  state.adminThresholds = thresholds;
  state.adminPolicy = policy;
  state.adminRole = role;
  state.adminBirim = birim;
  if (!state.adminFilterBirim) state.adminFilterBirim = "all";
  return { ok: true };
}

/* --------------------------- Admin: dashboard ----------------------------- */

function viewAdminDashboard() {
  const allSubs = state.adminSubmissions;
  const isCentral = state.adminRole === "central";
  const scopeUnits = isCentral ? UNITS : [unitById(state.adminBirim)];

  // Her birim için (filtreden bağımsız) kendi kayıtları — eşik/varsayılan ortalama hesapları için
  const subsByUnit = {};
  UNITS.forEach((u) => { subsByUnit[u.id] = allSubs.filter((s) => (s.info.birim || "muhendislik") === u.id); });
  function unitAvg(unitId) {
    const totals = subsByUnit[unitId].map((s) => s.calc.adjustedGrand);
    return totals.length ? totals.reduce((a, b) => a + b, 0) / totals.length : 0;
  }
  function effectiveThreshold(unitId) {
    const raw = state.adminThresholds[unitId];
    if (raw !== undefined && raw !== null && raw !== "") return Number(raw);
    return Math.round(unitAvg(unitId));
  }

  // Filtre (yalnızca merkezi yönetici için) — hangi birim(ler) tabloda/grafikte gösterilsin
  const filterSel = isCentral ? el("select", {}, [
    el("option", { value: "all", selected: state.adminFilterBirim === "all" ? "selected" : null }, ["Tüm Birimler"]),
    ...UNITS.map((u) => el("option", { value: u.id, selected: state.adminFilterBirim === u.id ? "selected" : null }, [u.label])),
  ]) : null;
  if (filterSel) filterSel.addEventListener("change", () => { state.adminFilterBirim = filterSel.value; render(); });

  const subs = (!isCentral || state.adminFilterBirim === "all") ? (isCentral ? allSubs : subsByUnit[state.adminBirim]) : subsByUnit[state.adminFilterBirim];

  const totals = subs.map((s) => s.calc.adjustedGrand);
  const n = totals.length;
  const avg = n ? totals.reduce((a, b) => a + b, 0) / n : 0;
  const sortedTotals = [...totals].sort((a, b) => a - b);
  const median = n ? (n % 2 ? sortedTotals[(n - 1) / 2] : (sortedTotals[n / 2 - 1] + sortedTotals[n / 2]) / 2) : 0;
  const max = n ? Math.max(...totals) : 0;

  const meetsCount = subs.filter((s) => s.calc.meetsMinimums && s.calc.adjustedGrand >= effectiveThreshold(s.info.birim || "muhendislik")).length;
  const belowMinCount = subs.filter((s) => !s.calc.meetsMinimums).length;

  const statCards = el("div", { class: "stat-grid" }, [
    statCard("Toplam Kayıt", n, ""),
    statCard("Ortalama Puan", fmt(avg, 1), "gold"),
    statCard("Medyan Puan", fmt(median, 1), ""),
    statCard("Kriterleri Karşılayan", meetsCount + " / " + n, "green"),
  ]);

  // Birim bazlı ayarlar: eşik + asgari I. Bölüm şartı + III+IV tavanı — hepsi tek satırda, tek Kaydet ile
  function previewStats(unitId, thrVal, minPuanVal, minAdetVal, capVal) {
    const list = subsByUnit[unitId];
    const n = list.length;
    if (!n) return { n: 0, meetsMin: 0, aboveThr: 0 };
    const policy = { minSec1Puan: minPuanVal, minSec1Adet: minAdetVal, capIIIIVPercent: capVal };
    const thr = thrVal === "" || thrVal === undefined ? Math.round(unitAvg(unitId)) : Number(thrVal);
    let meetsMin = 0, aboveThr = 0;
    list.forEach((s) => {
      const ev = computeEvaluation(s.values, unitId, policy);
      if (ev.meetsMinimums) meetsMin++;
      if (ev.meetsMinimums && ev.adjustedGrand >= thr) aboveThr++;
    });
    return { n, meetsMin, aboveThr };
  }

  const unitSettingsRows = scopeUnits.map((u) => {
    const pol = state.adminPolicy[u.id] || {};
    const thrInput = el("input", { type: "number", step: "1", value: state.adminThresholds[u.id] || "", placeholder: "örn. " + Math.round(unitAvg(u.id)) });
    const minPuanInput = el("input", { type: "number", step: "1", value: pol.minSec1Puan || "", placeholder: "sınırsız" });
    const minAdetInput = el("input", { type: "number", step: "1", value: pol.minSec1Adet || "", placeholder: "sınırsız" });
    const capInput = el("input", { type: "number", step: "1", min: "0", max: "100", value: pol.capIIIIVPercent || "", placeholder: "sınırsız" });
    const previewBox = el("div", { style: "grid-column:1/-1;font-size:11.5px;color:var(--navy-2);margin-top:2px;font-weight:600" }, []);
    const updatePreview = () => {
      const st = previewStats(u.id, thrInput.value, minPuanInput.value, minAdetInput.value, capInput.value);
      previewBox.textContent = st.n === 0
        ? "Bu birimde henüz kayıt yok — önizleme yapılamıyor."
        : "Bu değerlerle: " + st.meetsMin + "/" + st.n + " kişi asgari şartı karşılıyor (%" + Math.round(100 * st.meetsMin / st.n) + ") · " +
          st.aboveThr + "/" + st.n + " kişi “Kriterleri Karşılıyor” olur (%" + Math.round(100 * st.aboveThr / st.n) + ").";
    };
    [thrInput, minPuanInput, minAdetInput, capInput].forEach((inp) => inp.addEventListener("input", updatePreview));
    const saveBtn = el("button", { class: "btn btn-primary", style: "padding:7px 14px;font-size:12.5px", onclick: async () => {
      const payload = {
        threshold: thrInput.value === "" ? "" : (Number(thrInput.value) || 0),
        minSec1Puan: minPuanInput.value === "" ? "" : (Number(minPuanInput.value) || 0),
        minSec1Adet: minAdetInput.value === "" ? "" : (Number(minAdetInput.value) || 0),
        capIIIIVPercent: capInput.value === "" ? "" : (Number(capInput.value) || 0),
      };
      if (STORAGE_MODE === "cloud") {
        await apiCall("setUnitSettings", Object.assign({ password: state.adminPassword, birim: u.id }, payload));
      } else {
        localSet("settings:threshold:" + u.id, String(payload.threshold), true);
        localSet("settings:policy:" + u.id, JSON.stringify({ minSec1Puan: payload.minSec1Puan, minSec1Adet: payload.minSec1Adet, capIIIIVPercent: payload.capIIIIVPercent }), true);
      }
      state.adminThresholds[u.id] = payload.threshold;
      state.adminPolicy[u.id] = { minSec1Puan: payload.minSec1Puan, minSec1Adet: payload.minSec1Adet, capIIIIVPercent: payload.capIIIIVPercent };
      await loadAdminData();
      render();
      toast(u.label + " ayarları kaydedildi.");
    } }, ["Kaydet"]);
    updatePreview();
    return el("div", { style: "padding:14px 0;border-bottom:1px solid var(--line)" }, [
      el("div", { style: "font-weight:600;color:var(--navy);margin-bottom:2px" }, [u.label]),
      el("div", { style: "font-size:11px;color:var(--ink-soft);margin-bottom:10px" }, [subsByUnit[u.id].length + " kayıt · ham ortalama " + fmt(unitAvg(u.id), 1)]),
      el("div", { style: "display:grid;grid-template-columns:repeat(4,1fr) auto;gap:10px;align-items:end" }, [
        el("div", { class: "field" }, [el("label", {}, ["Asgari Performans Eşiği"]), thrInput]),
        el("div", { class: "field" }, [el("label", {}, ["Asgari I. Bölüm Puanı"]), minPuanInput]),
        el("div", { class: "field" }, [el("label", {}, ["Asgari I. Bölüm Faaliyet Adedi"]), minAdetInput]),
        el("div", { class: "field" }, [el("label", {}, ["III+IV Tavanı (%, I+II'ye göre)"]), capInput]),
        saveBtn,
      ]),
      previewBox,
    ]);
  });
  const thresholdBar = el("div", { class: "card" }, [
    el("h2", {}, ["Birim Bazında Değerlendirme Ayarları"]),
    el("div", { class: "sub" }, [
      "Her birimin puan ölçeği farklı olabileceğinden bu ayarlar birim başına ayrı belirlenir. Bir değer girerken altında, o an sistemde kayıtlı verilere göre kaç kişinin bu şartı karşılayacağı canlı olarak gösterilir — kaydetmeden önce etkisini görebilirsiniz. “Asgari I. Bölüm Puanı” ve “Asgari I. Bölüm Faaliyet Adedi” birlikte tanımlandığında VEYA mantığıyla çalışır — ikisinden biri karşılanırsa yeterlidir (örn. tek ama yüksek puanlı bir yayın, az sayıda madde şartını telafi edebilir). Yalnızca biri girilirse, o tek şart geçerli olur. “III+IV Tavanı”, Hizmet+Tanıtım toplamının Araştırma+Proje toplamına oranla en fazla belirtilen yüzde kadar sayılmasını sağlar. Boş bırakılan alanlar sınırsızdır.",
    ]),
    el("div", { style: "margin-top:14px" }, unitSettingsRows),
  ]);

  const barsSection = n === 0 ? null : el("div", { class: "card" }, [
    el("h2", {}, ["Puan Dağılımı"]),
    el("div", { class: "sub" }, ["Değerlendirme toplamına (politika uygulanmış) göre sıralanmış tüm kayıtlar."]),
    el("div", { class: "bars" }, subs.map((s) => {
      const pct = max > 0 ? Math.max(4, (s.calc.adjustedGrand / max) * 100) : 0;
      return el("div", { class: "bar-row" }, [
        el("div", { style: "overflow:hidden;text-overflow:ellipsis;white-space:nowrap" }, [s.info.adSoyad || "—"]),
        el("div", { class: "bar-track" }, [el("div", { class: "bar-fill", style: "width:" + pct + "%" })]),
        el("div", { style: "font-family:var(--mono);text-align:right" }, [fmt(s.calc.adjustedGrand, 0)]),
      ]);
    })),
  ]);

  const sectionAvg = [1, 2, 3, 4].map((sNo) => {
    const vals = subs.map((s) => s.calc.bySection[sNo]);
    const a = vals.length ? vals.reduce((x, y) => x + y, 0) / vals.length : 0;
    return { sNo, a };
  });
  const sectionAvgCard = n === 0 ? null : el("div", { class: "card" }, [
    el("h2", {}, ["Bölüm Bazında Ortalamalar"]),
    el("div", { class: "bars", style: "margin-top:14px" }, sectionAvg.map(({ sNo, a }) => {
      const maxA = Math.max(...sectionAvg.map((x) => x.a), 1);
      const pct = Math.max(4, (a / maxA) * 100);
      return el("div", { class: "bar-row" }, [
        el("div", {}, [SECTIONS[sNo].title]),
        el("div", { class: "bar-track" }, [el("div", { class: "bar-fill", style: "width:" + pct + "%" })]),
        el("div", { style: "font-family:var(--mono);text-align:right" }, [fmt(a, 1)]),
      ]);
    })),
  ]);

  const exportBtn = el("button", { class: "btn btn-ghost", onclick: () => exportCSV(subs) }, ["CSV Olarak Dışa Aktar"]);
  const backupBtn = el("button", { class: "btn btn-ghost", onclick: () => {
    const payload = buildBackupPayload(subs.map((s) => ({ key: s.key, record: { info: s.info, values: s.values, calc: s.calc, submittedAt: s.submittedAt, rektorluk: s.rektorluk } })));
    downloadJSON("ostim-performans-yedek.json", payload);
  } }, ["Tüm Verileri Yedekle (JSON)"]);
  const importInput = el("input", { type: "file", accept: "application/json", multiple: "multiple", style: "display:none" });
  importInput.addEventListener("change", async (e) => {
    const files = [...e.target.files];
    let count = 0;
    for (const file of files) {
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        const list = parsed.submissions ? parsed.submissions : (parsed.info ? [{ key: "submission:" + (parsed.info.birim || "muhendislik") + "__" + slugify(parsed.info.adSoyad) + "__" + slugify(parsed.info.donem), record: parsed }] : []);
        for (const entry of list) {
          await putRecord(entry.key, entry.record);
          count++;
        }
      } catch (err) { console.error("İçe aktarma hatası", err); }
    }
    await loadAdminData();
    render();
    toast(count + " kayıt içe aktarıldı.");
  });
  const importBtn = el("button", { class: "btn btn-ghost", onclick: () => importInput.click() }, ["Kayıt Dosyalarını İçe Aktar"]);
  const refreshBtn = el("button", { class: "btn btn-ghost", onclick: async () => { await loadAdminData(); render(); toast("Liste güncellendi."); } }, ["Yenile"]);
  const logoutBtn = el("button", { class: "btn btn-ghost", onclick: () => { state.adminAuthed = false; setView("landing"); } }, ["Çıkış Yap"]);

  const showBirimCol = isCentral;
  const headers = ["Ad Soyad"].concat(showBirimCol ? ["Birim"] : []).concat(["Unvan", "Bölüm", "Dönem", "I. Bil. Araş.", "II. Proje", "III. Hizmet", "IV. Tanıtım", "Toplam", "Durum", ""]);
  const table = n === 0 ? el("div", { class: "empty-state" }, [
    el("div", { class: "glyph" }, ["◌"]),
    el("div", {}, ["Henüz kaydedilmiş bir form bulunmuyor."]),
  ]) : el("table", { class: "data-table" }, [
    el("thead", {}, [el("tr", {}, headers.map((h) => el("th", {}, [h])))]),
    el("tbody", {}, subs.map((s) => renderAdminRow(s, effectiveThreshold(s.info.birim || "muhendislik"), showBirimCol))),
  ]);

  return el("div", { class: "view" }, [
    el("div", { class: "detail-header" }, [
      el("div", {}, [
        el("h2", { style: "font-family:var(--display);font-size:24px;color:var(--navy);margin:0" }, ["Yönetici Paneli"]),
        el("div", { class: "sub" }, [
          isCentral ? "Merkezi yönetici — tüm birimlerin performans verilerini görüntüleyin ve karşılaştırın."
            : unitById(state.adminBirim).label + " — yalnızca bu birimin verileri.",
        ]),
      ]),
      el("div", { style: "display:flex;gap:8px;flex-wrap:wrap;align-items:center" }, [filterSel, refreshBtn, importBtn, importInput, backupBtn, exportBtn, logoutBtn].filter(Boolean)),
    ]),
    statCards,
    thresholdBar,
    el("div", { class: "card" }, [
      el("h2", {}, ["Kayıtlar"]),
      el("div", { class: "sub" }, [n + " öğretim elemanı kaydı listeleniyor."]),
      el("div", { style: "overflow-x:auto" }, [table]),
    ]),
    barsSection,
    sectionAvgCard,
  ].filter(Boolean));
}

function statCard(label, value, cls) {
  return el("div", { class: "stat-card" }, [
    el("div", { class: "k" }, [label]),
    el("div", { class: "v" + (cls ? " " + cls : "") }, [String(value)]),
  ]);
}

function renderAdminRow(s, threshold, showBirimCol) {
  const meetsMin = s.calc.meetsMinimums;
  const above = meetsMin && s.calc.adjustedGrand >= threshold;
  let badgeClass = "badge-below";
  let badgeText = "Eşik Altı";
  if (!meetsMin) { badgeClass = "badge-warn"; badgeText = "Asgari Şart Eksik"; }
  else if (above) { badgeClass = "badge-above"; badgeText = "Kriterleri Karşılıyor"; }
  const totalCell = s.calc.capped
    ? el("span", { title: "Ham toplam " + fmt(s.calc.rawGrand, 1) + " idi; III+IV tavanı nedeniyle " + fmt(s.calc.adjustedGrand, 1) + "'e düşürüldü." }, [fmt(s.calc.adjustedGrand, 1) + " †"])
    : fmt(s.calc.adjustedGrand, 1);
  return el("tr", {}, [
    el("td", { class: "name-cell" }, [s.info.adSoyad || "—"]),
    showBirimCol ? el("td", {}, [unitById(s.info.birim || "muhendislik").label]) : null,
    el("td", {}, [s.info.unvan || "—"]),
    el("td", {}, [s.info.bolum || "—"]),
    el("td", {}, [s.info.donem || "—"]),
    el("td", { style: "font-family:var(--mono)" }, [fmt(s.calc.bySection[1], 1)]),
    el("td", { style: "font-family:var(--mono)" }, [fmt(s.calc.bySection[2], 1)]),
    el("td", { style: "font-family:var(--mono)" }, [fmt(s.calc.bySection[3], 1)]),
    el("td", { style: "font-family:var(--mono)" }, [fmt(s.calc.bySection[4], 1)]),
    el("td", { style: "font-family:var(--mono);font-weight:700;color:var(--navy)" }, [totalCell]),
    el("td", {}, [el("span", { class: "badge " + badgeClass }, [badgeText])]),
    el("td", {}, [el("div", { class: "row-actions" }, [
      el("button", { class: "icon-btn", onclick: () => { state.adminDetailKey = s.key; setView("admin-detail"); } }, ["Görüntüle"]),
      el("button", { class: "icon-btn danger", onclick: async () => {
        if (confirm(s.info.adSoyad + " kaydını silmek istediğinize emin misiniz?")) {
          if (STORAGE_MODE === "cloud") await apiCall("delete", { password: state.adminPassword, key: s.key });
          else localDelete(s.key, true);
          await loadAdminData();
          render();
          toast("Kayıt silindi.");
        }
      } }, ["Sil"]),
    ])]),
  ].filter(Boolean));
}

function exportCSV(subs) {
  const header = ["Ad Soyad", "Birim", "E-posta", "Unvan", "Bölüm", "Dönem", "I. Bilimsel Araştırma", "II. Proje", "III. Hizmet", "IV. Tanıtım", "Ham Toplam", "Değerlendirme Toplamı", "Asgari Şart", "Rektörlük Görüşü"];
  const rows = subs.map((s) => [
    s.info.adSoyad, unitById(s.info.birim || "muhendislik").label, s.info.email || "", s.info.unvan, s.info.bolum, s.info.donem,
    fmt(s.calc.bySection[1], 2), fmt(s.calc.bySection[2], 2), fmt(s.calc.bySection[3], 2), fmt(s.calc.bySection[4], 2),
    fmt(s.calc.rawGrand, 2), fmt(s.calc.adjustedGrand, 2), s.calc.meetsMinimums ? "Karşılıyor" : "Karşılamıyor", s.rektorluk || "",
  ]);
  const csv = [header, ...rows].map((r) => r.map((c) => '"' + String(c || "").replace(/"/g, '""') + '"').join(";")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "ostim-performans-degerlendirme.csv";
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

/* --------------------------- Admin: detay -------------------------------- */

function viewAdminDetail() {
  const s = state.adminSubmissions.find((x) => x.key === state.adminDetailKey);
  if (!s) {
    return el("div", { class: "view" }, [el("div", { class: "empty-state" }, ["Kayıt bulunamadı."])]);
  }
  const unitId = s.info.birim || "muhendislik";
  const rektInput = el("input", { type: "number", min: "0", max: "10", step: "1", value: s.rektorluk || "" });
  const saveRekt = async () => {
    if (STORAGE_MODE === "cloud") {
      await apiCall("setRektorluk", { password: state.adminPassword, key: s.key, value: rektInput.value });
    } else {
      const raw = localGet(s.key, true);
      let parsed = {};
      try { parsed = raw ? JSON.parse(raw) : {}; } catch (e) {}
      parsed.rektorluk = rektInput.value;
      localSet(s.key, JSON.stringify(parsed), true);
    }
    s.rektorluk = rektInput.value;
    toast("Rektörlük görüşü kaydedildi.");
  };

  const wordBtn = el("button", {
    class: "btn btn-primary",
    onclick: () => downloadWordDoc({ info: s.info, values: s.values, evidence: s.evidence || {}, calc: s.calc }),
  }, ["📄 İmza İçin Word Belgesi İndir"]);

  const sectionsView = [1, 2, 3, 4].map((sNo) => {
    const items = CRITERIA.filter((i) => i.section === sNo && (computeItemScore(i, s.values[i.id], unitId) > 0));
    if (items.length === 0) return null;
    const pct = s.calc.rawGrand > 0 ? (s.calc.bySection[sNo] / s.calc.rawGrand) * 100 : 0;
    return el("div", { class: "section-block" }, [
      el("div", { class: "section-head open", style: "cursor:default" }, [
        el("div", { class: "s-left" }, [
          el("span", { class: "s-no" }, ["Bölüm " + sNo]),
          el("h3", {}, [SECTIONS[sNo].title]),
        ]),
        el("span", { class: "s-total" }, [fmt(s.calc.bySection[sNo], 1) + " puan  ·  %" + fmt(pct, 1)]),
      ]),
      el("div", { class: "section-body open" }, items.map((item) => {
        const val = s.values[item.id];
        const scoreVal = computeItemScore(item, val, unitId);
        const detailStr = item.author_row
          ? AUTHOR_LABELS.map((lab, idx) => {
              const c = Number((val || {})["a" + (idx + 1)]) || 0;
              return c > 0 ? c + "× " + lab : null;
            }).filter(Boolean).join(", ")
          : String(val) + (item.maxValue ? " / " + item.maxValue : " adet");
        const evText = (s.evidence && s.evidence[item.id] && s.evidence[item.id].text) || "";
        return el("div", { class: "crit-row" + (item.author_row ? " author-row" : "") }, [
          el("div", { class: "crit-tag" }, [item.id]),
          el("div", {}, [
            el("div", { class: "crit-text" }, [item.text]),
            el("div", { class: "crit-meta" }, [detailStr]),
            evText ? el("div", { class: "crit-meta", style: "margin-top:4px;white-space:pre-line;color:var(--ink)" }, ["Kanıt: " + evText]) : null,
          ].filter(Boolean)),
          el("div", { class: "crit-score" }, [fmt(scoreVal, 1)]),
        ]);
      })),
    ]);
  }).filter(Boolean);

  const DIST_COLORS = { 1: "var(--navy)", 2: "var(--navy-2)", 3: "var(--gold)", 4: "var(--orange)" };
  const distBar = el("div", { class: "card", style: "margin-bottom:20px" }, [
    el("h2", {}, ["Puan Dağılımı"]),
    el("div", { class: "sub" }, ["Ham Toplam içindeki bölüm payları — bir bölümün orantısız büyük olması, değerlendirmede dikkat edilmesi gereken bir işaret olabilir."]),
    el("div", { style: "display:flex;height:28px;border-radius:8px;overflow:hidden;margin-top:12px;background:#EFE9D9" },
      [1, 2, 3, 4].map((sNo) => {
        const pct = s.calc.rawGrand > 0 ? (s.calc.bySection[sNo] / s.calc.rawGrand) * 100 : 0;
        if (pct <= 0) return null;
        return el("div", {
          style: "width:" + pct + "%;background:" + DIST_COLORS[sNo] + ";display:flex;align-items:center;justify-content:center",
          title: SECTIONS[sNo].title + ": %" + fmt(pct, 1),
        }, [pct > 7 ? el("span", { style: "color:#fff;font-size:10.5px;font-weight:700" }, ["%" + fmt(pct, 0)]) : null].filter(Boolean));
      }).filter(Boolean)),
    el("div", { style: "display:flex;gap:16px;margin-top:10px;flex-wrap:wrap" },
      [1, 2, 3, 4].map((sNo) => {
        const pct = s.calc.rawGrand > 0 ? (s.calc.bySection[sNo] / s.calc.rawGrand) * 100 : 0;
        return el("div", { style: "display:flex;align-items:center;gap:6px;font-size:11.5px;color:var(--ink-soft)" }, [
          el("span", { style: "width:9px;height:9px;border-radius:2px;background:" + DIST_COLORS[sNo] + ";display:inline-block" }),
          SECTIONS[sNo].title + ": %" + fmt(pct, 1) + " (" + fmt(s.calc.bySection[sNo], 1) + " puan)",
        ]);
      })),
  ]);

  return el("div", { class: "view" }, [
    el("div", { class: "detail-header" }, [
      el("div", {}, [
        el("h2", {}, [s.info.adSoyad]),
        el("div", { class: "chip-row" }, [
          el("span", { class: "chip" }, [unitById(unitId).label]),
          el("span", { class: "chip" }, [s.info.unvan]),
          el("span", { class: "chip" }, [s.info.bolum]),
          el("span", { class: "chip" }, [s.info.donem]),
          el("span", { class: "chip" }, ["Ham Toplam: " + fmt(s.calc.rawGrand, 2)]),
          s.calc.capped ? el("span", { class: "chip", style: "background:var(--orange-soft);color:var(--orange)" }, ["Değerlendirme Toplamı: " + fmt(s.calc.adjustedGrand, 2) + " (III+IV tavanı uygulandı)"]) : null,
          el("span", { class: "chip", style: s.calc.meetsMinimums ? "background:var(--green-soft);color:var(--green)" : "background:var(--orange-soft);color:var(--orange)" }, [s.calc.meetsMinimums ? "Asgari Şart: Karşılıyor" : "Asgari Şart: Karşılamıyor"]),
        ].filter(Boolean)),
      ]),
      el("div", { style: "display:flex;gap:8px" }, [wordBtn, el("button", { class: "btn btn-ghost", onclick: () => setView("admin-dashboard") }, ["← Panele Dön"])]),
    ]),
    distBar,
    el("div", { class: "card", style: "margin-bottom:20px" }, [
      el("h2", {}, ["Rektörlük Görüşü"]),
      el("div", { class: "sub" }, ["0 ile 10 arasında, yalnızca yönetici tarafından girilir. Genel toplam puana dahil edilmez (resmi formdaki tanıma göre bilgi amaçlıdır)."]),
      el("div", { style: "display:flex;gap:10px;align-items:flex-end" }, [
        el("div", { class: "field" }, [el("label", {}, ["Puan (0-10)"]), rektInput]),
        el("button", { class: "btn btn-primary", onclick: saveRekt }, ["Kaydet"]),
      ]),
    ]),
    items_empty_note(items_count(s, unitId)),
    ...sectionsView,
  ]);
}
function items_count(s, unitId){ return CRITERIA.filter(i => computeItemScore(i, s.values[i.id], unitId) > 0).length; }
function items_empty_note(count){
  if (count > 0) return null;
  return el("div", { class: "empty-state" }, ["Bu kayıt için henüz puan üreten bir faaliyet girilmemiş."]);
}

/* --------------------------- Başlangıç ------------------------------------ */

render();
