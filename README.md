# OSTİM Teknik Üniversitesi — Bilimsel Performans Değerlendirme

*Prof. Dr. Demiral AKBAR tarafından geliştirilmiştir · 08.2026*


Bu, OSTİM Teknik Üniversitesi'nin **tüm birimleri** için ortak "Bilimsel Performans
Değerlendirme Formu"nu (Excel/Word tabanlı resmi süreç) dijitalleştiren bir web
uygulamasıdır. Öğretim elemanları faaliyetlerini girer, her birimin yöneticisi
kendi verilerini, merkezi yönetici ise tüm üniversitenin verilerini görebilir.

## Kapsanan birimler

| Birim | Bildirim e-postası | Puanlama tablosu |
|---|---|---|
| Mühendislik Fakültesi | mf.dekanlik@ostimteknik.edu.tr | Kaynak Excel'deki "Müh Fak Puanlaması" sütunu |
| Mimarlık ve Tasarım Fakültesi (MTF) | sare.sahil@ostimteknik.edu.tr | *(geçici: Mühendislik ile aynı — bkz. not aşağıda)* |
| İktisadi ve İdari Bilimler Fakültesi (İİBF) | ihsan.alp@ostimteknik.edu.tr | Kaynak Excel'deki "İİBF Puanlaması" sütunu |
| Yüksek Meslek Okulu (YMO) | zeynep.aydemir@ostimteknik.edu.tr | *(geçici: Mühendislik ile aynı)* |
| Bilişim Teknolojileri MYO | btmyo.bilgi@ostimteknik.edu.tr | *(geçici: Mühendislik ile aynı)* |

> **Not:** Kaynak Excel dosyasında yalnızca Mühendislik Fakültesi ve İİBF için ayrı
> puanlama sütunu vardı. MTF, YMO ve Bilişim Teknolojileri MYO için 180 kriter
> aynı kabul edilerek Mühendislik puan tablosu geçici olarak kullanılıyor. Bu
> birimler için gerçek puan tabloları sağlandığında `app.js` içindeki `UNITS`
> dizisinde ilgili `scoreKey` alanı güncellenmeli ve `data.js` içine yeni bir
> puan sütunu eklenmelidir.

GitHub Pages gibi **statik** bir barındırmada çalıştığı için, verilerin farklı
bilgisayarlar arasında otomatik olarak toplanabilmesi için ücretsiz bir
**Google Apps Script + Google E-Tablo** arka ucu kullanır (5 dakikada kurulur,
kod yazmanız gerekmez). Tüm birimler **aynı** tek Google E-Tablo'yu paylaşır;
birim ayrımı satır bazında (Birim sütunu) yapılır.

---

## 1) Hızlı kurulum (önerilen — gerçek, paylaşımlı veri toplama)

### Adım A — Google E-Tablo ve Apps Script'i kurun

1. [sheets.google.com](https://sheets.google.com) üzerinden **boş bir e-tablo** oluşturun.
2. Üst menüden **Uzantılar (Extensions) → Apps Script**'i açın.
3. Açılan editördeki örnek kodu silin, bu depodaki **`apps-script/Code.gs`** dosyasının tamamını yapıştırın.
4. Üstteki fonksiyon açılır listesinden **`setup`**'ı seçip **▶ Çalıştır (Run)** butonuna basın.
   - İlk çalıştırmada Google izin isteyecektir; kendi hesabınız olduğu için onaylayabilirsiniz.
   - Bu adım `Submissions` ve `Settings` sayfalarını oluşturur; **her birim için ayrı bir varsayılan şifre, bildirim e-postası ve boş eşik** değeri otomatik eklenir (aşağıdaki tabloya bakın).
5. Sağ üstten **Dağıt (Deploy) → Yeni Dağıtım (New deployment)**:
   - Tür: **Web Uygulaması (Web app)**
   - Yürütme kişisi (Execute as): **Ben (Me)**
   - Erişimi olanlar (Who has access): **Herkes (Anyone)**
   - **Dağıt**'a basın ve çıkan **Web App URL**'sini kopyalayın.

### Varsayılan yönetici şifreleri (Settings sayfasından değiştirin)

| Anahtar (Settings sayfası) | Varsayılan şifre | Kapsam |
|---|---|---|
| `AdminSifre` | `Ostim2025!` | **Merkezi yönetici** — tüm birimleri görür |
| `AdminSifre_muhendislik` | `Muh2025!` | Yalnızca Mühendislik Fakültesi |
| `AdminSifre_mtf` | `Mtf2025!` | Yalnızca MTF |
| `AdminSifre_iibf` | `Iibf2025!` | Yalnızca İİBF |
| `AdminSifre_myo` | `Myo2025!` | Yalnızca Yüksek Meslek Okulu |
| `AdminSifre_btmyo` | `Btmyo2025!` | Yalnızca Bilişim Teknolojileri MYO |

Her birimin dekanlığına/müdürlüğüne kendi şifresini iletebilirsiniz — kendi
girdikleri şifreyle yalnızca kendi birimlerinin verilerini görürler. Siz
(merkezi yönetici şifresiyle) tüm birimleri tek panelde, birim filtresiyle
görebilirsiniz.

Bildirim e-postaları `BildirimEposta_<birim>`, eşik değerleri `Esik_<birim>`
anahtarlarıyla aynı Settings sayfasında tutulur — hepsi kod değişikliği
gerekmeden doğrudan düzenlenebilir.

### E-posta bildirimleri

Bir öğretim elemanı formu gönderdiğinde iki e-posta otomatik gider:
- **İlgili birimin dekanlığına/müdürlüğüne** (`Settings!BildirimEposta_<birim>`): formun özeti.
- **Formu dolduran öğretim elemanına** (formda girdiği e-posta adresine): kendi kaydının bir kopyası.

E-postalar, Apps Script projesinin bağlı olduğu Google hesabından (`MailApp`) gönderilir.
Kişisel Gmail hesapları için günlük gönderim kotası ~100 e-postadır; bir OSTİM Google
Workspace hesabıyla dağıtırsanız bu kota çok daha yüksektir.

### Adım B — Siteyi yapılandırın

1. Bu klasördeki **`config.js`** dosyasını açın.
2. `window.APPS_SCRIPT_URL = "";` satırını, Adım A'da kopyaladığınız URL ile değiştirin:
   ```js
   window.APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx.../exec";
   ```
3. Kaydedin.

### Adım C — GitHub Pages'e yükleyin

1. GitHub'da yeni bir repo oluşturun (örn. `ostim-performans`).
2. Bu klasördeki tüm dosyaları (`index.html`, `style.css`, `app.js`, `data.js`, `config.js`) reponun kök dizinine yükleyin.
3. Repo **Settings → Pages** bölümünden: Source: **Deploy from a branch**, Branch: `main` / `(root)`.
4. Birkaç dakika içinde siteniz yayında olacaktır.

Artık her birimden öğretim elemanları bu bağlantıyı açıp formu doldurduğunda
veriler otomatik olarak aynı Google E-Tablo'ya (birim etiketiyle) kaydolur.

---

## 2) Hızlı deneme (kurulum yapmadan)

`config.js` içindeki `APPS_SCRIPT_URL` boş bırakılırsa uygulama **yerel modda**
çalışır: turuncu bir uyarı şeridi görürsünüz, veriler yalnızca o an kullanılan
tarayıcıda (localStorage) saklanır, birimler arasında **otomatik paylaşım
olmaz** ve yalnızca tek bir merkezi şifre geçerlidir (`LOCAL_ADMIN_PASSWORD`,
varsayılan `Ostim2025!`). Bu mod yalnızca test/demo amaçlıdır.

Bu modda gerçek veri toplamak isterseniz:
- Öğretim elemanı formu gönderdiğinde çıkan **"Kayıt Dosyasını İndir (.json)"**
  butonuyla bir dosya indirir ve bunu size (yöneticiye) e-posta ile gönderir.
- Siz yönetici panelinde **"Kayıt Dosyalarını İçe Aktar"** ile bu dosyaları
  yükleyip panele dahil edersiniz.

---

## Yarıda bırakıp devam etme

Öğretim elemanları formu iki şekilde yarıda bırakıp sonra devam edebilir:

1. **Otomatik (bulut modunda):** Aynı Ad Soyad + Birim + Dönem ile tekrar giriş
   yapan kişi, sistemde kayıtlı en son verisiyle otomatik olarak kaldığı yerden
   devam eder — herhangi bir işlem yapmasına gerek yoktur.
2. **Manuel yedek dosyasıyla (her modda çalışır, farklı cihazlar arasında da):**
   Form doldururken sağdaki panelden veya "Form Kaydedildi" ekranından
   **"💾 İlerlemeyi Yedekle (.json)"** ile o ana kadar girilenler (kanıt
   metinleri ve görseller dahil) bir dosyaya indirilir. Bu dosya, "Öğretim
   Elemanı Bilgileri" sayfasındaki **"📤 Taslak veya Word Belgesi Yükle"**
   alanına yüklendiğinde, tüm girilen veriler aynen geri yüklenir ve kaldığı
   yerden devam edilir — farklı bir bilgisayardan bile.
3. **"İmza İçin Word Belgesi" (.doc) ile de devam edilebilir:** Aynı yükleme
   alanı, indirdiğiniz `.doc` dosyasını (veya kişi bunu Microsoft Word'de
   gerçek `.docx` olarak kaydettiyse onu da, en iyi çaba ilkesiyle) okuyup
   girilen faaliyet adetlerini, çoklu yazarlı maddelerin yazar dağılımını,
   kanıt metinlerini/görsellerini ve kimlik bilgilerini geri yükler. `.doc`
   dosyası (bizim ürettiğimiz, hiç Word'de değiştirilmemiş hali) her zaman
   güvenilir çalışır; gerçek `.docx`'e dönüştürülüp düzenlenmiş dosyalarda
   (tablo yapısı çok değiştirilmediyse) en iyi çaba ile okunur — bu durumda
   `mammoth.js` adlı bir dönüştürücü kütüphanesi internet üzerinden
   otomatik yüklenir (yalnızca `.docx` yüklendiğinde).

## Kanıt (delil) ekleme ve ıslak imzalı Word belgesi

Formdaki her madde satırının sağında bir **"+ Kanıt"** butonu vardır. Buna basınca
o madde için bir metin kutusu (atıf, DOI, sözleşme no vb.) ve isteğe bağlı
**görsel ekleme** (ekran görüntüsü) alanı açılır. Girilen kanıt metni formla
birlikte kaydedilir.

Sticky özet panelindeki (ve "Form Kaydedildi" ekranındaki) **"📄 İmza İçin Word
Belgesi"** butonuna basıldığında, OSTİM'in resmi ıslak imzalı teslim şablonuyla
birebir uyumlu bir **.doc dosyası** indirilir: kimlik bilgileri tablosu, tüm
kriter tablosu (girilen faaliyet adetleri ve alan puanlarıyla), Genel Toplam,
ardından her beyan edilen madde için girdiğiniz kanıt metni/görselleri, en
altta da imza satırı. Bu dosya doğrudan Microsoft Word ile açılır; gözden
geçirip yazdırıp ıslak imzalayabilir veya imzalı halini tarayıp ilgili
birime teslim edebilirsiniz.

**Önemli:** Kanıt olarak eklenen **görseller yalnızca sizin tarayıcınızda**
tutulur ve sunucuya (Google E-Tablo'ya) gönderilmez — yalnızca o an
oluşturduğunuz Word belgesine gömülürler. Kanıt **metni** ise formla birlikte
kaydedilir ve yönetici panelinden de görülebilir/Word olarak indirilebilir
(ama bu durumda görseller olmadan, yalnızca metinle).

---

## Adil değerlendirme mekanizmaları (asgari şart + tavan)

Salt toplam puana dayalı sistemlerde, bir kişi hiç araştırma/yayın yapmadan
yalnızca Tanıtım/Hizmet faaliyetleriyle yüksek sıraya çıkabilir; ya da eski
yayınlardan gelen birikmiş atıf puanıyla, güncel dönemde üretim olmadan öne
çıkabilir. Bunu önlemek için YÖK/ÜAK doçentlik sisteminden esinlenen iki
mekanizma eklendi — **varsayılan olarak kapalıdır** (mevcut davranışı bozmaz),
komisyon karar verdiğinde Yönetici Paneli → "Birim Bazında Değerlendirme
Ayarları" kartından, kod değiştirmeden, birim başına ayrı ayrı açılabilir:

- **Asgari I. Bölüm Puanı / Asgari I. Bölüm Faaliyet Adedi** — Bu şartlardan
  ikisi birden tanımlanırsa **VEYA** mantığıyla çalışır: biri karşılanırsa
  yeterlidir (örn. tek ama yüksek puanlı bir yayın, az sayıda madde şartını
  telafi edebilir). Yalnızca biri girilirse, o tek şart geçerli olur. Hiçbiri
  karşılanmazsa kişi toplam puanı ne olursa olsun **"Asgari Şart Eksik"**
  sayılır. Bu, Tanıtım/Hizmet ağırlıklı bir profille üst sıraya çıkmayı
  engeller.
- **III+IV Tavanı (%)** — Hizmet (III) ve Tanıtım (IV) bölümlerinin toplamı,
  Araştırma+Proje (I+II) toplamına oranla belirtilen yüzdeyi aşamaz; aşan kısım
  "Değerlendirme Toplamı"na eklenmez (ham toplam yine de ayrıca gösterilir/dışa
  aktarılır, şeffaflık için).

Yönetici panelindeki tablo, CSV dışa aktarım ve indirilen Word belgesi hem
**Ham Toplam** (orijinal Excel mantığıyla birebir) hem **Değerlendirme
Toplamı** (politika uygulanmış) rakamlarını ayrı ayrı gösterir.

> **Not:** "Geçmişten gelen atıf puanıyla güncel üretim olmadan yüksek puan
> alma" sorunu için üçüncü bir mekanizma (dönemsel/güncellik ayrımı) da
> konuşuldu; bu, formda hangi maddelerin "atıf/geçmiş" sayılacağının ayrıca
> tanımlanmasını gerektirdiğinden şimdilik uygulanmadı — komisyon isterse
> ayrı bir aşamada eklenebilir.

## Dosya yapısı

```
index.html            Ana sayfa (tüm görünümleri yükler)
style.css              Görsel tasarım
data.js                180 kriterin tamamı, hem Müh. Fak. hem İİBF puanlarıyla
config.js              Apps Script bağlantı ayarı (BURAYI DÜZENLEYİN)
app.js                 Uygulama mantığı: birimler, hesaplama, form, yönetici paneli
apps-script/Code.gs     Google Apps Script arka uç kodu (e-tabloya yapıştırılır)
```

## Bilinen sınırlamalar

- MTF, YMO ve Bilişim Teknolojileri MYO için ayrı puan tablosu yok — yukarıdaki
  nota bakın.
- Genel Toplam, orijinal Excel'deki davranışla birebir aynıdır: I–IV. bölüm
  ağırlık yüzdeleri (%45/%25/%17,5/%7,5) yalnızca bilgi amaçlıdır, toplama
  formülünde fiilen uygulanmaz (orijinal dosyada da böyleydi).
- "Rektörlük Görüşü" (0-10) alanı, orijinal formdaki kırık (`#REF!`) formül
  nedeniyle Genel Toplam'a dahil edilmez; yalnızca bilgi amaçlı, yönetici
  tarafından girilen ayrı bir alandır.
- "Asgari performans eşiği" resmi belgede tanımlı değildir; her birimin
  yöneticisi kendi eşiğini belirler (varsayılan: o birimin katılımcı ortalaması).
