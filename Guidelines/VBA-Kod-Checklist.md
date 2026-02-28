# VBA Kodu İçin Detaylı Checklist

```
================================================================================
  Teslim / kod incelemesi öncesi: zorunlu ve önerilen maddeler (12 bölüm + özet)
================================================================================
```

3DExperience VBA makrolarını **yazarken**, **teslim etmeden önce** veya **kod incelemesinde** kullanabileceğiniz detaylı kontrol listesi. Her bölümde **zorunlu (✓)** ve **önerilen (○)** maddeler ayrılmıştır. Bkz. [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md), [16-Iyilestirme-Onerileri.md](16-Iyilestirme-Onerileri.md). **Help dosyalarını ne zaman/nasıl kullanacağınız:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md). **Sık hatalar ve dikkat edilecekler:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

------------------------------------------------------------

## 1. Modül başlığı ve tanım

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 1.1 | **Purpose** (amaç) yorumu var; 2–3 cümle ile ne yaptığı yazılmış. | ✓ | Help: zorunlu başlık. |
| 1.2 | **Assumptions** (varsayımlar) yazılmış: hangi workbench, belge türü, seçim durumu. | ✓ | Örn: "Part Design açık, aktif belge Part." |
| 1.3 | **Language: VBA** belirtilmiş. | ✓ | CATScript, VBScript, VBA, VB.Net, C#, Python vb. |
| 1.4 | **Release** (sürüm) yazılmış; örn. 3DEXPERIENCE R2024x. | ✓ | İlk desteklenen sürüm. |
| 1.5 | **Regional Settings** belirtilmiş; örn. English (United States). | ✓ | Farklı locale’de davranış değişebilir. |
| 1.6 | **Author** ve **Copyright** (teslim/kurumsal kullanımda) doldurulmuş. | ○ | Örnek script’lerde boş bırakılabilir. |
| 1.7 | Versiyon etiketi veya revizyon satırı var; örn. `' REV 1.2 – 2025-03-01`. | ○ | Dağıtım ve takip için. |

------------------------------------------------------------

## 2. Option Explicit ve değişkenler

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 2.1 | **Option Explicit** modülün **ilk satırında** yer alıyor. | ✓ | Tanımsız değişken kullanımını engeller. |
| 2.2 | Tüm değişkenler **Dim** ile tanımlı; tip belirtilmiş (Object, String, Long, Double vb.). | ✓ | Variant gereksiz yere kullanılmamalı. |
| 2.3 | Resmi **önekler** kullanılmış: b (Boolean), d (Double), s (String), i (Integer/Long), o (Object), c (Collection). | ○ | bReadOnly, dLength, sPartName, oPart. |
| 2.4 | **Const** ile sabitler tanımlanmış; sihirli sayılar (örn. 100, 0.001) koddan çıkarılmış. | ○ | MAX_ITERATION, LOG_PATH vb. |
| 2.5 | Uzun satırlar 80 karakteri aşmıyor veya satır devamı `_` ile bölünmüş. | ○ | Help: kod sunum kuralı. |

------------------------------------------------------------

## 3. Hata yönetimi (On Error)

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 3.1 | En az bir **On Error** kullanılıyor: **On Error GoTo etiket** veya **On Error Resume Next** + kısa blok. | ✓ | Hata durumunda makro kontrollü davranmalı. |
| 3.2 | **On Error Resume Next** kullanıldıysa hemen sonrasında **Err.Number** (veya Err) kontrolü yapılıyor. | ✓ | Aksi halde hata sessizce yutulur. |
| 3.3 | **On Error Resume Next** sonrası **On Error GoTo 0** ile normale dönülüyor. | ✓ | Sonraki satırlarda hata yakalama açık olsun. |
| 3.4 | **On Error GoTo HataYakala** (veya benzeri) kullanıldıysa **Exit Sub** / **Exit Function** ile etiket bloğuna düşmeden çıkış var. | ✓ | Aksi halde hata olmasa da HataYakala çalışır. |
| 3.5 | Hata mesajında **Err.Number** ve **Err.Description** kullanıcıya veya log’a yazılıyor. | ✓ | "Hata oluştu" yerine somut bilgi. |
| 3.6 | Kritik senaryolarda **Err.Raise** (9000–9999) ile özel hata fırlatılıyor (test/otomasyon için). | ○ | Help önerisi. |
| 3.7 | **On Error Resume Next** sadece tek satır veya çok kısa blokta; tüm makro boyunca açık değil. | ✓ | Hata yönetimi dar tutulmalı. |

------------------------------------------------------------

## 4. 3DExperience API ve nesne erişimi

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 4.1 | **GetObject(, "CATIA.Application")** veya uygun giriş noktası sonrası **oApp Is Nothing** kontrolü yapılıyor. | ✓ | Uygulama kapalıysa anlamlı mesaj. |
| 4.2 | **ActiveDocument** alındıktan sonra **oDoc Is Nothing** kontrolü var. | ✓ | Açık belge yoksa çıkış. |
| 4.3 | **GetItem("Part")**, **GetItem("Product")**, **GetItem("DrawingRoot")** vb. sonrası **Nothing** veya geçerli nesne kontrolü yapılıyor. | ✓ | Belge türü yanlışsa hata vermemeli. |
| 4.4 | Koleksiyonlara (**Parameters**, **Shapes**, **Children**, **Sheets** vb.) erişimden önce **Is Nothing** ve gerekiyorsa **.Count > 0** kontrolü var. | ✓ | Boş koleksiyonda Item(1) hata verir. |
| 4.5 | **Part** veya **Product** değiştiriliyorsa **Update** yalnızca **bir kez**, tüm değişiklikler bittikten sonra (döngü dışında) çağrılıyor. | ✓ | Döngü içinde Update performans ve tutarlılık sorunudur. |
| 4.6 | Eski **V5 API** kullanılmıyor: **Documents.Add**, **HybridShapeFactoryOld** vb. | ✓ | 3DExperience’ta desteklenmeyebilir. |
| 4.7 | **Editor-level** servis kullanılıyorsa aktif pencerenin doğru türde (Part/Product/Drawing) olduğu varsayılıyor veya kontrol ediliyor. | ○ | GetService Nothing dönebilir. |

------------------------------------------------------------

## 5. Kod yapısı ve isimlendirme

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 5.1 | **Sub/Function** adları fiil veya fiil ifadesi; kelimelerin ilk harfi büyük (mixed case). | ○ | DoItBetter, GetActivePart. |
| 5.2 | Uzun tek Sub (50+ satır) yerine mantıklı **Sub/Function** bölünmesi yapılmış. | ○ | Okunabilirlik ve tekrar kullanım. |
| 5.3 | Tekrarlayan bloklar (Application al, Part al, log yaz) **ortak Sub/Function**’a taşınmış. | ○ | GetActivePart(), LogSatir(). |
| 5.4 | **Girinti** 4 boşluk; yorumlar `'` ile yazılmış. | ○ | Help: kod sunum. |
| 5.5 | Her Sub/Function için kısa **başlık yorumu** (amaç, parametreler, dönüş) var. | ○ | Help önerisi. |

------------------------------------------------------------

## 6. Dosya, yol ve dış kaynaklar

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 6.1 | Dosya/klasör yolu kullanılıyorsa **var mı / yazılabilir mi** kontrolü yapılıyor (FileSystem.Exists vb.). | ○ | Kaydetmeden önce hata almamak için. |
| 6.2 | Log, çıktı klasörü, varsayılan dosya adı **Const** veya modül başında tek yerde toplanmış. | ○ | Dağıtımda değiştirmek kolay olsun. |
| 6.3 | Sabit yol (C:\Temp vb.) dağıtım notunda veya yorumda belirtilmiş. | ○ | Farklı ortamda değiştirilecek. |
| 6.4 | **Hassas bilgi** (şifre, token) koda veya log’a yazılmıyor. | ✓ | Güvenlik. |

------------------------------------------------------------

## 7. Kullanıcı arayüzü ve mesajlar

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 7.1 | Kullanıcıya **en az bir** açıklayıcı çıktı veriliyor: **MsgBox** veya eşdeğeri. | ✓ | "Bitti" / "Hata" / sonuç özeti. |
| 7.2 | Başarı/hata mesajları **somut**: "12 parametre güncellendi", "Length.1 bulunamadı" gibi. | ○ | "Bitti" yerine sayı/ad. |
| 7.3 | Uzun süren işlemlerde (10+ saniye) başta bilgi mesajı veya "X öğe işlenecek" uyarısı var. | ○ | Kullanıcı deneyimi. |
| 7.4 | **InputBox** iptal (boş dönüş) kontrol ediliyor; anlamlı mesaj veya çıkış yapılıyor. | ✓ | Cancel’a basınca hata vermemeli. |

------------------------------------------------------------

## 8. Log ve izlenebilirlik

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 8.1 | Hata veya kritik adım log’a yazılıyorsa **bağlam** eklenmiş: belge adı, parametre adı, adım no. | ○ | LogSatir "ERROR", 9100, "Parametre yok", "Doc=" & oDoc.Name |
| 8.2 | Log dosyası yolu ve rotasyonu (varsa) dokümante edilmiş. | ○ | Kurumsal senaryoda. |
| 8.3 | Uzun işlemde **süre** (Timer) ölçülüp log’a veya rapora yazılmış. | ○ | Performans analizi. |

------------------------------------------------------------

## 9. Rollback ve çok adımlı işlemler

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 9.1 | Birden fazla nesneyi değiştiriyorsa hata durumunda **rollback** (eski değerlere dönme) veya "kısmen uygulandı" uyarısı düşünülmüş. | ○ | Veri tutarlılığı. |
| 9.2 | Kritik değerler değiştirilmeden önce **eski değer** saklanıyorsa (geri alma için) bu yorumla belirtilmiş. | ○ | Okunabilirlik. |

------------------------------------------------------------

## 10. Test ve senaryolar

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 10.1 | **Hiç belge açık değilken** makro çalıştırıldı; mesaj veya çıkış doğru. | ✓ | ActiveDocument = Nothing. |
| 10.2 | **Yanlış belge türü** açıkken (örn. Part beklenirken Drawing) test edildi; hata vermeden veya anlamlı mesajla çıkıyor. | ✓ | GetItem("Part") = Nothing. |
| 10.3 | **Sınır durumlar** denendi: 0 parametre, 1 parametre, kullanıcı iptal (InputBox boş). | ○ | Count = 0, Cancel. |
| 10.4 | Kod **derleniyor**; manuel çalıştırmada (F5) beklenen davranış görülüyor. | ✓ | Temel geçerlilik. |

------------------------------------------------------------

## 11. Dağıtım ve dokümantasyon

| # | Madde | Zorunlu | Not |
|---|--------|:-------:|-----|
| 11.1 | **3 satırlık kullanım yönergesi** verilmiş veya yazılmış: 1) Ne açılacak, 2) Makro nasıl çalıştırılacak, 3) Bitti nasıl anlaşılacak. | ○ | Kullanıcı talimatı. |
| 11.2 | Gerekli **yetkiler** (Part yazma, ağ sürücüsüne kaydetme vb.) Assumptions veya ayrı notta yazılmış. | ○ | Dağıtım için. |
| 11.3 | **Değişiklik günlüğü** veya revizyon notu (kurumsal teslimde) güncel. | ○ | Versiyon takibi. |

------------------------------------------------------------

## 12. Kurumsal / opsiyonel (genişletilmiş)

| # | Madde | Not |
|---|--------|-----|
| 12.1 | İhtiyaç formu (Katman A–E) yanıtlandı; doğru workbench ve API modül matrisinde işaretli. | 11. doküman. |
| 12.2 | Özel hata numaraları 9000–9999 aralığında; Err.Raise ile tutarlı kullanılıyor. | 9. doküman. |
| 12.3 | 10K+ occurrence veya büyük döngüde timeout/performans testi yapıldı; süre log’a yazılıyor. | Performans. |
| 12.4 | Workbench varlığı veya read-only belge kontrolü (gerekiyorsa) yapılıyor ve loglanıyor. | Güvenlik/lisans. |
| 12.5 | Talep sahibi / kullanıcı "Çalıştı" (UAT) onayı verdi. | Teslim kriteri. |

------------------------------------------------------------

## Özet: Zorunlu minimum (tek sayfa)

Aşağıdakilerin **hepsi** işaretli olmalı:

- [ ] **Option Explicit** modül başında.
- [ ] **Purpose**, **Assumptions**, **Language**, **Release**, **Regional Settings** başlıkta.
- [ ] **On Error** (GoTo etiket veya Resume Next + Err kontrolü + GoTo 0) kullanılıyor.
- [ ] **GetObject** / **ActiveDocument** / **GetItem** sonrası **Nothing** kontrolü.
- [ ] Koleksiyon (Parameters, Shapes vb.) **Nothing** ve **Count** kontrolü.
- [ ] **oPart.Update** (veya eşdeğeri) **yalnızca bir kez**, döngü dışında.
- [ ] **Eski V5 API** (Documents.Add, HybridShapeFactoryOld vb.) yok.
- [ ] Kod **derleniyor** ve **manuel testte** çalışıyor.
- [ ] Kullanıcıya **en az bir** mesaj (MsgBox/echo) veriliyor.
- [ ] **InputBox** iptal durumu kontrol ediliyor.
- [ ] Hata mesajında **Err.Number** ve **Err.Description** kullanılıyor.
- [ ] **Hassas bilgi** koda/log’a yazılmıyor.

------------------------------------------------------------

**İlgili dokümanlar:** [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md) (TAMAM/HAZIR, kod sunum) · [16-Iyilestirme-Onerileri.md](16-Iyilestirme-Onerileri.md) (iyileştirme önerileri) · [09-Hata-Yakalama-ve-Debug.md](09-Hata-Yakalama-ve-Debug.md) (On Error, log). **Tüm rehber:** [README](README.md).
