# 3DExperience VBA – Sıfırdan Kod Yazma Rehberi

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  18 dokümanlık rehber seti  |  Help uyumlu                                   ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

Kodlamaya **yeni başlayan** biri için, **3DExperience VBA** ile makro yazmayı adım adım anlatan **18 dokümanlık** rehber seti. **Proje sayfası (GitHub/GitLab):** repository kökündeki [README.md](../README.md). İçerik, proje kökündeki **Help** klasöründeki resmi dokümanlarla (**Help-Automation Development Guidelines**, **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ**, **Help-Native Apps Automation** vb.) uyumlu ve genişletilmiştir.

**Nasıl ilerlenir:** Aşağıdaki tablodan dokümanları **sırayla** (01 → 18) takip edin; belirli bir konu arıyorsanız doğrudan ilgili dokümana atlayabilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Dokümanlar (sırayla takip edin)

| # | Dosya | İçerik |
|---|-------|--------|
| 1 | [01-Giris-Neden-3DExperience-VBA.md](01-Giris-Neden-3DExperience-VBA.md) | Giriş – Neden 3DExperience VBA? (Dassault bakışı, desteklenen diller, bölgesel ayarlar) |
| 2 | [02-Ortam-Kurulumu.md](02-Ortam-Kurulumu.md) | Ortam kurulumu (Language/Release başlığı, makro konumu, dağıtım) |
| 3 | [03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md](03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md) | VBA temelleri – Değişkenler ve veri tipleri (resmi önekler: b, d, s, i, o, c) |
| 4 | [04-VBA-Temelleri-Kosullar-ve-Donguler.md](04-VBA-Temelleri-Kosullar-ve-Donguler.md) | VBA temelleri – Koşullar ve döngüler |
| 5 | [05-VBA-Temelleri-Prosedurler-ve-Fonksiyonlar.md](05-VBA-Temelleri-Prosedurler-ve-Fonksiyonlar.md) | VBA temelleri – Prosedürler ve fonksiyonlar |
| 6 | [06-3DExperience-Nesne-Modeli.md](06-3DExperience-Nesne-Modeli.md) | 3DExperience nesne modeli (Application, Editor, FileSystem, hiyerarşi ağacı) |
| 7 | [07-Makro-Kayit-ve-Inceleme.md](07-Makro-Kayit-ve-Inceleme.md) | Makro kayıt ve inceleme |
| 8 | [08-Sik-Kullanilan-APIler.md](08-Sik-Kullanilan-APIler.md) | Sık kullanılan API’ler |
| 9 | [09-Hata-Yakalama-ve-Debug.md](09-Hata-Yakalama-ve-Debug.md) | Hata yakalama ve debug (hata sınıflandırması, log tasarımı) |
| 10 | [10-Ornek-Proje-Bastan-Sona-Bir-Makro.md](10-Ornek-Proje-Bastan-Sona-Bir-Makro.md) | Örnek proje – Baştan sona bir makro (Design/Draft/Harden/Finalize, kontrol listesi) |
| 11 | [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md) | **Resmi kurallar ve hazırlık fazları** — Help’ten özet: kod sunum, isimlendirme, Option Explicit, hata yönetimi, ihtiyaç analizi, modül matrisi, kod taslağı fazları, hata/log, TAMAM/HAZIR listesi |
| 12 | [12-Servisler-ve-Yapilabilecek-Islemler.md](12-Servisler-ve-Yapilabilecek-Islemler.md) | **Servisler ve yapılabilecek işlemler** — Editor-level / Session-level servisler (tablo ve kod), FileSystem, Part/Product/Drawing işlemleri detayı |
| 13 | [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md) | **Neye nereden erişilir, neyi nasıl kullanırsın** — Erişim yolları (VBA) tablosu, kullanım özeti, kod kalıpları, tek sayfa zincir özeti |
| 14 | [14-VBA-ve-Excel-Etkilesimi.md](14-VBA-ve-Excel-Etkilesimi.md) | **VBA’dan Excel ile etkileşim** — CreateObject/GetObject, çalışma kitabı açma, hücre okuma/yazma, 3DExperience parametrelerini Excel’e yazma / Excel’den Part’a yazma |
| 15 | [15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md) | **Dosya seçtirme ve kaydetme diyalogları** — FileDialog (aç/kaydet/klasör), GetOpenFileName/GetSaveFileName (Windows API), InputBox ile yol, tam akış örnekleri |
| 16 | [16-Iyilestirme-Onerileri.md](16-Iyilestirme-Onerileri.md) | **İyileştirme önerileri** — Kod kalitesi, performans, bakım, test, kullanıcı deneyimi, dağıtım; isteğe bağlı kontrol listesi |
| 17 | [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md) | **Help içindeki dosyaları ne zaman ve nasıl kullanacaksınız** — Help klasörü yapısı, hangi dosya ne işe yarar, aşamaya göre kullanım, arama yöntemleri |
| 18 | [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) | **Sık yapılan hatalar ve dikkat edilmesi gereken özel noktalar** — Nothing/Update/On Error, V5 API, InputBox iptal, locale, servis sırası, özet tablo |

════════════════════════════════════════════════════════════════════════════════

## Nasıl kullanılır?

- **Sıfırdan başlıyorsanız:** 1. dokümandan başlayıp sırayla 10’a kadar ilerleyin; kurumsal standartlar için 11’i, servisler ve işlem detayı için 12’yi okuyun.  
- **Belirli konu arıyorsanız:** Yukarıdaki tablodan ilgili dokümanı açın. **“Buna nereden erişirim, bunu nasıl kullanırım?”** için **13. doküman**; **VBA’dan Excel’e veri yazma/okuma** için **14. doküman**; **dosya seçtirme / kaydetme diyaloğu** için **15. doküman**; **kod ve süreç iyileştirme önerileri** için **16. doküman** kullanın.  
- **Resmi kurallar ve fazlar:** Help’e dayalı özet ve kontrol listeleri için **11. doküman** kullanın.  
- **VBA kodu checklist (detaylı):** Teslim veya kod incelemesi öncesi **[VBA-Kod-Checklist.md](VBA-Kod-Checklist.md)** dosyasındaki maddeleri işaretleyin.  
- **API detayı için:** Proje kökündeki **VBA_API_REFERENCE.md** (varsa) veya **Help/VBA_CALL_LIST.txt** ve **Help/text/** klasöründeki metinleri kullanın.
- **Çalıştırılabilir örnek makrolar:** Proje kökündeki **[Examples/](../Examples/README.md)** klasöründe `.bas` dosyaları bulunur.  
- **Help dosyalarını ne zaman/nasıl kullanacağınız:** **[17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md)** — Help klasörü yapısı, hangi dosyayı hangi aşamada açacağınız, arama yöntemleri.  
- **Sık hatalar ve dikkat edilecekler:** **[18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md)** — Option Explicit, Nothing/Update, On Error, V5 API, InputBox iptal, locale, servis sırası vb.

Tüm dokümanlar **3DExperience VBA** özelinde yazılmıştır; örnekler ve terimler bu platforma göredir. Help klasöründeki PDF’ler (Automation Development Guidelines, 3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ, Native Apps Automation vb.) tam ve güncel referanstır.

════════════════════════════════════════════════════════════════════════════════

## Hızlı referans – Sık kullanılan kalıplar

| İhtiyaç | Örnek kod (kavramsal) |
|--------|------------------------|
| Uygulama al | `Set oApp = GetObject(, "CATIA.Application")` |
| Aktif belge | `Set oDoc = oApp.ActiveDocument` |
| Part al | `Set oPart = oDoc.GetItem("Part")` veya `Set oPart = oDoc` |
| Parametre oku | `Set oParam = oPart.Parameters.Item("Length.1")` → `dVal = oParam.Value` |
| Parametre yaz | `oParam.Value = 50.5` → `oPart.Update` |
| Shapes döngüsü | `For i = 1 To oPart.Shapes.Count` … `Set oSh = oPart.Shapes.Item(i)` |
| Hata yakala | `On Error GoTo HataYakala` … `Exit Sub` … `HataYakala: MsgBox Err.Description` |
| Nothing kontrolü | `If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub` |

Bu kalıplar sürüme göre değişebilir; tam API için **VBA_API_REFERENCE.md** ve **Help/text/** kullanın.

════════════════════════════════════════════════════════════════════════════════

## Doküman satır sayıları (rehber genişliği)

Rehber seti, girişten resmi kurallara kadar **binlerce satır** örnek ve açıklama içerir. Her dokümanda:

- Temel kavramlar ve kurallar  
- 3DExperience’a özel VBA örnekleri  
- Help dokümanlarından alıntılar ve uyumlu kod blokları  
- Kontrol listeleri ve sonraki adım önerileri  

bulunur. Baştan sona takip edildiğinde sıfırdan makro yazıp dağıtım öncesi kontrol listesini uygulayabilecek seviyeye gelirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Örnek makro türleri (dokümanlarda dağılım)

| Doküman | Örnek türü |
|---------|------------|
| 01 | Senaryo tabloları, kavramsal “parça adı göster”, otomasyon türleri |
| 02 | İlk MsgBox, InputBox, Language/Release başlığı, F5/F8 |
| 03 | Değişken tipleri, Set, Const, Variant, Date, önekli isimler |
| 04 | If/Else, Select Case, For/For Each, Do While, Nothing kontrolü, Exit For |
| 05 | Sub/Function, parametreler, ByVal/ByRef, Optional, Call |
| 06 | GetObject, ActiveDocument, Part/Product/Drawing, Shapes, FileSystem, GetSessionService |
| 07 | Kayıt çıktısı, sadeleştirme, Nothing kontrolü, tek Update, sabit→değişken |
| 08 | Parametre okuma/yazma, Shapes döngüsü, Drawing Sheets/Views, Product Children, GetService |
| 09 | On Error GoTo, Resume Next, Err.Number, breakpoint, Immediate, LogYaz |
| 10 | Tam makro iskeletleri: bilgi göster, parametre oku/yaz, listele, dosyaya yaz, log, modüler yapı |
| 11 | Başlık, cross-platform, Err.Raise, Design/Draft/Harden/Finalize, risk matrisi, 3 satır kullanım |

| 12–15 | Servisler, erişim tabloları, Excel, dosya diyalogları (12–15. dokümanlar) |
| 16 | İyileştirme önerileri, kontrol listesi (kalite, performans, bakım, test, UX, dağıtım) |
| 17 | Help dosyalarını ne zaman/nasıl kullanacağınız, hangi dosya hangi aşamada |
| 18 | Sık yapılan hatalar, dikkat edilecek özel noktalar (Nothing, Update, On Error, V5 API, locale) |

Bu dağılım, belirli bir örnek türünü nerede bulacağınızı hızlıca göstermek içindir.

════════════════════════════════════════════════════════════════════════════════

## Başlamak için en kısa yol

1. **01–02** ile giriş ve ortam kurulumunu yapın; ilk MsgBox makrosunu çalıştırın.  
2. **03–05** ile değişken, koşul, döngü ve Sub/Function temellerini öğrenin.  
3. **06** ile Application → Document → Part/Product/Drawing zincirini ve nesne modelini inceleyin.  
4. **07** ile makro kaydı yapıp üretilen kodu sadeleştirmeyi deneyin.  
5. **08** ile parametre, Shapes, Drawing, Product API örneklerini inceleyin.  
6. **09** ile On Error ve debug (breakpoint, Immediate) kullanımını öğrenin.  
7. **10** ile baştan sona birkaç tam makro örneğini kopyalayıp kendi ihtiyacınıza uyarlayın.  
8. **11** ile resmi kurallar (başlık, isimlendirme, Design/Draft/Harden/Finalize) ve kontrol listelerini uygulayın.  
9. **12** ile servisler (Editor/Session) ve yapılabilecek işlemlerin detaylı listesi ile kod örneklerini inceleyin.  
10. **13** ile “neye nereden erişilir, neyi nasıl kullanırsın” tablolarını ve kod kalıplarını kullanın (hızlı referans).  
11. **14** ile VBA’dan Excel’e veri yazma/okuma (Part ↔ Excel).  
12. **15** ile dosya seçtirme ve kaydetme diyalogları (FileDialog, GetOpenFileName/GetSaveFileName).  
13. **16** ile kod kalitesi, performans, bakım, test ve dağıtım için iyileştirme önerilerini ve isteğe bağlı kontrol listesini uygulayın.  
14. **17** ile Help klasöründeki dosyaları **ne zaman ve nasıl** kullanacağınızı (hangi dosya, hangi aşama, arama yöntemleri) öğrenin.  
15. **18** ile **sık yapılan hataları** ve **dikkat edilmesi gereken özel noktaları** (Nothing, Update, On Error, V5 API, locale, servis sırası vb.) inceleyin.

Bu sırayla ilerlediğinizde toplam rehber **5000 satır civarı** örnek ve açıklama içerir; her dokümanda birden fazla VBA kodu ve senaryo bulunur.

════════════════════════════════════════════════════════════════════════════════

## Doküman başına örnek sayısı (yaklaşık)

| Doküman | Örnek sayısı (yaklaşık) | İçerik türü |
|---------|--------------------------|-------------|
| 01 | 8+ | Senaryo tabloları, kod parçaları, otomasyon türleri |
| 02 | 6+ | MsgBox, InputBox, başlık, F5/F8, dağıtım |
| 03 | 10+ | Dim, Set, Const, Variant, Date, Enum, Type |
| 04 | 15+ | If, Select Case, For, For Each, Do, Exit For, bayrak |
| 05 | 12+ | Sub, Function, ByVal/ByRef, Optional, Public/Private |
| 06 | 12+ | GetObject, Part, Product, Drawing, FileSystem, Services |
| 07 | 8+ | Kayıt öncesi/sonrası, sadeleştirme, değişken tiplendirme |
| 08 | 14+ | Parametre, Shapes, Drawing, Product, GetService örnekleri |
| 09 | 14+ | On Error, Err, breakpoint, Immediate, Log, Retry |
| 10 | 19+ | Tam makro iskeletleri (bilgi, parametre, listele, dosya, rollback) |
| 11 | 33+ | Başlık, kurallar, Design/Harden/Finalize, risk matrisi, 3 satır kullanım |
| 12 | 25+ | Servisler (tablo + kod), FileSystem, Part/Product/Drawing işlemleri |
| 13 | — | Erişim tabloları, kullanım özeti, kod kalıpları, zincir özeti |
| 14 | 8+ | Excel CreateObject/GetObject, aç/kapat, hücre okuma/yazma, Part ↔ Excel örnekleri |
| 15 | 10+ | FileDialog (aç/kaydet/klasör), GetOpenFileName/GetSaveFileName API, InputBox yol, akış örnekleri |
| 16 | — | İyileştirme önerileri (kalite, performans, bakım, test, UX, dağıtım), kontrol listesi |
| 17 | — | Help klasörü yapısı, hangi dosya ne işe yarar, aşamaya göre kullanım, arama yöntemleri |
| 18 | 22+ | Sık hatalar (12 madde) + dikkat noktaları (10 madde), özet tablo, kod örnekleri |

Toplamda **150’den fazla** ayrı örnek veya kod blokları rehberde yer alır; tekrarlar ve varyasyonlarla birlikte toplam satır sayısı 5000’e ulaşır.

════════════════════════════════════════════════════════════════════════════════

## İçerik türleri

- **Kod blokları:** VBA örnekleri (Sub, Function, tek satırlık çağrılar).  
- **Tablolar:** Senaryo listeleri, modül matrisi, hata seviyeleri, kontrol listeleri.  
- **Açıklamalar:** “Ne yapar”, “Neden”, “Dikkat”, “Help’e göre” metinleri.  
- **Başlık örnekleri:** Language, Release, Purpose, Assumptions, Copyright.  
- **Kullanım yönergeleri:** 3 satırlık talimat, dağıtım notu, sonraki adım önerileri.

Tüm bu içerikler 3DExperience VBA bağlamında yazılmıştır; API isimleri sürüme göre **VBA_API_REFERENCE.md** ve **Help/text/** ile doğrulanmalıdır.
