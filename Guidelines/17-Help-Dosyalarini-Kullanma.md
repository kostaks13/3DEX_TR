# 17. Help İçindeki Dosyaları Ne Zaman ve Nasıl Kullanacaksınız?

**Bu dokümanda:** Help klasörü yapısı; hangi dosya ne zaman kullanılır; arama yöntemleri (grep, editör); aşamaya göre kullanım.

Bu doküman, proje kökündeki **Help** klasöründeki dosyaların **ne işe yaradığını**, **hangi aşamada** ve **nasıl** kullanılacağını özetler. Rehberdeki diğer dokümanlar (01-18) sık sık "Help'e bakın", "VBA_API_REFERENCE.md ve Help/text/" der; burada hepsini tek yerde topluyoruz.

════════════════════════════════════════════════════════════════════════════════

## Help klasörünün yapısı

```
Help/
  VBA_CALL_LIST.txt          <- Çağrılabilir API listesi (ham)
  API_REPORT.csv             <- API raporu (ek kaynak)
  text/                      <- PDF'lerden çevrilmiş metin dosyaları
    Help-Automation Development Guidelines.txt
    Help-3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.txt
    Help-Native Apps Automation.txt
    Help-Common Services.txt
    Help-Automation Reference.txt
    Help-3D Modeling.txt
    Help-Simulation.txt
    Help-Social and Collaborative.txt
    Help-Adoption.txt
    3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.txt
```

Proje kökünde ayrıca **VBA_API_REFERENCE.md** vardır; Help ve VBA_CALL_LIST/API_REPORT'tan üretilmiş, okunaklı API referansıdır.

════════════════════════════════════════════════════════════════════════════════

## Hangi dosyayı ne zaman kullanacaksınız?

| İhtiyaç / Soru | Kullanılacak dosya | Ne zaman |
|----------------|---------------------|----------|
| **Kod nasıl yazılır?** (girinti, yorum, isimlendirme, Option Explicit, başlık) | **Help-Automation Development Guidelines.txt** | Kod yazmaya başlamadan önce; 11. dokümanla birlikte. |
| **Makro hazırlık aşamaları** (Design, Draft, Harden, Finalize), TAMAM/HAZIR listesi, ihtiyaç analizi | **Help-3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.txt** | Proje planlama ve teslim öncesi; 10. ve 11. dokümanla birlikte. |
| **Nesne modeli** (Application, Editors, ActiveEditor, FileSystem, Documents, Part, Product, Drawing) | **Help-Native Apps Automation.txt** | 6. ve 8. dokümanı okurken; "buna nereden ulaşırım?" diye sorduğunuzda. |
| **Servisler** (GetSessionService, GetService, Service Identifier listesi) | **Help-Common Services.txt**, **Help-Native Apps Automation.txt** | 12. doküman (Servisler) ve editor/session servis kullanırken. |
| **Belirli bir sınıf/metod imzası** (parametreler, dönüş tipi, örnek) | **VBA_API_REFERENCE.md** (öncelik), **Help-Automation Reference.txt**, **Help/text/** içinde arama | 7. (kayıt sonrası), 8., 13. dokümanlarda; API adını biliyorsanız. |
| **Hiyerarşi ağacı** (Application altında ne var, Part/Product/Drawing altında ne var) | **3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.txt** | 6. doküman (Nesne modeli) ve 13. (Erişim rehberi) ile birlikte. |
| **Hata yönetimi, log, Err.Raise** (resmi öneriler) | **Help-Automation Development Guidelines.txt**, **Help-3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.txt** | 9. ve 11. dokümanlarda hata/log tasarımı yaparken. |
| **Çağrılabilir API listesi** (ham liste) | **Help/VBA_CALL_LIST.txt** | VBA_API_REFERENCE.md yeterli değilse; script ile referans üretirken. |

════════════════════════════════════════════════════════════════════════════════

## Nasıl kullanacaksınız?

### 1. Metin dosyalarında arama (Help/text/*.txt)

- **Editörünüzde** (VS Code, Cursor, Notepad++) **Ctrl+F / Cmd+F** ile açın; aradığınız kelimeyi yazın (örneğin `Part`, `GetSessionService`, `ActiveDocument`, `FileSystem`).
- **Terminalde:** `grep -l "GetItem" Help/text/*.txt` ile hangi dosyada geçtiğini bulun; sonra ilgili dosyayı açıp bağlamı okuyun.
- **Ne arayacaksınız?** Makro kaydında çıkan sınıf/method adı, rehberde geçen API adı (Parameters, Shapes, Children, Service Identifier vb.).

### 2. VBA_API_REFERENCE.md ile hızlı bakış

- Proje kökündeki **VBA_API_REFERENCE.md** dosyası sınıf ve metodları listeler; her biri için kısa amaç ve örnek VBA çağrısı vardır.
- **Ne zaman:** Belirli bir nesne veya metodun tam adını ve örnek kullanımı görmek istediğinizde. Help/text/ içinde detay aramak yerine önce buradan bakın; yetersizse Help metinlerinde arama yapın.

### 3. Hangi aşamada hangi dosya?

| Rehber adımı / Aşama | Açın / Kullanın |
|----------------------|------------------|
| **01-02 (Giriş, Ortam)** | Zorunlu değil; merak ederseniz Help-Automation Development Guidelines (genel bakış). |
| **03-05 (VBA temelleri)** | Help-Automation Development Guidelines (isimlendirme, Option Explicit, yorum kuralları). |
| **06 (Nesne modeli)** | Help-Native Apps Automation, 3DEXPERIENCE Otomasyon Hiyerarşi Ağacı; Application, ActiveEditor, FileSystem bölümleri. |
| **07 (Makro kayıt)** | Kayıt çıktısındaki sınıf/method adlarını VBA_API_REFERENCE.md veya Help/text/ içinde arayın; Help-Automation Development Guidelines (kayıt sonrası düzenlemeler). |
| **08 (Sık kullanılan API)** | VBA_API_REFERENCE.md, Help-Native Apps Automation, Help-Common Services (servisler için). |
| **09 (Hata, debug)** | Help-Automation Development Guidelines, MACRO HAZIRLIK YÖNERGESİ (hata sınıflandırma, log). |
| **10-11 (Örnek proje, Resmi kurallar)** | Help-3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ (fazlar, TAMAM/HAZIR), Help-Automation Development Guidelines (başlık, teslim). |
| **12 (Servisler)** | Help-Common Services.txt, Help-Native Apps Automation (Service Identifier, GetSessionService). |
| **13-16 (Erişim, Excel, Dosya, İyileştirme)** | Gerekirse VBA_API_REFERENCE.md ve Help-Native Apps Automation (erişim yolu); diğerleri rehberde yeterli. |

════════════════════════════════════════════════════════════════════════════════

## Özet: "Şimdi ne yapayım?"

1. **Kod yazmaya başlıyorum** → 11. doküman + **Help-Automation Development Guidelines.txt** (başlık, isimlendirme, Option Explicit).
2. **Application / Part / Product nereden alınır bilmiyorum** → 6. ve 13. doküman + **Help-Native Apps Automation.txt**, **3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.txt**.
3. **Kayıt ettim, çıkan kodu anlamadım** → Sınıf/method adını **VBA_API_REFERENCE.md** veya **Help/text/** içinde arayın; 7. doküman (sadeleştirme).
4. **Servis kullanacağım (GetSessionService, GetService)** → 12. doküman + **Help-Common Services.txt**, **Help-Native Apps Automation.txt** (Service Identifier).
5. **Teslim / kurumsal standart** → 11. doküman + **Help-3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.txt** (fazlar, TAMAM/HAZIR).

════════════════════════════════════════════════════════════════════════════════

## İlgili dokümanlar

**Tüm rehber:** [README](README.md). **Resmi kurallar özeti:** [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md). **Nesne modeli / erişim:** [06-3DExperience-Nesne-Modeli.md](06-3DExperience-Nesne-Modeli.md), [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md). **Sık hatalar ve dikkat edilecekler:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md). **API listesi (proje kökü):** [VBA_API_REFERENCE.md](../VBA_API_REFERENCE.md).

**Gezinme:** Önceki: [16-Iyilestirme](16-Iyilestirme-Onerileri.md) | [Rehber listesi](README.md) | Sonraki: [18-Sik-Hatalar](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) →
