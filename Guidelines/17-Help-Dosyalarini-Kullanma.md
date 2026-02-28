# 17. Help İçindeki Dosyaları Ne Zaman ve Nasıl Kullanacaksınız?

**Bu dokümanda:** Help klasörü yapısı (PDF'ler); hangi PDF ne zaman kullanılır; PDF'de arama (Ctrl+F). Rehberdeki diğer dokümanlar (01-18) "Help'e bakın" dediğinde burada hepsini tek yerde topluyoruz.

Bu doküman, proje kökündeki **Help** klasöründeki **resmi referans PDF'lerinin** ne işe yaradığını, hangi aşamada ve nasıl kullanılacağını özetler. **Tüm referanslar PDF üzerinedir;** proje içinde ayrıca [Help/ARAMA_REHBERI.md](../Help/ARAMA_REHBERI.md) ile hangi konunun hangi PDF'te olduğu ve PDF'de nasıl arayacağınız anlatılır.

════════════════════════════════════════════════════════════════════════════════

## Help klasörünün yapısı (PDF'ler)

Help klasöründe **resmi Dassault Systèmes / 3DExperience dokümantasyonundan** referans alınan **PDF dosyaları** kullanılır. Örnek isimler (sürüme göre değişebilir):

```
Help/
  Help-Automation Development Guidelines.pdf    <- Kod kuralları, isimlendirme, Option Explicit, başlık
  3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf   <- Design/Draft/Harden/Finalize, TAMAM/HAZIR, ihtiyaç analizi
  Help-Native Apps Automation.pdf             <- Nesne modeli: Application, Part, Product, Drawing, FileSystem
  Help-Common Services.pdf                     <- GetService, GetSessionService, Service Identifier listesi
  Help-Automation Reference.pdf                <- Sınıf/metod imzaları, parametreler, örnekler
  Help-3D Modeling.pdf                         <- 3D modelleme API'leri
  Help-Simulation.pdf                          <- Simülasyon API'leri
  Help-Social and Collaborative.pdf            <- Sosyal / işbirliği API'leri
  Help-Adoption.pdf                            <- Adoption konuları
  3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.pdf   <- Application altında nesne hiyerarşisi
  ARAMA_REHBERI.md                            <- Hangi konu hangi PDF'te; PDF'de Ctrl+F rehberi
```

Proje **docs** klasöründe **VBA_API_REFERENCE.md** vardır; sık kullanılan API'lerin özeti ve Help PDF'lerine yönlendirme.

════════════════════════════════════════════════════════════════════════════════

## Hangi PDF'i ne zaman kullanacaksınız?

| İhtiyaç / Soru | Kullanılacak PDF | Ne zaman |
|----------------|-------------------|----------|
| **Kod nasıl yazılır?** (girinti, yorum, isimlendirme, Option Explicit, başlık) | **Help-Automation Development Guidelines.pdf** | Kod yazmaya başlamadan önce; 11. dokümanla birlikte. |
| **Makro hazırlık aşamaları** (Design, Draft, Harden, Finalize), TAMAM/HAZIR listesi, ihtiyaç analizi | **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf** | Proje planlama ve teslim öncesi; 10. ve 11. dokümanla birlikte. |
| **Nesne modeli** (Application, Editors, ActiveEditor, FileSystem, Documents, Part, Product, Drawing) | **Help-Native Apps Automation.pdf** | 6. ve 8. dokümanı okurken; "buna nereden ulaşırım?" diye sorduğunuzda. |
| **Servisler** (GetSessionService, GetService, Service Identifier listesi) | **Help-Common Services.pdf**, **Help-Native Apps Automation.pdf** | 12. doküman (Servisler) ve editor/session servis kullanırken. |
| **Belirli bir sınıf/metod imzası** (parametreler, dönüş tipi, örnek) | **VBA_API_REFERENCE.md** (öncelik), **Help-Automation Reference.pdf**, **Help-Native Apps Automation.pdf** | 7. (kayıt sonrası), 8., 13. dokümanlarda; API adını biliyorsanız. |
| **Hiyerarşi ağacı** (Application altında ne var, Part/Product/Drawing altında ne var) | **3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.pdf** | 6. doküman (Nesne modeli) ve 13. (Erişim rehberi) ile birlikte. |
| **Hata yönetimi, log, Err.Raise** (resmi öneriler) | **Help-Automation Development Guidelines.pdf**, **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf** | 9. ve 11. dokümanlarda hata/log tasarımı yaparken. |

════════════════════════════════════════════════════════════════════════════════

## Nasıl kullanacaksınız? (PDF'de arama)

- **PDF okuyucunuzda (Adobe, Foxit, tarayıcı vb.)** ilgili PDF'i açın; **Ctrl+F** (veya Cmd+F) ile arama yapın. Aradığınız kelime: örn. `Part`, `GetSessionService`, `ActiveDocument`, `Parameters`, `Shapes`.
- **Hangi PDF'te arayacağınızı bilmiyorsanız:** [Help/ARAMA_REHBERI.md](../Help/ARAMA_REHBERI.md) içindeki "Belirli bir konuyu hangi PDF'te bulacağınız" tablosunu kullanın.
- **Önce proje referansı:** Belirli bir API için önce **docs/VBA_API_REFERENCE.md** dosyasına bakın; yetersizse ilgili Help PDF'ini açıp Ctrl+F ile arayın.

════════════════════════════════════════════════════════════════════════════════

## Hangi aşamada hangi PDF?

| Rehber adımı / Aşama | Açın / Kullanın |
|----------------------|------------------|
| **01-02 (Giriş, Ortam)** | Zorunlu değil; merak ederseniz Help-Automation Development Guidelines (genel bakış). |
| **03-05 (VBA temelleri)** | Help-Automation Development Guidelines (isimlendirme, Option Explicit, yorum kuralları). |
| **06 (Nesne modeli)** | Help-Native Apps Automation, 3DEXPERIENCE Otomasyon Hiyerarşi Ağacı; Application, ActiveEditor, FileSystem bölümleri. |
| **07 (Makro kayıt)** | Kayıt çıktısındaki sınıf/method adlarını VBA_API_REFERENCE.md veya ilgili Help PDF'inde Ctrl+F ile arayın; Help-Automation Development Guidelines (kayıt sonrası düzenlemeler). |
| **08 (Sık kullanılan API)** | VBA_API_REFERENCE.md, Help-Native Apps Automation, Help-Common Services (servisler için). |
| **09 (Hata, debug)** | Help-Automation Development Guidelines, MACRO HAZIRLIK YÖNERGESİ (hata sınıflandırma, log). |
| **10-11 (Örnek proje, Resmi kurallar)** | 3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ (fazlar, TAMAM/HAZIR), Help-Automation Development Guidelines (başlık, teslim). |
| **12 (Servisler)** | Help-Common Services, Help-Native Apps Automation (Service Identifier, GetSessionService). |
| **13-16 (Erişim, Excel, Dosya, İyileştirme)** | Gerekirse VBA_API_REFERENCE.md ve Help-Native Apps Automation (erişim yolu); diğerleri rehberde yeterli. |

════════════════════════════════════════════════════════════════════════════════

## Özet: "Şimdi ne yapayım?"

1. **Kod yazmaya başlıyorum** → 11. doküman + **Help-Automation Development Guidelines.pdf** (başlık, isimlendirme, Option Explicit).
2. **Application / Part / Product nereden alınır bilmiyorum** → 6. ve 13. doküman + **Help-Native Apps Automation.pdf**, **3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.pdf**.
3. **Kayıt ettim, çıkan kodu anlamadım** → Sınıf/method adını **VBA_API_REFERENCE.md** veya ilgili **Help PDF**'inde Ctrl+F ile arayın; 7. doküman (sadeleştirme).
4. **Servis kullanacağım (GetSessionService, GetService)** → 12. doküman + **Help-Common Services.pdf**, **Help-Native Apps Automation.pdf** (Service Identifier).
5. **Teslim / kurumsal standart** → 11. doküman + **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf** (fazlar, TAMAM/HAZIR).

════════════════════════════════════════════════════════════════════════════════

## Uygulamalı alıştırma – Yaparak öğren

**Amaç:** Help klasöründe ilgili PDF'i bulup içinde arama yapmak.  
**Süre:** Yaklaşık 10 dakika.  
**Zorluk:** Orta

| Adım | Ne yapacaksınız | Kontrol |
|------|------------------|--------|
| **1** | Help klasörünü açın. "Parameters" veya "Parametre" konusu için [ARAMA_REHBERI.md](../Help/ARAMA_REHBERI.md) tablosuna bakın; hangi PDF'te anlatıldığını bulun. O PDF'i açın. | Doğru PDF seçildi mi? |
| **2** | PDF'de **Ctrl+F** ile "GetItem" veya "Update" araması yapın. İlgili bölümü okuyun. | İlgili sayfa bulundu mu? |
| **3** | Bu dokümandaki "Hangi PDF'i ne zaman kullanacaksınız?" tablosuna bakın. "Belirli bir sınıf/metod imzası" ihtiyacı için hangi PDF öneriliyor? O PDF'i açıp bir arama yapın. | Doğru kaynak kullanıldı mı? |
| **4** | [Help/ARAMA_REHBERI.md](../Help/ARAMA_REHBERI.md) içindeki "Hangi konu hangi PDF'te" tablosundan bir konu seçip ilgili PDF'te Ctrl+F ile arayın. | Arama rehberi kullanıldı mı? |

**Beklenen sonuç:** En az bir Help PDF'i açıldı; içinde arama yapıldı ve ilgili bölüm bulundu.

════════════════════════════════════════════════════════════════════════════════

## İlgili dokümanlar

**Tüm rehber:** [README](README.md). **Resmi kurallar özeti:** [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md). **Nesne modeli / erişim:** [06-3DExperience-Nesne-Modeli.md](06-3DExperience-Nesne-Modeli.md), [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md). **Sık hatalar:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md). **API özeti (docs):** [VBA_API_REFERENCE.md](../docs/VBA_API_REFERENCE.md).

---

### Gezinme

| [← Önceki: 16 İyileştirme](16-Iyilestirme-Onerileri.md) | [Rehber listesi](README.md) | [Sonraki: 18 Sık hatalar →](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) |
| :--- | :--- | :--- |
