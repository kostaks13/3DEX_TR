# Help PDF'lerinde Arama Rehberi

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  Hangi konu hangi PDF'te?  |  PDF'de Ctrl+F ile arama                       ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

Bu doküman, **Help** klasöründeki **resmi referans PDF'lerinde** arama yaparken kullanacağınız rehberdir. Tüm referanslar **PDF** üzerinedir; hangi konuyu hangi PDF'te bulacağınız ve PDF'de nasıl arayacağınız anlatılır.

> **İpucu:** Önce aşağıdaki "Belirli bir konuyu hangi PDF'te bulacağınız" tablosunu kullanın; sonra ilgili PDF'i açıp **Ctrl+F** (veya Cmd+F) ile arayın.

---

## 1. Belirli bir konuyu hangi PDF'te bulacağınız

| Aradığınız | Önerilen PDF |
|------------|-------------------------------|
| Nesne modeli (Application, Part, Product, Document) | **Help-Native Apps Automation.pdf**, **3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.pdf** |
| Servis listesi (GetService, GetSessionService, Service Identifier) | **Help-Common Services.pdf**, **Help-Native Apps Automation.pdf** |
| Kod kuralları (Option Explicit, başlık, isimlendirme) | **Help-Automation Development Guidelines.pdf** |
| Makro hazırlık fazları (Design, Draft, Harden, Finalize) | **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf** |
| Parametre / Parameters, Value, Update | **Help-Native Apps Automation.pdf**, **Help-Automation Reference.pdf** |
| Shapes, MainBody, geometri | **Help-Native Apps Automation.pdf**, **Help-Automation Reference.pdf** |
| Drawing, Sheets, Views | **Help-Native Apps Automation.pdf** |
| FileSystem, dosya/klasör işlemleri | **Help-Native Apps Automation.pdf** |
| Hata yönetimi, log, Err.Raise | **Help-Automation Development Guidelines.pdf**, **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf** |

---

## 2. PDF'de arama (Ctrl+F / Cmd+F)

1. Help klasöründen ilgili **PDF'i** açın (Adobe Reader, Foxit, tarayıcı vb.).
2. **Ctrl+F** (Windows/Linux) veya **Cmd+F** (macOS) ile arama kutusunu açın.
3. Aradığınız terimi yazın: örn. `GetItem`, `Parameters`, `ActiveDocument`, `Update`, `Shapes`, `GetService`.
4. Gerekirse "Sonraki" / "Önceki" ile tüm eşleşmeleri gezin.

**Ne arayacaksınız?** Makro kaydında çıkan sınıf veya metod adı, rehberde geçen API adı (Parameters, Shapes, Children, Service Identifier vb.).

---

## 3. Pratik örnekler

| Sorunuz | Hangi PDF | PDF'de arayın |
|---------|-----------|----------------|
| Parametre değeri nasıl yazarım? | Help-Native Apps Automation, Help-Automation Reference | `Parameter`, `Value`, `Update` |
| Shapes koleksiyonuna nasıl erişirim? | Help-Native Apps Automation, Help-Automation Reference | `Shapes`, `MainBody` |
| Hangi servisleri GetService ile alabilirim? | Help-Common Services, Help-Native Apps Automation | `GetService`, `Service Identifier` |
| Drawing sayfaları ve Views | Help-Native Apps Automation | `Sheet`, `View`, `Drawing` |
| FileDialog veya dosya yolu | Help-Native Apps Automation | `FileSystem`, `File`, `Folder` |

---

## 4. Proje içi API özeti

- **docs/VBA_API_REFERENCE.md** — Sık kullanılan API'lerin kısa listesi ve örnek VBA çağrıları. Detay için ilgili Help PDF'ine yönlendirir.
- Önce bu dosyaya bakın; yetersizse yukarıdaki tabloya göre ilgili PDF'i açıp Ctrl+F ile arayın.

---

## İlgili dokümanlar

- **Help dosyalarını ne zaman kullanacağınız:** [Guidelines/17-Help-Dosyalarini-Kullanma.md](../Guidelines/17-Help-Dosyalarini-Kullanma.md)
- **API referansı (proje):** [docs/VBA_API_REFERENCE.md](../docs/VBA_API_REFERENCE.md)

**Gezinme:** [Ana sayfa](../README.md) · [Docs](../docs/README.md) · [Rehber](../Guidelines/README.md) · [Örnek makrolar](../Examples/README.md)
