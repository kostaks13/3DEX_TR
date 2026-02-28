<div align="center">

# 3DExperience VBA Macro

### Rehber ve Referans

*3DExperience (Dassault Systèmes) platformunda VBA ile makro yazmayı sıfırdan, adım adım öğrenin.*

`v1.2` · **Türkçe** · 3DEX · CATIA · Otomasyon

---

```
██████╗ ██████╗ ███████╗██╗  ██╗     ████████╗██████╗ 
╚════██╗██╔══██╗██╔════╝╚██╗██╔╝     ╚══██╔══╝██╔══██╗
 █████╔╝██║  ██║█████╗   ╚███╔╝         ██║   ██████╔╝
 ╚═══██╗██║  ██║██╔══╝   ██╔██╗         ██║   ██╔══██╗
██████╔╝██████╔╝███████╗██╔╝ ██╗███████╗██║   ██║  ██║
╚═════╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝   ╚═╝  ╚═╝
```

**Sıfırdan ileri seviyeye** — tekrarlayan işleri otomatikleştirin, parametreleri yönetin, raporlar üretin.

</div>

---

## Nasıl kullanılır?

| Adım | Ne yapmalı? |
| :--- | :--- |
| **1** | Rehberi açıp 01. dokümandan başlayarak sırayla ilerleyin. |
| **2** | Her dokümanda "Uygulamalı alıştırma" bölümünü kendi ortamınızda yapın. |
| **3** | İlk makroyu 02. dokümanda yazıp F5 ile çalıştırın; sonra örneklerden birini deneyin. |
| **4** | API veya terimlerde takılırsanız [Help](Help/) ve [docs/GLOSSARY.md](docs/GLOSSARY.md) kullanın. |

---

## Ne sunar? (özet)

| Özellik | Açıklama |
| :--- | :--- |
| **Rehber** | 19 doküman: VBA temelleri → nesne modeli → makro kayıt → hata yakalama → örnek proje → resmi kurallar → isimlendirme. |
| **Örnekler** | Part, parametre, Shapes, dosyaya yazma, log, modüler yapı; her biri çalıştırılabilir `.bas`. |
| **Referans** | API referansı, hızlı başlangıç, terimler (TR↔EN), sorun giderme, checklist. |
| **Bakım** | `npm run check-links` ile link kontrolü; CI/CD (GitHub Actions) ile otomatik kontroller. |

**Bu sayfada:** [Neden bu rehber?](#neden-bu-rehber) · [Hemen başla](#hemen-başla) · [İçindekiler](#i̇çindekiler) · [Hızlı başlangıç](#hızlı-başlangıç) · [Proje yapısı](#proje-yapısı) · [Guidelines özeti](#guidelines-rehber-özeti)

---

## Neden bu rehber?

| Ne sunar? | Açıklama |
| :--- | :--- |
| **19 doküman** | VBA temellerinden nesne modeline, hata yakalama ve resmi kurallara kadar adım adım rehber. |
| **10+ örnek makro** | Kopyala–yapıştır çalışan `.bas` dosyaları; Part, parametre, Shapes, dosya, log örnekleri. |
| **Help uyumlu** | Dassault Systèmes Automation Guidelines ve 3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ ile uyumlu içerik. |
| **Uygulamalı alıştırmalar** | Her dokümanda "Yaparak öğren" bölümü; kendi ortamınızda deneyerek ilerleyin. |

---

## Hemen başla

> **Yeni başlıyorsanız:** Rehberi 01’den itibaren sırayla takip edin; ilk makroyu 02’de yazıp F5 ile çalıştırın.  
> **Zaten VBA biliyorsanız:** [İlk 5 dk](docs/QUICK_START.md) veya [örnek makrolardan](Examples/README.md) birini açıp hemen deneyin.

| Yeni başlıyorsanız | Zaten VBA biliyorsanız |
| :--- | :--- |
| [Rehber (01→19)](Guidelines/README.md) → 02’de ilk makro, F5 | [İlk 5 dk](docs/QUICK_START.md) |
| [İlerleme listesi](docs/ILERLEME-LISTESI.md) (süre, zorluk, checklist) | [Örnek makrolar](Examples/README.md) |
| | [API referansı](docs/VBA_API_REFERENCE.md) · [Checklist](Guidelines/VBA-Kod-Checklist.md) |

**Tüm linkler:** [Rehber](Guidelines/README.md) · [İlerleme](docs/ILERLEME-LISTESI.md) · [İlk 5 dk](docs/QUICK_START.md) · [Örnekler](Examples/README.md) · [API](docs/VBA_API_REFERENCE.md) · [Checklist](Guidelines/VBA-Kod-Checklist.md) · [Sorun giderme](docs/TROUBLESHOOTING.md)

---

## İçindekiler

| Bölüm | Açıklama |
| :--- | :--- |
| [**Guidelines**](Guidelines/README.md) | 19 dokümanlık kod yazma rehberi (VBA temelleri, nesne modeli, makro kayıt, hata yakalama, örnek projeler, resmi kurallar, isimlendirme) |
| [**Examples**](Examples/README.md) | Çalıştırılabilir örnek makrolar (`.bas`); rehberle uyumlu |
| [**Help**](Help/) | Ham API (`VBA_CALL_LIST.txt`), özet (`SIK_KULLANILAN_API.txt`), arama rehberi (`ARAMA_REHBERI.md`), `text/` (metin dosyaları); PDF’ler isteğe bağlı eklenebilir |
| [**Docs**](docs/) | [API referansı](docs/VBA_API_REFERENCE.md), [hızlı başlangıç](docs/QUICK_START.md), [hızlı referans](docs/CHEATSHEET.md), [terimler](docs/GLOSSARY.md), [FAQ](docs/FAQ.md), [sorun giderme](docs/TROUBLESHOOTING.md), [sürüm notları](docs/CHANGELOG.md) |

---

## Hızlı başlangıç

> **Başlamak için:** Rehberi sırayla takip edin veya hızlı denemek için [docs/QUICK_START.md](docs/QUICK_START.md) ve [Examples/](Examples/README.md) kullanın.

1. **Rehberi takip et:** [Guidelines/README.md](Guidelines/README.md) → 1. dokümandan başlayıp sırayla ilerleyin.
2. **API’ye bak:** [docs/VBA_API_REFERENCE.md](docs/VBA_API_REFERENCE.md) veya `Help/VBA_CALL_LIST.txt`, `Help/text/` içindeki dosyalar.
3. **Sık hataları özetle:** [Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).
4. **Teslim / kod incelemesi öncesi:** [Guidelines/VBA-Kod-Checklist.md](Guidelines/VBA-Kod-Checklist.md) dosyasındaki zorunlu ve önerilen maddeleri mutlaka kontrol edin (Option Explicit, Nothing, tek Update, başlık, On Error).

---

## Proje yapısı

```
.
├── README.md                  ← Bu dosya (proje sayfası)
├── .gitignore
├── LICENSE
├── Guidelines/                ← Kod yazma rehberi (19 doküman + VBA-Kod-Checklist)
├── Examples/                  ← Örnek makrolar (.bas)
├── Help/                      ← API listeleri, arama rehberi, text/
└── docs/                      ← API referansı, hızlı başlangıç, terimler, FAQ, sorun giderme, CHANGELOG
```

> **Bakım:** Tüm Markdown linklerini kontrol etmek için `npm install` sonrası `npm run check-links` çalıştırın → [scripts/check-links.sh](scripts/check-links.sh).

---

## Guidelines (rehber) özeti

- **01–05:** VBA temelleri (değişkenler, koşullar, döngüler, prosedürler).
- **06:** 3DExperience nesne modeli (Application, Document, Part, Product, Drawing, FileSystem).
- **07–08:** Makro kayıt, inceleme ve sık kullanılan API’ler.
- **09–10:** Hata yakalama, debug ve baştan sona örnek makro.
- **11:** Resmi kurallar ve hazırlık fazları (Design/Draft/Harden/Finalize).
- **12–13:** Servisler (Editor/Session), erişim ve kullanım rehberi.
- **14–15:** VBA–Excel etkileşimi, dosya seçme/kaydetme diyalogları.
- **16–19:** İyileştirme önerileri, Help dosyalarını kullanma, sık hatalar ve dikkat noktaları, isimlendirme rehberi.

Tam liste ve tablolar: **[Guidelines/README.md](Guidelines/README.md)**.

---

**Gezinme:** [Docs](docs/README.md) · [Rehber](Guidelines/README.md) · [Örnek makrolar](Examples/README.md) · [Help](Help/)

*Anahtar kelimeler: 3DExperience, VBA, macro, Dassault Systèmes, CATIA, automation, 3DEX, makro rehberi.*
