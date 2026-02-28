# 3DExperience VBA Macro – Rehber ve Referans

**Anahtar kelimeler:** 3DExperience, VBA, macro, Dassault Systèmes, CATIA, automation, 3DEX, makro rehberi.

**Rehber sürümü:** v1.1

```
██████╗ ██████╗ ███████╗██╗  ██╗     ████████╗██████╗ 
╚════██╗██╔══██╗██╔════╝╚██╗██╔╝     ╚══██╔══╝██╔══██╗
 █████╔╝██║  ██║█████╗   ╚███╔╝         ██║   ██████╔╝
 ╚═══██╗██║  ██║██╔══╝   ██╔██╗         ██║   ██╔══██╗
██████╔╝██████╔╝███████╗██╔╝ ██╗███████╗██║   ██║  ██║
╚═════╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝   ╚═╝  ╚═╝
```

**3DExperience** (Dassault Systèmes) platformunda **VBA ile makro** yazmak için sıfırdan rehber, API referansı ve yardımcı dokümanlar. Yeni başlayanlar için adım adım anlatım, sık kullanılan kalıplar ve resmi Help dokümanlarıyla uyumlu içerik.

---

## İçindekiler

| Bölüm | Açıklama |
|-------|----------|
| [**Guidelines**](Guidelines/README.md) | 19 dokümanlık kod yazma rehberi (VBA temelleri, nesne modeli, makro kayıt, hata yakalama, örnek projeler, resmi kurallar, isimlendirme) |
| [**VBA API Referansı**](VBA_API_REFERENCE.md) | Sık kullanılan API imzaları + açıklamalar; tam liste `Help/VBA_CALL_LIST.txt` ve `Help/text/` |
| [**Examples**](Examples/README.md) | Çalıştırılabilir örnek makrolar (`.bas`); rehberle uyumlu |
| [**Help**](Help/) | Ham API (`VBA_CALL_LIST.txt`), özet (`SIK_KULLANILAN_API.txt`), arama rehberi (`ARAMA_REHBERI.md`), `text/` |
| [**QUICK_START**](QUICK_START.md) | İlk 5 dakikada tek sayfa hızlı başlangıç |
| [**Terimler**](GLOSSARY.md) | Sözlük (Part, Parameter, Nothing, Update, Editor/Session-level vb.) |
| [**SSS / FAQ**](FAQ.md) | Sık sorulan sorular |
| [**Sorun giderme**](TROUBLESHOOTING.md) | Hata senaryoları ve çözüm önerileri |

---

## Hızlı başlangıç

1. **Rehberi takip et:** [Guidelines/README.md](Guidelines/README.md) → 1. dokümandan başlayıp sırayla ilerleyin.
2. **API’ye bak:** [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) (varsa) veya `Help/VBA_CALL_LIST.txt`, `Help/text/` içindeki dosyalar.
3. **Sık hataları özetle:** [Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).
4. **Teslim / kod incelemesi öncesi:** [Guidelines/VBA-Kod-Checklist.md](Guidelines/VBA-Kod-Checklist.md) dosyasındaki zorunlu ve önerilen maddeleri mutlaka kontrol edin (Option Explicit, Nothing, tek Update, başlık, On Error).

---

## Proje yapısı

```
.
├── README.md                 ← Bu dosya (proje sayfası)
├── QUICK_START.md            ← İlk 5 dakika hızlı başlangıç
├── GLOSSARY.md               ← Terimler sözlüğü
├── FAQ.md                    ← Sık sorulan sorular
├── TROUBLESHOOTING.md        ← Sorun giderme
├── VBA_API_REFERENCE.md      ← API referansı (sık kullanılan imzalar + Help kaynakları)
├── .gitignore
├── LICENSE
├── Guidelines/                ← Kod yazma rehberi (19 doküman + önceki/sonraki gezinme)
│   ├── README.md
│   ├── 01-Giris…md … 18-Sik-Hatalar…md
│   └── VBA-Kod-Checklist.md   ← Teslim öncesi mutlaka kontrol et
├── Examples/                  ← Örnek makrolar (.bas)
│   ├── README.md
│   └── *.bas
└── Help/                      ← Referans ve ham veri
    ├── VBA_CALL_LIST.txt      ← Çağrılabilir API listesi
    ├── SIK_KULLANILAN_API.txt ← Sık kullanılan API özeti
    ├── ARAMA_REHBERI.md       ← grep/arama örnekleri
    ├── API_REPORT.csv
    ├── *.pdf                  ← Resmi Help PDF’leri (isteğe bağlı)
    └── text/                  ← Metin (.txt) versiyonları
```

---

## Guidelines (rehber) özeti

- **01–05:** VBA temelleri (değişkenler, koşullar, döngüler, prosedürler).
- **06:** 3DExperience nesne modeli (Application, Document, Part, Product, Drawing, FileSystem).
- **07–08:** Makro kayıt, inceleme ve sık kullanılan API’ler.
- **09–10:** Hata yakalama, debug ve baştan sona örnek makro.
- **11:** Resmi kurallar ve hazırlık fazları (Design/Draft/Harden/Finalize).
- **12–13:** Servisler (Editor/Session), erişim ve kullanım rehberi.
- **14–15:** VBA–Excel etkileşimi, dosya seçme/kaydetme diyalogları.
- **16–18:** İyileştirme önerileri, Help dosyalarını kullanma, sık hatalar ve dikkat noktaları.

Tam liste ve tablolar: **[Guidelines/README.md](Guidelines/README.md)**.
