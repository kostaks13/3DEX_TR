# 3DExperience VBA Macro – Rehber ve Referans

**3DExperience** (Dassault Systèmes) platformunda **VBA ile makro** yazmak için sıfırdan rehber, API referansı ve yardımcı dokümanlar. Yeni başlayanlar için adım adım anlatım, sık kullanılan kalıplar ve resmi Help dokümanlarıyla uyumlu içerik.

---

## İçindekiler

| Bölüm | Açıklama |
|-------|----------|
| [**Guidelines**](Guidelines/README.md) | 18 dokümanlık kod yazma rehberi (VBA temelleri, nesne modeli, makro kayıt, hata yakalama, örnek projeler, resmi kurallar) |
| [**VBA API Referansı**](VBA_API_REFERENCE.md) | Çağrılabilir API listesi ve kısa açıklamalar (proje kökünde) |
| [**Help**](Help/) | Ham API listesi (`VBA_CALL_LIST.txt`), rapor (`API_REPORT.csv`) ve PDF’lerden çevrilmiş metin dosyaları (`text/`) |

---

## Hızlı başlangıç

1. **Rehberi takip et:** [Guidelines/README.md](Guidelines/README.md) → 1. dokümandan başlayıp sırayla ilerleyin.
2. **API’ye bak:** [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) (varsa) veya `Help/VBA_CALL_LIST.txt`, `Help/text/` içindeki dosyalar.
3. **Sık hataları özetle:** [Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

---

## Proje yapısı

```
.
├── README.md                 ← Bu dosya (proje sayfası)
├── VBA_API_REFERENCE.md      ← API referansı (varsa)
├── .gitignore
├── Guidelines/               ← Kod yazma rehberi (18 doküman + checklist)
│   ├── README.md             ← Rehber giriş ve doküman listesi
│   ├── 01-Giris-Neden-3DExperience-VBA.md
│   ├── ...
│   ├── 18-Sik-Hatalar-ve-Dikkat-Edilecekler.md
│   └── VBA-Kod-Checklist.md
├── Help/                     ← Referans ve ham veri
│   ├── VBA_CALL_LIST.txt
│   ├── API_REPORT.csv
│   └── text/                 ← PDF’lerden çevrilmiş metin dosyaları
└── scripts/                  ← Yardımcı scriptler (varsa)
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

---

## Gereksinimler

- **3DExperience** (Native Client) kurulu ve lisanslı ortam.
- Makro yazmak için **VBA** editörü erişimi (Tools → Macro → Edit vb.).
- Rehber ve Help metinleri için herhangi bir metin/markdown okuyucu.

---

## Katkı ve lisans

- **Katkı:** Hata düzeltmesi veya öneri için Issue açabilir veya Merge Request gönderebilirsiniz.
- **Lisans:** Proje içeriği eğitim ve referans amaçlıdır. 3DExperience ve ilgili ticari markalar Dassault Systèmes’e aittir. Dokümanların kullanım koşulları için repository’deki `LICENSE` dosyasına bakın.

---

## Bağlantılar

- [Dassault Systèmes 3DExperience](https://www.3ds.com/products-services/3dexperience/)
- Rehberin tamamı: **[Guidelines/README.md](Guidelines/README.md)**
