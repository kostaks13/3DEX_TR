# 3DExperience VBA Macro – Rehber ve Referans

```
 333333333333333   DDDDDDDDDDDDD      EEEEEEEEEEEEEEEEEEEEEEXXXXXXX       XXXXXXX                        TTTTTTTTTTTTTTTTTTTTTTTRRRRRRRRRRRRRRRRR
3:::::::::::::::33 D::::::::::::DDD   E::::::::::::::::::::EX:::::X       X:::::X                        T:::::::::::::::::::::TR::::::::::::::::R
3::::::33333::::::3D:::::::::::::::DD E::::::::::::::::::::EX:::::X       X:::::X                        T:::::::::::::::::::::TR::::::RRRRRR:::::R
3333333     3:::::3DDD:::::DDDDD:::::DEE::::::EEEEEEEEE::::EX::::::X     X::::::X                        T:::::TT:::::::TT:::::TRR:::::R     R:::::R
            3:::::3  D:::::D    D:::::D E:::::E       EEEEEEXXX:::::X   X:::::XXX                        TTTTTT  T:::::T  TTTTTT  R::::R     R:::::R
            3:::::3  D:::::D     D:::::DE:::::E                X:::::X X:::::X                                   T:::::T          R::::R     R:::::R
    33333333:::::3   D:::::D     D:::::DE::::::EEEEEEEEEE       X:::::X:::::X                                    T:::::T          R::::RRRRRR:::::R
    3:::::::::::3    D:::::D     D:::::DE:::::::::::::::E        X:::::::::X                                     T:::::T          R:::::::::::::RR
    33333333:::::3   D:::::D     D:::::DE:::::::::::::::E        X:::::::::X                                     T:::::T          R::::RRRRRR:::::R
            3:::::3  D:::::D     D:::::DE::::::EEEEEEEEEE       X:::::X:::::X                                    T:::::T          R::::R     R:::::R
            3:::::3  D:::::D     D:::::DE:::::E                X:::::X X:::::X                                   T:::::T          R::::R     R:::::R
            3:::::3  D:::::D    D:::::D E:::::E       EEEEEEXXX:::::X   X:::::XXX                                T:::::T          R::::R     R:::::R
3333333     3:::::3DDD:::::DDDDD:::::DEE::::::EEEEEEEE:::::EX::::::X     X::::::X                              TT:::::::TT      RR:::::R     R:::::R
3::::::33333::::::3D:::::::::::::::DD E::::::::::::::::::::EX:::::X       X:::::X                              T:::::::::T      R::::::R     R:::::R
3:::::::::::::::33 D::::::::::::DDD   E::::::::::::::::::::EX:::::X       X:::::X                              T:::::::::T      R::::::R     R:::::R
 333333333333333   DDDDDDDDDDDDD      EEEEEEEEEEEEEEEEEEEEEEXXXXXXX       XXXXXXX                              TTTTTTTTTTT      RRRRRRRR     RRRRRRR
                                                                                 ________________________
                                                                                 _::::::::::::::::::::::_
                                                                                 ________________________
```

**3DExperience** (Dassault Systèmes) platformunda **VBA ile makro** yazmak için sıfırdan rehber, API referansı ve yardımcı dokümanlar. Yeni başlayanlar için adım adım anlatım, sık kullanılan kalıplar ve resmi Help dokümanlarıyla uyumlu içerik.

---

## İçindekiler

| Bölüm | Açıklama |
|-------|----------|
| [**Guidelines**](Guidelines/README.md) | 18 dokümanlık kod yazma rehberi (VBA temelleri, nesne modeli, makro kayıt, hata yakalama, örnek projeler, resmi kurallar) |
| [**VBA API Referansı**](VBA_API_REFERENCE.md) | Çağrılabilir API listesi (varsa); yoksa `Help/VBA_CALL_LIST.txt` ve `Help/text/` kullanın |
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
├── VBA_API_REFERENCE.md      ← API referansı (isteğe bağlı; yoksa Help/ kullanın)
├── .gitignore
├── LICENSE
├── Guidelines/               ← Kod yazma rehberi (18 doküman + checklist)
│   ├── README.md             ← Rehber giriş ve doküman listesi
│   ├── 01-Giris-Neden-3DExperience-VBA.md … 18-Sik-Hatalar-ve-Dikkat-Edilecekler.md
│   └── VBA-Kod-Checklist.md
└── Help/                     ← Referans ve ham veri
    ├── VBA_CALL_LIST.txt      ← Çağrılabilir API listesi
    ├── API_REPORT.csv
    ├── *.pdf                  ← Resmi Help PDF’leri (Automation, Native Apps, Common Services vb.)
    └── text/                  ← Aynı içeriğin metin (.txt) versiyonları
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
