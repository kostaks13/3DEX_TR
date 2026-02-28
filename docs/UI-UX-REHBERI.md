# Dokümantasyon UI/UX Rehberi

Bu rehber, projedeki Markdown sayfalarında **tutarlı okunabilirlik ve gezinme** için kullanılacak kalıpları tanımlar.

---

## 1. Callout (uyarı / ipucu) blokları

Önemli bilgiyi vurgulamak için aynı format kullanın:

| Tür | Kullanım | Örnek |
|-----|----------|--------|
| **İpucu** | Kısayol, iyi uygulama | `> **İpucu:** Önce 02. dokümanda ilk makroyu çalıştırın.` |
| **Hızlı yol** | Tek adımda gidebilecek yer | `> **Hızlı yol:** [İlk 5 dk](docs/QUICK_START.md) ile hemen deneyin.` |
| **Dikkat** | Sık yapılan hata, dikkat edilmesi gereken | `> **Dikkat:** Update'ü döngü içinde çağırmayın.` |
| **Önemli** | Zorunlu adım, kurumsal kural | `> **Önemli:** Teslim öncesi [VBA-Kod-Checklist](Guidelines/VBA-Kod-Checklist.md) kullanın.` |

- Callout’lar **tek satır** veya **birkaç satır** olabilir; ikinci satırda `>` ile devam edin.
- Her sayfada aynı tür için aynı etiket kullanın (**İpucu**, **Hızlı yol**, **Dikkat**, **Önemli**).

---

## 2. “Bu sayfada” (içindekiler)

Sayfa uzunsa, başta **Bu sayfada** ile bölüm linklerini verin:

```markdown
**Bu sayfada:** [Bölüm 1](#bölüm-1) · [Bölüm 2](#bölüm-2) · [Bölüm 3](#bölüm-3)
```

- Başlık anchor’ları genelde otomatik üretilir (küçük harf, boşluk → tire).
- Çok bölüm varsa satır sonunda tek satırda kalabilir; 5’ten fazla link varsa iki satıra bölünebilir.

---

## 3. İki kullanıcı yolu (yeni / deneyimli)

Ana sayfada **iki net yol** sunun:

- **Yeni başlıyorsanız:** Rehber 01 → 02, ilk makro, F5.
- **Zaten biliyorsanız:** İlk 5 dk, örnek makrolar, API referansı.

Bunu tablo veya iki blok (blockquote / liste) ile gösterin; her yolda **tek tıkla** gidilecek ana link olsun.

---

## 4. Adım numaraları ve akış

Adım adım anlatımlarda:

1. **Kısa başlık** verin (örn. “3 adım”).
2. Gerekirse **ASCII akış** ile görselleştirin: `[1] → [2] → [3]`.
3. Numaralı liste kullanın; her adımda **bir sonraki adım** veya “Neden?” kısaca belirtilebilir.

---

## 5. Gezinme (alt sayfa linkleri)

Her sayfanın altında **Gezinme** satırı olsun:

```markdown
**Gezinme:** [Ana sayfa](../README.md) · [Rehber](../Guidelines/README.md) · [Örnekler](../Examples/README.md)
```

- Sıra: Ana sayfa → ilgili bölümler (Rehber, Docs, Examples, Help).
- Rehber dokümanlarında **Önceki / Sonraki** (01 → 02 → 03) eklenebilir.

---

## 6. Tablolar

- Başlık satırından sonra `|---|` ile ayırın.
- Hücre içi uzun metinleri kısaltın; detayı ayrı paragrafa alın.
- “Süre”, “Zorluk” gibi sütunlar okunabilirliği artırır.

---

## 7. Kod blokları

- Dil etiketi kullanın: ` ```vba `, ` ```bash `.
- Önce **ne yaptığını** bir cümleyle yazın; gerekirse “Beklenen çıktı” ile bitirin.

---

---

## 8. Guidelines sayfa standardı

- **Başlık:** H1 (örn. `# 1. Giriş – …`), altında ASCII kutu (isteğe bağlı), sonra kısa giriş paragrafı.
- **Bu dokümanda:** Tek satır, bölümler **·** (orta nokta) ile ayrılır; örn. `**Bu dokümanda:** A · B · C · D.`
- **İlk adım / Sonraki adım:** Blockquote ile vurgula (`> **İlk adım:** …`).
- **Bölüm ayırıcı:** Uzun `═══` yerine `---` kullanılabilir (hafif görünüm).
- **Alt gezinme:** Sayfa sonunda `---` + `### Gezinme` + tablo: `| [← Önceki: XX](...) | [Rehber listesi](README.md) | [Sonraki: YY →](...) |`

Bu rehber, yeni doküman eklerken veya mevcut sayfaları güncellerken referans alınır. Tutarlılık, okunabilirlik ve gezinme kolaylığını artırır.
