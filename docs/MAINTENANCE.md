# Bakım ve otomasyon

Proje dokümantasyonu için kullanılan araçlar ve commit öncesi önerilen adımlar.

---

## Komutlar

| Komut | Açıklama |
|-------|----------|
| `npm install` | Bağımlılıkları yükler (link kontrolü, lint, spell için gerekli). |
| `npm run check-links` | Tüm `.md` dosyalarındaki linkleri kontrol eder. |
| `npm run lint:md` | Markdown stil kurallarını kontrol eder (`.markdownlint.json`). |
| `npm run spell` | Türkçe + proje sözlüğü ile yazım kontrolü (`cspell.json`). |

---

## Commit öncesi (isteğe bağlı)

Değişiklikleri pushlamadan önce yerel ortamda şunları çalıştırabilirsiniz:

```bash
npm install
npm run check-links
npm run lint:md
npm run spell
```

**Pre-commit hook** kullanmak isterseniz (ör. Husky):

1. `npx husky init` ile husky kurun.
2. `.husky/pre-commit` içine `npm run check-links` ve isteğe bağlı `npm run lint:md` ekleyin.

Hook kurmadan da CI (GitHub Actions) her push/PR'da link kontrolü yapar.

---

## CI (GitHub Actions)

- **Workflow:** `.github/workflows/check-links.yml`
- **Tetikleyici:** `main` dalına push veya PR.
- **Adımlar:** `npm ci` → `npm run check-links`. Link hatası varsa workflow başarısız olur.

---

## Yapılandırma dosyaları

| Dosya | Açıklama |
|-------|----------|
| `scripts/mlc-config.json` | markdown-link-check ayarları (timeout, yok sayılan desenler). |
| `.markdownlint.json` | markdownlint kuralları (satır uzunluğu vb. kapatılabilir). |
| `cspell.json` | cspell: dil (tr), proje sözcükleri, yok sayılan dosyalar. |
