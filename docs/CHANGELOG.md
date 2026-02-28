# Sürüm notları (Changelog)

```
┌─────────────────────────────────────────────────────────────────┐
│  v1.2 · v1.1 · v1.0  |  Örnekler, dokümantasyon, alıştırmalar   │
└─────────────────────────────────────────────────────────────────┘
```

Rehber ve proje içeriğinde yapılan **önemli güncellemelerin** kısa listesi. Detay için ilgili commit mesajlarına bakın.

---

## [v1.2] – Örnekler checklist uyumu, başlık tutarlılığı, dokümantasyon, uygulamalı alıştırmalar

- **UI / Görünüm:** Tüm ana sayfalara tutarlı başlık blokları (┌──┐), blockquote vurguları ve "Gezinme" linkleri eklendi. README’de hızlı başlangıç ve link kontrolü için kısa çağrı kutuları; docs/, QUICK_START, FAQ, GLOSSARY, TROUBLESHOOTING, CHEATSHEET, CHANGELOG’da ortak görsel dil.
- **Link kontrolü:** Tüm Markdown dosyalarındaki linkleri kontrol etmek için `scripts/check-links.sh` ve `npm run check-links` eklendi; `markdown-link-check` ile yerel çalıştırma. Yapılandırma: `scripts/mlc-config.json`. README’de kullanım notu.
- **Sözlük:** [GLOSSARY.md](GLOSSARY.md) genişletildi: ActiveEditor, ByVal/ByRef, Collection, Exit Sub/Exit Function, GetObject, Item, Children, On Error, Option Explicit, Sheets, Views, Value ve E–L bölümü eklendi.
- **Guidelines 01:** "Uygulamalı alıştırma – Yaparak öğren" bölümü eklendi; 2. dokümana geçiş ve oradaki ilk makro alıştırmasını adım adım yönlendiriyor.
- **Örnekler:** [ParametreListesiniDosyayaYaz.bas](../Examples/ParametreListesiniDosyayaYaz.bas) VBA-Kod-Checklist ile tam uyumlu hale getirildi (On Error, adım adım Nothing kontrolü, tam başlık, Const ile çıktı yolu). Tüm örnek .bas dosyalarına **Assumptions** ve **Regional Settings** başlık alanları eklendi.
- **Hata yönetimi:** SadecePartKontrol, IkiParametreTakas, MinMaxParametreDeger, ShapesBilgisi ve GetActivePart_AnaParametreListesi’nde On Error GoTo ve adım adım oApp/oDoc/oPart kontrolü eklendi. [LogOrnekMakro.bas](../Examples/LogOrnekMakro.bas) içinde LogSatir, dosya yazma hatasında makroyu çökertmeyecek şekilde On Error Resume Next ile korundu.
- **Dokümantasyon:** README’de Help klasörü açıklaması netleştirildi (text/ mevcut, PDF isteğe bağlı). Guidelines özetinde 19. doküman (isimlendirme rehberi) vurgulandı. Yardımcı dokümanlar docs/ klasörüne taşındı; linkler güncellendi.
- **Uygulamalı alıştırmalar:** Guidelines 02–19 dokümanlarına "Uygulamalı alıştırma – Yaparak öğren" bölümleri eklendi; adım tabloları ile aşama aşama ilerleme.

---

## [v1.1] – İçerik, örnekler, yapı, kalite

- **İçerik:** [GLOSSARY.md](GLOSSARY.md) (terimler sözlüğü), [FAQ.md](FAQ.md) (sık sorulan sorular), [TROUBLESHOOTING.md](TROUBLESHOOTING.md) (sorun giderme) eklendi.
- **Örnekler:** [QUICK_START.md](QUICK_START.md) (ilk 5 dakika); Examples’a SadecePartKontrol, IkiParametreTakas, LogOrnekMakro, MinMaxParametreDeger eklendi; her örnek için “Beklenen çıktı” açıklaması.
- **Yapı:** Her Guidelines dokümanında (01–19) “Bu dokümanda” özeti; README ve Guidelines/README’de VBA-Kod-Checklist vurgusu; 19. doküman (İsimlendirme Rehberi) rehber setine dahil.
- **Kalite:** CHANGELOG.md; rehber sürümü README’de v1.1 olarak güncellendi. VBA kod blokları rehberde ```vba dil etiketi ile işaretlidir (tutarlı sözdizimi vurgusu).

---

## [v1.0] – İlk tam rehber seti

- 18 dokümanlık Guidelines rehberi; VBA_API_REFERENCE genişletmesi; Examples klasörü ve örnek makrolar; Help (SIK_KULLANILAN_API, ARAMA_REHBERI); önceki/sonraki gezinme; ASCII diyagramları; README (anahtar kelimeler, sürüm, dağıtım/PDF).

---

*Tarih formatı: YYYY-MM-DD. Sürüm numaraları [Semantic Versioning](https://semver.org/) benzeri (rehber için Major.Minor) kullanılabilir.*
