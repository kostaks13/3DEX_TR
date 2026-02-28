# Sık Sorulan Sorular (SSS)

```
┌──────────────────────────────────────────────────────────────────────────────┐
│  Genel · Hata ve davranış · API ve Help · Teslim ve kalite                    │
└──────────────────────────────────────────────────────────────────────────────┘
```

3DExperience VBA makro rehberiyle ilgili **sık sorulan sorular** ve kısa yanıtlar. Detay için linkli dokümanlara bakın.

**Bu sayfada:** Genel · Hata ve davranış · API ve Help · Teslim ve kalite

---

## Genel

> **Rehberi nereden başlamalıyım?**  
[Guidelines/README.md](../Guidelines/README.md) → 1. dokümandan (01-Giris) sırayla ilerleyin. Hızlı denemek için [QUICK_START.md](QUICK_START.md) ve [Examples/](../Examples/README.md) kullanın.

> **Parametre nasıl değiştirilir?**  
`oPart.Parameters.Item("ParametreAdi").Value = yeniDeger` yazıp ardından `oPart.Update` çağırın (döngü dışında bir kez). Örnek: [Examples/ParametreYaz.bas](../Examples/ParametreYaz.bas), [Guidelines/08-Sik-Kullanilan-APIler.md](../Guidelines/08-Sik-Kullanilan-APIler.md).

> **Excel’den Part parametrelerine nasıl yazarım?**  
Excel’i `CreateObject("Excel.Application")` veya `GetObject` ile açın; hücreleri okuyup Part parametrelerine `.Value = ...` ile yazın. Tam akış: [Guidelines/14-VBA-ve-Excel-Etkilesimi.md](../Guidelines/14-VBA-ve-Excel-Etkilesimi.md).

---

## Hata ve davranış

**Makro “Nothing” veya “Object variable not set” diyor. Neden?**  
Uygulama kapalı, belge yok veya `GetItem("Part")` çizim/montaj belgesinde Part döndürmedi. Her `Set` sonrası `If ... Is Nothing Then MsgBox "...": Exit Sub` ekleyin. Bkz. [Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](../Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md), [TROUBLESHOOTING.md](TROUBLESHOOTING.md).

> **Update çağırdım ama ekranda değişiklik görünmüyor.**  
Değişiklikleri yaptıktan sonra tek bir `oPart.Update` çağrısı yeterli; bazen görüntü yenilemesi veya belge türü (read-only) etkiler. [TROUBLESHOOTING.md](TROUBLESHOOTING.md) içinde “Update sonrası değişiklik yok” maddesine bakın.

> **GetItem("Part") Part bulamıyor / hata veriyor.**  
Aktif belge Part değildir (çizim veya montaj açık olabilir). Önce `oDoc.GetItem("Part")` sonrası `If oPart Is Nothing Then` ile kontrol edin; kullanıcıyı “Lütfen bir parça belgesi açın” diye uyarın. Örnek: [Examples/SadecePartKontrol.bas](../Examples/SadecePartKontrol.bas).

---

## API ve Help

**İstediğim API’nin tam imzasını nerede bulurum?**  
[VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) (sık kullanılanlar); tam liste için [Help/VBA_CALL_LIST.txt](../Help/VBA_CALL_LIST.txt) ve [Help/text/](../Help/). Arama: [Help/ARAMA_REHBERI.md](../Help/ARAMA_REHBERI.md).

> **Editor-level ve Session-level servis farkı ne?**  
Editor-level: aktif Part/Product/Drawing’e bağlı (InertiaService vb.). Session-level: oturum geneli (Search, PLMOpenService vb.). Editor-level kullanırken session-level çağırma; önce editor işini bitir. Bkz. [Guidelines/12-Servisler-ve-Yapilabilecek-Islemler.md](../Guidelines/12-Servisler-ve-Yapilabilecek-Islemler.md).

---

## Teslim ve kalite

> **Teslim / kod incelemesi öncesi ne kontrol etmeliyim?**  
[Guidelines/VBA-Kod-Checklist.md](../Guidelines/VBA-Kod-Checklist.md) dosyasındaki zorunlu ve önerilen maddeleri uygulayın: Option Explicit, Nothing/Count, tek Update, başlık (Language, Release), On Error. Özet: [Guidelines/11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](../Guidelines/11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md).

> **Terimlerin kısa tanımı nerede?**  
[GLOSSARY.md](GLOSSARY.md) — Part, Parameter, Nothing, Update, Editor-level, Session-level vb.

> **Sorun giderme (hata senaryoları) nerede?**  
[TROUBLESHOOTING.md](TROUBLESHOOTING.md) — “Makro çalışmıyor”, “Update sonrası değişiklik yok”, “GetItem Part bulamıyor” ve diğer senaryolar.

---

**İlgili:** [QUICK_START.md](QUICK_START.md) · [CHEATSHEET.md](CHEATSHEET.md) · [TROUBLESHOOTING.md](TROUBLESHOOTING.md) · [Guidelines/README.md](../Guidelines/README.md)
