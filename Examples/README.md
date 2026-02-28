# Örnek Makrolar

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  Çalıştırılabilir .bas örnekleri  |  Rehber 08 + 10 ile uyumlu               ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

Bu klasörde **3DExperience VBA** rehberindeki (Guidelines 08 ve 10) örnek makroların **çalıştırılabilir** kopyaları bulunur. Her dosya `.bas` formatındadır; 3DExperience VBA editörüne modül olarak aktarabilir veya içeriği kopyalayıp kendi modülünüze yapıştırabilirsiniz.

> **İlk kez kullanıyorsanız:** Aşağıdaki **Önerilen sıra** tablosundan 1 → 2 → 3 → 4 ile başlayın.

**Öne çıkan:** [AktifParcaBilgisi.bas](AktifParcaBilgisi.bas) (ilk makro) · [ParametreOkuVeGoster.bas](ParametreOkuVeGoster.bas) · [ParametreYaz.bas](ParametreYaz.bas) · [SadecePartKontrol.bas](SadecePartKontrol.bas)

---

## Kullanım (3 adım)

```
  [1] 3DExperience aç   ──►  [2] VBA Editör: Modül ekle, .bas yapıştır   ──►  [3] F5 ile çalıştır
       Part belgesi aç         (Tools → Macro → Edit)                           Mesaj / dosya çıktısı
```

1. **3DExperience** açın; gerekirse bir **Part** belgesi açın.
2. **Tools → Macro → Edit** (veya eşdeğeri) ile VBA editörünü açın; yeni modül ekleyip ilgili `.bas` dosyasının içeriğini yapıştırın.
3. Makroyu çalıştırın (F5 veya Run); mesajları veya çıktı dosyasını kontrol edin.

---

## Önerilen sıra (yeni başlıyorsanız)

Örnekleri şu sırayla deneyin: **1** → **2** → **3** → **4** (temel zincir ve parametre). Sonra ihtiyacınıza göre diğerlerine geçin.

| Sıra | Dosya | Neden bu sıra? |
|------|-------|----------------|
| 1 | AktifParcaBilgisi.bas | Uygulama → Belge → Part zinciri; ilk çalışan makro. |
| 2 | ParametreOkuVeGoster.bas | Parametre okuma; InputBox ile kullanıcı girişi. |
| 3 | ParametreYaz.bas | Parametre yazma + **tek** Update. |
| 4 | SadecePartKontrol.bas | Belge türü kontrolü; hata vermeden çıkış. |
| — | Diğerleri | Shapes, dosyaya yazma, log, modüler yapı (GetActivePart), takas, min/max. |

---

## Dosyalar

| Dosya | Ne yapar | Gereksinim | Beklenen çıktı |
|-------|----------|------------|----------------|
| **AktifParcaBilgisi.bas** | Aktif belgenin adı, tam yolu ve (Part ise) Shapes sayısını MsgBox ile gösterir. | Açık Part belgesi. | MsgBox: belge adı, tam yol, Shapes sayısı. |
| **ParametreOkuVeGoster.bas** | Kullanıcıdan parametre adı (örn. Length.1) alır; değerini mesajla gösterir. | Açık Part, parametre adı. | MsgBox: "Length.1 = 100" benzeri. |
| **ParametreYaz.bas** | Kullanıcıdan parametre adı ve yeni değer alır; Part'ta günceller ve Update çağırır. | Açık Part, yazılabilir parametre. | MsgBox: "… güncellendi."; modelde değer değişir. |
| **ShapesBilgisi.bas** | Shapes sayısı ve ilk 10 şeklin adını listeler. | Açık Part. | MsgBox: "Shapes sayısı: N" + ilk 10 ad. |
| **ParametreListesiniDosyayaYaz.bas** | Tüm parametreleri `Parametre;Değer` formatında `C:\Temp\parametre_listesi.txt` dosyasına yazar. | Açık Part; C:\Temp yazılabilir. | Dosyada "Parametre;Değer" satırları; MsgBox "Liste yazıldı: …". |
| **GetActivePart_AnaParametreListesi.bas** | Ortak `GetActivePart()` + tüm parametreleri mesajda listeleyen `AnaParametreListesi`. | Açık Part. | MsgBox: parametre sayısı + ad = değer listesi. |
| **SadecePartKontrol.bas** | Aktif belgenin Part olup olmadığını kontrol eder; Part değilse uyarır. | 3DExperience açık. | Part açıksa: "Parça belgesi hazır: …"; değilse uyarı. |
| **IkiParametreTakas.bas** | Length.1 ve Length.2 değerlerini takas eder. | Açık Part, Length.1 ve Length.2 var. | MsgBox "Takas edildi."; modelde değerler yer değişir. |
| **LogOrnekMakro.bas** | Log dosyasına START/END (ve hata durumunda END ERR) yazar. | C:\Temp yazılabilir. | C:\Temp\macro_log.txt içinde tarih + mesaj satırları. |
| **MinMaxParametreDeger.bas** | Tüm parametreler arasında min ve max sayısal değeri bulur. | Açık Part, en az bir parametre. | MsgBox: "Min: … Max: …". |

---

## Notlar

- **API isimleri** (GetItem, Parameters, Shapes vb.) 3DExperience sürümüne göre değişebilir. Kendi ortamınızda makro kaydı yapıp üretilen kodu bu örneklerle karşılaştırın.
- **Yol:** `ParametreListesiniDosyayaYaz` içindeki `C:\Temp` yolunu kendi ortamınıza göre değiştirin.
- Tam rehber ve API referansı için proje kökündeki [README.md](../README.md), [VBA_API_REFERENCE.md](../docs/VBA_API_REFERENCE.md) ve [Guidelines/README.md](../Guidelines/README.md) dosyalarına bakın.

---

**Gezinme:** [Ana sayfa](../README.md) · [Rehber (Guidelines)](../Guidelines/README.md) · [Docs](../docs/README.md)
