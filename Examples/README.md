# Örnek Makrolar

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  Çalıştırılabilir .bas örnekleri  |  Rehber 08 + 10 ile uyumlu               ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

Bu klasörde **3DExperience VBA** rehberindeki (Guidelines 08 ve 10) örnek makroların **çalıştırılabilir** kopyaları bulunur. Her dosya `.bas` formatındadır; 3DExperience VBA editörüne modül olarak aktarabilir veya içeriği kopyalayıp kendi modülünüze yapıştırabilirsiniz.

> **İlk kez kullanıyorsanız:** Aşağıdaki **Önerilen sıra** tablosundan 1 → 2 → 3 → 4 ile başlayın.

**Bu sayfada:** [Kullanım (3 adım)](#kullanım-3-adım) · [Önerilen sıra](#önerilen-sıra-yeni-başlıyorsanız) · [Kategoriye göre](#kategoriye-göre-örnekler) · [Dosyalar](#dosyalar) · [Notlar](#notlar)

**Öne çıkan:** [AktifParcaBilgisi.bas](AktifParcaBilgisi.bas) (ilk makro) · [ParametreOkuVeGoster.bas](ParametreOkuVeGoster.bas) · [ParametreYaz.bas](ParametreYaz.bas) · [SadecePartKontrol.bas](SadecePartKontrol.bas) · [ParametreleriExceleYaz.bas](ParametreleriExceleYaz.bas) · [ExceldenPartaParametreYaz.bas](ExceldenPartaParametreYaz.bas) (Excel↔Part) · [FileDialogParametreListesiYaz.bas](FileDialogParametreListesiYaz.bas) · [FileSystemDosyaBilgisi.bas](FileSystemDosyaBilgisi.bas) (dosya boyutu)

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
| — | Diğerleri | Shapes, dosyaya yazma, log, modüler yapı (GetActivePart), takas, min/max, **Excel** (ParametreleriExceleYaz), **FileDialog** (FileDialogParametreListesiYaz). |

---

## Kategoriye göre örnekler

| Kategori | Örnekler | Açıklama |
|----------|----------|----------|
| **Zincir / Kontrol** | AktifParcaBilgisi, SadecePartKontrol | Application → Document → Part; belge türü kontrolü. |
| **Parametre** | ParametreOkuVeGoster, ParametreYaz, IkiParametreTakas, MinMaxParametreDeger, GetActivePart_AnaParametreListesi | Okuma, yazma, takas, min/max, listeleme. |
| **Shapes** | ShapesBilgisi | Shapes sayısı ve isim listesi. |
| **Drawing** | DrawingSayfaVeGorusumSayisi | Drawing belgesi; Sheets, Views sayısı. |
| **Product / BOM** | BomListesiChildren | Montajda Children ile alt bileşen listesi. |
| **FileSystem** | FileSystemKlasorKontrol, FileSystemDosyaBilgisi, FileSystemKlasorListele | Klasör var mı; dosya var mı / boyut; klasördeki dosya listesi. |
| **Dosya** | ParametreListesiniDosyayaYaz, FileDialogParametreListesiYaz, FileDialogDosyaAcOku, KlasorSecParametreListesiYaz | Parametreleri dosyaya yazma; FileDialog ile kayıt/klasör seçimi; dosya açıp okuma. |
| **Excel** | ParametreleriExceleYaz, ExceldenPartaParametreYaz | Part → Excel; Excel → Part (Rehber 14). |
| **Log** | LogOrnekMakro | Log dosyasına START/END yazma. |
| **Servisler (Rehber 12)** | InertiaServiceKutleGoster, VisuServicesKameraListesi, HybridBodiesListele | InertiaService; VisuServices Cameras; Part HybridBodies/HybridShapes. |
| **Modüler** | GetActivePart_AnaParametreListesi | Ortak GetActivePart() + ana parametre listesi. |

**Ortam notu:** Örnekler 3DExperience R2024x (veya uyumlu sürüm), Windows üzerinde test edilmiştir. API isimleri sürüme göre değişebilir; kendi ortamınızda makro kaydı ile doğrulayın.

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
| **DrawingSayfaVeGorusumSayisi.bas** | Aktif Drawing’de sayfa sayısı ve ilk sayfadaki görünüm sayısını gösterir. | Açık Drawing (.CATDrawing). | MsgBox: sayfa sayısı, ilk sayfa Views sayısı. |
| **BomListesiChildren.bas** | Aktif Product’taki alt bileşenleri (Children) listeler. | Açık Product (.CATProduct). | MsgBox: bileşen sayısı + ilk 20 ad. |
| **FileSystemKlasorKontrol.bas** | Kullanıcının girdiği klasör yolunun var olup olmadığını kontrol eder. | 3DExperience açık; FileSystem API mevcut. | MsgBox: "Klasör mevcut" / "bulunamadı". |
| **ParametreleriExceleYaz.bas** | Aktif parçanın parametre listesini Excel'e yazar (A: Parametre, B: Değer); C:\Temp\ParametreListe.xlsx kaydeder. | Açık Part; Excel yüklü. | Excel açılır, liste yazılır, dosya kaydedilir; MsgBox "… yazıldı: …". |
| **FileDialogParametreListesiYaz.bas** | FileDialog (Kaydet) ile kullanıcının seçtiği yere parametre listesini (Parametre;Değer) yazar. | Açık Part; Excel yüklü (FileDialog için). | Kaydet diyaloğu açılır; seçilen .txt/.csv dosyasına liste yazılır veya "İptal edildi.". |
| **ExceldenPartaParametreYaz.bas** | Excel dosyasında (A: parametre adı, B: değer) satırları okuyup Part parametrelerini günceller; Update çağırır. | Açık Part; C:\Temp\ParametreGiris.xlsx mevcut (veya yolu değiştirin). | MsgBox "… Part'a yazıldı"; modelde değerler güncellenir. |
| **FileSystemDosyaBilgisi.bas** | Kullanıcının girdiği dosya yolunun var olup olmadığını ve boyutunu (byte) gösterir. | 3DExperience açık; FileSystem API. | MsgBox: "Dosya: … Boyut: N byte" veya "Dosya bulunamadı". |
| **FileDialogDosyaAcOku.bas** | FileDialog (Aç) ile kullanıcı .txt/.xlsx seçer; makro dosyayı açar ve ilk satırları MsgBox ile gösterir. | Açık Part; FileDialog API. | Aç diyaloğu; seçilen dosyadan ilk satırlar mesajda. |
| **KlasorSecParametreListesiYaz.bas** | FileDialog (Klasör seçici, msoFileDialogFolderPicker) ile klasör seçtirir; parametre listesini seçilen klasöre yazar. | Açık Part; FileDialog API. | Klasör seçici; seçilen klasöre parametre listesi .txt. |
| **FileSystemKlasorListele.bas** | Kullanıcının girdiği klasördeki dosyaları listeler (ad, boyut). | 3DExperience açık; FileSystem API. | MsgBox: dosya adı + boyut listesi. |
| **InertiaServiceKutleGoster.bas** | Aktif editörde InertiaService alır; kütle bilgisi için Help referansı verir (API sürüme göre değişir). | Part/Product penceresi aktif. | MsgBox: "InertiaService alındı" + API notu. |
| **HybridBodiesListele.bas** | Part.HybridBodies ve her body içindeki HybridShapes isimlerini listeler. | Açık Part. | MsgBox: body adları + shape adları. |
| **VisuServicesKameraListesi.bas** | GetSessionService("VisuServices"), Cameras koleksiyonu; kamera adlarını listeler. | 3DExperience açık. | MsgBox: kamera sayısı + ad listesi. |

---

## Eklenebilecek örnekler (fikir listesi)

Rehberde anlatılıp henüz ayrı bir `.bas` örneği olmayan konular. İhtiyaca göre eklenebilir.

| Konu | Açıklama | Rehber |
|------|----------|--------|
| ~~Excel'den Part'a parametre yazma~~ | ✅ [ExceldenPartaParametreYaz.bas](ExceldenPartaParametreYaz.bas) eklendi. | 14 |
| ~~FileDialog ile dosya açma~~ | ✅ [FileDialogDosyaAcOku.bas](FileDialogDosyaAcOku.bas) eklendi. | 15 |
| ~~FileDialog ile klasör seçme~~ | ✅ [KlasorSecParametreListesiYaz.bas](KlasorSecParametreListesiYaz.bas) eklendi. | 15 |
| ~~FileSystem: dosya var mı / boyut~~ | ✅ [FileSystemDosyaBilgisi.bas](FileSystemDosyaBilgisi.bas) eklendi. | 12 |
| ~~FileSystem: klasördeki dosya listesi~~ | ✅ [FileSystemKlasorListele.bas](FileSystemKlasorListele.bas) eklendi. | 12 |
| ~~InertiaService – kütle~~ | ✅ [InertiaServiceKutleGoster.bas](InertiaServiceKutleGoster.bas) eklendi. | 12 |
| ~~HybridBodies / Bodies listesi~~ | ✅ [HybridBodiesListele.bas](HybridBodiesListele.bas) eklendi. | 12 |
| ~~VisuServices – kamera listesi~~ | ✅ [VisuServicesKameraListesi.bas](VisuServicesKameraListesi.bas) eklendi. | 12 |

*İşaretli (✅) satırlar projeye eklenmiş örneklerdir.*

---

## Notlar

- **API isimleri** (GetItem, Parameters, Shapes vb.) 3DExperience sürümüne göre değişebilir. Kendi ortamınızda makro kaydı yapıp üretilen kodu bu örneklerle karşılaştırın.
- **Yol:** `ParametreListesiniDosyayaYaz` ve `ParametreleriExceleYaz` içindeki `C:\Temp` (veya varsayılan dosya adı) yolunu kendi ortamınıza göre değiştirin. `FileDialogParametreListesiYaz` kayıt yerini kullanıcı seçer.
- Tam rehber ve API referansı için proje kökündeki [README.md](../README.md), [VBA_API_REFERENCE.md](../docs/VBA_API_REFERENCE.md) ve [Guidelines/README.md](../Guidelines/README.md) dosyalarına bakın.

---

**Gezinme:** [Ana sayfa](../README.md) · [Rehber (Guidelines)](../Guidelines/README.md) · [Docs](../docs/README.md)
