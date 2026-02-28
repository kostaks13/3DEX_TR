# Örnek Makrolar

Bu klasörde **3DExperience VBA** rehberindeki (Guidelines 08 ve 10) örnek makroların çalıştırılabilir kopyaları bulunur. Her dosya `.bas` formatındadır; 3DExperience VBA editörüne modül olarak aktarabilir veya içeriği kopyalayıp kendi modülünüze yapıştırabilirsiniz.

---

## Kullanım (3 satır)

1. **3DExperience** açın; gerekirse bir **Part** belgesi açın.
2. **Tools → Macro → Edit** (veya eşdeğeri) ile VBA editörünü açın; yeni modül ekleyip ilgili `.bas` dosyasının içeriğini yapıştırın.
3. Makroyu çalıştırın (F5 veya Run); mesajları veya çıktı dosyasını kontrol edin.

---

## Dosyalar

| Dosya | Ne yapar | Gereksinim |
|-------|----------|------------|
| **AktifParcaBilgisi.bas** | Aktif belgenin adı, tam yolu ve (Part ise) Shapes sayısını MsgBox ile gösterir. | Açık Part belgesi. |
| **ParametreOkuVeGoster.bas** | Kullanıcıdan parametre adı (örn. Length.1) alır; değerini mesajla gösterir. | Açık Part, parametre adı. |
| **ParametreYaz.bas** | Kullanıcıdan parametre adı ve yeni değer alır; Part'ta günceller ve Update çağırır. | Açık Part, yazılabilir parametre. |
| **ShapesBilgisi.bas** | Shapes sayısı ve ilk 10 şeklin adını listeler. | Açık Part. |
| **ParametreListesiniDosyayaYaz.bas** | Tüm parametreleri `Parametre;Değer` formatında `C:\Temp\parametre_listesi.txt` dosyasına yazar. | Açık Part; C:\Temp yazılabilir olmalı. |
| **GetActivePart_AnaParametreListesi.bas** | Ortak `GetActivePart()` fonksiyonu + tüm parametreleri mesajda listeleyen `AnaParametreListesi` makrosu (modüler yapı örneği). | Açık Part. |

---

## Notlar

- **API isimleri** (GetItem, Parameters, Shapes vb.) 3DExperience sürümüne göre değişebilir. Kendi ortamınızda makro kaydı yapıp üretilen kodu bu örneklerle karşılaştırın.
- **Yol:** `ParametreListesiniDosyayaYaz` içindeki `C:\Temp` yolunu kendi ortamınıza göre değiştirin.
- Tam rehber ve API referansı için proje kökündeki [README.md](../README.md), [VBA_API_REFERENCE.md](../VBA_API_REFERENCE.md) ve [Guidelines/README.md](../Guidelines/README.md) dosyalarına bakın.
