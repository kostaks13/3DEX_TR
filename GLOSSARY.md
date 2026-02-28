# Terimler Sözlüğü – 3DExperience VBA

Rehberde ve API referansında geçen temel terimlerin kısa tanımları. Detay için ilgili Guidelines dokümanına ve [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) dosyasına bakın.

---

## A–D

| Terim | Açıklama |
|-------|----------|
| **ActiveDocument** | O anda ön planda olan (aktif) belge. `oApp.ActiveDocument` ile alınır; açık belge yoksa **Nothing** döner. |
| **Application** | 3DExperience uygulama nesnesi. Genelde `GetObject(, "CATIA.Application")` ile alınır. Tüm oturuma erişimin giriş noktasıdır. |
| **Editor-level servis** | Aktif editördeki Part/Product/Drawing’e bağlı servis. `ActiveEditor.GetService("ServisAdi")` ile alınır (örn. InertiaService, MeasureService). İçeriden session-level servis çağrılmamalıdır. |
| **GetItem** | Belge veya koleksiyondan isimle nesne alır. Örn. `oDoc.GetItem("Part")` → Part nesnesi. Sürüme göre API adı değişebilir. |
| **DrawingRoot** | Çizim (Drawing) belgesinin kök nesnesi. Sayfalar (Sheets), görünümler (Views) buradan erişilir. |

---

## M–P

| Terim | Açıklama |
|-------|----------|
| **MainBody** | Parça (Part) belgesindeki ana gövde. Şekiller çoğunlukla `oPart.Shapes` veya `oPart.MainBody.Shapes` ile alınır. |
| **Nothing** | VBA’da “nesne atanmamış” anlamına gelen özel değer. `If oDoc Is Nothing Then` ile kontrol edilir; kontrol edilmezse sonraki satırlarda hata oluşur. |
| **Parameter** | Part’ta tanımlı sayısal veya metin değer (uzunluk, açı vb.). `oPart.Parameters.Item("Length.1")` ile alınır; `.Value` ile okunur/yazılır. |
| **Part** | Parça belgesi nesnesi. Parametreler, şekiller (Shapes), MainBody bu nesneden erişilir. |
| **Product** | Montaj belgesi nesnesi. Alt bileşenler `oProduct.Children` ile alınır. |

---

## S–U

| Terim | Açıklama |
|-------|----------|
| **Session-level servis** | Oturum genelinde, açık belgeden bağımsız işlemler (PLM arama, belge açma vb.). `oApp.GetSessionService("ServisAdi")` ile alınır. Editor-level iş bittikten sonra kullanılmalıdır. |
| **Shapes** | Part’taki şekiller koleksiyonu (Pad, Pocket, Sketch vb.). `oPart.Shapes` veya `oPart.MainBody.Shapes`; `.Count`, `.Item(i)` kullanılır. |
| **Update** | Part veya Product üzerinde yapılan değişiklikleri uygulayan metod (`oPart.Update`). Performans için döngü içinde değil, tüm değişiklikler bittikten sonra **bir kez** çağrılmalıdır. |

---

## V–Z

| Terim | Açıklama |
|-------|----------|
| **VBA** | Visual Basic for Applications. 3DExperience makrolarında kullanılan dil. |
| **V5 API** | Eski CATIA V5 API’si (örn. Documents.Add, HybridShapeFactoryOld). 3DExperience’ta desteklenmeyebilir; yeni API’ye geçilmelidir. |

---

*Tam liste için [Guidelines](Guidelines/README.md) ve [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) kullanın.*
