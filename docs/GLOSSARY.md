# Terimler Sözlüğü – 3DExperience VBA

```
┌──────────────────────────────────────────────────────────────────────────────┐
│  Alfabetik terim listesi  |  Part, Parameter, Nothing, Update, GetItem…      │
└──────────────────────────────────────────────────────────────────────────────┘
```

Rehberde ve API referansında geçen temel terimlerin **kısa tanımları**. Detay için ilgili Guidelines dokümanına ve [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) dosyasına bakın.

**TR ↔ EN:** Aşağıdaki tablolarda **EN** sütunu İngilizce karşılığı gösterir (Help ve API’de geçen isimler).

**Bölümler:** [A–D](#a-d) · [E–L](#e-l) · [M–P](#m-p) · [S–U](#s-u) · [V–Z](#v-z)

---

## A–D

| Terim (TR/API) | EN | Açıklama |
|----------------|-----|----------|
| **ActiveDocument** | ActiveDocument | O anda ön planda olan (aktif) belge. `oApp.ActiveDocument` ile alınır; açık belge yoksa **Nothing** döner. |
| **ActiveEditor** | ActiveEditor | O anda ön plandaki editör (Part, Product veya Drawing ile ilişkili). Editor-level servislere `ActiveEditor.GetService("...")` ile erişilir. |
| **Application** | Application | 3DExperience uygulama nesnesi. Genelde `GetObject(, "CATIA.Application")` ile alınır. Tüm oturuma erişimin giriş noktasıdır. |
| **ByVal / ByRef** | ByVal / ByRef | Parametre geçirme: **ByVal** kopya gönderir (değişiklik çağıranı etkilemez); **ByRef** (varsayılan) referans gönderir, değişiklik çağıranı etkiler. |
| **Collection** | Collection | Nesnelerin listesi (Parameters, Shapes, Children vb.). `.Count`, `.Item(index)` veya `.Item("isim")` ile kullanılır. |
| **DrawingRoot** | DrawingRoot | Çizim (Drawing) belgesinin kök nesnesi. Sayfalar (Sheets), görünümler (Views) buradan erişilir. |

---

## E–L

| Terim (TR/API) | EN | Açıklama |
|----------------|-----|----------|
| **Children** | Children | Montaj (Product) içindeki alt bileşenler koleksiyonu. `oProduct.Children` ile erişilir; BOM veya ağaç dolaşımında kullanılır. |
| **Editor-level servis** | Editor-level service | Aktif editördeki Part/Product/Drawing’e bağlı servis. `ActiveEditor.GetService("ServisAdi")` ile alınır (örn. InertiaService, MeasureService). İçeriden session-level servis çağrılmamalıdır. |
| **Exit Sub / Exit Function** | Exit Sub / Exit Function | Prosedürü veya fonksiyonu o noktada sonlandırır. Hata durumunda veya erken çıkışta kullanılır. |
| **GetItem** | GetItem | Belge veya koleksiyondan isimle nesne alır. Örn. `oDoc.GetItem("Part")` → Part nesnesi. Sürüme göre API adı değişebilir. |
| **GetObject** | GetObject | COM üzerinden çalışan uygulamaya bağlanır. `GetObject(, "CATIA.Application")` 3DExperience oturumunu döndürür; yoksa hata veya Nothing. |
| **Item** | Item | Koleksiyondan tek öğe alır. `oPart.Parameters.Item("Length.1")` veya `oPart.Parameters.Item(1)` (indeks ile). |

---

## M–P

| Terim (TR/API) | EN | Açıklama |
|----------------|-----|----------|
| **MainBody** | MainBody | Parça (Part) belgesindeki ana gövde. Şekiller çoğunlukla `oPart.Shapes` veya `oPart.MainBody.Shapes` ile alınır. |
| **Nothing** | Nothing | VBA’da “nesne atanmamış” anlamına gelen özel değer. `If oDoc Is Nothing Then` ile kontrol edilir; kontrol edilmezse sonraki satırlarda hata oluşur. |
| **On Error** | On Error | Hata yakalama: `On Error GoTo Etiket` (etiketli satıra git), `On Error Resume Next` (sonraki satıra geç), `On Error GoTo 0` (varsayılana dön). |
| **Option Explicit** | Option Explicit | Modül başında yazılır; tüm değişkenlerin `Dim` ile tanımlanmasını zorunlu kılar. Yazım hatalarını azaltır; önerilir. |
| **Parameter** | Parameter | Part’ta tanımlı sayısal veya metin değer (uzunluk, açı vb.). `oPart.Parameters.Item("Length.1")` ile alınır; `.Value` ile okunur/yazılır. |
| **Part** | Part | Parça belgesi nesnesi. Parametreler, şekiller (Shapes), MainBody bu nesneden erişilir. |
| **Product** | Product | Montaj belgesi nesnesi. Alt bileşenler `oProduct.Children` ile alınır. |

---

## S–U

| Terim (TR/API) | EN | Açıklama |
|----------------|-----|----------|
| **Session-level servis** | Session-level service | Oturum genelinde, açık belgeden bağımsız işlemler (PLM arama, belge açma vb.). `oApp.GetSessionService("ServisAdi")` ile alınır. Editor-level iş bittikten sonra kullanılmalıdır. |
| **Sheets** | Sheets | Çizim (Drawing) belgesindeki sayfalar koleksiyonu. `oDrawingRoot.Sheets`; her sayfa ölçek (Scale), görünümler (Views) içerir. |
| **Shapes** | Part’taki şekiller koleksiyonu (Pad, Pocket, Sketch vb.). `oPart.Shapes` veya `oPart.MainBody.Shapes`; `.Count`, `.Item(i)` kullanılır. |
| **Update** | Update | Part veya Product üzerinde yapılan değişiklikleri uygulayan metod (`oPart.Update`). Performans için döngü içinde değil, tüm değişiklikler bittikten sonra **bir kez** çağrılmalıdır. |
| **Views** | Views | Çizim sayfasındaki görünümler koleksiyonu. Her görünüm ölçek, konum ve ilişkili geometri bilgisi taşır. |

---

## V–Z

| Terim (TR/API) | EN | Açıklama |
|----------------|-----|----------|
| **VBA** | VBA | Visual Basic for Applications. 3DExperience makrolarında kullanılan dil. |
| **V5 API** | V5 API | Eski CATIA V5 API’si (örn. Documents.Add, HybridShapeFactoryOld). 3DExperience’ta desteklenmeyebilir; yeni API’ye geçilmelidir. |
| **Value** | Value | Parametre veya benzeri nesnelerin değerini okumak/yazmak için kullanılan property. Örn. `oParam.Value = 100`. |

---

*Tam liste için [Guidelines](../Guidelines/README.md) ve [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) kullanın.*

**Gezinme:** [Ana sayfa](../README.md) · [Docs](README.md) · [Rehber](../Guidelines/README.md) · [Örnek makrolar](../Examples/README.md) · [VBA_API_REFERENCE](VBA_API_REFERENCE.md)
