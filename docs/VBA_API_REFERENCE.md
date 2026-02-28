# VBA API Referansı – 3DEX_TR

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  Sık kullanılan API imzaları  |  Help: resmi PDF'ler (Native Apps Automation, Automation Reference)  ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

Bu sayfa **3DExperience VBA** makro geliştirirken kullanılan API’lere yönlendirme ve kısa açıklamalar içerir. Tam liste ve imzalar proje içindeki aşağıdaki kaynaklarda bulunur.

---

## Hızlı erişim

| Kaynak | Açıklama |
|--------|----------|
| **Help klasörü (PDF)** | Resmi referans PDF'leri: *Help-Native Apps Automation.pdf*, *Help-Automation Reference.pdf*, *Help-Common Services.pdf* — API imzaları, nesne modeli, servis listesi. |
| [**Help/ARAMA_REHBERI.md**](../Help/ARAMA_REHBERI.md) | Hangi konuyu hangi PDF'te bulacağınız; PDF'de arama (Ctrl+F) rehberi. |

---

## Erişim zinciri (özet diyagram)

```
  GetObject(,"CATIA.Application")  ──►  oApp
           │
           ├── .ActiveDocument     ──►  oDoc   ──►  .GetItem("Part")     ──►  oPart  ──►  .Parameters, .Shapes, .Update
           │                                └──►  .GetItem("Product")   ──►  oProduct ──►  .Children
           │                                └──►  .GetItem("DrawingRoot") ──►  oDraw  ──►  .Sheets, .Views
           │
           ├── .ActiveEditor      ──►  oEditor  ──►  .GetService("...")   (Editor-level servis)
           └── .GetSessionService("...")         (Session-level servis)
```

---

## Sık kullanılan kalıplar (özet)

| İhtiyaç | Örnek (kavramsal) |
|--------|--------------------|
| Uygulama al | `Set oApp = GetObject(, "CATIA.Application")` |
| Aktif belge | `Set oDoc = oApp.ActiveDocument` |
| Part al | `Set oPart = oDoc.GetItem("Part")` veya `Set oPart = oDoc` |
| Parametre oku | `Set oParam = oPart.Parameters.Item("Length.1")` → `dVal = oParam.Value` |
| Parametre yaz | `oParam.Value = 50.5` → `oPart.Update` |
| Shapes döngüsü | `For i = 1 To oPart.Shapes.Count` … `Set oSh = oPart.Shapes.Item(i)` |
| Editor servis | `Set oEditor = oApp.ActiveEditor` → `Set oSvc = oEditor.GetService("InertiaService")` |

Tam imzalar ve sürüme özel API adları için Help klasöründeki **resmi PDF'leri** (Native Apps Automation, Automation Reference) kullanın; [Help/ARAMA_REHBERI.md](../Help/ARAMA_REHBERI.md) ile hangi konunun hangi PDF'te olduğunu ve PDF'de nasıl arayacağınızı bulabilirsiniz.

---

## Sık kullanılan API'ler (imza + açıklama)

Aşağıdaki tabloda günlük makro yazarken en çok ihtiyaç duyacağınız API'ler, VBA imzası ve kısa açıklama ile verilmiştir. İsimler sürüme göre değişebilir; şüphede Help klasöründeki **resmi PDF'lerle** (Native Apps Automation, Automation Reference) doğrulayın.

### Uygulama ve belge erişimi

| API | VBA imzası / kullanım | Açıklama |
|-----|------------------------|----------|
| **GetObject** | `Set oApp = GetObject(, "CATIA.Application")` | Çalışan 3DExperience oturumuna bağlanır. İlk parametre boş bırakılır. |
| **ActiveDocument** | `Set oDoc = oApp.ActiveDocument` | Aktif (ön plandaki) belgeyi döndürür. Açık belge yoksa Nothing. |
| **Documents** | `oApp.Documents` | Tüm açık belgelerin koleksiyonu. Count, Item(i) ile kullanılır. |
| **GetItem** | `Set oPart = oDoc.GetItem("Part")` | Belge içinde isimle nesne alır. "Part", "Product", "DrawingRoot" vb. |
| **Name** | `sAd = oDoc.Name` | Belge/nesne adı (dosya adı veya görünen ad). |
| **FullName** | `sYol = oDoc.FullName` | Belgenin tam dosya yolu (varsa). |

### Part – Parça

| API | VBA imzası / kullanım | Açıklama |
|-----|------------------------|----------|
| **Part (nesne)** | `Set oPart = oDoc` veya `oDoc.GetItem("Part")` | Parça belgesinin kök nesnesi. |
| **Parameters** | `Set oParams = oPart.Parameters` | Parametre koleksiyonu (uzunluk, açı vb.). |
| **Parameters.Item** | `Set oParam = oParams.Item("Length.1")` veya `Item(i)` | İsim veya indeks ile parametre alır. |
| **Parameter.Value** | `dVal = oParam.Value` / `oParam.Value = 50.5` | Parametre değerini okur veya yazar. |
| **Parameter.Name** | `sAd = oParam.Name` | Parametre adı. |
| **Update** | `oPart.Update` | Part üzerinde yapılan değişiklikleri uygular. Döngü dışında **bir kez** çağrılmalı. |
| **Shapes** | `Set oShapes = oPart.Shapes` (veya `oPart.MainBody.Shapes`) | Parçadaki şekiller (Pad, Pocket, Sketch vb.). |
| **Shapes.Count** | `iCount = oShapes.Count` | Şekil sayısı. |
| **Shapes.Item** | `Set oSh = oShapes.Item(i)` | İndeks ile şekil alır (1'den başlar). |
| **MainBody** | `Set oBody = oPart.MainBody` | Ana gövde nesnesi. |

### Product – Montaj

| API | VBA imzası / kullanım | Açıklama |
|-----|------------------------|----------|
| **Children** | `oProduct.Children` | Alt bileşenler (occurrence) koleksiyonu. |
| **Children.Item** | `Set oChild = oProduct.Children.Item("Parça1")` veya `Item(i)` | İsim veya indeks ile alt bileşen. |

### Drawing – Çizim

| API | VBA imzası / kullanım | Açıklama |
|-----|------------------------|----------|
| **Sheets** | `oDrawing.Sheets` | Çizim sayfaları koleksiyonu. |
| **ActiveSheet** | `Set oSheet = oDrawing.Sheets.ActiveSheet` | Aktif sayfa. |
| **Views** | `oSheet.Views` | Sayfadaki görünümler. |
| **Dimensions** | `oView.Dimensions` | Görünümdeki ölçüler (varsa). |

### Editor ve servisler

| API | VBA imzası / kullanım | Açıklama |
|-----|------------------------|----------|
| **ActiveEditor** | `Set oEditor = oApp.ActiveEditor` | Aktif editör penceresi. |
| **GetService** | `Set oSvc = oEditor.GetService("InertiaService")` | Editor-level servis alır (Part/Product/Drawing bağlamında). |
| **GetSessionService** | `Set oSvc = oApp.GetSessionService("SearchService")` | Oturum genelinde servis (PLM arama, açma vb.). Editor-level iş bitmeden çağırılmamalı. |

### Dosya ve ortam

| API | VBA imzası / kullanım | Açıklama |
|-----|------------------------|----------|
| **FileSystem** | `Set oFS = oApp.FileSystem` | Dosya/klasör işlemleri (CreateFolder, FileExists, OpenAsTextStream vb.). |

### Koleksiyon ortakları

| API | Kullanım | Açıklama |
|-----|----------|----------|
| **Count** | `i = oColl.Count` | Koleksiyon eleman sayısı (çoğunda property). |
| **Item** | `Set oElem = oColl.Item(i)` veya `Item("isim")` | İndeks (1 tabanlı) veya isimle eleman. |

---

## Rehberle ilişki

- **Nesne modeli ve erişim:** [Guidelines/06-3DExperience-Nesne-Modeli.md](../Guidelines/06-3DExperience-Nesne-Modeli.md), [Guidelines/13-Erisim-ve-Kullanim-Rehberi.md](../Guidelines/13-Erisim-ve-Kullanim-Rehberi.md)
- **Sık kullanılan API örnekleri:** [Guidelines/08-Sik-Kullanilan-APIler.md](../Guidelines/08-Sik-Kullanilan-APIler.md)
- **Help dosyalarını ne zaman/nasıl kullanacağınız:** [Guidelines/17-Help-Dosyalarini-Kullanma.md](../Guidelines/17-Help-Dosyalarini-Kullanma.md)

---

**Gezinme:** [Ana sayfa](../README.md) · [Docs](README.md) · [Rehber](../Guidelines/README.md) · [Örnek makrolar](../Examples/README.md) · [CHEATSHEET](CHEATSHEET.md) · [Arama rehberi](../Help/ARAMA_REHBERI.md)

*API isimleri 3DExperience sürümüne göre değişir; güncel referans için Help klasöründeki dosyaları kullanın.*
