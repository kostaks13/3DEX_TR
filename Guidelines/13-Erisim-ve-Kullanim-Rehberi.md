# 13. Neye Nereden Erişilir, Neyi Nasıl Kullanırsın

**Bu dokümanda:** Erişim tabloları (nereden ne alınır); kullanım tabloları; tek sayfa zincir özeti; "nereden alacağım?" cevapları.

Bu dokümanda **hangi nesneye/servise nereden ulaşacağınız** ve **o nesneyi/servisi nasıl kullanacağınız** tek yerde toplanmıştır. Her satırda: erişim yolu (VBA), ne işe yarar, nasıl kullanılır (okuma/yazma/çağrı) ve kısa kod örneği.

════════════════════════════════════════════════════════════════════════════════

## 1. Erişim: Nereden ne alınır?

Aşağıdaki tabloda **soldaki sütun** “buna ihtiyacım var”, **ortadaki** “bunu nereden alırım (VBA yolu)”, **sağdaki** “nasıl kullanırım (kısa)” şeklindedir.

### 1.1 Uygulama ve belge

| İhtiyaç | Nereden erişilir (VBA yolu) | Nasıl kullanılır |
|--------|-----------------------------|------------------|
| 3DExperience uygulaması | `GetObject(, "CATIA.Application")` veya `Application` | Tüm diğer nesnelere giriş noktası; önce `Set oApp = ...`, sonra `If oApp Is Nothing` kontrolü. |
| Aktif belge (açık olan) | `oApp.ActiveDocument` | Belge adı: `oDoc.Name`, tam yol: `oDoc.FullName`; yoksa `Nothing` döner. |
| Part (parça) nesnesi | `oDoc.GetItem("Part")` veya `Set oPart = oDoc` (parça belgesiyse) | Parametreler, Shapes, Bodies, HybridBodies vb. Part’a bağlı. |
| Product (montaj) nesnesi | `oDoc.GetItem("Product")` veya `Set oProduct = oDoc` (montaj belgesiyse) | Children, occurrence ağacı, BOM. |
| Drawing (çizim) nesnesi | `oDoc.GetItem("DrawingRoot")` veya `Set oDraw = oDoc` (çizim belgesiyse) | Sheets, Views, Dimensions. |
| Aktif editör | `oApp.ActiveEditor` | GetService (editor-level servis), ActiveObject, Selection. |
| Dosya/klasör işlemleri | `oApp.FileSystem` | GetFile, GetFolder, Exists; taşınabilir kod için bu kullanılır. |
| Ortam değişkeni, yol birleştirme | `oApp.SystemService` | Environ("DEĞİŞKEN"), ConcatenatePaths(yol1, yol2). |

**Kod kalıbı – Uygulama ve belge:**

```vba
Dim oApp As Object, oDoc As Object, oPart As Object
Set oApp = GetObject(, "CATIA.Application")
If oApp Is Nothing Then MsgBox "Uygulama yok.": Exit Sub
Set oDoc = oApp.ActiveDocument
If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
Set oPart = oDoc.GetItem("Part")
If oPart Is Nothing Then Set oPart = oDoc
' Bu noktada oPart kullanılabilir (parça belgesiyse)
```

════════════════════════════════════════════════════════════════════════════════

### 1.2 Part (parça) altındakiler

| İhtiyaç | Nereden erişilir (VBA yolu) | Nasıl kullanılır |
|--------|-----------------------------|------------------|
| Parametreler koleksiyonu | `oPart.Parameters` | Count, Item(i) veya Item("Length.1"); her eleman Parameter (Name, Value). |
| Tek parametre (isimle) | `oPart.Parameters.Item("Length.1")` | Okuma: `deger = oParam.Value`; yazma: `oParam.Value = 50.5`; sonra `oPart.Update` (bir kez). |
| Shapes (şekiller) | `oPart.Shapes` veya `oPart.MainBody.Shapes` | Count, Item(i); her eleman Shape (Name, tip). Pad, Pocket, Sketch vb. |
| Bodies (katı gövdeler) | `oPart.Bodies` | Count, Item(i); her eleman Body (Name). |
| HybridBodies (geometrik setler) | `oPart.HybridBodies` | Count, Item(i); her HybridBody içinde HybridShapes. |
| HybridShapes (yüzey, eğri, nokta) | `oHB = oPart.HybridBodies.Item(i)` sonra `oHB.HybridShapes` | Count, Item(j); nokta, eğri, yüzey eklemek için HybridShapeFactory kullanılır. |
| MainBody | `oPart.MainBody` | Ana katı gövde; Shapes bazen buradan alınır. |

**Kod kalıbı – Parametre okuma/yazma:**

```vba
Dim oParams As Object, oParam As Object
Set oParams = oPart.Parameters
If oParams Is Nothing Then Exit Sub
Set oParam = oParams.Item("Length.1")
If oParam Is Nothing Then Exit Sub
dDeger = oParam.Value                    ' Okuma
oParam.Value = 100                       ' Yazma
oPart.Update                             ' Tek Update (döngü dışında)
```

**Kod kalıbı – Shapes döngüsü:**

```vba
Dim oShapes As Object, oShape As Object, i As Long
Set oShapes = oPart.Shapes
If oShapes Is Nothing Then Exit Sub
For i = 1 To oShapes.Count
    Set oShape = oShapes.Item(i)
    ' oShape.Name, oShape tipi kullanılır
Next i
```

════════════════════════════════════════════════════════════════════════════════

### 1.3 Product (montaj) altındakiler

| İhtiyaç | Nereden erişilir (VBA yolu) | Nasıl kullanılır |
|--------|-----------------------------|------------------|
| Alt bileşenler (Children) | `oProduct.Children` | Count, Item(i) veya For Each; her eleman alt ürün/occurrence. |
| Kök occurrence (montaj ağacı) | `oEditor.GetService("PLMProductService")` → `.RootOccurrence` | Montaj kökü; alt occurrence’lara Occurrences ile geçilir. |
| Tek bileşen (isimle) | `oProduct.Children.Item("Parça1")` | Name, pozisyon vb. property’ler. |

**Kod kalıbı – BOM (Children listesi):**

```vba
Dim oProduct As Object, oChild As Object
Set oProduct = oDoc.GetItem("Product")
If oProduct Is Nothing Then Set oProduct = oDoc
If oProduct.Children Is Nothing Then Exit Sub
For Each oChild In oProduct.Children
    ' oChild.Name, listeye ekle
Next oChild
```

**Kod kalıbı – PLMProductService ile kök:**

```vba
Dim oEditor As Object, oProductSvc As Object, oRoot As Object
Set oEditor = oApp.ActiveEditor
Set oProductSvc = oEditor.GetService("PLMProductService")
If oProductSvc Is Nothing Then Exit Sub
Set oRoot = oProductSvc.RootOccurrence
' oRoot.Name, oRoot.Occurrences vb.
```

════════════════════════════════════════════════════════════════════════════════

### 1.4 Drawing (çizim) altındakiler

| İhtiyaç | Nereden erişilir (VBA yolu) | Nasıl kullanılır |
|--------|-----------------------------|------------------|
| Sayfalar | `oDraw.Sheets` | Count, Item(i); her eleman Sheet. |
| Aktif sayfa | `oDraw.ActiveSheet` veya `oDraw.Sheets.Item(1)` | Tek sayfa. |
| Bir sayfadaki görünümler | `oSheet.Views` | Count, Item(i); her eleman View. |
| Görünüm ölçeği | `oView.Scale` (veya benzeri property) | Okuma/yazma (sürüme göre). |
| Görünümdeki ölçüler | `oView.Dimensions` | Count, Item(i); metin, değer. |
| Görünümdeki metinler | `oView.DrawingTexts` | Çizim metinleri. |

**Kod kalıbı – Sayfa ve görünüm sayısı:**

```vba
Dim oDraw As Object, oSheets As Object, oSheet As Object, oViews As Object
Set oDraw = oDoc.GetItem("DrawingRoot")
If oDraw Is Nothing Then Set oDraw = oDoc
Set oSheets = oDraw.Sheets
For i = 1 To oSheets.Count
    Set oSheet = oSheets.Item(i)
    Set oViews = oSheet.Views
    ' oSheet.Name, oViews.Count
Next i
```

════════════════════════════════════════════════════════════════════════════════

### 1.5 Servisler – Nereden alınır?

| İhtiyaç | Nereden erişilir (VBA yolu) | Nasıl kullanılır |
|--------|-----------------------------|------------------|
| Session-level servis (arama, yeni belge, kamera vb.) | `oApp.GetSessionService("Tanımlayıcı")` | Tanımlayıcı: "VisuServices", "Search", "PLMNewService", "PLMOpenService" vb. Dönen nesne servise özel metodlara sahiptir. |
| Editor-level servis (kütle, ölçüm, çizim, montaj kökü) | `oApp.ActiveEditor.GetService("Tanımlayıcı")` | Tanımlayıcı: "InertiaService", "MeasureService", "CATDrawingService", "PLMProductService" vb. Aktif pencere ilgili türde (Part/Product/Drawing) olmalı. |

**Kod kalıbı – Session servis:**

```vba
Dim oSvc As Object
Set oSvc = oApp.GetSessionService("VisuServices")
If oSvc Is Nothing Then Exit Sub
' oSvc.Cameras, oSvc.Windows vb.
```

**Kod kalıbı – Editor servis:**

```vba
Dim oEditor As Object, oSvc As Object
Set oEditor = oApp.ActiveEditor
If oEditor Is Nothing Then Exit Sub
Set oSvc = oEditor.GetService("InertiaService")
If oSvc Is Nothing Then Exit Sub
' oSvc ile kütle/atalet işlemleri (API'ye bakın)
```

════════════════════════════════════════════════════════════════════════════════

### 1.6 FileSystem – Dosya ve klasör

| İhtiyaç | Nereden erişilir (VBA yolu) | Nasıl kullanılır |
|--------|-----------------------------|------------------|
| FileSystem nesnesi | `oApp.FileSystem` | GetFile, GetFolder, Exists, Create, Copy, Delete (Help’te imzalar). |
| Var olan dosya | `oApp.FileSystem.GetFile("C:\Tam\yol\dosya.txt")` | .Path, .Size, .OpenAsTextStream (metin okuma/yazma). |
| Var olan klasör | `oApp.FileSystem.GetFolder("C:\Tam\yol")` | .Files (dosya listesi), .SubFolders (alt klasörler). |
| Dosya/klasör var mı? | `oApp.FileSystem.Exists("C:\yol")` | Boolean döner. |

**Kod kalıbı – Dosya var mı ve boyutu:**

```vba
Dim oFS As Object, oFile As Object
Set oFS = oApp.FileSystem
If oFS.Exists("C:\Temp\log.txt") Then
    Set oFile = oFS.GetFile("C:\Temp\log.txt")
    lBoyut = oFile.Size
End If
```

════════════════════════════════════════════════════════════════════════════════

## 2. Kullanım: Neyi nasıl yaparsın?

Aşağıda **yaygın işlemler** tek tek “neyi kullanırsın” ve “nasıl yaparsın” olarak özetleniyor.

### 2.1 Parametre: Okuma, yazma, listeleme

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Tek parametrenin değerini okumak | `oPart.Parameters.Item("Length.1")` → `.Value` | `Set oP = oPart.Parameters.Item("Length.1")` → `dVal = oP.Value` |
| Tek parametreye değer yazmak | Aynı + `.Value = deger` ve `oPart.Update` | `oP.Value = 100` → `oPart.Update` (döngü içinde değil) |
| Tüm parametreleri listelemek | `oPart.Parameters` → Count, Item(i) | `For i = 1 To oParams.Count` → `Set oP = oParams.Item(i)` → Name, Value |
| İsme göre parametre aramak | `oParams.Item(sParamAdi)` | On Error ile Nothing kontrolü; yoksa Item bulunamaz. |

════════════════════════════════════════════════════════════════════════════════

### 2.2 Shapes / Bodies: Listeleme, sayım

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Şekil sayısını almak | `oPart.Shapes.Count` | `iCount = oPart.Shapes.Count` |
| Her şeklin adını almak | `oPart.Shapes.Item(i).Name` | For i = 1 To Count döngüsü |
| Body listesi | `oPart.Bodies` | Count, Item(i).Name |
| HybridBody ve içindeki geometriler | `oPart.HybridBodies.Item(i).HybridShapes` | İç içe döngü; her HybridShape.Name |

════════════════════════════════════════════════════════════════════════════════

### 2.3 Montaj: BOM, kök, occurrence

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Alt bileşen listesi (BOM) | `oProduct.Children` | For Each oChild In oProduct.Children → oChild.Name |
| Montaj kökü (occurrence ağacı) | `PLMProductService.RootOccurrence` | GetService("PLMProductService") → .RootOccurrence |
| Belirli bileşeni isimle almak | `oProduct.Children.Item("Parça1")` | Nothing kontrolü |

════════════════════════════════════════════════════════════════════════════════

### 2.4 Çizim: Sayfa, görünüm, ölçü

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Sayfa listesi | `oDraw.Sheets` | Count, Item(i).Name |
| Bir sayfadaki görünüm sayısı | `oSheet.Views.Count` | Set oViews = oSheet.Views → oViews.Count |
| Görünüm ölçeğini okumak | `oView.Scale` (veya benzeri) | Property’yi oku; sürüme göre ad değişir. |
| Ölçü listesi | `oView.Dimensions` | Count, Item(i); değer/metin property’leri |

════════════════════════════════════════════════════════════════════════════════

### 2.5 Servisler: Ne zaman hangisi?

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Kütle / atalet (Part veya Product) | **InertiaService** (Editor) | GetService("InertiaService") → Inertia nesnesi; kütle property’si (API’ye bakın) |
| Mesafe, alan, açı ölçümü | **MeasureService** (Editor) | GetService("MeasureService") → seçim/referans ile ölçüm |
| Çizim sayfa/görünüm işlemleri | **CATDrawingService** (Editor) | GetService("CATDrawingService"); çizim penceresi aktif olmalı |
| Montaj kökü (occurrence) | **PLMProductService** (Editor) | GetService("PLMProductService") → RootOccurrence |
| PLM’de arama | **SearchService** (Session) | GetSessionService("Search") → arama sorgusu ve Results |
| Yeni Part/Product/Drawing oluşturma | **PLMNewService** (Session) | GetSessionService("PLMNewService") → oluşturma metodu |
| Belge açma (PLM’den) | **PLMOpenService** (Session) | GetSessionService("PLMOpenService") → açma metodu |
| Kamera / katman listesi | **VisuServices** (Session) | GetSessionService("VisuServices") → .Cameras vb. |

════════════════════════════════════════════════════════════════════════════════

### 2.6 Dosya: Var mı, listele, yaz

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Dosya var mı? | `oApp.FileSystem.Exists(yol)` | Boolean; True/False |
| Dosya boyutu | `oApp.FileSystem.GetFile(yol).Size` | Byte cinsinden |
| Klasördeki dosyaları listeleme | `oApp.FileSystem.GetFolder(yol).Files` | For Each oFile In oFolder.Files |
| Metin dosyasına satır yazma | `File.OpenAsTextStream` (FileSystem) veya VBA `Open ... For Append` | FileSystem taşınabilir; Open/Print/Close klasik VBA. |

════════════════════════════════════════════════════════════════════════════════

## 3. Tek sayfa özet: Erişim zinciri

```
GetObject(, "CATIA.Application")     → oApp
oApp.ActiveDocument                  → oDoc (Nothing kontrolü)
oDoc.GetItem("Part")                 → oPart
oDoc.GetItem("Product")              → oProduct
oDoc.GetItem("DrawingRoot")          → oDraw

oPart.Parameters                     → oParams  → .Item(i) veya .Item("Length.1")
oPart.Shapes                         → oShapes → .Item(i)
oPart.Bodies                         → oBodies → .Item(i)
oPart.HybridBodies.Item(i)           → oHB     → .HybridShapes

oProduct.Children                    → For Each oChild

oDraw.Sheets.Item(i)                 → oSheet  → .Views
oSheet.Views.Item(j)                 → oView   → .Dimensions, .Scale

oApp.ActiveEditor                    → oEditor
oEditor.GetService("InertiaService") → oSvc (Editor-level)
oApp.GetSessionService("VisuServices") → oSvc (Session-level)

oApp.FileSystem                     → oFS → GetFile, GetFolder, Exists
oApp.SystemService                   → Environ, ConcatenatePaths
```

════════════════════════════════════════════════════════════════════════════════

## 4. Sık hata: “Nereden alacağım?” cevabı

| Soru | Cevap (nereden) |
|------|------------------|
| Parametre listesi nerede? | `oPart.Parameters` (Part’ı önce oDoc.GetItem("Part") ile al) |
| Şekiller nerede? | `oPart.Shapes` veya `oPart.MainBody.Shapes` |
| Montajdaki alt parçalar nerede? | `oProduct.Children` (Product’ı oDoc.GetItem("Product") ile al) |
| Çizimdeki sayfalar nerede? | `oDraw.Sheets` (Drawing’i oDoc.GetItem("DrawingRoot") ile al) |
| Kütle hesabı nerede? | Editor-level **InertiaService**: `oApp.ActiveEditor.GetService("InertiaService")` |
| PLM araması nerede? | Session-level **Search**: `oApp.GetSessionService("Search")` |
| Dosya var mı nerede? | `oApp.FileSystem.Exists(yol)` |
| Yeni belge oluşturma nerede? | Session-level **PLMNewService**: `oApp.GetSessionService("PLMNewService")` |
| Excel’e erişim nerede? | `CreateObject("Excel.Application")` veya `GetObject(, "Excel.Application")`; kitap: Workbooks.Open / Add; hücre: Worksheet.Range("A1") veya Cells(satir,sutun). **14. doküman:** [14-VBA-ve-Excel-Etkilesimi.md](14-VBA-ve-Excel-Etkilesimi.md). |
| Dosya seçtirme / kaydetme diyaloğu nerede? | FileDialog (Excel üzerinden: aç=1, kaydet=2, klasör=4) veya Windows API GetOpenFileName/GetSaveFileName. **15. doküman:** [15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md). |

════════════════════════════════════════════════════════════════════════════════

## 5. İlgili dokümanlar

**Tüm rehber:** [README](README.md). Konuya göre: [06](06-3DExperience-Nesne-Modeli.md) (hiyerarşi), [08](08-Sik-Kullanilan-APIler.md) (API örnekleri), [12](12-Servisler-ve-Yapilabilecek-Islemler.md) (servisler), [14](14-VBA-ve-Excel-Etkilesimi.md) (Excel), [15](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md) (dosya diyalogları), [18](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) (sık hatalar ve dikkat noktaları). İmzalar: **VBA_API_REFERENCE.md**. **Help dosyalarını ne zaman/nasıl kullanacağınız** için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md).

**Gezinme:** Önceki: [12-Servisler](12-Servisler-ve-Yapilabilecek-Islemler.md) | [Rehber listesi](README.md) | Sonraki: [14-Excel](14-VBA-ve-Excel-Etkilesimi.md) →
