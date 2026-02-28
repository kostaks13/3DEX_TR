# 6. 3DExperience Nesne Modeli

```
  Konu: Application -> ActiveDocument -> Part/Product/DrawingRoot, GetObject, FileSystem
```

3DExperience’ta her şey **nesne** (object) hiyerarşisi içindedir. Makro yazarken önce **uygulama**ya, sonra **açık belge**ye, oradan **parça**, **montaj** veya **çizim** nesnelerine ulaşırsınız. Bu dokümanda bu yapıyı öğreneceksiniz.

════════════════════════════════════════════════════════════════════════════════

## Hiyerarşi özeti

```
Application (3DExperience uygulaması)
  └── ActiveDocument (açık olan belge: parça, montaj veya çizim)
        └── Part / Product / DrawingRoot (belge türüne göre)
              └── MainBody / Shapes / Children / Sheets ... (alt nesneler)
```

**Özet – tek satır akış:**

```
  GetObject  ──►  oApp  ──►  .ActiveDocument  ──►  oDoc  ──►  .GetItem("Part"|"Product"|"DrawingRoot")  ──►  oPart / oProduct / oDraw
```

- **Application:** Tüm oturuma erişim. Belgeleri listeleme, aktif belgeyi alma, yeni belge açma vb.  
- **ActiveDocument:** O anda odakta olan belge (Part, Product veya Drawing).  
- **Part:** Parça belgesi — MainBody, Shapes, Parameters vb.  
- **Product:** Montaj belgesi — Children, Constraints vb.  
- **DrawingRoot:** Çizim belgesi — Sheets, Views vb.

**“Neye nereden erişilir, neyi nasıl kullanırım?”** sorusunun tablo ve kod kalıplarıyla yanıtı için **13. doküman:** [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md). **Help dosyalarını hangi aşamada nasıl kullanacağınız** için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md).

════════════════════════════════════════════════════════════════════════════════

## Uygulama nesnesini alma

3DExperience zaten açıkken VBA’dan uygulama nesnesini almak için genelde şu yöntemler kullanılır:

**Yöntem 1 – GetObject (çalışan oturuma bağlan):**

```vba
Dim oApp As Object
On Error Resume Next
Set oApp = GetObject(, "CATIA.Application")
On Error GoTo 0

If oApp Is Nothing Then
    MsgBox "3DExperience (CATIA) çalışmıyor. Önce uygulamayı açın."
    Exit Sub
End If
```

**Yöntem 2 – Bazı ortamlarda Application doğrudan:**

```vba
Dim oApp As Object
Set oApp = Application   ' 3DExperience VBA içinden bazen bu kullanılır
```

Hangi yöntemin geçerli olduğu 3DExperience sürümüne ve rolüne göre değişir; dokümantasyon veya kayıt edilmiş makroyu kontrol edin.

════════════════════════════════════════════════════════════════════════════════

## Aktif belge ve türü

```vba
Dim oApp As Object
Set oApp = GetObject(, "CATIA.Application")

Dim oDoc As Object
Set oDoc = oApp.ActiveDocument

If oDoc Is Nothing Then
    MsgBox "Açık belge yok."
    Exit Sub
End If

' Belge adı ve türü
MsgBox "Belge: " & oDoc.Name & ", FullName: " & oDoc.FullName
```

Belge türüne göre **Part**, **Product** veya **Drawing** nesnesine geçersiniz:

```vba
Dim oPart As Object
' Sadece parça belgesi açıksa:
Set oPart = oDoc   ' veya oDoc.GetItem("Part") vb. (sürüme göre API farklı olabilir)
' oPart.MainBody, oPart.Shapes, oPart.Parameters vb. kullanılır
```

Proje içindeki `VBA_API_REFERENCE.md` ve Help dosyalarında ilgili sınıfların tam API’si listelenir; sürümünüze göre `Part`, `Product`, `DrawingRoot` nasıl alınır oradan bakılmalıdır.

════════════════════════════════════════════════════════════════════════════════

## Parça (Part) içinde dolaşma

Parça belgesinde genelde şunlar vardır:

- **MainBody:** Ana gövde.  
- **Shapes:** Gövde içindeki şekiller (Pad, Pocket, Sketch vb.) — bir koleksiyondur.  
- **Parameters:** Parametreler (Length, Angle vb.).

Kavramsal örnek (gerçek property/method isimleri sürüme göre değişir):

```vba
Dim oPart As Object
Set oPart = oDoc

Dim oShapes As Object
Set oShapes = oPart.Shapes   ' veya Part.MainBody.Shapes

Dim i As Long
For i = 1 To oShapes.Count
    Dim oShape As Object
    Set oShape = oShapes.Item(i)
    MsgBox "Shape " & i & ": " & oShape.Name
Next i
```

`VBA_CALL_LIST.txt` / `VBA_API_REFERENCE.md` içinde `Part`, `Shapes`, `Item` gibi sınıf ve metodları arayarak doğru API’yi bulun.

════════════════════════════════════════════════════════════════════════════════

## Montaj (Product) ve çocuklar

Montajda **Children** (veya benzeri) koleksiyonu alt bileşenleri verir. Yine sürüme göre isim değişir; Help veya referans dokümanına bakın.

```vba
Dim oProduct As Object
Set oProduct = oDoc   ' Montaj belgesi açıksa

Dim oChildren As Object
Set oChildren = oProduct.Children   ' veya Products, Components vb.

Dim oChild As Object
For Each oChild In oChildren
    MsgBox "Bileşen: " & oChild.Name
Next oChild
```

════════════════════════════════════════════════════════════════════════════════

## Çizim (Drawing)

Çizim belgesinde **Sheets** (sayfalar), her sayfada **Views** (görünümler) vb. vardır. Erişim örneği (API sürüme göre uyarlanmalı):

```vba
Dim oDrawing As Object
Set oDrawing = oDoc   ' Çizim belgesi açıksa

Dim oSheets As Object
Set oSheets = oDrawing.Sheets
' oSheets.Item(1).Views vb.
```

════════════════════════════════════════════════════════════════════════════════

## Önemli noktalar

1. **Her zaman Nothing kontrolü:** `If oDoc Is Nothing Then Exit Sub` gibi.  
2. **Belge türü kontrolü:** Parça mı, montaj mı, çizim mi buna göre farklı nesnelere geçin.  
3. **Referans dokümanı:** Bu projedeki `VBA_API_REFERENCE.md` ve `Help/text/` altındaki metinler, hangi sınıfta hangi property/method olduğunu gösterir; sürümünüzle eşleştirin.

════════════════════════════════════════════════════════════════════════════════

## Application nesnesinin yapısı (Help-Native Apps Automation)

**Help-Native Apps Automation** (Foundation Object Model Map) dokümanına göre **Application** nesnesi tüm oturumun köküdür ve şunları toplar:

- **Editors** — Uygulamanın yönettiği tüm editörler. `CATIA.Editors` ile alınır.
- **ActiveEditor** — O anda etkin olan editör; aktif penceredeki nesneye ve komutlara erişim sağlar. `CATIA.ActiveEditor`.
- **ActiveEditor.ActiveObject** — O anda düzenlenen verinin kökü: **Part**, **DrawingRoot** veya **VPMReference** (Product) olabilir.
- **Windows** — Uygulama tarafından açılan pencereler.
- **FileSystem** — Klasör ve dosya işlemleri; platformdan bağımsız kullanım için **CreateObject("Scripting.FileSystemObject")** yerine `CATIA.FileSystem` kullanılır (Help’e göre taşınabilirlik için önerilir).
- **GetSessionService(ad)** — Oturum seviyesi servisler: PLMNewService, PLMOpenService, SearchService, VisuServices, MaterialService, SystemService vb.
- **SystemService** — Ortam değişkenleri (`Environ`), yol birleştirme (`ConcatenatePaths`) gibi sistem bilgisi ve yardımcıları.
- **SettingControllers** — Ayarlar (Options) deposuna erişim.

Editör seviyesi servisler için **ActiveEditor** üzerinden **GetService(ad)** kullanılır: örn. CATDrawingService, InertiaService, MeasureService, PLMProductService.

════════════════════════════════════════════════════════════════════════════════

## 3DExperience otomasyon hiyerarşi ağacı (özet)

**Help** klasöründeki **3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.pdf** metninden özet:

```
CATIA (Application)
├── Editors
│   ├── Item(i): Editor
│   ├── ActiveEditor
│   │   ├── ActiveWindow → ActiveViewer, Viewpoint2D/3D, LightSources
│   │   ├── ActiveObject
│   │   │   ├── Part
│   │   │   │   ├── HybridBodies → HybridShapes, AppendHybridShape
│   │   │   │   ├── Bodies (OrderedGeometricalSets)
│   │   │   │   ├── Parameters (ParameterSet)
│   │   │   │   ├── Relations
│   │   │   │   ├── ShapeFactory (AddNewPad, AddNewPocket, …)
│   │   │   │   ├── HybridShapeFactory (AddNewPlaneOffset, AddNewPointCoord, …)
│   │   │   │   ├── InWorkObject, Update
│   │   │   │   └── Parent → VPMRepReference
│   │   │   ├── Product (VPMReference) → RepInstances, Occurrences, AssemblyConstraints
│   │   │   └── DrawingRoot → Sheets, Views, Dimensions, Tables, GDT
│   │   ├── Selection
│   │   ├── GetService("InertiaService"), GetService("MeasureService"), …
│   │   └── GetService("PLMProductService") → RootOccurrence
│   └── …
├── Windows, Printers, SettingControllers
├── GetSessionService(name) → SearchService, MaterialService, PLMNewService, …
├── FileSystem → GetFolder, GetFile, Exists, Copy, Delete, Create
├── SystemService → Environ, ConcatenatePaths
└── …
```

Bu ağaç, “nesneye nereden ulaşırım?” sorusunda yol haritasıdır; tam API listesi için `VBA_API_REFERENCE.md` ve Help PDF’lerine bakın.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Tam makro – Uygulama → Belge → Part iskeleti

Aşağıdaki makro, Application alır, ActiveDocument alır, Nothing kontrolleri yapar; ardından Part nesnesine geçmeye çalışır (sürüme göre `GetItem("Part")` veya doğrudan `oDoc` kullanılabilir). Part alındıysa Shapes sayısını mesajla gösterir.

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Uygulama → Belge → Part zincirini gösterir.

Sub UygulamaBelgePartZinciri()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oShapes As Object
    Dim iCount As Long

    On Error Resume Next
    Set oApp = GetObject(, "CATIA.Application")
    On Error GoTo 0
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil."
        Exit Sub
    End If

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok."
        Exit Sub
    End If

    ' Part nesnesi (sürüme göre: oDoc veya oDoc.GetItem("Part"))
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc

    On Error Resume Next
    Set oShapes = oPart.Shapes
    On Error GoTo 0
    If Not oShapes Is Nothing Then
        iCount = oShapes.Count
        MsgBox "Belge: " & oDoc.Name & vbCrLf & "Shapes sayısı: " & iCount
    Else
        MsgBox "Belge: " & oDoc.Name & vbCrLf & "Shapes alınamadı (farklı belge türü olabilir)."
    End If
End Sub
```

Bu iskeleti kopyalayıp kendi iş mantığınızı (parametre okuma, geometri ekleme vb.) ekleyebilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Örnek: ActiveEditor ve ActiveObject (CATIA referansı)

Help’e göre bazen **CATIA.ActiveEditor** ve **CATIA.ActiveEditor.ActiveObject** kullanılır. ActiveObject, o anda düzenlenen kök nesnedir (Part, DrawingRoot veya VPMReference/Product).

```vba
Sub EditorVeActiveObjectOrnek()
    Dim oEditor As Object
    Dim oActiveObj As Object

    On Error Resume Next
    Set oEditor = GetObject(, "CATIA.Application").ActiveEditor
    On Error GoTo 0
    If oEditor Is Nothing Then Exit Sub

    Set oActiveObj = oEditor.ActiveObject
    If oActiveObj Is Nothing Then
        MsgBox "Aktif nesne yok."
        Exit Sub
    End If
    ' oActiveObj Part, Product veya DrawingRoot olabilir
    MsgBox "Aktif nesne var (Part/Product/Drawing olabilir)."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: FileSystem ile dosya kontrolü (taşınabilir kod)

Help’e göre platformlar arası script için **CATIA.FileSystem** kullanılmalı; Windows’a özel `Scripting.FileSystemObject` yerine. Dosya var mı kontrolü:

```vba
Sub DosyaVarMiOrnek()
    Dim oApp As Object
    Dim oFS As Object
    Dim sPath As String
    Dim bVar As Boolean

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then Exit Sub

    sPath = "C:\Temp\test.txt"
    On Error Resume Next
    bVar = oFS.Exists(sPath)
    On Error GoTo 0
    MsgBox "Dosya var mı: " & bVar
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: GetSessionService – SearchService fikri

Session-level servisler **GetSessionService** ile alınır. Örnek: arama servisi (gerçek identifier Help’teki tabloda).

```vba
Sub SessionServiceOrnek()
    Dim oApp As Object
    Dim oSearch As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub

    On Error Resume Next
    Set oSearch = oApp.GetSessionService("SearchService")
    On Error GoTo 0
    If oSearch Is Nothing Then
        MsgBox "SearchService alınamadı (sürüm/rol farkı olabilir)."
        Exit Sub
    End If
    ' oSearch.Search(...) ile PLM araması yapılabilir
    MsgBox "SearchService alındı."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: HybridBodies ve HybridShapes (GSD)

Parça içinde **HybridBodies** (Geometrical Sets) ve her birinin **HybridShapes** koleksiyonu vardır. Nokta, eğri, yüzey gibi geometriler burada tutulur. Kavramsal erişim:

```vba
Sub HybridBodiesOrnek()
    Dim oPart As Object
    Dim oHBs As Object
    Dim oHB As Object
    Dim oHSs As Object
    Dim i As Long
    Dim j As Long
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oHBs = oPart.HybridBodies
    If oHBs Is Nothing Then Exit Sub
    For i = 1 To oHBs.Count
        Set oHB = oHBs.Item(i)
        If Not oHB Is Nothing Then
            Set oHSs = oHB.HybridShapes
            If Not oHSs Is Nothing Then
                For j = 1 To oHSs.Count
                    Debug.Print "Body " & i & " Shape " & j & ": " & oHSs.Item(j).Name
                Next j
            End If
        End If
    Next i
End Sub
```

API isimleri (HybridBodies, HybridShapes) sürüme göre değişebilir; kayıt veya referanstan kontrol edin.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Drawing – Sheets ve ActiveSheet

Çizim belgesinde **Sheets** koleksiyonu sayfaları içerir; **ActiveSheet** o anki sayfayı verir:

```vba
Sub CizimSayfalariOrnek()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oDraw As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim i As Long
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    Set oDraw = oDoc.GetItem("DrawingRoot")
    If oDraw Is Nothing Then Set oDraw = oDoc
    Set oSheets = oDraw.Sheets
    If oSheets Is Nothing Then Exit Sub
    For i = 1 To oSheets.Count
        Set oSheet = oSheets.Item(i)
        If Not oSheet Is Nothing Then MsgBox "Sayfa " & i & ": " & oSheet.Name
    Next i
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Product – Root ve Occurrences

Montajda **Root** kök ürün; **Children** veya **Products** alt bileşenler. Bazen **Occurrences** veya **RepInstances** ile occurrence listesi alınır (API’ye göre):

```vba
Sub ProductRootOrnek()
    Dim oDoc As Object
    Dim oProduct As Object
    Dim oRoot As Object
    Set oDoc = GetObject(, "CATIA.Application").ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    Set oProduct = oDoc.GetItem("Product")
    If oProduct Is Nothing Then Set oProduct = oDoc
    Set oRoot = oProduct   ' veya oProduct.Parent / GetRoot vb. (sürüme göre)
    MsgBox "Kök ürün: " & oRoot.Name
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Bodies (OrderedGeometricalSets) – Part Design

Part’ta **Bodies** (veya OrderedGeometricalSets) katı geometri (Pad, Pocket vb.) içerir. Shapes ile benzer şekilde Count ve Item(i) ile dolaşılır (API isimleri sürüme göre değişir):

```vba
Sub BodiesOrnek()
    Dim oPart As Object
    Dim oBodies As Object
    Dim i As Long
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oBodies = oPart.Bodies
    If oBodies Is Nothing Then Exit Sub
    For i = 1 To oBodies.Count
        Debug.Print "Body " & i & ": " & oBodies.Item(i).Name
    Next i
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Application.Documents – Açık belgeler listesi

Bazen aktif belge yerine **tüm açık belgeleri** taramak gerekir. **Application.Documents** (veya **Documents** koleksiyonu) açık belgeleri verir; Count ve Item(i) ile döngü yapılır. API adı sürüme göre değişir; “Documents”, “OpenDocuments” vb. Help’te aranmalıdır.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Editor.GetService – Editor-level servis

Aktif editördeki **Part** veya **Product**’a bağlı servisler **ActiveEditor.GetService("ServisAdi")** ile alınır. Örnek servis isimleri: InertiaService, MeasureService, DrawingService. Bu servisler o anda düzenlenen kök nesneye (Part/Product/Drawing) özeldir; oturum genelinde değil.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Selection – Aktif seçim (kavramsal)

Kullanıcının ekranda seçtiği nesnelere erişmek için **ActiveEditor.Selection** veya **Document.Selection** kullanılır (sürüme göre). Count ve Item(i) ile seçili elemanlar taranır; Filter veya Search ile belirli tipte nesne aranabilir. Tam API için Help-Native Apps Automation ve VBA_API_REFERENCE.md içinde “Selection” araması yapın.

════════════════════════════════════════════════════════════════════════════════

## Kontrol listesi

- [ ] Application nesnesini (GetObject veya Application) alabiliyorum  
- [ ] ActiveDocument ile açık belgeye erişebiliyorum  
- [ ] Part / Product / Drawing ayrımını biliyorum  
- [ ] Koleksiyonda For veya For Each ile dolaşacağımı biliyorum  

════════════════════════════════════════════════════════════════════════════════

## Sonraki adım

**7. doküman:** [07-Makro-Kayit-ve-Inceleme.md](07-Makro-Kayit-ve-Inceleme.md) — Makro kaydetme ve oluşan kodu inceleyip düzenleme.

**Gezinme:** Önceki: [05-Prosedurler](05-VBA-Temelleri-Prosedurler-ve-Fonksiyonlar.md) | [Rehber listesi](README.md) | Sonraki: [07-Makro-Kayit](07-Makro-Kayit-ve-Inceleme.md) →
