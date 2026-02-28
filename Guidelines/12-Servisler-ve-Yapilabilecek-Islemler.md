# 12. Servisler ve Yapılabilecek İşlemler (Detay)

Bu dokümanda **servisler** (Editor-level ve Session-level) ile **yapılabilecek işlemler** daha detaylı anlatılır; her biri için kullanım örnekleri ve kod verilir. Kaynak: **Help-Native Apps Automation** (Foundation Object Model Map) ve **Help-Common Services**. **Help dosyalarını hangi aşamada nasıl kullanacağınız** için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md).

------------------------------------------------------------

## 1. Servis nedir?

**Service** (servis), nesneden bağımsız işlemler sunan bir soyut nesnedir. İki tür vardır:

- **Session-level servis:** Oturum genelinde, **aktif editör veya açık belgeden bağımsız** işlemler (PLM arama, yeni belge açma, malzeme atama, kamera/katman yönetimi vb.). Erişim: **Application.GetSessionService("ServisTanimi")**.
- **Editor-level servis:** Aktif editördeki **belirli PLM kök nesnesine** (Part, Product, Drawing) bağlı işlemler. Örneğin kütle hesaplama (InertiaService) aktif parça/montaj için geçerlidir. Erişim: **ActiveEditor.GetService("ServisTanimi")**.

**Önemli kural:** Editor-level bir servis kullanırken **içeriden** session-level servis çağrısı yapmayın; kilitleşme riski vardır. Önce editor-level işinizi bitirip sonra session-level’a geçin.

------------------------------------------------------------

## 2. Session-level servisler (GetSessionService)

Aşağıdaki tablo, Help’teki “Service Identifier” listesine göre **GetSessionService** ile alınan servislerdir. Parametre olarak verilen string tam olarak bu tanımlayıcı olmalıdır (sürüm/rol bazı servisleri kapatabilir).

| Servis tanımlayıcı | Açıklama (kısa) |
|--------------------|------------------|
| **MATPLMService** | Malzeme oluşturma, çekirdek/kaplama malzeme atama veya okuma, atanmış malzemeleri kaldırma, oturumdaki tüm malzemeleri listeleme. |
| **PLMNewService** | Yeni PLM kök nesnesi (Part, Product, Drawing vb.) oluşturma. |
| **PLMOpenService** | PLM nesnesini (belge) açma. |
| **PLMPropagateService** | Editörde tutulan değişiklikleri kaydetme (save/propagate). |
| **PLMRefreshService** | Yenileme (refresh) işlemini yönetme. |
| **PLMScriptService** | Veritabanında saklanan script’leri yönetme. |
| **ProductSessionService** | Shape3D nesnelerinin koleksiyonunu toplar. |
| **SearchService** (tanımlayıcı: **"Search"**) | PLM nesnelerinde arama yapma. |
| **SimInitializationService** | Simülasyon başlatma. |
| **SIMPLMService** | SimulationReference ve diğer simülasyon nesnelerini yönetme. |
| **VisuServices** | Katmanlar, katman filtreleri, pencereler ve kameralarla çalışma. |
| **OlpTranslatorHelper** | OLP veri nesnelerine erişim. |
| **PnOService** | Kişi (person) nesnesine erişim. |

### Örnek: VisuServices – Kamera koleksiyonu

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Session-level VisuServices ile kamera listesine erişim.

Sub VisuServicesKameraListesi()
    Dim oApp As Object
    Dim oVisu As Object
    Dim oCameras As Object
    Dim oCam As Object
    Dim i As Long
    Dim sOut As String

    On Error GoTo HataYakala
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oVisu = oApp.GetSessionService("VisuServices")
    If oVisu Is Nothing Then
        MsgBox "VisuServices alınamadı."
        Exit Sub
    End If
    Set oCameras = oVisu.Cameras
    If oCameras Is Nothing Then Exit Sub

    sOut = "Kameralar: " & vbCrLf
    For i = 1 To oCameras.Count
        Set oCam = oCameras.Item(i)
        If Not oCam Is Nothing Then sOut = sOut & i & ": " & oCam.Name & vbCrLf
    Next i
    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

### Örnek: SearchService – PLM araması (kavramsal)

```vba
Sub SearchServiceOrnek()
    Dim oApp As Object
    Dim oSearch As Object
    Dim oResults As Object
    Dim i As Long

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    On Error Resume Next
    Set oSearch = oApp.GetSessionService("Search")
    On Error GoTo 0
    If oSearch Is Nothing Then
        MsgBox "SearchService (Search) alınamadı."
        Exit Sub
    End If
    ' oSearch ile arama sorgusu oluşturulur, çalıştırılır; sonuçlar Results ile alınır.
    ' Tam API: DatabaseSearch, Query, Execute vb. Help ve VBA_API_REFERENCE.md'ye bakın.
    MsgBox "SearchService alındı; arama API dokümantasyonuna göre kullanılır."
End Sub
```

### Örnek: PLMOpenService – Belge açma (kavramsal)

```vba
Sub PLMOpenServiceOrnek()
    Dim oApp As Object
    Dim oOpen As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    On Error Resume Next
    Set oOpen = oApp.GetSessionService("PLMOpenService")
    On Error GoTo 0
    If oOpen Is Nothing Then
        MsgBox "PLMOpenService alınamadı."
        Exit Sub
    End If
    ' oOpen ile PLM kaynağından belge açılır (metod imzası Help'te).
    MsgBox "PLMOpenService alındı."
End Sub
```

### Örnek: PLMNewService – Yeni kök nesne oluşturma (kavramsal)

```vba
Sub PLMNewServiceOrnek()
    Dim oApp As Object
    Dim oNew As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    On Error Resume Next
    Set oNew = oApp.GetSessionService("PLMNewService")
    On Error GoTo 0
    If oNew Is Nothing Then
        MsgBox "PLMNewService alınamadı."
        Exit Sub
    End If
    ' oNew ile yeni Part/Product/Drawing oluşturulur (V5 Documents.Add yerine).
    MsgBox "PLMNewService alındı."
End Sub
```

------------------------------------------------------------

## 3. Editor-level servisler (GetService)

Aşağıdaki tablo, **ActiveEditor.GetService("...")** ile alınan servislerdir. Hepsi **aktif editördeki kök nesneye** (Part, Drawing veya Product/VPMReference) bağlıdır.

| Servis tanımlayıcı | Açıklama (kısa) |
|--------------------|------------------|
| **CATDrawingGenService** | Çizim üretimi (DrawingGen). |
| **CATDrawingService** | Çizim temsiline (Drawing) uygulanan işlemler. |
| **InertiaService** | Verilen geometrik nesne için **Inertia** (kütle, atalet) nesnesini döndürür. |
| **InertiaBoxService** | Verilen geometrik nesnenin sınırlayıcı kutusu (**InertiaBox**) nesnesini döndürür. |
| **MeasurableService** | Eğri ve yüzeyleri **ölçmek** için nesneler döndürür. |
| **MeasureService** | Eğri ve yüzeyleri ölçmek için nesneler döndürür. |
| **PLMProductService** | Editörün **kök occurrence**’ını (RootOccurrence) döndürür; montaj ağacında gezinme. |
| **KnowledgeServices** | **Units** koleksiyonu veya **KnowledgeCollection** döndürür. |
| **InterferenceServices** | Girişim (interference) simülasyonu oluşturur. |
| **FCBService** | Esnek panolar (flexible boards) oluşturur veya getirir. |
| **FittingService** | Tracks ve Tpoints oluşturur veya yönetir. |
| **RfgService** | Referans düzlemleri ve yüzeyleri yönetir. |
| **SectionService** | Mevcut incelemedeki Section nesnelerini yönetir. |
| **SimExecutionService** | Simülasyon çalıştırma. |
| **SimSimulationService** | Simülasyonun referans verdiği ürünün kök occurrence’ını döndürür. |
| **StrService** | Structure nesnelerine uygulanır. |
| **ValidationService** | Doğrulama (validation) nesnelerine uygulanır. |

### Örnek: InertiaService – Kütle ve atalet

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Aktif editördeki Part/Product için kütle bilgisi (InertiaService).

Sub InertiaServiceKutle()
    Dim oApp As Object
    Dim oEditor As Object
    Dim oInertiaSvc As Object
    Dim oInertia As Object
    Dim dMass As Double

    On Error GoTo HataYakala
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then
        MsgBox "Aktif editör yok."
        Exit Sub
    End If
    Set oInertiaSvc = oEditor.GetService("InertiaService")
    If oInertiaSvc Is Nothing Then
        MsgBox "InertiaService alınamadı (Part/Product penceresi aktif olmalı)."
        Exit Sub
    End If
    ' Servis üzerinden Inertia nesnesi alınır; genelde kök geometri veya seçim verilir.
    ' Örnek: Set oInertia = oInertiaSvc.GetInertia(oActiveObject) vb. (API'ye bakın)
    ' dMass = oInertia.Mass veya benzeri property
    dMass = 0#
    MsgBox "InertiaService alındı. Kütle API dokümantasyonuna göre okunur (örn. " & dMass & ")."
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

### Örnek: MeasureService – Ölçüm

```vba
Sub MeasureServiceOrnek()
    Dim oApp As Object
    Dim oEditor As Object
    Dim oMeasure As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then Exit Sub
    On Error Resume Next
    Set oMeasure = oEditor.GetService("MeasureService")
    On Error GoTo 0
    If oMeasure Is Nothing Then
        MsgBox "MeasureService alınamadı."
        Exit Sub
    End If
    ' oMeasure ile mesafe, açı, yüzey alanı vb. hesaplanır (seçim veya referans gerekir).
    MsgBox "MeasureService alındı."
End Sub
```

### Örnek: PLMProductService – Kök occurrence (montaj ağacı)

```vba
Sub PLMProductServiceRootOccurrence()
    Dim oApp As Object
    Dim oEditor As Object
    Dim oProductSvc As Object
    Dim oRoot As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then Exit Sub
    On Error Resume Next
    Set oProductSvc = oEditor.GetService("PLMProductService")
    On Error GoTo 0
    If oProductSvc Is Nothing Then
        MsgBox "PLMProductService alınamadı (montaj penceresi aktif olmalı)."
        Exit Sub
    End If
    Set oRoot = oProductSvc.RootOccurrence
    If Not oRoot Is Nothing Then
        MsgBox "Kök occurrence: " & oRoot.Name
    Else
        MsgBox "RootOccurrence alınamadı."
    End If
End Sub
```

### Örnek: CATDrawingService – Çizim servisi

```vba
Sub DrawingServiceOrnek()
    Dim oApp As Object
    Dim oEditor As Object
    Dim oDrawSvc As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then Exit Sub
    On Error Resume Next
    Set oDrawSvc = oEditor.GetService("CATDrawingService")
    On Error GoTo 0
    If oDrawSvc Is Nothing Then
        MsgBox "CATDrawingService alınamadı (çizim penceresi aktif olmalı)."
        Exit Sub
    End If
    MsgBox "CATDrawingService alındı."
End Sub
```

------------------------------------------------------------

## 4. FileSystem (Application.FileSystem) – Dosya ve klasör

**FileSystem** bir servis değil; **Application** nesnesinin bir **property**’sidir. Platformdan bağımsız dosya/klasör işlemleri için kullanılır (Help: Windows Scripting Host FileSystemObject yerine bu kullanılmalı).

| İşlem | Açıklama |
|-------|-----------|
| **GetFile(yol)** | Var olan bir dosyayı **File** nesnesi olarak döndürür. |
| **GetFolder(yol)** | Var olan bir klasörü **Folder** nesnesi olarak döndürür. |
| **Exists(yol)** | Dosya veya klasörün var olup olmadığını döndürür. |
| **Create**, **Copy**, **Delete** | Dosya/klasör oluşturma, kopyalama, silme (metod imzaları Help’te). |

**File** nesnesi: **Path**, **Size**, **OpenAsTextStream** (metin okuma/yazma için **TextStream**).  
**Folder** nesnesi: **Path**, **Files** (dosyalar koleksiyonu), **SubFolders** (alt klasörler).

### Örnek: FileSystem – Dosya var mı, boyutu

```vba
Sub FileSystemDosyaBilgi()
    Dim oApp As Object
    Dim oFS As Object
    Dim oFile As Object
    Dim sPath As String
    Dim bVar As Boolean
    Dim lSize As Long

    sPath = "C:\Temp\macro_log.txt"
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then Exit Sub

    On Error Resume Next
    bVar = oFS.Exists(sPath)
    On Error GoTo 0
    If Not bVar Then
        MsgBox "Dosya yok: " & sPath
        Exit Sub
    End If

    Set oFile = oFS.GetFile(sPath)
    If oFile Is Nothing Then Exit Sub
    lSize = oFile.Size
    MsgBox "Dosya: " & sPath & vbCrLf & "Boyut: " & lSize & " byte"
End Sub
```

### Örnek: FileSystem – Klasördeki dosyaları listele

```vba
Sub FileSystemKlasorListele()
    Dim oApp As Object
    Dim oFS As Object
    Dim oFolder As Object
    Dim oFiles As Object
    Dim oFile As Object
    Dim sPath As String
    Dim sOut As String

    sPath = "C:\Temp"
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then Exit Sub
    On Error Resume Next
    Set oFolder = oFS.GetFolder(sPath)
    On Error GoTo 0
    If oFolder Is Nothing Then
        MsgBox "Klasör bulunamadı: " & sPath
        Exit Sub
    End If
    Set oFiles = oFolder.Files
    sOut = "Dosyalar (" & sPath & "):" & vbCrLf
    For Each oFile In oFiles
        sOut = sOut & oFile.Path & " (" & oFile.Size & " B)" & vbCrLf
    Next oFile
    MsgBox sOut
End Sub
```

------------------------------------------------------------

## 5. Yapılabilecek işlemler – Parça (Part)

Aşağıdaki işlemler **Part** kök nesnesi ve alt nesneleriyle yapılır. Hepsi için önce **Application → ActiveDocument → GetItem("Part")** (veya oDoc) ile Part alınır.

| İşlem | Kullanılan nesne/koleksiyon | Açıklama |
|-------|-----------------------------|----------|
| Parametre okuma | **Part.Parameters.Item(isim)** veya **Item(i)** | Name, Value property’leri. |
| Parametre yazma | **Parameter.Value = değer**, ardından **Part.Update** | Tek Update kuralına uyun. |
| Parametre listesi | **Parameters.Count**, 1’den Count’a döngü, **Item(i).Name**, **Item(i).Value** | Rapor veya dosyaya yazma. |
| Shapes listesi | **Part.Shapes** (veya MainBody.Shapes), **Count**, **Item(i).Name** | Pad, Pocket, Sketch vb. |
| Bodies listesi | **Part.Bodies** (OrderedGeometricalSets), **Count**, **Item(i)** | Katı gövde listesi. |
| HybridBodies / HybridShapes | **Part.HybridBodies**, her **HybridBody.HybridShapes** | Yüzey, eğri, nokta vb. |
| Geometri ekleme (Part Design) | **ShapeFactory** (Pad, Pocket vb.) | AddNewPad, AddNewPocket vb. (referansta arayın). |
| Geometri ekleme (GSD) | **HybridShapeFactory** | AddNewPointCoord, AddNewPlaneOffset vb. |
| Inertia (kütle) | **InertiaService** (Editor-level) | GetService("InertiaService") → Inertia nesnesi. |
| Ölçüm | **MeasureService** / **MeasurableService** (Editor-level) | Seçim veya referans ile mesafe, alan vb. |

### Örnek: Part – HybridBodies ve HybridShapes döngüsü

```vba
Sub PartHybridBodiesDetay()
    Dim oApp As Object
    Dim oPart As Object
    Dim oHBs As Object
    Dim oHB As Object
    Dim oHSs As Object
    Dim oHS As Object
    Dim i As Long
    Dim j As Long
    Dim sOut As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oHBs = oPart.HybridBodies
    If oHBs Is Nothing Then Exit Sub

    sOut = "HybridBodies ve HybridShapes:" & vbCrLf
    For i = 1 To oHBs.Count
        Set oHB = oHBs.Item(i)
        If Not oHB Is Nothing Then
            sOut = sOut & "BODY " & i & ": " & oHB.Name & vbCrLf
            Set oHSs = oHB.HybridShapes
            If Not oHSs Is Nothing Then
                For j = 1 To oHSs.Count
                    Set oHS = oHSs.Item(j)
                    If Not oHS Is Nothing Then sOut = sOut & "  - " & oHS.Name & vbCrLf
                Next j
            End If
        End If
    Next i
    MsgBox sOut
End Sub
```

### Örnek: Part – Tüm parametreleri CSV’ye yaz

```vba
Sub PartParametreleriCSV()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim iFile As Integer
    Dim sDosya As String
    Dim sSatir As String

    sDosya = "C:\Temp\parametreler.csv"
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub

    iFile = FreeFile
    Open sDosya For Output As #iFile
    Print #iFile, "ParametreAdi;Deger"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            sSatir = oParam.Name & ";" & Replace(CStr(oParam.Value), ";", ",")
            Print #iFile, sSatir
        End If
    Next i
    Close #iFile
    MsgBox "Yazıldı: " & sDosya
End Sub
```

------------------------------------------------------------

## 6. Yapılabilecek işlemler – Montaj (Product)

| İşlem | Kullanılan nesne/koleksiyon | Açıklama |
|-------|-----------------------------|----------|
| Alt bileşen listesi | **Product.Children** (veya **Products**), **For Each** / **Item(i)** | BOM, gezinme. |
| Kök occurrence | **PLMProductService.RootOccurrence** | Montaj ağacı kökü. |
| Occurrence gezinme | **VPMRootOccurrence** / **VPMOccurrence**, **Occurrences** koleksiyonu | Alt occurrence’lar. |
| PLMEntity | **VPMOccurrence.PLMEntity** | GetAttributeValue / SetAttributeValue ile PLM öznitelikleri. |
| Montaj kısıtları | **AssemblyConstraints** vb. (API’ye göre) | Constraint listesi. |
| Kütle (montaj) | **InertiaService** (editör montaj penceresindeyken) | Toplam kütle. |

### Örnek: Product – Children ile BOM listesi

```vba
Sub ProductBOMListe()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oProduct As Object
    Dim oChild As Object
    Dim sOut As String
    Dim iCount As Long

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    Set oProduct = oDoc.GetItem("Product")
    If oProduct Is Nothing Then Set oProduct = oDoc
    If oProduct.Children Is Nothing Then
        MsgBox "Alt bileşen yok."
        Exit Sub
    End If

    iCount = 0
    sOut = "BOM (Children):" & vbCrLf
    For Each oChild In oProduct.Children
        iCount = iCount + 1
        sOut = sOut & iCount & ": " & oChild.Name & vbCrLf
    Next oChild
    MsgBox sOut
End Sub
```

------------------------------------------------------------

## 7. Yapılabilecek işlemler – Çizim (Drawing)

| İşlem | Kullanılan nesne/koleksiyon | Açıklama |
|-------|-----------------------------|----------|
| Sayfa listesi | **DrawingRoot.Sheets**, **Count**, **Item(i)** | Sayfa adı, sayısı. |
| Aktif sayfa | **DrawingRoot.ActiveSheet** veya **Sheets.Item(1)** | Tek sayfa. |
| Görünüm listesi | **Sheet.Views**, **Count**, **Item(i)** | Her sayfadaki görünümler. |
| Görünüm ölçeği | **DrawingView.Scale** (veya benzeri property) | Okuma/yazma. |
| Ölçüler | **View.Dimensions** | Dimension listesi, metin, değer. |
| Metinler | **View.DrawingTexts** | Çizim metinleri. |
| Çizim servisi | **CATDrawingService** (Editor-level) | Çizime özel işlemler. |

### Örnek: Drawing – Tüm sayfalar ve görünüm sayıları

```vba
Sub DrawingSayfaVeGorusSayilari()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oDraw As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oViews As Object
    Dim i As Long
    Dim j As Long
    Dim sOut As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    Set oDraw = oDoc.GetItem("DrawingRoot")
    If oDraw Is Nothing Then Set oDraw = oDoc
    Set oSheets = oDraw.Sheets
    If oSheets Is Nothing Then Exit Sub

    sOut = "Çizim: " & oDoc.Name & vbCrLf
    For i = 1 To oSheets.Count
        Set oSheet = oSheets.Item(i)
        If Not oSheet Is Nothing Then
            sOut = sOut & "Sayfa " & i & ": " & oSheet.Name
            Set oViews = oSheet.Views
            If Not oViews Is Nothing Then
                sOut = sOut & " (Görünüm sayısı: " & oViews.Count & ")"
            End If
            sOut = sOut & vbCrLf
        End If
    Next i
    MsgBox sOut
End Sub
```

------------------------------------------------------------

## 8. Yapılabilecek işlemler – Dosya ve ortam

| İşlem | Kullanılan | Açıklama |
|-------|-------------|----------|
| Dosya var mı | **FileSystem.Exists(yol)** | Boolean. |
| Dosya boyutu | **FileSystem.GetFile(yol).Size** | Byte. |
| Klasör içeriği | **FileSystem.GetFolder(yol).Files** | Dosya listesi. |
| Metin dosyası okuma/yazma | **File.OpenAsTextStream** (FileSystem ile) veya VBA **Open ... For Input/Output/Append** | Taşınabilirlik için FileSystem tercih edilir. |
| Ortam değişkeni | **SystemService.Environ("DEĞİŞKEN_ADI")** | Örn. ROOT_FOLDER. |
| Yol birleştirme | **SystemService.ConcatenatePaths(yol1, yol2)** | Platform uyumlu. |

------------------------------------------------------------

## 9. Özet tablo: Servis erişimi

| Erişim | Kullanım | Örnek servisler |
|--------|----------|------------------|
| **GetSessionService("...")** | Application üzerinden | VisuServices, Search, PLMNewService, PLMOpenService, MATPLMService |
| **GetService("...")** | ActiveEditor üzerinden | InertiaService, MeasureService, CATDrawingService, PLMProductService |
| **FileSystem** | Application.FileSystem | GetFile, GetFolder, Exists |
| **SystemService** | Application.SystemService | Environ, ConcatenatePaths |

------------------------------------------------------------

## Sonraki adım

**Tüm rehber:** [README](README.md). İlgili: [13](13-Erisim-ve-Kullanim-Rehberi.md) (erişim tabloları), [14](14-VBA-ve-Excel-Etkilesimi.md) (Excel), [15](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md) (dosya diyalogları). Servis adları ve metod imzaları: **VBA_API_REFERENCE.md** ve **Help/text/**.

