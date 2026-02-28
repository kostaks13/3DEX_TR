# 8. Sık Kullanılan API’ler

Bu dokümanda 3DExperience VBA ile sık karşılaşacağınız **parça**, **geometri** ve **çizim** tarafı API kullanımlarına kısa örnekler verilir. Tam liste ve parametreler için projedeki **VBA_API_REFERENCE.md** ve Help klasöründeki **resmi PDF'lere** (Native Apps Automation, Automation Reference) bakın. **Help PDF'lerini hangi aşamada nasıl kullanacağınız** için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md).

**Neye nereden erişilir ve neyi nasıl kullanırsın?** — Erişim yolları (VBA) ve kullanım tabloları için **13. doküman:** [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md). **Sık hatalar ve dikkat edilecekler** (Nothing, Update, V5 API vb.) için **18. doküman:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

**Bu dokümanda:** Part, Parameters, Shapes; Drawing (Sheets, Views); Product (Children); Editor-level ve Session-level servis; kod örnekleri.

════════════════════════════════════════════════════════════════════════════════

## Genel kural: Nesneye nasıl ulaşılır?

Önce **Application** → **ActiveDocument**; belge türüne göre **Part**, **Product** veya **Drawing** nesnesini alırsınız. Tüm aşağıdaki örnekler bu zincirin devamıdır; sürüme göre property/method isimleri değişebilir.

════════════════════════════════════════════════════════════════════════════════

## Parça (Part) – Belge ve gövde

```
  oPart (Part belgesi)
    ├── .MainBody          ──►  Ana gövde
    ├── .Shapes            ──►  Koleksiyon (Pad, Pocket, Sketch …)  →  .Count, .Item(i)
    ├── .Parameters        ──►  Koleksiyon (uzunluk, açı …)        →  .Item("Length.1"), .Value
    ├── .Bodies            ──►  Bodies (OrderedGeometricalSets)
    └── .Update            ──►  Değişiklikleri uygula (döngü dışında bir kez)
```

- **Aktif parça belgesi:** `oDoc` zaten parça ise `oPart = oDoc` veya `Set oPart = oDoc.GetItem("Part")` (API’ye göre).  
- **MainBody:** Ana gövde — `oPart.MainBody`.  
- **Shapes:** Gövdedeki şekiller (Pad, Pocket, Sketch vb.) — `oPart.Shapes` veya `oPart.MainBody.Shapes`.  
- **Parameters:** Uzunluk, açı vb. parametreler — `oPart.Parameters` veya `oPart.Parameters.Item("isim")`.

Referansta **Part**, **Shape**, **Parameters**, **Body** sınıflarını arayın; `GetItem`, `Item`, `Count` kullanımını inceleyin.

════════════════════════════════════════════════════════════════════════════════

## Parametre okuma ve yazma

Parametreler genelde **Name** ve **Value** (veya benzeri) property’lere sahiptir. Örnek (isimler sürüme göre değişir):

```vba
Dim oParam As Object
Set oParam = oPart.Parameters.Item("Length.1")   ' veya GetItem("Length.1")
Dim deger As Double
deger = oParam.Value   ' Okuma
oParam.Value = 50.5    ' Yazma (varsa)
oPart.Update           ' Güncelleme (gerekirse)
```

`VBA_API_REFERENCE.md` içinde **Parameter**, **Parameters**, **Value** geçen metodlara bakın.

════════════════════════════════════════════════════════════════════════════════

## Shapes (şekiller) üzerinde döngü

```vba
Dim oShapes As Object
Set oShapes = oPart.Shapes   ' veya MainBody.Shapes

Dim i As Long
For i = 1 To oShapes.Count
    Dim oShape As Object
    Set oShape = oShapes.Item(i)
    ' oShape.Name, oShape.Type vb. kullanılır
Next i
```

Koleksiyonlarda **Item(1)** çoğu yerde 1’den başlar; **For Each** de kullanılabilir (referansta ilgili sınıflarda `__iter__` veya koleksiyon açıklamalarına bakın).

════════════════════════════════════════════════════════════════════════════════

## Çizim (Drawing) – Sayfalar ve görünümler

- **DrawingRoot / Drawing:** Çizim belgesi — `oDoc` çizimse doğrudan veya `oDoc.GetItem("DrawingRoot")` vb.  
- **Sheets:** Sayfalar — `oDrawing.Sheets`, `oSheets.Item(1)`.  
- **Views:** Bir sayfadaki görünümler — `oSheet.Views`, `oView.Views.Item(1)`.

Örnek (API isimleri sürüme göre uyarlanmalı):

```vba
Dim oSheet As Object
Set oSheet = oDrawing.Sheets.ActiveSheet   ' veya .Item(1)
Dim oViews As Object
Set oViews = oSheet.Views
```

Referansta **DrawingSheet**, **DrawingView**, **DrawingViews** gibi sınıfları arayın.

════════════════════════════════════════════════════════════════════════════════

## Çizimde ölçü (Dimension) ve metin

- **Dimensions:** Görünümdeki ölçüler — `oView.Dimensions`.  
- **DrawingTexts:** Metinler — `oView.DrawingTexts` veya benzeri.  
- Yeni ölçü/metin ekleme: İlgili koleksiyonun **Add** metoduna bakın (`VBA_API_REFERENCE.md` içinde “Add” ve “DrawingDimension” / “DrawingText” araması yapın).

════════════════════════════════════════════════════════════════════════════════

## Montaj (Product) – Bileşenler

- **Children / Products:** Alt bileşenler — `oProduct.Children` veya `oProduct.Products`.  
- **GetItem / Item:** İsim veya indeks ile bileşen — `oProduct.Children.Item("Parça1")` veya `.Item(1)`.

```vba
Dim oChild As Object
For Each oChild In oProduct.Children
    ' oChild.Name, oChild.Position vb.
Next oChild
```

════════════════════════════════════════════════════════════════════════════════

## Ortak pattern’ler

| İhtiyaç | Nerede bakılır |
|--------|-----------------|
| Koleksiyondan eleman alma | `Item(i)`, `GetItem("isim")` |
| Eleman sayısı | `Count` (property veya metod) |
| Koleksiyonda dolaşma | `For i = 1 To .Count` veya `For Each ... In ...` |
| Yeni eleman ekleme | Koleksiyonun `Add` metodu |
| Değer okuma/yazma | İlgili nesnenin `Value`, `Name` vb. property’leri |

Tüm bu sınıf ve metodların tam imzaları **VBA_API_REFERENCE.md** içinde “Ne yapar” ve “Örnek” ile listelenmiştir; sürümünüze uygun olanı seçin.

════════════════════════════════════════════════════════════════════════════════

## Editor-level vs Session-level servis (Help’ten)

**Help-Native Apps Automation** (Foundation Object Model Map) dokümanına göre:

- **Editor-level servis:** Aktif editördeki **belirli PLM kök nesnesine** (Part, Product, Drawing) bağlıdır. Örnek: InertiaService, MeasureService, DrawingService, PLMProductService. Erişim: `CATIA.ActiveEditor.GetService("ServisAdi")`.
- **Session-level servis:** Oturum genelinde, editörden bağımsız işlemler (PLM arama, yeni belge açma, malzeme atama vb.). Örnek: SearchService, PLMNewService, PLMOpenService, MaterialService. Erişim: `CATIA.GetSessionService("ServisAdi")`.

**Kural:** Editor-level bir servis içinden session-level servis çağrısı yapmayın; kilitleşme riski vardır. Önce editor-level işinizi bitirip sonra session-level işleme geçin.

Servis tanımlayıcıları (GetService / GetSessionService’e verilecek string) Help’teki “Service Identifier” tablolarında listelenir; projedeki **VBA_API_REFERENCE.md** içinde de ilgili sınıflar aranabilir.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Tam makro – Tüm parametreleri listele

Aktif parçadaki parametrelerin adını ve değerini mesaj kutusunda (veya dosyaya) yazdıran örnek. Parametre koleksiyonunda Count ve Item(i) kullanılır (sürüme göre Item ile indeks veya isim alınabilir).

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Aktif parçanın parametre listesini mesajda gösterir.

Sub TumParametreleriListele()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim sListe As String
    Dim iCount As Long

    On Error GoTo HataYakala
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "Uygulama yok.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    iCount = oParams.Count
    sListe = "Parametre sayısı: " & iCount & vbCrLf
    For i = 1 To iCount
        On Error Resume Next
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            sListe = sListe & oParam.Name & " = " & oParam.Value & vbCrLf
        End If
        On Error GoTo 0
    Next i
    MsgBox sListe
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Parametre değeri oku – İsimle

Kullanıcıdan parametre adı alıp sadece değerini okuyan (yazmayan) makro:

```vba
Sub ParametreDegeriOku()
    Dim oApp As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim sParamAdi As String
    Dim dDeger As Double

    sParamAdi = InputBox("Parametre adı (örn. Length.1):", "Oku", "Length.1")
    If sParamAdi = "" Then Exit Sub

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub

    On Error Resume Next
    Set oParam = oParams.Item(sParamAdi)
    On Error GoTo 0
    If oParam Is Nothing Then
        MsgBox "Parametre bulunamadı: " & sParamAdi
        Exit Sub
    End If
    dDeger = oParam.Value
    MsgBox sParamAdi & " = " & dDeger
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Shapes döngüsü – İsim ve tip

Parçadaki tüm şekillerin (Pad, Pocket, Sketch vb.) adını ve tip bilgisini toplayan örnek:

```vba
Sub ShapesListele()
    Dim oApp As Object
    Dim oPart As Object
    Dim oShapes As Object
    Dim oShape As Object
    Dim i As Long
    Dim sListe As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oShapes = oPart.Shapes
    If oShapes Is Nothing Then Exit Sub

    sListe = "Shapes:" & vbCrLf
    For i = 1 To oShapes.Count
        Set oShape = oShapes.Item(i)
        If Not oShape Is Nothing Then
            sListe = sListe & i & ": " & oShape.Name & vbCrLf
        End If
    Next i
    MsgBox sListe
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Çizim – Aktif sayfa ve görünüm sayısı

Çizim belgesinde aktif sayfayı alıp o sayfadaki görünüm sayısını gösteren kavramsal örnek (API isimleri sürüme göre değişir):

```vba
Sub CizimSayfaVeGorusSayisi()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oDraw As Object
    Dim oSheet As Object
    Dim oViews As Object
    Dim iCount As Long

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    Set oDraw = oDoc.GetItem("DrawingRoot")
    If oDraw Is Nothing Then Set oDraw = oDoc

    Set oSheet = oDraw.ActiveSheet
    If oSheet Is Nothing Then Set oSheet = oDraw.Sheets.Item(1)
    Set oViews = oSheet.Views
    If oViews Is Nothing Then
        MsgBox "Views alınamadı."
        Exit Sub
    End If
    iCount = oViews.Count
    MsgBox "Aktif sayfadaki görünüm sayısı: " & iCount
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Montaj – Children döngüsü

Product (montaj) kökü altındaki tüm alt bileşenlerin adını listeler:

```vba
Sub MontajBilesenListele()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oProduct As Object
    Dim oChild As Object
    Dim sListe As String

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
    sListe = "Bileşenler:" & vbCrLf
    For Each oChild In oProduct.Children
        sListe = sListe & oChild.Name & vbCrLf
    Next oChild
    MsgBox sListe
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: GetService (Editor) – InertiaService fikri

Aktif editördeki parça için kütle bilgisi almak üzere InertiaService kullanımı (Help’teki servis adı sürüme göre değişir):

```vba
Sub KutleBilgisiAl()
    Dim oApp As Object
    Dim oEditor As Object
    Dim oInertia As Object
    Dim dMass As Double

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then Exit Sub

    On Error Resume Next
    Set oInertia = oEditor.GetService("InertiaService")
    On Error GoTo 0
    If oInertia Is Nothing Then
        MsgBox "InertiaService alınamadı."
        Exit Sub
    End If
    ' oInertia ile kütle hesaplanır (API'ye göre GetMass, Compute vb.)
    MsgBox "InertiaService alındı; kütle API dokümantasyonuna göre hesaplanır."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Parametre adına göre filtreleme (Like / Left)

Sadece adı “Length” ile başlayan parametreleri işlemek için VBA **Like** veya **Left** kullanılabilir:

```vba
Sub LengthParametreleriniGuncelle()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim sName As String
    Dim dYeni As Double
    dYeni = 75
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            sName = oParam.Name
            If Left(sName, 6) = "Length" Or sName Like "Length*" Then
                oParam.Value = dYeni
            End If
        End If
    Next i
    oPart.Update
    MsgBox "Length parametreleri güncellendi."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Çizimde görünüm ölçeği (Scale)

Bir görünümün ölçeğini okumak veya yazmak (API’de genelde DrawingView veya View’da Scale property’si):

```vba
Sub GorusOlcegiOrnek()
    Dim oDraw As Object
    Dim oSheet As Object
    Dim oViews As Object
    Dim oView As Object
    Dim dScale As Double
    Set oDraw = GetObject(, "CATIA.Application").ActiveDocument.GetItem("DrawingRoot")
    If oDraw Is Nothing Then Exit Sub
    Set oSheet = oDraw.ActiveSheet
    If oSheet Is Nothing Then Exit Sub
    Set oViews = oSheet.Views
    If oViews Is Nothing Or oViews.Count = 0 Then Exit Sub
    Set oView = oViews.Item(1)
    If oView Is Nothing Then Exit Sub
    On Error Resume Next
    dScale = oView.Scale
    On Error GoTo 0
    MsgBox "İlk görünüm ölçeği: " & dScale
End Sub
```

Property adı (Scale, ScaleValue vb.) sürüme göre değişir; referansta DrawingView araması yapın.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Parametre adını Name property’sinden okuma

Parametre nesnesinin **Name** property’si parametre adını (örn. "Length.1") verir; döngüde her parametrenin adını böyle alırsınız:

```vba
Sub ParametreAdiOrnek()
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Set oParams = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part").Parameters
    If oParams Is Nothing Then Exit Sub
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then Debug.Print oParam.Name & " = " & oParam.Value
    Next i
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: GetItem ile isim veya indeks

Çoğu koleksiyonda **Item** hem indeks (1’den başlayan sayı) hem isim (string) alır: `Item(1)` veya `Item("Length.1")`. İsim kullanırken yazım hatası veya farklı dilde isimlendirme (locale) konusuna dikkat edin.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Koleksiyon Count kontrolü

Koleksiyona erişmeden önce **Count** kontrolü yapın; 0 ise döngüye girmeyin:

```vba
Sub GuvenliCountOrnek()
    Dim oShapes As Object
    Dim i As Long
    Set oShapes = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part").Shapes
    If oShapes Is Nothing Then MsgBox "Shapes yok.": Exit Sub
    If oShapes.Count = 0 Then MsgBox "Hiç şekil yok.": Exit Sub
    For i = 1 To oShapes.Count
        ' ... Item(i) güvenle kullanılır
    Next i
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Parametre tipi kontrolü (Double vs Integer)

Bazı parametreler açı (Angle), bazıları uzunluk (Length). Değer atarken tip uyumluluğuna dikkat edin; gerekirse **CDbl**, **CLng** kullanın:

```vba
Sub ParametreTipOrnek()
    Dim oParam As Object
    Dim vVal As Variant
    Set oParam = oPart.Parameters.Item("Length.1")
    vVal = 100.5
    oParam.Value = CDbl(vVal)
    ' Açı parametresi için: oParam.Value = CDbl(açıDerece)
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Çizimde Dimensions koleksiyonu (kavramsal)

Bir görünümdeki ölçüler **Dimensions** (veya benzeri) koleksiyonunda tutulur. Count ve Item(i) ile döngü; her ölçünün metni, değeri veya birimi property’lerden okunur/yazılır. Tam API adları sürüme göre değişir; VBA_API_REFERENCE.md içinde “Dimension”, “DrawingDimension” araması yapın.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Product – Occurrence vs Component

Montajda bazen **Occurrence** (örnek) ve **Component** (tanım) ayrımı yapılır. Aynı parçanın montajda birden fazla örneği olabilir; her biri bir occurrence’dır. API’de **RepInstances**, **Children**, **Occurrences** gibi koleksiyon isimleri kullanılır; Help’teki “Product”, “VPMOccurrence” bölümlerine bakın.

════════════════════════════════════════════════════════════════════════════════

## Yapılabilecek işlemler – Kısa detay

Aşağıdaki işlemler **Part**, **Product** ve **Drawing** ile sık yapılır. Detaylı açıklama ve ek kod örnekleri için **12. doküman:** [12-Servisler-ve-Yapilabilecek-Islemler.md](12-Servisler-ve-Yapilabilecek-Islemler.md).

### Part ile yapılabilecekler

- **Parametre:** Okuma (Name, Value), yazma (Value + Update), listeleme (Count, Item(i)), filtreleme (Like "Length*"), toplu güncelleme (döngü + tek Update).
- **Shapes:** Listeleme (Count, Item(i).Name), tip kontrolü; **Bodies:** Katı gövde listesi.
- **HybridBodies / HybridShapes:** Geometrik setler ve içindeki yüzey/eğri/nokta; **HybridShapeFactory** ile yeni geometri ekleme (nokta, düzlem vb.).
- **Kütle/atalet:** **InertiaService** (Editor-level) ile kütle ve atalet bilgisi; **MeasureService** ile mesafe, alan, açı ölçümü.

### Product ile yapılabilecekler

- **Children:** Alt bileşen listesi (BOM), For Each ile dolaşma; **PLMProductService.RootOccurrence** ile montaj kökü.
- **Occurrence modeli:** VPMRootOccurrence, VPMOccurrence, Occurrences koleksiyonu ile ağaç gezinme; **PLMEntity** ile GetAttributeValue / SetAttributeValue.

### Drawing ile yapılabilecekler

- **Sheets:** Sayfa listesi, ActiveSheet; **Views:** Her sayfadaki görünümler, Scale (ölçek).
- **Dimensions:** Görünümdeki ölçüler; **DrawingTexts:** Metinler. **CATDrawingService** (Editor-level) çizime özel işlemler için.

### Dosya ve ortam

- **FileSystem:** GetFile, GetFolder, Exists, dosya boyutu (Size), klasördeki dosyalar (Files). **SystemService:** Environ, ConcatenatePaths.

════════════════════════════════════════════════════════════════════════════════

## Servisler özeti ve kod kalıbı

| Tür | Erişim | Örnek tanımlayıcı |
|-----|--------|--------------------|
| Session-level | `oApp.GetSessionService("...")` | "VisuServices", "Search", "PLMNewService", "PLMOpenService" |
| Editor-level | `oApp.ActiveEditor.GetService("...")` | "InertiaService", "MeasureService", "CATDrawingService", "PLMProductService" |

**Örnek: Editor-level servis alıp kullanma**

```vba
Sub EditorServiceKalip()
    Dim oApp As Object
    Dim oEditor As Object
    Dim oSvc As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then Exit Sub
    On Error Resume Next
    Set oSvc = oEditor.GetService("InertiaService")
    On Error GoTo 0
    If oSvc Is Nothing Then
        MsgBox "Servis alınamadı (pencere türü uygun olmalı)."
        Exit Sub
    End If
    ' oSvc ile işlem (örn. kütle hesaplama)
End Sub
```

**Örnek: Session-level servis alıp kullanma**

```vba
Sub SessionServiceKalip()
    Dim oApp As Object
    Dim oSvc As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    On Error Resume Next
    Set oSvc = oApp.GetSessionService("VisuServices")
    On Error GoTo 0
    If oSvc Is Nothing Then
        MsgBox "VisuServices alınamadı."
        Exit Sub
    End If
    ' oSvc.Cameras vb. kullanılır
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: FileSystem ile metin dosyasına satır ekleme (OpenAsTextStream)

Help’e göre taşınabilir kod için **FileSystem** kullanın. **File.OpenAsTextStream** ile metin dosyası açılıp yazılabilir (imza Help’te):

```vba
Sub FileSystemLogSatir()
    Dim oApp As Object
    Dim oFS As Object
    Dim oFile As Object
    Dim oStream As Object
    Dim sPath As String
    Dim sSatir As String

    sPath = "C:\Temp\macro_log.txt"
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then Exit Sub
    On Error Resume Next
    If Not oFS.Exists(sPath) Then
        ' Dosya yoksa oluşturmak için CreateFile veya önce boş dosya oluşturma (API'ye bakın)
    End If
    Set oFile = oFS.GetFile(sPath)
    On Error GoTo 0
    If oFile Is Nothing Then Exit Sub
    Set oStream = oFile.OpenAsTextStream(2, 0)
    If oStream Is Nothing Then Exit Sub
    sSatir = Format(Now, "yyyy-mm-dd hh:nn:ss") & "  INFO  Makro çalıştı" & vbCrLf
    oStream.Write sSatir
    oStream.Close
    MsgBox "Log satırı yazıldı."
End Sub
```

*(OpenAsTextStream parametreleri sürüme göre değişir; ForAppend = 2 genelde ekleme modudur. VBA_API_REFERENCE veya Help’te TextStream açın.)*

════════════════════════════════════════════════════════════════════════════════

════════════════════════════════════════════════════════════════════════════════

## Uygulamalı alıştırma – Yaparak öğren

**Amaç:** Part’tan parametre okumak, tek parametre yazmak ve Shapes sayısını göstermek.  
**Süre:** Yaklaşık 20 dakika. **Gereksinim:** 3DExperience açık, en az bir parametre içeren Part belgesi açık.  
**Zorluk:** Orta

| Adım | Ne yapacaksınız | Kontrol |
|------|------------------|--------|
| **1** | Yeni Sub: `ParametreOkuAlistirma`. GetObject → ActiveDocument → GetItem("Part") zinciri ile oPart alın; Nothing kontrolleri ekleyin. `Set oParams = oPart.Parameters`. `If oParams Is Nothing Or oParams.Count = 0 Then MsgBox "Parametre yok": Exit Sub`. | Part ve Parameters alındı mı? |
| **2** | Parametre adını bilmiyorsanız: `For i = 1 To oParams.Count` döngüsünde `Set oParam = oParams.Item(i)`, `Debug.Print oParam.Name & " = " & oParam.Value`. F5 çalıştırıp **Ctrl+G** (Immediate) ile çıktıyı görün. İlk parametre adını (örn. Length.1) not alın. | Parametre adları listelendi mi? |
| **3** | Bir parametre değerini mesajla gösterin: `Set oParam = oParams.Item("Length.1")` (veya not aldığınız ad). `If oParam Is Nothing Then MsgBox "Parametre yok": Exit Sub`. `MsgBox oParam.Name & " = " & oParam.Value`. | Tek parametre değeri görünüyor mu? |
| **4** | Yeni Sub: `ParametreYazAlistirma`. oPart alın. `Set oParam = oParams.Item("Length.1")`. Eski değeri bir değişkende saklayın: `dEski = oParam.Value`. `oParam.Value = dEski + 10` (veya sabit bir değer). **Döngü dışında** `oPart.Update` çağırın. MsgBox ile “Güncellendi” deyin. Çalıştırın; modelde değişiklik görmelisiniz. | Update tek sefer, döngü dışında mı? |
| **5** | Yeni Sub: `ShapesSayisiAlistirma`. oPart alın. `Set oShapes = oPart.Shapes` (veya MainBody.Shapes; sürüme göre). `If oShapes Is Nothing Then MsgBox "Shapes yok": Exit Sub`. `MsgBox "Shapes sayisi: " & oShapes.Count`. Çalıştırın. | Shapes sayısı doğru mu? |

**Beklenen sonuç:** Parametre listesi Immediate’da; tek parametre değeri mesajda; bir parametre güncellenip Update ile uygulandı; Shapes sayısı mesajda.

════════════════════════════════════════════════════════════════════════════════

## Kontrol listesi

- [ ] Part, Shapes, Parameters erişimini biliyorum  
- [ ] Parametre okuma/yazma örneğini referanstan bulabiliyorum  
- [ ] Drawing Sheets/Views kullanımını biliyorum  
- [ ] Product Children ile bileşen listesini alabileceğimi biliyorum  
- [ ] GetService (editor) ile GetSessionService (session) farkını biliyorum  

════════════════════════════════════════════════════════════════════════════════

## Kendinizi test edin

1. **Parametre değerini yazdıktan sonra** mutlaka çağırmanız gereken tek metod nedir ve nerede çağrılmalı (döngü içinde mi, dışında mı)?  
2. Drawing belgesinde **sayfa (sheet) listesine** hangi nesne zinciri ile erişirsiniz? (Document → ? → Sheets)  
3. **Product** (montaj) içindeki alt bileşenler koleksiyonunun API adı nedir?

<details>
<summary>Yanıtlar (tıklayarak açın)</summary>

1. **oPart.Update** — Tüm parametre/şekil değişikliklerinden **sonra**, döngü **dışında** bir kez çağrılmalı.  
2. **oDoc** → **GetItem("DrawingRoot")** → **oDraw.Sheets**. Her sheet’te Views, Scale vb. erişilir.  
3. **Children** — `oProduct.Children` ile montajdaki alt bileşenlere erişilir; BOM veya ağaç dolaşımında kullanılır.

</details>

════════════════════════════════════════════════════════════════════════════════

## Sonraki adım

**9. doküman:** [09-Hata-Yakalama-ve-Debug.md](09-Hata-Yakalama-ve-Debug.md) — On Error, kesme noktası ve Immediate penceresi.

---

### Gezinme

| [← Önceki: 07 Makro kayıt](07-Makro-Kayit-ve-Inceleme.md) | [Rehber listesi](README.md) | [Sonraki: 09 Hata yakalama →](09-Hata-Yakalama-ve-Debug.md) |
| :--- | :--- | :--- |

**İlgili:** [12-Servisler-ve-Yapilabilecek-Islemler.md](12-Servisler-ve-Yapilabilecek-Islemler.md) (servisler, işlem detayı) · [15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md) (dosya seç/kaydet diyaloğu).
