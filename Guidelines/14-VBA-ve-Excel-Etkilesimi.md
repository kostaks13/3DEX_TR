# 14. VBA’dan Excel ile Etkileşim

3DExperience VBA makrolarından **Excel**’e bağlanıp çalışma kitabı açma, hücre okuma/yazma, kaydetme ve kapatma yapabilirsiniz. Bu dokümanda erişim yolları ve kullanım örnekleri verilir.

------------------------------------------------------------

## 1. Excel’e nereden erişilir?

| İhtiyaç | VBA yolu | Açıklama |
|--------|----------|----------|
| Yeni Excel oturumu başlatmak | `CreateObject("Excel.Application")` | Excel arka planda açılır; görünür yapmak için `Visible = True`. |
| Zaten açık Excel’e bağlanmak | `GetObject(, "Excel.Application")` | Açık Excel yoksa hata verir; On Error ile kontrol edin. |
| Belirli dosyaya bağlanmak | `GetObject("C:\Tam\yol\dosya.xlsx")` | Dosya açıksa o örneğe bağlanır. |

3DExperience VBA ortamında genelde **late binding** (tüm nesneleri `As Object` ile tanımlama) kullanılır; böylece projeye “Microsoft Excel Object Library” referansı eklemeniz gerekmez.

------------------------------------------------------------

## 2. Temel zincir: Application → Workbook → Worksheet → Range/Cells

```
Excel.Application    → oExcel
oExcel.Workbooks     → çalışma kitapları koleksiyonu
oExcel.Workbooks.Open("C:\yol\dosya.xlsx")  → oWb (Workbook)
oWb.Worksheets.Item(1) veya oWb.Sheets("Sayfa1")  → oWs (Worksheet)
oWs.Range("A1") veya oWs.Cells(satir, sutun)  → hücre
oWs.Range("A1").Value  → okuma/yazma
```

------------------------------------------------------------

## 3. Excel’i başlatma ve çalışma kitabı açma

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Excel uygulamasını başlatır, çalışma kitabı açar.

Sub ExcelAcVeOku()
    Dim oExcel As Object
    Dim oWb As Object
    Dim oWs As Object
    Dim sYol As String
    Dim vDeger As Variant

    On Error GoTo HataYakala
    Set oExcel = CreateObject("Excel.Application")
    If oExcel Is Nothing Then
        MsgBox "Excel başlatılamadı. Excel yüklü mü?"
        Exit Sub
    End If

    oExcel.Visible = True
    oExcel.DisplayAlerts = False

    sYol = "C:\Temp\Ornek.xlsx"
    Set oWb = oExcel.Workbooks.Open(sYol)
    If oWb Is Nothing Then
        oExcel.Quit
        Exit Sub
    End If

    Set oWs = oWb.Worksheets.Item(1)
    vDeger = oWs.Range("A1").Value
    MsgBox "A1 değeri: " & vDeger

    oWb.Close SaveChanges:=False
    oExcel.Quit
    Set oWs = Nothing
    Set oWb = Nothing
    Set oExcel = Nothing
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then oExcel.Quit
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

------------------------------------------------------------

## 4. Hücre okuma ve yazma

| Ne yapmak istiyorsun? | Nasıl yaparsın? |
|----------------------|-----------------|
| Tek hücre okumak | `oWs.Range("A1").Value` veya `oWs.Cells(1, 1).Value` |
| Tek hücre yazmak | `oWs.Range("A1").Value = "Metin"` veya `oWs.Cells(1, 1).Value = 100` |
| Aralık okumak | `oWs.Range("A1:B10").Value` → 2 boyutlu dizi döner |
| Aralık yazmak | `oWs.Range("A1:B2").Value = Array(Array(1,2), Array(3,4))` veya hücre hücre döngü |
| Satır/sütun numarasıyla | `oWs.Cells(iSatir, iSutun).Value` (örn. Cells(3, 2) = B3) |

**Örnek: Bir sütunu okuyup listeleme**

```vba
Sub ExcelSutunOku()
    Dim oExcel As Object
    Dim oWb As Object
    Dim oWs As Object
    Dim i As Long
    Dim v As Variant
    Dim sOut As String

    Set oExcel = CreateObject("Excel.Application")
    Set oWb = oExcel.Workbooks.Open("C:\Temp\Liste.xlsx")
    Set oWs = oWb.Worksheets.Item(1)

    sOut = "A sütunu:" & vbCrLf
    For i = 1 To 10
        v = oWs.Cells(i, 1).Value
        If Not IsEmpty(v) And v <> "" Then sOut = sOut & v & vbCrLf
    Next i

    oWb.Close SaveChanges:=False
    oExcel.Quit
    MsgBox sOut
End Sub
```

**Örnek: Hücrelere yazma ve kaydetme**

```vba
Sub ExcelYazVeKaydet()
    Dim oExcel As Object
    Dim oWb As Object
    Dim oWs As Object

    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = False
    Set oWb = oExcel.Workbooks.Add
    Set oWs = oWb.Worksheets.Item(1)

    oWs.Cells(1, 1).Value = "Parametre"
    oWs.Cells(1, 2).Value = "Değer"
    oWs.Cells(2, 1).Value = "Length.1"
    oWs.Cells(2, 2).Value = 100

    oWb.SaveAs "C:\Temp\Cikti.xlsx"
    oWb.Close SaveChanges:=False
    oExcel.Quit
    Set oExcel = Nothing
    MsgBox "Dosya kaydedildi: C:\Temp\Cikti.xlsx"
End Sub
```

------------------------------------------------------------

## 5. Zaten açık Excel’e bağlanma

```vba
Sub ExcelAcikOlanaBaglan()
    Dim oExcel As Object
    Dim oWb As Object

    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    On Error GoTo 0
    If oExcel Is Nothing Then
        MsgBox "Excel çalışmıyor. Önce Excel açın."
        Exit Sub
    End If

    If oExcel.Workbooks.Count = 0 Then
        MsgBox "Açık çalışma kitabı yok."
        Exit Sub
    End If

    Set oWb = oExcel.ActiveWorkbook
    MsgBox "Aktif kitap: " & oWb.Name
End Sub
```

------------------------------------------------------------

## 6. 3DExperience + Excel birlikte: Parametreleri Excel’e yazma

Aktif parçadaki parametreleri okuyup Excel’e (veya yeni oluşturulan bir kitaba) satır satır yazan örnek:

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Aktif Part parametrelerini Excel'e yazar.

Sub ParametreleriExcelYaz()
    Dim oApp As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim oExcel As Object
    Dim oWb As Object
    Dim oWs As Object
    Dim i As Long
    Dim sDosya As String

    On Error GoTo HataYakala

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience yok.": Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oApp.ActiveDocument
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametre yok.": Exit Sub

    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True
    Set oWb = oExcel.Workbooks.Add
    Set oWs = oWb.Worksheets.Item(1)

    oWs.Cells(1, 1).Value = "Parametre"
    oWs.Cells(1, 2).Value = "Değer"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            oWs.Cells(i + 1, 1).Value = oParam.Name
            oWs.Cells(i + 1, 2).Value = oParam.Value
        End If
    Next i

    sDosya = "C:\Temp\ParametreListe.xlsx"
    oWb.SaveAs sDosya
    MsgBox "Kaydedildi: " & sDosya
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then oExcel.Quit
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

------------------------------------------------------------

## 7. Excel’den okuyup Part parametrelerine yazma

Excel’de A sütununda parametre adı, B sütununda yeni değer varsa; bunları okuyup Part’a yazan örnek:

```vba
Sub ExceldenPartaYaz()
    Dim oApp As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim oExcel As Object
    Dim oWb As Object
    Dim oWs As Object
    Dim i As Long
    Dim sAd As String
    Dim dDeger As Double
    Dim sYol As String

    sYol = "C:\Temp\ParametreGiris.xlsx"
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub

    On Error Resume Next
    Set oExcel = GetObject(sYol)
    On Error GoTo 0
    If oExcel Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
        Set oWb = oExcel.Workbooks.Open(sYol)
    Else
        Set oWb = oExcel.ActiveWorkbook
    End If
    Set oWs = oWb.Worksheets.Item(1)

    i = 2
    Do While True
        sAd = Trim(oWs.Cells(i, 1).Value & "")
        If sAd = "" Then Exit Do
        dDeger = oWs.Cells(i, 2).Value
        Set oParam = oParams.Item(sAd)
        If Not oParam Is Nothing Then oParam.Value = dDeger
        i = i + 1
    Loop

    oPart.Update
    oWb.Close SaveChanges:=False
    If oExcel.Workbooks.Count = 0 Then oExcel.Quit
    MsgBox "Parametreler güncellendi."
End Sub
```

------------------------------------------------------------

## 8. Sık kullanılan Excel nesneleri

| Nesne / özellik | Açıklama |
|-----------------|----------|
| **Application** | Excel uygulaması. Visible, DisplayAlerts, Quit. |
| **Workbooks** | Açık kitaplar. Open(yol), Add, Count. |
| **Workbook** | Tek kitap. Worksheets, Sheets, Save, SaveAs(yol), Close(SaveChanges). |
| **Worksheets** / **Sheets** | Sayfalar. Item(1), Item("SayfaAdi"), Count. |
| **Worksheet** | Tek sayfa. Range("A1"), Cells(satir, sutun). |
| **Range("A1")** | Hücre veya aralık. .Value (okuma/yazma), .Formula. |
| **Cells(i, j)** | Satır i, sütun j (1’den başlar). .Value. |

------------------------------------------------------------

## 9. Dikkat edilecekler

- **Excel’i kapatma:** İşiniz bitince `oWb.Close SaveChanges:=True/False`, ardından `oExcel.Quit` çağırın; nesneleri `Set ... = Nothing` yapın. Aksi halde Excel süreci arkada kalabilir.
- **Dosya açık mı:** `GetObject(, "Excel.Application")` kullanıyorsanız Excel’in zaten çalıştığından emin olun; yoksa CreateObject ile yeni örnek başlatın.
- **Hata yakalama:** Excel dosyası yok, sayfa yok veya Excel yüklü değilse hata alırsınız; On Error GoTo ve Nothing kontrolleri kullanın.
- **Referans:** 3DExperience VBA’da Excel kütüphanesi referansı olmayabilir; tüm değişkenleri `As Object` tanımlayarak late binding kullanın.

------------------------------------------------------------

## 10. Özet: Excel erişim ve kullanım

| Ne yapmak istiyorsun? | Neyi kullanırsın? | Nasıl yaparsın? |
|----------------------|-------------------|------------------|
| Excel’i başlatmak | `CreateObject("Excel.Application")` | Set oExcel = ... ; Visible = True isteğe bağlı |
| Dosya açmak | `oExcel.Workbooks.Open("C:\yol\dosya.xlsx")` | Set oWb = ... |
| Yeni kitap | `oExcel.Workbooks.Add` | Set oWb = ... |
| İlk sayfa | `oWb.Worksheets.Item(1)` | Set oWs = ... |
| Hücre okumak | `oWs.Range("A1").Value` veya `oWs.Cells(1,1).Value` | Variant döner |
| Hücre yazmak | `oWs.Cells(1,1).Value = deger` | |
| Kaydetmek | `oWb.SaveAs("C:\yol\dosya.xlsx")` | Yeni dosyaysa SaveAs, mevcut dosyaysa Save |
| Kapatmak | `oWb.Close SaveChanges:=False` | Sonra oExcel.Quit |

------------------------------------------------------------

## İlgili dokümanlar

**Tüm rehber:** [README](README.md). İlgili: [13](13-Erisim-ve-Kullanim-Rehberi.md) (erişim), [15](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md) (dosya seç/kaydet diyaloğu), [10](10-Ornek-Proje-Bastan-Sona-Bir-Makro.md) (parametre → dosya örnekleri).
