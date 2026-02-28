# 15. Dosya Seçtirme ve Kaydetme Diyalogları

Makroda kullanıcının **dosya seçmesi** (açılacak veya işlenecek dosya) veya **kaydetme yeri ve adı belirlemesi** için diyalog pencereleri kullanılır. Bu dokümanda **dosya açma diyaloğu**, **dosya kaydetme diyaloğu** ve **klasör seçme** yöntemleri anlatılır.

------------------------------------------------------------

## 1. Yöntemlerin özeti

| İhtiyaç | Yöntem | Kullanım |
|--------|--------|----------|
| Dosya seçtirme (açılacak dosya) | FileDialog (Office) veya Windows API GetOpenFileName | Kullanıcı tek dosya seçer; tam yol döner. |
| Dosya kaydetme (yer + ad) | FileDialog (Save As) veya Windows API GetSaveFileName | Kullanıcı konum ve dosya adı seçer. |
| Klasör seçme | FileDialog (Folder Picker) veya API | Sadece klasör yolu döner. |
| Basit yol girişi | InputBox | Diyaloğa gerek yok; kullanıcı yolu yazar. |

3DExperience VBA ortamında **FileDialog** bazen kullanılamayabilir (Application farklı olduğu için). Bu yüzden **Windows API** ile GetOpenFileName / GetSaveFileName kullanımı aşağıda ayrıca verilir; bu yöntem sadece Windows’ta çalışır ancak 3DExperience VBA’da genelde kullanılabilir.

------------------------------------------------------------

## 2. FileDialog ile dosya seçtirme (Office)

Microsoft Office nesne modelinde **FileDialog** vardır. 3DExperience VBA’da bazen **Excel** veya **Word** CreateObject ile oluşturulup onun FileDialog’u kullanılabilir; ya da ortam FileDialog destekliyorsa doğrudan kullanılır.

**Excel üzerinden FileDialog kullanımı (örnek):**

```vba
' Excel.Application.FileDialog ile dosya seçtirme
Sub DosyaSecFileDialog()
    Dim oExcel As Object
    Dim oDlg As Object
    Dim sSecilen As String

    On Error GoTo HataYakala
    Set oExcel = CreateObject("Excel.Application")
    Set oDlg = oExcel.FileDialog(1)
    If oDlg Is Nothing Then
        oExcel.Quit
        MsgBox "FileDialog alınamadı."
        Exit Sub
    End If

    oDlg.Title = "Açılacak dosyayı seçin"
    oDlg.InitialFileName = "C:\Temp\"
    oDlg.Filters.Clear
    oDlg.Filters.Add "Excel dosyaları", "*.xlsx; *.xls"
    oDlg.FilterIndex = 1
    oDlg.AllowMultiSelect = False

    If oDlg.Show = -1 Then
        sSecilen = oDlg.SelectedItems(1)
        MsgBox "Seçilen: " & sSecilen
    Else
        MsgBox "İptal edildi."
    End If

    oExcel.Quit
    Set oExcel = Nothing
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then oExcel.Quit
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

**FileDialog türleri (Excel/Office):**

- **1** — msoFileDialogOpen: Dosya açma (tek veya çoklu seçim).
- **2** — msoFileDialogSaveAs: Farklı kaydet / kaydetme yeri ve adı.
- **4** — msoFileDialogFolderPicker: Sadece klasör seçimi.

**Özellikler:** `.Title`, `.InitialFileName`, `.Filters.Add "Açıklama", "*.uzanti"`, `.AllowMultiSelect`, `.Show` (‑1 = Tamam), `.SelectedItems(1)` (seçilen yol).

------------------------------------------------------------

## 3. FileDialog ile dosya kaydetme

Kullanıcıya “Kaydet” diyaloğu göstermek için FileDialog(2) kullanılır; seçilen yol ve dosya adı `.SelectedItems(1)` ile alınır.

```vba
Sub DosyaKaydetFileDialog()
    Dim oExcel As Object
    Dim oDlg As Object
    Dim sHedefYol As String

    Set oExcel = CreateObject("Excel.Application")
    Set oDlg = oExcel.FileDialog(2)

    oDlg.Title = "Kaydedilecek yeri ve dosya adını seçin"
    oDlg.InitialFileName = "C:\Temp\Cikti.xlsx"
    oDlg.Filters.Clear
    oDlg.Filters.Add "Excel dosyaları", "*.xlsx"
    oDlg.FilterIndex = 1

    If oDlg.Show = -1 Then
        sHedefYol = oDlg.SelectedItems(1)
        MsgBox "Kayıt yolu: " & sHedefYol
        ' Bu yolu kullanarak dosyayı kaydedin (örn. oWb.SaveAs sHedefYol)
    Else
        MsgBox "İptal edildi."
    End If

    oExcel.Quit
    Set oExcel = Nothing
End Sub
```

------------------------------------------------------------

## 4. FileDialog ile klasör seçme

Sadece klasör seçtirmek için FileDialog(4) (Folder Picker) kullanılır.

```vba
Sub KlasorSecFileDialog()
    Dim oExcel As Object
    Dim oDlg As Object
    Dim sKlasor As String

    Set oExcel = CreateObject("Excel.Application")
    Set oDlg = oExcel.FileDialog(4)

    oDlg.Title = "Klasör seçin"
    oDlg.AllowMultiSelect = False

    If oDlg.Show = -1 Then
        sKlasor = oDlg.SelectedItems(1)
        MsgBox "Seçilen klasör: " & sKlasor
    Else
        MsgBox "İptal edildi."
    End If

    oExcel.Quit
    Set oExcel = Nothing
End Sub
```

------------------------------------------------------------

## 5. Windows API ile dosya seçtirme (GetOpenFileName)

3DExperience VBA’da FileDialog kullanılamıyorsa **Windows Common Dialog** API’si kullanılır. Aşağıdaki kod modülün üstünde (Declare) ve bir yardımcı Function içinde tanımlanır.

```vba
Option Explicit
' Windows API: Dosya açma diyaloğu
#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#Else
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As LongPtr
    hInstance         As LongPtr
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As LongPtr
    lpfnHook          As LongPtr
    lpTemplateName    As String
End Type

Public Function DosyaSecAPI(Optional sBaslik As String = "Dosya seçin", _
                            Optional sFiltre As String = "Tüm dosyalar (*.*)|*.*", _
                            Optional sBaslangicDir As String = "C:\") As String
    Dim ofn As OPENFILENAME
    Dim sDosya As String
    Dim i As Long

    sDosya = Space(2000)
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = 0
    ofn.lpstrFilter = Replace(sFiltre, "|", Chr(0)) & Chr(0)
    ofn.nFilterIndex = 1
    ofn.lpstrFile = sDosya
    ofn.nMaxFile = 2000
    ofn.lpstrFileTitle = Space(256)
    ofn.nMaxFileTitle = 256
    ofn.lpstrInitialDir = sBaslangicDir
    ofn.lpstrTitle = sBaslik
    ofn.flags = 0

    If GetOpenFileName(ofn) <> 0 Then
        DosyaSecAPI = Trim(Left(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr(0)) - 1))
    Else
        DosyaSecAPI = ""
    End If
End Function

Sub OrnekDosyaSec()
    Dim sYol As String
    sYol = DosyaSecAPI("Açılacak Excel dosyası", "Excel (*.xlsx)|*.xlsx|Tüm (*.*)|*.*", "C:\Temp")
    If sYol <> "" Then
        MsgBox "Seçilen: " & sYol
    Else
        MsgBox "İptal edildi."
    End If
End Sub
```

**Filtre formatı:** Metinde `|` ile ayrılmış çiftler: "Açıklama|*.uzanti". API’de `|` yerine `Chr(0)` konur; sonuna bir `Chr(0)` daha eklenir.

------------------------------------------------------------

## 6. Windows API ile dosya kaydetme (GetSaveFileName)

Kaydetme diyaloğu için **GetSaveFileName** kullanılır; yapı aynı OPENFILENAME’dir, sadece farklı API fonksiyonu çağrılır.

```vba
#If VBA7 Then
    Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
#Else
    Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If

Public Function DosyaKaydetAPI(Optional sBaslik As String = "Kaydet", _
                              Optional sVarsayilanAd As String = "Cikti.xlsx", _
                              Optional sFiltre As String = "Excel (*.xlsx)|*.xlsx|Tüm (*.*)|*.*", _
                              Optional sBaslangicDir As String = "C:\Temp") As String
    Dim ofn As OPENFILENAME
    Dim sDosya As String

    sDosya = sVarsayilanAd & Space(2000 - Len(sVarsayilanAd))
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = 0
    ofn.lpstrFilter = Replace(sFiltre, "|", Chr(0)) & Chr(0)
    ofn.nFilterIndex = 1
    ofn.lpstrFile = sDosya
    ofn.nMaxFile = 2000
    ofn.lpstrFileTitle = Space(256)
    ofn.nMaxFileTitle = 256
    ofn.lpstrInitialDir = sBaslangicDir
    ofn.lpstrTitle = sBaslik
    ofn.lpstrDefExt = "xlsx"
    ofn.flags = 0

    If GetSaveFileName(ofn) <> 0 Then
        DosyaKaydetAPI = Trim(Left(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr(0)) - 1))
    Else
        DosyaKaydetAPI = ""
    End If
End Function

Sub OrnekDosyaKaydet()
    Dim sYol As String
    sYol = DosyaKaydetAPI("Raporu kaydet", "Rapor.xlsx", "Excel (*.xlsx)|*.xlsx", "C:\Temp")
    If sYol <> "" Then
        MsgBox "Kayıt yolu: " & sYol
    Else
        MsgBox "İptal edildi."
    End If
End Sub
```

------------------------------------------------------------

## 7. InputBox ile basit yol girişi

Diyalog kullanmadan kullanıcıdan dosya veya klasör yolu almak için **InputBox** kullanılabilir. Varsayılan değer vererek test edilebilirliği artırın (Help kuralı).

```vba
Sub DosyaYoluInputBox()
    Dim sYol As String
    sYol = Trim(InputBox("Dosya veya klasör yolunu girin:", "Yol", "C:\Temp\Ornek.xlsx"))
    If sYol = "" Then
        MsgBox "İptal edildi."
        Exit Sub
    End If
    ' sYol ile devam (var mı kontrolü için FileSystem.Exists kullanılabilir)
    MsgBox "Girilen yol: " & sYol
End Sub
```

------------------------------------------------------------

## 8. Tam akış örneği: Dosya seç → Excel’de aç → Bir hücreyi oku

Önce dosya seçtirme (FileDialog veya API), sonra seçilen yolu Excel’de açma:

```vba
Sub DosyaSecSonraExcelAc()
    Dim sYol As String
    Dim oExcel As Object
    Dim oWb As Object
    Dim vDeger As Variant

    ' Yöntem 1: FileDialog (Excel üzerinden)
    Set oExcel = CreateObject("Excel.Application")
    With oExcel.FileDialog(1)
        .Title = "Excel dosyası seçin"
        .InitialFileName = "C:\Temp\"
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            oExcel.Quit
            MsgBox "İptal."
            Exit Sub
        End If
        sYol = .SelectedItems(1)
    End With

    Set oWb = oExcel.Workbooks.Open(sYol)
    vDeger = oWb.Worksheets(1).Range("A1").Value
    MsgBox "A1 = " & vDeger
    oWb.Close SaveChanges:=False
    oExcel.Quit
    Set oExcel = Nothing
End Sub
```

------------------------------------------------------------

## 9. Tam akış örneği: Veriyi hazırla → Kaydet diyaloğu → Kaydet

Parametre listesini hazırlayıp kullanıcıya “Nereye kaydedeceksin?” diyaloğu gösterme; seçilen yola yazma (örnek: metin dosyası).

```vba
Sub VeriHazirlaSonraKaydetDiyalog()
    Dim sIcerik As String
    Dim sKayitYolu As String
    Dim oExcel As Object
    Dim oDlg As Object
    Dim iFile As Integer

    sIcerik = "Parametre;Değer" & vbCrLf & "Length.1;100" & vbCrLf

    Set oExcel = CreateObject("Excel.Application")
    Set oDlg = oExcel.FileDialog(2)
    oDlg.Title = "Raporu kaydedeceğiniz yeri seçin"
    oDlg.InitialFileName = "C:\Temp\Rapor.csv"
    oDlg.Filters.Clear
    oDlg.Filters.Add "CSV", "*.csv"
    oDlg.FilterIndex = 1

    If oDlg.Show <> -1 Then
        oExcel.Quit
        MsgBox "İptal."
        Exit Sub
    End If
    sKayitYolu = oDlg.SelectedItems(1)
    oExcel.Quit
    Set oExcel = Nothing

    iFile = FreeFile
    Open sKayitYolu For Output As #iFile
    Print #iFile, sIcerik;
    Close #iFile
    MsgBox "Kaydedildi: " & sKayitYolu
End Sub
```

------------------------------------------------------------

## 10. Özet tablo

| Ne yapmak istiyorsun? | Nasıl yaparsın? |
|----------------------|-----------------|
| Kullanıcıya dosya seçtirmek | FileDialog(1) (Excel üzerinden) veya GetOpenFileName API |
| Kullanıcıya kaydet yeri ve adı seçtirmek | FileDialog(2) veya GetSaveFileName API |
| Kullanıcıya sadece klasör seçtirmek | FileDialog(4) (Folder Picker) |
| Diyalog olmadan yol almak | InputBox("Yol girin", "Başlık", "C:\Temp\") |
| Seçilen dosyayı açmak | Seçilen yol → Workbooks.Open(sYol) veya 3DExperience API |
| Seçilen yola kaydetmek | Seçilen yol → SaveAs(sYol) veya Open ... For Output / Print # |

------------------------------------------------------------

## 11. Dikkat edilecekler

- **FileDialog** için Excel (veya Office) CreateObject ile açılıyorsa iş bitince **Quit** ile kapatın.
- **GetOpenFileName / GetSaveFileName** sadece Windows’ta çalışır; 64 bit VBA’da `PtrSafe` ve `LongPtr` kullanın (#If VBA7). 32 bit ortamda OPENFILENAME içindeki `LongPtr` alanlarını **Long** yapmanız gerekebilir.
- Kullanıcı diyaloğu iptal ederse seçim döndürülmez; dönen dizi boş veya API 0 döner. Her zaman **boş/iptal** kontrolü yapın.
- Filtre metninde **|** karakteri API’de **Chr(0)** ile değiştirilip sonuna **Chr(0)** eklenir.

------------------------------------------------------------

## İlgili dokümanlar

**Tüm rehber:** [README](README.md). İlgili: [13](13-Erisim-ve-Kullanim-Rehberi.md) (erişim, FileSystem), [14](14-VBA-ve-Excel-Etkilesimi.md) (Excel, SaveAs).

**Gezinme:** Önceki: [14-Excel](14-VBA-ve-Excel-Etkilesimi.md) | [Rehber listesi](README.md) | Sonraki: [16-Iyilestirme](16-Iyilestirme-Onerileri.md) →
