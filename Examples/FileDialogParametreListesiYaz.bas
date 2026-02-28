Attribute VB_Name = "FileDialogParametreListesiYaz"
Option Explicit

' Örnek: FileDialogParametreListesiYaz | Rehber: 15 (Dosya diyalogları) | FileDialog
' ============================================================
' Purpose: Kullanıcıya "Kaydet" diyaloğu (FileDialog) gösterir;
'          seçilen yere parametre listesini (Parametre;Değer) yazar.
' Assumptions: 3DExperience açık, aktif belge Part; Excel yüklü (FileDialog için).
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub FileDialogParametreListesiYaz()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim oExcel As Object
    Dim oDlg As Object
    Dim sHedefYol As String
    Dim i As Long
    Dim iFile As Integer

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil. Önce uygulamayı açın."
        Exit Sub
    End If

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok. Bir parça açın."
        Exit Sub
    End If

    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    If oPart Is Nothing Then
        MsgBox "Bu belge Part değil. Bir parça belgesi açın."
        Exit Sub
    End If

    Set oParams = oPart.Parameters
    If oParams Is Nothing Then
        MsgBox "Parametreler alınamadı."
        Exit Sub
    End If

    If oParams.Count = 0 Then
        MsgBox "Parçada parametre yok."
        Exit Sub
    End If

    Set oExcel = CreateObject("Excel.Application")
    If oExcel Is Nothing Then
        MsgBox "Excel başlatılamadı (FileDialog için gerekli). Excel yüklü mü?"
        Exit Sub
    End If

    Set oDlg = oExcel.FileDialog(2)
    If oDlg Is Nothing Then
        oExcel.Quit
        Set oExcel = Nothing
        MsgBox "FileDialog alınamadı."
        Exit Sub
    End If

    oDlg.Title = "Parametre listesini kaydedeceğiniz yeri ve dosya adını seçin"
    oDlg.InitialFileName = "C:\Temp\parametre_listesi.txt"
    oDlg.Filters.Clear
    oDlg.Filters.Add "Metin dosyaları (*.txt)", "*.txt"
    oDlg.Filters.Add "CSV (*.csv)", "*.csv"
    oDlg.FilterIndex = 1
    oDlg.AllowMultiSelect = False

    If oDlg.Show <> -1 Then
        oExcel.Quit
        Set oExcel = Nothing
        MsgBox "İptal edildi."
        Exit Sub
    End If

    sHedefYol = oDlg.SelectedItems(1)
    oExcel.Quit
    Set oDlg = Nothing
    Set oExcel = Nothing

    iFile = FreeFile
    Open sHedefYol For Output As #iFile
    Print #iFile, "Parametre;Değer"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            Print #iFile, oParam.Name & ";" & oParam.Value
        End If
    Next i
    Close #iFile

    MsgBox "Parametre listesi yazıldı: " & sHedefYol
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then
        On Error Resume Next
        oExcel.Quit
        On Error GoTo HataYakala
    End If
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
