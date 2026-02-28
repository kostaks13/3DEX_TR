Attribute VB_Name = "FileDialogDosyaAcOku"
Option Explicit

' Örnek: FileDialogDosyaAcOku | Rehber: 15 (Dosya diyalogları) | FileDialog Open
' ============================================================
' Purpose: FileDialog (Aç) ile kullanıcının seçtiği metin dosyasını
'          açar; ilk satırları (en fazla 15) MsgBox ile gösterir.
' Assumptions: 3DExperience açık; Excel yüklü (FileDialog için); seçilen dosya .txt.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Private Const MAX_SATIR As Long = 15

Sub FileDialogDosyaAcOku()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oFS As Object
    Dim oExcel As Object
    Dim oDlg As Object
    Dim oTS As Object
    Dim sSecilen As String
    Dim sOut As String
    Dim sSatir As String
    Dim i As Long

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil. Önce uygulamayı açın."
        Exit Sub
    End If

    Set oExcel = CreateObject("Excel.Application")
    If oExcel Is Nothing Then
        MsgBox "Excel başlatılamadı (FileDialog için gerekli)."
        Exit Sub
    End If

    Set oDlg = oExcel.FileDialog(1)
    If oDlg Is Nothing Then
        oExcel.Quit
        Set oExcel = Nothing
        MsgBox "FileDialog alınamadı."
        Exit Sub
    End If

    oDlg.Title = "Açılacak metin dosyasını seçin"
    oDlg.InitialFileName = "C:\Temp\"
    oDlg.Filters.Clear
    oDlg.Filters.Add "Metin dosyaları (*.txt)", "*.txt"
    oDlg.Filters.Add "Tüm dosyalar (*.*)", "*.*"
    oDlg.FilterIndex = 1
    oDlg.AllowMultiSelect = False

    If oDlg.Show <> -1 Then
        oExcel.Quit
        Set oExcel = Nothing
        MsgBox "İptal edildi."
        Exit Sub
    End If

    sSecilen = oDlg.SelectedItems(1)
    oExcel.Quit
    Set oDlg = Nothing
    Set oExcel = Nothing

    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then
        MsgBox "FileSystem alınamadı."
        Exit Sub
    End If

    If Not oFS.Exists(sSecilen) Then
        MsgBox "Dosya bulunamadı: " & sSecilen
        Exit Sub
    End If

    On Error Resume Next
    Set oTS = oFS.GetFile(sSecilen).OpenAsTextStream(1)
    On Error GoTo HataYakala
    If oTS Is Nothing Then
        MsgBox "Dosya açılamadı (OpenAsTextStream): " & sSecilen
        Exit Sub
    End If

    sOut = "İlk " & MAX_SATIR & " satır (" & sSecilen & "):" & vbCrLf & vbCrLf
    i = 0
    Do While i < MAX_SATIR
        sSatir = oTS.ReadLine
        If oTS.AtEndOfStream And sSatir = "" Then Exit Do
        sOut = sOut & (i + 1) & ": " & sSatir & vbCrLf
        i = i + 1
        If oTS.AtEndOfStream Then Exit Do
    Loop
    oTS.Close

    MsgBox sOut
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then
        On Error Resume Next
        oExcel.Quit
        On Error GoTo HataYakala
    End If
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
