Attribute VB_Name = "FileSystemKlasorListele"
Option Explicit

' Örnek: FileSystemKlasorListele | Rehber: 12 (Servisler, FileSystem) | FileSystem
' ============================================================
' Purpose: Kullanıcının girdiği klasördeki dosyaları listeler
'          (dosya adı ve boyut byte).
' Assumptions: 3DExperience açık; FileSystem API mevcut.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub FileSystemKlasorListele()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oFS As Object
    Dim oFolder As Object
    Dim oFiles As Object
    Dim oFile As Object
    Dim sPath As String
    Dim sOut As String
    Dim iMax As Long
    Dim i As Long

    sPath = InputBox("Listelenecek klasör yolunu girin:", "Klasör", "C:\Temp")
    If Trim(sPath) = "" Then
        MsgBox "İptal edildi."
        Exit Sub
    End If

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil. Önce uygulamayı açın."
        Exit Sub
    End If

    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then
        MsgBox "FileSystem alınamadı."
        Exit Sub
    End If

    On Error Resume Next
    Set oFolder = oFS.GetFolder(sPath)
    On Error GoTo HataYakala
    If oFolder Is Nothing Then
        MsgBox "Klasör bulunamadı: " & sPath
        Exit Sub
    End If

    Set oFiles = oFolder.Files
    If oFiles Is Nothing Then
        MsgBox "Klasörde dosya listesi alınamadı."
        Exit Sub
    End If

    iMax = 30
    sOut = "Dosyalar (" & sPath & ") – ilk " & iMax & ":" & vbCrLf & vbCrLf
    i = 0
    For Each oFile In oFiles
        If Not oFile Is Nothing Then
            sOut = sOut & oFile.Path & "  (" & oFile.Size & " B)" & vbCrLf
            i = i + 1
            If i >= iMax Then Exit For
        End If
    Next oFile
    If i = 0 Then sOut = sOut & "(dosya yok)"

    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
