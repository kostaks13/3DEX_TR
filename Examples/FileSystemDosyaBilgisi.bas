Attribute VB_Name = "FileSystemDosyaBilgisi"
Option Explicit

' Örnek: FileSystemDosyaBilgisi | Rehber: 12 (Servisler, FileSystem) | FileSystem
' ============================================================
' Purpose: Kullanıcının girdiği dosya yolunun var olup olmadığını
'          kontrol eder; varsa dosya boyutunu (byte) gösterir.
' Assumptions: 3DExperience açık; FileSystem API mevcut.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub FileSystemDosyaBilgisi()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oFS As Object
    Dim oFile As Object
    Dim sPath As String
    Dim bVar As Boolean
    Dim lSize As Long

    sPath = InputBox("Kontrol edilecek dosya yolunu girin:", "Dosya yolu", "C:\Temp\macro_log.txt")
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
    bVar = oFS.Exists(sPath)
    On Error GoTo HataYakala

    If Not bVar Then
        MsgBox "Dosya bulunamadı: " & sPath
        Exit Sub
    End If

    Set oFile = oFS.GetFile(sPath)
    If oFile Is Nothing Then
        MsgBox "Dosya bilgisi alınamadı: " & sPath
        Exit Sub
    End If

    lSize = oFile.Size
    MsgBox "Dosya: " & sPath & vbCrLf & "Boyut: " & lSize & " byte"
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
