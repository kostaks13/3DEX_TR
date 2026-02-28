Attribute VB_Name = "FileSystemKlasorKontrol"
Option Explicit

' Örnek: FileSystemKlasorKontrol | Rehber: 12 | FileSystem
' Purpose: FileSystem ile belirtilen klasörün var olup olmadığını kontrol eder (veya listeler).
' Assumptions: 3DExperience açık; FileSystem API sürümünüzde mevcutsa çalışır.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' Not: FileSystem, CreateFolder, FileExists vb. sürüme göre değişir; Help ve makro kaydı ile doğrulayın.

Sub FileSystemKlasorKontrol()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oFS As Object
    Dim sYol As String
    Dim bVar As Boolean

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub

    sYol = InputBox("Kontrol edilecek klasör yolunu girin:", "Klasör kontrolü", "C:\Temp")
    If Len(Trim(sYol)) = 0 Then Exit Sub

    On Error Resume Next
    Set oFS = oApp.FileSystem
    If oFS Is Nothing Then
        MsgBox "FileSystem bu sürümde kullanılamıyor. Help'te FileSystem/FileExists arayın."
        Exit Sub
    End If

    bVar = oFS.FolderExists(sYol)
    On Error GoTo HataYakala

    If bVar Then
        MsgBox "Klasör mevcut: " & sYol
    Else
        MsgBox "Klasör bulunamadı veya erişilemiyor: " & sYol
    End If
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description & vbCrLf & "FileSystem API adı sürüme göre farklı olabilir (FolderExists / CreateFolder vb.)."
End Sub
