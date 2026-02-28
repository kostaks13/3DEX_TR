Attribute VB_Name = "DrawingSayfaVeGorusumSayisi"
Option Explicit

' Örnek: DrawingSayfaVeGorusumSayisi | Rehber: 06, 08 | Drawing
' Purpose: Aktif çizim (Drawing) belgesinde sayfa sayısı ve ilk sayfadaki görünüm sayısını gösterir.
' Assumptions: 3DExperience açık, aktif belge Drawing (.CATDrawing).
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' Not: API adları (DrawingRoot, Sheets, Views) sürüme göre değişebilir; makro kaydı ile doğrulayın.

Sub DrawingSayfaVeGorusumSayisi()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oDoc As Object
    Dim oDraw As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oViews As Object
    Dim sOut As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub

    Set oDraw = oDoc.GetItem("DrawingRoot")
    If oDraw Is Nothing Then
        MsgBox "Aktif belge bir çizim değil. Lütfen bir Drawing açın."
        Exit Sub
    End If

    On Error Resume Next
    Set oSheets = oDraw.Sheets
    On Error GoTo HataYakala
    If oSheets Is Nothing Then MsgBox "Sheets alınamadı (API sürümü).": Exit Sub

    sOut = "Sayfa sayısı: " & oSheets.Count & vbCrLf

    If oSheets.Count >= 1 Then
        Set oSheet = oSheets.Item(1)
        If Not oSheet Is Nothing Then
            Set oViews = oSheet.Views
            If Not oViews Is Nothing Then
                sOut = sOut & "İlk sayfadaki görünüm sayısı: " & oViews.Count
            Else
                sOut = sOut & "İlk sayfada Views alınamadı."
            End If
        End If
    End If

    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
