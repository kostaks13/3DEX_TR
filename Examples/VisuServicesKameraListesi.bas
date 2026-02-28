Attribute VB_Name = "VisuServicesKameraListesi"
Option Explicit

' Örnek: VisuServicesKameraListesi | Rehber: 12 (Servisler) | Session-level
' ============================================================
' Purpose: GetSessionService("VisuServices") ile kamera koleksiyonuna
'          erişir; kamera adlarını listeler.
' Assumptions: 3DExperience açık; VisuServices API mevcut.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub VisuServicesKameraListesi()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oVisu As Object
    Dim oCameras As Object
    Dim oCam As Object
    Dim i As Long
    Dim sOut As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil. Önce uygulamayı açın."
        Exit Sub
    End If

    On Error Resume Next
    Set oVisu = oApp.GetSessionService("VisuServices")
    On Error GoTo HataYakala
    If oVisu Is Nothing Then
        MsgBox "VisuServices alınamadı."
        Exit Sub
    End If

    Set oCameras = oVisu.Cameras
    If oCameras Is Nothing Then
        MsgBox "Cameras koleksiyonu alınamadı."
        Exit Sub
    End If

    sOut = "Kameralar (" & oCameras.Count & "):" & vbCrLf & vbCrLf
    For i = 1 To oCameras.Count
        Set oCam = oCameras.Item(i)
        If Not oCam Is Nothing Then sOut = sOut & i & ": " & oCam.Name & vbCrLf
    Next i

    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
