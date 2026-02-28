Attribute VB_Name = "InertiaServiceKutleGoster"
Option Explicit

' Örnek: InertiaServiceKutleGoster | Rehber: 12 (Servisler) | Editor-level servis
' ============================================================
' Purpose: Aktif editördeki Part/Product için InertiaService alır;
'          kütle bilgisi varsa gösterir (API sürüme göre değişir).
' Assumptions: 3DExperience açık; Part veya Product penceresi aktif.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub InertiaServiceKutleGoster()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oEditor As Object
    Dim oInertiaSvc As Object
    Dim sOut As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil. Önce uygulamayı açın."
        Exit Sub
    End If

    Set oEditor = oApp.ActiveEditor
    If oEditor Is Nothing Then
        MsgBox "Aktif editör yok. Bir Part veya Product penceresi açın."
        Exit Sub
    End If

    On Error Resume Next
    Set oInertiaSvc = oEditor.GetService("InertiaService")
    On Error GoTo HataYakala
    If oInertiaSvc Is Nothing Then
        MsgBox "InertiaService alınamadı (Part/Product penceresi aktif olmalı)."
        Exit Sub
    End If

    sOut = "InertiaService alındı." & vbCrLf & vbCrLf
    sOut = sOut & "Kütle okumak için Help'te InertiaService API'ye bakın (GetInertia, Mass vb. sürüme göre değişir)."
    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
