Attribute VB_Name = "IkiParametreTakas"
Option Explicit

' Purpose: Length.1 ve Length.2 parametrelerinin değerini takas eder.
' Assumptions: 3DExperience açık, aktif belge Part; Length.1 ve Length.2 mevcut.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States) – sayı formatı.

Sub IkiParametreTakas()
    On Error GoTo HataYakala
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oP1 As Object
    Dim oP2 As Object
    Dim d1 As Double
    Dim d2 As Double

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    If oPart Is Nothing Then MsgBox "Bu belge Part değil.": Exit Sub

    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler alınamadı.": Exit Sub
    Set oP1 = oParams.Item("Length.1")
    Set oP2 = oParams.Item("Length.2")
    If oP1 Is Nothing Or oP2 Is Nothing Then MsgBox "Length.1 veya Length.2 bulunamadı.": Exit Sub

    d1 = oP1.Value
    d2 = oP2.Value
    oP1.Value = d2
    oP2.Value = d1
    oPart.Update
    MsgBox "Length.1 ve Length.2 takas edildi."
    Exit Sub
HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
