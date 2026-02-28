Attribute VB_Name = "GetActivePart_AnaParametreListesi"
Option Explicit

' Purpose: Modüler yapı örneği – GetActivePart ortak fonksiyonu ve tüm parametreleri listeleyen AnaParametreListesi.
' Assumptions: 3DExperience açık, aktif belge Part.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States) – sayı formatı.

Private Function GetActivePart() As Object
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    On Error Resume Next
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Set GetActivePart = Nothing: Exit Function
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Set GetActivePart = Nothing: Exit Function
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set GetActivePart = oPart
End Function

Sub AnaParametreListesi()
    On Error GoTo HataYakala
    Dim oPart As Object
    Dim oParams As Object
    Dim i As Long
    Dim sOut As String

    Set oPart = GetActivePart()
    If oPart Is Nothing Then MsgBox "Aktif parça alınamadı.": Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    sOut = "Parametre sayısı: " & oParams.Count & vbCrLf
    For i = 1 To oParams.Count
        On Error Resume Next
        sOut = sOut & oParams.Item(i).Name & " = " & oParams.Item(i).Value & vbCrLf
        On Error GoTo HataYakala
    Next i
    MsgBox sOut
    Exit Sub
HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
