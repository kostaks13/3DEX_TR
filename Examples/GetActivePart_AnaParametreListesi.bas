Attribute VB_Name = "GetActivePart_AnaParametreListesi"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Modüler yapı örneği: GetActivePart ortak fonksiyon + AnaParametreListesi.

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
        On Error GoTo 0
    Next i
    MsgBox sOut
End Sub
