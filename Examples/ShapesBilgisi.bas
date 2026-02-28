Attribute VB_Name = "ShapesBilgisi"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Purpose: Aktif parçadaki Shapes sayısı ve ilk 10 şeklin adını gösterir.

Sub ShapesBilgisi()
    Dim oApp As Object
    Dim oPart As Object
    Dim oShapes As Object
    Dim oShape As Object
    Dim i As Long
    Dim iMax As Long
    Dim sOut As String

    On Error GoTo HataYakala
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oShapes = oPart.Shapes
    If oShapes Is Nothing Then MsgBox "Shapes yok.": Exit Sub

    sOut = "Shapes sayısı: " & oShapes.Count & vbCrLf
    iMax = oShapes.Count
    If iMax > 10 Then iMax = 10
    For i = 1 To iMax
        Set oShape = oShapes.Item(i)
        If Not oShape Is Nothing Then sOut = sOut & i & ": " & oShape.Name & vbCrLf
    Next i
    If oShapes.Count > 10 Then sOut = sOut & "..."
    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
