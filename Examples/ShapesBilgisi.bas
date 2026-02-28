Attribute VB_Name = "ShapesBilgisi"
Option Explicit

' Purpose: Aktif parçadaki Shapes sayısı ve ilk 10 şeklin adını gösterir.
' Assumptions: 3DExperience açık, aktif belge Part.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).

Sub ShapesBilgisi()
    On Error GoTo HataYakala
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oShapes As Object
    Dim oShape As Object
    Dim i As Long
    Dim iMax As Long
    Dim sOut As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    If oPart Is Nothing Then MsgBox "Bu belge Part değil.": Exit Sub
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
