Attribute VB_Name = "IkiParametreTakas"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Purpose: Length.1 ve Length.2 parametrelerinin değerini takas eder.

Sub IkiParametreTakas()
    Dim oPart As Object
    Dim oParams As Object
    Dim oP1 As Object
    Dim oP2 As Object
    Dim d1 As Double
    Dim d2 As Double
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    Set oP1 = oParams.Item("Length.1")
    Set oP2 = oParams.Item("Length.2")
    If oP1 Is Nothing Or oP2 Is Nothing Then MsgBox "Parametre bulunamadı.": Exit Sub
    d1 = oP1.Value
    d2 = oP2.Value
    oP1.Value = d2
    oP2.Value = d1
    oPart.Update
    MsgBox "Length.1 ve Length.2 takas edildi."
End Sub
