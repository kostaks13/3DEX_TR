Attribute VB_Name = "MinMaxParametreDeger"
Option Explicit

' Purpose: Aktif parçadaki tüm parametreler arasında min ve max sayısal değeri bulup gösterir.
' Assumptions: 3DExperience açık, aktif belge Part; en az bir parametre var.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States) – sayı formatı.

Sub MinMaxParametreDeger()
    On Error GoTo HataYakala
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim dMin As Double
    Dim dMax As Double
    Dim dV As Double

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    If oPart Is Nothing Then MsgBox "Bu belge Part değil.": Exit Sub

    Set oParams = oPart.Parameters
    If oParams Is Nothing Or oParams.Count = 0 Then MsgBox "Parametre yok.": Exit Sub

    dMin = oParams.Item(1).Value
    dMax = dMin
    For i = 2 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            dV = oParam.Value
            If dV < dMin Then dMin = dV
            If dV > dMax Then dMax = dV
        End If
    Next i
    MsgBox "Min: " & dMin & vbCrLf & "Max: " & dMax
    Exit Sub
HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
