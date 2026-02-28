Attribute VB_Name = "MinMaxParametreDeger"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Purpose: Aktif parçadaki tüm parametreler arasında min ve max sayısal değeri bulup gösterir.

Sub MinMaxParametreDeger()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim dMin As Double
    Dim dMax As Double
    Dim dV As Double
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Or oParams.Count = 0 Then Exit Sub
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
End Sub
