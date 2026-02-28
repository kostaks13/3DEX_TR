Attribute VB_Name = "ParametreOkuVeGoster"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Purpose: Aktif parçada belirtilen parametrenin değerini okur ve gösterir.

Sub ParametreOkuVeGoster()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim sParamAdi As String
    Dim dDeger As Double

    On Error GoTo HataYakala

    sParamAdi = InputBox("Parametre adı (örn. Length.1):", "Parametre oku", "Length.1")
    If Trim(sParamAdi) = "" Then Exit Sub

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler alınamadı.": Exit Sub

    On Error Resume Next
    Set oParam = oParams.Item(sParamAdi)
    On Error GoTo HataYakala
    If oParam Is Nothing Then
        MsgBox "Parametre bulunamadı: " & sParamAdi
        Exit Sub
    End If

    dDeger = oParam.Value
    MsgBox sParamAdi & " = " & dDeger
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
