Attribute VB_Name = "ParametreYaz"
Option Explicit

' Purpose: Kullanıcıdan parametre adı ve değer alır; Part'ta günceller.
' Assumptions: 3DExperience açık, aktif belge Part; parametre yazılabilir.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States) – sayı formatı.

Sub ParametreYaz()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim sParamAdi As String
    Dim sDegerStr As String
    Dim dDeger As Double

    On Error GoTo HataYakala

    sParamAdi = InputBox("Parametre adı:", "Parametre yaz", "Length.1")
    If Trim(sParamAdi) = "" Then Exit Sub
    sDegerStr = InputBox("Yeni değer:", "Parametre yaz", "100")
    If Trim(sDegerStr) = "" Then Exit Sub

    dDeger = CDbl(sDegerStr)

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "Uygulama yok.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    On Error Resume Next
    Set oParam = oParams.Item(sParamAdi)
    On Error GoTo HataYakala
    If oParam Is Nothing Then MsgBox "Parametre yok: " & sParamAdi: Exit Sub

    oParam.Value = dDeger
    oPart.Update
    MsgBox sParamAdi & " = " & dDeger & " olarak güncellendi."
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
