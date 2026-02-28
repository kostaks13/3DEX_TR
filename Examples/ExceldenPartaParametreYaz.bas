Attribute VB_Name = "ExceldenPartaParametreYaz"
Option Explicit

' Örnek: ExceldenPartaParametreYaz | Rehber: 14 (VBA-Excel) | Excel → Part
' ============================================================
' Purpose: Excel dosyasındaki A sütunu (parametre adı) ve B sütunu (değer)
'          ile Part parametrelerini günceller. A2'den itibaren satır satır okur.
' Assumptions: 3DExperience açık, aktif belge Part; Excel dosyası mevcut.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Private Const EXCEL_YOL As String = "C:\Temp\ParametreGiris.xlsx"

Sub ExceldenPartaParametreYaz()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim oExcel As Object
    Dim oWb As Object
    Dim oWs As Object
    Dim i As Long
    Dim sAd As String
    Dim dDeger As Double

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil. Önce uygulamayı açın."
        Exit Sub
    End If

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok. Bir parça açın."
        Exit Sub
    End If

    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    If oPart Is Nothing Then
        MsgBox "Bu belge Part değil. Bir parça belgesi açın."
        Exit Sub
    End If

    Set oParams = oPart.Parameters
    If oParams Is Nothing Then
        MsgBox "Parametreler alınamadı."
        Exit Sub
    End If

    On Error Resume Next
    Set oWb = GetObject(EXCEL_YOL)
    On Error GoTo HataYakala
    If oWb Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
        Set oWb = oExcel.Workbooks.Open(EXCEL_YOL)
    Else
        Set oExcel = oWb.Application
    End If
    If oWb Is Nothing Then
        MsgBox "Excel dosyası açılamadı: " & EXCEL_YOL
        If Not oExcel Is Nothing Then oExcel.Quit
        Exit Sub
    End If

    Set oWs = oWb.Worksheets.Item(1)
    i = 2
    Do While True
        sAd = Trim(oWs.Cells(i, 1).Value & "")
        If sAd = "" Then Exit Do
        dDeger = oWs.Cells(i, 2).Value
        Set oParam = oParams.Item(sAd)
        If Not oParam Is Nothing Then oParam.Value = dDeger
        i = i + 1
    Loop

    oPart.Update
    oWb.Close SaveChanges:=False
    If oExcel.Workbooks.Count = 0 Then oExcel.Quit
    Set oExcel = Nothing

    MsgBox "Excel'den okunan parametreler Part'a yazıldı; Update çağrıldı."
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then
        On Error Resume Next
        oExcel.Quit
        On Error GoTo HataYakala
    End If
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
