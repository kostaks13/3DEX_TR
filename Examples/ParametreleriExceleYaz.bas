Attribute VB_Name = "ParametreleriExceleYaz"
Option Explicit

' Örnek: ParametreleriExceleYaz | Rehber: 14 (VBA-Excel) | Excel
' ============================================================
' Purpose: Aktif parçanın parametre listesini Excel çalışma
'          kitabına yazar (A: Parametre adı, B: Değer); dosyayı kaydeder.
' Assumptions: 3DExperience açık, aktif belge Part; Excel yüklü.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Private Const EXCEL_SAVE_PATH As String = "C:\Temp\ParametreListe.xlsx"

Sub ParametreleriExceleYaz()
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

    If oParams.Count = 0 Then
        MsgBox "Parçada parametre yok."
        Exit Sub
    End If

    Set oExcel = CreateObject("Excel.Application")
    If oExcel Is Nothing Then
        MsgBox "Excel başlatılamadı. Excel yüklü mü?"
        Exit Sub
    End If

    oExcel.Visible = True
    oExcel.DisplayAlerts = False
    Set oWb = oExcel.Workbooks.Add
    Set oWs = oWb.Worksheets.Item(1)

    oWs.Cells(1, 1).Value = "Parametre"
    oWs.Cells(1, 2).Value = "Değer"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            oWs.Cells(i + 1, 1).Value = oParam.Name
            oWs.Cells(i + 1, 2).Value = oParam.Value
        End If
    Next i

    oWb.SaveAs EXCEL_SAVE_PATH
    oWb.Close SaveChanges:=False
    oExcel.Quit
    Set oWs = Nothing
    Set oWb = Nothing
    Set oExcel = Nothing

    MsgBox "Parametre listesi Excel'e yazıldı: " & EXCEL_SAVE_PATH
    Exit Sub

HataYakala:
    If Not oExcel Is Nothing Then
        On Error Resume Next
        oExcel.Quit
        On Error GoTo HataYakala
    End If
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
