Attribute VB_Name = "ParametreListesiniDosyayaYaz"
Option Explicit

' ============================================================
' Purpose: Aktif parçanın parametre listesini CSV formatında
'          (Parametre;Değer) dosyaya yazar.
' Assumptions: 3DExperience açık, aktif belge Part, çıktı
'              klasörü mevcut/yazılabilir.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States) – sayı formatı.
' ============================================================

Private Const OUT_PATH As String = "C:\Temp\parametre_listesi.txt"

Sub ParametreListesiniDosyayaYaz()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim iFile As Integer

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

    iFile = FreeFile
    Open OUT_PATH For Output As #iFile
    Print #iFile, "Parametre;Değer"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            Print #iFile, oParam.Name & ";" & oParam.Value
        End If
    Next i
    Close #iFile

    MsgBox "Liste yazıldı: " & OUT_PATH
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
