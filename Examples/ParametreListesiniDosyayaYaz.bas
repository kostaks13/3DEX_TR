Attribute VB_Name = "ParametreListesiniDosyayaYaz"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Purpose: Aktif parçanın parametre listesini CSV formatında dosyaya yazar.
' Varsayılan yol: C:\Temp\parametre_listesi.txt (dağıtımda değiştirilebilir).

Sub ParametreListesiniDosyayaYaz()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim iFile As Integer
    Dim sDosya As String

    sDosya = "C:\Temp\parametre_listesi.txt"
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then MsgBox "Parça yok.": Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    iFile = FreeFile
    Open sDosya For Output As #iFile
    Print #iFile, "Parametre;Değer"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then Print #iFile, oParam.Name & ";" & oParam.Value
    Next i
    Close #iFile
    MsgBox "Liste yazıldı: " & sDosya
End Sub
