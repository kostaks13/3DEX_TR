Attribute VB_Name = "SadecePartKontrol"
Option Explicit

' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Purpose: Aktif belgenin Part olup olmadığını kontrol eder; Part değilse uyarır.

Sub SadecePartKontrol()
    Dim oDoc As Object
    Dim oPart As Object
    Set oDoc = GetObject(, "CATIA.Application").ActiveDocument
    If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
    On Error Resume Next
    Set oPart = oDoc.GetItem("Part")
    On Error GoTo 0
    If oPart Is Nothing Then
        MsgBox "Aktif belge bir parça değil. Lütfen bir .CATPart açın."
        Exit Sub
    End If
    MsgBox "Parça belgesi hazır: " & oDoc.Name
End Sub
