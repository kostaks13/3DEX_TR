Attribute VB_Name = "SadecePartKontrol"
Option Explicit

' Örnek: SadecePartKontrol | Rehber: 06, 08 | Zincir/Kontrol
' Purpose: Aktif belgenin Part olup olmadığını kontrol eder; Part değilse uyarır.
' Assumptions: 3DExperience açık; açık belge olabilir veya olmayabilir.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).

Sub SadecePartKontrol()
    On Error GoTo HataYakala
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub

    On Error Resume Next
    Set oPart = oDoc.GetItem("Part")
    On Error GoTo HataYakala
    If oPart Is Nothing Then
        MsgBox "Aktif belge bir parça değil. Lütfen bir .CATPart açın."
        Exit Sub
    End If
    MsgBox "Parça belgesi hazır: " & oDoc.Name
    Exit Sub
HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
