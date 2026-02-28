Attribute VB_Name = "BomListesiChildren"
Option Explicit

' Örnek: BomListesiChildren | Rehber: 06, 08 | Product / BOM
' Purpose: Aktif montaj (Product) belgesindeki alt bileşenleri (Children) listeler.
' Assumptions: 3DExperience açık, aktif belge Product (.CATProduct).
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' Not: API adları (Product, Children) sürüme göre değişebilir; makro kaydı ile doğrulayın.

Sub BomListesiChildren()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oDoc As Object
    Dim oProduct As Object
    Dim oChildren As Object
    Dim i As Long
    Dim oChild As Object
    Dim sAd As String
    Dim sOut As String
    Dim iMax As Long

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub

    Set oProduct = oDoc.GetItem("Product")
    If oProduct Is Nothing Then
        MsgBox "Aktif belge bir montaj değil. Lütfen bir Product açın."
        Exit Sub
    End If

    Set oChildren = oProduct.Children
    If oChildren Is Nothing Then MsgBox "Children alınamadı (API sürümü).": Exit Sub

    If oChildren.Count = 0 Then
        MsgBox "Alt bileşen yok."
        Exit Sub
    End If

    sOut = "Alt bileşen sayısı: " & oChildren.Count & vbCrLf & vbCrLf
    iMax = oChildren.Count
    If iMax > 20 Then iMax = 20

    For i = 1 To iMax
        On Error Resume Next
        Set oChild = oChildren.Item(i)
        If Not oChild Is Nothing Then
            sAd = oChild.Name
            If Len(sAd) = 0 Then sAd = "(isim yok)"
            sOut = sOut & i & ". " & sAd & vbCrLf
        End If
        On Error GoTo HataYakala
    Next i

    If oChildren.Count > 20 Then sOut = sOut & vbCrLf & "... (ilk 20 gösterildi)"

    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
