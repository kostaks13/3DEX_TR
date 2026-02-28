Attribute VB_Name = "AktifParcaBilgisi"
Option Explicit

' Örnek: AktifParcaBilgisi | Rehber: 06, 08 | Zincir/Kontrol
' ============================================================
' Purpose: Aktif parça belgesinin adı, tam yolu ve Shapes
'          sayısını gösterir.
' Assumptions: 3DExperience açık, aktif belge Part.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub AktifParcaBilgisi()
    On Error GoTo HataYakala

    Dim oApp As Object
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience (CATIA) çalışmıyor. Önce uygulamayı açın."
        Exit Sub
    End If

    Dim oDoc As Object
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok. Bir parça açın."
        Exit Sub
    End If

    Dim bilgi As String
    bilgi = "Belge adı: " & oDoc.Name & vbCrLf
    bilgi = bilgi & "Tam yol: " & oDoc.FullName & vbCrLf

    Dim oPart As Object
    Set oPart = oDoc.GetItem("Part")
    If Not oPart Is Nothing Then
        Dim oShapes As Object
        On Error Resume Next
        Set oShapes = oPart.Shapes
        If Err.Number = 0 And Not oShapes Is Nothing Then
            bilgi = bilgi & "Shapes sayısı: " & oShapes.Count & vbCrLf
        End If
        On Error GoTo HataYakala
    End If

    MsgBox bilgi
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
