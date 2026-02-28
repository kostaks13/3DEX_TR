Attribute VB_Name = "HybridBodiesListele"
Option Explicit

' Örnek: HybridBodiesListele | Rehber: 12 (Servisler, Part işlemleri) | Part
' ============================================================
' Purpose: Aktif Part'taki HybridBodies ve her birindeki
'          HybridShapes isimlerini listeler.
' Assumptions: 3DExperience açık, aktif belge Part.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States).
' ============================================================

Sub HybridBodiesListele()
    On Error GoTo HataYakala

    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oHBs As Object
    Dim oHB As Object
    Dim oHSs As Object
    Dim oHS As Object
    Dim i As Long
    Dim j As Long
    Dim sOut As String
    Dim iMaxShape As Long

    iMaxShape = 50

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

    Set oHBs = oPart.HybridBodies
    If oHBs Is Nothing Then
        MsgBox "HybridBodies alınamadı (parça boş veya API farklı olabilir)."
        Exit Sub
    End If

    sOut = "HybridBodies ve HybridShapes:" & vbCrLf & vbCrLf
    For i = 1 To oHBs.Count
        Set oHB = oHBs.Item(i)
        If Not oHB Is Nothing Then
            sOut = sOut & "BODY " & i & ": " & oHB.Name & vbCrLf
            Set oHSs = oHB.HybridShapes
            If Not oHSs Is Nothing Then
                For j = 1 To oHSs.Count
                    If j > iMaxShape Then
                        sOut = sOut & "  ... (" & oHSs.Count & " toplam, ilk " & iMaxShape & " gösterildi)" & vbCrLf
                        Exit For
                    End If
                    Set oHS = oHSs.Item(j)
                    If Not oHS Is Nothing Then sOut = sOut & "  - " & oHS.Name & vbCrLf
                Next j
            End If
        End If
    Next i

    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
