Attribute VB_Name = "LogOrnekMakro"
Option Explicit

' Örnek: LogOrnekMakro | Rehber: 09, 10 | Log
' ============================================================
' Purpose: Log dosyasına başlangıç/bitiş ve hata satırı yazar
'          (log pattern örneği).
' Assumptions: 3DExperience açık; C:\Temp mevcut ve yazılabilir.
' Language: VBA  |  Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States) – tarih/saat.
' ============================================================

Private Const LOG_PATH As String = "C:\Temp\macro_log.txt"

Sub LogOrnekMakro()
    On Error GoTo HataYakala
    LogSatir LOG_PATH, "START " & "ParametreListesi"
    ' ... işlemler ...
    LogSatir LOG_PATH, "END OK"
    MsgBox "Bitti. Log: " & LOG_PATH
    Exit Sub
HataYakala:
    LogSatir LOG_PATH, "END ERR " & Err.Number & " " & Err.Description
    MsgBox "Hata: " & Err.Description
End Sub

Private Sub LogSatir(sDosya As String, sMesaj As String)
    Dim iFile As Integer
    On Error Resume Next
    iFile = FreeFile
    Open sDosya For Append As #iFile
    If Err.Number <> 0 Then Exit Sub
    Print #iFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & sMesaj
    Close #iFile
    On Error GoTo 0
End Sub
