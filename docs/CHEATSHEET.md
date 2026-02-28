# Hızlı referans (Cheat Sheet)

```
┌─────────────────────────────────────────────────────────────────┐
│  Zincir · Parametre · Hata yakalama · Döngü · Kısa terimler      │
└─────────────────────────────────────────────────────────────────┘
```

Tek sayfada **sık kullanılan kalıplar**. Detay için [Guidelines/README.md](../Guidelines/README.md) ve [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md).

**Bu sayfada:** [Uygulama → Part zinciri](#uygulama--part-zinciri) · [Parametre](#parametre-okuma--yazma) · [Hata yakalama](#hata-yakalama-özet) · [Döngü](#koleksiyon-döngüsü-parametreler) · [Kısa terimler](#kısa-terimler)

---

## Uygulama → Part zinciri

```vba
Dim oApp As Object, oDoc As Object, oPart As Object
Set oApp = GetObject(, "CATIA.Application")
If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
Set oDoc = oApp.ActiveDocument
If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub
Set oPart = oDoc.GetItem("Part")
If oPart Is Nothing Then MsgBox "Part bulunamadı.": Exit Sub
```

---

## Parametre okuma / yazma

```vba
' Okuma
Dim val As Double
val = oPart.Parameters.Item("Length.1").Value

' Yazma (sonunda tek Update)
oPart.Parameters.Item("Length.1").Value = 150
oPart.Update
```

---

## Hata yakalama (özet)

```vba
Sub Ornek()
    On Error GoTo HataYakala
    ' ... kod ...
    Exit Sub
HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

---

## Koleksiyon döngüsü (parametreler)

```vba
Dim i As Long, oParam As Object
For i = 1 To oPart.Parameters.Count
    Set oParam = oPart.Parameters.Item(i)
    ' oParam.Name, oParam.Value kullan
Next i
```

---

## Kısa terimler

| Terim | Anlamı |
|-------|--------|
| **Nothing** | Nesne atanmamış; `If oDoc Is Nothing Then` ile kontrol et. |
| **Update** | Değişiklikleri uygula; döngü içinde değil, **bir kez** sonunda çağır. |
| **GetItem("Part")** | Aktif belgeden Part nesnesi; Part değilse Nothing döner. |

Tam liste: [GLOSSARY.md](GLOSSARY.md).

---

**Gezinme:** [Ana sayfa](../README.md) · [Docs](README.md) · [Rehber](../Guidelines/README.md) · [Örnek makrolar](../Examples/README.md) · [VBA_API_REFERENCE](VBA_API_REFERENCE.md)
