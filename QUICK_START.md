# Hızlı başlangıç – İlk 5 dakika

Tek sayfada: 3 adım + tek çalışan makro. Detaylı rehber için [Guidelines/README.md](Guidelines/README.md) ve [Examples/](Examples/README.md) kullanın.

---

## 3 adım

```
  [1] 3DExperience aç (Part belgesi aç)   ──►   [2] VBA Editör: Modül ekle, aşağıdaki kodu yapıştır   ──►   [3] F5 ile çalıştır
```

1. **3DExperience**’ı açın; bir **Part** belgesi açın (veya yeni parça oluşturun).  
2. **Tools → Macro → Edit** (veya **Alt+F11**) ile VBA editörünü açın. Sol tarafta projeye sağ tıklayıp **Insert → Module** ile yeni modül ekleyin. Aşağıdaki kodu modüle yapıştırın.  
3. **F5** (veya Run) ile makroyu çalıştırın. Mesaj kutusunda belge adı görünmelidir.

---

## Tek örnek makro (kopyala–yapıştır)

```vba
Option Explicit

Sub IlkMakro()
    On Error GoTo HataYakala
    Dim oApp As Object
    Dim oDoc As Object

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub

    MsgBox "Belge adı: " & oDoc.Name
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

**Beklenen çıktı:** Part (veya herhangi bir belge) açıksa bir mesaj kutusunda “Belge adı: …” görünür. 3DExperience kapalıysa veya belge yoksa uyarı mesajı çıkar.

---

## Sonraki adımlar

- Daha fazla örnek: [Examples/README.md](Examples/README.md)  
- Parametre okuma/yazma: [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md), [Guidelines/08-Sik-Kullanilan-APIler.md](Guidelines/08-Sik-Kullanilan-APIler.md)  
- Rehberin tamamı: [Guidelines/README.md](Guidelines/README.md)
