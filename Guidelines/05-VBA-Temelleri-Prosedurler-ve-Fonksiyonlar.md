# 5. VBA Temelleri – Prosedürler ve Fonksiyonlar

Kodu **parçalara bölmek** için **Sub** (prosedür) ve **Function** (fonksiyon) kullanırız. Böylece kod tekrarı azalır ve makro daha okunabilir olur.

------------------------------------------------------------

## Sub (prosedür)

**Sub** bir işi yapan, ancak **değer döndürmeyen** bloktur. Makroyu çalıştırdığınızda aslında bir Sub çalışır.

```vba
Option Explicit

Sub AnaIslem()
    MsgBox "Başladı"
    ParcaKontrolEt
    MsgBox "Bitti"
End Sub

Sub ParcaKontrolEt()
    ' Burada parça kontrolü yapılacak (ileride dolduracağız)
    MsgBox "Parça kontrolü yapıldı."
End Sub
```

`AnaIslem` çalışınca önce “Başladı”, sonra `ParcaKontrolEt` çağrılır, en sonda “Bitti” görünür.  
3DExperience’ta: “Belgeyi aç”, “Ölçüleri güncelle”, “Rapor yaz” gibi adımları ayrı Sub’lara bölebilirsiniz.

------------------------------------------------------------

## Function (fonksiyon)

**Function** bir **değer hesaplayıp döndürür**. Dönüş tipini ve `Exit Function` / `FunctionAdi = deger` kullanımını bilmek yeterli.

```vba
Function Topla(a As Long, b As Long) As Long
    Topla = a + b
End Function

Sub Test()
    Dim sonuc As Long
    sonuc = Topla(3, 5)
    MsgBox "Toplam: " & sonuc   ' 8
End Sub
```

- **Function Topla(...) As Long** → İki sayı alır, Long döndürür.  
- **Topla = a + b** → Dönüş değerini atar.  
- Çağırırken: `sonuc = Topla(3, 5)`.

3DExperience’ta: “Parça adını getir”, “Ölçü değerini oku”, “Birim dönüştür” gibi işlerde Function kullanılır.

------------------------------------------------------------

## Parametreler (argümanlar)

Hem Sub hem Function parametre alabilir. **ByVal** (kopya) veya **ByRef** (referans); varsayılan ByRef’tir.

```vba
Sub MesajGoster(ByVal metin As String)
    MsgBox metin
End Sub

Function Kare(ByVal x As Double) As Double
    Kare = x * x
End Function

Sub Cagri()
    MesajGoster "Merhaba"
    MsgBox "Kare(4) = " & Kare(4)   ' 16
End Sub
```

3DExperience’ta: parça adı, ölçü adı, dosya yolu gibi bilgileri parametre olarak geçirirsiniz.

------------------------------------------------------------

## Optional ve varsayılan değer

Bazı parametreleri isteğe bağlı yapabilirsiniz:

```vba
Sub Kaydet(Optional ByVal dosyaAdi As String = "Varsayilan.catpart")
    MsgBox "Kaydedilecek: " & dosyaAdi
End Sub

Sub Test()
    Kaydet                    ' Varsayilan.catpart
    Kaydet "OzelParça.catpart" ' OzelParça.catpart
End Sub
```

------------------------------------------------------------

## Özet tablo

| | Sub | Function |
|---|-----|----------|
| Amaç | İş yap (mesaj, dosya, 3DExperience işlemi) | Değer hesapla ve döndür |
| Dönüş | Yok | Var (tip belirtilir) |
| Çağrı | `IslemAdi` veya `Call IslemAdi` | `sonuc = FonksiyonAdi(...)` |

------------------------------------------------------------

## 3DExperience’ta kullanım örneği (iskelet)

```vba
Sub MakroyuCalistir()
    Dim oApp As Object
    Set oApp = GetObject(, "CATIA.Application")  ' İleride anlatılacak
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil."
        Exit Sub
    End If
    
    Dim oDoc As Object
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Aktif belge yok."
        Exit Sub
    End If
    
    ' İşlemleri ayrı Sub'lara bölmek
    ParcaBilgisiniYaz oDoc
End Sub

Sub ParcaBilgisiniYaz(oDoc As Object)
    ' oDoc üzerinden parça bilgisi yazdırma (6. ve 8. dokümanlarda doldurulacak)
    MsgBox "Belge: " & oDoc.Name
End Sub
```

Bu yapıyı 6. dokümanda **nesne modeli** ile somutlaştıracağız.

------------------------------------------------------------

## Örnek: Parametreli Sub – Belge adı yazdırma

Sub’a Object parametresi geçirerek “herhangi bir belge” için aynı işlemi yapabilirsiniz.

```vba
Option Explicit

Sub BelgeBilgisiYaz(oDoc As Object)
    If oDoc Is Nothing Then
        MsgBox "Belge nesnesi yok."
        Exit Sub
    End If
    MsgBox "Ad: " & oDoc.Name & vbCrLf & "Tam yol: " & oDoc.FullName
End Sub

Sub CagiranSub()
    Dim oApp As Object
    Dim oDoc As Object
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oDoc = oApp.ActiveDocument
    BelgeBilgisiYaz oDoc
End Sub
```

------------------------------------------------------------

## Örnek: Function – Birim dönüştürme (mm → inch)

Function, hesaplanan değeri döndürür; böylece aynı hesaplamayı birçok yerde kullanabilirsiniz.

```vba
Function MmToInch(dMm As Double) As Double
    MmToInch = dMm / 25.4
End Function

Function InchToMm(dInch As Double) As Double
    InchToMm = dInch * 25.4
End Function

Sub BirimTest()
    Dim dMm As Double
    dMm = 25.4
    MsgBox dMm & " mm = " & MmToInch(dMm) & " inch"
    MsgBox "1 inch = " & InchToMm(1) & " mm"
End Sub
```

------------------------------------------------------------

## Örnek: Function – Boolean döndüren (geçerli mi?)

“Nesne geçerli mi?” gibi True/False döndüren yardımcı fonksiyon yazmak okunabilirliği artırır.

```vba
Function BelgeGecerliMi(oDoc As Object) As Boolean
    If oDoc Is Nothing Then
        BelgeGecerliMi = False
    Else
        BelgeGecerliMi = True
    End If
End Function

Sub Kullanim()
    Dim oDoc As Object
    Set oDoc = Nothing
    If BelgeGecerliMi(oDoc) Then
        MsgBox "Belge var."
    Else
        MsgBox "Belge yok."
    End If
End Sub
```

------------------------------------------------------------

## Örnek: Call ile Sub çağrısı

Sub’ı çağırırken `Call` kullanmak okunabilirliği artırır.

```vba
Sub Adim1()
    MsgBox "Adım 1"
End Sub

Sub TumAdimlar()
    Call Adim1
    MsgBox "Bitti"
End Sub
```

------------------------------------------------------------

## Prosedür/fonksiyon başlığı (Help’ten)

**Help-Automation Development Guidelines**’a göre her **iç prosedür veya fonksiyon** için bir başlık yorumu bulunmalıdır:

- **Amaç (Purpose)** — Ne yapar?
- **Parametreler** — Anlamları kısaca.
- **Dönüş değeri** — Function ise ne döner?

Örnek:

```vba
' Purpose: Aktif parçanın HybridBodies koleksiyonunu döndürür.
' Parametre: oPart — Part nesnesi.
' Return: HybridBodies koleksiyonu veya Nothing.
Function GetHybridBodies(oPart As Object) As Object
    If oPart Is Nothing Then Set GetHybridBodies = Nothing: Exit Function
    Set GetHybridBodies = oPart.GetItem("HybridBodies")
End Function
```

Bu, kodun bakımını ve yeniden kullanımını kolaylaştırır.

------------------------------------------------------------

## Kontrol listesi

------------------------------------------------------------

## Örnek: Function – String döndüren (parametre adı listesi)

Parametre adlarını virgülle ayrılmış tek metin olarak döndüren yardımcı fonksiyon:

```vba
Function ParametreAdlariniGetir(oParams As Object) As String
    Dim i As Long
    Dim sOut As String
    ParametreAdlariniGetir = ""
    If oParams Is Nothing Then Exit Function
    For i = 1 To oParams.Count
        sOut = sOut & oParams.Item(i).Name & ", "
    Next i
    If Len(sOut) > 0 Then sOut = Left(sOut, Len(sOut) - 2)
    ParametreAdlariniGetir = sOut
End Function
```

------------------------------------------------------------

## Örnek: ByRef – Çağıran tarafın değişkenini güncelleme

ByRef ile Sub içinde değişen parametre, çağıran tarafta da güncellenir:

```vba
Sub Artir(ByRef x As Long)
    x = x + 1
End Sub

Sub TestByRef()
    Dim n As Long
    n = 5
    Artir n
    MsgBox "n = " & n
End Sub
```

------------------------------------------------------------

## Örnek: Optional – İsteğe bağlı parametre

Optional parametre verilmezse varsayılan değer kullanılır:

```vba
Sub MesajOptional(Optional ByVal sMetin As String = "Varsayılan")
    MsgBox sMetin
End Sub
```

------------------------------------------------------------

## Örnek: Public vs Private – Görünürlük

Modül içinde **Public Sub** veya **Private Sub** yazılabilir. **Public** (varsayılan) makro listesinde görünür ve başka modüllerden çağrılabilir. **Private** sadece aynı modül içinden çağrılır; makro listesinde görünmez:

```vba
Public Sub AnaMakro()
    Call YardimciSub
End Sub

Private Sub YardimciSub()
    MsgBox "Sadece bu modülden çağrılır."
End Sub
```

------------------------------------------------------------

## Örnek: Exit Sub / Exit Function – Erken çıkış

Koşul sağlanınca prosedürden veya fonksiyondan hemen çıkmak için **Exit Sub** veya **Exit Function** kullanın:

```vba
Function ParametreVarMi(oPart As Object, sAd As String) As Boolean
    Dim oParams As Object
    ParametreVarMi = False
    If oPart Is Nothing Then Exit Function
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Function
    On Error Resume Next
    If Not oParams.Item(sAd) Is Nothing Then ParametreVarMi = True
End Function
```

------------------------------------------------------------

- [ ] Sub tanımlayıp başka Sub’dan çağırabiliyorum  
- [ ] Function yazıp dönüş değerini kullanabiliyorum  
- [ ] Parametreli Sub/Function yazabiliyorum  
- [ ] Optional parametre kullanımını biliyorum  

------------------------------------------------------------

## Sonraki adım

**6. doküman:** [06-3DExperience-Nesne-Modeli.md](06-3DExperience-Nesne-Modeli.md) — Uygulama, belge, parça hiyerarşisi ve erişim.

**Gezinme:** ← [04-Kosullar-ve-Donguler](04-VBA-Temelleri-Kosullar-ve-Donguler.md) | [Rehber listesi](README.md) | Sonraki: [06-Nesne-Modeli](06-3DExperience-Nesne-Modeli.md) →
