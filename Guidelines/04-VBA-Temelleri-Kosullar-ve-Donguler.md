# 4. VBA Temelleri – Koşullar ve Döngüler

Programda “eğer şu olursa bunu yap” ve “bir işi N kez tekrarla” demek için **koşul** ve **döngü** yapılarını kullanırız. 3DExperience makrolarında sık karşılaşacaksınız.

------------------------------------------------------------

## Koşul: If / Then / Else

Belirli bir koşul doğruysa bir blok, değilse başka bir blok çalışır.

```vba
Option Explicit

Sub KosulOrnek()
    Dim sayi As Long
    sayi = 5
    
    If sayi > 10 Then
        MsgBox "Sayı 10'dan büyük."
    ElseIf sayi > 0 Then
        MsgBox "Sayı pozitif ama 10'dan küçük."
    Else
        MsgBox "Sayı sıfır veya negatif."
    End If
End Sub
```

- **And** / **Or:** Birden fazla koşulu birleştirir: `If sayi > 0 And sayi < 100 Then`  
- 3DExperience’ta örnek: “Belge açık mı?”, “Parça var mı?” gibi kontrollerde kullanılır.

------------------------------------------------------------

## Select Case (çoklu seçim)

Bir değişkenin değerine göre farklı dallara gidecekseniz **Select Case** okunabilir olur.

```vba
Sub SelectCaseOrnek()
    Dim secim As Long
    secim = 2
    
    Select Case secim
        Case 1
            MsgBox "Parça işlemi"
        Case 2
            MsgBox "Montaj işlemi"
        Case 3
            MsgBox "Çizim işlemi"
        Case Else
            MsgBox "Bilinmeyen seçim"
    End Select
End Sub
```

------------------------------------------------------------

## For döngüsü (belirli sayıda tekrar)

Belirli sayıda tekrar için **For ... Next** kullanın. Örneğin koleksiyondaki her eleman için döngü (Count ile):

```vba
Sub ForOrnek()
    Dim i As Long
    For i = 1 To 5
        MsgBox "Adım: " & i
    Next i
End Sub
```

3DExperience’ta: “Tüm parçaları tara”, “1’den 10’a kadar ölçü güncelle” gibi senaryolarda kullanılır.

------------------------------------------------------------

## For Each (koleksiyonda dolaşma)

Bir koleksiyondaki **her eleman** için döngü yapmak istediğinizde **For Each** idealidir. 3DExperience API’sinde sık görülür.

```vba
Sub ForEachOrnek()
    Dim oKoleksiyon As Object   ' Örnek: parçalar koleksiyonu
    Dim oElem As Object
    
    ' oKoleksiyon = ... (ileride 3DExperience'tan alacağız)
    ' For Each oElem In oKoleksiyon
    '     MsgBox oElem.Name
    ' Next oElem
End Sub
```

Gerçek kullanımda `oKoleksiyon` yerine Part’ın Shapes’i, Product’ın Children’ı vb. gelecek.

------------------------------------------------------------

## Do döngüsü (koşula bağlı tekrar)

“Koşul doğru olduğu sürece tekrarla” için **Do While** / **Do Until** kullanılır.

```vba
Sub DoOrnek()
    Dim sayac As Long
    sayac = 0
    
    Do While sayac < 3
        sayac = sayac + 1
        MsgBox "Sayac: " & sayac
    Loop
End Sub
```

- **Do Until** koşul doğru olana kadar döner: `Do Until sayac >= 3`  
- Sonsuz döngüye düşmemek için koşulun bir yerde sağlanacağından emin olun.

------------------------------------------------------------

## 3DExperience’a uyarlanmış örnek fikirleri

- **If:** `If oDoc Is Nothing Then MsgBox "Belge açık değil"`  
- **For:** `For i = 1 To oShapes.Count`  
- **For Each:** `For Each oShape In oPart.Shapes`  

Bu yapıları 6. ve 8. dokümanlarda nesne modeli ve API örnekleriyle birleştireceğiz.

------------------------------------------------------------

## Örnek: 3DExperience tarzı Nothing kontrolü

Aktif belge veya parça alındıktan sonra mutlaka **Is Nothing** kontrolü yapılır. Bu pattern’i her makroda kullanacaksınız.

```vba
Option Explicit

Sub NothingKontrolOrnek()
    Dim oDoc As Object
    ' Simülasyon: oDoc henüz atanmadı
    Set oDoc = Nothing

    If oDoc Is Nothing Then
        MsgBox "Belge yok. Önce bir parça veya montaj açın."
        Exit Sub
    End If

    ' Buraya sadece oDoc geçerliyse gelinir
    MsgBox "Belge adı: " & oDoc.Name
End Sub
```

Gerçek kodda `Set oDoc = oApp.ActiveDocument` sonrası aynı `If oDoc Is Nothing` kontrolü yapılır.

------------------------------------------------------------

## Örnek: For döngüsü ile 1’den N’e işlem

Parametre indeksi veya şekil indeksi 1’den başlar (çoğu COM koleksiyonunda). For döngüsü ile her birine erişim:

```vba
Sub ForIndeksOrnek()
    Dim i As Long
    Dim sListe As String
    sListe = ""
    For i = 1 To 10
        sListe = sListe & "Öğe " & i & vbCrLf
    Next i
    MsgBox sListe
End Sub
```

3DExperience’ta: `For i = 1 To oShapes.Count` ile tüm şekilleri tarayabilirsiniz.

------------------------------------------------------------

## Örnek: Do While ile “koşul sağlanana kadar” döngü

Sayacı artırıp belirli bir değere ulaşınca çıkış:

```vba
Sub DoWhileOrnek()
    Dim iSayac As Long
    iSayac = 0
    Do While iSayac < 5
        iSayac = iSayac + 1
        MsgBox "Adım " & iSayac
    Loop
End Sub
```

Dikkat: Koşulun bir gün sağlanacağından emin olun; yoksa sonsuz döngü oluşur.

------------------------------------------------------------

## Örnek: And / Or ile bileşik koşul

Birden fazla koşulu birleştirerek “hem bu hem şu” veya “bu veya şu” yazabilirsiniz.

```vba
Sub BilesikKosulOrnek()
    Dim dDeger As Double
    dDeger = 50.5

    If dDeger >= 0 And dDeger <= 100 Then
        MsgBox "Değer 0–100 aralığında."
    End If

    If dDeger < 0 Or dDeger > 1000 Then
        MsgBox "Değer aralık dışında."
    End If
End Sub
```

3DExperience’ta: “Belge açık **ve** parça **ve** read-only değil” gibi kontrollerde kullanılır.

------------------------------------------------------------

## Örnek: Select Case ile belge türüne göre işlem

Aktif belgenin Part, Product veya Drawing olmasına göre farklı mesaj veya işlem (TypeName veya Document Type API’sine göre):

```vba
Sub BelgeTuruneGoreIslem()
    Dim oDoc As Object
    Dim sTip As String
    Set oDoc = GetObject(, "CATIA.Application").ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    ' Sürüme göre: oDoc.Type veya TypeName(oDoc) vb.
    sTip = "Part"   ' Örnek: gerçekte oDoc.DocumentType veya benzeri okuyun
    Select Case sTip
        Case "Part"
            MsgBox "Parça belgesi — parametreler ve shapes kullanılabilir."
        Case "Product"
            MsgBox "Montaj belgesi — Children ve bileşenler kullanılabilir."
        Case "Drawing"
            MsgBox "Çizim belgesi — Sheets ve Views kullanılabilir."
        Case Else
            MsgBox "Desteklenmeyen belge türü."
    End Select
End Sub
```

------------------------------------------------------------

## Örnek: For Each ile koleksiyon dolaşma (kavramsal)

Koleksiyon nesnelerinde For Each kullanımı; 3DExperience’ta Children, Shapes, Parameters vb. için geçerlidir:

```vba
Sub ForEachKoleksiyonOrnek()
    Dim oProduct As Object
    Dim oChild As Object
    Dim sListe As String
    Set oProduct = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Product")
    If oProduct Is Nothing Then Exit Sub
    If oProduct.Children Is Nothing Then Exit Sub
    sListe = ""
    For Each oChild In oProduct.Children
        sListe = sListe & oChild.Name & vbCrLf
    Next oChild
    MsgBox "Bileşenler:" & vbCrLf & sListe
End Sub
```

------------------------------------------------------------

## Örnek: Do Until – “Koşul sağlanana kadar” alternatif

Do While’ın tersi: Koşul **yanlış** olduğu sürece döngü devam eder; koşul doğru olunca çıkılır:

```vba
Sub DoUntilOrnek()
    Dim i As Long
    i = 0
    Do Until i >= 5
        i = i + 1
        MsgBox "Adım " & i
    Loop
End Sub
```

------------------------------------------------------------

## Örnek: İç içe For – Parametre ve şekil indeksleri

Bazen iki koleksiyonu birlikte taramak gerekir (ör. sayfalar ve her sayfadaki görünümler). İç içe For yapısı:

```vba
Sub IcIceForOrnek()
    Dim iSayfa As Long
    Dim iGorus As Long
    Dim iMaxSayfa As Long
    Dim iMaxGorus As Long
    iMaxSayfa = 3
    iMaxGorus = 2
    For iSayfa = 1 To iMaxSayfa
        For iGorus = 1 To iMaxGorus
            Debug.Print "Sayfa " & iSayfa & ", Görünüm " & iGorus
        Next iGorus
    Next iSayfa
End Sub
```

3DExperience’ta: Sheets sayısı kadar dış döngü, her sayfada Views sayısı kadar iç döngü kullanılabilir.

------------------------------------------------------------

## Örnek: Exit For – Döngüden erken çıkış

Belirli bir koşulda döngüyü yarıda bırakmak için Exit For kullanın:

```vba
Sub ExitForOrnek()
    Dim i As Long
    For i = 1 To 100
        If i > 10 Then Exit For
        Debug.Print i
    Next i
End Sub
```

Örnek senaryo: Parametreler arasında “Length.1” adını bulunca döngüden çıkıp o parametreyi güncellemek.

------------------------------------------------------------

## Örnek: Koşulda String karşılaştırma (StrComp, =)

Metin değişkenlerini karşılaştırırken **=** kullanılabilir; büyük/küçük harf duyarlılığı için **StrComp** kullanın:

```vba
Sub StringKarsilastirOrnek()
    Dim sAd As String
    sAd = "Length.1"
    If sAd = "Length.1" Then MsgBox "Eşleşti."
    If StrComp(sAd, "length.1", vbTextCompare) = 0 Then MsgBox "Metin olarak eşit."
End Sub
```

3DExperience’ta parametre adı veya belge adı karşılaştırmalarında bu yapı kullanılır.

------------------------------------------------------------

## Örnek: For Step – İkişer veya geriye sayma

**Step** ile artış miktarını belirleyebilirsiniz; negatif Step ile geriye sayılır:

```vba
Sub ForStepOrnek()
    Dim i As Long
    For i = 0 To 10 Step 2
        Debug.Print i
    Next i
    For i = 5 To 1 Step -1
        Debug.Print i
    Next i
End Sub
```

------------------------------------------------------------

## Örnek: İç içe If – Çoklu koşul

Önce belge var mı, sonra Part mı, sonra Parameters var mı diye iç içe kontrol:

```vba
Sub IcIceIfOrnek()
    Dim oDoc As Object
    Dim oPart As Object
    Set oDoc = GetObject(, "CATIA.Application").ActiveDocument
    If Not oDoc Is Nothing Then
        Set oPart = oDoc.GetItem("Part")
        If Not oPart Is Nothing Then
            If Not oPart.Parameters Is Nothing Then
                MsgBox "Parametre sayısı: " & oPart.Parameters.Count
            End If
        End If
    End If
End Sub
```

------------------------------------------------------------

## Örnek: Boolean değişken ile bayrak (flag)

Koşulu bir kez hesaplayıp sonucu Boolean değişkende tutmak; aynı koşulu tekrar tekrar yazmamak için:

```vba
Sub BayrakOrnek()
    Dim oPart As Object
    Dim bParametreVar As Boolean
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    bParametreVar = (Not oPart Is Nothing) And (Not oPart.Parameters Is Nothing) And (oPart.Parameters.Count > 0)
    If bParametreVar Then
        MsgBox "Parametre işlemi yapılabilir."
    Else
        MsgBox "Parametre yok veya Part yok."
    End If
End Sub
```

------------------------------------------------------------

## Örnek: For döngüsünde Step -1 (geriye sayma)

Listeyi sondan başa doğru taramak için (bazen koleksiyonda silme veya sıra değişince indeks kayması önlenir):

```vba
Sub GeriyeSayOrnek()
    Dim i As Long
    For i = 10 To 1 Step -1
        Debug.Print i
    Next i
End Sub
```

------------------------------------------------------------

## Test edilebilirlik notu (Help’ten)

**Help-Automation Development Guidelines**’a göre otomatik test senaryolarında:

- **InputBox** kullanıyorsanız **varsayılan değer** (üçüncü parametre) verin; test ortamı bu değeri kullanabilir. Örnek: `InputBox("Gövde adı:", "", "NEWBODY")`.
- **MsgBox** yerine hata durumunda **Err.Raise** kullanırsanız, test framework’ü hatayı “hata” olarak tanır. Örnek: `If iShape > cShapes.Count Then Err.Raise 9999, "MyMacro", "Shape numarası çok büyük"`.

Bu sayede makrolarınız hem elle hem otomatik test suite’lerde daha tutarlı çalışır.

------------------------------------------------------------

## Kontrol listesi

- [ ] If / ElseIf / Else ve End If yazabiliyorum  
- [ ] Select Case kullanabiliyorum  
- [ ] For ... Next ile sayılı döngü kuruyorum  
- [ ] For Each ... Next ile koleksiyon döngüsü kuracağımı biliyorum  
- [ ] Do While / Do Until farkını biliyorum  

------------------------------------------------------------

## Sonraki adım

**5. doküman:** [05-VBA-Temelleri-Prosedurler-ve-Fonksiyonlar.md](05-VBA-Temelleri-Prosedurler-ve-Fonksiyonlar.md) — Sub, Function ve parametreler.

**Gezinme:** Önceki: [03-Degiskenler](03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md) | [Rehber listesi](README.md) | Sonraki: [05-Prosedurler](05-VBA-Temelleri-Prosedurler-ve-Fonksiyonlar.md) →
