# 7. Makro Kayıt ve İnceleme

Kodlamaya yeni başlarken **makro kaydı** çok işe yarar: 3DExperience’ta yaptığınız adımlar VBA koduna dönüşür. Bu kodu inceleyerek hangi API’lerin nasıl kullanıldığını öğrenir ve kendi makronuzu yazarken örnek alırsınız.

------------------------------------------------------------

```
  Kayıt öncesi          Kayıt                 Kayıt sonrası
  (hangi iş?)    ──►  Start Recording   ──►  Stop Recording
       │                    │                        │
       │              İşlemleri yap                  │
       │              (parça aç, parametre            │
       │               değiştir, tıkla...)            ▼
       │                                        Kodu bul (Alt+F11)
       │                                        ──►  İncele  ──►  Sadeleştir (Option Explicit, Nothing, tek Update)
       └──────────────────────────────────────────────────────────────────────────────────────────────────►  Kullan / genelleştir
```

------------------------------------------------------------

## Makro kaydı nasıl açılır?

1. 3DExperience’ı açın ve ilgili rolü seçin (ör. Mechanical Designer).  
2. Menüden **Tools** → **Macro** → **Start Recording** (veya **Record Macro**).  
3. Bir makro adı ve isteğe bağlı açıklama girin; kayıt başlar.  
4. Yapmak istediğiniz işlemleri **normal şekilde** yapın: parça açın, bir ölçüyü değiştirin, bir komut tıklayın vb.  
5. **Tools** → **Macro** → **Stop Recording** ile kaydı durdurun.

Kayıt, VBA projesine bir modül ekleyip tüm adımları Sub içinde kod olarak yazar.

------------------------------------------------------------

## Kaydedilen kodu bulma ve açma

1. **Alt+F11** ile VBA editörünü açın.  
2. Sol tarafta **Project Explorer**’da yeni bir modül görünür (ör. Module1, RecordedMacro1).  
3. Modüle çift tıklayın; sağda kaydedilen VBA kodu açılır.

------------------------------------------------------------

## Kayıt çıktısını inceleme

Kayıt genelde şunları üretir:

- **Application / Document erişimi:** `GetObject(, "CATIA.Application")`, `ActiveDocument` vb.  
- **Nesne zincirleri:** Örn. `oDoc.Part.MainBody.Shapes.Item(1)`.  
- **Property atamaları:** Örn. `Parameter.Value = 100`.  
- **Method çağrıları:** Örn. `Part.Update`, `Selection.Add` vb.

Yapmanız gerekenler:

1. **Hangi nesneden başlıyor?** (Application → Document → Part/Product)  
2. **Hangi property/method kullanılmış?** İsimleri not alın; `VBA_API_REFERENCE.md` içinde arayın.  
3. **Sabit değerler:** Sayılar, metinler kayıtta sabit yazılır; bunları ileride değişkene veya parametreye çevirebilirsiniz.

------------------------------------------------------------

## Kayıt kodu nasıl sadeleştirilir?

Kayıt gereksiz adımlar da ekler (seçim temizleme, fazla güncelleme vb.). İnceleyerek:

- Sadece **işinize yarayan satırları** bırakın.  
- Tekrarlayan blokları **döngü** veya **Sub/Function** yapın.  
- Sabit sayı/metinleri **değişken** veya **parametre** yapın.  
- **Option Explicit** ekleyin ve değişkenleri `Dim` ile tanımlayın.

Örnek (kavramsal):

```vba
' Kayıtta: Shapes.Item(1), Shapes.Item(2), ... ayrı ayrı yazılmış olabilir
' Sadeleştirilmiş:
Dim i As Long
For i = 1 To oShapes.Count
    Set oShape = oShapes.Item(i)
    ' ... işlem
Next i
```

------------------------------------------------------------

## Kayıtla öğrenme alışkanlığı

- Yeni bir komut veya pencerede bir işlem kullanacaksanız, önce **kayıt yapıp** nasıl kod üretildiğine bakın.  
- Üretilen sınıf/method isimlerini `VBA_API_REFERENCE.md` veya Help’te arayın; benzer işler için nasıl kullanıldığını görün.  
- Küçük bir **test makrosu** yazıp sadece o birkaç satırı çalıştırarak davranışı doğrulayın.

------------------------------------------------------------

## Dikkat edilecekler

- Kayıt, **o an açık olan belge ve seçime** göre kod üretir. Başka belgede veya farklı seçimle çalıştırırsanız hata alabilirsiniz; bu yüzden kodu genelleştirmek (ActiveDocument, parametreler vb.) önemlidir.  
- Bazı işlemler kayda **yansımaz** veya farklı API ile yapılır; böyle durumda Help veya API referansına bakın.

------------------------------------------------------------

## Kayıt sonrası zorunlu düzenlemeler (Help’ten)

**Help-Automation Development Guidelines** ve **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ**’ne göre kayıt edilmiş makroyu temel alıyorsanız:

1. **Option Explicit** ekleyin; kayıtta oluşan kullanılmayan değişkenleri (ör. `Language = "VBScript"`) **silin**, tanımsız değişken bırakmayın.
2. **Tüm Set atamalarından sonra** ilgili nesne için **Nothing** veya **Count** kontrolü ekleyin; böylece farklı belge/ortamda güvenli çıkış yapılır.
3. **Tek Update** kuralına uyun: Döngü içinde `oPart.Update` çağırmayın; gerekirse işlemler bittikten sonra **bir kez** Update çağırın (performans ve tutarlılık için).
4. **Eski V5 API** kullanımı (Documents.Add, Selection.Search, HybridShapeFactoryOld vb.) 3DExperience’ta yasak veya farklıdır; kayıtta çıkmışsa **Help-Native Apps Automation** ve **VBA_API_REFERENCE.md** ile doğru nesne/service’e geçin.
5. **Language** ve **Release** başlık yorumunu ekleyin; böylece hangi sürümde kaydedildiği belli olur.

------------------------------------------------------------

## Örnek: Kayıt öncesi – Kayıtta oluşabilecek kod (kavramsal)

Kayda başlamadan önce ne yapacağınızı planlayın. Örneğin “parçada Length.1 parametresini 100 yap” işlemi kayıtta kabaca şöyle bir blok üretebilir:

```vba
' KAYIT ÇIKTISI (sadeleştirilmemiş, örnek)
Sub RecordedChangeParameter()
    Dim partDocument1
    Set partDocument1 = GetObject(, "CATIA.Application").ActiveDocument
    Dim part1
    Set part1 = partDocument1.Part
    Dim parameters1
    Set parameters1 = part1.Parameters
    Dim length1
    Set length1 = parameters1.Item("Length.1")
    length1.Value = 100
    part1.Update
End Sub
```

Bu kodu alıp Nothing kontrolleri ve Option Explicit ile sadeleştirirsiniz (aşağıdaki “Örnek: Kayıt sonrası sadeleştirilmiş” bölümüne bakın).

------------------------------------------------------------

## Örnek: Kayıt sonrası sadeleştirilmiş – Parametre değiştirme

Aynı işlemin, zorunlu düzenlemeler uygulanmış hali:

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Aktif parçada Length.1 parametresini verilen değere ayarlar.

Sub ParametreDegistir()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object

    On Error GoTo HataYakala
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil."
        Exit Sub
    End If

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok."
        Exit Sub
    End If

    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then
        MsgBox "Parametreler alınamadı."
        Exit Sub
    End If

    Set oParam = oParams.Item("Length.1")
    If oParam Is Nothing Then
        MsgBox "Length.1 parametresi bulunamadı."
        Exit Sub
    End If

    oParam.Value = 100
    oPart.Update
    MsgBox "Length.1 = 100 olarak güncellendi."
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

Burada: Option Explicit, Language/Release, her Set sonrası Nothing/Count kontrolü, tek Update, On Error GoTo kullanıldı.

------------------------------------------------------------

## Örnek: Kayıttan döngüye – Birden fazla parametre

Kayıtta sadece `Item("Length.1")`, `Item("Length.2")` gibi tek tek satırlar oluşur. Bunları **For** döngüsü ile tekrarlamayın; parametre adlarını bir dizi veya liste ile verip döngüde Item alın:

```vba
Option Explicit

Sub BirdenFazlaParametreGuncelle()
    Dim oApp As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim arrNames As Variant
    Dim i As Long
    Dim dValue As Double

    arrNames = Array("Length.1", "Length.2", "Length.3")
    dValue = 50

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub

    For i = LBound(arrNames) To UBound(arrNames)
        On Error Resume Next
        Set oParam = oParams.Item(arrNames(i))
        On Error GoTo 0
        If Not oParam Is Nothing Then
            oParam.Value = dValue
        End If
    Next i
    oPart.Update
    MsgBox "Parametreler güncellendi."
End Sub
```

Döngü içinde **Update** çağrılmadı; sadece en sonda bir kez `oPart.Update` var.

------------------------------------------------------------

## Örnek: Sabitleri değişkene çevirme

Kayıtta `length1.Value = 100` gibi sabit değer vardı. Bunu kullanıcıdan almak için InputBox ile değişkene çevirin:

```vba
Sub ParametreDegeriniKullanicidanAl()
    Dim sInput As String
    Dim dValue As Double

    sInput = InputBox("Length.1 için yeni değer:", "Parametre", "100")
    If sInput = "" Then Exit Sub

    On Error Resume Next
    dValue = CDbl(sInput)
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "Geçerli bir sayı girin."
        Exit Sub
    End If

    ' Burada oPart, oParam alınıp oParam.Value = dValue ve oPart.Update yapılır
    MsgBox "Atanacak değer: " & dValue
End Sub
```

Bu pattern’i tüm “sabit değer” yerlerinde uygulayabilirsiniz.

------------------------------------------------------------

## Örnek: Kayıtta oluşan gereksiz satırları temizleme

Kayıt bazen şunları ekler: fazla **StartWorkbench**, **Selection.Clear**, **Window.Activate**, tekrarlayan **Update**. Bunları tespit edip sadece iş mantığıyla ilgili satırları bırakın. Örnek temizlik öncesi/sonrası:

```vba
' TEMİZLİK ÖNCESİ (kayıt çıktısı)
' StartWorkbench "PartDesign"
' ...
' part1.Update
' part1.Update
' Selection.Clear

' TEMİZLİK SONRASI: Tek Update, gereksiz Selection.Clear kaldırıldı.
oPart.Update
```

------------------------------------------------------------

## Örnek: Kayıt sırasında yapılacak işlem listesi

Kayda başlamadan önce şunları planlayın: (1) Hangi workbench açık olacak? (2) Hangi belge açık olacak? (3) Hangi tek işlem yapılacak? (4) Kayıt bitince tekrarlayan adımları (ör. aynı menüye tıklama) not alın; kodda bunları döngüye çevirebilirsiniz.

------------------------------------------------------------

## Örnek: Kayıt sonrası değişken tiplendirme

Kayıt `Dim part1` gibi tip belirtmeden yazar; **Option Explicit** ekledikten sonra her değişkeni kullanımına göre tiplendirin: `Dim oPart As Object`, `Dim oParams As Object`, `Dim i As Long` vb. Böylece IntelliSense ve hata kontrolü iyileşir.

------------------------------------------------------------

## Örnek: Kayıtta oluşan uzun nesne zincirini değişkene atama

Kayıt bazen `GetObject(, "CATIA.Application").ActiveDocument.Part.Parameters.Item(1)` gibi uzun zincirler üretir. Okunabilirlik ve Nothing kontrolü için ara değişkenlere bölün:

```vba
Dim oApp As Object
Dim oDoc As Object
Dim oPart As Object
Dim oParams As Object
Set oApp = GetObject(, "CATIA.Application")
If oApp Is Nothing Then Exit Sub
Set oDoc = oApp.ActiveDocument
If oDoc Is Nothing Then Exit Sub
Set oPart = oDoc.GetItem("Part")
If oPart Is Nothing Then Exit Sub
Set oParams = oPart.Parameters
If oParams Is Nothing Then Exit Sub
' Artık oParams.Item(1) güvenle kullanılır
```

------------------------------------------------------------

## Örnek: Kayıtta sabit indeks – 1’den Count’a genelleme

Kayıt sırasında sadece bir öğe seçildiyse kod `Item(1)` veya `Item(2)` gibi sabit indeks üretir. Bunu **genel** hale getirmek için döngü kullanın: `For i = 1 To oKoleksiyon.Count` ve `Item(i)`. Böylece aynı makro farklı sayıda parametre/shape içeren belgelerde de çalışır.

------------------------------------------------------------

## Kontrol listesi

- [ ] Makro kaydını başlatıp durdurabiliyorum  
- [ ] Kaydedilen modülü VBA editöründe bulup açabiliyorum  
- [ ] Kodda Application/Document/Part (veya Product) zincirini tanıyorum  
- [ ] İlgili method/property isimlerini referans dokümanda arayacağımı biliyorum  

------------------------------------------------------------

## Sonraki adım

**8. doküman:** [08-Sik-Kullanilan-APIler.md](08-Sik-Kullanilan-APIler.md) — Parça, geometri ve çizimle ilgili sık kullanılan API’ler ve kısa örnekler. **Help:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md). **Sık hatalar:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

**Gezinme:** Önceki: [06-Nesne-Modeli](06-3DExperience-Nesne-Modeli.md) | [Rehber listesi](README.md) | Sonraki: [08-Sik-Kullanilan-APIler](08-Sik-Kullanilan-APIler.md) →
