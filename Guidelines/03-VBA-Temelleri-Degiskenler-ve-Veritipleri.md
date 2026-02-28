# 3. VBA Temelleri – Değişkenler ve Veri Tipleri

```
  Konu: Dim, Set, veri tipleri, Option Explicit, önekler (b, d, s, i, o, c), Const
```

Kod yazarken bilgiyi geçici olarak saklamak için **değişken** kullanırız. Bu dokümanda değişken tanımlama ve veri tiplerini 3DExperience VBA bağlamında öğreneceksiniz.

------------------------------------------------------------

## Değişken nedir?

Değişken, adı verilen bir kutudur; içine sayı, metin, tarih veya nesne referansı koyarız. Örnek:

```vba
Sub DegiskenOrnek()
    Dim ad As String
    ad = "Parça_01"
    
    Dim sayi As Long
    sayi = 10
    
    MsgBox "Ad: " & ad & ", Sayı: " & sayi
End Sub
```

- `Dim ad As String` → “ad” adında bir metin (string) değişkeni tanımlar.  
- `ad = "Parça_01"` → Bu değişkene değer atar.  
- `&` → Metinleri birleştirir.

------------------------------------------------------------

## Sık kullanılan veri tipleri

| Tip | Açıklama | Örnek |
|-----|----------|--------|
| **String** | Metin | `"Parça_01"`, `"C:\Dosyalar"` |
| **Long** | Uzun tamsayı | `1`, `1000` |
| **Integer** | Kısa tamsayı | Küçük sayılar |
| **Double** | Ondalıklı sayı | `3.14`, `0.5` |
| **Boolean** | True / False | Koşul sonuçları |
| **Object** | Nesne referansı | 3DExperience parça, çizim vb. |
| **Variant** | Her türlü değer | Tip belirtmek istemediğinizde (dikkatli kullanın) |

3DExperience’ta parça, montaj, çizim gibi şeylere erişirken çoğu zaman **Object** kullanırız; `Set` ile atama yaparız (aşağıda örnek var).

------------------------------------------------------------

## Option Explicit kullanın

Her modülün **en üstüne** şunu yazın:

```vba
Option Explicit
```

Böylece yazım hatası yaptığınızda (örneğin `Parca` yerine `Parca1` yazarsanız) VBA sizi uyarır; bu da hatayı azaltır. Yeni modül açtığınızda **Tools** → **Options** → **Editor** içinde **Require Variable Declaration** işaretli olsa da, manuel yazmak iyi alışkanlıktır.

------------------------------------------------------------

## Nesne değişkeni ve Set

Sayı veya metin atarken `=` kullanırız. **Nesne** (Object) atarken mutlaka **Set** kullanırız:

```vba
Option Explicit

Sub NesneOrnek()
    Dim oParca As Object   ' 3DExperience parça nesnesi (ileride göreceğiz)
    ' oParca = ...  ' YANLIŞ: Object için = kullanılmaz
    ' Set oParca = ...  ' DOĞRU: Set ile atanır
End Sub
```

3DExperience’tan gelen her “şey” (Part, Product, Drawing vb.) bir nesnedir; ilerideki dokümanlarda hep `Set` ile atayacağız.

------------------------------------------------------------

## Değişken isimlendirme (kısa özet)

- Harf veya alt çizgi ile başlayın; sonra harf, rakam, alt çizgi.  
- Anlamlı isim verin: `parcaAdi`, `olcuDegeri`, `sayac`.  
- 3DExperience nesneleri için birçok örnekte `o` öneki kullanılır: `oParca`, `oDoc`, `oApp`.

------------------------------------------------------------

## Resmi isimlendirme önekleri (Help-Automation Development Guidelines)

Dassault Systèmes **Automation Development Guidelines** dokümanında VBA/VBScript için **tip önekleri** önerilir. Değişken adının başına tek harf ekleyerek tipi anında tanırsınız:

| Önek | Tip | Örnek |
|------|-----|--------|
| **b** | Boolean | bIsUpToDate, bReadOnly |
| **d** | Double | dLength, dAngle |
| **s** | String | sName, sFilePath |
| **i** | Integer | iNumberOfElements, iIndex |
| **o** | Object | oPart, oDoc, oSketch1 |
| **c** | Collection | cDrwViews, cShapes |

Sabitler için: Tamamı **büyük harf**, bileşenler **alt çizgi** ile ayrılır. Örnek: `MAX_VALUE`, `LOG_PATH`, `DEFAULT_BODY_NAME`.

------------------------------------------------------------

## Örnek: Basit hesaplama

```vba
Option Explicit

Sub Hesapla()
    Dim uzunluk As Double
    Dim genislik As Double
    Dim alan As Double
    
    uzunluk = 100.5
    genislik = 50.25
    alan = uzunluk * genislik
    
    MsgBox "Alan: " & alan
End Sub
```

Burada sadece sayısal değişkenler kullanılıyor; 3DExperience’tan henüz veri almıyoruz. Bunu bir sonraki aşamalarda nesne modeli ile birleştireceğiz.

------------------------------------------------------------

## Örnek: Resmi öneklerle 3DExperience tarzı değişkenler

Aşağıdaki örnekte Help’teki önek kuralı (b, d, s, i, o, c) kullanılıyor. Henüz 3DExperience nesnesi almadığımız için nesneleri `Nothing` veya yorum satırıyla bırakıyoruz.

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x

Sub OrnekOnekliDegiskenler()
    Dim bReadOnly As Boolean
    Dim dLength As Double
    Dim sPartName As String
    Dim iCount As Long
    Dim oPart As Object
    Dim cShapes As Object

    bReadOnly = False
    dLength = 100.5
    sPartName = "Parça_01"
    iCount = 0
    ' oPart ve cShapes ileride Set ile atanacak
    ' Set oPart = ...
    ' Set cShapes = oPart.Shapes

    MsgBox "Parça: " & sPartName & vbCrLf & _
           "Uzunluk: " & dLength & vbCrLf & _
           "ReadOnly: " & bReadOnly
End Sub
```

------------------------------------------------------------

## Örnek: Sabit (Const) kullanımı

Tekrarlayan sayı veya metni **Const** ile tanımlayın; değişmesi gerektiğinde tek yerden güncellersiniz. Sabit isimleri büyük harf ve alt çizgi ile yazın (Help kuralı).

```vba
Option Explicit

Sub SabitOrnek()
    Const MAX_ITERATION As Long = 100
    Const DEFAULT_TOLERANCE As Double = 0.001
    Const LOG_PATH As String = "C:\Temp\macro_log.txt"

    Dim i As Long
    For i = 1 To MAX_ITERATION
        ' ... işlem
    Next i
    MsgBox "Tolerans: " & DEFAULT_TOLERANCE & ", Log: " & LOG_PATH
End Sub
```

------------------------------------------------------------

## Örnek: Variant (dikkatli kullanın)

**Variant** her türlü değeri alabilir; ancak tip güvenliği azalır ve hata riski artar. Mümkünse her zaman belirli tip (String, Long, Double, Object) kullanın. Sadece gerçekten “bazen sayı bazen metin” gelen durumlarda Variant düşünün.

```vba
Sub VariantOrnek()
    Dim v As Variant
    v = 10
    MsgBox "v = " & v
    v = "şimdi metin"
    MsgBox "v = " & v
End Sub
```

------------------------------------------------------------

## Örnek: Tarih ve saat (Date tipi)

VBA’da **Date** tipi hem tarih hem saat tutar. `Now` o anki tarih/saati verir.

```vba
Sub TarihOrnek()
    Dim dtSimdi As Date
    dtSimdi = Now
    MsgBox "Şimdi: " & Format(dtSimdi, "yyyy-mm-dd hh:nn:ss")
End Sub
```

Log dosyası veya rapor başlığında tarih/saat yazarken bu yapı kullanılır.

------------------------------------------------------------

## Örnek: Birden fazla değişkeni tek satırda tanımlama

Aynı tipte birkaç değişkeni tek **Dim** satırında tanımlayabilirsiniz; her biri virgülle ayrılır:

```vba
Sub CokluDimOrnek()
    Dim i As Long, j As Long, k As Long
    Dim sAd As String, sSoyad As String
    i = 1
    j = 2
    k = i + j
    sAd = "Parça"
    sSoyad = "01"
    MsgBox sAd & sSoyad & " - Toplam: " & k
End Sub
```

Dikkat: `Dim i, j As Long` yazarsanız sadece `j` Long olur, `i` Variant kalır; bu yüzden her değişkene açıkça tip yazmak daha güvenlidir.

------------------------------------------------------------

## Örnek: Metin birleştirme (&) ve satır sonu (vbCrLf)

VBA’da metinleri **&** ile birleştiririz; satır sonu için **vbCrLf** (veya **vbNewLine**) kullanılır:

```vba
Sub MetinBirlestirOrnek()
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    s1 = "Satır 1"
    s2 = "Satır 2"
    s3 = s1 & vbCrLf & s2 & vbCrLf & "Satır 3"
    MsgBox s3
End Sub
```

Log veya rapor metninde çok satırlı çıktı oluştururken bu yapı kullanılır.

------------------------------------------------------------

## Örnek: Sayıyı metne çevirme (CStr, Format)

MsgBox veya dosyaya yazarken sayıyı metne çevirmek için **CStr** veya **Format** kullanın:

```vba
Sub SayiyiMetinOrnek()
    Dim dVal As Double
    dVal = 3.14159
    MsgBox "Değer: " & CStr(dVal)
    MsgBox "Format: " & Format(dVal, "0.00")
End Sub
```

Format ile ondalık basamak sayısı veya tarih/saat biçimi belirlenir.

------------------------------------------------------------

## Örnek: Enum (sabit listesi) – Okunabilir kod

VBA’da **Enum** ile sayısal sabitlere anlamlı isim verilir; kod okunabilirliği artar:

```vba
Public Enum BelgeTuru
    PartDoc = 1
    ProductDoc = 2
    DrawingDoc = 3
End Enum

Sub EnumOrnek()
    Dim t As BelgeTuru
    t = PartDoc
    If t = PartDoc Then MsgBox "Part belgesi."
End Sub
```

3DExperience’ta belge türü kontrolünde sayı yerine böyle sabitler kullanılabilir (API’nin DocumentType değerleri ile eşleştirilir).

------------------------------------------------------------

## Örnek: Type (kullanıcı tanımlı tip) – İlgili verileri gruplama

Birkaç alanı tek yapıda toplamak için **Type** kullanılır (ileri seviye; basit makrolarda gerekmez):

```vba
Type ParametreBilgi
    sAd As String
    dDeger As Double
End Type

Sub TypeOrnek()
    Dim p As ParametreBilgi
    p.sAd = "Length.1"
    p.dDeger = 100.5
    MsgBox p.sAd & " = " & p.dDeger
End Sub
```

------------------------------------------------------------

## Örnek: Empty ve Null (Variant durumları)

**Variant** bazen **Empty** (henüz atanmamış) veya **Null** (veri yok) değer alır. Kontrol için **IsEmpty** ve **IsNull** kullanılır; 3DExperience API’sinden dönen bazı değerler Null olabilir, bu yüzden Optional parametre veya veritabanı alanlarıyla uğraşırken bu kontroller işe yarar.

------------------------------------------------------------

## Kontrol listesi

- [ ] `Dim degiskenAdi As Tip` ile değişken tanımlayabiliyorum  
- [ ] String, Long, Double, Boolean kullanımını biliyorum  
- [ ] Nesne atarken `Set` kullanacağımı biliyorum  
- [ ] Modül başına `Option Explicit` yazıyorum  

------------------------------------------------------------

## Sonraki adım

**4. doküman:** [04-VBA-Temelleri-Kosullar-ve-Donguler.md](04-VBA-Temelleri-Kosullar-ve-Donguler.md) — If, Select Case, For ve Do döngüleri.

**Gezinme:** ← [02-Ortam-Kurulumu](02-Ortam-Kurulumu.md) | [Rehber listesi](README.md) | Sonraki: [04-Kosullar-ve-Donguler](04-VBA-Temelleri-Kosullar-ve-Donguler.md) →
