# 1. Giriş – Neden 3DExperience VBA?

```
================================================================================
  Bu rehber: 3DExperience VBA ile makro yazmayı adım adım anlatan başlangıç rehberi
================================================================================
```

Bu rehber, kodlamaya yeni başlayan biri için **3DExperience platformunda VBA ile makro yazmayı** adım adım anlatır. İlk dokümanda ne yapacağımızı ve neden VBA kullandığımızı netleştiriyoruz.

--------------------------------------------------------------------------------

## Bu rehber kimler için?

- Daha önce hiç kod yazmamış veya çok az yazmış olanlar  
- 3DExperience (CATIA, DELMIA vb.) kullanıp tekrarlayan işleri otomatikleştirmek isteyenler  
- VBA’yı sıfırdan, 3DExperience’a özel örneklerle öğrenmek isteyenler  

------------------------------------------------------------

## VBA nedir?

**VBA (Visual Basic for Applications)**, Microsoft’un ofis ve mühendislik yazılımlarına gömülü bir programlama dilidir. 3DExperience içinde de makro ve otomasyon script’leri VBA ile yazılır.

- **Avantajlar:** Yazılımın içinden çalışır, kurulumu ayrı değildir; öğrenmesi nispeten kolaydır; 3DExperience’ın tüm nesnelerine (parça, montaj, çizim vb.) doğrudan erişirsiniz.  
- **Sınır:** Genelde tek bir bilgisayarda, 3DExperience açıkken çalışır; büyük kurumsal otomasyon için C#/VB.NET veya API’ler de kullanılabilir ama başlangıç için VBA yeterlidir.

------------------------------------------------------------

## 3DExperience’ta VBA ile neler yapılabilir?

- **Parça (Part):** Parametre okuma/yazma (Length, Angle vb.), parametre listesini dosyaya yazma, Shapes/Bodies/HybridBodies listeleme, geometri ekleme (HybridShapeFactory, ShapeFactory), kütle/atalet (InertiaService), ölçüm (MeasureService).
- **Montaj (Product):** Alt bileşen listesi (Children), BOM çıkarma, kök occurrence (PLMProductService), occurrence ağacında gezinme, PLM öznitelikleri (GetAttributeValue/SetAttributeValue).
- **Çizim (Drawing):** Sayfa listesi (Sheets), görünüm listesi (Views), ölçek (Scale), ölçüler (Dimensions), metinler (DrawingTexts); CATDrawingService ile çizime özel işlemler.
- **Dosya ve oturum:** FileSystem ile dosya/klasör var mı, boyut, klasör listesi; belge açma/kapama (PLMOpenService, PLMNewService), kaydetme (PLMPropagateService), arama (SearchService).
- **Arayüz:** Mesaj kutusu (MsgBox), giriş kutusu (InputBox) ile kullanıcıdan bilgi almak.

Yani tekrarlayan, kuralı belli işleri **makro** ile otomatikleştirirsiniz. **Servisler** ve **yapılabilecek işlemlerin** detaylı listesi için **12. doküman:** [12-Servisler-ve-Yapilabilecek-Islemler.md](12-Servisler-ve-Yapilabilecek-Islemler.md). **“Buna nereden erişirim, bunu nasıl kullanırım?”** sorusunun tek sayfa cevabı için **13. doküman:** [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md); **Excel** için [14-VBA-ve-Excel-Etkilesimi.md](14-VBA-ve-Excel-Etkilesimi.md); **dosya seç/kaydet diyaloğu** için [15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md). **Help dosyalarını ne zaman/nasıl kullanacağınız** için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md).

------------------------------------------------------------

## Örnek senaryolar – Ne otomatikleştirilebilir?

Aşağıdaki senaryolar, 3DExperience VBA ile sık karşılaşılan otomasyon fikirleridir. Her biri ilerideki dokümanlarda adım adım kodlanabilir.

### Parça (Part) senaryoları

| # | Senaryo | Kısa açıklama | İleride kullanılacak API (ör.) |
|---|---------|----------------|----------------------------------|
| 1 | Tüm parametreleri listele | Part’taki Length, Angle vb. parametrelerin ad ve değerini mesaj veya dosyaya yaz | Parameters, Item, Name, Value |
| 2 | Tek parametreyi güncelle | Kullanıcıdan parametre adı ve yeni değer al; Part’ı güncelle | Parameters.Item("..."), Value, Part.Update |
| 3 | Gövdeye nokta ekle | HybridBodies içine (x,y,z) koordinatında nokta ekle | HybridShapeFactory, AddNewPointCoord, AppendHybridShape |
| 4 | Şekil sayısını raporla | MainBody’deki Shapes sayısını ve her birinin adını göster | Shapes.Count, Shapes.Item(i), Name |
| 5 | Ölçü birimini oku | Part’ın uzunluk birimi (mm, inch vb.) | Part’ın unit ile ilgili property’leri (referansta arayın) |

### Montaj (Product) senaryoları

| # | Senaryo | Kısa açıklama |
|---|---------|----------------|
| 6 | BOM listesi | Tüm alt bileşenlerin adını ve sayısını listele (Children döngüsü) |
| 7 | Kütle toplamı | Montajdaki her parçanın kütlesini topla (InertiaService) |
| 8 | Bileşen sayısı | Kök altındaki toplam occurrence sayısı |

### Çizim (Drawing) senaryoları

| # | Senaryo | Kısa açıklama |
|---|---------|----------------|
| 9 | Sayfa ölçeğini ayarla | Aktif sayfanın Scale değerini oku veya yaz |
| 10 | Görünüm sayısı | Bir sayfadaki Views koleksiyonunun Count’u |
| 11 | Ölçü metnini değiştir | Belirli bir dimension’ın metnini (prefix/suffix) güncelle |

### Dosya ve oturum senaryoları

| # | Senaryo | Kısa açıklama |
|---|---------|----------------|
| 12 | Aktif belge adı | ActiveDocument.Name, FullName |
| 13 | Belgeyi kaydet | Save veya SaveAs (API’ye göre) |
| 14 | Yeni parça aç | PLMNewService veya Documents.Add (sürüme göre) |

Bu tablolar, “ilk hangi makroyu yazayım?” sorusuna yanıt verir; 6–14. dokümanlarda nesne modeli ve API’yi öğrendikten sonra bu senaryolardan birini seçip kodu yazarsınız.

------------------------------------------------------------

## Örnek kod parçası – Sadece fikir (henüz çalıştırmayın)

Aşağıdaki kod **kavramsal** bir örnektir; 3DExperience’a bağlanıp aktif belge adını göstermeyi hedefler. Sözdizimi ve API adları sürüme göre değişebilir; bu rehberin 6. ve 8. dokümanlarında doğru erişim anlatılır.

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x

Sub AktifBelgeAdiniGoster()
    Dim oApp As Object
    Dim oDoc As Object

    On Error Resume Next
    Set oApp = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Or oApp Is Nothing Then
        MsgBox "3DExperience (CATIA) çalışmıyor."
        Exit Sub
    End If
    On Error GoTo 0

    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok."
        Exit Sub
    End If

    MsgBox "Aktif belge: " & oDoc.Name
End Sub
```

```
  AKIS:  Uygulama al  -->  Aktif belge al  -->  Kontrol et  -->  Isle
```
Bu örnekte görülen yapı: **Uygulama al → Aktif belge al → Kontrol et → İşle.** Tüm makrolarınızda benzer bir giriş kullanacaksınız.

------------------------------------------------------------

## Rehberin yapısı

Rehber **18 dokümanlıdır**. Tam liste, özet ve bağlantılar **[README](README.md)** sayfasındaki tabloda yer alır. Temel akış 1–10’dur; 11–18 kurallar, servisler, erişim rehberi, Excel, dosya diyalogları, iyileştirme önerileri, Help kullanımı ve sık hatalar/dikkat noktaları için ek dokümanlardır.

------------------------------------------------------------

## Dassault Systèmes resmi bakış (Help referansı)

**Help-Automation Development Guidelines** dokümanına göre:

- Dassault Systèmes Native Client, makroların **birden fazla dilde** yazılmasını destekler; örnekler ve sözdizimi ağırlıklı olarak **VBScript** ve **Visual Basic for Applications (VBA)** için verilir.
- Desteklenen seviyeler ve diller, Program Directory’deki Native Client bölümünde listelenir.
- Karmaşık bir uygulama (ör. tam Visual Basic) geliştirmek için gerekli tüm konular bu dokümanla kapatılmaz; ancak kuralların çoğu bu tür uygulamalara da uygulanır.
- Genel kural: Bu dokümanla çakışmadığı sürece, dil tedarikçisinin (ör. Microsoft VBScript Coding Conventions, Python Style Guide) belirttiği kurallar geçerlidir.

Yani 3DExperience tarafında **VBA resmi olarak desteklenen** dillerden biridir; script’lerinizi bu dilde yazarken Help’teki kod sunum ve isimlendirme kurallarına uymanız önerilir.

------------------------------------------------------------

## Desteklenen diller (Help’ten)

Automation Development Guidelines’da **Language** alanı için geçerli değerler:

- **CATScript** (uyumluluk amacıyla tutulur)
- **VBScript**
- **VBA**
- **VB.Net**
- **Python**
- **C#**

Bu rehberde yalnızca **VBA** kullanımı anlatılmaktadır; diğer dillerde de benzer nesne modeli ve API’ler kullanılır.

------------------------------------------------------------

## Bölgesel ayarlar uyarısı

Help’e göre: Belirli bir **locale** (bölgesel ayar) için kaydedilen veya yazılan makrolar, başka bir locale’de **çalışmayabilir**. Etkileşimli ürün dokümantasyonunda anlatılan dil adlarını kullanın (örn. English (United States), French (France)). Makroyu paylaşırken veya dokümante ederken hangi bölgesel ayarda test edildiğini belirtin.

------------------------------------------------------------

## Örnek: “Parça adını göster” – Tek satırlık mantık

Makro özetle şunu yapar: Uygulama al → Aktif belge al → Belge adını göster. Aşağıdaki kod bu mantığın basit bir ifadesidir (gerçek API sürüme göre uyarlanır):

```vba
Sub ParcaAdiGoster()
    Dim oApp As Object
    Dim oDoc As Object
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "Uygulama yok.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
    MsgBox "Aktif belge: " & oDoc.Name
End Sub
```

Bu üç adım (uygulama → belge → bilgi) neredeyse tüm makrolarınızda tekrarlanacak.

------------------------------------------------------------

## Örnek: Hangi işler için makro yazılır?

Aşağıdaki liste, 3DExperience kullanıcılarının sık otomatikleştirdiği işlerden örneklerdir. Her biri rehberin ilerleyen dokümanlarında kod örnekleriyle desteklenir.

1. **Parametre toplu güncelleme:** Aynı part’taki birçok Length/Angle parametresini bir tablodan veya kullanıcı girişinden okuyup tek seferde yazma.  
2. **Parametre raporu:** Tüm parametre adlarını ve değerlerini CSV veya metin dosyasına yazma.  
3. **BOM (bill of materials) çıkarma:** Montajdaki tüm bileşenlerin adını, sayısını veya revizyonunu listeleme.  
4. **Çizim sayfa/görünüm bilgisi:** Aktif çizimdeki sayfa sayısı, her sayfadaki görünüm sayısı veya ölçek bilgisini okuma.  
5. **Belge bilgisi:** Açık belgelerin adını, tam yolunu veya türünü (Part/Product/Drawing) listeleme.  
6. **Geometri sayımı:** Part’taki Shapes veya HybridShapes sayısını raporlama.  
7. **Basit veri girişi:** InputBox ile kullanıcıdan parametre adı veya değer alıp tek bir parametreyi güncelleme.  
8. **Birim dönüştürme:** Okunan parametre değerini mm’den inch’e (veya tersi) çevirip başka yerde kullanma.  
9. **Log tutma:** Makro adımlarını veya hatalarını tarih/saat ile dosyaya yazma.  
10. **Koşullu işlem:** Örn. “Sadece adı Length ile başlayan parametreleri güncelle” gibi filtreli döngüler.

Bu liste sizin ilk makro fikrinizi netleştirmenize yardımcı olur; 6. ve 8. dokümanlarda nesne modeli ve API kullanımını öğrendikten sonra bu işlerden birini seçip kodu yazabilirsiniz.

------------------------------------------------------------

## Örnek: Otomasyon türleri (kavramsal)

| Tür | Açıklama | VBA ile |
|-----|----------|---------|
| Etkileşimli | Kullanıcı bir düğmeye basar, makro çalışır | Evet; Tools → Macro → Run |
| Toplu (batch) | Birçok dosya sırayla işlenir | Kısmen; döngü ile aç/kapat/kaydet |
| Raporlama | Veri okunup dosyaya/dış sisteme yazılır | Evet; Parameters, BOM, ölçüler |
| Geometri | Nokta, çizgi, yüzey ekleme | Evet; HybridShapeFactory, ShapeFactory |

3DExperience VBA ile bu dört türün hepsine dokunabilirsiniz; başlangıçta genelde “etkileşimli + tek belge” ile başlanır.

------------------------------------------------------------

## Örnek: Önce kayıt, sonra kod – Önerilen yol

Yeni bir işlemde hangi API’nin kullanıldığını bilmiyorsanız: (1) Makro kaydını başlatın. (2) İşlemi 3DExperience’ta elle bir kez yapın. (3) Kaydı durdurun. (4) Oluşan kodu inceleyin; Application, Document, Part, Parameters, Shapes vb. nesne ve metod isimlerini not alın. (5) Bu isimleri **VBA_API_REFERENCE.md** veya Help metinlerinde arayarak tam imzaları ve alternatiflerini görün. (6) Kodu sadeleştirip Nothing kontrolleri ve Option Explicit ekleyin. Bu yol, 7. dokümanda detaylı anlatılır.

------------------------------------------------------------

## Örnek: Neden VBA, neden C#/Python değil?

Bu rehber **VBA** odaklıdır çünkü: (1) 3DExperience ile birlikte gelir, ek kurulum gerektirmez. (2) Makro kaydı doğrudan VBA kodu üretir; öğrenme hızlanır. (3) Tek bir makinede, etkileşimli kullanım için hızlı geliştirme yapılır. C# veya Python ile otomasyon da mümkündür; ancak COM/API erişimi, ayrı IDE ve dağıtım modeli farklıdır. Kurumsal ve toplu (batch) senaryolarda C#/Python tercih edilebilir; başlangıç için VBA yeterlidir.

------------------------------------------------------------

## Örnek: Rehberdeki dokümanların birbirine bağlantısı

- **01 → 02:** Girişten sonra ortam kurulumu ve ilk makro.  
- **02 → 03–05:** Kurulumdan sonra VBA temelleri (değişken, koşul, döngü, Sub/Function).  
- **05 → 06:** Prosedürlerden nesne modeline (Application, Document, Part/Product/Drawing).  
- **06 → 07–08:** Nesne modelinden makro kaydı ve sık kullanılan API’lere.  
- **08 → 09–10:** API kullanımından hata yakalama ve tam örnek makrolara.  
- **10 → 11:** Örnek projeden resmi kurallar ve hazırlık fazlarına.

Bu sıra takip edildiğinde adım adım 3DExperience VBA ile makro yazma becerisi oluşur.

------------------------------------------------------------

## Sonraki adım

**2. doküman:** [02-Ortam-Kurulumu.md](02-Ortam-Kurulumu.md) — 3DExperience'ı açıp VBA editörüne girmeyi ve ilk makroyu çalıştırmayı öğreneceksiniz. **Sık hatalar ve dikkat edilecekler:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

**Gezinme:** [Rehber listesi](README.md) | Sonraki: [02-Ortam-Kurulumu](02-Ortam-Kurulumu.md) →
