# 11. Resmi Kurallar ve Hazırlık Fazları (Help Referansı)

Bu doküman, **Help** klasöründeki resmi dokümanlardan özetlenmiştir: **Help-Automation Development Guidelines.pdf** ve **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf**. Kodlamaya başladıktan sonra bu kuralları uygulayarak makrolarınızı kurumsal standartlara yaklaştırabilirsiniz. **Help içindeki dosyaları ne zaman ve nasıl kullanacağınız** (hangi dosyayı hangi aşamada açacağınız, arama yöntemleri) için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md).

------------------------------------------------------------

## 1. Kod sunum kuralları (Automation Development Guidelines)

Dassault Systèmes, script’lerde okunabilirlik için aşağıdaki kuralları önerir. Mümkün olduğunca uyun.

### 1.1 Genel düzen (Layout)

| Kural | Açıklama |
|-------|----------|
| **Girinti** | 4 boşluk kullanın. |
| **Satır uzunluğu** | Satırları 80 karakterden kısa tutun. |
| **Yorum** | `REM` yerine `'` (tek tırnak) kullanın. |
| **Yorum hizası** | İç içe bloklardaki yorumlar, o bloğun girintisiyle hizalı olsun. |
| **Başlık** | Her Sub/Function için başlık yorumu: amaç, parametreler, dönüş değeri. |
| **Değişken yorumu** | Önemli değişkenleri aynı satırda kısa yorumla açıklayın. |

### 1.2 Zorunlu başlık (global header)

Teslim edilen script’lerde aşağıdaki bilgiler **zorunludur**:

- **Copyright** — Dassault Systèmes ürünü içinde teslim ediliyorsa telif bildirimi.
- **Purpose** — Script’in ne yaptığı (2–3 cümle).
- **Assumptions** — Oturum durumu, seçili nesneler gibi dış etkenler.
- **Author** — Müşteri script’iyse doldurulur; örnek script’lerde boş bırakılabilir.
- **Language** — Geçerli değerler: `CATScript`, `VBScript`, `VBA`, `VB.Net`, `Python`, `C#`.
- **Version / Release** — İlk desteklenen sürüm (örn. 3DEXPERIENCE R2017x).
- **Regional Settings** — Makro hangi yerel ayarda kaydedildi/yazıldı (örn. English (United States), French (France)). Farklı lokalde çalışmayabilir; dokümantasyonda belirtin.

Örnek başlık:

```vba
' Purpose: Aktif parçanın HybridBodies içine (0,0,0) noktası ekler.
' Assumptions: GSD workbench açık, aktif belge Part, HybridBodies mevcut.
' Language: VBA
' Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States)
```

### 1.3 İsimlendirme kuralları (VBA / VBScript)

**Sub/Function adları:** Fiil veya fiil ifadesi, her kelimenin ilk harfi büyük (mixed case).

```vba
Sub DoItBetter()
    ...
End Sub
```

**Değişken önekleri (minimal set):** Tipi belirten tek harf öneki kullanın.

| Önek | Tip | Örnek |
|------|-----|--------|
| b | Boolean | bIsUpToDate |
| d | Double | dLength |
| s | String | sName |
| i | Integer | iNumberOfElements |
| o | Object | oSketch1 |
| c | Collection | cDrwViews |

**Sabitler:** Tamamı büyük harf, bileşenler alt çizgi ile ayrılmış. Örnek: `MAX_VALUE`, `LOG_PATH`.

------------------------------------------------------------

## 2. Genel programlama kuralları (Help’ten)

### 2.1 Option Explicit

Derleyici/yorumlayıcının tanımsız değişken oluşturmasını engeller; yazım hatalarını azaltır. **Modülün ilk satırı** olmalı:

```vba
Option Explicit
```

Kayıt edilmiş makroyu temel alıyorsanız, `Language = "VBScript"` gibi gereksiz atamaları silin; değişkeni kaldırın.

### 2.2 Hata yönetimi

- Hata yönetimi **sistematik** olmalı; yoksa kullanıcı sorunu geç fark eder.
- Varsayılan davranış: Hata oluşunca yorumlayıcı durur ve hata kutusu gösterir.
- **Düzeltici aksiyon** alacaksanız `On Error Resume Next` kullanmak zorundasınız; kullanırken **mutlaka** hataları manuel kontrol edin ve **kısa süre** içinde `On Error GoTo 0` ile normale dönün.

Örnek (Help’ten):

```vba
Dim CATIA As Object
On Error Resume Next
Set CATIA = GetObject(, "CATIA.Application")
iErr = Err.Number
If iErr <> 0 Then
    On Error GoTo 0
    Set CATIA = CreateObject("CATIA.Application")
End If
On Error GoTo 0
```

### 2.3 Değişkenleri tipleyin (VBA/CATScript/VSTA)

Debug’u kolaylaştırır:

```vba
Dim VariableOfTypeX As TypeX
```

İstisna: Bir metodun argümanı dizi (array) istiyorsa ve VBA/VSTA tipi kabul etmiyorsa, o nesneyi tipleyemeyebilirsiniz; dokümantasyonda belirtilir.

### 2.4 Dil ve sürüm bildirimi

Aşağıdaki yorumlar **zorunludur**. Language: `VBScript`, `CATScript`, `VBA`, `VB.NET`, `C#`. Release: ilk desteklenen sürüm.

```vba
' Language: VBA
' Release:  V6R2010
```

### 2.5 Test edilebilirlik kuralları

- **InputBox:** Otomatik test ortamında kullanılabilmesi için **varsayılan değer** (üçüncü parametre) verin.  
  Örnek: `InputBox("Enter Body Name: ", "", "NEWBODY")`
- **Hata çıkışı:** Test senaryosunda hatanın “hata” olarak tanınması için **MsgBox** yerine **Err.Raise** kullanın.  
  Örnek: `If (iShape > cShapes.Count) Then Err.Raise 9999, "MyMacro", "Shape Number is too big"`

### 2.6 Platformlar arası script (Cross-Platform)

- **Sadece V6 nesneleri:** Windows’ta `CreateObject`/`GetObject` kullanılabilir; Unix’te değil. V6 alternatifi varsa onu kullanın.
- **Dosya sistemi:** Windows Scripting Host’un `Scripting.FileSystemObject` taşınabilir değildir. Bunun yerine **Application.FileSystem** (CATIA.FileSystem) üzerinden **File**, **Folder** nesnelerini kullanın.
- **Yol birleştirme:** Farklı kaynaklardan gelen yolları `&` ile birleştirmek yerine **CATIA.SystemService.ConcatenatePaths** kullanın; böylece `/` ve `\` karışımı ve platform farkları doğru yönetilir.

Örnek (Help’ten):

```vba
Dim sRootPath As String
sRootPath = CATIA.SystemService.Environ("ROOT_FOLDER")
Dim sFilePath As String
sFilePath = CATIA.SystemService.ConcatenatePaths(sRootPath, "drw/myData.txt")
```

------------------------------------------------------------

## 3. İhtiyaç analizi çerçevesi (3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ)

Kodu yazmadan önce “ne, neden, hangi kısıtlarla?” sorularını yanıtlamak için beş katmanlı analiz önerilir.

### Katman A – Amaç ve iş sonucu

1. **İş hedefi** — Parça mı oluşturulacak, montaj verisi mi raporlanacak?  
2. **Başarı ölçütü** — “Makro çalışınca X işlemi Y sürede bitmeli.”  
3. **Kullanıcı profili** — Tasarım mühendisi mi, simülasyon mühendisi mi?  
4. **Zaman / sıklık** — Bir kerelik mi, günlük/kronik mi?  
5. **Paydaşlar** — Kim onaylayacak; kimlerin modeline dokunuluyor?

### Katman B – Teknik kapsam

1. **Çalışma nesnesi** — Part / Product / Drawing / VPMReference?  
2. **Workbench** — GSD, Part Design, Composites, Simulation…  
3. **Ana API modülü** — HybridShapeFactory, PLMProductService, VOCServices vb.  
4. **Girdi** — Kullanıcı seçimi, parametre listesi, dosya, PLM araması?  
5. **Çıktı** — Geometri, öznitelik güncellemesi, rapor dosyası?

### Katman C – Operasyon ve bağımlılıklar

1. **Veri kaynağı** — Yerel .CATPart mı, PLM’den referans mı?  
2. **Erişim yetkisi** — Kullanıcının lisansı/workbench’i var mı?  
3. **Altyapı** — Malzeme kütüphanesi, rulebase, harici DLL, ağ paylaşımları?  
4. **Kritik yol** — Hangi adımda durursa iş aksar?  
5. **Performans sınırı** — Örn. 10.000 occurrence’lı montaj taranacak mı?

### Katman D – Kısıtlar ve riskler

- **Politika** — “V5 API yasak”, “yalnızca VBScript” vb.  
- **Güvenlik** — PLM’de sadece okuma / yazma izni?  
- **Rollback** — Hatalı yazımda geri alma gerekir mi?  
- **Versiyonlama** — Yeni nesne oluşturulacaksa revizyon stratejisi?  
- **Hata yansımaları** — Başarısız güncelleme montaj ilişkilerini bozar mı?

### Katman E – Dokümantasyon ve doğrulama

1. **Kabul testi** — “Seçilen iki yüz arasında ölçü raporu (mm) = 25 ± 0.1 yazdırılmalı.”  
2. **Günlük / geri bildirim** — Başarıda mesaj; hata varsa log dosyası.  
3. **Versiyon notasyonu** — Makro sürümü, API değişiklik zamanı, Help referansı.  
4. **Eğitim / el kitabı** — Kullanıcıya kısa kullanım yönergesi.  
5. **Bakım planı** — Hangi Help sürümü değişirse makro revize edilecek?

### Uygulanabilir kontrol formu (minimum soru seti)

1. Ne yapılacak?  
2. Hangi model(ler) / workbench?  
3. Girdi nereden geliyor?  
4. Çıktı ne olacak, kime gidecek?  
5. Ölçek? (parça / montaj büyüklüğü)  
6. Süre/lisans kısıtı var mı?  
7. Başarının teknik ölçütü?  
8. Hata durumda ne olmalı?  
9. Ek risk veya politika kısıtı?

Bu form doldurulmadan teknik tasarıma (modül seçimi vb.) geçilmemeli; eksik alan varsa kullanıcıdan netleştirme istenir.

------------------------------------------------------------

## 4. Modül ve API eşleştirme matrisi (Help’ten)

Gereksinimin hangi “dünyada” (Part Design, GSD, Assembly, Simulation vb.) çözüleceğini aşağıdaki tabloya göre seçin.

| Gereksinim tipi | Workbench | Temel nesne | Factory / Service |
|-----------------|-----------|-------------|--------------------|
| Basit katı (Pad, Pocket) | Part Design | Part / Bodies | ShapeFactory |
| İleri yüzey (Sweep, Loft) | GSD | Part / HybridBodies | HybridShapeFactory |
| Montaj ağacı | Assembly Design | Product / VPMOccurrence | PLMProductService |
| Kütle, atalet | — | Part/Product | InertiaService |
| Ölçüm (mesafe, açı) | — | Seçim veya ref | MeasureService / MeasurableService |
| Parametrik kural | Knowledge | ParameterSet | KnowledgeFactory |
| Kompozit | Composites | Part | CompositesServices |
| Hacim analizi | — | VPMRepReference | VOCServices |
| Malzeme atama | — | PLMEntity | MaterialService |
| PLM arama | — | — | SearchService |
| Simülasyon koşusu | Simulation | ScenarioRepresentation | SimExecutionService |

**Editor-level vs Session-level servis:**  
- **Editor-level:** Nesneye özgü, anlık işlem (geometri, kütle, ölçüm). Genelde RAM odaklı, hızlı.  
- **Session-level:** Yeni nesne yaratma, PLM arama, malzeme kütüphanesi, toplu kaydetme. Ağ gecikmesi ve PLM doğrulaması olur.  
- **Kural:** Editor-level servis içinde session-level servis çağrısı yapmayın; kilitleşme riski doğar.

------------------------------------------------------------

## 5. Kod taslağı fazları: Design → Draft → Harden → Finalize

### 5.1 Design (akış tasarımı)

- **High-level pseudo-code** — Her satır tek iş: “Geometrical Set’i bul”, “Nokta oluştur”, “Update”.  
- **Veri sözleşmesi** — Girdi/çıktı listesi, parametre adları, birim (mm/deg), seçim gerekli mi?  
- **Risk matrisi** — Null koleksiyon, lisans yok, read-only attribute…  
- **Kod başlığı** — Amaç, sürüm, yazar, tarih, kullanılan API’ler.

### 5.2 Draft (iskelet kod)

- **Option Explicit** + başlık yorum bloğu (JIRA ID, Help referansı, sürüm).  
- **Sub … End Sub** — Tek giriş noktası; gerekirse alt prosedürler.  
- **Değişken bildirimi** — Önekli adlandırma (oPart, dOffset, vSel).  
- **Hata bloğu** — On Error Resume Next (ve sonrasında GoTo 0 veya yapısal handler).  
- **Ana API nesneleri** — Late-binding ile al: `Dim oHSF As Object : Set oHSF = ...`  
- **Boş/null kontrolleri** — Her `Set` sonrası `If obj Is Nothing Then Exit Sub` (veya Err.Raise).  
- **İş mantığı** — Sadece AddNew* veya Service çağrıları; parametre yer tutucu.

### 5.3 Harden (güvenli ve performanslı kod)

- **Null ve Count** — Tüm koleksiyonlar için `If Is Nothing Or .Count = 0`.  
- **Workbench doğrulama** — Gerekirse `CATIA.StartWorkbench("Generative Shape Design")`.  
- **Read-only** — `If oPart.ReadOnly Then Exit Sub`.  
- **Tek Update** — Döngü içinde değil, sonda bir kez `oPart.Update`.  
- **Hata log** — FileSystem + OpenAsTextStream(ForAppend).  
- **Süre ölçümü** — `t0 = Timer` … `Debug.Print "Süre: ", Timer - t0`.  
- **Bellek** — Büyük döngü sonunda `Set oHB = Nothing`.

### 5.4 Finalize (temizlik ve belgeleme)

- **Kod inceleme** — Eski V5 API yok; tüm Set sonrası null kontrolü; Option Explicit, sürüm, Help sayfa no; On Error GoTo 0 kapatılmış.  
- **Yorum** — Her 10–15 satırda bir kısa açıklama.  
- **Versiyon etiketi** — `'-- REV 1.1 – 2025-07-13: Offset parametre eklendi`  
- **Dağıtım** — “Makro %CATStartupPath%\Macros\ konumuna kopyalanacak.”  
- **Kullanım (3 satır)** — “GSD aktifken çalıştır → Part güncel → ‘Done’ mesajı.”

------------------------------------------------------------

## 6. Hata sınıflandırması ve log (Help’ten)

| Seviye | Kaynak | Örnek | Etki | Eylem |
|--------|--------|--------|------|--------|
| 0 | Derleme | Option Explicit eksik | Makro açılmaz | Geliştirici düzeltir |
| 1 | Ortam | Workbench/Lisans yok | Makro atlanır | Kullanıcı uyar + log |
| 2 | Nesne | Null/Count, HybridBodies boş | Güvenli çıkış | Log + uyarı |
| 3 | İş mantığı | Offset mm < 0 | Devam veya geri al | Log + hata mesajı |
| 4 | Altyapı | Reports dizini yok | Geri al | Log + Retry/Sapma |
| 5 | Kritik | Çıktı tutarsız (örn. kütle sapması > %5) | Rollback & engelle | Log + paydaş bilgisi |

**Log tasarımı önerisi:**  
- Konum: `%TEMP%` veya `\\Server\MacroLogs\`  
- Format: TSV — Timestamp, ErrNo, Message, Context.  
- Boyut: 5 MB’a ulaşınca _old yapıp yeni dosya.  
- Düzey: INFO, WARN, ERROR, CRIT tag’leri.  
- Her makro başında: `"--- START v1.3 ---"` satırı.

------------------------------------------------------------

## 7. TAMAM / HAZIR kontrol listesi (özet)

**Çekirdek – Minimum (mutlaka ✓):**

1. İhtiyaç formu (Katman A–E) yanıtlandı.  
2. Doğru workbench ve API seçimi (matriste işaretli).  
3. Option Explicit var.  
4. On Error Resume Next + On Error GoTo 0 veya GoTo handler var.  
5. Tüm `Set ... = ... GetItem` sonrası Is Nothing / .Count kontrolü.  
6. Tek bir `oPart.Update` (döngü içinde yok).  
7. Eski V5 API yok (Documents.Add, HybridShapeFactoryOld vb.).  
8. Kod derleniyor ve manuel testte çalışıyor.  
9. Kullanıcıya en az bir açıklama (MsgBox/echo).

**Kurumsal – Geniş (önerilen ✓):**

- Hata: Özel Err.Raise 9000–9999; log konumu ve rotasyonu; rollback tanımlı.  
- Performans: Timer ile süre rapora yazıldı; 10K+ occurrence testinde timeout yok.  
- Güvenlik/Lisans: Workbench varlığı test edilip loglandı; read-only tespiti.  
- Versiyon: Başlıkta sürüm, tarih, Help sayfa no; değişiklik günlüğü güncel.  
- Dokümantasyon: 7 bölümlü teslim paketi; kullanım adımları; en az 1 “sonraki adım” önerisi.  
- Paydaş: Talep sahibi “Çalıştı” (UAT PASS) onayı verdi.

------------------------------------------------------------

## 8. Örnek: Tam başlık ve Option Explicit (uyumlu kod)

Aşağıdaki blok, Automation Development Guidelines’daki zorunlu başlık ve isimlendirme kurallarına uygun bir modül başlangıcıdır:

```vba
Option Explicit
' ============================================================
' Purpose: Aktif parçanın parametre sayısını ve ilk parametre adını gösterir.
' Assumptions: 3DExperience açık, aktif belge Part.
' Author: [Ekip / Müşteri]
' Language: VBA
' Release: 3DEXPERIENCE R2024x
' Regional Settings: English (United States)
' ============================================================

Sub ShowParameterCount()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim iCount As Long
    Dim sName As String

    On Error GoTo ErrHandler
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub

    iCount = oParams.Count
    If iCount > 0 Then
        sName = oParams.Item(1).Name
    Else
        sName = "(yok)"
    End If
    MsgBox "Parametre sayısı: " & iCount & vbCrLf & "İlk parametre: " & sName
    Exit Sub

ErrHandler:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

Bu örnekte: 4 boşluk girinti, önekli değişkenler (o, i, s), tek başlık bloğu, On Error GoTo, Nothing kontrolleri ve tek bir Exit Sub kullanılmıştır.

------------------------------------------------------------

## 9. Örnek: Cross-Platform – FileSystem ve ConcatenatePaths

Help’e göre taşınabilir script için **CATIA.FileSystem** ve **CATIA.SystemService.ConcatenatePaths** kullanılmalıdır. Aşağıdaki örnek, ortam değişkeninden kök yol alıp alt dizinle birleştirir (sözdizimi sürüme göre değişebilir):

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Not: Cross-platform için FileSystem ve SystemService kullanın.

Sub BuildPathExample()
    Dim oApp As Object
    Dim oSys As Object
    Dim sRoot As String
    Dim sSub As String
    Dim sFull As String

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oSys = oApp.SystemService
    If oSys Is Nothing Then Exit Sub

    sRoot = oSys.Environ("ROOT_FOLDER")
    If sRoot = "" Then sRoot = "C:\Temp"
    sSub = "Macros\Reports"
    sFull = oSys.ConcatenatePaths(sRoot, sSub)
    MsgBox "Birleşik yol: " & sFull
End Sub
```

Dosya var mı kontrolü için `oApp.FileSystem.Exists(sFull)` kullanılabilir (FileSystem API’sine bakın).

------------------------------------------------------------

## 10. Örnek: InputBox varsayılan değer (test edilebilirlik)

Help’teki test edilebilirlik kuralına göre InputBox’ta **üçüncü parametrede varsayılan değer** verin; böylece otomatik test senaryolarında aynı değer kullanılabilir:

```vba
Sub TestEdilebilirInput()
    Dim sBodyName As String
    Dim sParamName As String
    ' Varsayılan: "NEWBODY" ve "Length.1" — test script'leri bu değerleri kullanabilir
    sBodyName = InputBox("Gövde adı:", "Giriş", "NEWBODY")
    sParamName = InputBox("Parametre adı:", "Giriş", "Length.1")
    If sBodyName = "" Or sParamName = "" Then Exit Sub
    ' ... işlem
End Sub
```

------------------------------------------------------------

## 11. Örnek: Err.Raise ile test senaryosunda hata (Help’ten)

Hata durumunda test otomasyonunun “hata” olarak tanıması için **MsgBox** yerine **Err.Raise** kullanın:

```vba
Sub CheckShapeIndex(ByVal iIndex As Long)
    Dim oPart As Object
    Dim oShapes As Object
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Err.Raise 9001, "CheckShapeIndex", "Part not found"
    Set oShapes = oPart.Shapes
    If oShapes Is Nothing Then Err.Raise 9002, "CheckShapeIndex", "Shapes not found"
    If iIndex < 1 Or iIndex > oShapes.Count Then
        Err.Raise 9999, "CheckShapeIndex", "Shape Number is too big"
    End If
    ' ... devam
End Sub
```

Çağıran kodda `On Error GoTo` ile 9999 veya 9001 yakalanıp log’a yazılabilir.

------------------------------------------------------------

## 12. Örnek: Design fazı – Pseudo-code

Design fazında önce yüksek seviye adımlar yazılır; sonra her biri kod satırına dönüşür:

```
1. Uygulama nesnesini al (GetObject).
2. ActiveDocument al; Nothing ise çık.
3. Part nesnesini al (GetItem("Part") veya oDoc).
4. Parameters koleksiyonunu al; Nothing veya Count=0 ise çık.
5. For i = 1 To Parameters.Count:
   - Item(i) ile parametreyi al.
   - Name ve Value'yu log veya dosyaya yaz.
6. Hata durumunda: log + MsgBox, Exit Sub.
```

Bu liste “Draft” fazında `Dim`, `Set`, `For ... Next` ve `On Error GoTo` ile VBA’ya çevrilir.

------------------------------------------------------------

## 13. Örnek: Harden fazı – Tek Update ve Timer

Döngü içinde **Update** çağrılmaz; tüm değişiklikler bittikten sonra **bir kez** Update. İsteğe bağlı süre ölçümü:

```vba
Sub HardenPhaseExample()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim t0 As Single

    t0 = Timer
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub

    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            ' Örnek: sadece belirli parametreleri güncelle
            ' oParam.Value = yeniDeger
        End If
    Next i
    ' Döngü dışında tek Update
    oPart.Update
    Debug.Print "Süre (sn): "; Timer - t0
End Sub
```

------------------------------------------------------------

## 14. Örnek: Finalize – Versiyon etiketi ve dağıtım notu

Kod başlığına versiyon ve dağıtım bilgisi eklenir:

```vba
' -- REV 1.2 – 2025-02-15: Parametre listesi dosyaya yazma eklendi.
' -- Dağıtım: Bu makro %CATStartupPath%\Macros\ klasörüne kopyalanacak.
' -- Kullanım: 1) GSD veya Part Design aktif. 2) Part açık. 3) Makroyu çalıştır; "Done" mesajı görünce işlem tamam.
```

------------------------------------------------------------

## 15. Örnek: Log satırı formatı (TSV)

Help’teki log tasarımına uygun örnek satırlar:

```
2025-02-15 10:30:01  INFO  0  Makro başladı  Part=Wing_A.CATPart
2025-02-15 10:30:02  INFO  0  Parametre sayısı: 12  Part=Wing_A.CATPart
2025-02-15 10:30:03  ERROR 9100  Length.1 bulunamadı  Part=Wing_A.CATPart
2025-02-15 10:30:04  INFO  0  Makro bitti (hata ile)  Part=Wing_A.CATPart
```

Kodda `LogYaz "ERROR", 9100, "Length.1 bulunamadı", "Part=" & oDoc.Name` gibi bir yardımcı Sub ile bu satırlar üretilir.

------------------------------------------------------------

## 16. Örnek: Modül matrisi kullanımı

“Parametreleri toplu güncelle” gereksinimi için matristen seçim:

- **Gereksinim tipi:** Parametrik kural / Part parametreleri.  
- **Workbench:** Part Design.  
- **Temel nesne:** Part.  
- **API:** Part.Parameters, Parameter.Value, Part.Update.

Sonuç: Part.Parameters.Item("...").Value ataması ve döngü sonunda tek Part.Update kullanılır; KnowledgeFactory bu senaryoda gerekmez.

------------------------------------------------------------

## 17. Örnek: Katman A–E formu (kısa doldurma)

Örnek senaryo: “Aktif parçadaki tüm Length parametrelerini 100 mm yap.”

| Katman | Örnek yanıt |
|--------|-------------|
| A – Amaç | İş hedefi: Part parametrelerini toplu güncelleme. Başarı: Tüm Length.* parametreleri 100. |
| B – Teknik | Part, Part Design, Parameters API. Girdi: yok (veya parametre listesi dosyası). Çıktı: Part güncel. |
| C – Operasyon | Veri: Aktif belge. Lisans: Part Design. Kritik yol: Parameters erişimi. |
| D – Kısıt | V5 API yok. Read-only part’ta yazma yapılmaz. |
| E – Dokümantasyon | Kabul: Parametre değerleri 100 ± 0.001. Log: INFO/ERROR. |

Bu form doldurulduktan sonra Draft fazına geçilir.

------------------------------------------------------------

## 18. Örnek: Girinti ve satır uzunluğu (4 boşluk, 80 karakter)

Help’e göre 4 boşluk girinti ve satırları 80 karakterden kısa tutun. Uzun satırları alt satıra bölmek için `_` (satır devamı) kullanın:

```vba
Sub UzunSatirOrnek()
    Dim sMesaj As String
    sMesaj = "Bu çok uzun bir mesaj satırıdır ve 80 karakteri " & _
             "aşmamak için alt satıra bölünmüştür."
    MsgBox sMesaj
End Sub
```

------------------------------------------------------------

## 19. Örnek: Sub/Function başlık yorumu (amaç, parametreler, dönüş)

Her prosedür için kısa başlık yorumu yazın:

```vba
' Purpose: Aktif parçadaki parametre sayısını döndürür.
' Params:  Yok.
' Return:  Long — parametre sayısı; hata/Part yoksa 0.

Function GetParameterCount() As Long
    Dim oPart As Object
    GetParameterCount = 0
    On Error Resume Next
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Function
    If oPart.Parameters Is Nothing Then Exit Function
    GetParameterCount = oPart.Parameters.Count
End Function
```

------------------------------------------------------------

## 20. Örnek: V5 API’den kaçınma (yasak / farklı API)

Help’e göre aşağıdaki gibi **eski V5** çağrıları 3DExperience’ta kullanılmamalı veya farklı API ile değiştirilmelidir:

| Eski (V5) | 3DExperience uygun alternatif (ör.) |
|-----------|-------------------------------------|
| Documents.Add("Part") | PLMNewService veya Session/Application dokümantasyonundaki “yeni belge” API’si |
| Selection.Search | ActiveEditor.Selection + ilgili Filter/Add (referansa bakın) |
| HybridShapeFactoryOld | HybridShapeFactory (Native Apps Automation) |
| CreateReferenceFromName | Referans oluşturma için güncel Reference API’si |

Kayıt edilmiş makroda bu tür isimler görürseniz Help-Native Apps Automation ve VBA_API_REFERENCE.md ile doğru sınıfı bulun.

------------------------------------------------------------

## 21. Örnek: Rollback stratejisi (kavramsal)

Hata durumunda yapılan değişiklikleri geri almak için tasarım aşamasında karar verin:

- **Değişiklik yok:** Sadece okuma (parametre listesi, rapor) — rollback gerekmez.
- **Tek parametre yazıldı:** Hata sonrası eski değeri tekrar yazmak (değeri başta bir değişkende saklayın).
- **Birden fazla nesne değişti:** İşlem öncesi durumu (parametre değerleri, seçim) kaydedin; hata olursa geri yükleyin veya kullanıcıyı uyarın (“Kısmen uygulandı, lütfen manuel kontrol edin”).

Kodda: `On Error GoTo Rollback` ve `Rollback:` etiketinde saklanan değerleri geri yazma veya log’a “ROLLBACK” yazma.

------------------------------------------------------------

## 22. Örnek: 7 bölümlü teslim paketi (kurumsal)

Kurumsal teslimatta aşağıdaki bölümlerin doldurulması önerilir (Help – Hazırlık Yönergesi):

1. **İhtiyaç özeti** — Ne yapılacak, hangi workbench, girdi/çıktı.  
2. **Risk matrisi** — Null, lisans, read-only, timeout.  
3. **Kod başlığı** — Purpose, Assumptions, Language, Release, Author.  
4. **Test senaryosu** — Adımlar ve beklenen sonuç.  
5. **Kullanım yönergesi** — 3 satırlık kısa talimat.  
6. **Dağıtım** — Nereye kopyalanacak, hangi rol/güvenlik.  
7. **Sonraki adım** — Opsiyonel iyileştirmeler (log rotasyonu, timeout ayarı vb.).

------------------------------------------------------------

**Sık yapılan hatalar ve dikkat edilmesi gereken özel noktalar** (Nothing, Update, On Error, V5 API, InputBox iptal, locale vb.) için **18. doküman:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

------------------------------------------------------------

## 23. Örnek: Kod inceleme kontrol listesi (Finalize)

Dağıtım öncesi kod incelemesinde şunları kontrol edin:

- [ ] **Option Explicit** modül başında var.  
- [ ] **Language** ve **Release** yorumu var.  
- [ ] Tüm **Set** atamalarından sonra **Is Nothing** veya **Count** kontrolü yapılıyor.  
- [ ] **On Error Resume Next** kullanılan yerlerde kısa süre sonra **On Error GoTo 0** veya **GoTo** handler var.  
- [ ] **oPart.Update** (veya benzeri) döngü **içinde** değil, işlemler bittikten sonra **bir kez** çağrılıyor.  
- [ ] **Documents.Add**, **HybridShapeFactoryOld**, **Selection.Search** gibi V5 API’ler kullanılmıyor.  
- [ ] Değişken isimleri önekli (o, d, s, i, b, c) veya en azından anlamlı.  
- [ ] Her Sub/Function için kısa **Purpose** (ve gerekirse Params, Return) yorumu var.

------------------------------------------------------------

## 24. Örnek: Risk matrisi satırları (Design fazı)

Tasarım aşamasında risk matrisine eklenecek örnek satırlar:

| Risk | Olasılık | Etki | Önlem |
|------|----------|------|--------|
| ActiveDocument Nothing | Orta | Makro çalışmaz | Başta Nothing kontrolü, MsgBox |
| Part belgesi değil | Orta | Parametre/Shapes erişilemez | GetItem("Part") Nothing ise uyar |
| Parametre adı yanlış | Orta | Item bulunamaz | On Error, kullanıcıya mesaj |
| Read-only belge | Düşük | Yazma başarısız | ReadOnly kontrolü (varsa API’de) |
| Çok büyük Part (10K+ shape) | Düşük | Yavaşlama | Döngüde Timer veya max iterasyon |

Bu tabloyu ihtiyacınıza göre genişletebilirsiniz; 11. dokümandaki Katman D (Kısıtlar ve riskler) ile uyumludur.

------------------------------------------------------------

## 25. Örnek: 3 satırlık kullanım yönergesi (Finalize)

Help’teki “3 satırlık kullanım” önerisine uygun örnek metin:

1. **GSD** veya **Part Design** rolünde çalışın; açılacak belge bir **Part** (.CATPart) olsun.  
2. Makroyu **Tools → Macro → Run** (veya atadığınız düğme) ile çalıştırın.  
3. Ekranda **“Done”** veya **“İşlem tamamlandı”** mesajını gördüğünüzde işlem tamamdır; hata durumunda log dosyasına bakın.

Bu metni makro teslim paketine veya başlık yorumuna ekleyin.

------------------------------------------------------------

## 26. Örnek: Sabit (Const) isimlendirme – Büyük harf ve alt çizgi

Help’e göre sabitler tamamı büyük harf, kelimeler alt çizgi ile ayrılır:

```vba
Const MAX_ITERATION As Long = 100
Const DEFAULT_TOLERANCE As Double = 0.001
Const LOG_FILE_PATH As String = "C:\Temp\macro_log.txt"
```

------------------------------------------------------------

## 27. Örnek: Sub/Function adı – Mixed case fiil

Sub ve Function adları fiil veya fiil ifadesi; her kelimenin ilk harfi büyük (mixed case):

```vba
Sub UpdateParameterValue()
Sub GetActivePart()
Function ComputeTotalMass() As Double
```

------------------------------------------------------------

## 28. Örnek: Değişken yorumu – Aynı satırda kısa açıklama

Önemli değişkenleri aynı satırda kısa yorumla açıklayın (Help – Layout):

```vba
Dim oPart As Object   ' Aktif parça belgesi
Dim iCount As Long   ' Parametre sayısı
Dim sLogPath As String   ' Log dosyası yolu
```

------------------------------------------------------------

## 29. Örnek: Katman B – Teknik kapsam (form doldurma)

“Parametreleri dosyadan okuyup Part’a yaz” gereksinimi için Katman B örnek yanıtı:

- **Çalışma nesnesi:** Part.  
- **Workbench:** Part Design.  
- **Ana API modülü:** Part.Parameters, Parameter.Value, Part.Update.  
- **Girdi:** Metin/CSV dosyası (parametre adı; değer).  
- **Çıktı:** Part güncel; isteğe bağlı log dosyası.

------------------------------------------------------------

## 30. Örnek: Log dosyası rotasyonu (5 MB’a ulaşınca _old)

Help’teki log tasarımında dosya belirli boyuta (örn. 5 MB) ulaşınca _old yapılıp yeni dosya açılır. VBA’da **FileLen** ile dosya boyutu alınabilir; 5 MB’ı aşarsa dosyayı yeniden adlandırıp yeni dosya açılır (kod örneği bu dokümanın kapsamı dışında; mantık olarak Finalize fazında eklenebilir).

------------------------------------------------------------

## 31. Örnek: Kod başlığı – Copyright (telif)

Dassault ürünü içinde teslim edilen script’lerde **Copyright** satırı zorunludur (Help). Müşteri veya ekip script’iyse kendi telif metninizi yazın:

```vba
' Copyright (c) 2025 [Şirket Adı]. Tüm hakları saklıdır.
' Purpose: ...
' Language: VBA
' Release: 3DEXPERIENCE R2024x
```

------------------------------------------------------------

## 32. Örnek: Regional Settings yorumu

Makro hangi bölgesel ayarda (locale) kaydedildi/yazıldı ise bunu belirtin; farklı locale’de çalışmayabilir uyarısı Help’te yer alır:

```vba
' Regional Settings: English (United States)
' Not: Farklı dil/bölge ortamında test edilmedi.
```

------------------------------------------------------------

## 33. Örnek: Assumptions (varsayımlar) – Başlıkta

Teslim edilen script’lerde **Assumptions** alanı doldurulmalıdır: Hangi workbench açık olacak, hangi belge türü, kullanıcı ne seçmiş olacak vb.

```vba
' Assumptions: Part Design veya GSD açık; aktif belge Part; en az bir HybridBody mevcut.
```

------------------------------------------------------------

## 34. Örnek: Katman C – Operasyon (form doldurma)

“Parametreleri dosyadan okuyup Part’a yaz” için Katman C örnek yanıtı:

- **Veri kaynağı:** Yerel metin/CSV dosyası.  
- **Erişim yetkisi:** Kullanıcı Part Design lisansına sahip olmalı.  
- **Altyapı:** Dosya yolu yazılabilir olmalı.  
- **Kritik yol:** Dosya açma, Parameters.Item, Part.Update.  
- **Performans sınırı:** Parametre sayısı 1000’i aşmamalı (isteğe bağlı).

------------------------------------------------------------

## 35. Örnek: TAMAM/HAZIR – Geniş listeden seçmeler

Kurumsal teslimatta ek olarak şunlar önerilir: **Err.Raise 9000–9999** ile özel hata; **log dosyası** konumu ve rotasyonu; **rollback** tanımı; **Timer** ile süre raporu; **Workbench** varlığı testi; **versiyon etiketi** ve **değişiklik günlüğü**; **7 bölümlü teslim paketi**; **talep sahibi onayı** (UAT PASS).

------------------------------------------------------------

## 36. Örnek: Kod sunum – 80 karakter satır sınırı

Help’e göre satırları 80 karakterden kısa tutun. Uzun ifadeleri alt satıra bölmek için **satır devam karakteri** (_) kullanın:

```vba
MsgBox "Bu çok uzun bir mesaj satırıdır ve 80 karakteri aşmamak için " & _
       "alt satıra bölünmüştür."
```

------------------------------------------------------------

## 37. Örnek: Yorum hizası – İç içe bloklarda

İç içe bloklardaki yorumlar, o bloğun girintisiyle hizalı olmalıdır (Help – Layout):

```vba
Sub Ornek()
    ' Dış blok yorumu
    If True Then
        ' İç blok yorumu – 4 boşluk daha girintili
        MsgBox "Test"
    End If
End Sub
```

------------------------------------------------------------

## 38. Örnek: Versiyon etiketi – Değişiklik günlüğü

Finalize fazında başlığa veya ayrı bir “Changelog” bölümüne versiyon notları ekleyin:

```vba
' -- REV 1.0 – 2025-02-01: İlk sürüm.
' -- REV 1.1 – 2025-02-15: Parametre listesi dosyaya yazma eklendi.
' -- REV 1.2 – 2025-02-28: Log seviyesi (INFO/ERROR) eklendi.
```

------------------------------------------------------------

## Referanslar

- **Help-Automation Development Guidelines.pdf** — Kod sunum kuralları, isimlendirme, hata yönetimi, cross-platform.  
- **3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.pdf** — İhtiyaç analizi, modül eşleştirme, kod taslağı fazları, hata ve log, teslim protokolü, TAMAM/HAZIR listesi.  
- **Help-Native Apps Automation.pdf** — Application, Editors, ActiveEditor, Services, FileSystem nesne modeli.

Bu doküman, yukarıdaki kaynaklardan özetlenmiştir; tam ve güncel kurallar için ilgili PDF’lere bakın.

**Tüm rehber listesi:** [README](README.md). İlgili: [12](12-Servisler-ve-Yapilabilecek-Islemler.md) servisler, [13](13-Erisim-ve-Kullanim-Rehberi.md) erişim, [14](14-VBA-ve-Excel-Etkilesimi.md) Excel, [15](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md) dosya diyalogları, [16](16-Iyilestirme-Onerileri.md) iyileştirme.

**Gezinme:** ← [10-Ornek-Proje](10-Ornek-Proje-Bastan-Sona-Bir-Makro.md) | [Rehber listesi](README.md) | Sonraki: [12-Servisler](12-Servisler-ve-Yapilabilecek-Islemler.md) →
