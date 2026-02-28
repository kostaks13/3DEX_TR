# 9. Hata Yakalama ve Debug

**Bu dokümanda:** On Error GoTo / Resume Next; Nothing ve belge kontrolü; breakpoint, Immediate, Locals; log tasarımı; Err.Raise, Retry.

Makro çalışırken hata almak normaldir. Bu dokümanda **hata yakalama** (On Error) ve **debug** (kesme noktası, Immediate penceresi) ile hataları bulup düzeltmeyi öğreneceksiniz.

════════════════════════════════════════════════════════════════════════════════

## On Error – Hata yakalama

VBA’da hata oluşunca çalışma durur ve hata mesajı çıkar. Bunu **yönetmek** için **On Error** kullanırız.

```
  Sub başlar  ──►  On Error GoTo HataYakala  ──►  işlemler  ──►  Exit Sub  ──►  normal çıkış
       │                    │                        │
       │                    └── hata olursa ─────────┴────────►  HataYakala:  ──►  MsgBox  ──►  çıkış
```

### On Error GoTo etiket

Belirli bir satıra atlayıp orada mesaj gösterip çıkış yapabilirsiniz:

```vba
Sub GuvenliCalistir()
    On Error GoTo HataYakala
    
    Dim oApp As Object
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience açık değil."
        Exit Sub
    End If
    
    Dim oDoc As Object
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok."
        Exit Sub
    End If
    
    ' ... işlemler ...
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

- **On Error GoTo HataYakala** → Hata olursa `HataYakala:` satırına gider.  
- **Exit Sub** → Hata yoksa `HataYakala` bloğuna düşmeden çıkar.  
- **Err.Number**, **Err.Description** → Hata kodu ve açıklaması.

### On Error Resume Next

Bir satırda hata olsa bile **bir sonraki satıra** geçer. Dikkatli kullanın; sadece “bu satır hata verebilir, yok say” durumunda kullanın, hemen sonra **Err.Number** kontrolü yapın:

```vba
On Error Resume Next
Set oApp = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then
    MsgBox "CATIA bulunamadı."
    Exit Sub
End If
On Error GoTo 0   ' Sonraki hatalarda normale dön
```

### On Error GoTo 0

Sonrasında hata yakalamayı kapatır; hata tekrar mesaj kutusunda görünür.

════════════════════════════════════════════════════════════════════════════════

## Nothing ve belge kontrolü

3DExperience nesneleri bazen **Nothing** döner (uygulama kapalı, belge yok vb.). Her kritik nesneden sonra kontrol edin:

```vba
Set oDoc = oApp.ActiveDocument
If oDoc Is Nothing Then
    MsgBox "Aktif belge yok. Önce bir parça veya montaj açın."
    Exit Sub
End If
```

════════════════════════════════════════════════════════════════════════════════

## Kesme noktası (Breakpoint)

Kodu **belirli bir satırda durdurup** değişkenleri incelemek için:

1. VBA editöründe satırın solundaki **boş alana** tıklayın; kırmızı nokta (kesme noktası) gelir.  
2. **F5** ile makroyu çalıştırın; o satıra gelince durur.  
3. **F8** (Step Into) ile satır satır ilerleyin.  
4. İmleci bir değişkenin üzerine getirirseniz o anki değeri tooltip’te görürsünüz.

════════════════════════════════════════════════════════════════════════════════

## Immediate penceresi (Ctrl+G)

**View** → **Immediate Window** (veya **Ctrl+G**) ile açar. Kod durduğunda burada komut yazıp Enter’a basarsınız; sonuç hemen görünür.

Örnekler:

- `? oDoc.Name` → Aktif belge adı.  
- `? oShapes.Count` → Shapes sayısı.  
- `? deger` → `deger` değişkeninin o anki değeri.

**Print** yerine **?** kullanılır. Debug sırasında değişkenleri böyle kontrol edin.

════════════════════════════════════════════════════════════════════════════════

## Locals penceresi

**View** → **Locals** ile açar. Kod durduğunda o anda geçerli tüm yerel değişkenler ve değerleri listelenir; nesneler açılıp alt property’ler incelenebilir.

════════════════════════════════════════════════════════════════════════════════

**Sık yapılan hatalar ve dikkat edilmesi gereken özel noktalar** için ayrıca **18. doküman:** [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md).

════════════════════════════════════════════════════════════════════════════════

## İyi alışkanlıklar

1. **Option Explicit** kullanın; yazım hataları azalır.  
2. Kritik API çağrılarından sonra **Nothing** ve **Err** kontrolü yapın.  
3. Uzun makrolarda **Exit Sub** ile erken çıkışları net yazın; böylece hata etiketi sadece gerçek hata için çalışır.  
4. Mümkünse hata mesajına **nerede** olduğunu ekleyin: `MsgBox "Parametre okunurken hata: " & Err.Description`.

════════════════════════════════════════════════════════════════════════════════

## Hata sınıflandırması (3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ’nden)

Help’teki hazırlık yönergesine göre hatalar seviyeye ayrılabilir; her seviye için farklı tepki uygun olur:

| Seviye | Kaynak | Örnek | Etki | Önerilen eylem |
|--------|--------|--------|------|-----------------|
| 0 | Derleme | Option Explicit eksik, sözdizimi hatası | Makro açılmaz / derlenmez | Geliştirici düzeltir |
| 1 | Ortam | Workbench yüklü değil, lisans yok | Makro atlanır | Kullanıcı uyar + log |
| 2 | Nesne | Null, Count=0, HybridBodies boş | Güvenli çıkış | Log + uyarı mesajı |
| 3 | İş mantığı | Geometri/veri hatası (örn. offset < 0) | Devam veya geri al | Log + anlamlı hata mesajı |
| 4 | Altyapı | Dosya/klasör yok, ağ/PLM timeout | Makro geri al | Log + Retry veya sapma raporu |
| 5 | Kritik | Çıktı tutarsız (örn. kütle sapması > %5) | Rollback ve işlemi engelle | Log + paydaş bilgisi |

Kurumsal ortamda **Err.Raise** ile kendi hata numara aralığınızı (örn. 9000–9999) kullanmanız önerilir; böylece VBA/ sistem hata kodlarıyla karışmaz. Detay için **11. doküman** (Resmi Kurallar ve Hazırlık Fazları) ve Help’teki “Hataları Yakala & Günlüğe Kaydet” bölümüne bakın.

════════════════════════════════════════════════════════════════════════════════

## Log (günlük) tasarımı önerisi (Help’ten)

Hata ve izlenebilirlik için basit bir log yapısı:

- **Konum:** `%TEMP%` veya kurumsal paylaşım (örn. `\\Server\MacroLogs\`).
- **Format:** TSV — Timestamp, ErrNo, Message, Context (ve isteğe bağlı INFO/WARN/ERROR/CRIT etiketi).
- **Boyut:** Dosya belirli bir boyuta (örn. 5 MB) ulaşınca _old yapıp yeni dosya açın.
- **Başlık:** Her makro çalışmasında `"--- START v1.3 ---"` gibi bir satır yazın.

Örnek log satırı: `2025-07-12 10:42:15  ERROR  9100  HybridBodies missing  Part=Wing_Rib_A.CATPart`

════════════════════════════════════════════════════════════════════════════════

## Örnek: Tam hata yakalama şablonu (Exit Sub ile)

Uzun makrolarda her kritik adımdan sonra erken çıkış yapıp, sadece gerçek runtime hatalarında `HataYakala` etiketine düşülmesini sağlayın:

```vba
Option Explicit

Sub GuvenliMakroSablonu()
    Dim oApp As Object
    Dim oDoc As Object

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

    ' ... asıl işlemler ...
    MsgBox "İşlem tamamlandı."
    Exit Sub

HataYakala:
    MsgBox "Beklenmeyen hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: On Error Resume Next ve Err temizleme

GetObject ile uygulama alırken hata verebilir; hemen sonra Err kontrolü yapıp Err.Clear ile temizleyin (bir sonraki Resume Next bloğunda karışmasın diye):

```vba
Sub GetObjectGuvenli()
    Dim oApp As Object
    On Error Resume Next
    Set oApp = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Or oApp Is Nothing Then
        MsgBox "CATIA bulunamadı. Err: " & Err.Number
        Err.Clear
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    ' Buradan sonra normal hata yakalama
    MsgBox "Uygulama alındı: " & oApp.Name
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Err.Raise ile özel hata (Help önerisi)

Kendi hata numaranızı (9000–9999) kullanarak çağıran kodu bilgilendirmek:

```vba
Sub ParametreKontrolOrnek()
    Dim sAd As String
    sAd = InputBox("Parametre adı:")
    If sAd = "" Then
        Err.Raise 9001, "ParametreKontrolOrnek", "Parametre adı boş olamaz."
    End If
    ' Devam...
End Sub
```

Çağıran Sub’da `On Error GoTo` ile 9001’i yakalayıp kullanıcıya anlamlı mesaj gösterebilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Basit log dosyasına yazma

Hata ve bilgi mesajlarını dosyaya yazan bir yardımcı Sub (FileSystemObject veya 3DExperience FileSystem kullanılabilir):

```vba
Sub LogYaz(sMesaj As String, sDosyaYolu As String)
    Dim iFile As Integer
    Dim sSatir As String
    sSatir = Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & sMesaj & vbCrLf
    iFile = FreeFile
    Open sDosyaYolu For Append As #iFile
    Print #iFile, sSatir;
    Close #iFile
End Sub

Sub OrnekKullanim()
    LogYaz "Makro başladı.", "C:\Temp\macro_log.txt"
    On Error GoTo HataYakala
    ' ... işlemler ...
    LogYaz "Makro bitti.", "C:\Temp\macro_log.txt"
    Exit Sub
HataYakala:
    LogYaz "HATA " & Err.Number & ": " & Err.Description, "C:\Temp\macro_log.txt"
    MsgBox "Hata oluştu. Log dosyasına bakın."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Immediate penceresinde debug

Kod kesme noktasında durduğunda Immediate penceresinde (Ctrl+G) şunları deneyin:

```
? oDoc.Name
? oPart.Shapes.Count
? i
```

Değişken atamak için:

```
i = 10
sAd = "Length.1"
```

Böylece değişkenleri değiştirip bir sonraki adımda davranışı test edebilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Watch penceresi

Kritik bir değişkenin her adımda değerini izlemek için: **Debug** → **Add Watch** ile değişkeni ekleyin. Kod breakpoint’te durduğunda Watch penceresinde güncel değer görünür; bu özellikle döngü içinde değişen sayaçlar için faydalıdır.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Hata numarasına göre farklı mesaj (Select Case Err.Number)

Bazı hatalarda kullanıcıya özel mesaj göstermek için **Err.Number** kullanın:

```vba
Sub HataNumarasinaGoreMesaj()
    On Error GoTo HataYakala
    ' ... işlemler ...
    Exit Sub

HataYakala:
    Select Case Err.Number
        Case 91
            MsgBox "Nesne atanmamış (Object variable not set)."
        Case 424
            MsgBox "Gerekli nesne yok."
        Case Else
            MsgBox "Hata " & Err.Number & ": " & Err.Description
    End Select
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Err.Clear ile hata bilgisini temizleme

**On Error Resume Next** kullandıktan sonra hata kontrolü yaptıysanız, bir sonraki blokta eski **Err** değerinin karışmaması için **Err.Clear** kullanın:

```vba
On Error Resume Next
Set oApp = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then
    MsgBox "Uygulama bulunamadı."
    Err.Clear
    Exit Sub
End If
Err.Clear
On Error GoTo 0
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Debug.Print ile pencereye yazma

**MsgBox** yerine çıktıyı **Immediate** penceresine yazmak için **Debug.Print** kullanın; böylece birçok satırı hızlıca inceleyebilirsiniz:

```vba
Sub DebugPrintOrnek()
    Dim i As Long
    For i = 1 To 5
        Debug.Print "Adım " & i
    Next i
End Sub
```

Çalıştırdıktan sonra **Ctrl+G** ile Immediate penceresini açıp çıktıyı görün.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Hata mesajında konum bilgisi

Uzun makrolarda hatanın nerede oluştuğunu anlamak için mesaja konum ekleyin:

```vba
Sub KonumluHataMesaji()
    On Error GoTo HataYakala
    Dim oApp As Object
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "[1] Uygulama yok.": Exit Sub
    Dim oDoc As Object
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "[2] Belge yok.": Exit Sub
    Dim oPart As Object
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then MsgBox "[3] Part alınamadı.": Exit Sub
    Exit Sub
HataYakala:
    MsgBox "Hata [" & Err.Number & "] " & Err.Description
End Sub
```

[1], [2], [3] gibi etiketler hangi adımda takıldığını gösterir.

════════════════════════════════════════════════════════════════════════════════

## Örnek: Yeniden deneme (Retry) – Basit döngü

Bazen geçici hata (ağ, dosya kilitli) olur; birkaç kez yeniden denemek isteyebilirsiniz:

```vba
Sub RetryOrnek()
    Dim iDeneme As Long
    Dim bBasarili As Boolean
    bBasarili = False
    For iDeneme = 1 To 3
        On Error Resume Next
        ' ... riskli işlem (örn. dosya açma) ...
        If Err.Number = 0 Then bBasarili = True: Exit For
        Err.Clear
    Next iDeneme
    On Error GoTo 0
    If Not bBasarili Then MsgBox "3 denemede başarısız."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Örnek: Hata seviyesine göre log etiketi (INFO / WARN / ERROR)

Help’teki log tasarımında her satıra seviye etiketi (INFO, WARN, ERROR, CRIT) eklenebilir. Kodda bir yardımcı Sub ile tek bir noktadan log yazılır:

```vba
Sub LogSeviyeli(sDosya As String, sSeviye As String, sMesaj As String)
    Dim iFile As Integer
    iFile = FreeFile
    Open sDosya For Append As #iFile
    Print #iFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & sSeviye & "  " & sMesaj
    Close #iFile
End Sub

' Kullanım: LogSeviyeli "C:\Temp\log.txt", "ERROR", "Parametre bulunamadı"
```

════════════════════════════════════════════════════════════════════════════════

## Kontrol listesi

- [ ] On Error GoTo etiket ile hata yakalayabiliyorum  
- [ ] On Error Resume Next ve Err.Number kontrolü yapabiliyorum  
- [ ] Kesme noktası koyup F8 ile adım adım ilerleyebiliyorum  
- [ ] Immediate penceresinde ? ile değişken değerine bakabiliyorum  

════════════════════════════════════════════════════════════════════════════════

## Sonraki adım

**10. doküman:** [10-Ornek-Proje-Bastan-Sona-Bir-Makro.md](10-Ornek-Proje-Bastan-Sona-Bir-Makro.md) — Baştan sona örnek makro ve yazım kuralları özeti.

**Gezinme:** Önceki: [08-Sik-Kullanilan-APIler](08-Sik-Kullanilan-APIler.md) | [Rehber listesi](README.md) | Sonraki: [10-Ornek-Proje](10-Ornek-Proje-Bastan-Sona-Bir-Makro.md) →
