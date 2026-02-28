# 10. Örnek Proje – Baştan Sona Bir Makro

**Bu dokümanda:** Tam makro iskeleti (AktifParcaBilgisi); parametre oku/yaz, listele, dosyaya yaz; modüler yapı (GetActivePart); Design/Draft/Harden/Finalize.

Bu son dokümanda, önceki dokümanlarda öğrendiklerinizi kullanarak **baştan sona** bir makro iskeleti yazıyoruz. Ardından kısa bir **yazım kuralları** özeti veriyoruz.

════════════════════════════════════════════════════════════════════════════════

## Senaryo

- 3DExperience açık olacak.  
- Aktif belge bir **parça** (Part) olacak.  
- Makro: Uygulamaya ve belgeye bağlanacak, belge adını gösterecek, (varsa) parametre sayısını veya şekil sayısını mesajla bildirecek.  
- Hata olursa anlamlı mesaj verip çıkacak.

Bu iskeleti kendi ihtiyacınıza göre (parametre değiştirme, çizim sayfası işleme vb.) genişletebilirsiniz.

```
  GetObject → oApp   ──►   ActiveDocument → oDoc   ──►   GetItem("Part") → oPart
       │                          │                              │
       │ Nothing?                 │ Nothing?                     │ Nothing?
       ▼                          ▼                              ▼
  [1] Uygulama al            [2] Belge al                   [3] Part al
                                                                  │
                                                                  ├── .Parameters, .Shapes, .Name …
                                                                  │
                                                                  ▼
  [4] İşlem (oku/yaz/listele)   ──►   [5] (gerekirse) oPart.Update   ──►   [6] MsgBox / çıkış
```

════════════════════════════════════════════════════════════════════════════════

## Tam örnek makro (iskelet)

```vba
Option Explicit

' ============================================================
' Örnek: Aktif parça belgesinin bilgisini göster
' 3DExperience açık olmalı, aktif belge bir Part olmalı.
' ============================================================

Sub AktifParcaBilgisi()
    On Error GoTo HataYakala
    
    Dim oApp As Object
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then
        MsgBox "3DExperience (CATIA) çalışmıyor. Önce uygulamayı açın."
        Exit Sub
    End If
    
    Dim oDoc As Object
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "Açık belge yok. Bir parça açın."
        Exit Sub
    End If
    
    Dim bilgi As String
    bilgi = "Belge adı: " & oDoc.Name & vbCrLf
    bilgi = bilgi & "Tam yol: " & oDoc.FullName & vbCrLf
    
    ' Parça nesnesi (sürüme göre API farklı olabilir; gerekirse kayıt makrosu ile doğrulayın)
    Dim oPart As Object
    Set oPart = oDoc.GetItem("Part")   ' veya: Set oPart = oDoc
    If Not oPart Is Nothing Then
        Dim oShapes As Object
        On Error Resume Next
        Set oShapes = oPart.Shapes
        If Err.Number = 0 And Not oShapes Is Nothing Then
            bilgi = bilgi & "Shapes sayısı: " & oShapes.Count & vbCrLf
        End If
        On Error GoTo HataYakala
    End If
    
    MsgBox bilgi
    Exit Sub

HataYakala:
    MsgBox "Hata (" & Err.Number & "): " & Err.Description
End Sub
```

**Not:** `GetItem("Part")`, `Shapes`, `Count` gibi isimler 3DExperience sürümüne göre değişebilir. Kendi ortamınızda makro kaydı yapıp üretilen kodu bu iskeletle karşılaştırın; gerekirse referanstaki Part/Shapes API’sine bakın.

════════════════════════════════════════════════════════════════════════════════

## Adım adım ne yaptık?

1. **Option Explicit** — Değişkenleri tanımlı kullandık.  
2. **On Error GoTo HataYakala** — Hata olursa tek yerde mesaj verip çıkıyoruz.  
3. **GetObject(, "CATIA.Application")** — Çalışan 3DExperience oturumuna bağlandık.  
4. **Nothing kontrolleri** — Application ve Document yoksa anlamlı mesaj.  
5. **oDoc.GetItem("Part")** — Parça nesnesini aldık (sürüme göre uyarlayın).  
6. **oPart.Shapes** — Şekil koleksiyonu; Count ile sayı (varsa).  
7. **Exit Sub** — Normal çıkışta hata bloğuna düşmemek için.  
8. **HataYakala:** — Hata numarası ve açıklamasını gösteriyoruz.

════════════════════════════════════════════════════════════════════════════════

## Yazım kuralları özeti (guideline)

| Kural | Açıklama |
|-------|----------|
| **Option Explicit** | Her modülün başında olsun. |
| **Anlamlı isimler** | `parcaAdi`, `olcuDegeri`, `oPart`, `oDoc` gibi. |
| **Nothing kontrolü** | Application, Document, Part gibi nesneleri aldıktan sonra kontrol edin. |
| **On Error** | Makro içinde en az bir kez hata yakalayın; kullanıcıya net mesaj verin. |
| **Kısa Sub’lar** | Uzun işleri mantıklı Sub veya Function’lara bölün. |
| **Yorum satırı** | Makronun ne yaptığını ve önemli adımları kısa yorumla açıklayın. |
| **Referans kullanın** | Yeni API kullanırken `VBA_API_REFERENCE.md` ve Help metinlerine bakın. |
| **Kayıt + sadeleştirme** | Bilmediğiniz işlemde önce makro kaydedin, sonra kodu sadeleştirip genelleştirin. |

════════════════════════════════════════════════════════════════════════════════

## Kod taslağı fazları (Help – 3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ)

Resmi hazırlık yönergesine göre makro kodu dört fazda olgunlaştırılır:

1. **Design (akış tasarımı)** — Pseudo-code, veri sözleşmesi (girdi/çıktı), risk matrisi (null, lisans, read-only), kod başlığı.
2. **Draft (iskelet)** — Option Explicit, başlık yorumu, tek Sub girişi, önekli değişkenler, On Error, API nesnelerinin alınması, her Set sonrası Nothing/Count kontrolü, iskelet iş satırları.
3. **Harden (güvenlik ve performans)** — Null/Count, workbench doğrulama, read-only kontrolü, **tek** Update (döngü içinde değil), hata log’u, süre ölçümü (Timer), büyük döngü sonunda nesneleri Nothing yapma.
4. **Finalize (temizlik ve belgeleme)** — Kod inceleme (V5 API yok, null kontrolü, On Error GoTo 0), yorum, versiyon etiketi, dağıtım notu, 3 satırlık kullanım yönergesi.

Bu fazların tam listesi, hata/log/rollback detayları ve dağıtım öncesi kontrol listesi (TAMAM/HAZIR) için **11. doküman:** [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md) ve Help’teki “Kodu Standart Şablona Göre Üret” bölümüne bakın.

════════════════════════════════════════════════════════════════════════════════

## Kendi makronuzu büyütmek için

- **Parametre değiştirme:** `Parameters.Item("...").Value = ...` ve referansta Parameters/Value araması.  
- **Çizim:** DrawingRoot → Sheets → Views → Dimensions/DrawingTexts; referansta Drawing ile başlayan sınıflara bakın.  
- **Montaj:** Product → Children; bileşen ekleme/çıkarma için ilgili Add/Remove metodlarını referansta arayın.  
- **Dosya:** Yeni belge açma, kaydetme, farklı kaydetme için Application/Document dokümantasyonunu kullanın.

════════════════════════════════════════════════════════════════════════════════

Rehber özeti ve tüm doküman listesi için [README](README.md).

════════════════════════════════════════════════════════════════════════════════

## Sonraki adımlar

- Projedeki **VBA_API_REFERENCE.md** dosyasında ihtiyacınız olan sınıf ve metodu bulun.  
- **Help/text/** altındaki metinlerde ilgili bölümleri okuyun (hangi dosyayı ne zaman açacağınız için **17. doküman:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md)).  
- Küçük bir görev seçip (ör. “bir parametreyi okuyup mesajla göster”) makroyu yazın, çalıştırın; gerekirse 9. dokümandaki debug yöntemleriyle hatayı bulun.  
- **Erişim/kullanım** için [13-Erisim-ve-Kullanim-Rehberi.md](13-Erisim-ve-Kullanim-Rehberi.md); **Excel** için [14-VBA-ve-Excel-Etkilesimi.md](14-VBA-ve-Excel-Etkilesimi.md); **dosya seç/kaydet diyaloğu** için [15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md](15-Dosya-Secme-ve-Kaydetme-Diyaloglar.md); **iyileştirme önerileri** için [16-Iyilestirme-Onerileri.md](16-Iyilestirme-Onerileri.md); **sık hatalar ve dikkat edilecekler** için [18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) kullanın.

Bu rehber, 3DExperience VBA ile sıfırdan kod yazmaya başlamanız için 18 dokümanlık bir guideline setidir; her doküman bir sonrakine bağlanır ve hepsi 3DExperience VBA’ye özeldir.

════════════════════════════════════════════════════════════════════════════════

## İkinci örnek: Parametre oku ve mesajla göster

Bu makro, kullanıcıdan parametre adı alır, aktif parçada o parametreyi bulur ve değerini mesajla gösterir. Tüm adımlar (Application → Document → Part → Parameters → Item) ve hata kontrolleri tek Sub içinde:

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x
' Purpose: Aktif parçada belirtilen parametrenin değerini okur ve gösterir.

Sub ParametreOkuVeGoster()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim sParamAdi As String
    Dim dDeger As Double

    On Error GoTo HataYakala

    sParamAdi = InputBox("Parametre adı (örn. Length.1):", "Parametre oku", "Length.1")
    If Trim(sParamAdi) = "" Then Exit Sub

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "3DExperience açık değil.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Açık belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler alınamadı.": Exit Sub

    On Error Resume Next
    Set oParam = oParams.Item(sParamAdi)
    On Error GoTo HataYakala
    If oParam Is Nothing Then
        MsgBox "Parametre bulunamadı: " & sParamAdi
        Exit Sub
    End If

    dDeger = oParam.Value
    MsgBox sParamAdi & " = " & dDeger
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Üçüncü örnek: Parametre yaz (kullanıcıdan değer al)

Aynı zincir; bu sefer kullanıcı hem parametre adı hem yeni değer girer, makro değeri yazar ve Part’ı günceller:

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x

Sub ParametreYaz()
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim sParamAdi As String
    Dim sDegerStr As String
    Dim dDeger As Double

    On Error GoTo HataYakala

    sParamAdi = InputBox("Parametre adı:", "Parametre yaz", "Length.1")
    If Trim(sParamAdi) = "" Then Exit Sub
    sDegerStr = InputBox("Yeni değer:", "Parametre yaz", "100")
    If Trim(sDegerStr) = "" Then Exit Sub

    dDeger = CDbl(sDegerStr)

    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then MsgBox "Uygulama yok.": Exit Sub
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    On Error Resume Next
    Set oParam = oParams.Item(sParamAdi)
    On Error GoTo HataYakala
    If oParam Is Nothing Then MsgBox "Parametre yok: " & sParamAdi: Exit Sub

    oParam.Value = dDeger
    oPart.Update
    MsgBox sParamAdi & " = " & dDeger & " olarak güncellendi."
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Dördüncü örnek: Shapes sayısı ve ilk birkaç şeklin adı

Parça nesnesinden Shapes koleksiyonunu alıp Count ve Item(i).Name ile bilgi toplayan makro:

```vba
Option Explicit

Sub ShapesBilgisi()
    Dim oApp As Object
    Dim oPart As Object
    Dim oShapes As Object
    Dim oShape As Object
    Dim i As Long
    Dim iMax As Long
    Dim sOut As String

    On Error GoTo HataYakala
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Exit Sub
    Set oPart = oApp.ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oShapes = oPart.Shapes
    If oShapes Is Nothing Then MsgBox "Shapes yok.": Exit Sub

    sOut = "Shapes sayısı: " & oShapes.Count & vbCrLf
    iMax = oShapes.Count
    If iMax > 10 Then iMax = 10
    For i = 1 To iMax
        Set oShape = oShapes.Item(i)
        If Not oShape Is Nothing Then sOut = sOut & i & ": " & oShape.Name & vbCrLf
    Next i
    If oShapes.Count > 10 Then sOut = sOut & "..."
    MsgBox sOut
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Beşinci örnek: Modüler yapı – Yardımcı Sub ve Function

İş mantığını Sub’lara bölerek okunabilirliği artırma örneği. Ana Sub sadece akışı çağırır; Application ve Part alma ayrı bir Function’da:

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x

Private Function GetActivePart() As Object
    Dim oApp As Object
    Dim oDoc As Object
    Dim oPart As Object
    On Error Resume Next
    Set oApp = GetObject(, "CATIA.Application")
    If oApp Is Nothing Then Set GetActivePart = Nothing: Exit Function
    Set oDoc = oApp.ActiveDocument
    If oDoc Is Nothing Then Set GetActivePart = Nothing: Exit Function
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Set oPart = oDoc
    Set GetActivePart = oPart
End Function

Sub AnaParametreListesi()
    Dim oPart As Object
    Dim oParams As Object
    Dim i As Long
    Dim sOut As String

    Set oPart = GetActivePart()
    If oPart Is Nothing Then MsgBox "Aktif parça alınamadı.": Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    sOut = "Parametre sayısı: " & oParams.Count & vbCrLf
    For i = 1 To oParams.Count
        On Error Resume Next
        sOut = sOut & oParams.Item(i).Name & " = " & oParams.Item(i).Value & vbCrLf
        On Error GoTo 0
    Next i
    MsgBox sOut
End Sub
```

Bu yapıda `GetActivePart` tekrar kullanılabilir; yeni makrolarda aynı Function’ı çağırabilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Altıncı örnek: Dosyaya parametre listesi yaz

Parametre adı ve değerini ekranda değil, bir metin dosyasına yazan makro. VBA’nın `Open ... For Output` kullanımı:

```vba
Option Explicit

Sub ParametreListesiniDosyayaYaz()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim iFile As Integer
    Dim sDosya As String

    sDosya = "C:\Temp\parametre_listesi.txt"
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then MsgBox "Parça yok.": Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then MsgBox "Parametreler yok.": Exit Sub

    iFile = FreeFile
    Open sDosya For Output As #iFile
    Print #iFile, "Parametre;Değer"
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then Print #iFile, oParam.Name & ";" & oParam.Value
    Next i
    Close #iFile
    MsgBox "Liste yazıldı: " & sDosya
End Sub
```

`GetActivePart()` kullanmak isterseniz önceki örnekteki Function’ı bu modüle ekleyin.

════════════════════════════════════════════════════════════════════════════════

## Yedinci örnek: Tarih/saat ile log satırı

Makro başlangıcında ve bitişinde (ve hata durumunda) tek bir log dosyasına satır yazan pattern:

```vba
Sub LogOrnekMakro()
    Const LOG_PATH As String = "C:\Temp\macro_log.txt"
    On Error GoTo HataYakala
    LogSatir LOG_PATH, "START " & "ParametreListesi"
    ' ... işlemler ...
    LogSatir LOG_PATH, "END OK"
    Exit Sub
HataYakala:
    LogSatir LOG_PATH, "END ERR " & Err.Number & " " & Err.Description
    MsgBox "Hata: " & Err.Description
End Sub

Private Sub LogSatir(sDosya As String, sMesaj As String)
    Dim iFile As Integer
    iFile = FreeFile
    Open sDosya For Append As #iFile
    Print #iFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & sMesaj
    Close #iFile
End Sub
```

Bu yapıyı 9. dokümandaki “Log tasarımı” ile birleştirerek daha zengin log formatı (ERROR/INFO etiketi, kontekst) ekleyebilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## Sekizinci örnek: Sadece Part belgesi kontrolü

Belge Part değilse (ör. çizim veya montaj açıksa) işlem yapmayıp kullanıcıyı uyaran makro:

```vba
Sub SadecePartKontrol()
    Dim oDoc As Object
    Dim oPart As Object
    Set oDoc = GetObject(, "CATIA.Application").ActiveDocument
    If oDoc Is Nothing Then MsgBox "Belge yok.": Exit Sub
    On Error Resume Next
    Set oPart = oDoc.GetItem("Part")
    On Error GoTo 0
    If oPart Is Nothing Then
        MsgBox "Aktif belge bir parça değil. Lütfen bir .CATPart açın."
        Exit Sub
    End If
    MsgBox "Parça belgesi hazır: " & oDoc.Name
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Dokuzuncu örnek: İki parametreyi takas et

İki parametrenin değerini geçici değişkenle takas eden makro (Length.1 ↔ Length.2):

```vba
Sub IkiParametreTakas()
    Dim oPart As Object
    Dim oParams As Object
    Dim oP1 As Object
    Dim oP2 As Object
    Dim d1 As Double
    Dim d2 As Double
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    Set oP1 = oParams.Item("Length.1")
    Set oP2 = oParams.Item("Length.2")
    If oP1 Is Nothing Or oP2 Is Nothing Then MsgBox "Parametre bulunamadı.": Exit Sub
    d1 = oP1.Value
    d2 = oP2.Value
    oP1.Value = d2
    oP2.Value = d1
    oPart.Update
    MsgBox "Length.1 ve Length.2 takas edildi."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## Onuncu örnek: Minimum–maksimum parametre değeri bul

Tüm parametreler arasında en küçük ve en büyük sayısal değeri bulup mesajla gösterir:

```vba
Sub MinMaxParametreDeger()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim dMin As Double
    Dim dMax As Double
    Dim dV As Double
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Or oParams.Count = 0 Then Exit Sub
    dMin = oParams.Item(1).Value
    dMax = dMin
    For i = 2 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            dV = oParam.Value
            If dV < dMin Then dMin = dV
            If dV > dMax Then dMax = dV
        End If
    Next i
    MsgBox "Min: " & dMin & vbCrLf & "Max: " & dMax
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On birinci örnek: Başarı / hata sonunda tek mesaj

Uzun işlemlerde her adımda MsgBox açmak yerine, sonda tek bir özet mesajı göstermek daha iyidir:

```vba
Sub TekOzetMesaj()
    Dim bBasarili As Boolean
    Dim sMesaj As String
    bBasarili = True
    On Error GoTo HataYakala
    ' ... birçok adım ...
    ' Bir adımda hata olursa bBasarili = False yapılır
    If bBasarili Then
        sMesaj = "İşlem tamamlandı."
    Else
        sMesaj = "İşlem hata ile tamamlandı. Log dosyasına bakın."
    End If
    MsgBox sMesaj
    Exit Sub
HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On ikinci örnek: Kullanıcıdan iki sayı al ve parametrelere yaz

İki parametre adı ve iki değer kullanıcıdan alınır; Part’ta bu parametreler varsa değerler yazılır:

```vba
Sub IkiParametreKullanicidanYaz()
    Dim oPart As Object
    Dim oParams As Object
    Dim oP1 As Object, oP2 As Object
    Dim s1 As String, s2 As String
    Dim v1 As String, v2 As String
    Dim d1 As Double, d2 As Double

    s1 = InputBox("1. parametre adı:", "", "Length.1")
    s2 = InputBox("2. parametre adı:", "", "Length.2")
    v1 = InputBox("1. parametre değeri:", "", "100")
    v2 = InputBox("2. parametre değeri:", "", "50")
    If s1 = "" Or s2 = "" Or v1 = "" Or v2 = "" Then Exit Sub
    d1 = CDbl(v1)
    d2 = CDbl(v2)

    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    On Error Resume Next
    Set oP1 = oParams.Item(s1)
    Set oP2 = oParams.Item(s2)
    On Error GoTo 0
    If Not oP1 Is Nothing Then oP1.Value = d1
    If Not oP2 Is Nothing Then oP2.Value = d2
    oPart.Update
    MsgBox "Güncellendi: " & s1 & "=" & d1 & ", " & s2 & "=" & d2
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On üçüncü örnek: Parametre sayısını Function ile döndür

Aktif parçanın parametre sayısını döndüren Function; ana makro bu sayıyı kullanır:

```vba
Function AktifParcaParametreSayisi() As Long
    Dim oPart As Object
    AktifParcaParametreSayisi = 0
    On Error Resume Next
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Function
    If oPart.Parameters Is Nothing Then Exit Function
    AktifParcaParametreSayisi = oPart.Parameters.Count
End Function

Sub ParametreSayisiGoster()
    Dim n As Long
    n = AktifParcaParametreSayisi()
    MsgBox "Parametre sayısı: " & n
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On dördüncü örnek: Belge adına göre işlem (Like ile filtre)

Sadece belge adı belirli bir pattern’e uyan Part’larda işlem yapmak için **Like** kullanın:

```vba
Sub SadeceWingParcalarinda()
    Dim oDoc As Object
    Dim oPart As Object
    Set oDoc = GetObject(, "CATIA.Application").ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    If Not oDoc.Name Like "Wing*" Then
        MsgBox "Bu makro sadece adı Wing ile başlayan parçalarda çalışır."
        Exit Sub
    End If
    Set oPart = oDoc.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    MsgBox "İşlem yapılıyor: " & oDoc.Name
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On beşinci örnek: Toplu parametre güncelleme – Dizi kullanımı

Parametre adları ve yeni değerler iki ayrı dizi (veya Array) ile verilip döngüde güncellenir:

```vba
Sub TopluParametreGuncelle()
    Dim oPart As Object
    Dim oParams As Object
    Dim oParam As Object
    Dim arrNames As Variant
    Dim arrValues As Variant
    Dim i As Long
    arrNames = Array("Length.1", "Length.2", "Length.3")
    arrValues = Array(100, 200, 300)
    Set oPart = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part")
    If oPart Is Nothing Then Exit Sub
    Set oParams = oPart.Parameters
    If oParams Is Nothing Then Exit Sub
    For i = LBound(arrNames) To UBound(arrNames)
        On Error Resume Next
        Set oParam = oParams.Item(arrNames(i))
        If Not oParam Is Nothing Then oParam.Value = arrValues(i)
        On Error GoTo 0
    Next i
    oPart.Update
    MsgBox "Toplu güncelleme tamamlandı."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On altıncı örnek: Hata durumunda eski değeri geri yazma (Rollback)

Tek parametre yazarken hata olursa eski değeri geri yazmak için değeri başta saklayın:

```vba
Sub ParametreYazRollback()
    Dim oParam As Object
    Dim dEski As Double
    Dim dYeni As Double
    Set oParam = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part").Parameters.Item("Length.1")
    If oParam Is Nothing Then Exit Sub
    dEski = oParam.Value
    dYeni = 999
    On Error GoTo Rollback
    oParam.Value = dYeni
    GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part").Update
    MsgBox "Güncellendi."
    Exit Sub
Rollback:
    oParam.Value = dEski
    MsgBox "Hata oluştu; eski değer geri yüklendi: " & dEski
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On yedinci örnek: Timer ile süre ölçümü (Harden fazı)

Help’teki Harden fazında işlem süresini ölçüp log’a veya Debug’a yazmak önerilir. VBA’da **Timer** fonksiyonu saniye cinsinden değer döndürür:

```vba
Sub SureOlcumOrnek()
    Dim t0 As Single
    Dim t1 As Single
    t0 = Timer
    ' ... ağır işlemler (parametre döngüsü, Update vb.) ...
    t1 = Timer
    Debug.Print "Süre (sn): "; t1 - t0
    MsgBox "İşlem süresi: " & Format(t1 - t0, "0.00") & " saniye"
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On sekizinci örnek: Sadece sayısal parametreleri listele (tip kontrolü)

Bazı parametreler metin (String) veya tarih olabilir. Sadece sayısal (Double) parametreleri listelemek için Value’yu CDbl ile deneyip hata vermeyenleri sayın veya parametre tipi property’sine bakın (API’de Parameter.Type veya benzeri varsa):

```vba
Sub SadeceSayisalParametreler()
    Dim oParams As Object
    Dim oParam As Object
    Dim i As Long
    Dim dVal As Double
    Set oParams = GetObject(, "CATIA.Application").ActiveDocument.GetItem("Part").Parameters
    If oParams Is Nothing Then Exit Sub
    For i = 1 To oParams.Count
        Set oParam = oParams.Item(i)
        If Not oParam Is Nothing Then
            On Error Resume Next
            dVal = oParam.Value
            If Err.Number = 0 Then Debug.Print oParam.Name & " = " & dVal
            On Error GoTo 0
        End If
    Next i
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## On dokuzuncu örnek: Kısa kullanım yönergesi (3 satır – Finalize)

Her makronun tesliminde kullanıcıya verilecek 3 satırlık talimat (Help – Finalize):

1. **Part Design** veya **GSD** rolünde bir **Part** belgesi açın.  
2. **Tools → Macro → Run** ile bu makroyu çalıştırın.  
3. **“İşlem tamamlandı”** mesajını görünce işlem bitti; hata mesajı çıkarsa log dosyasını kontrol edin.

Bu metni makro başlığına veya ayrı “Kullanım” dosyasına ekleyin.

════════════════════════════════════════════════════════════════════════════════

## Yirminci örnek: Özet – Tüm adımlar tek listede

Baştan sona bir makro yazarken izlenecek adımların özeti:

1. **Uygulama al:** `GetObject(, "CATIA.Application")`; Nothing kontrolü.  
2. **Aktif belge al:** `ActiveDocument`; Nothing kontrolü.  
3. **Part/Product/Drawing al:** `GetItem("Part")` vb.; Nothing kontrolü.  
4. **Koleksiyon al:** Parameters, Shapes, Sheets, Children vb.; Nothing ve Count kontrolü.  
5. **Döngü veya tek erişim:** For/For Each ile Item(i) veya Item("isim").  
6. **Değer oku/yaz:** Property’ler (Value, Name vb.).  
7. **Güncelle:** Part.Update (döngü dışında bir kez).  
8. **Hata yakala:** On Error GoTo; Exit Sub; mesaj veya log.  
9. **Başlık:** Language, Release, Purpose, Assumptions.  
10. **Test:** F5, F8, Immediate penceresi; dağıtım öncesi kontrol listesi (11. doküman).

**Gezinme:** Önceki: [09-Hata-Yakalama](09-Hata-Yakalama-ve-Debug.md) | [Rehber listesi](README.md) | Sonraki: [11-Resmi-Kurallar](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md) →
