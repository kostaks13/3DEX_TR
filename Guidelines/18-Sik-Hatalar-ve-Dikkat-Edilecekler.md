# 18. Sık Yapılan Hatalar ve Dikkat Edilmesi Gereken Özel Noktalar

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  Sık yapılan hatalar + dikkat edilmesi gereken özel noktalar (3DExperience VBA)║
╚══════════════════════════════════════════════════════════════════════════════╝
```

**Bu dokümanda:** Sık yapılan hatalar (Option Explicit, Nothing, Update, V5 API, On Error, InputBox iptal…); dikkat edilecek özel noktalar (locale, servis sırası…); özet tablo.

Bu dokümanda 3DExperience VBA makrolarında **sık karşılaşılan hatalar** ve **özel dikkat gerektiren noktalar** toplu halde listelenir. Rehberin diğer bölümlerinde (özellikle 9, 11, checklist) dağınık geçen kurallar burada tek yerde toplanmıştır.

════════════════════════════════════════════════════════════════════════════════

## 1. Sık yapılan hatalar

### 1.1 Option Explicit unutmak

**Hata:** Modülün başına `Option Explicit` yazılmaz; yanlış yazılan değişken adı yeni bir Variant olarak kabul edilir, derleme hatası vermez, çalışma anında mantık hatası oluşur.

**Doğrusu:** Her modülün **ilk satırı** `Option Explicit` olsun.

```vba
Option Explicit   ' <- Mutlaka ilk satır

Sub Ornek()
    Dim sAd As String
    sAd = "Test"
    MsgBox sAd
End Sub
```

════════════════════════════════════════════════════════════════════════════════

### 1.2 On Error kullanmamak veya tüm makro boyunca Resume Next

**Hata:** Hiç `On Error` yok; uygulama kapalı veya belge yokken makro çöküp kullanıcıya anlamsız bir hata kutusu gösterir. Veya `On Error Resume Next` makro başında açılıp hiç kapatılmaz; tüm hatalar yutulur.

**Doğrusu:** En az bir kez `On Error GoTo HataYakala` kullanın; `On Error Resume Next` sadece tek satır / kısa blok için, hemen ardından `Err` kontrolü ve `On Error GoTo 0` yapın.

```vba
On Error GoTo HataYakala
' ... ana kod ...
Exit Sub
HataYakala:
    MsgBox "Hata: " & Err.Number & " - " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

### 1.3 Nothing / Count kontrolü yapmamak

**Hata:** `Set oDoc = oApp.ActiveDocument` veya `Set oPart = oDoc.GetItem("Part")` sonrası kontrol yok; belge kapalı veya yanlış türdeyse bir sonraki satırda "Object variable not set" (424) hatası.

**Doğrusu:** Her kritik `Set` sonrası `If ... Is Nothing Then` (ve gerekiyorsa `Count = 0`) kontrolü yapın; anlamlı mesaj verip `Exit Sub` / `Exit Function` yapın.

```vba
Set oDoc = oApp.ActiveDocument
If oDoc Is Nothing Then
    MsgBox "Açık belge yok. Önce bir parça veya montaj açın."
    Exit Sub
End If
Set oPart = oDoc.GetItem("Part")
If oPart Is Nothing Then Set oPart = oDoc
If oPart Is Nothing Then
    MsgBox "Bu belge Part değil."
    Exit Sub
End If
```

════════════════════════════════════════════════════════════════════════════════

### 1.4 Döngü içinde Part.Update / Update çağırmak

**Hata:** Parametreleri döngüyle güncellerken her iterasyonda `oPart.Update` çağrılır; hem yavaş hem tutarsızlık riski.

**Doğrusu:** Tüm değişiklikler bittikten **sonra**, döngü **dışında** **tek bir** `oPart.Update` (veya eşdeğeri) çağırın.

```vba
For i = 1 To oParams.Count
    Set oParam = oParams.Item(i)
    ' ... oParam.Value = ...
Next i
oPart.Update   ' <- Sadece bir kez, döngü dışında
```

════════════════════════════════════════════════════════════════════════════════

### 1.5 Eski V5 API kullanmak

**Hata:** Kayıttan kalan veya eski örneklerden kopyalanan `Documents.Add`, `HybridShapeFactoryOld`, `Selection.Search` vb. 3DExperience’ta desteklenmez veya farklı API ile değiştirilmesi gerekir.

**Doğrusu:** Yeni belge için **PLMNewService** / **PLMOpenService**; geometri için güncel **HybridShapeFactory**; arama için **SearchService** kullanın. API adlarını **VBA_API_REFERENCE.md** ve **Help-Native Apps Automation** ile kontrol edin.

════════════════════════════════════════════════════════════════════════════════

### 1.6 InputBox iptal (Cancel) kontrolü yapmamak

**Hata:** Kullanıcı InputBox’ta Cancel’a basınca boş string döner; kod boş değeri parametre adı veya yol olarak kullanırsa hata veya beklenmeyen davranış.

**Doğrusu:** InputBox’tan dönen değeri kontrol edin; boşsa iptal mesajı gösterip çıkın.

```vba
sAd = InputBox("Parametre adı girin:", "Giriş", "Length.1")
If sAd = "" Then
    MsgBox "İptal edildi."
    Exit Sub
End If
```

════════════════════════════════════════════════════════════════════════════════

### 1.7 Koleksiyonda Count = 0 kontrolü yapmamak

**Hata:** `oParams.Item(1)` veya `oShapes.Item(1)` çağrılır ama koleksiyon boş; "Invalid index" veya benzeri hata.

**Doğrusu:** Koleksiyona erişmeden önce `Is Nothing` ve **mümkünse** `Count > 0` kontrolü yapın.

```vba
If oParams Is Nothing Then Exit Sub
If oParams.Count = 0 Then
    MsgBox "Parametre yok."
    Exit Sub
End If
Set oParam = oParams.Item(1)
```

════════════════════════════════════════════════════════════════════════════════

### 1.8 Exit Sub unutmak (HataYakala’ya düşme)

**Hata:** `On Error GoTo HataYakala` kullanılıyor ama normal başarılı çıkışta `Exit Sub` yok; kod akışı her durumda `HataYakala:` etiketine de gelir ve "Hata: 0" gibi yanıltıcı mesaj gösterilir.

**Doğrusu:** Hata yakalama etiketinden **önce** mutlaka `Exit Sub` veya `Exit Function` yazın.

```vba
Sub Ornek()
    On Error GoTo HataYakala
    ' ... işlemler ...
    MsgBox "Tamamlandı."
    Exit Sub        ' <- Olmazsa HataYakala'ya düşer
HataYakala:
    MsgBox "Hata: " & Err.Description
End Sub
```

════════════════════════════════════════════════════════════════════════════════

### 1.9 Sihirli sayılar ve sabit yolları koda gömmek

**Hata:** `For i = 1 To 100`, `"C:\Temp\log.txt"` gibi değerler doğrudan kod içinde; değişmesi gerektiğinde birçok yeri değiştirmek gerekir.

**Doğrusu:** Sayı ve yolları modül başında **Const** veya anlamlı değişkenle tanımlayın.

```vba
Const MAX_ITER As Long = 100
Const LOG_PATH As String = "C:\Temp\macro_log.txt"
```

════════════════════════════════════════════════════════════════════════════════

### 1.10 Hassas bilgiyi koda veya log’a yazmak

**Hata:** Şifre, token veya kişisel veri koda veya log dosyasına yazılır; güvenlik ihlali.

**Doğrusu:** Hassas bilgi koda ve log’a **yazmayın**; gerekirse ortam değişkeni veya harici güvenli config kullanın.

════════════════════════════════════════════════════════════════════════════════

### 1.11 GetItem("Part") sonrası doğrudan Part varsaymak

**Hata:** Çizim veya montaj açıkken `GetItem("Part")` Nothing dönebilir; kod `oPart.Parameters` çağırırsa hata.

**Doğrusu:** `GetItem("Part")` sonrası Nothing kontrolü; parça belgesi değilse `Set oPart = oDoc` deneyin (sadece Part belgesiyse geçerlidir) veya kullanıcıya "Bu belge Part değil" deyip çıkın.

════════════════════════════════════════════════════════════════════════════════

### 1.12 Bölgesel ayar (locale) farkını hesaba katmamak

**Hata:** Makro bir dilde (örn. İngilizce) yazılıp başka locale’de (örn. Türkçe) çalıştırılınca sayı/ tarih formatı veya API davranışı farklı olabilir (Help uyarısı).

**Doğrusu:** Başlıkta **Regional Settings** belirtin; makroyu paylaşırken hangi locale’de test edildiğini yazın. Mümkünse sayı/tarih için tutarlı format kullanın.

════════════════════════════════════════════════════════════════════════════════

## 2. Dikkat edilmesi gereken özel noktalar

### 2.1 Tek Update kuralı (Part / Product)

**Neden önemli:** Part veya Product’ta birden fazla değişiklik yapıyorsanız, **hepsini yaptıktan sonra** **bir kez** `Update` çağırın. Döngü içinde her seferinde Update hem performansı düşürür hem bazı senaryolarda tutarsızlığa yol açabilir.

**Özet:** Değişiklik döngüsü **dışında** tek `oPart.Update` (veya eşdeğeri).

════════════════════════════════════════════════════════════════════════════════

### 2.2 Editor-level ve Session-level servis sırası

**Neden önemli:** Editor-level servis (ör. InertiaService) kullanırken **içeriden** session-level servis (ör. SearchService) çağırırsanız kilitleşme riski vardır (Help uyarısı).

**Özet:** Önce editor-level işinizi bitirin; sonra session-level servise geçin. Aynı akışta ikisini iç içe çağırmayın.

════════════════════════════════════════════════════════════════════════════════

### 2.3 FileSystem: Taşınabilirlik

**Neden önemli:** Windows’a özel `Scripting.FileSystemObject` veya sabit `C:\` yolları farklı ortamda (farklı dil, sanal makine) çalışmayabilir.

**Özet:** Mümkünse **Application.FileSystem** ve **SystemService.ConcatenatePaths** kullanın; yol birleştirmeyi sabit `\` ile yapmayın. Help’e göre taşınabilir script için bu önerilir.

════════════════════════════════════════════════════════════════════════════════

### 2.4 Windows API (GetOpenFileName / GetSaveFileName) ve 32/64 bit

**Neden önemli:** 64 bit VBA’da API `Declare` satırında **PtrSafe** ve **LongPtr** kullanılmazsa derleme veya çalışma hatası olur.

**Özet:** `#If VBA7 Then ... #Else ... #End If` ile 64 bit için `Declare PtrSafe` ve `LongPtr` kullanın. 32 bit ortamda `LongPtr` yerine **Long** gerekebilir.

════════════════════════════════════════════════════════════════════════════════

### 2.5 GetObject vs CreateObject (Application)

**Neden önemli:** 3DExperience **zaten açıkken** `GetObject(, "CATIA.Application")` ile mevcut oturuma bağlanırsınız. `CreateObject` yeni bir uygulama örneği açar; genelde makro içinden çalışırken **GetObject** kullanılır.

**Özet:** Uygulama açıkken **GetObject**; kapalıysa (ve bazen otomasyon senaryosunda) **CreateObject**. GetObject başarısız olursa (uygulama yok) Nothing döner; mutlaka kontrol edin.

════════════════════════════════════════════════════════════════════════════════

### 2.6 ActiveDocument ve belge türü

**Neden önemli:** Kullanıcı çizim veya montaj açıkken makro Part bekliyorsa `GetItem("Part")` Nothing döner. Hiç belge açık değilse `ActiveDocument` Nothing’dir.

**Özet:** Her zaman **ActiveDocument Is Nothing** ve (Part/Product/Drawing kullanıyorsanız) **doğru belge türü** kontrolü yapın; yanlış türde anlamlı mesaj verip çıkın.

════════════════════════════════════════════════════════════════════════════════

### 2.7 Koleksiyon indeksleri: 1’den başlama

**Neden önemli:** 3DExperience VBA koleksiyonlarında **Item(1)** çoğu yerde ilk elemandır (0 değil). Count = 0 ise Item(1) çağrılmamalı.

**Özet:** Döngüde `For i = 1 To col.Count` kullanın; `Count = 0` ise döngüye girmeyin veya erken çıkın.

════════════════════════════════════════════════════════════════════════════════

### 2.8 Hata mesajında Err.Number ve Err.Description

**Neden önemli:** Sadece "Hata oluştu" yazmak destek ve log incelemesini zorlaştırır.

**Özet:** Kullanıcıya veya log’a **Err.Number** ve **Err.Description** yazın; mümkünse hangi adımda (belge adı, parametre adı vb.) hata oluştuğunu ekleyin.

════════════════════════════════════════════════════════════════════════════════

### 2.9 Kayıt edilmiş kodu olduğu gibi bırakmak

**Neden önemli:** Kayıt gereksiz satırlar, sabit indeksler, bazen eski V5 API üretir; Option Explicit ve Nothing kontrolleri eklenmemiştir.

**Özet:** Kayıt sonrası **mutlaka** sadeleştirin: Option Explicit ekleyin, değişkenleri tipleyin, her kritik Set sonrası Nothing/Count kontrolü ekleyin, döngü içinde Update varsa dışarı alın, V5 API’leri referansla değiştirin (7. ve 11. doküman).

════════════════════════════════════════════════════════════════════════════════

### 2.10 Servis tanımlayıcı (identifier) yazımı

**Neden önemli:** `GetSessionService("VisuServices")` gibi çağrıda string **tam olarak** Help’teki Service Identifier ile aynı olmalıdır; büyük/küçük harf ve yazım hatası servisin Nothing dönmesine neden olur.

**Özet:** Servis adını **Help-Common Services** veya **Help-Native Apps Automation** içindeki tablolardan kopyalayın; sürüm/rol bazı servisleri kapatabilir.

════════════════════════════════════════════════════════════════════════════════

## 3. Özet tablo: Hata / Dikkat noktası → Ne yapmalı?

| Konu | Ne yapmalı? |
|------|-------------|
| Option Explicit | Modülün ilk satırına yaz. |
| On Error | GoTo etiket veya Resume Next + kısa blok + Err kontrolü + GoTo 0. |
| Nothing | Her kritik Set sonrası If ... Is Nothing Then kontrolü. |
| Update | Part/Product’ta sadece **bir kez**, döngü dışında. |
| V5 API | Kullanma; PLMNewService, güncel Factory, SearchService kullan. |
| InputBox iptal | Dönüş = "" ise iptal say; mesaj ver, Exit Sub. |
| Boş koleksiyon | Count = 0 kontrolü; Item(1) çağırma. |
| Exit Sub | Hata etiketinden önce mutlaka yaz. |
| Sabitler | Const ve anlamlı değişken; sihirli sayı/yol kullanma. |
| Hassas bilgi | Koda ve log’a yazma. |
| GetItem("Part") | Sonrası Nothing kontrolü; Part değilse uyar. |
| Regional Settings | Başlıkta belirt; paylaşırken locale yaz. |
| Editor/Session servis | İç içe çağırma; önce editor işi bitir. |
| FileSystem | Mümkünse Application.FileSystem, ConcatenatePaths. |
| GetObject | Uygulama açıkken; sonrası Nothing kontrolü. |
| Hata mesajı | Err.Number ve Err.Description yaz. |
| Kayıt sonrası | Sadeleştir, Option Explicit, Nothing, Update, V5 kontrolü. |
| Servis adı | Help’teki identifier’ı birebir kullan. |

════════════════════════════════════════════════════════════════════════════════

## İlgili dokümanlar

**Hata yakalama ve debug:** [09-Hata-Yakalama-ve-Debug.md](09-Hata-Yakalama-ve-Debug.md). **Resmi kurallar ve kontrol listesi:** [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md). **Detaylı checklist:** [VBA-Kod-Checklist.md](VBA-Kod-Checklist.md). **İyileştirme önerileri:** [16-Iyilestirme-Onerileri.md](16-Iyilestirme-Onerileri.md). **Help dosyalarını kullanma:** [17-Help-Dosyalarini-Kullanma.md](17-Help-Dosyalarini-Kullanma.md). **Tüm rehber:** [README](README.md).

**Gezinme:** Önceki: [17-Help-Dosyalari](17-Help-Dosyalarini-Kullanma.md) | [Rehber listesi](README.md) | Sonraki: [19-Isimlendirme-Rehberi](19-Isimlendirme-Rehberi.md)
