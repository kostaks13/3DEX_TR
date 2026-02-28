# 19. İsimlendirme Rehberi

**Bu dokümanda:** Değişken, sabit, Sub/Function, modül, dosya, parametre ve etiket isimlendirme kuralları tek yerde toplanmıştır. Help (Automation Development Guidelines) ve rehberdeki dağınık maddeler burada özetlenip genişletilmiştir.

**İlgili:** [03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md](03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md) (önekler), [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md) (Help kuralları), [VBA-Kod-Checklist.md](VBA-Kod-Checklist.md) (bölüm 5 ve 13).

------------------------------------------------------------

## 1. Özet tablo

| Ne | Kural | Örnek |
|----|--------|--------|
| **Değişken** | Tip öneki + anlamlı ad; mixed case (ilk harfler büyük). | `oPart`, `sPartName`, `iCount`, `dLength`, `bReadOnly`, `cParams` |
| **Sabit (Const)** | Tamamı BÜYÜK_HARF, kelimeler alt çizgi ile. | `MAX_ITERATION`, `LOG_PATH`, `ERR_PARAM_NOT_FOUND` |
| **Sub / Function** | Fiil veya fiil ifadesi; her kelimenin ilk harfi büyük (PascalCase). | `GetActivePart`, `UpdateParameterValue`, `ParametreAdlariniGetir` |
| **Modül adı** | İşlevi anlatan isim; boşluksuz, kelime başları büyük. | `ParametreIslemleri`, `DosyaVeLog`, `AppHelpers` |
| **.bas dosyası** | Proje/Modül adı ile uyumlu; tire veya alt çizgi ile. | `Acme_ParametreListesi.bas`, `Common_AppHelpers.bas` |
| **Parametre (Sub/Function)** | Anlamlı ad; önek kullanılabilir (ByVal sAd As String). | `sParamName`, `oPart`, `iMaxCount` |
| **Hata etiketi** | Anlamlı, kısa; sonunda iki nokta. | `HataYakala:`, `CikisTemiz:` |
| **Özel hata sabiti** | ERR_ veya benzeri önek; 9000–9999 değeri. | `Const ERR_PARAM_NOT_FOUND As Long = 9001` |

------------------------------------------------------------

## 2. Değişken isimlendirme

### 2.1 Resmi önekler (Help – Automation Development Guidelines)

Tipi hemen anlamak için değişken adının **başına tek harf** ekleyin:

| Önek | Tip | Örnek |
|------|-----|--------|
| **b** | Boolean | bReadOnly, bIsUpToDate, bFound |
| **d** | Double | dLength, dAngle, dTolerance |
| **s** | String | sPartName, sFilePath, sLogPath |
| **i** | Integer / Long | iCount, iIndex, iNumberOfShapes |
| **o** | Object | oPart, oDoc, oApp, oParam |
| **c** | Collection | cParams, cShapes, cChildren |

- **Integer vs Long:** Sayac ve indeks için genelde **Long** kullanın (Count vb. Long döner); Help’teki “i” öneki her iki tip için de kullanılır. Büyük döngülerde Long tercih edin.
- Önek sonrası kelimeler **mixed case**: ilk harf büyük (`sPartName`, `oActiveDocument`).

### 2.2 Ek önekler (isteğe bağlı)

Rehberde sık geçmeyen ancak tutarlı kullanırsanız faydalı olabilecek önekler:

| Önek | Tip | Örnek |
|------|-----|--------|
| **dt** veya **d** | Date | dtStart, dToday |
| **v** | Variant (mümkünse az kullanın) | vTemp |

### 2.3 Kaçınılacaklar

- **Anlamsız kısaltma:** `x`, `tmp`, `obj1` yerine `iIndex`, `oPart`, `sTempPath` gibi anlamlı isim.
- **Tek harf (döngü dışında):** Sadece döngü sayacı için `i`, `j` kabul edilebilir; diğer değişkenlerde tam isim kullanın.
- **Türkçe karakter:** VBA’da çalışır ancak locale ve aktarımda sorun çıkarabilir; İngilizce tercih edin (sPartName, oParca yerine oPart).

------------------------------------------------------------

## 3. Sabit (Const) isimlendirme

- **Tamamı büyük harf**, kelimeler **alt çizgi** ile ayrılır.
- Anlamlı önek kullanılabilir: `LOG_`, `MAX_`, `DEFAULT_`, `ERR_`.

```vba
Const MAX_ITERATION As Long = 100
Const LOG_PATH As String = "C:\Temp\macro_log.txt"
Const DEFAULT_TOLERANCE As Double = 0.001
Public Const ERR_PARAM_NOT_FOUND As Long = 9001
```

------------------------------------------------------------

## 4. Sub ve Function isimlendirme

- **Fiil** veya **fiil ifadesi** (ne yaptığı belli olsun).
- **PascalCase:** Her kelimenin ilk harfi büyük.

| Tür | Örnek |
|-----|--------|
| Get/Set | GetActivePart, GetParameterValue, SetParameterValue |
| İşlem | UpdateParameters, ExportToExcel, LogSatir |
| Kontrol / sorgu | ParametreVarMi, BelgePartMi |
| Başlat / bitir | Baslat, Temizle (veya Initialize, Cleanup) |

- **Function** dönüşü ima edebilir: `ParametreAdlariniGetir`, `ToplamParcaSayisi`.
- **Private** yardımcı prosedürler de aynı kurala uysun; “internal” anlamı için önek zorunlu değildir.

------------------------------------------------------------

## 5. Modül ve dosya isimlendirme

### 5.1 VBA modül adı (Project Explorer’da)

- Boşluk yok; kelime başları büyük (PascalCase) veya tümü bitişik.
- İşlevi yansıtsın: `ParametreIslemleri`, `DosyaVeLog`, `AnaMakrolar`, `AppHelpers`.

### 5.2 .bas dosya adı (dışa aktarılan / paylaşılan)

- Tutarlı kural: **ProjeAdi_Islev.bas** veya **ModulAdi.bas**.
- Geçersiz karakter kullanmayın; tire veya alt çizgi ile kelimeleri ayırın.
- Örnek: `Acme_ParametreListesi.bas`, `Common_AppHelpers.bas`, `ParametreYaz.bas`.

------------------------------------------------------------

## 6. Parametre (argüman) isimlendirme

- Sub/Function parametrelerinde de **anlamlı isim** ve isteğe bağlı **önek** kullanın.
- ByVal ile geçirilen nesne parametreleri genelde `o` ile başlar: `oPart`, `oDoc`.

```vba
Function ParametreVarMi(oPart As Object, sParamName As String) As Boolean
Sub LogSatir(sMesaj As String, Optional iSeviye As Long = 0)
```

------------------------------------------------------------

## 7. Hata etiketi ve özel hata sabitleri

- **On Error GoTo** için etiket: kısa, anlamlı, sonunda iki nokta. Örn: `HataYakala:`, `CikisTemiz:`.
- **Özel hata numaraları (9000–9999):** Sabitleri tek yerde tanımlayın; isim **ERR_** veya benzeri önek ile başlasın.

```vba
Const ERR_NO_ACTIVE_DOC As Long = 9000
Const ERR_PARAM_NOT_FOUND As Long = 9001
Const ERR_INVALID_INPUT As Long = 9002
```

------------------------------------------------------------

## 8. Kısa kontrol listesi (isimlendirme)

- [ ] Değişkenlerde resmi önek (b, d, s, i, o, c) kullanılıyor.
- [ ] Sabitler BÜYÜK_HARF_ALT_CIZGI.
- [ ] Sub/Function adları fiil veya fiil ifadesi; PascalCase.
- [ ] Modül ve .bas dosya adları tutarlı; anlamlı.
- [ ] Parametreler anlamlı; nesne parametrelerinde o öneki (isteğe bağlı).
- [ ] Hata etiketleri kısa ve anlamlı; özel hata numaraları Const ile tanımlı.

════════════════════════════════════════════════════════════════════════════════

## Uygulamalı alıştırma – Yaparak öğren

**Amaç:** Mevcut bir makrodaki isimleri bu rehbere göre gözden geçirmek.  
**Süre:** Yaklaşık 10 dakika.  
**Zorluk:** Orta

| Adım | Ne yapacaksınız | Kontrol |
|------|------------------|--------|
| **1** | Yazdığınız bir makroyu açın. Değişken listesine bakın (Dim satırları). Önekler (b, d, s, i, o, c) kullanılmış mı? Eksikse en az 2 değişkene uygun önek ekleyin (örn. `ad` → `sAd`, `sayi` → `iSayi`). | Önekler uygulandı mı? |
| **2** | Const kullandıysanız isimleri BÜYÜK_HARF_ALT_CIZGI yapın (örn. `logPath` → `LOG_PATH`). Sub/Function adlarında fiil ve mixed case var mı? (örn. `GetActivePart`, `UpdateParameterValue`). | Const ve Sub/Function isimleri rehbere uygun mu? |
| **3** | “Kaçınılacaklar” listesini okuyun. Makronuzda anlamsız kısaltma (x, tmp, obj1), tek harf (döngü dışında) veya gereksiz Variant var mı? Varsa birini düzeltin. | Kaçınılacaklar temizlendi mi? |
| **4** | Hata etiketi kullanıyorsanız (HataYakala:) kısa ve anlamlı mı? Sonunda `:` olduğundan emin olun. | Etiket isimlendirmesi doğru mu? |

**Beklenen sonuç:** En az bir isimlendirme iyileştirmesi yapıldı; özet tablo ve kaçınılacaklar kontrol edildi.

------------------------------------------------------------

## İlgili dokümanlar

**Tüm rehber:** [README](README.md). **Değişkenler ve önekler:** [03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md](03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md). **Resmi kurallar:** [11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md). **Checklist (bölüm 5, 13):** [VBA-Kod-Checklist.md](VBA-Kod-Checklist.md).

---

### Gezinme

| [← Önceki: 18 Sık hatalar](18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) | [Rehber listesi](README.md) | Rehber sonu |
| :--- | :--- | :--- |
