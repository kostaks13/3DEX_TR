# 2. Ortam Kurulumu

```
╔══════════════════════════════════════════════════════════════════════════════╗
║  Ortam: 3DExperience, VBA editörü, makro güvenliği, ilk makroyu yazıp        ║
║  çalıştırma                                                                  ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

3DExperience VBA ile kod yazabilmek için doğru ortamı kurmanız ve makro güvenlik ayarlarını yapmanız gerekir. Bu dokümanda adım adım ilerliyoruz.

**Bu dokümanda:** Ortam gereksinimleri; VBA editörü ve makro güvenliği; ilk makro (MsgBox, InputBox); Language/Release başlığı; F5/F8.

════════════════════════════════════════════════════════════════════════════════

## Gereksinimler

- **3DExperience** (R2024x veya uyumlu sürüm) kurulu ve lisansınızın **Makro / Scripting** yetkisi olmalı.  
- Windows üzerinde çalışır; VBA editörü 3DExperience ile birlikte gelir, ayrı kurulum gerekmez.

════════════════════════════════════════════════════════════════════════════════

## 1. 3DExperience’ı açın

- 3DExperience uygulamasını başlatın.  
- Bir **rol** seçin (ör. **Mechanical Designer**, **Assembly Designer**).  
- İhtiyacınıza göre yeni bir parça veya montaj açın veya mevcut bir belgeyi açın.

════════════════════════════════════════════════════════════════════════════════

## 2. VBA editörünü açın

- Menüden: **Tools** → **Macro** → **Edit Macro** (veya **Visual Basic Editor**).  
- Alternatif kısayol: **Alt + F11** (birçok sürümde geçerlidir).

Açılan pencerede:

- **Sol:** Proje ağacı (Project Explorer) — modüller, formlar listelenir.  
- **Orta:** Kod penceresi — burada VBA kodunu yazacaksınız.  
- **Alt:** Properties (özellikler) — seçili nesnenin özellikleri.

════════════════════════════════════════════════════════════════════════════════

## 3. Makro güvenliği

Makroların çalışması için güvenlik ayarı düşük veya orta olmalı (kurumsal politika izin veriyorsa).

- 3DExperience içinde: **Tools** → **Options** (veya **Settings**).  
- **General** / **Security** / **Macro** benzeri bir bölümde **Macro security** veya **Enable macros** seçeneğini açın.  
- Şirket politikası varsa IT ile “Makro çalıştırma” iznini kontrol edin.

════════════════════════════════════════════════════════════════════════════════

## 4. İlk makroyu yazıp çalıştırma

1. VBA editöründe **Insert** → **Module** ile yeni bir modül ekleyin.  
2. Açılan kod penceresine şunu yazın:

```vba
Sub IlkMakro()
    MsgBox "Merhaba, 3DExperience VBA!"
End Sub
```

3. İmleci `Sub IlkMakro()` satırının içine getirip **F5**’e basın veya **Run** → **Run Sub/UserForm**.  
4. 3DExperience penceresinde “Merhaba, 3DExperience VBA!” yazan bir mesaj kutusu görünür.

Bu, ortamınızın doğru kurulduğunu ve makronun çalıştığını gösterir.

════════════════════════════════════════════════════════════════════════════════

## 4b. Örnek 2 – InputBox ile kullanıcıdan veri almak

Kullanıcıdan metin veya sayı almak için **InputBox** kullanılır. Help’e göre test edilebilirlik için **üçüncü parametrede varsayılan değer** verin.

```vba
Option Explicit

Sub KullaniciAdiAl()
    Dim sAd As String
    ' Üçüncü parametre: varsayılan değer (test ortamında kullanılır)
    sAd = InputBox("Parça veya gövde adını girin:", "Giriş", "Parça_01")
    If sAd = "" Then
        MsgBox "İptal edildi."
        Exit Sub
    End If
    MsgBox "Girilen ad: " & sAd
End Sub
```

Çalıştırınca bir kutu açılır; kullanıcı metin yazıp OK’e basar. Boş bırakırsa `sAd = ""` olur; buna göre iptal mesajı gösterebilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## 4c. Örnek 3 – Birkaç değişken ve mesaj birleştirme

Sayı ve metni birleştirip tek mesajda göstermek için `&` kullanın. Sayıyı metne çevirmek için `CStr` veya `&` ile birleştirmek yeterlidir.

```vba
Option Explicit

Sub MesajBirlestir()
    Dim sUstBilgi As String
    Dim iSayi As Long
    Dim dOndalik As Double

    sUstBilgi = "3DExperience VBA"
    iSayi = 42
    dOndalik = 3.14159

    MsgBox "Başlık: " & sUstBilgi & vbCrLf & _
           "Tamsayı: " & iSayi & vbCrLf & _
           "Ondalık: " & dOndalik
End Sub
```

`vbCrLf` satır sonu karakteridir; mesaj kutusunda birkaç satır göstermek için kullanılır.

════════════════════════════════════════════════════════════════════════════════

## 4d. Örnek 4 – Başlık yorumu ile tam mini makro

Resmi kurallara uygun kısa bir makro: başlık, Option Explicit, Language/Release, tek Sub.

```vba
Option Explicit
' Purpose: Kullanıcıya merhaba deyip makro adını gösterir.
' Assumptions: Yok.
' Language: VBA
' Release:  3DEXPERIENCE R2024x

Sub MerhabaVeBilgi()
    Dim sMesaj As String
    sMesaj = "Merhaba." & vbCrLf & "Çalışan makro: MerhabaVeBilgi"
    MsgBox sMesaj
End Sub
```

Bu yapıyı alışkanlık edinin: Her Sub’ın üstüne amaç ve varsayımları kısa yazın.

════════════════════════════════════════════════════════════════════════════════

## 5. Makroyu kaydetme

- VBA editöründe **File** → **Save**.  
- Proje, 3DExperience oturumunuzla birlikte (veya açık olan belgeyle) kaydedilir.  
- Bazı sürümlerde makrolar **.CATPart** / **.CATProduct** içinde veya ayrı bir **.vba** projesi olarak saklanır; menüdeki “Save” ile kaydedilen yeri not alın.

════════════════════════════════════════════════════════════════════════════════

## 6. Resmi kurallardan: Dil ve sürüm bildirimi (Help referansı)

**Help-Automation Development Guidelines**’a göre teslim edilen script’lerde aşağıdaki yorumlar **zorunludur**:

- **Language** — Geçerli değerler: `VBScript`, `CATScript`, `VBA`, `VB.NET`, `C#`.
- **Release** — İlk desteklenen sürümü tanımlar (örn. V6R2010, 3DEXPERIENCE R2024x).

Örnek:

```vba
' Language: VBA
' Release:  3DEXPERIENCE R2024x
```

Bunu modülünüzün en üstüne (Option Explicit’ten hemen sonra veya başlık bloğu içinde) ekleyin. Böylece hangi sürümde test edildiği ve hangi dilde yazıldığı kayıtlı olur; API değişikliklerinde referans olur.

════════════════════════════════════════════════════════════════════════════════

## 7. Makro konumu ve dağıtım (Hazırlık Yönergesi’nden)

**3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ** dokümanında “Finalize” fazında şu not yer alır:

- Makro dağıtımı için konum örneği: **%CATStartupPath%\Macros\** (veya kurulumunuzdaki makro kök dizini).
- Dağıtım notunu kod başlığına veya ayrı bir “Dağıtım” bölümüne yazın: “Bu makro şu klasöre kopyalanacak: …”
- Kullanıcıya **3 satırlık kullanım** yönergesi verin: Örn. “GSD aktifken çalıştır → Part güncel → ‘Done’ mesajı görünce işlem tamamlandı.”

════════════════════════════════════════════════════════════════════════════════

## 8. Örnek: Modül başına zorunlu başlık (Language + Release)

Her yeni modül açtığınızda en üste şu iki satırı ekleyin; böylece hangi dil ve sürüm için yazıldığı belli olur:

```vba
Option Explicit
' Language: VBA
' Release:  3DEXPERIENCE R2024x

Sub OrnekSub()
    MsgBox "Bu modül VBA ve R2024x için yazıldı."
End Sub
```

════════════════════════════════════════════════════════════════════════════════

## 9. Örnek: Proje Explorer’da modül adı değiştirme

VBA editöründe sol tarafta **Project Explorer** görünür. Modüle sağ tıklayıp **Properties** (F4) ile **Name** alanından modül adını değiştirebilirsiniz (ör. Module1 → ParametreMakrolari). Bu, büyük projelerde modülleri ayırt etmeyi kolaylaştırır.

════════════════════════════════════════════════════════════════════════════════

## 10. Örnek: Makroyu araç çubuğuna veya menüye ekleme

3DExperience’ta **Tools** → **Customize** (veya **Options**) ile yeni bir düğme ekleyip makroyu atayabilirsiniz. Böylece kullanıcı **Alt+F11** açmadan doğrudan düğmeye basarak makroyu çalıştırır. Adımlar sürüme göre değişir; menüde “Macro”, “Command” veya “Script” benzeri bir bölüm arayın.

════════════════════════════════════════════════════════════════════════════════

## 11. Örnek: F5 vs F8 – Çalıştır vs Adım adım

- **F5:** Makroyu baştan sona çalıştırır.  
- **F8 (Step Into):** Satır satır ilerler; her F8’de bir satır çalışır. Debug için kullanılır.  
- **Shift+F8 (Step Over):** Alt prosedürü tek adımda geçer.  
- Kesme noktası koyduğunuzda F5 ile o satıra kadar çalışır, orada durur; sonra F8 ile devam edebilirsiniz.

════════════════════════════════════════════════════════════════════════════════

## 12. Örnek: Dağıtım notu – Kod başlığına ekleme

Help’teki “Finalize” fazında dağıtım konumunu kod başlığına yazın:

```vba
' Dağıtım: Bu makro %CATStartupPath%\Macros\ klasörüne kopyalanacaktır.
' Kullanım: Part Design veya GSD açık, Part belgesi aktif; makroyu çalıştırın.
```

════════════════════════════════════════════════════════════════════════════════

## 13. Örnek: Birden fazla modül – Proje yapısı

Büyük projelerde modülleri işlevine göre ayırabilirsiniz: bir modül “ParametreIslemleri”, bir modül “DosyaVeLog”, bir modül “AnaMakrolar”. AnaMakrolar’daki Sub’lar diğer modüllerdeki Public Sub/Function’ları çağırır. Aynı proje içinde oldukları sürece modül adı ile çağrı gerekmez; doğrudan SubAdi veya FunctionAdi yeterlidir.

════════════════════════════════════════════════════════════════════════════════

════════════════════════════════════════════════════════════════════════════════

## Uygulamalı alıştırma – Yaparak öğren

**Amaç:** İlk makroyu sıfırdan yazıp çalıştırmak; F5 ve F8 ile farkı görmak.  
**Süre:** Yaklaşık 10 dakika.

| Adım | Ne yapacaksınız | Kontrol |
|------|------------------|--------|
| **1** | 3DExperience’ı açın; bir rol seçin (örn. Mechanical Designer). | Uygulama açık mı? |
| **2** | **Alt+F11** (veya Tools → Macro → Edit Macro) ile VBA editörünü açın. | Sol tarafta proje ağacı görünüyor mu? |
| **3** | **Insert → Module** ile yeni modül ekleyin. Kod penceresine `Option Explicit` yazın, alt satıra `Sub IlkMakrom()` yazın (bilerek yanlış: sonunda fazladan `m`). | — |
| **4** | **F5** ile çalıştırmayı deneyin. Makro listesinde “IlkMakrom” görünür; çalıştırınca boş bir makro çalışır. Sonra Sub adını `IlkMakrom` → `IlkMakro` yapıp içine `MsgBox "Merhaba"` ekleyin. | Sub adını düzelttiniz mi? |
| **5** | Sub adını `IlkMakrom` → `IlkMakro` yapıp içine `MsgBox "Merhaba, 3DExperience VBA!"` yazın. **F5** ile çalıştırın. | Mesaj kutusu “Merhaba, 3DExperience VBA!” gösteriyor mu? |
| **6** | İmleci `MsgBox` satırına getirip **F8**’e iki kez basın (Step Into). Satırın sarı ile vurgulandığını görün; F8 ile tek tek ilerleyin. | Adım adım çalıştığını gördünüz mü? |
| **7** | **File → Save** ile projeyi kaydedin. | Proje kaydedildi mi? |

**Beklenen sonuç:** Mesaj kutusunda “Merhaba, 3DExperience VBA!” görünmeli; F8 ile satır satır ilerleme deneyimlendi.

════════════════════════════════════════════════════════════════════════════════

## Kontrol listesi

- [ ] 3DExperience açılıyor ve bir rol seçiliyor  
- [ ] **Alt+F11** veya **Tools → Macro → Edit Macro** ile VBA editörü açılıyor  
- [ ] Yeni modül eklenip `MsgBox "Merhaba..."` yazılabiliyor  
- [ ] **F5** ile makro çalışıyor ve mesaj kutusu görünüyor  
- [ ] Proje kaydedildi  

════════════════════════════════════════════════════════════════════════════════

## Sonraki adım

**3. doküman:** [03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md](03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md) — Değişkenler, veri tipleri ve `Option Explicit` kullanımı.

**Gezinme:** Önceki: [01-Giris](01-Giris-Neden-3DExperience-VBA.md) | [Rehber listesi](README.md) | Sonraki: [03-Degiskenler](03-VBA-Temelleri-Degiskenler-ve-Veritipleri.md) →
