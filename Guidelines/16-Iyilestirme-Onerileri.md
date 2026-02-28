# 16. İyileştirme Önerileri

Makronuz çalıştıktan sonra **kalite**, **bakım kolaylığı** ve **kullanıcı deneyimi**ni artırmak için uygulayabileceğiniz öneriler aşağıda kategorilere ayrılmıştır. Zorunlu değildir; ihtiyaca göre seçin.

------------------------------------------------------------

## 1. Kod kalitesi

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Sihirli sayıları kaldırın** | Tekrarlayan veya anlamı belirsiz sayıları **Const** veya başta tanımlı değişkene alın. | `Const MAX_ITERATION As Long = 100`; döngüde `For i = 1 To MAX_ITERATION` |
| **Uzun Sub’ları bölün** | 50+ satırlık tek Sub yerine anlamlı adımları ayrı Sub/Function yapın; ana Sub sadece çağrıları sıralasın. | `ParametreleriOku`, `DosyayaYaz`, `KullaniciyaRaporla` |
| **Tekrarlayan kodu fonksiyona alın** | Aynı “Application → Part al” veya “log satırı yaz” bloğu iki yerde varsa ortak Function/Sub yazın. | `GetActivePart()`, `LogSatir(sMesaj)` |
| **Anlamlı isim kullanın** | Değişken ve prosedür adları ne yaptığını anlatsın; kısaltma yerine tam kelime tercih edin (okunabilirlik için). | `iParametreSayisi` yerine `iCount` sadece sayı için uygun; `sParametreAdi` anlamlı. |
| **Yorumu güncel tutun** | Kod değişince başlık yorumunu (Purpose, Assumptions) da güncelleyin; eski davranışı anlatan yorumları silin. | Her revizyonda `' REV 1.2 – 2025-03-01: ...` ekleyin. |

------------------------------------------------------------

## 2. Performans

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Tek Update kuralı** | Part/Product değişikliklerinde döngü içinde değil, tüm değişiklikler bittikten sonra **bir kez** `oPart.Update` çağırın. | 10. ve 11. dokümanlarda vurgulandı. |
| **Gereksiz erişimi azaltın** | Aynı koleksiyonu defalarca `Item(i)` ile almak yerine bir kez değişkene alıp kullanın. | `Set oParam = oParams.Item(i)` bir kez; sonra `oParam.Name`, `oParam.Value`. |
| **Büyük döngüde nesneleri serbest bırakın** | Çok sayıda eleman tarandığında döngü içinde geçici nesneyi **Set ... = Nothing** yapmak (bazı senaryolarda) bellek baskısını azaltır. | Özellikle 1000+ occurrence/shape tararken. |
| **Süreyi ölçün** | Uzun süren işlemlerde **Timer** ile süreyi ölçüp log’a veya Debug’a yazın; darboğazı tespit edin. | `t0 = Timer` … işlem … `Debug.Print "Süre: "; Timer - t0` |
| **Dosya/yol kontrolünü başta yapın** | Kaydetmeden önce hedef klasörün var olduğunu veya yazılabilir olduğunu kontrol edin; işlem ortasında hata almayın. | FileSystem.Exists veya API ile klasör kontrolü. |

------------------------------------------------------------

## 3. Bakım ve konfigürasyon

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Yolları tek yerde toplayın** | Log dosyası, çıktı klasörü, varsayılan dosya adı gibi değerleri modül başında **Const** veya tek bir “config” bloğunda toplayın. | `Const LOG_PATH As String = "C:\Temp\macro_log.txt"` |
| **Versiyon etiketi kullanın** | Her teslimde başlıkta veya ayrı satırda sürüm ve tarih yazın; hangi sürümün çalıştığı belli olsun. | `' -- REV 1.2 – 2025-03-01` |
| **Varsayılanları dokümante edin** | InputBox’ta varsayılan değer veriyorsanız, bu değerin ne anlama geldiğini yorumda belirtin (test ve bakım için). | `' Varsayılan: "Length.1" – test script’leri bu değeri kullanır.` |
| **Bağımlılıkları yazın** | Makronun hangi workbench’te, hangi belge türünde çalıştığını Assumptions’ta net yazın. | `' Assumptions: Part Design açık, aktif belge Part.` |

------------------------------------------------------------

## 4. Test ve hata senaryoları

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Boş / yanlış belge ile deneyin** | Hiç belge açık değilken, çizim açıkken (Part beklenirken) makroyu çalıştırın; Nothing ve mesajlar doğru mu kontrol edin. | ActiveDocument = Nothing, GetItem("Part") = Nothing |
| **Sınır değerleri test edin** | 0 parametre, 1 parametre, çok uzun parametre adı; kullanıcı iptal (boş InputBox) senaryolarını deneyin. | Count = 0, InputBox’ta Cancel. |
| **Hata numarasını mesajda gösterin** | Kullanıcıya sadece “Hata oluştu” değil, **Err.Number** ve **Err.Description** verin; destek ve log analizi kolaylaşır. | `MsgBox "Hata: " & Err.Number & " - " & Err.Description` |
| **Log’a bağlam ekleyin** | Log satırına hangi belge, hangi parametre veya adım bilgisi eklensin; sonradan hatayı tekrarlamak kolaylaşır. | `LogSatir "ERROR", 9100, "Parametre yok", "Doc=" & oDoc.Name` |

------------------------------------------------------------

## 5. Kullanıcı deneyimi

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Uzun işlemde bilgi verin** | 10 saniyeden uzun sürecek işlemde başta “İşlem başlıyor…” veya “X adet öğe işlenecek” mesajı gösterin. | MsgBox veya status bar (API varsa). |
| **İptal seçeneği** | Çok uzun döngülerde kullanıcıya “Devam?” sorusu veya iptal düğmesi sunun (mümkünse). | Her N iterasyonda bir MsgBox veya form ile Cancel. |
| **Başarı/hata mesajını net verin** | “Bitti” yerine “12 parametre güncellendi” veya “Hata: Length.1 bulunamadı” gibi somut ifade kullanın. | MsgBox içeriğini iş sonucuna göre doldurun. |
| **3 satırlık kullanım yönergesi** | Dağıtırken kullanıcıya “1) Şunu aç, 2) Makroyu çalıştır, 3) Şu mesajı görünce bitti” şeklinde kısa talimat verin. | 11. dokümanda Finalize maddesi. |

------------------------------------------------------------

## 6. Dağıtım ve güvenlik

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Sabit yolu dağıtım notuna yazın** | Log veya çıktı için kullandığınız varsayılan yolu başlıkta veya dağıtım notunda belirtin; farklı ortamda değiştirilmesi gerektiği belli olsun. | `' Dağıtım: Log yolu kuruluma göre LOG_PATH değiştirilmeli.` |
| **Hassas bilgi yazmayın** | Şifre, token vb. koda veya log dosyasına yazmayın; gerekirse ortam değişkeni veya harici config kullanın. | — |
| **Bölgesel ayarı belirtin** | Hangi dil/locale’de test edildiğini başlıkta yazın; farklı bölgede davranış değişebilir (Help uyarısı). | `' Regional Settings: English (United States)` |
| **Gerekli yetkiyi dokümante edin** | Makronun “Part yazma”, “ağ sürücüsüne kaydetme” gibi hangi yetkileri gerektirdiğini kısa not edin. | Assumptions veya ayrı “Gereksinimler” bölümü. |

------------------------------------------------------------

## 7. Hata yönetimi ve geri alma

| Öneri | Açıklama | Örnek |
|-------|----------|--------|
| **Kritik adımlarda rollback düşünün** | Birden fazla nesneyi değiştiriyorsanız, hata durumunda önceki değerleri geri yazma veya “kısmen uygulandı” uyarısı verin. | Tek parametre yazarken eski değeri saklayıp hata olursa geri yazma (10. dokümanda örnek). |
| **Err.Raise ile test edilebilir çıkış** | Otomatik test senaryolarında hatanın “hata” olarak tanınması için MsgBox yerine **Err.Raise** kullanın (Help önerisi). | `If ... Then Err.Raise 9001, "MacroAdi", "Açıklama"` |
| **On Error’ü dar tutun** | **On Error Resume Next** sadece tek satır veya çok kısa blok için kullanın; hemen sonra **Err** kontrolü ve **On Error GoTo 0** yapın. | 9. dokümanda örnek. |

------------------------------------------------------------

## 8. Özet kontrol listesi (isteğe bağlı)

Kodunuzu teslim etmeden veya paylaşmadan önce aşağıdakilerden ihtiyacınız olanları işaretleyebilirsiniz:

- [ ] Option Explicit ve Language/Release başlığı var.
- [ ] Tüm Set sonrası Nothing veya Count kontrolü yapılıyor.
- [ ] Döngü içinde Update yok; tek Update kullanılıyor.
- [ ] Sihirli sayılar Const veya anlamlı değişkene alındı.
- [ ] Uzun Sub’lar mantıklı Sub/Function’lara bölündü.
- [ ] Log veya mesajlara bağlam (belge adı, adım) ekleniyor.
- [ ] Boş/yanlış belge ve iptal senaryoları test edildi.
- [ ] Kullanıcıya 3 satırlık kullanım yönergesi verildi.
- [ ] Versiyon etiketi ve Assumptions güncel.
- [ ] Bölgesel ayar ve dağıtım notu (gerekirse) yazıldı.

------------------------------------------------------------

## İlgili dokümanlar

**Tüm rehber:** [README](README.md). İlgili: [11](11-Resmi-Kurallar-ve-Hazirlik-Fazlari.md) (zorunlu kurallar, TAMAM/HAZIR), [09](09-Hata-Yakalama-ve-Debug.md) (On Error, log), [10](10-Ornek-Proje-Bastan-Sona-Bir-Makro.md) (rollback, kod kalıpları).
