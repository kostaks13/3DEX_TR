# Sorun Giderme (Troubleshooting)

Makro yazarken sık karşılaşılan hata veya davranış senaryoları ve olası nedenler / adımlar. Detay için [Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md) ve [Guidelines/09-Hata-Yakalama-ve-Debug.md](Guidelines/09-Hata-Yakalama-ve-Debug.md) kullanın.

---

## Makro hiç çalışmıyor / “Sub or Function not defined”

| Olası neden | Ne yapmalı |
|-------------|------------|
| Makro güvenlik ayarı | 3DExperience içinde makro çalıştırmaya izin verin (Tools → Options → makro güvenliği). |
| Modül yanlış / kod yanlış yerde | Kodu doğru modüle yapıştırdığınızdan ve Sub adının tam yazıldığından emin olun. |
| Option Explicit + tanımsız değişken | Tüm değişkenleri `Dim` ile tanımlayın; yazım hatası (typo) varsa düzeltin. |

---

## “Object variable not set” / Nothing hatası

| Olası neden | Ne yapmalı |
|-------------|------------|
| 3DExperience kapalı | Önce 3DExperience’ı açın; sonra makroyu çalıştırın. |
| Açık belge yok | `ActiveDocument` Nothing döner. Kullanıcıya “Önce bir belge açın” mesajı verin; `If oDoc Is Nothing Then Exit Sub` ekleyin. |
| Aktif belge Part değil | Çizim veya montaj açıkken `GetItem("Part")` Nothing dönebilir. Part kontrolü yapın; örnek: [Examples/SadecePartKontrol.bas](Examples/SadecePartKontrol.bas). |
| Koleksiyon boş | `Parameters.Count = 0` veya `Shapes.Item(1)` boş koleksiyonda hata verir. `Count > 0` kontrolü ekleyin. |

---

## GetItem("Part") Part bulamıyor / hata veriyor

| Olası neden | Ne yapmalı |
|-------------|------------|
| Aktif belge çizim veya montaj | Sadece Part belgesinde Part nesnesi vardır. Belge türünü kontrol edin; Part değilse kullanıcıyı uyarın. |
| API adı sürüme göre farklı | Kendi sürümünüzde makro kaydı yapıp `GetItem` veya eşdeğer metod adını doğrulayın; Help ve [VBA_API_REFERENCE.md](VBA_API_REFERENCE.md) kullanın. |

---

## Update çağırdım ama ekranda değişiklik görünmüyor

| Olası neden | Ne yapmalı |
|-------------|------------|
| Update tek sefer, doğru yerde | Tüm parametre/şekil değişikliklerinden **sonra** bir kez `oPart.Update` çağırın; döngü içinde çağırmayın. |
| Belge read-only | Part read-only ise değişiklik uygulanmaz. Gerekirse `ReadOnly` kontrolü ve kullanıcı uyarısı ekleyin. |
| Görüntü yenilemesi | Nadiren ekran güncellenmesi gecikebilir; belgeyi kapatıp açmayı veya görünüm yenilemeyi deneyin. |

---

## Parametre bulunamadı / Parametre adı yanlış

| Olası neden | Ne yapmalı |
|-------------|------------|
| Parametre adı farklı (locale / sürüm) | Length.1, L.1 vb. sürüm veya dilde değişebilir. Makro kaydı ile gerçek parametre adını görün; `Parameters` döngüsüyle tüm adları listeleyebilirsiniz. |
| Parametre yok | Part’ta o isimde parametre tanımlı değildir. `Parameters.Item("isim")` sonrası `If oParam Is Nothing Then` kontrolü ekleyin. |

---

## Dosyaya yazarken hata (Permission denied / Path not found)

| Olası neden | Ne yapmalı |
|-------------|------------|
| Klasör yok | `C:\Temp` veya kullandığınız yol mevcut değilse oluşturun veya kullanıcıya seçtirin (FileDialog). |
| Yazma izni yok | Klasör veya sürücü salt okunur; farklı bir yol kullanın veya kullanıcıyı uyarın. |
| Dosya başka programda açık | Çıktı dosyası başka uygulamada açıksa kilitlenebilir; kapatıp tekrar deneyin. |

---

## Daha fazla yardım

- **Hata yakalama ve debug:** [Guidelines/09-Hata-Yakalama-ve-Debug.md](Guidelines/09-Hata-Yakalama-ve-Debug.md)  
- **Sık hatalar ve dikkat noktaları:** [Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md](Guidelines/18-Sik-Hatalar-ve-Dikkat-Edilecekler.md)  
- **Kısa soru–cevap:** [FAQ.md](FAQ.md)
