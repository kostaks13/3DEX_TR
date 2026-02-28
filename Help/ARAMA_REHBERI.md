# Help ve API Listesinde Arama Rehberi

Bu doküman, **Help/text/** ve **VBA_CALL_LIST.txt** içinde arama yaparken kullanabileceğiniz **grep** (veya **ripgrep `rg`**) örneklerini listeler. Terminalde proje kökünden veya Help klasöründen çalıştırın.

---

## 1. Help/text/ içinde arama

Tüm Help metin dosyalarında bir kelime veya ifade aramak:

```bash
# "GetItem" geçen satırları göster (dosya adı + satır)
grep -n "GetItem" Help/text/*.txt

# Sadece hangi dosyalarda geçtiğini listele
grep -l "GetItem" Help/text/*.txt

# "Parameter" veya "Parameters" (büyük/küçük harf duyarsız)
grep -in "parameter" Help/text/*.txt
```

**ripgrep (rg)** kullanıyorsanız:

```bash
rg "GetItem" Help/text/
rg -i "ActiveDocument" Help/text/
rg "GetService" Help/text/ -l    # Sadece dosya adları
```

---

## 2. Belirli bir konuyu hangi dosyada bulacağınız

| Aradığınız | Önerilen dosya / arama |
|------------|-------------------------|
| Nesne modeli (Application, Part, Product) | `Help/text/Help-Native Apps Automation.txt`, `3DEXPERIENCE Otomasyon Hiyerarşi Ağacı.txt` |
| Servis listesi (GetService, GetSessionService) | `Help/text/Help-Common Services.txt`, `Help-Native Apps Automation.txt` |
| Kod kuralları (Option Explicit, başlık) | `Help/text/Help-Automation Development Guidelines.txt` |
| Makro hazırlık fazları (Design, Draft, Harden, Finalize) | `Help/text/Help-3DEXPERIENCE MACRO HAZIRLIK YÖNERGESİ.txt` |
| Parametre / Parameters | `rg -l "Parameters" Help/text/` ile dosya bulun; ilgili dosyada okuyun. |

---

## 3. VBA_CALL_LIST.txt içinde arama

VBA_CALL_LIST çok büyük olduğu için doğrudan açmak yerine önce arama yapmak işe yarar.

```bash
# "Part" sınıfı veya Part geçen bölümler (başlık satırları genelde ## ile)
grep -n "Part" Help/VBA_CALL_LIST.txt | head -80

# Belirli bir metod: "GetItem"
grep -n "GetItem" Help/VBA_CALL_LIST.txt

# "Parameters" ve "Item" birlikte (aynı satırda)
grep "Parameters\|Item" Help/VBA_CALL_LIST.txt

# Sadece VBA: satırlarını bul (imza satırı)
grep "VBA:" Help/VBA_CALL_LIST.txt | grep -i "update"
```

**rg** ile:

```bash
rg "\.Update" Help/VBA_CALL_LIST.txt
rg "Parameters" Help/VBA_CALL_LIST.txt -C 1   # 1 satır bağlam
rg "GetService" Help/VBA_CALL_LIST.txt
```

---

## 4. Pratik örnekler

- **“Parametre değeri nasıl yazarım?”**  
  → `rg -i "value\|parameter" Help/text/Help-Native*` veya SIK_KULLANILAN_API.txt içinde `Parameter.Value`.

- **“Shapes koleksiyonuna nasıl erişirim?”**  
  → `rg "Shapes" Help/VBA_CALL_LIST.txt` ve `rg "Shapes" Help/text/*.txt`.

- **“Hangi servisleri GetService ile alabilirim?”**  
  → `rg "GetService\|Service Identifier" Help/text/Help-Common*.txt` ve Help-Native Apps Automation.

- **“Drawing sayfaları ve Views”**  
  → `rg "Sheet\|View\|Drawing" Help/text/Help-Native*.txt`.

---

## 5. Editörde arama (Ctrl+F / Cmd+F)

VS Code, Cursor veya Notepad++ kullanıyorsanız:

1. **Help/text/** klasörünü açın; **Search in folder** ile "GetItem", "Parameters", "ActiveDocument" vb. arayın.
2. **VBA_CALL_LIST.txt** için dosyayı açıp Ctrl+F ile arama yapın; çok büyük olduğu için tercihen "Search in Files" ile Help/text/ içinde arama daha hızlı olabilir.

---

## İlgili dokümanlar

- **Help dosyalarını ne zaman kullanacağınız:** [Guidelines/17-Help-Dosyalarini-Kullanma.md](Guidelines/17-Help-Dosyalarini-Kullanma.md)
- **Sık kullanılan API özeti:** [SIK_KULLANILAN_API.txt](SIK_KULLANILAN_API.txt)
- **API referansı (proje kökü):** [VBA_API_REFERENCE.md](../VBA_API_REFERENCE.md)
