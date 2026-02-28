# VBA API Referansı – 3DEX_TR

Bu sayfa **3DExperience VBA** makro geliştirirken kullanılan API’lere yönlendirme ve kısa açıklamalar içerir. Tam liste ve imzalar proje içindeki aşağıdaki kaynaklarda bulunur.

---

## Hızlı erişim

| Kaynak | Açıklama |
|--------|----------|
| [**Help/VBA_CALL_LIST.txt**](Help/VBA_CALL_LIST.txt) | Çağrılabilir API listesi (method imzaları, VBA karşılıkları) |
| [**Help/API_REPORT.csv**](Help/API_REPORT.csv) | API raporu (ek kaynak) |
| [**Help/text/**](Help/text/) | Resmi Help dokümanlarından üretilmiş metin dosyaları (Automation, Native Apps, Common Services vb.) |

---

## Sık kullanılan kalıplar (özet)

| İhtiyaç | Örnek (kavramsal) |
|--------|--------------------|
| Uygulama al | `Set oApp = GetObject(, "CATIA.Application")` |
| Aktif belge | `Set oDoc = oApp.ActiveDocument` |
| Part al | `Set oPart = oDoc.GetItem("Part")` veya `Set oPart = oDoc` |
| Parametre oku | `Set oParam = oPart.Parameters.Item("Length.1")` → `dVal = oParam.Value` |
| Parametre yaz | `oParam.Value = 50.5` → `oPart.Update` |
| Shapes döngüsü | `For i = 1 To oPart.Shapes.Count` … `Set oSh = oPart.Shapes.Item(i)` |
| Editor servis | `Set oEditor = oApp.ActiveEditor` → `Set oSvc = oEditor.GetService("InertiaService")` |

Tam imzalar ve sürüme özel API adları için **Help/VBA_CALL_LIST.txt** ve **Help/text/** içinde arama yapın.

---

## Rehberle ilişki

- **Nesne modeli ve erişim:** [Guidelines/06-3DExperience-Nesne-Modeli.md](Guidelines/06-3DExperience-Nesne-Modeli.md), [13-Erisim-ve-Kullanim-Rehberi.md](Guidelines/13-Erisim-ve-Kullanim-Rehberi.md)
- **Sık kullanılan API örnekleri:** [Guidelines/08-Sik-Kullanilan-APIler.md](Guidelines/08-Sik-Kullanilan-APIler.md)
- **Help dosyalarını ne zaman/nasıl kullanacağınız:** [Guidelines/17-Help-Dosyalarini-Kullanma.md](Guidelines/17-Help-Dosyalarini-Kullanma.md)

---

*API isimleri 3DExperience sürümüne göre değişir; güncel referans için Help klasöründeki dosyaları kullanın.*
