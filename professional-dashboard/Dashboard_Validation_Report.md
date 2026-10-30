# Coffee Sales Dashboard — Doğrulama Raporu

**Tarih:** 30 Ekim 2026  
**Kaynak repository:** https://github.com/hakankoksalcan/coffee-sales-excel-analysis  
**İncelenen kaynak commit:** `87b4b0a`  
**Teslimat:** `Coffee_Sales_Dashboard_Professional.xlsx`

## Sonuç

669 gerçek işlem korunarak yeni bir Excel dashboard'u oluşturuldu. Üç KPI, dört yerel Excel grafiği ve gerçekten hesaplamaları değiştiren tarih/içecek girişleri eklendi. Kaynak çalışma kitabının beş özgün sayfası, beş PivotTable'ı, pivot önbelleği, üç özgün grafiği ve kalite kontrol kayıtları korundu.

Dosya LibreOffice ile açıldı, gerçek formülleri bağımsız olarak yeniden hesaplatıldı ve yedi senaryoda Python ile kaynak satırlarından hesaplanan sonuçlarla karşılaştırıldı. Tüm senaryolar geçti. Teslim edilen dosya LibreOffice ile yeniden kaydedilmedi; özgün PivotTable parçalarını koruyan paket teslim edildi.

**Microsoft Excel masaüstü uygulaması bu ortamda yoktur. Excel'deki açılış/onarım uyarısı, PivotTable Refresh All işlemi ve yeni tablo satırına otomatik formül doldurma davranışı doğrudan test edilmedi.** Bu işlevler tamamlandı diye işaretlenmemiştir. LibreOffice açılışı, OOXML yapısı, veri bütünlüğü, yerel grafikler ve formül hesaplamaları için aşağıda ayrı kanıtlar verilmiştir.

## 1. Bağımsız veri denetimi

`Ham_veri` sayfasındaki altı kaynak alanın tamamı okundu; gelir ondalık Decimal hesabıyla toplandı. Tam yinelenen kayıt kontrolünde altı alan birlikte, milisaniyeli zaman damgası dahil karşılaştırıldı. `Temiz_veri` sayfasının ilk altı alanı ham veriyle kayıt sırası dahil birebir karşılaştırıldı.

| Kontrol | Bağımsız sonuç | Durum |
| --- | ---: | --- |
| İşlem sayısı | 669 | Doğrulandı |
| Gelir | 22.834,18 currency units | Doğrulandı |
| Ortalama işlem tutarı | 34,1318086696562; gösterim 34,13 | Doğrulandı |
| İlk işlem tarihi | 1 Mart 2024 | Doğrulandı |
| Son işlem tarihi | 31 Mayıs 2024 | Doğrulandı |
| Farklı takvim ayı | 3 | Doğrulandı |
| Boş kart kimliği | 85 | Doğrulandı |
| Boş kart kimliği olan nakit işlem | 85 | Doğrulandı |
| Boş kart kimliği olan kartlı işlem | 0 | Doğrulandı |
| Tam yinelenen kayıt | 0 | Doğrulandı |
| Ham/temiz kaynak alanı farkı | 0 | Doğrulandı |
| Silinen veya eklenen teslimat kaydı | 0 | Doğrulandı |

85 boş kart kimliği doldurulmadı; nakit satışlar korundu. Test için değiştirilen/eklenen satırlar yalnızca geçici test kopyalarındadır. Teslimatta iki veri sayfasında da 669 kayıt vardır.

Para birimi bağımsız olarak doğrulanmadığından USD, EUR veya TL kullanılmadı. Gelir başlıkları ve birim açıklamaları **Revenue (currency units)** ifadesini kullanır. Galeride İngilizce sayı biçimi kullanılır: `22,834.18`.

Kaynak notlarında sözü geçen 3.636 satırlık CSV bu repository'de bulunmuyor. Onun tamamı bağımsız olarak denetlenmiş sayılmadı; denetim teslim edilen 669 satırlık Mart–Mayıs özetiyle sınırlıdır.

### Aylık sonuçlar

| Ay | Gelir (currency units) | İşlem |
| --- | ---: | ---: |
| Mart 2024 | 7.050,20 | 206 |
| Nisan 2024 | 6.720,56 | 196 |
| Mayıs 2024 | 9.063,42 | 267 |
| Toplam | 22.834,18 | 669 |

### İçecek sonuçları

| İçecek | Gelir (currency units) |
| --- | ---: |
| Latte | 5.265,62 |
| Cappuccino | 4.518,38 |
| Americano with Milk | 4.470,02 |
| Americano | 3.395,54 |
| Cortado | 1.892,32 |
| Hot Chocolate | 1.889,38 |
| Cocoa | 805,78 |
| Espresso | 597,14 |

Kartlı satışlar: **584 (%87,29)**; nakit satışlar: **85 (%12,71)**. Dashboard yüzdeleri bir ondalıkla %87,3 ve %12,7 gösterir. En yoğun kayıtlı saat **10:00–10:59**, **87 işlem**. Grafik 24 saatin tamamını içerir; sıfır işlem, işletmenin o saatte açık olduğunu kanıtlamaz.

## 2. Çalışma kitabı ve gerçek işlevler

| Sayfa | İşlev |
| --- | --- |
| `Dashboard` | Yeni yönetici görünümü; 3 KPI, 4 yerel Excel grafiği, tarih ve içecek girişleri |
| `Dashboard_Calc` | Yeni formüller, grafik kaynak aralıkları, kapsam mutabakatı |
| `Dashboard_Guide` | Kullanım, yenileme, kapsam ve teknik sınırlamalar |
| `Özet` | Özgün dashboard; hücreler/grafikler korundu, yalnızca sekme seçimi kapatıldı |
| `Analiz` | Özgün analiz, 4 PivotTable ve mevcut formüller; aynen korundu |
| `Ham_veri` | Özgün 669 satırlık veri anlık görüntüsü; aynen korundu |
| `Temiz_veri` | Düzenlenebilir mevcut `ham_veri__24` Excel tablosu ve 3 kalite/yardımcı formül sütunu; aynen korundu |
| `Kontrol_Notları` | Özgün kalite kontrol kayıtları ve 1 PivotTable; aynen korundu |

### Filtreler

- `Dashboard!D6`: başlangıç tarihi.
- `Dashboard!K6`: bitiş tarihi, dahil.
- `Dashboard!S6`: Excel veri doğrulama açılır listesi; All beverages veya mevcut sekiz içecekten biri.
- Tarih doğrulamaları Mart–Mayıs 2024 kapsamını ve başlangıç/bitiş sırasını kontrol eder. Normal Excel davranışı gereği yapıştırma işlemi doğrulamayı atlayabilir; onaylı değerler kullanılmalıdır.
- `SUMIFS`, `COUNTIFS` ve `IFERROR` formülleri üç KPI'yı ve dört grafik aralığını günceller. Bunlar sahte düğmeler veya statik KPI rakamları değildir.
- Boş seçim sonucunda KPI'lar 0, 0 ve 0,00 olur; sıfıra bölme hatası üretilmez. İşlem olmayan kapsamda ödeme dağılımı grafiğinin çizilecek dilimi yoktur.
- Bu girişler yeni dashboard'u filtreler. Özgün PivotTable'lar ve `Özet` sayfası tam dönem kapsamını korur.

### Grafik kaynakları

| Yerel grafik | Gerçek hücre kaynakları |
| --- | --- |
| Monthly Revenue Trend | `Dashboard_Calc!A8:A10`, `B8:B10` |
| Revenue by Beverage | `Dashboard_Calc!E8:E15`, `F8:F15` |
| Transactions by Hour | `Dashboard_Calc!A19:A42`, `B19:B42` |
| Payment Method Breakdown | `Dashboard_Calc!E19:E20`, `F19:F20` |

Grafikler gömülü resim değildir; düzenlenebilir Excel grafik nesneleridir. Gelir ve işlem eksenleri sıfırdan başlar. Gelir ölçeği binlik ayırıcılarla, aylık etiketler iki ondalıkla gösterilir. İçecek çubuk etiketleri okunabilirlik için tam birime yuvarlanır; hücrelerde iki ondalıklı tutar korunur. İçecek sırası tam dönem gelirine göre sabittir; filtre değişiminde sıralama yeniden yapılmaz.

Yeni 61 formül ve özgün 2.193 formül olmak üzere toplam 2.254 formül vardır. Yeni grafik XML önbelleklerindeki 37 değer, ilgili hesaplama hücreleriyle birebir karşılaştırıldı. Yeni formüller mevcut `ham_veri__24` tablo sütunlarına başvurur; harici sunucu, Power Query bağlantısı, makro veya ücretli eklenti gerektirmez.

## 3. PivotTable koruma yöntemi

openpyxl yalnızca okuma/denetim için kullanıldı. Özgün çalışma kitabı openpyxl ile yeniden kaydedilmedi. Yeni sayfalar XlsxWriter ile üretildi ve OOXML paketine eklendi; stil referansları, ilişki hedefleri, içerik türleri ve sayfa kimlikleri doğrulandı.

Özgün veri/analiz/kontrol sayfaları, tablo tanımları, pivot tanımları, pivot önbelleği, özgün grafikler ve çizim parçaları **bayt düzeyinde aynı** kaldı. `Özet` sayfasındaki yalnızca `tabSelected` görünüm bayrağı 1 yerine 0 yapıldı; yeni dashboard tek seçili sekme olarak açılır. Özgün hücre içeriği ve formüllere dokunulmadı. Toplam 34 özgün paket parçası aynen korundu. Bu tek görünüm değişikliği dışında beş paket metadata/stil parçası yeni içeriğe yer açmak için birleştirildi: `[Content_Types].xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`, `docProps/app.xml`.

Eski `calcChain.xml` kaldırıldı; yeni formüllerle birlikte bağımlılık zincirinin uygulama tarafından yeniden kurulması sağlandı. Bu işlem veri veya PivotTable silmez. Otomatik hesaplama, açılışta tam hesaplama ve tam hesaplamayı zorlama bayrakları ayarlandı.

Beş PivotTable, `CoffeePivotSource` adlı mevcut dinamik kaynağa bağlıdır. Tanım, düzenlenebilir tablonun satır sayısına göre `Temiz_veri` kaynağını genişletir. Beş PivotTable hem nihai dosyayı openpyxl ile okurken hem LibreOffice test kopyalarını yeniden açarken bulundu. **Bu, Microsoft Excel'de Refresh All işleminin çalıştırıldığı anlamına gelmez.** Pivot yenileme/doğrudan Excel etkileşimi bu ortamda doğrulanamadı.

Native slicer oluşturulmadı. Kullanılan dosya yazma araçlarıyla güvenilir slicer üretmek yerine standart Excel veri doğrulama listesi ve tarih girişleri uygulandı. Galeride slicer özelliği iddia edilmez.

## 4. Uygulama ve hesaplama testleri

Ortam: Python 3.12, openpyxl, XlsxWriter, lxml ve **LibreOfficeDev 26.8.0.0.alpha0**. Görsel kontrol, gerçek XLSX'in LibreOffice Calc PDF çıktısı üzerinden yapıldı; Microsoft Excel ekran görüntüsü olarak sunulmaz.

LibreOffice normal yüklemede Excel'in önbellekli formül değerlerini kullanabildiği için, bağımsız hesaplama testlerinde yalnızca geçici kopyaların formül sonuç önbellekleri temizlendi. Gerçek formüller uygulama tarafından yeniden hesaplandı. Teslim edilen dosyanın doğrulanmış sonuç önbellekleri korunmuştur.

| Test senaryosu | Beklenen gelir | Beklenen işlem | Beklenen ortalama | Sonuç |
| --- | ---: | ---: | ---: | --- |
| Tam dönem / tüm içecekler | 22,834.18 | 669 | 34.13180867 | Geçti |
| Yalnızca Nisan | 6,720.56 | 196 | 34.28857143 | Geçti |
| Yalnızca Latte | 5,265.62 | 137 | 38.43518248 | Geçti |
| Nisan + Latte | 1,193.12 | 31 | 38.48774194 | Geçti |
| 1 Mart + Espresso / satış yok | 0.00 | 0 | 0.00000000 | Geçti |
| Test kopyasında ilk tutar +1,25 | 22,835.43 | 669 | 34.13367713 | Geçti |
| Test kopyasında tabloya 1 satır ekleme | 22,872.88 | 670 | 34.13862687 | Geçti |

Her senaryoda yalnızca KPI'lar değil; üç aylık gelir, sekiz içecek geliri, 24 saatlik işlem ve iki ödeme yöntemi sayısı da kaynak satırlarından yeniden hesaplanan değerlerle karşılaştırıldı. Karşılaştırma toleransı 0,0000001 birimdi. Aylık/içecek gelir mutabakatı ve saat/ödeme işlem mutabakatı kontrollerinin tamamı sıfırdır (kayan nokta toleransı dahil). Yedi yeniden hesaplanan dosyada da hücre formül hatası sayısı **0**.

Yeni satır testi, geçici kopyada mevcut Excel tablosunun aralığı bir satır genişletilerek ve G:I formülleri yeni satıra taşınarak yapıldı. Böylece yapılandırılmış tablo referanslarının yeni satırı kapsadığı doğrulandı. **Excel arayüzünün bu genişletmeyi/formül doldurmayı otomatik yapması ayrıca test edilmedi.**

Nihai dosyanın ZIP bütünlüğü, tüm dahili ilişki hedefleri, toplam 7 yerel grafik nesnesi (4 yeni + 3 özgün), 5 PivotTable ve formül hata önbellekleri kontrol edildi. Eksik parça, bozuk dahili hedef veya `externalLinks` bulunmadı. Veri kaynağına yönelik harici bağlantı yoktur; özgün kaynak açıklamalarındaki internet adresleri veri bağlantısı değildir.

## 5. Fiverr galerisi ve görsel kalite

Görsellerin tamamı **nihai XLSX'in gerçek LibreOffice PDF render'ından** üretildi. PDF'in yalnızca dashboard sayfası ve tam sayfa içeriği kırpıldı. Kapak ve özellik görselinde aynı gerçek görüntü ölçeklendi; grafikler yeniden çizilmedi, rakamlar değiştirilmedi. Yapay zekâ görsel üretimi kullanılmadı. Pillow yalnızca kırpma, ölçekleme, metin ve yerleşim için kullanıldı.

| Dosya | Boyut | İçerik |
| --- | --- | --- |
| `Fiverr_Dashboard_Cover.png` | 1280 × 769 px | EXCEL DASHBOARDS; Charts • KPIs • PivotTables; gerçek dashboard |
| `Fiverr_Dashboard_Showcase.png` | 2843 × 1710 px | Tam dashboard'un yüksek çözünürlüklü gerçek render'ı |
| `Fiverr_Dashboard_Features.png` | 1280 × 769 px | 669 transactions; 3 months of sales data; Excel charts; KPI reporting; PivotTables |

Kapak ve özellik görseli görsel olarak açılıp incelendi. Başlıklar, KPI tutarları, sekiz içecek etiketi, aylık değerler ve ödeme yüzdeleri render üzerinde kontrol edildi. Grafik etiketlerinin yalnızca istenen değerleri göstermesi için OOXML veri etiketi bayrakları açıkça tanımlandı. Kesik başlık, çakışan etiket, fare imleci, sahte filtre düğmesi ve 3D grafik bulunmadı. Dashboard tek baskı sayfasına sığar; ekran görünümü %85 yakınlaştırmayla açılır. Excel şeridi/ekran çözünürlüğüne göre kullanıcı yakınlaştırmayı değiştirebilir.

Galeri `PivotTables` özelliğini, dosyada gerçekten bulunan ve korunan beş yerel PivotTable'a dayanarak gösterir. Slicer, otomatik kategori keşfi, kârlılık veya canlı sunucu bağlantısı iddia edilmez.

## 6. Yenileme ve bakım

1. `Ham_veri` özgün anlık görüntüsünü değiştirmeyin.
2. Onaylı düzeltmeleri `Temiz_veri` içindeki `ham_veri__24` tablosuna uygulayın. Yeni kayıtları tablonun son satırından Tab ile ilerleyerek/sonraki tablo satırına ekleyin. Tablo aralığının genişlediğini ve G:I yardımcı formüllerinin dolduğunu kontrol edin.
3. Microsoft Excel'de **Formulas > Calculation Options > Automatic** seçin ve **Ctrl+Alt+F9** ile tam hesaplayın.
4. Özgün PivotTable'lar için **Data > Refresh All / Ctrl+Alt+F5** uygulayın; tamamlanmasını bekleyin ve kaydedin. Bu adım Excel'de kullanıcı tarafından ayrıca doğrulanmalıdır.
5. `Dashboard!D6`, `K6`, `S6` seçimlerini yapın. Tam dönem için 01 Mar 2024, 31 May 2024, All beverages değerlerine dönün.
6. `Dashboard_Calc!J2:J5` değerlerinin sıfır olduğunu kontrol edin; gelirde 0,000001 altındaki kayan nokta farkını tolere edin. `Kontrol_Notları` ve `Temiz_veri` kalite sütunlarını inceleyin. Onaylı ekleme/düzeltmeden sonra ham veriye göre fark oluşması beklenir; farkı açıklayın, ham görüntüyü değiştirmeyin.
7. Yeni bir ay veya yeni içecek için model satırlarını, grafik aralıklarını, doğrulama listesini ve kapsam başlıklarını genişletip tekrar test edin. Bu teslimat Mart–Mayıs 2024 ve mevcut sekiz içecekle sınırlıdır.

## 7. Kontrol listesi

| İstenen kontrol | Sonuç |
| --- | --- |
| Orijinal veri satırı korunuyor | Geçti — ham ve temiz kaynakta 669 |
| Gelir ve ortalama doğru | Geçti — Decimal hesabı + bağımsız uygulama hesabı |
| Grafik rakamları kaynakla eşleşiyor | Geçti — kaynak aralıkları, 37 önbellek değeri ve senaryolar karşılaştırıldı |
| Excel dosyası açılıyor | LibreOffice ile açılış/dışa aktarma geçti; Microsoft Excel doğrudan test edilmedi |
| Okunabilir metin ve sayılar | Gerçek dosya render'ında görsel kontrol geçti |
| Tek ekranda yönetici özeti | Tek dashboard/tek baskı sayfası; farklı Excel ekranları ayrıca test edilmedi |
| Galeride kesik başlık/hizalama/imleç | Görsel kontrol geçti |
| Fiverr kapağı 1280 × 769 | Piksel ölçümü geçti |
| Çalışmayan/sahte filtre | Yok — gerçek hücre/formül filtreleri hesaplama testlerinden geçti; liste tıklaması Excel UI'da test edilmedi |
| Özgün PivotTable yapıları | Bayt düzeyinde korundu; Excel Refresh All doğrudan test edilmedi |
| Makro/harici ücretli eklenti zorunluluğu | Yok |

## 8. GitHub güvenliği ve teslimat konumu

Özgün `Coffee_Sales_Analysis.xlsx`, `dashboard.png`, `Executive_Summary.pdf` ve `README.md` içerikleri değiştirilmedi. Teslimatlar ayrı `redesign/professional-dashboard` dalında yeni `professional-dashboard/` klasöründe hazırlanır. Ana dala push/merge yapılmaz.

Oturumda belirtilen `/codex/browser/projectless/.../output` dizini çalıştırma ortamında yazılabilir değildi (`/codex` oluşturulurken Permission denied). Bu yüzden dosyalar önce `/workspace/work/output` altında üretildi ve repository'deki yeni teslimat klasörüne kopyalandı. Dosya içeriği bu konum farkından etkilenmez.

## 9. Dosya bütünlüğü

Özgün Excel SHA-256:

`173d9f50a3c37b946a3fdf520bef13954c1081a6de64b98a5bef44d79ffca649`

Teslimat dosyalarının SHA-256 değerleri (bu raporun kendi hash'i döngü oluşturmamak için listelenmedi):

| Dosya | SHA-256 |
| --- | --- |
| `Coffee_Sales_Dashboard_Professional.xlsx` | `001a987bda91e3979d9f32e96bfd3e07f440b001a6ed022c597b1d4706d6de97` |
| `Fiverr_Dashboard_Cover.png` | `92167d439d38697de1fff5e8a63c36160d5b150c359c74151cbfec3f489bdd5d` |
| `Fiverr_Dashboard_Features.png` | `1f6edcdcffbc2fa8963cca53b6a198d4e516db0d5d76b00dcd1ee5f23cdf1353` |
| `Fiverr_Dashboard_Showcase.png` | `48dc6a71d2496712b03dfa98cd171e645b536fdb5582c9e6a8e60880e9eadea5` |
