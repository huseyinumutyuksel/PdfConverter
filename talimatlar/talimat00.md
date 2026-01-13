1. Word (DOC / DOCX) Desteği – Düşük Risk, Yüksek Getiri
1.1 Teknik Gerçeklik

Word → PDF dönüşümü, PowerPoint ile neredeyse birebir aynı COM API mantığına sahiptir.

Kullanılacak teknoloji

win32com.client.Dispatch("Word.Application")

Document.ExportAsFixedFormat

1.2 Word → PDF Önerilen Uygulama Stratejisi

Word’ü headless (Visible=False) çalıştır

Sayfa ayarlarını dokümana dokunmadan Word’ün kendi engine’ine bırak

DOC ve DOCX arasında fark gözetme (Word otomatik handle eder)

doc.ExportAsFixedFormat(
    OutputFileName=pdf_path,
    ExportFormat=17  # wdExportFormatPDF
)

1.3 Risk Analizi
Risk	Durum
Layout bozulması	❌ Çok düşük
Tablo kırılması	❌ Nadiren
Font kaybı	❌ Yok (native engine)

Sonuç:
Word desteği projeye minimum eforla eklenmeli. Teknik borç yaratmaz.

2. Excel (XLS / XLSX) – Asıl Kritik Nokta

Excel → PDF dönüşümü basit bir SaveAs problemi değildir, bir layout optimizasyon problemidir.

2.1 Temel Problemler (Doğru Tespit)

Tablolar tek sayfaya sığmıyor

Otomatik page break Excel tarafından kötü hesaplanıyor

Yatay / dikey yönlendirme yanlış

Geniş kolonlar PDF’de taşma yapıyor

Çok sayfalı sheet’lerde okunabilirlik düşüyor

Bu problemler ExportAsFixedFormat çağrısından önce çözülmelidir.

3. “Akıllı Excel → PDF” Yaklaşımı (Önerilen Gold Standard)
3.1 Ana Fikir

PDF’e geçmeden önce Excel Sheet’i programatik olarak normalize etmek.

PDF sadece çıktı formatı.
Asıl iş Excel sayfa ayarlarında yapılmalı.

4. Akıllı Excel Layout Pipeline (Önerilen Algoritma)
4.1 Sheet Bazlı İşleme (Zorunlu)

Her Worksheet için ayrı layout analizi yapılmalı.

Workbook
 ├─ Sheet1 → Layout Analyze → Normalize → Export
 ├─ Sheet2 → Layout Analyze → Normalize → Export

4.2 Layout Analizi (Ön Aşama)

Excel COM üzerinden ölçülebilen değerler:

Ölçüm	COM Property
Kullanılan alan	UsedRange
Kolon sayısı	UsedRange.Columns.Count
Satır sayısı	UsedRange.Rows.Count
Toplam genişlik	Columns(i).Width
Toplam yükseklik	Rows(i).Height
4.3 Akıllı Karar Mekanizması (Rule-Based)
1️⃣ Yönlendirme Kararı
Eğer toplam kolon genişliği > A4 genişliği → Landscape
Aksi halde → Portrait

sheet.PageSetup.Orientation = xlLandscape or xlPortrait

2️⃣ Ölçekleme (Scaling) Kararı

Excel’in tek sayfaya sığdırma özelliği burada kritik:

sheet.PageSetup.Zoom = False
sheet.PageSetup.FitToPagesWide = 1
sheet.PageSetup.FitToPagesTall = False  # Çok uzun tablolarda serbest bırak


Neden?

Yatay taşmayı kesin çözer

Dikeyde okunabilirliği korur

3️⃣ Page Break Yönetimi

Excel’in otomatik page break’leri genellikle kötüdür.

sheet.ResetAllPageBreaks()


Opsiyonel:

Her N satırda manuel page break

Ya da sadece Excel’e bırak

4️⃣ Yazdırma Alanı (Print Area)

Sadece kullanılan alanı PDF’e al:

sheet.PageSetup.PrintArea = sheet.UsedRange.Address

4.4 Export Aşaması
sheet.ExportAsFixedFormat(
    Type=0,  # PDF
    Filename=pdf_path,
    Quality=0,  # Standard
    IncludeDocProperties=True,
    IgnorePrintAreas=False
)

5. Mimari Öneri (Excel Eklenecekse ŞART)
5.1 Mutlaka Yapılması Gereken Refactor
converters/
 ├─ base_converter.py
 ├─ ppt_converter.py
 ├─ word_converter.py
 ├─ excel_converter.py
services/
 ├─ office_app_manager.py
ui/
 ├─ main_window.py

5.2 ExcelConverter Sorumluluğu

Sheet analiz

Layout kararları

Export

GUI tek satır Excel kodu görmemeli.

6. Threading + COM Uyarısı (ÖNEMLİ)

Excel ve Word COM thread-safe değildir.

Doğru Model:

1 worker thread

COM instance thread içinde yaratılmalı

GUI → Queue → Worker

Yanlış Model:
❌ Ana thread’de COM, worker’da SaveAs

7. Gelişmiş (Opsiyonel ama Güçlü) İyileştirme
Excel için “Akıllı Modlar”

Kullanıcıya seçenek sunabilirsiniz:

Mod	Açıklama
Auto (Önerilen)	Yukarıdaki rule-based sistem
Single Page	Her sheet tek sayfa
Readability	Daha az sıkıştırma
Raw	Excel default davranışı
8. Sonuç – Net Değerlendirme
Word Desteği

Kolay

Risksiz

Hemen eklenmeli

Excel Desteği

Basit değil

Layout-aware olmak zorunda

Yukarıdaki yaklaşım akademik + endüstriyel olarak doğrudur

Genel Yorum

Bu Excel stratejisi uygulanırsa:

“Basit dönüştürücü” değil

Akıllı Office PDF Engine seviyesine çıkalım.