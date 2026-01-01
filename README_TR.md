# PowerPoint PDF Dönüştürücü

Windows için basit ve kullanıcı dostu masaüstü uygulaması. PowerPoint sunuları (PPT/PPTX) hızlı bir şekilde PDF formatına dönüştürür.

## Özellikler

- 🎨 **Kullanıcı Dostu Arayüz**: Tkinter ile yapılmış temiz ve sade arayüz
- 📁 **Toplu Dönüştürme**: Birden fazla PowerPoint dosyasını aynı anda dönüştürün
- ⚡ **Hızlı İşlem**: Windows COM API ile hızlı dönüştürme
- 🖥️ **Windows Entegrasyonu**: Microsoft Office ile doğrudan entegrasyon

## Gereksinimler

- **Windows** işletim sistemi
- **Python 3.7+**
- **Microsoft PowerPoint** bilgisayarınızda kurulu olmalıdır

## Kurulum

### 1. Projeyi İndirin veya Klonlayın

```bash
git clone <depo-url>
cd PdfConverter
```

### 2. Sanal Ortam Oluşturun (Önerilir)

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Bağımlılıkları Yükleyin

```bash
pip install -r requirements.txt
```

### 4. pywin32 Kurulumunu Tamamlayın (Zorunlu!)

`win32com.client` modülünün çalışması için bu adım gereklidir:

```bash
python -m Scripts.pywin32_postinstall -install
```

## Kullanım

Uygulamayı başlatın:

```bash
python converters/ppt_to_pdf.py
```

1. **"Klasör Seç"** butonuna tıklayarak PowerPoint dosyalarını içeren bir klasör seçin
2. Uygulama otomatik olarak `.ppt` ve `.pptx` dosyalarını bulacaktır
3. **"PDF'e Dönüştür"** butonuna tıklayarak dönüştürme işlemini başlatın
4. PDF dosyaları, orijinal dosyalarla aynı dizine kaydedilecektir

## Proje Yapısı

```
PdfConverter/
├── converters/
│   └── ppt_to_pdf.py          # Ana uygulama dosyası
├── requirements.txt            # Python paket gereksinimleri
├── .gitignore                 # Git gözardı kuralları
├── README.md                  # İngilizce dokumentasyon
└── README_TR.md               # Bu dosya (Türkçe)
```

## Gelecek Planları & Yeni Özellikler

Bu proje aktif olarak geliştirilmektedir. Ilerleyen zamanlarda aşağıdaki özellikler eklenecektir:

- ✅ PowerPoint PDF Dönüştürme (Mevcut)
- 📄 Word (DOCX/DOC) → PDF dönüştürme
- 📊 Excel (XLSX/XLS) → PDF dönüştürme
- 🎯 İlerleme göstergesi ile toplu işlem
- ⚙️ Dönüştürme ayarları ve konfigürasyon seçenekleri
- 📱 Komut satırı arayüzü (CLI)
- 🌐 Web tabanlı arayüz

Güncellemeyi bekleyiniz!

## Sorun Giderme

### "Module not found: win32com" Hatası

Kurulum adım 4'ü tamamladığınızdan emin olun.

### "You do not have the permissions to install COM objects" Uyarısı

Bu uyarı kritik değildir. pywin32 uzantıları başarıyla yüklendiyse güvenle göz ardı edebilirsiniz.

### Dönüştürme Sessizce Başarısız Oluyor

Microsoft PowerPoint'in kurulu olduğundan ve PowerPoint dosyasının bozuk veya şifre korumalı olmadığından emin olun.

## Lisans

Bu proje açık kaynaktır ve MIT Lisansı altında sunulur.

## Katkıda Bulunma

Katkılarınız hoş karşılanır! Sorunları bildirin veya pull request gönderin.

## Geliştirici

Windows sistemlerinde PowerPoint'i toplu olarak PDF'e dönüştürmek için bir utility olarak oluşturulmuştur.
