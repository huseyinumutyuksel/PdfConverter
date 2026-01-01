# Kurulum Adımları

## 1. Requirements Kurulumu

```bash
pip install -r requirements.txt
```

## 2. pywin32 Post-Install (Zorunlu!)

`win32com.client` modülünün çalışması için bu adım gereklidir:

```bash
python Scripts/pywin32_postinstall.py -install
```

Veya eğer yukarıdaki çalışmazsa:

```bash
python -m pip install --upgrade pywin32
python -m Scripts.pywin32_postinstall -install
```

## 3. Uygulamayı Çalıştırma

```bash
python converters/ppt_to_pdf.py
```

## Not

- Bu uygulama **sadece Windows** sisteminde çalışır (Microsoft Office COM API gerektirir)
- Microsoft Office (PowerPoint) bilgisayarınızda kurulu olmalıdır
