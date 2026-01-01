# PPT to PDF Converter

A simple and user-friendly Windows desktop application that converts PowerPoint presentations (PPT/PPTX) to PDF format with a graphical interface.

## Features

- 🎨 **User-Friendly Interface**: Clean and intuitive GUI built with Tkinter
- 📁 **Batch Conversion**: Convert multiple PowerPoint files at once
- ⚡ **Fast Processing**: Quick conversion using Windows COM API
- 🖥️ **Windows-Native**: Direct integration with Microsoft Office

## Requirements

- **Windows** operating system
- **Python 3.7+**
- **Microsoft PowerPoint** installed on your system

## Installation

### 1. Clone or Download the Project

```bash
git clone <repository-url>
cd PdfConverter
```

### 2. Create a Virtual Environment (Recommended)

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Complete pywin32 Setup (Important!)

This step is essential for `win32com.client` to work properly:

```bash
python -m Scripts.pywin32_postinstall -install
```

## Usage

Run the application:

```bash
python converters/ppt_to_pdf.py
```

1. Click **"Select Folder"** to choose a directory containing PowerPoint files
2. The app will automatically detect all `.ppt` and `.pptx` files
3. Click **"Convert to PDF"** to start the conversion
4. PDF files will be saved in the same directory as the original files

## Project Structure

```
PdfConverter/
├── converters/
│   └── ppt_to_pdf.py          # Main application file
├── requirements.txt            # Python dependencies
├── .gitignore                 # Git ignore rules
├── README.md                  # This file
└── README_TR.md               # Turkish documentation
```

## Roadmap & Future Features

This project is actively being developed. Future enhancements will include:

- ✅ PowerPoint to PDF (Current)
- 📄 Word (DOCX/DOC) to PDF conversion
- 📊 Excel (XLSX/XLS) to PDF conversion
- 🎯 Batch processing with progress tracking
- ⚙️ Configuration options for conversion settings
- 📱 Command-line interface (CLI)
- 🌐 Web-based interface

Stay tuned for updates!

## Troubleshooting

### "Module not found: win32com"

Make sure you completed step 4 of the installation process.

### "You do not have the permissions to install COM objects"

This is a non-critical warning and can be safely ignored if the pywin32 extensions were successfully installed.

### Conversion fails silently

Ensure Microsoft PowerPoint is installed and the PowerPoint file is not corrupted or password-protected.

## License

This project is open-source and available under the MIT License.

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

## Author

Created as a utility for batch PowerPoint to PDF conversions on Windows systems.
