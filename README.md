# PdfConverter v2.1

A production-grade Windows desktop application that converts Microsoft Office documents to PDF format with intelligent file management and user control.

## Features

- 🎨 **Intuitive Interface**: Clean GUI with file type selection and preview
- 📁 **Smart Batch Conversion**: Convert multiple files with granular type selection
- 📊 **File Preview**: See exactly what will be converted before starting
- 📂 **Organized Output**: PDFs saved in timestamped folders for easy management
- ⚡ **Non-Blocking UI**: Background processing keeps the interface responsive
- 🏗️ **Clean Architecture**: SOLID principles, extensible design
- 🖥️ **Windows-Native**: Direct integration with Microsoft Office via COM
- 📊 **Excel Support**: NEW in v2.1 - Smart layout optimization for spreadsheets
- 🔧 **Production-Ready**: Structured logging, error handling, thread safety
- 📦 **Standalone Executable**: No Python installation required

## Supported Formats

- **PowerPoint**: `.ppt`, `.pptx`
- **Word**: `.doc`, `.docx`
- **Excel**: `.xls`, `.xlsx`, `.xlsm` ✨ NEW in v2.1

## Requirements

- **Windows** operating system
- **Python 3.7+**
- **Microsoft Office** (PowerPoint and/or Word) installed on your system

## Installation

### 1. Clone the Repository

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

## Usage

### Option 1: Standalone Executable (Recommended)

1. Download `PdfConverter.exe` from releases
2. Double-click to run (no installation needed)

### Option 2: Run from Source

Run the application:

```bash
python -m app.main
```

### Using the Application

1. **Select File Types**: Check the boxes for file types you want to convert (PowerPoint, Word, Excel)
2. **Select Folder**: Click "Select Folder" and choose a directory containing Office files
3. **Preview Files**: Review the list of files that will be converted
4. **Convert**: Click "Convert to PDF" to start
5. **Monitor Progress**: Watch the progress bar
6. **Access PDFs**: Click "Open Output Folder" to view converted files

**Output Location**: PDFs are saved in a timestamped folder (e.g., `PDF_Output_20260113_141500`) within the selected directory.

## Project Structure

```
PdfConverter/
├── app/
│   ├── main.py                # Application entry point
│   └── config.py              # Configuration
├── core/
│   ├── interfaces/
│   │   └── converter.py       # IConverter interface
│   ├── models/
│   │   └── conversion_job.py  # Domain models
│   └── services/
│       ├── conversion_service.py
│       └── file_scanner.py
├── adapters/
│   └── office/
│       ├── powerpoint_adapter.py
│       └── word_adapter.py
├── ui/
│   └── desktop/
│       └── main_window.py     # Tkinter GUI
├── utils/
│   ├── exceptions.py          # Custom exceptions
│   ├── logging.py             # Logging setup
│   └── threading.py           # COM-safe worker thread
└── requirements.txt
```

## Architecture

This project follows **Clean Architecture** principles:

- **Core Domain**: Business logic and interfaces (framework-agnostic)
- **Adapters**: Office COM integration (infrastructure)
- **UI**: Presentation layer (Tkinter)
- **Dependency Inversion**: UI and adapters depend on core abstractions

### Key Design Principles

- **SOLID**: Single Responsibility, Open/Closed, Liskov Substitution, Interface Segregation, Dependency Inversion
- **Thread Safety**: COM objects are created and used within worker threads
- **Extensibility**: New converters can be added without modifying existing code

## Roadmap

| Version | Features |
|---------|----------|
| v2.0 ✅ | Refactored architecture + Word support |
| v2.1 ✅ | Excel support + File type selection + Output folder management + Standalone .exe |
| v2.2 | Image file support (.jpg, .png) + Advanced Excel layout modes |
| v3.0 | Web API (FastAPI) + Multi-language support |
| v3.1 | Desktop installer (MSI) |
| v4.0 | Cross-platform support |

## Troubleshooting

### "Module not found: win32com"

Make sure you installed `pywin32`:

```bash
pip install pywin32
```

### "Failed to initialize Word/PowerPoint"

Ensure Microsoft Office is installed and properly licensed on your system.

### Application freezes during conversion

This should not happen in v2.0. If it does, please report it as a bug.

## Contributing

Contributions are welcome! Please ensure:

- Code follows SOLID principles
- New converters implement the `IConverter` interface
- Changes are documented

## License

This project is open-source and available under the MIT License.

## Author

Created as a production-grade utility for batch Office to PDF conversions on Windows systems.

