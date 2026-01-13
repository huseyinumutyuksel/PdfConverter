"""
Application configuration.
"""
import os
from pathlib import Path

# Application metadata
APP_NAME = "PdfConverter"
APP_VERSION = "2.0.0"
APP_DESCRIPTION = "Production-grade Office to PDF converter"

# Logging configuration
LOG_LEVEL = os.getenv("PDFCONVERTER_LOG_LEVEL", "INFO")
LOG_FILE = os.getenv("PDFCONVERTER_LOG_FILE", None)  # None = console only

# Paths
PROJECT_ROOT = Path(__file__).parent.parent
LOGS_DIR = PROJECT_ROOT / "logs"

# Ensure directories exist
LOGS_DIR.mkdir(exist_ok=True)
