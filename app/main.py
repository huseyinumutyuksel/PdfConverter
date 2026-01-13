"""
PdfConverter v2.0 - Main Entry Point

A production-grade Office to PDF converter following Clean Architecture principles.
"""
import sys
from app.config import APP_NAME, APP_VERSION, LOG_LEVEL, LOG_FILE
from utils.logging import setup_logging, get_logger
from core.services.conversion_service import ConversionService
from adapters.office.powerpoint_adapter import PowerPointAdapter
from adapters.office.word_adapter import WordAdapter
from adapters.office.excel_adapter import ExcelAdapter
from ui.desktop.main_window import MainWindow


def main():
    """Application entry point."""
    # Setup logging
    setup_logging(log_level=LOG_LEVEL, log_file=LOG_FILE)
    logger = get_logger(__name__)
    
    logger.info(f"Starting {APP_NAME} v{APP_VERSION}")
    
    try:
        # Initialize conversion service
        service = ConversionService()
        
        # Register converters (Dependency Injection)
        service.register_converter(PowerPointAdapter())
        service.register_converter(WordAdapter())
        service.register_converter(ExcelAdapter())  # NEW in v2.1
        
        logger.info(f"Registered converters for: {', '.join(service.get_supported_extensions())}")
        
        # Launch UI
        app = MainWindow(service)
        app.run()
        
    except Exception as e:
        logger.critical(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)
    
    logger.info("Application shutdown complete")


if __name__ == "__main__":
    main()
