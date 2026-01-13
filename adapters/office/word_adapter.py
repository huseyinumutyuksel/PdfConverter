"""
Word to PDF converter adapter.
"""
import os
import win32com.client
from typing import List
from core.interfaces.converter import IConverter
from core.models.conversion_job import ConversionJob, ConversionResult
from utils.exceptions import ConversionError, OfficeApplicationError
from utils.logging import get_logger

logger = get_logger(__name__)


class WordAdapter(IConverter):
    """
    Adapter for converting Word documents to PDF using COM automation.
    """
    
    def supported_extensions(self) -> List[str]:
        """Returns supported Word extensions."""
        return ['.doc', '.docx']
    
    def convert(self, job: ConversionJob) -> ConversionResult:
        """
        Convert a Word document to PDF.
        
        Args:
            job: Conversion job with input/output paths
            
        Returns:
            ConversionResult indicating success or failure
        """
        word = None
        doc = None
        
        try:
            logger.info(f"Starting Word conversion: {job.input_path}")
            
            # Initialize Word application (headless)
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
            except Exception as e:
                raise OfficeApplicationError(
                    f"Failed to initialize Word. Ensure it is installed. Error: {e}"
                )
            
            # Open document
            input_abs = os.path.abspath(job.input_path)
            output_abs = os.path.abspath(job.output_path)
            
            try:
                doc = word.Documents.Open(input_abs)
            except Exception as e:
                raise ConversionError(f"Failed to open document: {e}")
            
            # Export as PDF (format code 17 = wdExportFormatPDF)
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=output_abs,
                    ExportFormat=17,  # wdExportFormatPDF
                    OpenAfterExport=False,
                    OptimizeFor=0,  # Standard quality
                    CreateBookmarks=1,  # Create bookmarks from headings
                    DocStructureTags=True
                )
                logger.info(f"Word conversion successful: {output_abs}")
                return ConversionResult.success_result(
                    output_path=output_abs,
                    message=f"Successfully converted {os.path.basename(job.input_path)}"
                )
            except Exception as e:
                raise ConversionError(f"Failed to export as PDF: {e}")
                
        except (ConversionError, OfficeApplicationError) as e:
            logger.error(f"Word conversion failed: {e}")
            return ConversionResult.failure_result(error=e, message=str(e))
            
        except Exception as e:
            logger.error(f"Unexpected error during Word conversion: {e}", exc_info=True)
            return ConversionResult.failure_result(
                error=e,
                message=f"Unexpected error: {e}"
            )
            
        finally:
            # Cleanup
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
