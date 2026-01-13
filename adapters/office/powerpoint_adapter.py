"""
PowerPoint to PDF converter adapter.
"""
import os
import win32com.client
from typing import List
from core.interfaces.converter import IConverter
from core.models.conversion_job import ConversionJob, ConversionResult
from utils.exceptions import ConversionError, OfficeApplicationError
from utils.logging import get_logger

logger = get_logger(__name__)


class PowerPointAdapter(IConverter):
    """
    Adapter for converting PowerPoint files to PDF using COM automation.
    """
    
    def supported_extensions(self) -> List[str]:
        """Returns supported PowerPoint extensions."""
        return ['.ppt', '.pptx']
    
    def convert(self, job: ConversionJob) -> ConversionResult:
        """
        Convert a PowerPoint file to PDF.
        
        Args:
            job: Conversion job with input/output paths
            
        Returns:
            ConversionResult indicating success or failure
        """
        powerpoint = None
        deck = None
        
        try:
            logger.info(f"Starting PowerPoint conversion: {job.input_path}")
            
            # Initialize PowerPoint application
            try:
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            except Exception as e:
                raise OfficeApplicationError(
                    f"Failed to initialize PowerPoint. Ensure it is installed. Error: {e}"
                )
            
            # Open presentation
            input_abs = os.path.abspath(job.input_path)
            output_abs = os.path.abspath(job.output_path)
            
            try:
                deck = powerpoint.Presentations.Open(input_abs, WithWindow=False)
            except Exception as e:
                raise ConversionError(f"Failed to open presentation: {e}")
            
            # Save as PDF (format code 32 = ppSaveAsPDF)
            try:
                deck.SaveAs(output_abs, 32)
                logger.info(f"PowerPoint conversion successful: {output_abs}")
                return ConversionResult.success_result(
                    output_path=output_abs,
                    message=f"Successfully converted {os.path.basename(job.input_path)}"
                )
            except Exception as e:
                raise ConversionError(f"Failed to save as PDF: {e}")
                
        except (ConversionError, OfficeApplicationError) as e:
            logger.error(f"PowerPoint conversion failed: {e}")
            return ConversionResult.failure_result(error=e, message=str(e))
            
        except Exception as e:
            logger.error(f"Unexpected error during PowerPoint conversion: {e}", exc_info=True)
            return ConversionResult.failure_result(
                error=e,
                message=f"Unexpected error: {e}"
            )
            
        finally:
            # Cleanup
            if deck:
                try:
                    deck.Close()
                except:
                    pass
            if powerpoint:
                try:
                    powerpoint.Quit()
                except:
                    pass
