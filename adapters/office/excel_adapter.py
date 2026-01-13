"""
Excel to PDF converter adapter with smart layout optimization.
"""
import os
import win32com.client
from typing import List
from core.interfaces.converter import IConverter
from core.models.conversion_job import ConversionJob, ConversionResult
from utils.exceptions import ConversionError, OfficeApplicationError
from utils.logging import get_logger

logger = get_logger(__name__)

# Excel constants
xlLandscape = 2
xlPortrait = 1


class ExcelAdapter(IConverter):
    """
    Adapter for converting Excel files to PDF using COM automation.
    
    Implements smart layout optimization:
    - Automatic orientation detection (landscape for wide sheets)
    - Intelligent scaling (fit to page width)
    - Print area normalization
    """
    
    def supported_extensions(self) -> List[str]:
        """Returns supported Excel extensions."""
        return ['.xls', '.xlsx', '.xlsm']
    
    def convert(self, job: ConversionJob) -> ConversionResult:
        """
        Convert an Excel file to PDF with smart layout optimization.
        
        Args:
            job: Conversion job with input/output paths
            
        Returns:
            ConversionResult indicating success or failure
        """
        excel = None
        workbook = None
        
        try:
            logger.info(f"Starting Excel conversion: {job.input_path}")
            
            # Initialize Excel application (headless)
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
            except Exception as e:
                raise OfficeApplicationError(
                    f"Failed to initialize Excel. Ensure it is installed. Error: {e}"
                )
            
            # Open workbook
            input_abs = os.path.abspath(job.input_path)
            output_abs = os.path.abspath(job.output_path)
            
            try:
                workbook = excel.Workbooks.Open(input_abs, ReadOnly=True)
            except Exception as e:
                raise ConversionError(f"Failed to open workbook: {e}")
            
            # Process each visible worksheet
            try:
                for sheet in workbook.Worksheets:
                    if sheet.Visible:
                        self._optimize_sheet_layout(sheet)
            except Exception as e:
                logger.warning(f"Layout optimization failed, using default settings: {e}")
            
            # Export as PDF
            try:
                workbook.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=output_abs,
                    Quality=0,  # xlQualityStandard
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                logger.info(f"Excel conversion successful: {output_abs}")
                return ConversionResult.success_result(
                    output_path=output_abs,
                    message=f"Successfully converted {os.path.basename(job.input_path)}"
                )
            except Exception as e:
                raise ConversionError(f"Failed to export as PDF: {e}")
                
        except (ConversionError, OfficeApplicationError) as e:
            logger.error(f"Excel conversion failed: {e}")
            return ConversionResult.failure_result(error=e, message=str(e))
            
        except Exception as e:
            logger.error(f"Unexpected error during Excel conversion: {e}", exc_info=True)
            return ConversionResult.failure_result(
                error=e,
                message=f"Unexpected error: {e}"
            )
            
        finally:
            # Cleanup
            if workbook:
                try:
                    workbook.Close(SaveChanges=False)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
    
    def _optimize_sheet_layout(self, sheet):
        """
        Apply smart layout optimization to a worksheet.
        
        Args:
            sheet: Excel Worksheet COM object
        """
        try:
            # Get used range to analyze content
            used_range = sheet.UsedRange
            
            if not used_range:
                return
            
            # Analyze dimensions
            col_count = used_range.Columns.Count
            row_count = used_range.Rows.Count
            
            logger.debug(f"Sheet '{sheet.Name}': {row_count} rows x {col_count} columns")
            
            # Decision 1: Orientation
            # If more than 8 columns, use landscape
            if col_count > 8:
                sheet.PageSetup.Orientation = xlLandscape
                logger.debug(f"Sheet '{sheet.Name}': Set to Landscape (wide content)")
            else:
                sheet.PageSetup.Orientation = xlPortrait
                logger.debug(f"Sheet '{sheet.Name}': Set to Portrait")
            
            # Decision 2: Scaling
            # Fit to page width, allow vertical overflow
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False  # Allow multiple pages vertically
            
            # Decision 3: Print Area
            # Only print used range (exclude empty cells)
            sheet.PageSetup.PrintArea = used_range.Address
            
            # Decision 4: Page Breaks
            # Reset automatic page breaks
            try:
                sheet.ResetAllPageBreaks()
            except:
                pass  # Not critical if this fails
            
            # Additional optimizations
            sheet.PageSetup.CenterHorizontally = True
            sheet.PageSetup.CenterVertically = False
            
            logger.debug(f"Sheet '{sheet.Name}': Layout optimization complete")
            
        except Exception as e:
            logger.warning(f"Failed to optimize sheet '{sheet.Name}': {e}")
            # Don't raise - use default settings if optimization fails
