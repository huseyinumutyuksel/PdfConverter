"""
Conversion service - orchestrates the conversion workflow.
"""
import os
from pathlib import Path
from typing import List, Dict, Optional
from core.interfaces.converter import IConverter
from core.models.conversion_job import ConversionJob, ConversionResult
from utils.exceptions import UnsupportedFileTypeError, ValidationError
from utils.logging import get_logger

logger = get_logger(__name__)


class ConversionService:
    """
    Service that orchestrates document conversion.
    
    This service follows the Dependency Inversion Principle:
    it depends on the IConverter abstraction, not concrete implementations.
    """
    
    def __init__(self):
        """Initialize the conversion service."""
        self._converters: Dict[str, IConverter] = {}
        
    def register_converter(self, converter: IConverter):
        """
        Register a converter for specific file types.
        
        Args:
            converter: Converter instance implementing IConverter
        """
        for ext in converter.supported_extensions():
            self._converters[ext.lower()] = converter
            logger.debug(f"Registered converter for {ext}: {converter.__class__.__name__}")
            
    def get_supported_extensions(self) -> List[str]:
        """
        Get all supported file extensions.
        
        Returns:
            List of supported extensions
        """
        return list(self._converters.keys())
    
    def get_available_converters(self) -> Dict[str, str]:
        """
        Get available converters grouped by type.
        
        Returns:
            Dict mapping file extensions to converter names
        """
        result = {}
        for ext, converter in self._converters.items():
            result[ext] = converter.__class__.__name__
        return result
        return list(self._converters.keys())
    
    def create_job(self, input_path: str, output_folder: str = None, custom_output_name: str = None) -> ConversionJob:
        """
        Create a conversion job.
        
        Args:
            input_path: Path to the input file
            output_folder: Optional output folder (defaults to same as input)
            custom_output_name: Optional custom output filename (without extension)
            
        Returns:
            ConversionJob instance
            
        Raises:
            ValidationError: If the file doesn't exist or is unsupported
        """
        if not os.path.isfile(input_path):
            raise ValidationError(f"File not found: {input_path}")
        
        # Determine output path
        input_file = Path(input_path)
        ext = input_file.suffix.lower()
        
        if ext not in self._converters:
            raise UnsupportedFileTypeError(
                f"Unsupported file type: {ext}. Supported: {', '.join(self.get_supported_extensions())}"
            )
        
        # Determine output filename
        if custom_output_name:
            pdf_name = custom_output_name + ".pdf"
        else:
            pdf_name = input_file.stem + ".pdf"
        
        # Determine output path
        if output_folder:
            output_path = os.path.join(output_folder, pdf_name)
        else:
            output_path = str(input_file.with_suffix(".pdf"))
            
        return ConversionJob(
            input_path=input_path,
            output_path=output_path,
            output_folder=output_folder
        )
    
    def convert(self, job: ConversionJob) -> ConversionResult:
        """
        Execute a conversion job.
        
        Args:
            job: The conversion job to execute
            
        Returns:
            ConversionResult
        """
        # Determine converter based on file extension
        ext = Path(job.input_path).suffix.lower()
        converter = self._converters.get(ext)
        
        if not converter:
            error = UnsupportedFileTypeError(f"No converter registered for {ext}")
            return ConversionResult.failure_result(error=error, message=str(error))
        
        logger.info(f"Converting {job.input_path} using {converter.__class__.__name__}")
        return converter.convert(job)
    
    def convert_batch(self, jobs: List[ConversionJob]) -> List[ConversionResult]:
        """
        Convert multiple files.
        
        Args:
            jobs: List of conversion jobs
            
        Returns:
            List of conversion results
        """
        results = []
        for job in jobs:
            result = self.convert(job)
            results.append(result)
        return results
