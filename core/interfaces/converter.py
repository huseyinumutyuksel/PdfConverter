"""
Core interface definitions for converters.
"""
from abc import ABC, abstractmethod
from typing import List
from core.models.conversion_job import ConversionJob, ConversionResult


class IConverter(ABC):
    """
    Abstract base class for all document converters.
    
    This interface enforces the Liskov Substitution Principle:
    any concrete converter can be used interchangeably.
    """
    
    @abstractmethod
    def supported_extensions(self) -> List[str]:
        """
        Returns a list of file extensions this converter supports.
        
        Returns:
            List of extensions (e.g., ['.ppt', '.pptx'])
        """
        pass
    
    @abstractmethod
    def convert(self, job: ConversionJob) -> ConversionResult:
        """
        Converts a single file to PDF.
        
        Args:
            job: The conversion job containing input/output paths and options
            
        Returns:
            ConversionResult indicating success or failure
        """
        pass
