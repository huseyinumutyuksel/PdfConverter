"""
Core domain models for the PdfConverter application.
"""
from dataclasses import dataclass, field
from typing import Optional, Dict, Any


@dataclass
class ConversionJob:
    """
    Represents a single file conversion job.
    
    Attributes:
        input_path: Absolute path to the source file
        output_path: Absolute path where PDF should be saved
        output_folder: Optional output folder override (if None, uses output_path's directory)
        options: Additional conversion options (for future use, e.g., Excel sheet selection)
    """
    input_path: str
    output_path: str
    output_folder: Optional[str] = None
    options: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Validate job parameters."""
        if not self.input_path:
            raise ValueError("input_path cannot be empty")
        if not self.output_path:
            raise ValueError("output_path cannot be empty")


@dataclass
class ConversionResult:
    """
    Represents the result of a conversion operation.
    
    Attributes:
        success: Whether the conversion succeeded
        output_path: Path to the generated PDF (if successful)
        message: Human-readable status message
        error: Error details (if failed)
    """
    success: bool
    output_path: Optional[str] = None
    message: str = ""
    error: Optional[Exception] = None
    
    @classmethod
    def success_result(cls, output_path: str, message: str = "Conversion successful") -> "ConversionResult":
        """Factory method for successful conversion."""
        return cls(success=True, output_path=output_path, message=message)
    
    @classmethod
    def failure_result(cls, error: Exception, message: str = "Conversion failed") -> "ConversionResult":
        """Factory method for failed conversion."""
        return cls(success=False, error=error, message=message)
