"""
Custom exceptions for the PdfConverter application.
"""


class PdfConverterException(Exception):
    """Base exception for all PdfConverter errors."""
    pass


class ConversionError(PdfConverterException):
    """Raised when a conversion operation fails."""
    pass


class UnsupportedFileTypeError(PdfConverterException):
    """Raised when attempting to convert an unsupported file type."""
    pass


class OfficeApplicationError(PdfConverterException):
    """Raised when there's an issue with the Office application (not installed, crashed, etc.)."""
    pass


class ValidationError(PdfConverterException):
    """Raised when job validation fails."""
    pass
