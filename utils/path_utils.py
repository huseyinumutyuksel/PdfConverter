"""
Path utilities for output folder management.
"""
import os
import subprocess
from datetime import datetime
from pathlib import Path
from utils.logging import get_logger

logger = get_logger(__name__)


def create_output_folder(base_path: str, folder_name: str = None, use_timestamp: bool = True) -> str:
    """
    Create an output folder for converted PDFs.
    
    Args:
        base_path: Base directory where output folder should be created
        folder_name: Custom folder name (optional)
        use_timestamp: If True, append timestamp to folder name
        
    Returns:
        Absolute path to the created output folder
        
    Raises:
        OSError: If folder creation fails
    """
    if not os.path.isdir(base_path):
        raise ValueError(f"Base path is not a valid directory: {base_path}")
    
    # Determine folder name
    if folder_name is None:
        folder_name = "PDF_Output"
    
    if use_timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"{folder_name}_{timestamp}"
    
    output_path = os.path.join(base_path, folder_name)
    
    # Create folder (handle existing)
    try:
        os.makedirs(output_path, exist_ok=True)
        logger.info(f"Created output folder: {output_path}")
        return os.path.abspath(output_path)
    except OSError as e:
        logger.error(f"Failed to create output folder: {e}")
        raise


def open_folder_in_explorer(folder_path: str) -> bool:
    """
    Open a folder in Windows Explorer.
    
    Args:
        folder_path: Absolute path to the folder
        
    Returns:
        True if successful, False otherwise
    """
    if not os.path.isdir(folder_path):
        logger.warning(f"Cannot open non-existent folder: {folder_path}")
        return False
    
    try:
        # Windows-specific: use explorer.exe
        subprocess.run(['explorer', os.path.abspath(folder_path)], check=True)
        logger.info(f"Opened folder in Explorer: {folder_path}")
        return True
    except Exception as e:
        logger.error(f"Failed to open folder in Explorer: {e}")
        return False


def validate_output_path(output_path: str) -> bool:
    """
    Validate that an output path is writable.
    
    Args:
        output_path: Path to validate
        
    Returns:
        True if path is writable, False otherwise
    """
    try:
        # Check if parent directory exists and is writable
        parent = Path(output_path).parent
        if not parent.exists():
            return False
        
        # Try to create a test file
        test_file = parent / ".write_test"
        test_file.touch()
        test_file.unlink()
        return True
    except Exception:
        return False
