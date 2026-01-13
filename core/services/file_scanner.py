"""
File scanning service for discovering convertible documents.
"""
import os
from pathlib import Path
from typing import List, Set
from utils.logging import get_logger

logger = get_logger(__name__)


class FileScanner:
    """
    Service for scanning directories and identifying convertible files.
    """
    
    def __init__(self, supported_extensions: Set[str]):
        """
        Initialize the file scanner.
        
        Args:
            supported_extensions: Set of file extensions to scan for (e.g., {'.ppt', '.docx'})
        """
        self.supported_extensions = {ext.lower() for ext in supported_extensions}
        
    def scan_folder(self, folder_path: str) -> List[str]:
        """
        Scan a folder for supported files.
        
        Args:
            folder_path: Path to the folder to scan
            
        Returns:
            List of absolute paths to supported files
        """
        if not os.path.isdir(folder_path):
            logger.warning(f"Invalid folder path: {folder_path}")
            return []
        
        found_files = []
        
        try:
            for entry in os.listdir(folder_path):
                file_path = os.path.join(folder_path, entry)
                
                if os.path.isfile(file_path):
                    ext = Path(file_path).suffix.lower()
                    if ext in self.supported_extensions:
                        found_files.append(os.path.abspath(file_path))
                        
            logger.info(f"Found {len(found_files)} convertible file(s) in {folder_path}")
            return found_files
            
        except Exception as e:
            logger.error(f"Error scanning folder {folder_path}: {e}")
            return []
