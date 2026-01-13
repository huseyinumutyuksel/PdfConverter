"""
PyInstaller build script for creating standalone executable.

Usage:
    python build_exe.py

Output:
    dist/PdfConverter.exe
"""
import PyInstaller.__main__
import os
import sys

# Ensure we're in the project root
project_root = os.path.dirname(os.path.abspath(__file__))
os.chdir(project_root)

print("Building PdfConverter standalone executable...")
print(f"Project root: {project_root}")

# PyInstaller arguments
args = [
    'app/main.py',                    # Entry point
    '--name=PdfConverter',            # Executable name
    '--onefile',                      # Single executable file
    '--windowed',                     # No console window (GUI app)
    '--clean',                        # Clean cache before building
    '--noconfirm',                    # Overwrite without asking
    
    # Hidden imports (COM libraries)
    '--hidden-import=win32com',
    '--hidden-import=win32com.client',
    '--hidden-import=pywintypes',
    
    # Add data files
    '--add-data=README.md;.',
    
    # Exclude unnecessary modules to reduce size
    '--exclude-module=matplotlib',
    '--exclude-module=numpy',
    '--exclude-module=pandas',
    '--exclude-module=scipy',
    
    # Icon (optional - create if available)
    # '--icon=assets/icon.ico',
]

print("\nPyInstaller arguments:")
for arg in args:
    print(f"  {arg}")

print("\nBuilding...")
PyInstaller.__main__.run(args)

print("\n" + "="*60)
print("Build complete!")
print(f"Executable location: {os.path.join(project_root, 'dist', 'PdfConverter.exe')}")
print("="*60)
