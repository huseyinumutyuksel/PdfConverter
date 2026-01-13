"""
Desktop GUI for PdfConverter v2.1.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Set, Dict
import os
from pathlib import Path
from core.services.conversion_service import ConversionService
from core.services.file_scanner import FileScanner
from core.models.conversion_job import ConversionJob, ConversionResult
from utils.threading import ConversionWorker
from utils.path_utils import create_output_folder, open_folder_in_explorer
from utils.logging import get_logger

logger = get_logger(__name__)


class MainWindow:
    """
    Main application window for PdfConverter v2.1.
    
    Features:
    - File type selection via checkboxes
    - File list preview with Treeview
    - Output folder management
    - Progress tracking
    """
    
    def __init__(self, conversion_service: ConversionService):
        """
        Initialize the main window.
        
        Args:
            conversion_service: The conversion service instance
        """
        self.service = conversion_service
        self.worker = ConversionWorker()
        
        # State
        self.selected_folder = ""
        self.selected_types: Set[str] = set(self.service.get_supported_extensions())
        self.scanned_files: List[str] = []
        self.filtered_files: List[str] = []
        self.jobs: List[ConversionJob] = []
        self.output_folder_path = ""
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("PdfConverter v2.1")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Create UI elements
        self._create_widgets()
        
        # Start worker thread
        self.worker.start()
        
        # Cleanup on close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
    def _create_widgets(self):
        """Create all UI widgets."""
        # Title
        title_label = tk.Label(
            self.root,
            text="Office to PDF Converter v2.1",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=15)
        
        # File Type Selection Frame
        self._create_file_type_panel()
        
        # Folder Selection
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.path_label = tk.Label(
            folder_frame,
            text="No folder selected...",
            fg="gray",
            wraplength=600,
            anchor="w"
        )
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        select_btn = tk.Button(
            folder_frame,
            text="Select Folder",
            command=self._select_folder,
            width=15
        )
        select_btn.pack(side=tk.RIGHT, padx=5)
        
        # File List Frame
        self._create_file_list()
        
        # Output Folder Configuration
        self._create_output_folder_panel()
        
        # Progress bar
        self.progress = ttk.Progressbar(
            self.root,
            mode='determinate',
            length=600
        )
        self.progress.pack(pady=10, padx=20, fill=tk.X)
        
        # Convert button
        self.convert_btn = tk.Button(
            self.root,
            text="Convert to PDF",
            command=self._start_conversion,
            state=tk.DISABLED,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
            height=2
        )
        self.convert_btn.pack(pady=15, fill=tk.X, padx=50)
        
        # Status label
        self.status_label = tk.Label(
            self.root,
            text="",
            font=("Arial", 9),
            fg="green"
        )
        self.status_label.pack(pady=5)
        
        # Open folder button (initially hidden)
        self.open_folder_btn = tk.Button(
            self.root,
            text="Open Output Folder",
            command=self._open_output_folder,
            state=tk.DISABLED
        )
        self.open_folder_btn.pack(pady=5)
        
    def _create_file_type_panel(self):
        """Create file type selection checkboxes."""
        panel = tk.LabelFrame(self.root, text="Select File Types to Convert", font=("Arial", 10, "bold"))
        panel.pack(pady=10, padx=20, fill=tk.X)
        
        # Group extensions by converter type
        converters = self.service.get_available_converters()
        type_groups = {}
        
        for ext, converter_name in converters.items():
            # Simplify converter name (e.g., "PowerPointAdapter" -> "PowerPoint")
            type_name = converter_name.replace("Adapter", "")
            if type_name not in type_groups:
                type_groups[type_name] = []
            type_groups[type_name].append(ext)
        
        # Create checkboxes
        self.type_vars = {}
        checkbox_frame = tk.Frame(panel)
        checkbox_frame.pack(pady=10, padx=10)
        
        for i, (type_name, extensions) in enumerate(sorted(type_groups.items())):
            var = tk.BooleanVar(value=True)
            self.type_vars[type_name] = (var, extensions)
            
            ext_str = ", ".join(extensions)
            cb = tk.Checkbutton(
                checkbox_frame,
                text=f"{type_name} ({ext_str})",
                variable=var,
                command=self._on_type_selection_change,
                font=("Arial", 9)
            )
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=10, pady=3)
        
    def _create_file_list(self):
        """Create file list Treeview."""
        list_frame = tk.LabelFrame(self.root, text="Files to Convert", font=("Arial", 10, "bold"))
        list_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Treeview with scrollbar
        tree_scroll = ttk.Scrollbar(list_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_tree = ttk.Treeview(
            list_frame,
            columns=("Filename", "Type", "Size"),
            show="headings",
            yscrollcommand=tree_scroll.set,
            height=8
        )
        tree_scroll.config(command=self.file_tree.yview)
        
        # Column headers
        self.file_tree.heading("Filename", text="Filename")
        self.file_tree.heading("Type", text="Type")
        self.file_tree.heading("Size", text="Size")
        
        # Column widths
        self.file_tree.column("Filename", width=350)
        self.file_tree.column("Type", width=100)
        self.file_tree.column("Size", width=80)
        
        self.file_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def _create_output_folder_panel(self):
        """Create output folder configuration panel."""
        panel = tk.Frame(self.root)
        panel.pack(pady=5, padx=20, fill=tk.X)
        
        label = tk.Label(panel, text="Output Folder:", font=("Arial", 9))
        label.pack(side=tk.LEFT)
        
        self.output_folder_label = tk.Label(
            panel,
            text="(will be created automatically)",
            fg="gray",
            font=("Arial", 9, "italic")
        )
        self.output_folder_label.pack(side=tk.LEFT, padx=10)
        
    def _on_type_selection_change(self):
        """Handle file type checkbox changes."""
        # Update selected types
        self.selected_types.clear()
        for type_name, (var, extensions) in self.type_vars.items():
            if var.get():
                self.selected_types.update(extensions)
        
        # Refresh file list if folder is selected
        if self.selected_folder:
            self._filter_and_display_files()
        
    def _select_folder(self):
        """Handle folder selection."""
        folder = filedialog.askdirectory()
        if not folder:
            return
            
        self.selected_folder = folder
        self.path_label.config(text=f"Selected: {folder}", fg="black")
        
        # Scan for all files
        scanner = FileScanner(set(self.service.get_supported_extensions()))
        self.scanned_files = scanner.scan_folder(folder)
        
        # Filter and display
        self._filter_and_display_files()
        
    def _filter_and_display_files(self):
        """Filter scanned files based on selected types and update display."""
        # Filter files
        self.filtered_files = [
            f for f in self.scanned_files
            if Path(f).suffix.lower() in self.selected_types
        ]
        
        # Clear tree
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # Populate tree
        for file_path in self.filtered_files:
            path_obj = Path(file_path)
            filename = path_obj.name
            ext = path_obj.suffix.lower()
            
            # Determine type name
            type_name = ext.upper()
            for name, (var, extensions) in self.type_vars.items():
                if ext in extensions:
                    type_name = name
                    break
            
            # Get file size
            try:
                size_bytes = path_obj.stat().st_size
                if size_bytes < 1024:
                    size_str = f"{size_bytes} B"
                elif size_bytes < 1024 * 1024:
                    size_str = f"{size_bytes / 1024:.1f} KB"
                else:
                    size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
            except:
                size_str = "N/A"
            
            self.file_tree.insert("", "end", values=(filename, type_name, size_str))
        
        # Update UI state
        count = len(self.filtered_files)
        if count > 0:
            self.convert_btn.config(state=tk.NORMAL)
        else:
            self.convert_btn.config(state=tk.DISABLED)
        
        logger.info(f"Filtered {count} files from {len(self.scanned_files)} total")
        
    def _start_conversion(self):
        """Start the conversion process in background thread."""
        if not self.filtered_files:
            return
        
        # Create output folder
        try:
            self.output_folder_path = create_output_folder(
                self.selected_folder,
                folder_name="PDF_Output",
                use_timestamp=True
            )
            self.output_folder_label.config(
                text=os.path.basename(self.output_folder_path),
                fg="blue"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create output folder: {e}")
            return
        
        # Create jobs
        try:
            self.jobs = [
                self.service.create_job(f, output_folder=self.output_folder_path)
                for f in self.filtered_files
            ]
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create conversion jobs: {e}")
            return
        
        # Disable UI during conversion
        self.convert_btn.config(state=tk.DISABLED)
        self.progress['maximum'] = len(self.jobs)
        self.progress['value'] = 0
        self.status_label.config(text="Converting...", fg="orange")
        
        # Submit conversion task to worker
        def conversion_task():
            """Task that runs in worker thread."""
            results = []
            for i, job in enumerate(self.jobs):
                result = self.service.convert(job)
                results.append(result)
                
                # Update progress (thread-safe UI update)
                self.root.after(0, lambda v=i+1: self._update_progress(v))
                
            return results
            
        def on_complete(results):
            """Callback when conversion completes."""
            self.root.after(0, lambda: self._on_conversion_complete(results))
            
        self.worker.submit(conversion_task, on_complete)
        
    def _update_progress(self, value: int):
        """Update progress bar (must be called from main thread)."""
        self.progress['value'] = value
        self.status_label.config(text=f"Converting... ({value}/{len(self.jobs)})")
        
    def _on_conversion_complete(self, results: List[ConversionResult]):
        """Handle conversion completion."""
        if isinstance(results, Exception):
            messagebox.showerror("Error", f"Conversion failed: {results}")
            self.status_label.config(text="Conversion failed", fg="red")
        else:
            success_count = sum(1 for r in results if r.success)
            failed_count = len(results) - success_count
            
            message = f"Conversion complete!\n\nSuccessful: {success_count}\nFailed: {failed_count}"
            
            if failed_count > 0:
                # Show details of failures
                failures = [r.message for r in results if not r.success]
                message += "\n\nErrors:\n" + "\n".join(failures[:5])
                
            messagebox.showinfo("Complete", message)
            self.status_label.config(
                text=f"Completed: {success_count} successful, {failed_count} failed",
                fg="green" if failed_count == 0 else "orange"
            )
            
            # Enable "Open Folder" button
            self.open_folder_btn.config(state=tk.NORMAL)
            
        # Re-enable UI
        self.convert_btn.config(state=tk.NORMAL)
        self.progress['value'] = 0
        
    def _open_output_folder(self):
        """Open the output folder in Explorer."""
        if self.output_folder_path and os.path.isdir(self.output_folder_path):
            open_folder_in_explorer(self.output_folder_path)
        else:
            messagebox.showwarning("Warning", "Output folder not found")
        
    def _on_close(self):
        """Handle window close event."""
        logger.info("Shutting down application")
        self.worker.stop()
        self.root.destroy()
        
    def run(self):
        """Start the application main loop."""
        logger.info("Starting PdfConverter v2.1 UI")
        self.root.mainloop()
