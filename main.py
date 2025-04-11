import tkinter as tk
from tkinter import ttk, filedialog, Toplevel, messagebox
import os
import logging
from datetime import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import re
import shutil
import glob
from logger import logger
from xml_processor import XMLProcessor
from preview_window import PreviewWindow
import threading

class XMLProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Processor")
        self.root.geometry("500x350")
        
        self.processor = XMLProcessor()
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input paths section
        ttk.Label(main_frame, text="XML Files Directory", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=(10, 5))
        self.xml_path_var = tk.StringVar()
        self.xml_path_var.trace_add("write", self.on_path_change)
        xml_path_entry = ttk.Entry(main_frame, textvariable=self.xml_path_var, width=50)
        xml_path_entry.grid(row=1, column=0, sticky=tk.W+tk.E, padx=(0, 10))
        ttk.Button(main_frame, text="Browse", command=self.browse_xml_dir).grid(row=1, column=1, sticky=tk.W)
        
        # XML status label
        self.xml_status_var = tk.StringVar()
        self.xml_status_var.set("")
        self.xml_status_label = ttk.Label(main_frame, textvariable=self.xml_status_var, font=("Arial", 8), foreground="black")
        self.xml_status_label.grid(row=2, column=0, sticky=tk.W)
        
        # XLSX file input
        ttk.Label(main_frame, text="XLSX Reference File", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=(10, 5))
        self.xlsx_path_var = tk.StringVar()
        self.xlsx_path_var.trace_add("write", self.on_path_change)
        xlsx_path_entry = ttk.Entry(main_frame, textvariable=self.xlsx_path_var, width=50)
        xlsx_path_entry.grid(row=4, column=0, sticky=tk.W+tk.E, padx=(0, 10))
        ttk.Button(main_frame, text="Browse", command=self.browse_xlsx_file).grid(row=4, column=1, sticky=tk.W)
        
        # XLSX status label
        self.xlsx_status_var = tk.StringVar()
        self.xlsx_status_var.set("")
        self.xlsx_status_label = ttk.Label(main_frame, textvariable=self.xlsx_status_var, font=("Arial", 8), foreground="blue")
        self.xlsx_status_label.grid(row=5, column=0, sticky=tk.W)
        
        # Output path section
        ttk.Label(main_frame, text="Output Directory", font=("Arial", 10)).grid(row=6, column=0, sticky=tk.W, pady=(10, 5))
        self.output_path_var = tk.StringVar()
        self.output_path_var.trace_add("write", self.on_path_change)
        output_path_entry = ttk.Entry(main_frame, textvariable=self.output_path_var, width=50)
        output_path_entry.grid(row=7, column=0, sticky=tk.W+tk.E, padx=(0, 10))
        ttk.Button(main_frame, text="Browse", command=self.browse_output_dir).grid(row=7, column=1, sticky=tk.W)
        
        # Output status label
        self.output_status_var = tk.StringVar()
        self.output_status_var.set("")
        self.output_status_label = ttk.Label(main_frame, textvariable=self.output_status_var, font=("Arial", 8), foreground="black")
        self.output_status_label.grid(row=8, column=0, sticky=tk.W)
        
        # Process button
        self.process_button = ttk.Button(main_frame, text="Process", command=self.process_files)
        self.process_button.grid(row=9, column=0, columnspan=2, pady=(20, 10))
        
        # Loading state
        self.is_loading_xlsx = False
        
        # Status label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, font=("Arial", 9, "italic"))
        status_label.grid(row=10, column=0, columnspan=2, sticky=tk.W)
        
        # Configure grid
        main_frame.columnconfigure(0, weight=1)
        
        # Initialize variables for file processing
        self.xml_files = []
        self.processed_files = []
        self.changes_made = {}
        
        # Create default backup directory in the application root
        self.backup_dir = os.path.join(os.getcwd(), "backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        logger.info(f"Default backup directory created: {self.backup_dir}")
        
    def browse_xml_dir(self):
        directory = filedialog.askdirectory(title="Select XML Files Directory")
        if directory:
            self.xml_path_var.set(directory)
            logger.info(f"XML directory selected: {directory}")
    
    def browse_xlsx_file(self):
        """Browse for XLSX file"""
        xlsx_file = filedialog.askopenfilename(
            title="Select XLSX File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if xlsx_file:
            self.xlsx_path_var.set(xlsx_file)
            self.xlsx_status_var.set("Loading Excel data...")
            self.xlsx_status_label.configure(foreground="blue")
            
            # Disable process button during loading
            self.is_loading_xlsx = True
            self.process_button.configure(state="disabled")
            
            # Show loading indicator with specific message
            self.show_progress_indicator(True, "Loading and parsing Excel data...")
            
            # Start loading in a separate thread
            loading_thread = threading.Thread(target=self._load_xlsx_thread, args=(xlsx_file,))
            loading_thread.daemon = True
            loading_thread.start()
            
            logger.info(f"XLSX file selected: {xlsx_file}")

    def _load_xlsx_thread(self, xlsx_file):
        """Load XLSX file in a background thread"""
        try:
            # Load reference data from XLSX file
            success = self.processor.load_reference_data(xlsx_file)
            
            if success:
                self.root.after(0, lambda: self.xlsx_status_var.set("Excel data loaded successfully"))
                self.root.after(0, lambda: self.xlsx_status_label.configure(foreground="green"))
                logger.info(f"Successfully loaded XLSX file: {xlsx_file}")
            else:
                self.root.after(0, lambda: self.xlsx_status_var.set("No valid data found in Excel file"))
                self.root.after(0, lambda: self.xlsx_status_label.configure(foreground="red"))
                logger.warning(f"No valid data found in XLSX file: {xlsx_file}")
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self.xlsx_status_var.set(f"Error: {error_msg[:50]}..."))
            self.root.after(0, lambda: self.xlsx_status_label.configure(foreground="red"))
            logger.error(f"Error loading XLSX file {xlsx_file}: {error_msg}", exc_info=True)
        finally:
            # Re-enable process button and update loading state
            self.root.after(0, lambda: self.process_button.configure(state="normal"))
            self.root.after(0, lambda: setattr(self, 'is_loading_xlsx', False))
            self.root.after(0, lambda: self.show_progress_indicator(False))
    
    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_path_var.set(directory)
            logger.info(f"Output directory selected: {directory}")
    
    def on_path_change(self, *args):
        """Called when any path variable changes to trigger verification"""
        # Use after to avoid too many immediate calls
        self.root.after(500, self.verify_paths)

    def verify_paths(self):
        """Verify each path and update status labels"""
        # Don't verify paths if we're currently loading XLSX
        if self.is_loading_xlsx:
            return

        xml_valid = False
        output_valid = False
        xlsx_valid = False

        # Verify XML directory
        if self.xml_path_var.get():
            if os.path.isdir(self.xml_path_var.get()):
                xml_files = glob.glob(os.path.join(self.xml_path_var.get(), "*.xml"))
                if xml_files:
                    self.xml_status_var.set(f"✓ Found {len(xml_files)} XML files")
                    self.xml_status_label.configure(foreground="green")
                    xml_valid = True
                else:
                    self.xml_status_var.set("✗ No XML files found")
                    self.xml_status_label.configure(foreground="red")
            else:
                self.xml_status_var.set("✗ Invalid directory")
                self.xml_status_label.configure(foreground="red")
        else:
            self.xml_status_var.set("")
            self.xml_status_label.configure(foreground="black")
        
        # Verify XLSX file - only check if file exists and has correct extension
        if self.xlsx_path_var.get():
            if os.path.isfile(self.xlsx_path_var.get()) and self.xlsx_path_var.get().lower().endswith('.xlsx'):
                xlsx_valid = True
            else:
                self.xlsx_status_var.set("✗ Invalid XLSX file")
                self.xlsx_status_label.configure(foreground="red")
        else:
            self.xlsx_status_var.set("")
            self.xlsx_status_label.configure(foreground="black")
        
        # Verify output directory
        if self.output_path_var.get():
            if os.path.isdir(self.output_path_var.get()):
                self.output_status_var.set("✓ Valid output directory")
                self.output_status_label.configure(foreground="green")
                output_valid = True
            else:
                self.output_status_var.set("✗ Invalid directory")
                self.output_status_label.configure(foreground="red")
        else:
            self.output_status_var.set("")
            self.output_status_label.configure(foreground="black")

        # Enable/disable process button based on all validations
        if xml_valid and xlsx_valid and output_valid and not self.is_loading_xlsx:
            self.process_button.configure(state="normal")
        else:
            self.process_button.configure(state="disabled")
    
    def process_files(self):
        # Validate inputs
        if not self.xml_path_var.get() or not os.path.isdir(self.xml_path_var.get()):
            self.status_var.set("Error: Invalid XML directory")
            return
        
        if not self.xlsx_path_var.get() or not os.path.isfile(self.xlsx_path_var.get()):
            self.status_var.set("Error: Invalid XLSX file")
            return
        
        if not self.output_path_var.get() or not os.path.isdir(self.output_path_var.get()):
            self.status_var.set("Error: Invalid output directory")
            return
        
        # Show processing indicator
        self.status_var.set("Processing files...")
        self.show_progress_indicator(True, "Processing XML files. This may take a while...")
        logger.info("Process button clicked, starting file processing")
        
        # Start processing in a separate thread
        processing_thread = threading.Thread(target=self._process_files_thread)
        processing_thread.daemon = True
        processing_thread.start()

    def _process_files_thread(self):
        """Process files in a background thread"""
        try:
            # Find all XML files in the input directory
            xml_dir = self.xml_path_var.get()
            self.xml_files = glob.glob(os.path.join(xml_dir, "*.xml"))
            
            if not self.xml_files:
                self.root.after(0, lambda: self.status_var.set("No XML files found in the input directory"))
                self.root.after(0, lambda: self.show_progress_indicator(False))
                return
            
            # Check if reference data is already loaded
            if not self.processor.reference_data:
                self.root.after(0, lambda: self.show_progress_indicator(True, "Loading Excel data..."))
                try:
                    success = self.processor.load_reference_data(self.xlsx_path_var.get())
                    if not success:
                        self.root.after(0, lambda: self.status_var.set("No valid reference data found in XLSX file"))
                        self.root.after(0, lambda: self.show_progress_indicator(False))
                        return
                except Exception as e:
                    self.root.after(0, lambda: self.status_var.set(f"Error loading Excel data: {str(e)}"))
                    self.root.after(0, lambda: self.show_progress_indicator(False))
                    return
            
            # Process XML files
            self.processed_files = []
            self.changes_made = {}
            total_files = len(self.xml_files)
            
            for i, xml_file in enumerate(self.xml_files):
                # Update progress message
                file_name = os.path.basename(xml_file)
                progress_msg = f"Processing file {i+1} of {total_files}: {file_name}"
                self.root.after(0, lambda msg=progress_msg: self.show_progress_indicator(True, msg))
                
                processed_file, changes = self.processor.process_xml_file(xml_file, self.output_path_var.get())
                if processed_file:
                    self.processed_files.append(processed_file)
                    self.changes_made[xml_file] = changes
            
            if not self.processed_files:
                self.root.after(0, lambda: self.status_var.set("No files were processed"))
                self.root.after(0, lambda: self.show_progress_indicator(False))
                return
            
            # Open preview window in the main thread
            self.root.after(0, self.show_preview_window)
            
        except Exception as e:
            logger.error(f"Error processing files: {str(e)}", exc_info=True)
            self.root.after(0, lambda: self.status_var.set(f"Error: {str(e)}"))
            self.root.after(0, lambda: self.show_progress_indicator(False))

    def show_preview_window(self):
        """Show the preview window in the main thread"""
        self.show_progress_indicator(False)
        preview = PreviewWindow(self.root, self.xml_files, self.processed_files, 
                               self.changes_made, self.backup_dir, self.output_path_var.get())

    def show_progress_indicator(self, show=True, message=None):
        """Show or hide a progress indicator with optional message"""
        if not hasattr(self, 'progress_frame'):
            # Create progress frame if it doesn't exist
            self.progress_frame = ttk.Frame(self.root)
            self.progress_indicator = ttk.Progressbar(self.progress_frame, mode='indeterminate', length=300)
            self.progress_indicator.pack(pady=5)
            self.progress_message = ttk.Label(self.progress_frame, text="", wraplength=300)
            self.progress_message.pack(pady=5)
        
        if show:
            if message:
                self.progress_message.config(text=message)
            self.progress_frame.pack(pady=10)
            self.progress_indicator.start(10)
            self.root.update_idletasks() 
        else:
            self.progress_indicator.stop()
            self.progress_frame.pack_forget()
            self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = XMLProcessorApp(root)
    root.mainloop() 