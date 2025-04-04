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

class XMLProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Processor")
        self.root.geometry("600x400")
        
        # Create processor
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
        self.xml_status_label = ttk.Label(main_frame, textvariable=self.xml_status_var, font=("Arial", 8), foreground="blue")
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
        self.output_status_label = ttk.Label(main_frame, textvariable=self.output_status_var, font=("Arial", 8), foreground="blue")
        self.output_status_label.grid(row=8, column=0, sticky=tk.W)
        
        # Process button
        process_button = ttk.Button(main_frame, text="Process", command=self.process_files)
        process_button.grid(row=9, column=0, columnspan=2, pady=(20, 10))
        
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
        file_path = filedialog.askopenfilename(title="Select XLSX Reference File", 
                                              filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            self.xlsx_path_var.set(file_path)
            logger.info(f"XLSX file selected: {file_path}")
    
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
        # Verify XML directory
        if self.xml_path_var.get():
            if os.path.isdir(self.xml_path_var.get()):
                xml_files = glob.glob(os.path.join(self.xml_path_var.get(), "*.xml"))
                if xml_files:
                    self.xml_status_var.set(f"✓ Found {len(xml_files)} XML files")
                    self.xml_status_label.configure(foreground="green")
                else:
                    self.xml_status_var.set("✗ No XML files found")
                    self.xml_status_label.configure(foreground="red")
            else:
                self.xml_status_var.set("✗ Invalid directory")
                self.xml_status_label.configure(foreground="red")
        else:
            self.xml_status_var.set("")
        
        # Verify XLSX file
        if self.xlsx_path_var.get():
            if os.path.isfile(self.xlsx_path_var.get()) and self.xlsx_path_var.get().lower().endswith('.xlsx'):
                # Try to load reference data to verify
                try:
                    reference_data = self.processor.load_reference_data(self.xlsx_path_var.get())
                    if reference_data:
                        self.xlsx_status_var.set(f"✓ Found {len(reference_data)} reference entries")
                        self.xlsx_status_label.configure(foreground="green")
                    else:
                        self.xlsx_status_var.set("✗ No valid reference data found")
                        self.xlsx_status_label.configure(foreground="red")
                except Exception as e:
                    self.xlsx_status_var.set(f"✗ Error reading XLSX: {str(e)[:30]}...")
                    self.xlsx_status_label.configure(foreground="red")
            else:
                self.xlsx_status_var.set("✗ Invalid XLSX file")
                self.xlsx_status_label.configure(foreground="red")
        else:
            self.xlsx_status_var.set("")
        
        # Verify output directory
        if self.output_path_var.get():
            if os.path.isdir(self.output_path_var.get()):
                self.output_status_var.set("✓ Valid output directory")
                self.output_status_label.configure(foreground="green")
            else:
                self.output_status_var.set("✗ Invalid directory")
                self.output_status_label.configure(foreground="red")
        else:
            self.output_status_var.set("")
    
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
        
        self.status_var.set("Processing files...")
        logger.info("Process button clicked, starting file processing")
        
        try:
            # Find all XML files in the input directory
            xml_dir = self.xml_path_var.get()
            self.xml_files = glob.glob(os.path.join(xml_dir, "*.xml"))
            
            if not self.xml_files:
                self.status_var.set("No XML files found in the input directory")
                return
            
            # Load reference data from XLSX file
            self.processor.load_reference_data(self.xlsx_path_var.get())
            
            if not self.processor.reference_data:
                self.status_var.set("No valid reference data found in XLSX file")
                return
            
            # Process XML files
            self.processed_files = []
            self.changes_made = {}
            
            for xml_file in self.xml_files:
                processed_file, changes = self.processor.process_xml_file(xml_file, self.output_path_var.get())
                if processed_file:
                    self.processed_files.append(processed_file)
                    self.changes_made[xml_file] = changes
            
            if not self.processed_files:
                self.status_var.set("No files were processed")
                return
            
            # Open preview window
            preview = PreviewWindow(self.root, self.xml_files, self.processed_files, 
                                   self.changes_made, self.backup_dir, self.output_path_var.get())
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            logger.error(f"Error processing files: {str(e)}", exc_info=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = XMLProcessorApp(root)
    root.mainloop() 