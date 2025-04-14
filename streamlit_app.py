import streamlit as st
import os
import logging
import glob
import pandas as pd
import threading
from datetime import datetime
import shutil
from xml_processor import XMLProcessor

# Configure logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class StreamlitXMLProcessor:
    def __init__(self):
        self.processor = XMLProcessor()
        
        # Set page configuration
        st.set_page_config(page_title="XML Processor", layout="wide")
        st.title("XML Processor")
        
        # Initialize session state if not exists
        if 'xml_files' not in st.session_state:
            st.session_state.xml_files = []
        if 'processed_files' not in st.session_state:
            st.session_state.processed_files = []
        if 'changes_made' not in st.session_state:
            st.session_state.changes_made = {}
        if 'reference_data_loaded' not in st.session_state:
            st.session_state.reference_data_loaded = False
            
        # Create default backup directory
        self.backup_dir = os.path.join(os.getcwd(), "backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        
        self.create_ui()
    
    def create_ui(self):
        # Create columns for the layout
        col1, col2 = st.columns(2)
        
        with col1:
            # XML Files Directory Input
            st.subheader("XML Files Directory")
            xml_dir = st.text_input("Enter path to XML files directory", key="xml_path")
            if xml_dir and os.path.isdir(xml_dir):
                xml_files = glob.glob(os.path.join(xml_dir, "*.xml"))
                if xml_files:
                    st.success(f"Found {len(xml_files)} XML files")
                    st.session_state.xml_files = xml_files
                else:
                    st.error("No XML files found in the directory")
            elif xml_dir:
                st.error("Invalid directory path")
                
            # XLSX Reference File Input
            st.subheader("XLSX Reference File")
            xlsx_file = st.text_input("Enter path to XLSX reference file", key="xlsx_path")
            if xlsx_file and os.path.isfile(xlsx_file) and xlsx_file.lower().endswith('.xlsx'):
                if st.button("Load XLSX Data"):
                    with st.spinner("Loading Excel data..."):
                        try:
                            success = self.processor.load_reference_data(xlsx_file)
                            if success:
                                st.success("Excel data loaded successfully")
                                st.session_state.reference_data_loaded = True
                            else:
                                st.error("No valid data found in Excel file")
                        except Exception as e:
                            st.error(f"Error loading XLSX file: {str(e)}")
            elif xlsx_file:
                st.error("Invalid XLSX file path")
        
        with col2:
            # Output Directory Input
            st.subheader("Output Directory")
            output_dir = st.text_input("Enter path to output directory", key="output_path")
            if output_dir and not os.path.isdir(output_dir):
                st.error("Invalid output directory path")
            
            # Process Button
            if (st.session_state.xml_files and 
                st.session_state.reference_data_loaded and 
                output_dir and os.path.isdir(output_dir)):
                
                if st.button("Process Files"):
                    with st.spinner("Processing XML files..."):
                        try:
                            self.process_files(output_dir)
                            if st.session_state.processed_files:
                                st.success(f"Processed {len(st.session_state.processed_files)} files")
                                self.display_preview()
                            else:
                                st.warning("No files were processed")
                        except Exception as e:
                            st.error(f"Error processing files: {str(e)}")
            else:
                st.info("Please provide all required inputs to process files")
    
    def process_files(self, output_dir):
        """Process XML files using the XMLProcessor"""
        st.session_state.processed_files = []
        st.session_state.changes_made = {}
        
        for i, xml_file in enumerate(st.session_state.xml_files):
            # Update progress
            file_name = os.path.basename(xml_file)
            st.text(f"Processing file {i+1} of {len(st.session_state.xml_files)}: {file_name}")
            
            processed_file, changes = self.processor.process_xml_file(xml_file, output_dir)
            if processed_file:
                st.session_state.processed_files.append(processed_file)
                st.session_state.changes_made[xml_file] = changes
    
    def display_preview(self):
        """Display preview of changes"""
        st.subheader("Preview of Changes")
        
        # Create tabs for each processed file
        if st.session_state.processed_files:
            tabs = st.tabs([os.path.basename(f) for f in st.session_state.processed_files])
            
            for i, (tab, processed_file) in enumerate(zip(tabs, st.session_state.processed_files)):
                with tab:
                    original_file = st.session_state.xml_files[i]
                    changes = st.session_state.changes_made.get(original_file, {})
                    
                    if changes:
                        st.write(f"Changes made to {os.path.basename(original_file)}:")
                        for key, value in changes.items():
                            st.write(f"• {key}: {value}")
                    else:
                        st.write("No changes were made to this file")
            
            # Add apply changes button
            if st.button("Apply All Changes"):
                with st.spinner("Applying changes..."):
                    # Create timestamped backup directory
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_subdir = os.path.join(self.backup_dir, f"backup_{timestamp}")
                    os.makedirs(backup_subdir, exist_ok=True)
                    
                    # Backup original files
                    for xml_file in st.session_state.xml_files:
                        if xml_file in st.session_state.changes_made:
                            shutil.copy2(xml_file, backup_subdir)
                    
                    # Copy processed files to original locations
                    for i, processed_file in enumerate(st.session_state.processed_files):
                        original_file = st.session_state.xml_files[i]
                        if os.path.exists(processed_file):
                            shutil.copy2(processed_file, original_file)
                    
                    st.success(f"Changes applied. Original files backed up to {backup_subdir}")
        else:
            st.info("No files were processed")

if __name__ == "__main__":
    app = StreamlitXMLProcessor() 