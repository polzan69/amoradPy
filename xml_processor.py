import os
import xml.etree.ElementTree as ET
import pandas as pd
import shutil
import tempfile
from logger import logger
from datetime import datetime

class XMLProcessor:
    def __init__(self):
        self.reference_data = {}
        
    def load_reference_data(self, xlsx_file):
        """Load reference data from XLSX file"""
        logger.debug(f"Reading XLSX file: {xlsx_file}")
        
        try:
            # Use pandas to read the Excel file
            df = pd.read_excel(xlsx_file)
            
            # Process the data
            self.reference_data = {}
            
            # Validate required columns
            required_columns = ['expression', 'attribute', 'new_value']
            if not all(col in df.columns for col in required_columns):
                logger.error(f"XLSX file missing required columns. Required: {required_columns}")
                return False
            
            # Process each row
            for _, row in df.iterrows():
                expression = str(row['expression']).strip()
                attribute = str(row['attribute']).strip()
                new_value = str(row['new_value']).strip()
                
                # Skip rows with empty values
                if not expression or not attribute or pd.isna(new_value):
                    continue
                    
                # Add to reference data
                if expression not in self.reference_data:
                    self.reference_data[expression] = {}
                
                self.reference_data[expression][attribute] = new_value
            
            logger.info(f"Total reference data entries loaded: {len(self.reference_data)}")
            return len(self.reference_data) > 0
            
        except Exception as e:
            logger.error(f"Error loading XLSX file {xlsx_file}: {str(e)}", exc_info=True)
            raise
    
    def process_xml_file(self, xml_file, output_dir):
        """Process a single XML file"""
        logger.debug(f"Processing XML file: {xml_file}")
        try:
            # Parse the XML file
            tree = ET.parse(xml_file)
            root = tree.getroot()
            
            # Track changes
            changes = []
            
            # Find all lrvalueadd elements
            for elem in root.findall(".//lrvalueadd"):
                # Check if the expression attribute exists and matches any in our reference data
                expression = elem.get('expression')
                if expression and expression in self.reference_data:
                    # Get the new date from reference data
                    new_date = self.reference_data[expression]
                    
                    # Get the current startEffectiveDate
                    old_date = elem.get('startEffectiveDate')
                    
                    # Update the startEffectiveDate if different
                    if old_date and old_date != new_date:
                        elem.set('startEffectiveDate', new_date)
                        
                        # Record the change
                        changes.append({
                            'expression': expression,
                            'attribute': 'startEffectiveDate',
                            'old_value': old_date,
                            'new_value': new_date
                        })
                        logger.debug(f"Changed {expression} startEffectiveDate from {old_date} to {new_date}")
            
            # If no changes were made, return None
            if not changes:
                logger.debug(f"No changes made to {xml_file}")
                return None, []
            
            # Create a temporary file for the processed XML
            fd, temp_file = tempfile.mkstemp(suffix='.xml', dir=output_dir)
            os.close(fd)
            
            # Write the modified XML to the temporary file
            tree.write(temp_file, encoding='utf-8', xml_declaration=True)
            
            logger.info(f"Processed {xml_file} with {len(changes)} changes")
            return temp_file, changes
            
        except Exception as e:
            logger.error(f"Error processing XML file {xml_file}: {str(e)}", exc_info=True)
            return None, []
    
    def apply_changes(self, original_files, processed_files, changes_made, backup_dir, output_dir):
        """Apply changes to original files and create backups"""
        logger.info("Applying changes to original files")
        
        # Create log directory if it doesn't exist
        log_dir = os.path.join(output_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        
        # Create a log file for this processing session
        log_file = os.path.join(log_dir, f"processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        
        try:
            with open(log_file, 'w') as f:
                f.write(f"XML Processing Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                # Process each file
                for i, processed_file in enumerate(processed_files):
                    original_file = original_files[i]
                    file_name = os.path.basename(original_file)
                    
                    # Log the file being processed
                    f.write(f"File: {file_name}\n")
                    
                    # Log the changes
                    file_changes = changes_made.get(original_file, [])
                    for change in file_changes:
                        f.write(f"  Expression: {change['expression']}\n")
                        f.write(f"  Changed {change['attribute']} from '{change['old_value']}' to '{change['new_value']}'\n\n")
                    
                    # Create backup
                    backup_file = os.path.join(backup_dir, file_name)
                    shutil.copy2(original_file, backup_file)
                    logger.debug(f"Created backup: {backup_file}")
                    
                    # Replace original with processed
                    shutil.copy2(processed_file, original_file)
                    logger.debug(f"Replaced original file with processed version")
            
            logger.info(f"Processing complete. {len(processed_files)} files were updated.")
            return True, f"Processing complete. {len(processed_files)} files updated."
            
        except Exception as e:
            logger.error(f"Error during final processing: {str(e)}", exc_info=True)
            return False, f"An error occurred during processing: {str(e)}" 