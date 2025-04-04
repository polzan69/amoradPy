import os
import xml.etree.ElementTree as ET
import pandas as pd
import shutil
import tempfile
from logger import logger

class XMLProcessor:
    def __init__(self):
        self.reference_data = {}
        
    def load_reference_data(self, xlsx_file):
        """Load reference data from XLSX file"""
        reference_data = {}
        
        try:
            # Read the Excel file
            logger.debug(f"Reading XLSX file: {xlsx_file}")
            df = pd.read_excel(xlsx_file)
            
            # Log the columns found in the file for debugging
            logger.debug(f"Columns found in {xlsx_file}: {list(df.columns)}")
            
            # Try different approaches to find the right columns
            expression_col = None
            date_col = None
            
            # First try exact column names
            expression_candidates = ['Expression', 'expression']
            date_candidates = ['SrcDate', 'srcdate', 'Date', 'date']
            
            # Check for exact matches first
            for col in expression_candidates:
                if col in df.columns:
                    expression_col = col
                    logger.debug(f"Found exact match for expression column: {col}")
                    break
                    
            for col in date_candidates:
                if col in df.columns:
                    date_col = col
                    logger.debug(f"Found exact match for date column: {col}")
                    break
            
            # If not found, try partial matching
            if not expression_col:
                for col in df.columns:
                    if 'express' in str(col).lower() or 'work' in str(col).lower():
                        expression_col = col
                        logger.debug(f"Found partial match for expression column: {col}")
                        break
            
            if not date_col:
                for col in df.columns:
                    if 'date' in str(col).lower() or 'src' in str(col).lower():
                        date_col = col
                        logger.debug(f"Found partial match for date column: {col}")
                        break
            
            # If still not found, try using column H (which might be SrcDate)
            if not date_col and len(df.columns) > 7:
                date_col = df.columns[7]  # Column H (0-indexed)
                logger.debug(f"Using column H as date column: {date_col}")
            
            # If still not found, try using column G (which might be Expression)
            if not expression_col and len(df.columns) > 6:
                expression_col = df.columns[6]  # Column G (0-indexed)
                logger.debug(f"Using column G as expression column: {expression_col}")
            
            if not expression_col or not date_col:
                logger.warning(f"XLSX file {xlsx_file} missing required columns. Found: {list(df.columns)}")
                return reference_data
            
            logger.info(f"Using columns: Expression={expression_col}, Date={date_col}")
            
            # Process each row
            for idx, row in df.iterrows():
                # Skip rows with missing data
                if pd.isna(row.get(expression_col)) or pd.isna(row.get(date_col)):
                    continue
                
                expression = str(row[expression_col]).strip()
                src_date = row[date_col]
                
                # Convert date to string format if it's a datetime
                if isinstance(src_date, pd.Timestamp):
                    src_date = src_date.strftime('%Y-%m-%d')
                elif isinstance(src_date, str):
                    # Try to parse and standardize date format if it's a string
                    try:
                        src_date = pd.to_datetime(src_date).strftime('%Y-%m-%d')
                    except:
                        # If parsing fails, use as is
                        pass
                elif isinstance(src_date, (int, float)):
                    # Handle numeric dates (Excel sometimes stores dates as numbers)
                    try:
                        # Convert Excel date number to datetime
                        src_date = pd.to_datetime('1899-12-30') + pd.Timedelta(days=int(src_date))
                        src_date = src_date.strftime('%Y-%m-%d')
                    except:
                        # If conversion fails, convert to string
                        src_date = str(src_date)
                
                # Add to reference data
                reference_data[expression] = src_date
                logger.debug(f"Added reference data: {expression} -> {src_date}")
            
        except Exception as e:
            logger.error(f"Error loading XLSX file {xlsx_file}: {str(e)}", exc_info=True)
        
        logger.info(f"Total reference data entries loaded: {len(reference_data)}")
        self.reference_data = reference_data
        return reference_data
    
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
        log_file = os.path.join(log_dir, f"changes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        
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