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
            # Get list of all sheets in the Excel file
            xlsx = pd.ExcelFile(xlsx_file)
            sheet_names = xlsx.sheet_names
            logger.debug(f"Found sheets in Excel file: {sheet_names}")
            
            # Initialize reference data
            self.reference_data = {}
            
            # Required columns we're interested in
            required_columns = ['Expression', 'BaseFilename', 'SrcDate']
            
            # Process each sheet
            for sheet_name in sheet_names:
                logger.debug(f"Processing sheet: {sheet_name}")
                
                try:
                    # Read only the columns we need
                    df = pd.read_excel(xlsx_file, sheet_name=sheet_name, usecols=lambda x: x in required_columns)
                    
                    # Log the columns found in this sheet
                    found_columns = list(df.columns)
                    logger.debug(f"Columns found in sheet {sheet_name}: {found_columns}")
                    
                    # Check if all required columns are present
                    if not all(col in df.columns for col in required_columns):
                        missing_cols = [col for col in required_columns if col not in df.columns]
                        logger.warning(f"Sheet {sheet_name} missing required columns: {missing_cols}. Skipping sheet.")
                        continue
                    
                    # Process rows in this sheet
                    valid_rows = 0
                    for _, row in df.iterrows():
                        expression = str(row['Expression']).strip()
                        base_file = str(row['BaseFilename']).strip()
                        src_date = str(row['SrcDate']).strip()
                        
                        # Skip rows with empty values
                        if not expression or not base_file or pd.isna(src_date):
                            continue
                        
                        # Add to reference data
                        if expression not in self.reference_data:
                            self.reference_data[expression] = {}
                        
                        self.reference_data[expression]['file'] = base_file
                        self.reference_data[expression]['date'] = src_date
                        valid_rows += 1
                    
                    logger.info(f"Processed {valid_rows} valid rows from sheet: {sheet_name}")
                    
                except Exception as e:
                    logger.error(f"Error processing sheet {sheet_name}: {str(e)}")
                    continue
            
            total_entries = len(self.reference_data)
            logger.info(f"Total reference data entries loaded from all sheets: {total_entries}")
            
            if total_entries == 0:
                logger.warning("No valid data found in any sheet")
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading XLSX file {xlsx_file}: {str(e)}", exc_info=True)
            raise
    
    def _process_normalized_data(self, df):
        """Process the normalized dataframe and return success status"""
        # Process each row
        for _, row in df.iterrows():
            expression = str(row['expression']).strip()
            base_file = str(row['basefilename']).strip()
            src_date = str(row['srcdate']).strip()
            
            # Skip rows with empty values
            if not expression or not base_file or pd.isna(src_date):
                continue
                
            # Add to reference data
            if expression not in self.reference_data:
                self.reference_data[expression] = {}
            
            self.reference_data[expression]['file'] = base_file
            self.reference_data[expression]['date'] = src_date
        
        logger.info(f"Total reference data entries loaded: {len(self.reference_data)}")
        return len(self.reference_data) > 0
    
    def _normalize_columns(self, df):
        """Normalize column names to handle case sensitivity and variations"""
        # Required column names and their possible variations
        required_mappings = {
            'expression': ['Expression', 'expression', 'EXPRESSION', 'Name', 'name', 'Work', 'work'],
            'basefilename': ['BaseFilename', 'basefilename', 'BaseFileName', 'Filename', 'filename', 'File', 'file'],
            'srcdate': ['SrcDate', 'srcdate', 'SourceDate', 'sourcedate', 'Date', 'date', 'SRC', 'SRCDate', 'SrcDate', 'Notes SRC', 'SrcDate']
        }
        
        # Check if there's a column with "SRC" in it that might contain dates
        for col in df.columns:
            if 'SRC' in col.upper():
                # Check a few values to see if they look like dates
                sample_values = df[col].dropna().head(5).tolist()
                logger.debug(f"Checking if column '{col}' contains dates. Sample values: {sample_values}")
                date_like = False
                for val in sample_values:
                    val_str = str(val)
                    # Check if value matches common date patterns
                    if any(pattern in val_str for pattern in ['-', '/']):
                        date_like = True
                        logger.debug(f"Column '{col}' contains potential date values like: {val_str}")
                        required_mappings['srcdate'].append(col)
                        break
        
        # Create a mapping of actual column names to normalized names
        column_mapping = {}
        column_names = list(df.columns)
        
        # Log column names as they appear in the file
        logger.debug(f"Attempting to normalize columns: {column_names}")
        
        for norm_name, variations in required_mappings.items():
            found = False
            for var in variations:
                if var in column_names:
                    column_mapping[var] = norm_name
                    found = True
                    logger.debug(f"Mapped column '{var}' to '{norm_name}'")
                    break
            
            if not found:
                logger.error(f"Could not find any variation of required column '{norm_name}'")
                logger.error(f"Expected one of: {variations}")
                logger.error(f"Available columns: {column_names}")
                
                # Special handling for srcdate column - check if there's a column containing "SRC"
                if norm_name == 'srcdate':
                    for col in column_names:
                        if 'SRC' in col.upper() or 'DATE' in col.upper():
                            logger.debug(f"Found potential SRC/date column: {col}")
                            column_mapping[col] = norm_name
                            found = True
                            logger.debug(f"Mapped column '{col}' to '{norm_name}'")
                            break
                
                if not found:
                    # Last resort for srcdate - try to use a different date column if available
                    if norm_name == 'srcdate' and 'DateModified' in column_names:
                        logger.warning(f"Using 'DateModified' as fallback for 'srcdate'")
                        column_mapping['DateModified'] = norm_name
                        found = True
                    else:
                        return None
        
        # Rename the columns
        try:
            df_normalized = df.rename(columns=column_mapping)
            logger.debug(f"Columns successfully normalized to: {list(df_normalized.columns)}")
            return df_normalized
        except Exception as e:
            logger.error(f"Error normalizing columns: {str(e)}")
            return None
    
    def _log_sample_data(self, df):
        """Log a sample of the data for debugging purposes"""
        try:
            if len(df) == 0:
                logger.warning("Excel file contains no data rows")
                return
                
            # Get the first few rows (up to 5)
            sample_size = min(5, len(df))
            sample_df = df.head(sample_size)
            
            logger.debug("Sample data from Excel file:")
            for i, row in sample_df.iterrows():
                logger.debug(f"Row {i+1}:")
                for col in df.columns:
                    value = row[col]
                    logger.debug(f"  {col}: {value}")
            
            # Additional debug for specific columns like SRC
            logger.debug("Looking for specific columns or patterns:")
            for col in df.columns:
                if 'SRC' in col.upper() or 'DATE' in col.upper():
                    logger.debug(f"Found potential date/SRC column: {col}")
                    # Show first 3 values to help with debugging
                    values = df[col].head(3).tolist()
                    logger.debug(f"  Sample values for {col}: {values}")
                    
        except Exception as e:
            logger.error(f"Error logging sample data: {str(e)}")
    
    def process_xml_file(self, xml_file, output_dir):
        """Process a single XML file"""
        logger.debug(f"Processing XML file: {xml_file}")
        try:
            # Get the base filename without path
            xml_base_filename = os.path.basename(xml_file)
            
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
                    # Check if this expression should be applied to this file
                    ref_file = self.reference_data[expression]['file']
                    if ref_file != xml_base_filename:
                        continue
                        
                    # Get the new date from reference data and ensure it's just the date part
                    new_date = self.reference_data[expression]['date'].strip()
                    # Remove any time component if present (e.g., "2024-03-14 00:00:00" -> "2024-03-14")
                    if ' ' in new_date:
                        new_date = new_date.split()[0]
                    
                    # Get the current startEffectiveDate and ensure we compare just the date part
                    old_date = elem.get('startEffectiveDate', '').strip()
                    old_date_part = old_date.split()[0] if ' ' in old_date else old_date
                    
                    # Compare just the date parts and update if different
                    if old_date_part != new_date:
                        # Set the new date (without any time component)
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