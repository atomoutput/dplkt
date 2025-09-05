import os
import shutil
import pandas as pd
from typing import Tuple, Optional
from datetime import datetime

try:
    import chardet
    HAS_CHARDET = True
except ImportError:
    HAS_CHARDET = False
    print("Warning: chardet not available. Using basic encoding detection.")

class CSVRepairer:
    """CSV repair functionality for fixing corrupted or malformed CSV files."""
    
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback
        
    def detect_encoding(self, filepath: str) -> str:
        """Detect the encoding of a CSV file."""
        if HAS_CHARDET:
            try:
                with open(filepath, 'rb') as f:
                    raw_data = f.read(10000)  # Read first 10KB for detection
                
                result = chardet.detect(raw_data)
                encoding = result['encoding']
                confidence = result['confidence']
                
                if self.progress_callback:
                    self.progress_callback(f"Detected encoding: {encoding} (confidence: {confidence:.2f})")
                
                return encoding if confidence > 0.7 else 'utf-8'
                
            except Exception:
                return 'utf-8'
        else:
            # Fallback: try common encodings
            encodings_to_try = ['utf-8', 'windows-1252', 'iso-8859-1']
            for encoding in encodings_to_try:
                try:
                    with open(filepath, 'r', encoding=encoding) as f:
                        f.read(1000)  # Try to read first 1KB
                    return encoding
                except UnicodeDecodeError:
                    continue
            return 'utf-8'
    
    def repair_csv(self, filepath: str, create_backup: bool = True, 
                   target_encoding: str = 'utf-8', overwrite: bool = False) -> Tuple[bool, str, Optional[str]]:
        """
        Repair a CSV file by fixing encoding issues, removing empty rows, and standardizing format.
        
        Args:
            filepath: Path to the CSV file to repair
            create_backup: Whether to create a backup of the original file
            target_encoding: Target encoding for the output file
            overwrite: Whether to overwrite the original file
            
        Returns:
            Tuple of (success: bool, message: str, output_path: Optional[str])
        """
        try:
            if self.progress_callback:
                self.progress_callback(f"Repairing: {os.path.basename(filepath)}")
            
            # Detect current encoding
            current_encoding = self.detect_encoding(filepath)
            
            # Create backup if requested
            if create_backup and not overwrite:
                backup_path = filepath + '.bak'
                shutil.copy2(filepath, backup_path)
                if self.progress_callback:
                    self.progress_callback(f"Backup created: {os.path.basename(backup_path)}")
            
            # Try to read the CSV with various encodings
            encodings_to_try = [current_encoding, 'utf-8', 'windows-1252', 'iso-8859-1', 'latin-1']
            df = None
            successful_encoding = None
            
            for encoding in encodings_to_try:
                try:
                    df = pd.read_csv(filepath, encoding=encoding, on_bad_lines='skip')
                    successful_encoding = encoding
                    if self.progress_callback:
                        self.progress_callback(f"Successfully read with {encoding} encoding")
                    break
                except (UnicodeDecodeError, pd.errors.ParserError) as e:
                    continue
                except Exception as e:
                    if self.progress_callback:
                        self.progress_callback(f"Failed to read with {encoding}: {str(e)[:100]}")
                    continue
            
            if df is None:
                return False, "Failed to read file with any encoding", None
            
            # Clean and validate data
            original_rows = len(df)
            original_columns = len(df.columns)
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            # Remove duplicate rows (but be conservative about this)
            df_before_dedup = len(df)
            df = df.drop_duplicates()
            
            cleaned_rows = len(df)
            duplicates_removed = df_before_dedup - cleaned_rows
            
            # Strip whitespace from string columns
            string_columns = df.select_dtypes(include=['object']).columns
            for col in string_columns:
                df[col] = df[col].astype(str).str.strip()
                # Replace 'nan' strings with actual NaN
                df[col] = df[col].replace('nan', pd.NA)
            
            # Determine output path
            if overwrite:
                output_path = filepath
            else:
                base, ext = os.path.splitext(filepath)
                output_path = f"{base}_repaired{ext}"
            
            # Write the repaired file
            df.to_csv(output_path, index=False, encoding=target_encoding)
            
            # Prepare summary message
            changes = []
            if original_rows != cleaned_rows:
                changes.append(f"Rows: {original_rows} → {cleaned_rows}")
            if duplicates_removed > 0:
                changes.append(f"Removed {duplicates_removed} duplicates")
            if successful_encoding != target_encoding:
                changes.append(f"Encoding: {successful_encoding} → {target_encoding}")
            
            change_summary = ", ".join(changes) if changes else "No changes needed"
            
            message = f"✓ Repaired successfully. {change_summary}. Columns: {original_columns}"
            
            if self.progress_callback:
                self.progress_callback(message)
            
            return True, message, output_path
            
        except Exception as e:
            error_msg = f"Error repairing CSV: {str(e)}"
            if self.progress_callback:
                self.progress_callback(f"✗ {error_msg}")
            return False, error_msg, None
    
    def quick_repair_if_needed(self, filepath: str, target_encoding: str = 'utf-8') -> Tuple[bool, str, Optional[str]]:
        """
        Quickly check if a CSV needs repair and fix it if necessary.
        This is a lightweight version for integration into the duplicate detection workflow.
        
        Returns:
            Tuple of (was_repaired: bool, message: str, repaired_file_path: Optional[str])
        """
        try:
            # First, try to read the file normally with full validation
            try:
                df = pd.read_csv(filepath, encoding='utf-8')
                # Check if it has the basic structure we need
                if len(df.columns) >= 4 and len(df) > 0:
                    # File appears to be readable and has data
                    return False, "File is already readable", filepath
            except (UnicodeDecodeError, pd.errors.ParserError):
                # File definitely needs repair due to encoding or parsing issues
                pass
            except Exception:
                # Other issues, try repair
                pass
            
            if self.progress_callback:
                self.progress_callback("File appears corrupted, attempting repair...")
            
            # Repair with temporary output
            base, ext = os.path.splitext(filepath)
            temp_output = f"{base}_temp_repaired{ext}"
            
            success, message, output_path = self.repair_csv(
                filepath, 
                create_backup=False,  # Don't create backup for temp repair
                target_encoding=target_encoding,
                overwrite=False
            )
            
            if success and output_path:
                # Rename temp file if needed
                if output_path != temp_output and os.path.exists(output_path):
                    shutil.move(output_path, temp_output)
                    return True, f"Auto-repaired: {message}", temp_output
                else:
                    return True, f"Auto-repaired: {message}", output_path
            else:
                return False, f"Auto-repair failed: {message}", None
                
        except Exception as e:
            return False, f"Auto-repair error: {str(e)}", None
    
    def validate_csv_structure(self, filepath: str) -> Tuple[bool, str]:
        """
        Validate that a CSV file has the expected structure for ServiceNow tickets.
        
        Returns:
            Tuple of (is_valid: bool, message: str)
        """
        try:
            # Try to read just the header
            df_header = pd.read_csv(filepath, nrows=0)
            columns = df_header.columns.tolist()
            
            # Check for required ServiceNow columns
            required_columns = ['Site', 'Number', 'Short description', 'Created']
            missing_columns = [col for col in required_columns if col not in columns]
            
            if missing_columns:
                return False, f"Missing required columns: {', '.join(missing_columns)}"
            
            # Try to read a few rows to check for basic structure issues
            df_sample = pd.read_csv(filepath, nrows=10)
            
            if len(df_sample) == 0:
                return False, "File appears to be empty"
            
            if len(df_sample.columns) < 4:
                return False, f"File has only {len(df_sample.columns)} columns, expected at least 4"
            
            return True, f"Valid CSV structure with {len(df_sample.columns)} columns"
            
        except Exception as e:
            return False, f"Structure validation failed: {str(e)}"