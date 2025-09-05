import pandas as pd
import datetime
import os
from typing import Tuple, Optional
from csv_repair import CSVRepairer

class CSVParser:
    """Handles CSV file loading, validation, and preprocessing for ServiceNow ticket data."""
    
    REQUIRED_COLUMNS = ['Site', 'Number', 'Short description', 'Created']
    OPTIONAL_COLUMNS = ['Resolved']
    DATE_FORMAT = '%d-%b-%Y %H:%M:%S'
    
    def __init__(self, repair_callback=None):
        self.data = None
        self.original_data = None
        self.csv_repairer = CSVRepairer(progress_callback=repair_callback)
        self.was_repaired = False
        self.repaired_file_path = None
        
    def load_and_validate(self, file_path: str, auto_repair: bool = True) -> Tuple[bool, str]:
        """
        Load CSV file and validate required columns, with optional auto-repair.
        
        Args:
            file_path: Path to the CSV file
            auto_repair: Whether to attempt automatic repair if the file is corrupted
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        working_file_path = file_path
        
        try:
            # First attempt: try to load the file directly
            try:
                self.original_data = pd.read_csv(file_path, encoding='utf-8')
                success, validation_msg = self._validate_required_columns()
                if success:
                    # File loaded successfully without repair
                    self._finalize_loading()
                    return True, f"Successfully loaded {len(self.data)} tickets from {len(self.data['Site'].unique())} unique sites."
            except (UnicodeDecodeError, pd.errors.ParserError, KeyError) as e:
                if not auto_repair:
                    return False, f"File corrupted and auto-repair disabled: {str(e)}"
                # File needs repair, continue to repair section
                pass
            
            # Second attempt: try auto-repair if enabled
            if auto_repair:
                was_repaired, repair_message, repaired_path = self.csv_repairer.quick_repair_if_needed(file_path)
                
                if was_repaired and repaired_path:
                    self.was_repaired = True
                    self.repaired_file_path = repaired_path
                    working_file_path = repaired_path
                    
                    # Try to load the repaired file
                    try:
                        self.original_data = pd.read_csv(working_file_path, encoding='utf-8')
                        success, validation_msg = self._validate_required_columns()
                        
                        if success:
                            self._finalize_loading()
                            return True, f"Auto-repaired and loaded: {len(self.data)} tickets from {len(self.data['Site'].unique())} unique sites. ({repair_message})"
                        else:
                            return False, f"Repaired file still invalid: {validation_msg}"
                            
                    except Exception as e:
                        return False, f"Failed to load repaired file: {str(e)}"
                else:
                    return False, f"Auto-repair failed: {repair_message}"
            
            # If we get here, both direct loading and repair failed
            return False, "File could not be loaded or repaired"
            
        except FileNotFoundError:
            return False, f"File not found: {file_path}"
        except pd.errors.EmptyDataError:
            return False, "The CSV file is empty."
        except Exception as e:
            return False, f"Unexpected error loading CSV: {str(e)}"
    
    def _validate_required_columns(self) -> Tuple[bool, str]:
        """Validate that required columns are present."""
        if self.original_data is None:
            return False, "No data loaded"
            
        missing_columns = []
        for col in self.REQUIRED_COLUMNS:
            if col not in self.original_data.columns:
                missing_columns.append(col)
        
        if missing_columns:
            return False, f"Missing required columns: {', '.join(missing_columns)}"
        
        return True, "All required columns present"
    
    def _finalize_loading(self):
        """Finalize the loading process by creating working copy and parsing dates."""
        # Make a working copy
        self.data = self.original_data.copy()
        
        # Parse and validate the Created column
        success, message = self._parse_created_column()
        if not success:
            raise ValueError(message)
    
    def _parse_created_column(self) -> Tuple[bool, str]:
        """
        Parse the Created column into datetime objects.
        
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Convert Created column to datetime
            self.data['Created_dt'] = pd.to_datetime(self.data['Created'], format=self.DATE_FORMAT)
            
            # Check for any parsing failures (NaT values)
            failed_parsing = self.data['Created_dt'].isna().sum()
            if failed_parsing > 0:
                return False, f"Failed to parse {failed_parsing} date entries. Expected format: DD-Mon-YYYY HH:MM:SS"
            
            return True, "Date parsing successful."
            
        except Exception as e:
            return False, f"Error parsing Created column: {str(e)}"
    
    def get_filtered_data(self, exclude_resolved: bool = False) -> pd.DataFrame:
        """
        Get filtered data based on user preferences.
        
        Args:
            exclude_resolved: Whether to exclude tickets with resolved dates
            
        Returns:
            Filtered DataFrame
        """
        if self.data is None:
            return pd.DataFrame()
        
        filtered_data = self.data.copy()
        
        if exclude_resolved and 'Resolved' in filtered_data.columns:
            # Exclude tickets that have a non-empty Resolved field
            filtered_data = filtered_data[filtered_data['Resolved'].isna() | (filtered_data['Resolved'] == '')]
        
        return filtered_data
    
    def get_sites(self) -> list:
        """Get list of unique sites in the data."""
        if self.data is None:
            return []
        return sorted(self.data['Site'].unique().tolist())
    
    def get_total_tickets(self) -> int:
        """Get total number of tickets."""
        if self.data is None:
            return 0
        return len(self.data)
    
    def get_data_summary(self) -> dict:
        """Get summary statistics of the loaded data."""
        if self.data is None:
            return {}
        
        return {
            'total_tickets': len(self.data),
            'unique_sites': len(self.data['Site'].unique()),
            'date_range': {
                'earliest': self.data['Created_dt'].min(),
                'latest': self.data['Created_dt'].max()
            },
            'resolved_tickets': len(self.data[self.data['Resolved'].notna() & (self.data['Resolved'] != '')]) if 'Resolved' in self.data.columns else 0,
            'was_repaired': self.was_repaired
        }
    
    def cleanup_temp_files(self):
        """Clean up any temporary repaired files created during processing."""
        if self.repaired_file_path and os.path.exists(self.repaired_file_path):
            try:
                if self.repaired_file_path.endswith('_temp_repaired.csv'):
                    os.remove(self.repaired_file_path)
            except Exception:
                pass  # Ignore cleanup errors
    
    def manual_repair(self, file_path: str, create_backup: bool = True, 
                     target_encoding: str = 'utf-8', overwrite: bool = False) -> Tuple[bool, str, Optional[str]]:
        """
        Manually repair a CSV file with full control over options.
        
        Args:
            file_path: Path to the CSV file to repair
            create_backup: Whether to create a backup file
            target_encoding: Target encoding for output
            overwrite: Whether to overwrite the original file
            
        Returns:
            Tuple of (success: bool, message: str, output_path: Optional[str])
        """
        return self.csv_repairer.repair_csv(file_path, create_backup, target_encoding, overwrite)