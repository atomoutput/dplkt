import pandas as pd
import os
from typing import Tuple

class ExportManager:
    """Handles exporting analysis results to various formats (CSV, Excel)."""
    
    def __init__(self):
        pass
    
    def export_data(self, data: pd.DataFrame, file_path: str) -> Tuple[bool, str]:
        """
        Export DataFrame to CSV or Excel format based on file extension.
        
        Args:
            data: DataFrame containing the results to export
            file_path: Target file path with extension
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            if data.empty:
                return False, "No data to export."
            
            file_extension = os.path.splitext(file_path)[1].lower()
            
            if file_extension == '.csv':
                return self._export_csv(data, file_path)
            elif file_extension in ['.xlsx', '.xls']:
                return self._export_excel(data, file_path)
            else:
                # Default to CSV if extension is unclear
                return self._export_csv(data, file_path + '.csv')
                
        except Exception as e:
            return False, f"Export failed: {str(e)}"
    
    def _export_csv(self, data: pd.DataFrame, file_path: str) -> Tuple[bool, str]:
        """
        Export DataFrame to CSV format.
        
        Args:
            data: DataFrame to export
            file_path: Target CSV file path
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            data.to_csv(file_path, index=False, encoding='utf-8')
            return True, f"Successfully exported {len(data)} records to CSV file."
            
        except PermissionError:
            return False, "Permission denied. Please ensure the file is not open in another application."
        except Exception as e:
            return False, f"CSV export failed: {str(e)}"
    
    def _export_excel(self, data: pd.DataFrame, file_path: str) -> Tuple[bool, str]:
        """
        Export DataFrame to Excel format with formatting.
        
        Args:
            data: DataFrame to export
            file_path: Target Excel file path
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Write main data
                data.to_excel(writer, sheet_name='Duplicate_Tickets', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Duplicate_Tickets']
                
                # Auto-adjust column widths
                self._adjust_excel_columns(worksheet)
                
                # Add summary sheet if there's data
                if not data.empty:
                    self._add_summary_sheet(writer, data)
            
            return True, f"Successfully exported {len(data)} records to Excel file."
            
        except ImportError:
            return False, "Excel export requires openpyxl. Please install it: pip install openpyxl"
        except PermissionError:
            return False, "Permission denied. Please ensure the file is not open in another application."
        except Exception as e:
            return False, f"Excel export failed: {str(e)}"
    
    def _adjust_excel_columns(self, worksheet):
        """Auto-adjust column widths in Excel worksheet."""
        try:
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        cell_length = len(str(cell.value)) if cell.value is not None else 0
                        max_length = max(max_length, cell_length)
                    except:
                        pass
                
                # Set width with some padding, but cap it at reasonable maximum
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
                
        except Exception:
            # If auto-adjustment fails, continue without it
            pass
    
    def _add_summary_sheet(self, writer, data: pd.DataFrame):
        """Add a summary sheet to the Excel export."""
        try:
            # Create summary statistics
            summary_data = []
            
            # Overall statistics
            total_pairs = len(data)
            unique_sites = data['Site'].nunique() if 'Site' in data.columns else 0
            
            # Time window breakdown
            if 'Time_Window_Hours' in data.columns:
                time_window_counts = data['Time_Window_Hours'].value_counts().sort_index()
                
                for window, count in time_window_counts.items():
                    summary_data.append({
                        'Metric': f'Duplicates within {window} hours',
                        'Value': count
                    })
            
            # Add overall metrics
            summary_data.extend([
                {'Metric': 'Total duplicate pairs', 'Value': total_pairs},
                {'Metric': 'Unique sites affected', 'Value': unique_sites}
            ])
            
            # Similarity statistics
            if 'Similarity_Score' in data.columns:
                avg_similarity = data['Similarity_Score'].mean()
                max_similarity = data['Similarity_Score'].max()
                min_similarity = data['Similarity_Score'].min()
                
                summary_data.extend([
                    {'Metric': 'Average similarity score', 'Value': f'{avg_similarity:.1f}%'},
                    {'Metric': 'Highest similarity score', 'Value': f'{max_similarity}%'},
                    {'Metric': 'Lowest similarity score', 'Value': f'{min_similarity}%'}
                ])
            
            # Create summary DataFrame and export
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Auto-adjust summary sheet columns
            summary_worksheet = writer.sheets['Summary']
            self._adjust_excel_columns(summary_worksheet)
            
        except Exception:
            # If summary creation fails, continue without it
            pass