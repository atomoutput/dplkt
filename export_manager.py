import pandas as pd
import os
from typing import Tuple

class ExportManager:
    """Handles exporting analysis results to various formats (CSV, Excel) with enhanced multi-sheet support."""
    
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
    
    def export_enhanced_data(self, enhanced_results: dict, file_path: str, 
                           enable_same_day: bool = False, enable_rapid_fire: bool = False, 
                           enable_exact_match: bool = False, enable_category_patterns: bool = False) -> Tuple[bool, str]:
        """
        Export enhanced analysis results with multiple sheets.
        
        Args:
            enhanced_results: Dictionary containing different analysis results
            file_path: Target file path with extension
            enable_same_day: Include same-day analysis sheet
            enable_rapid_fire: Include rapid-fire analysis sheet
            enable_exact_match: Include exact match analysis sheet
            enable_category_patterns: Include category patterns sheet
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            file_extension = os.path.splitext(file_path)[1].lower()
            
            if file_extension == '.csv':
                # For CSV, export the primary fuzzy matching results only
                primary_data = self._convert_fuzzy_results_to_dataframe(enhanced_results.get('fuzzy_matching', {}))
                return self._export_csv(primary_data, file_path)
            elif file_extension in ['.xlsx', '.xls']:
                return self._export_enhanced_excel(enhanced_results, file_path, 
                                                 enable_same_day, enable_rapid_fire, 
                                                 enable_exact_match, enable_category_patterns)
            else:
                # Default to CSV if extension is unclear
                primary_data = self._convert_fuzzy_results_to_dataframe(enhanced_results.get('fuzzy_matching', {}))
                return self._export_csv(primary_data, file_path + '.csv')
                
        except Exception as e:
            return False, f"Enhanced export failed: {str(e)}"
    
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
    
    def _export_enhanced_excel(self, enhanced_results: dict, file_path: str,
                              enable_same_day: bool, enable_rapid_fire: bool, 
                              enable_exact_match: bool, enable_category_patterns: bool) -> Tuple[bool, str]:
        """
        Export enhanced analysis results to Excel with multiple sheets.
        
        Args:
            enhanced_results: Dictionary containing different analysis results
            file_path: Target Excel file path
            enable_same_day: Include same-day analysis sheet
            enable_rapid_fire: Include rapid-fire analysis sheet
            enable_exact_match: Include exact match analysis sheet
            enable_category_patterns: Include category patterns sheet
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            sheets_created = 0
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                
                # Primary sheet - Fuzzy matching duplicates (always included)
                primary_data = self._convert_fuzzy_results_to_dataframe(enhanced_results.get('fuzzy_matching', {}))
                if not primary_data.empty:
                    primary_data.to_excel(writer, sheet_name='Duplicate_Tickets', index=False)
                    self._adjust_excel_columns(writer.sheets['Duplicate_Tickets'])
                    sheets_created += 1
                
                # Same-day duplicates sheet
                if enable_same_day and 'same_day' in enhanced_results:
                    same_day_data = self._convert_same_day_to_dataframe(enhanced_results['same_day'])
                    if not same_day_data.empty:
                        same_day_data.to_excel(writer, sheet_name='Same_Day_Duplicates', index=False)
                        self._adjust_excel_columns(writer.sheets['Same_Day_Duplicates'])
                        sheets_created += 1
                
                # Rapid-fire duplicates sheet
                if enable_rapid_fire and 'rapid_fire' in enhanced_results:
                    rapid_fire_data = self._convert_rapid_fire_to_dataframe(enhanced_results['rapid_fire'])
                    if not rapid_fire_data.empty:
                        rapid_fire_data.to_excel(writer, sheet_name='Rapid_Fire_Duplicates', index=False)
                        self._adjust_excel_columns(writer.sheets['Rapid_Fire_Duplicates'])
                        sheets_created += 1
                
                # Exact matches sheet
                if enable_exact_match and 'exact_match' in enhanced_results:
                    exact_match_data = self._convert_exact_match_to_dataframe(enhanced_results['exact_match'])
                    if not exact_match_data.empty:
                        exact_match_data.to_excel(writer, sheet_name='Exact_Matches', index=False)
                        self._adjust_excel_columns(writer.sheets['Exact_Matches'])
                        sheets_created += 1
                
                # Category patterns sheet
                if enable_category_patterns and 'category_patterns' in enhanced_results:
                    category_data = self._convert_category_patterns_to_dataframe(enhanced_results['category_patterns'])
                    if not category_data.empty:
                        category_data.to_excel(writer, sheet_name='Category_Patterns', index=False)
                        self._adjust_excel_columns(writer.sheets['Category_Patterns'])
                        sheets_created += 1
                
                # Enhanced summary sheet
                self._add_enhanced_summary_sheet(writer, enhanced_results, 
                                                enable_same_day, enable_rapid_fire,
                                                enable_exact_match, enable_category_patterns)
                sheets_created += 1
            
            return True, f"Successfully exported {sheets_created} analysis sheets to Excel file."
            
        except ImportError:
            return False, "Excel export requires openpyxl. Please install it: pip install openpyxl"
        except PermissionError:
            return False, "Permission denied. Please ensure the file is not open in another application."
        except Exception as e:
            return False, f"Enhanced Excel export failed: {str(e)}"
    
    def _convert_fuzzy_results_to_dataframe(self, fuzzy_results) -> pd.DataFrame:
        """Convert fuzzy matching results to DataFrame (handles both new list format and legacy dict format)."""
        all_results = []
        
        if isinstance(fuzzy_results, list):
            # New format: list of duplicate pairs
            for duplicate in fuzzy_results:
                all_results.append({
                    'Site': duplicate['site'],
                    'Ticket_1_Number': duplicate['ticket1_number'],
                    'Ticket_1_Description': duplicate['ticket1_description'],
                    'Ticket_1_Created': duplicate['ticket1_created'],
                    'Ticket_2_Number': duplicate['ticket2_number'],
                    'Ticket_2_Description': duplicate['ticket2_description'],
                    'Ticket_2_Created': duplicate['ticket2_created'],
                    'Time_Difference': duplicate['time_difference_formatted'],
                    'Time_Difference_Hours': duplicate.get('time_difference_hours', 0),
                    'Time_Category': duplicate.get('time_category', 'N/A'),
                    'Similarity_Score': duplicate['similarity_score']
                })
        else:
            # Legacy format: dictionary with time windows
            for time_window, duplicates in fuzzy_results.items():
                for duplicate in duplicates:
                    all_results.append({
                        'Time_Window_Hours': duplicate.get('time_window_hours', time_window),
                        'Site': duplicate['site'],
                        'Ticket_1_Number': duplicate['ticket1_number'],
                        'Ticket_1_Description': duplicate['ticket1_description'],
                        'Ticket_1_Created': duplicate['ticket1_created'],
                        'Ticket_2_Number': duplicate['ticket2_number'],
                        'Ticket_2_Description': duplicate['ticket2_description'],
                        'Ticket_2_Created': duplicate['ticket2_created'],
                        'Time_Difference': duplicate['time_difference_formatted'],
                        'Time_Difference_Hours': duplicate.get('time_difference_hours', 0),
                        'Time_Category': duplicate.get('time_category', 'N/A'),
                        'Similarity_Score': duplicate['similarity_score']
                    })
        
        return pd.DataFrame(all_results)
    
    def _convert_same_day_to_dataframe(self, same_day_results: list) -> pd.DataFrame:
        """Convert same-day results to DataFrame."""
        return pd.DataFrame([{
            'Site': result['site'],
            'Date': result['date'],
            'Ticket_Count': result['ticket_count'],
            'Ticket_Numbers': result['ticket_numbers'],
            'Category_Mix': result['category_mix'],
            'Priority_Mix': result['priority_mix'],
            'Time_Span': result['time_span'],
            'Earliest_Time': result['earliest_time'],
            'Latest_Time': result['latest_time']
        } for result in same_day_results])
    
    def _convert_rapid_fire_to_dataframe(self, rapid_fire_results: list) -> pd.DataFrame:
        """Convert rapid-fire results to DataFrame."""
        return pd.DataFrame([{
            'Site': result['site'],
            'Time_Window_Minutes': result['time_window_minutes'],
            'Ticket_1': result['ticket_1'],
            'Ticket_1_Description': result['ticket_1_description'],
            'Ticket_1_Created': result['ticket_1_created'],
            'Ticket_2': result['ticket_2'],
            'Ticket_2_Description': result['ticket_2_description'],
            'Ticket_2_Created': result['ticket_2_created'],
            'Time_Difference': result['time_difference'],
            'Similarity_Score': result['similarity_score']
        } for result in rapid_fire_results])
    
    def _convert_exact_match_to_dataframe(self, exact_match_results: list) -> pd.DataFrame:
        """Convert exact match results to DataFrame."""
        return pd.DataFrame([{
            'Site': result['site'],
            'Description': result['description'],
            'Ticket_Count': result['ticket_count'],
            'Ticket_Numbers': result['ticket_numbers'],
            'Date_Range': result['date_range'],
            'Category': result['category']
        } for result in exact_match_results])
    
    def _convert_category_patterns_to_dataframe(self, category_results: list) -> pd.DataFrame:
        """Convert category patterns results to DataFrame."""
        return pd.DataFrame([{
            'Site': result['site'],
            'Date': result['date'],
            'Category': result['category'],
            'Subcategory': result['subcategory'],
            'Ticket_Count': result['ticket_count'],
            'Ticket_Numbers': result['ticket_numbers'],
            'Priority_Distribution': result['priority_distribution']
        } for result in category_results])
    
    def _add_enhanced_summary_sheet(self, writer, enhanced_results: dict,
                                   enable_same_day: bool, enable_rapid_fire: bool,
                                   enable_exact_match: bool, enable_category_patterns: bool):
        """Add an enhanced summary sheet with statistics from all analysis types."""
        try:
            summary_data = []
            
            # Primary fuzzy matching summary
            fuzzy_results = enhanced_results.get('fuzzy_matching', {})
            total_fuzzy_pairs = sum(len(duplicates) for duplicates in fuzzy_results.values())
            
            if total_fuzzy_pairs > 0:
                summary_data.append({
                    'Analysis_Type': 'Fuzzy Matching Duplicates',
                    'Total_Pairs_or_Groups': total_fuzzy_pairs,
                    'Description': 'Similar tickets within time windows'
                })
                
                # Time window breakdown
                for window, duplicates in fuzzy_results.items():
                    if duplicates:
                        summary_data.append({
                            'Analysis_Type': f'  - Within {window} hours',
                            'Total_Pairs_or_Groups': len(duplicates),
                            'Description': f'Duplicates found within {window} hour window'
                        })
            
            # Same-day duplicates summary
            if enable_same_day and 'same_day' in enhanced_results:
                same_day_count = len(enhanced_results['same_day'])
                if same_day_count > 0:
                    summary_data.append({
                        'Analysis_Type': 'Same-Day Duplicates',
                        'Total_Pairs_or_Groups': same_day_count,
                        'Description': 'Multiple tickets created on same calendar day'
                    })
            
            # Rapid-fire duplicates summary
            if enable_rapid_fire and 'rapid_fire' in enhanced_results:
                rapid_fire_count = len(enhanced_results['rapid_fire'])
                if rapid_fire_count > 0:
                    summary_data.append({
                        'Analysis_Type': 'Rapid-Fire Duplicates',
                        'Total_Pairs_or_Groups': rapid_fire_count,
                        'Description': 'Similar tickets within 15-60 minutes'
                    })
            
            # Exact matches summary
            if enable_exact_match and 'exact_match' in enhanced_results:
                exact_match_count = len(enhanced_results['exact_match'])
                if exact_match_count > 0:
                    summary_data.append({
                        'Analysis_Type': 'Exact Matches',
                        'Total_Pairs_or_Groups': exact_match_count,
                        'Description': 'Tickets with identical descriptions'
                    })
            
            # Category patterns summary
            if enable_category_patterns and 'category_patterns' in enhanced_results:
                category_count = len(enhanced_results['category_patterns'])
                if category_count > 0:
                    summary_data.append({
                        'Analysis_Type': 'Category Patterns',
                        'Total_Pairs_or_Groups': category_count,
                        'Description': 'Same category/subcategory on same day'
                    })
            
            # Create and export summary DataFrame
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Analysis_Summary', index=False)
                self._adjust_excel_columns(writer.sheets['Analysis_Summary'])
            
        except Exception:
            # If enhanced summary creation fails, continue without it
            pass