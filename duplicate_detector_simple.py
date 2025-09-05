import pandas as pd
from datetime import timedelta
from typing import List, Dict, Tuple, Callable
import difflib

class SimpleDuplicateDetector:
    """Simple duplicate detector using Python's built-in difflib for string similarity."""
    
    def __init__(self, progress_callback: Callable[[str, int, int], None] = None):
        self.progress_callback = progress_callback
        self.results = {}
        
    def analyze(self, data: pd.DataFrame, time_windows: List[int], similarity_threshold: int = 85) -> Dict[int, List[Dict]]:
        """
        Analyze data for potential duplicates using built-in string similarity.
        """
        self.results = {window: [] for window in time_windows}
        
        if data.empty:
            return self.results
            
        # Group data by site
        sites = data.groupby('Site')
        total_sites = len(sites)
        
        for site_idx, (site_name, site_data) in enumerate(sites):
            if self.progress_callback:
                self.progress_callback(f"Processing site {site_idx + 1} of {total_sites}: {site_name}", 
                                     site_idx + 1, total_sites)
            
            # Sort by creation time for efficient comparison
            site_data_sorted = site_data.sort_values('Created_dt').reset_index(drop=True)
            
            # Find duplicates within this site for each time window
            for time_window in time_windows:
                site_duplicates = self._find_duplicates_in_site(
                    site_data_sorted, time_window, similarity_threshold
                )
                self.results[time_window].extend(site_duplicates)
        
        # Sort results by similarity score (descending)
        for time_window in self.results:
            self.results[time_window].sort(key=lambda x: x['similarity_score'], reverse=True)
            
        return self.results
    
    def _find_duplicates_in_site(self, site_data: pd.DataFrame, time_window_hours: int, 
                                 similarity_threshold: int) -> List[Dict]:
        """Find potential duplicates within a single site for a specific time window."""
        duplicates = []
        time_window_delta = timedelta(hours=time_window_hours)
        
        # Compare each ticket with subsequent tickets within the time window
        for i in range(len(site_data)):
            ticket1 = site_data.iloc[i]
            
            # Only compare with tickets created after this one and within time window
            for j in range(i + 1, len(site_data)):
                ticket2 = site_data.iloc[j]
                
                # Check if ticket2 is within the time window
                time_diff = ticket2['Created_dt'] - ticket1['Created_dt']
                if time_diff > time_window_delta:
                    break  # No more tickets within time window (data is sorted)
                
                # Calculate similarity
                similarity_score = self._calculate_similarity(
                    ticket1['Short description'], 
                    ticket2['Short description']
                )
                
                if similarity_score >= similarity_threshold:
                    duplicate_pair = {
                        'site': ticket1['Site'],
                        'ticket1_number': ticket1['Number'],
                        'ticket1_description': ticket1['Short description'],
                        'ticket1_created': ticket1['Created'],
                        'ticket1_created_dt': ticket1['Created_dt'],
                        'ticket2_number': ticket2['Number'],
                        'ticket2_description': ticket2['Short description'],
                        'ticket2_created': ticket2['Created'],
                        'ticket2_created_dt': ticket2['Created_dt'],
                        'time_difference': time_diff,
                        'time_difference_formatted': self._format_time_difference(time_diff),
                        'similarity_score': similarity_score,
                        'time_window_hours': time_window_hours
                    }
                    duplicates.append(duplicate_pair)
        
        return duplicates
    
    def _calculate_similarity(self, desc1: str, desc2: str) -> int:
        """Calculate similarity using Python's difflib."""
        if pd.isna(desc1) or pd.isna(desc2):
            return 0
        
        desc1_clean = str(desc1).lower().strip()
        desc2_clean = str(desc2).lower().strip()
        
        if not desc1_clean or not desc2_clean:
            return 0
        
        # Use difflib.SequenceMatcher for similarity ratio
        similarity = difflib.SequenceMatcher(None, desc1_clean, desc2_clean).ratio()
        
        return int(similarity * 100)
    
    def _format_time_difference(self, time_diff: timedelta) -> str:
        """Format time difference as H:M:S string."""
        total_seconds = int(time_diff.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    
    def get_summary_stats(self) -> Dict[int, Dict]:
        """Get summary statistics for each time window."""
        stats = {}
        
        for time_window, duplicates in self.results.items():
            unique_sites = set()
            unique_tickets = set()
            
            for dup in duplicates:
                unique_sites.add(dup['site'])
                unique_tickets.add(dup['ticket1_number'])
                unique_tickets.add(dup['ticket2_number'])
            
            stats[time_window] = {
                'total_pairs': len(duplicates),
                'affected_sites': len(unique_sites),
                'unique_tickets_involved': len(unique_tickets),
                'avg_similarity': sum(dup['similarity_score'] for dup in duplicates) / len(duplicates) if duplicates else 0
            }
        
        return stats
    
    def export_results(self, time_windows: List[int] = None) -> pd.DataFrame:
        """Export results as a pandas DataFrame."""
        all_results = []
        
        windows_to_export = time_windows if time_windows is not None else list(self.results.keys())
        
        for time_window in windows_to_export:
            for duplicate in self.results[time_window]:
                all_results.append({
                    'Time_Window_Hours': duplicate['time_window_hours'],
                    'Site': duplicate['site'],
                    'Ticket_1_Number': duplicate['ticket1_number'],
                    'Ticket_1_Description': duplicate['ticket1_description'],
                    'Ticket_1_Created': duplicate['ticket1_created'],
                    'Ticket_2_Number': duplicate['ticket2_number'],
                    'Ticket_2_Description': duplicate['ticket2_description'],
                    'Ticket_2_Created': duplicate['ticket2_created'],
                    'Time_Difference': duplicate['time_difference_formatted'],
                    'Similarity_Score': duplicate['similarity_score']
                })
        
        return pd.DataFrame(all_results)