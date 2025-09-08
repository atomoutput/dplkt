import pandas as pd
from datetime import timedelta
from typing import List, Dict, Tuple, Callable
import itertools

try:
    from fuzzywuzzy import fuzz
    HAS_FUZZYWUZZY = True
except ImportError:
    import difflib
    HAS_FUZZYWUZZY = False
    print("Warning: fuzzywuzzy not available. Using difflib for string similarity.")
    print("For better performance, install fuzzywuzzy: pip install fuzzywuzzy python-Levenshtein")

class DuplicateDetector:
    """Core engine for detecting potential duplicate tickets using fuzzy string matching and time windows."""
    
    def __init__(self, progress_callback: Callable[[str, int, int], None] = None):
        """
        Initialize the duplicate detector.
        
        Args:
            progress_callback: Optional callback function for progress updates (message, current, total)
        """
        self.progress_callback = progress_callback
        self.results = {}
        
    def analyze(self, data: pd.DataFrame, max_hours: int, similarity_threshold: int = 85) -> List[Dict]:
        """
        Analyze data for potential duplicates within a maximum timeframe (no duplicates across windows).
        
        Args:
            data: DataFrame with ticket data
            max_hours: Maximum time window in hours to search for duplicates
            similarity_threshold: Minimum similarity percentage (0-100)
            
        Returns:
            List of duplicate pairs (each pair appears only once)
        """
        if data.empty:
            return []
            
        all_duplicates = []
        
        # Group data by site
        sites = data.groupby('Site')
        total_sites = len(sites)
        
        for site_idx, (site_name, site_data) in enumerate(sites):
            if self.progress_callback:
                self.progress_callback(f"Processing site {site_idx + 1} of {total_sites}: {site_name}", 
                                     site_idx + 1, total_sites)
            
            # Sort by creation time for efficient comparison
            site_data_sorted = site_data.sort_values('Created_dt').reset_index(drop=True)
            
            # Find duplicates within this site
            site_duplicates = self._find_duplicates_in_site_max_timeframe(
                site_data_sorted, max_hours, similarity_threshold
            )
            all_duplicates.extend(site_duplicates)
        
        # Sort results by similarity score (descending)
        all_duplicates.sort(key=lambda x: x['similarity_score'], reverse=True)
            
        return all_duplicates
    
    def analyze_legacy(self, data: pd.DataFrame, time_windows: List[int], similarity_threshold: int = 85) -> Dict[int, List[Dict]]:
        """
        Legacy analyze method for backward compatibility (creates duplicate pairs across windows).
        
        Args:
            data: DataFrame with ticket data
            time_windows: List of time windows in hours
            similarity_threshold: Minimum similarity percentage (0-100)
            
        Returns:
            Dictionary with time windows as keys and lists of duplicate pairs as values
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
    
    def analyze_enhanced(self, data: pd.DataFrame, max_hours: int, similarity_threshold: int = 85, 
                        enable_same_day: bool = False, enable_rapid_fire: bool = False, 
                        enable_exact_match: bool = False, enable_category_patterns: bool = False) -> Dict:
        """
        Enhanced analysis with multiple detection approaches using maximum timeframe.
        
        Args:
            data: DataFrame with ticket data
            max_hours: Maximum time window in hours for fuzzy matching
            similarity_threshold: Minimum similarity percentage (0-100)
            enable_same_day: Enable same-day duplicate detection
            enable_rapid_fire: Enable rapid-fire (15-60min) detection
            enable_exact_match: Enable exact content match detection
            enable_category_patterns: Enable category clustering analysis
            
        Returns:
            Dictionary with different analysis types and their results
        """
        enhanced_results = {}
        
        # Primary fuzzy matching analysis (new single timeframe approach)
        enhanced_results['fuzzy_matching'] = self.analyze(data, max_hours, similarity_threshold)
        
        if data.empty:
            return enhanced_results
        
        # Additional analysis methods
        if enable_same_day:
            enhanced_results['same_day'] = self._analyze_same_day_duplicates(data)
            
        if enable_rapid_fire:
            enhanced_results['rapid_fire'] = self._analyze_rapid_fire_duplicates(data, similarity_threshold)
            
        if enable_exact_match:
            enhanced_results['exact_match'] = self._analyze_exact_matches(data)
            
        if enable_category_patterns:
            enhanced_results['category_patterns'] = self._analyze_category_patterns(data)
        
        return enhanced_results
    
    def _analyze_same_day_duplicates(self, data: pd.DataFrame) -> List[Dict]:
        """
        Find tickets created on the same calendar day regardless of description similarity.
        
        Args:
            data: DataFrame with ticket data
            
        Returns:
            List of same-day duplicate groups
        """
        if self.progress_callback:
            self.progress_callback("Analyzing same-day duplicates", 1, 4)
            
        results = []
        
        # Group by site and date
        data['Created_date'] = data['Created_dt'].dt.date
        
        for site_name, site_data in data.groupby('Site'):
            date_groups = site_data.groupby('Created_date')
            
            for date, day_tickets in date_groups:
                if len(day_tickets) > 1:  # Multiple tickets on same day
                    tickets = day_tickets.sort_values('Created_dt')
                    
                    # Calculate time span
                    earliest_time = tickets['Created_dt'].min()
                    latest_time = tickets['Created_dt'].max()
                    time_span = latest_time - earliest_time
                    
                    # Get category mix (handle missing columns gracefully)
                    categories = tickets['Category'].value_counts() if 'Category' in tickets.columns else pd.Series()
                    priorities = tickets['Priority'].value_counts() if 'Priority' in tickets.columns else pd.Series()
                    
                    result = {
                        'site': site_name,
                        'date': date.strftime('%Y-%m-%d'),
                        'ticket_count': len(tickets),
                        'ticket_numbers': ', '.join(tickets['Number'].astype(str)),
                        'category_mix': ', '.join([f"{cat}({count})" for cat, count in categories.items()]),
                        'priority_mix': ', '.join([f"{pri}({count})" for pri, count in priorities.items()]),
                        'time_span': self._format_time_difference(time_span),
                        'earliest_time': earliest_time.strftime('%H:%M:%S'),
                        'latest_time': latest_time.strftime('%H:%M:%S')
                    }
                    results.append(result)
        
        return sorted(results, key=lambda x: (x['site'], x['date']))
    
    def _analyze_rapid_fire_duplicates(self, data: pd.DataFrame, similarity_threshold: int) -> List[Dict]:
        """
        Find tickets created within rapid-fire time windows (15, 30, 60 minutes).
        
        Args:
            data: DataFrame with ticket data
            similarity_threshold: Minimum similarity percentage
            
        Returns:
            List of rapid-fire duplicate pairs
        """
        if self.progress_callback:
            self.progress_callback("Analyzing rapid-fire duplicates", 2, 4)
            
        rapid_windows = [0.25, 0.5, 1]  # 15, 30, 60 minutes in hours
        results = []
        
        for site_name, site_data in data.groupby('Site'):
            site_data_sorted = site_data.sort_values('Created_dt').reset_index(drop=True)
            
            for window_hours in rapid_windows:
                window_delta = timedelta(hours=window_hours)
                window_minutes = int(window_hours * 60)
                
                for i in range(len(site_data_sorted)):
                    ticket1 = site_data_sorted.iloc[i]
                    
                    # Skip if ticket1 has invalid datetime
                    if pd.isna(ticket1['Created_dt']):
                        continue
                    
                    for j in range(i + 1, len(site_data_sorted)):
                        ticket2 = site_data_sorted.iloc[j]
                        
                        # Skip if ticket2 has invalid datetime
                        if pd.isna(ticket2['Created_dt']):
                            continue
                        
                        time_diff = ticket2['Created_dt'] - ticket1['Created_dt']
                        if time_diff > window_delta:
                            break
                        
                        # Calculate similarity
                        similarity_score = self._calculate_similarity(
                            ticket1['Short description'], 
                            ticket2['Short description']
                        )
                        
                        if similarity_score >= similarity_threshold:
                            result = {
                                'site': site_name,
                                'time_window_minutes': window_minutes,
                                'ticket_1': ticket1['Number'],
                                'ticket_1_description': ticket1['Short description'],
                                'ticket_1_created': ticket1['Created'],
                                'ticket_2': ticket2['Number'],
                                'ticket_2_description': ticket2['Short description'],
                                'ticket_2_created': ticket2['Created'],
                                'time_difference': self._format_time_difference(time_diff),
                                'similarity_score': similarity_score
                            }
                            results.append(result)
        
        return sorted(results, key=lambda x: x['similarity_score'], reverse=True)
    
    def _analyze_exact_matches(self, data: pd.DataFrame) -> List[Dict]:
        """
        Find tickets with 100% identical descriptions regardless of time.
        
        Args:
            data: DataFrame with ticket data
            
        Returns:
            List of exact match groups
        """
        if self.progress_callback:
            self.progress_callback("Analyzing exact matches", 3, 4)
            
        results = []
        
        # Group by site and exact description
        for site_name, site_data in data.groupby('Site'):
            description_groups = site_data.groupby('Short description')
            
            for description, desc_tickets in description_groups:
                if len(desc_tickets) > 1 and pd.notna(description) and str(description).strip():
                    tickets = desc_tickets.sort_values('Created_dt')
                    
                    # Calculate date range
                    earliest_date = tickets['Created_dt'].min().strftime('%Y-%m-%d %H:%M')
                    latest_date = tickets['Created_dt'].max().strftime('%Y-%m-%d %H:%M')
                    date_range = f"{earliest_date} to {latest_date}" if earliest_date != latest_date else earliest_date
                    
                    # Get category info (handle missing columns gracefully)
                    if 'Category' in tickets.columns:
                        categories = tickets['Category'].unique()
                        category_info = ', '.join(categories) if len(categories) <= 3 else f"{categories[0]} (+{len(categories)-1} more)"
                    else:
                        category_info = 'N/A'
                    
                    result = {
                        'site': site_name,
                        'description': str(description)[:100] + '...' if len(str(description)) > 100 else str(description),
                        'ticket_count': len(tickets),
                        'ticket_numbers': ', '.join(tickets['Number'].astype(str)),
                        'date_range': date_range,
                        'category': category_info
                    }
                    results.append(result)
        
        return sorted(results, key=lambda x: x['ticket_count'], reverse=True)
    
    def _analyze_category_patterns(self, data: pd.DataFrame) -> List[Dict]:
        """
        Find same category/subcategory combinations on the same day.
        
        Args:
            data: DataFrame with ticket data
            
        Returns:
            List of category pattern groups
        """
        if self.progress_callback:
            self.progress_callback("Analyzing category patterns", 4, 4)
            
        results = []
        data['Created_date'] = data['Created_dt'].dt.date
        
        # Group by site, date, category, and subcategory (handle missing columns)
        for site_name, site_data in data.groupby('Site'):
            # Only group by available columns
            group_cols = ['Created_date']
            if 'Category' in site_data.columns:
                group_cols.append('Category')
            if 'Subcategory' in site_data.columns:
                group_cols.append('Subcategory')
            
            if len(group_cols) == 1:
                # Skip if no category columns available
                continue
                
            pattern_groups = site_data.groupby(group_cols)
            
            for group_key, pattern_tickets in pattern_groups:
                # Handle different grouping scenarios
                if len(group_cols) == 3:  # date, category, subcategory
                    date, category, subcategory = group_key
                elif len(group_cols) == 2:  # date, category
                    date, category = group_key
                    subcategory = 'N/A'
                else:  # just date (shouldn't happen due to continue above)
                    continue
                if len(pattern_tickets) > 1:  # Multiple tickets with same pattern
                    tickets = pattern_tickets.sort_values('Created_dt')
                    
                    # Get priority distribution (handle missing columns gracefully)
                    if 'Priority' in tickets.columns:
                        priorities = tickets['Priority'].value_counts()
                        priority_dist = ', '.join([f"{pri}({count})" for pri, count in priorities.items()])
                    else:
                        priority_dist = 'N/A'
                    
                    result = {
                        'site': site_name,
                        'date': date.strftime('%Y-%m-%d'),
                        'category': str(category) if pd.notna(category) else 'N/A',
                        'subcategory': str(subcategory) if pd.notna(subcategory) else 'N/A',
                        'ticket_count': len(tickets),
                        'ticket_numbers': ', '.join(tickets['Number'].astype(str)),
                        'priority_distribution': priority_dist
                    }
                    results.append(result)
        
        return sorted(results, key=lambda x: (x['site'], x['date'], x['ticket_count']), reverse=True)
    
    def _find_duplicates_in_site_max_timeframe(self, site_data: pd.DataFrame, max_hours: int, 
                                             similarity_threshold: int) -> List[Dict]:
        """
        Find potential duplicates within a single site using maximum timeframe (no duplicate pairs).
        
        Args:
            site_data: DataFrame containing tickets for one site, sorted by Created_dt
            max_hours: Maximum time window in hours
            similarity_threshold: Minimum similarity percentage
            
        Returns:
            List of unique duplicate pair dictionaries
        """
        duplicates = []
        max_time_delta = timedelta(hours=max_hours)
        
        # Compare each ticket with subsequent tickets within the max timeframe
        for i in range(len(site_data)):
            ticket1 = site_data.iloc[i]
            
            # Skip if ticket1 has invalid datetime
            if pd.isna(ticket1['Created_dt']):
                continue
            
            # Only compare with tickets created after this one and within max timeframe
            for j in range(i + 1, len(site_data)):
                ticket2 = site_data.iloc[j]
                
                # Skip if ticket2 has invalid datetime
                if pd.isna(ticket2['Created_dt']):
                    continue
                
                # Check if ticket2 is within the max timeframe
                time_diff = ticket2['Created_dt'] - ticket1['Created_dt']
                if time_diff > max_time_delta:
                    break  # No more tickets within timeframe (data is sorted)
                
                # Calculate similarity
                similarity_score = self._calculate_similarity(
                    ticket1['Short description'], 
                    ticket2['Short description']
                )
                
                if similarity_score >= similarity_threshold:
                    # Categorize the time difference for grouping
                    time_category = self._categorize_time_difference(time_diff)
                    
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
                        'time_difference_hours': time_diff.total_seconds() / 3600,
                        'time_category': time_category,
                        'similarity_score': similarity_score
                    }
                    duplicates.append(duplicate_pair)
        
        return duplicates
    
    def _categorize_time_difference(self, time_diff: timedelta) -> str:
        """
        Categorize time difference for easy grouping and analysis.
        
        Args:
            time_diff: Time difference as timedelta
            
        Returns:
            Category string for grouping
        """
        hours = time_diff.total_seconds() / 3600
        
        if hours <= 1:
            return "0-1h"
        elif hours <= 4:
            return "1-4h"
        elif hours <= 8:
            return "4-8h"
        elif hours <= 24:
            return "8-24h"
        elif hours <= 72:
            return "1-3d"
        elif hours <= 168:
            return "3-7d"
        else:
            return ">7d"
    
    def _find_duplicates_in_site(self, site_data: pd.DataFrame, time_window_hours: int, 
                                 similarity_threshold: int) -> List[Dict]:
        """
        Find potential duplicates within a single site for a specific time window.
        
        Args:
            site_data: DataFrame containing tickets for one site, sorted by Created_dt
            time_window_hours: Time window in hours
            similarity_threshold: Minimum similarity percentage
            
        Returns:
            List of duplicate pair dictionaries
        """
        duplicates = []
        time_window_delta = timedelta(hours=time_window_hours)
        
        # Compare each ticket with subsequent tickets within the time window
        for i in range(len(site_data)):
            ticket1 = site_data.iloc[i]
            
            # Skip if ticket1 has invalid datetime
            if pd.isna(ticket1['Created_dt']):
                continue
            
            # Only compare with tickets created after this one and within time window
            for j in range(i + 1, len(site_data)):
                ticket2 = site_data.iloc[j]
                
                # Skip if ticket2 has invalid datetime
                if pd.isna(ticket2['Created_dt']):
                    continue
                
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
        """
        Calculate similarity between two descriptions using fuzzy string matching.
        
        Args:
            desc1: First description
            desc2: Second description
            
        Returns:
            Similarity score (0-100)
        """
        if pd.isna(desc1) or pd.isna(desc2):
            return 0
        
        if HAS_FUZZYWUZZY:
            # Use fuzz.ratio for overall similarity, but also consider partial matches
            ratio_score = fuzz.ratio(str(desc1).lower(), str(desc2).lower())
            partial_score = fuzz.partial_ratio(str(desc1).lower(), str(desc2).lower())
            
            # Take the higher of the two scores to catch both exact and partial matches
            return max(ratio_score, partial_score)
        else:
            # Fallback to difflib
            desc1_clean = str(desc1).lower().strip()
            desc2_clean = str(desc2).lower().strip()
            
            if not desc1_clean or not desc2_clean:
                return 0
            
            similarity = difflib.SequenceMatcher(None, desc1_clean, desc2_clean).ratio()
            return int(similarity * 100)
    
    def _format_time_difference(self, time_diff: timedelta) -> str:
        """
        Format time difference as H:M:S string.
        
        Args:
            time_diff: Time difference as timedelta
            
        Returns:
            Formatted string (e.g., "2:30:15")
        """
        total_seconds = int(time_diff.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    
    def get_summary_stats(self) -> Dict[int, Dict]:
        """
        Get summary statistics for each time window.
        
        Returns:
            Dictionary with time windows as keys and stats as values
        """
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
    
    def export_results_new(self, duplicates: List[Dict] = None) -> pd.DataFrame:
        """
        Export new format results as a pandas DataFrame for CSV/Excel export.
        
        Args:
            duplicates: List of duplicate pairs (None to use internal results)
            
        Returns:
            DataFrame with all duplicate pairs (no time window duplicates)
        """
        if duplicates is None:
            duplicates = getattr(self, 'current_results', [])
        
        all_results = []
        
        for duplicate in duplicates:
            all_results.append({
                'Site': duplicate['site'],
                'Ticket_1_Number': duplicate['ticket1_number'],
                'Ticket_1_Description': duplicate['ticket1_description'],
                'Ticket_1_Created': duplicate['ticket1_created'],
                'Ticket_2_Number': duplicate['ticket2_number'],
                'Ticket_2_Description': duplicate['ticket2_description'],
                'Ticket_2_Created': duplicate['ticket2_created'],
                'Time_Difference': duplicate['time_difference_formatted'],
                'Time_Difference_Hours': duplicate['time_difference_hours'],
                'Time_Category': duplicate['time_category'],
                'Similarity_Score': duplicate['similarity_score']
            })
        
        return pd.DataFrame(all_results)
    
    def export_results(self, time_windows: List[int] = None) -> pd.DataFrame:
        """
        Export results as a pandas DataFrame for CSV/Excel export.
        
        Args:
            time_windows: List of time windows to include (None for all)
            
        Returns:
            DataFrame with all duplicate pairs
        """
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