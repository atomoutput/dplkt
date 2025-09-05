#!/usr/bin/env python3
"""
ServiceNow Duplicate Ticket Detection Tool - Command Line Interface
Optimized for Termux/Android environment
"""

import argparse
import os
import sys
from csv_parser import CSVParser
from duplicate_detector import DuplicateDetector
from export_manager import ExportManager

class DuplicateTicketCLI:
    """Command-line interface for duplicate ticket detection."""
    
    def __init__(self):
        self.csv_parser = CSVParser(repair_callback=self.repair_progress_callback)
        self.duplicate_detector = None
        self.export_manager = ExportManager()
    
    def run(self, args):
        """Main execution flow."""
        print("ServiceNow Duplicate Ticket Detection Tool")
        print("=" * 50)
        
        # Load and validate CSV (with optional repair)
        print(f"Loading CSV file: {args.input}")
        success, message = self.csv_parser.load_and_validate(args.input, auto_repair=not args.no_auto_repair)
        
        if not success:
            print(f"Error: {message}")
            return 1
        
        print(f"✓ {message}")
        
        # Show data summary
        summary = self.csv_parser.get_data_summary()
        print(f"  - Total tickets: {summary['total_tickets']}")
        print(f"  - Unique sites: {summary['unique_sites']}")
        print(f"  - Date range: {summary['date_range']['earliest'].strftime('%Y-%m-%d')} to {summary['date_range']['latest'].strftime('%Y-%m-%d')}")
        if summary.get('resolved_tickets', 0) > 0:
            print(f"  - Resolved tickets: {summary['resolved_tickets']}")
        if summary.get('was_repaired', False):
            print("  ⚠️ File was auto-repaired due to corruption")
        print()
        
        # Parse time windows
        try:
            time_windows = [int(w.strip()) for w in args.time_windows.split(',') if w.strip()]
            if not time_windows:
                raise ValueError("No time windows specified")
            print(f"Time windows: {time_windows} hours")
        except ValueError as e:
            print(f"Error parsing time windows: {e}")
            return 1
        
        # Get filtered data
        data = self.csv_parser.get_filtered_data(args.exclude_resolved)
        if args.exclude_resolved:
            print(f"Excluding resolved tickets. Analyzing {len(data)} tickets.")
        else:
            print(f"Including all tickets. Analyzing {len(data)} tickets.")
        
        if data.empty:
            print("No tickets to analyze after applying filters.")
            return 1
        
        print(f"Similarity threshold: {args.similarity}%")
        print()
        
        # Run analysis
        print("Running duplicate detection analysis...")
        self.duplicate_detector = DuplicateDetector(self.progress_callback)
        
        results = self.duplicate_detector.analyze(data, time_windows, args.similarity)
        
        # Display results
        self.display_results(results)
        
        # Export if requested
        if args.output:
            print(f"\nExporting results to: {args.output}")
            df = self.duplicate_detector.export_results()
            success, message = self.export_manager.export_data(df, args.output)
            
            if success:
                print(f"✓ {message}")
            else:
                print(f"✗ Export failed: {message}")
                return 1
        
        # Cleanup temporary files
        self.csv_parser.cleanup_temp_files()
        
        return 0
    
    def progress_callback(self, message: str, current: int, total: int):
        """Handle progress updates."""
        percent = (current / total) * 100 if total > 0 else 0
        print(f"\r{message} [{percent:.1f}%]", end="", flush=True)
        
        if current == total:
            print()  # New line when complete
    
    def repair_progress_callback(self, message: str):
        """Handle repair progress updates."""
        print(f"  {message}")
    
    def repair_only_mode(self, args):
        """Run in repair-only mode to fix a corrupted CSV file."""
        print("CSV Repair Mode")
        print("=" * 30)
        
        if not os.path.exists(args.input):
            print(f"Error: File '{args.input}' does not exist.")
            return 1
        
        print(f"Repairing CSV file: {args.input}")
        
        success, message, output_path = self.csv_parser.manual_repair(
            args.input,
            create_backup=args.create_backup,
            target_encoding=getattr(args, 'encoding', 'utf-8'),
            overwrite=args.overwrite_original
        )
        
        if success:
            print(f"✓ {message}")
            if output_path:
                print(f"  Output file: {output_path}")
            return 0
        else:
            print(f"✗ Repair failed: {message}")
            return 1
    
    def display_results(self, results):
        """Display analysis results in terminal."""
        print("\nAnalysis Results")
        print("=" * 50)
        
        total_duplicates = sum(len(duplicates) for duplicates in results.values())
        
        if total_duplicates == 0:
            print("No potential duplicates found.")
            return
        
        print(f"Found {total_duplicates} potential duplicate pairs\n")
        
        # Display summary for each time window
        for time_window in sorted(results.keys()):
            duplicates = results[time_window]
            if not duplicates:
                continue
                
            print(f"Within {time_window} hours: {len(duplicates)} pairs")
            
            if args.verbose:
                print("-" * 40)
                
                # Show top matches
                top_matches = duplicates[:min(5, len(duplicates))]
                for i, dup in enumerate(top_matches, 1):
                    print(f"  {i}. {dup['similarity_score']}% similarity")
                    print(f"     Site: {dup['site']}")
                    print(f"     Ticket 1: {dup['ticket1_number']} ({dup['ticket1_created']})")
                    print(f"     Description: {dup['ticket1_description'][:80]}...")
                    print(f"     Ticket 2: {dup['ticket2_number']} ({dup['ticket2_created']})")
                    print(f"     Description: {dup['ticket2_description'][:80]}...")
                    print(f"     Time difference: {dup['time_difference_formatted']}")
                    print()
                
                if len(duplicates) > 5:
                    print(f"  ... and {len(duplicates) - 5} more pairs")
                print()
        
        # Show summary statistics
        stats = self.duplicate_detector.get_summary_stats()
        print("Summary Statistics:")
        print("-" * 20)
        
        for time_window, stat in stats.items():
            if stat['total_pairs'] > 0:
                print(f"Within {time_window}h:")
                print(f"  - Duplicate pairs: {stat['total_pairs']}")
                print(f"  - Affected sites: {stat['affected_sites']}")
                print(f"  - Unique tickets involved: {stat['unique_tickets_involved']}")
                print(f"  - Average similarity: {stat['avg_similarity']:.1f}%")

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Detect duplicate tickets from ServiceNow CSV exports",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s current.csv
  %(prog)s current.csv -t "1,8,24" -s 90 -o results.csv
  %(prog)s current.csv --exclude-resolved --verbose
  %(prog)s current.csv --no-auto-repair
  %(prog)s corrupted.csv --repair-only --create-backup
  %(prog)s corrupted.csv --repair-only --overwrite-original --encoding utf-8
        """
    )
    
    parser.add_argument("input", help="Input CSV file path")
    parser.add_argument("-t", "--time-windows", default="1,8,24,72",
                       help="Comma-separated time windows in hours (default: 1,8,24,72)")
    parser.add_argument("-s", "--similarity", type=int, default=85, 
                       help="Similarity threshold percentage (default: 85)")
    parser.add_argument("-o", "--output", help="Output file path for results (CSV or Excel)")
    parser.add_argument("--exclude-resolved", action="store_true",
                       help="Exclude tickets that have been resolved")
    parser.add_argument("-v", "--verbose", action="store_true",
                       help="Show detailed results")
    
    # Repair options
    parser.add_argument("--repair-only", action="store_true",
                       help="Only repair the CSV file without running analysis")
    parser.add_argument("--no-auto-repair", action="store_true", 
                       help="Disable automatic repair of corrupted CSV files")
    parser.add_argument("--create-backup", action="store_true", default=True,
                       help="Create backup when repairing (default: enabled)")
    parser.add_argument("--overwrite-original", action="store_true",
                       help="Overwrite the original file when repairing")
    parser.add_argument("--encoding", default="utf-8",
                       help="Target encoding for repaired files (default: utf-8)")
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' does not exist.")
        return 1
    
    # Validate similarity threshold
    if not 50 <= args.similarity <= 100:
        print("Error: Similarity threshold must be between 50 and 100.")
        return 1
    
    # Store verbose flag globally for display_results
    global verbose_flag
    verbose_flag = args.verbose
    
    try:
        cli = DuplicateTicketCLI()
        
        # Check if running in repair-only mode
        if args.repair_only:
            return cli.repair_only_mode(args)
        else:
            return cli.run(args)
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        return 1
    except Exception as e:
        print(f"Unexpected error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

if __name__ == "__main__":
    # Make verbose flag accessible to display_results method
    verbose_flag = False
    # Monkey patch to fix verbose access
    def display_results_with_verbose(self, results):
        global verbose_flag
        args = type('Args', (), {'verbose': verbose_flag})()
        return self.display_results_original(results, args)
    
    # Save original method
    DuplicateTicketCLI.display_results_original = DuplicateTicketCLI.display_results
    
    # Replace with patched version
    def patched_display_results(self, results):
        print("\nAnalysis Results")
        print("=" * 50)
        
        total_duplicates = sum(len(duplicates) for duplicates in results.values())
        
        if total_duplicates == 0:
            print("No potential duplicates found.")
            return
        
        print(f"Found {total_duplicates} potential duplicate pairs\n")
        
        # Display summary for each time window
        for time_window in sorted(results.keys()):
            duplicates = results[time_window]
            if not duplicates:
                continue
                
            print(f"Within {time_window} hours: {len(duplicates)} pairs")
            
            if verbose_flag:
                print("-" * 40)
                
                # Show top matches
                top_matches = duplicates[:min(5, len(duplicates))]
                for i, dup in enumerate(top_matches, 1):
                    print(f"  {i}. {dup['similarity_score']}% similarity")
                    print(f"     Site: {dup['site']}")
                    print(f"     Ticket 1: {dup['ticket1_number']} ({dup['ticket1_created']})")
                    print(f"     Description: {dup['ticket1_description'][:80]}...")
                    print(f"     Ticket 2: {dup['ticket2_number']} ({dup['ticket2_created']})")
                    print(f"     Description: {dup['ticket2_description'][:80]}...")
                    print(f"     Time difference: {dup['time_difference_formatted']}")
                    print()
                
                if len(duplicates) > 5:
                    print(f"  ... and {len(duplicates) - 5} more pairs")
                print()
        
        # Show summary statistics
        stats = self.duplicate_detector.get_summary_stats()
        print("Summary Statistics:")
        print("-" * 20)
        
        for time_window, stat in stats.items():
            if stat['total_pairs'] > 0:
                print(f"Within {time_window}h:")
                print(f"  - Duplicate pairs: {stat['total_pairs']}")
                print(f"  - Affected sites: {stat['affected_sites']}")
                print(f"  - Unique tickets involved: {stat['unique_tickets_involved']}")
                print(f"  - Average similarity: {stat['avg_similarity']:.1f}%")
    
    DuplicateTicketCLI.display_results = patched_display_results
    
    sys.exit(main())