#!/usr/bin/env python3
"""
ServiceNow Duplicate Ticket Detection Tool
Main GUI application using Tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from csv_parser import CSVParser
from duplicate_detector import DuplicateDetector
from export_manager import ExportManager

class DuplicateTicketApp:
    """Main application class for the duplicate ticket detection tool."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("ServiceNow Duplicate Ticket Detection Tool")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        # Initialize components
        self.csv_parser = CSVParser()
        self.duplicate_detector = None
        self.export_manager = ExportManager()
        
        # Data storage
        self.current_file_path = None
        self.analysis_results = {}
        
        # Create GUI
        self.setup_styles()
        self.create_widgets()
        self.create_layout()
        
        # Initial state
        self.update_ui_state()
    
    def setup_styles(self):
        """Configure ttk styles for better appearance."""
        style = ttk.Style()
        
        # Configure button styles
        style.configure('Action.TButton', padding=(10, 5))
        style.configure('Primary.TButton', padding=(15, 8))
        
        # Configure frame styles
        style.configure('Card.TFrame', relief='raised', borderwidth=1)
        style.configure('Section.TLabelframe', padding=(10, 10))
    
    def create_widgets(self):
        """Create all GUI widgets."""
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="10")
        
        # Top panel - File operations
        self.file_frame = ttk.LabelFrame(self.main_frame, text="File Operations", 
                                        style='Section.TLabelframe')
        
        self.load_button = ttk.Button(self.file_frame, text="Load CSV File",
                                     command=self.load_file, style='Primary.TButton')
        
        self.file_path_var = tk.StringVar(value="No file selected")
        self.file_path_label = ttk.Label(self.file_frame, textvariable=self.file_path_var,
                                        foreground='gray')
        
        # Left panel - Configuration
        self.config_frame = ttk.LabelFrame(self.main_frame, text="Analysis Parameters",
                                          style='Section.TLabelframe')
        
        # Time windows configuration
        ttk.Label(self.config_frame, text="Time Windows (hours):").pack(anchor='w', pady=(0, 5))
        self.time_windows_var = tk.StringVar(value="1, 8, 24, 72")
        self.time_windows_entry = ttk.Entry(self.config_frame, textvariable=self.time_windows_var,
                                           width=30)
        self.time_windows_entry.pack(fill='x', pady=(0, 10))
        
        # Similarity threshold
        ttk.Label(self.config_frame, text="Similarity Threshold:").pack(anchor='w', pady=(0, 5))
        
        self.similarity_frame = ttk.Frame(self.config_frame)
        self.similarity_var = tk.IntVar(value=85)
        self.similarity_scale = ttk.Scale(self.similarity_frame, from_=50, to=100,
                                         variable=self.similarity_var, orient='horizontal')
        self.similarity_scale.pack(side='left', fill='x', expand=True)
        
        self.similarity_label_var = tk.StringVar(value="85%")
        self.similarity_label = ttk.Label(self.similarity_frame, textvariable=self.similarity_label_var,
                                         width=6)
        self.similarity_label.pack(side='right', padx=(10, 0))
        self.similarity_scale.configure(command=self.update_similarity_label)
        
        self.similarity_frame.pack(fill='x', pady=(0, 15))
        
        # Options
        self.exclude_resolved_var = tk.BooleanVar(value=False)
        self.exclude_resolved_check = ttk.Checkbutton(self.config_frame, 
                                                     text="Exclude resolved tickets",
                                                     variable=self.exclude_resolved_var)
        self.exclude_resolved_check.pack(anchor='w', pady=(0, 15))
        
        # Run analysis button
        self.analyze_button = ttk.Button(self.config_frame, text="Run Analysis",
                                        command=self.run_analysis_threaded,
                                        style='Primary.TButton')
        self.analyze_button.pack(fill='x', pady=(0, 10))
        
        # Results panel
        self.results_frame = ttk.LabelFrame(self.main_frame, text="Results",
                                           style='Section.TLabelframe')
        
        # Notebook for tabbed results
        self.results_notebook = ttk.Notebook(self.results_frame)
        self.results_notebook.pack(fill='both', expand=True, pady=(10, 0))
        
        # Bottom panel - Status and actions
        self.bottom_frame = ttk.Frame(self.main_frame)
        
        self.status_var = tk.StringVar(value="Ready.")
        self.status_label = ttk.Label(self.bottom_frame, textvariable=self.status_var,
                                     foreground='blue')
        self.status_label.pack(side='left')
        
        self.export_button = ttk.Button(self.bottom_frame, text="Export Results",
                                       command=self.export_results, style='Action.TButton')
        self.export_button.pack(side='right')
        
        # Progress bar (initially hidden)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.bottom_frame, variable=self.progress_var,
                                           mode='determinate', length=300)
    
    def create_layout(self):
        """Arrange widgets in the main layout."""
        self.main_frame.pack(fill='both', expand=True)
        
        # Top panel
        self.file_frame.pack(fill='x', pady=(0, 10))
        self.load_button.pack(side='left', padx=(0, 10))
        self.file_path_label.pack(side='left', fill='x', expand=True)
        
        # Create main content area
        content_frame = ttk.Frame(self.main_frame)
        content_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Left panel (configuration)
        self.config_frame.pack(side='left', fill='y', padx=(0, 10))
        self.config_frame.configure(width=300)
        
        # Right panel (results)
        self.results_frame.pack(side='right', fill='both', expand=True)
        
        # Bottom panel
        self.bottom_frame.pack(fill='x')
    
    def update_similarity_label(self, value):
        """Update the similarity threshold label."""
        self.similarity_label_var.set(f"{int(float(value))}%")
    
    def load_file(self):
        """Load and validate CSV file."""
        file_path = filedialog.askopenfilename(
            title="Select ServiceNow CSV Export",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        self.status_var.set("Loading and validating file...")
        self.root.update()
        
        success, message = self.csv_parser.load_and_validate(file_path)
        
        if success:
            self.current_file_path = file_path
            self.file_path_var.set(os.path.basename(file_path))
            self.status_var.set(message)
            
            # Clear previous results
            self.clear_results()
            
            # Show data summary
            summary = self.csv_parser.get_data_summary()
            info_msg = (f"Loaded: {summary['total_tickets']} tickets\\n"
                       f"Sites: {summary['unique_sites']}\\n"
                       f"Date range: {summary['date_range']['earliest'].strftime('%Y-%m-%d')} "
                       f"to {summary['date_range']['latest'].strftime('%Y-%m-%d')}")
            
            if summary.get('resolved_tickets', 0) > 0:
                info_msg += f"\\nResolved tickets: {summary['resolved_tickets']}"
            
            messagebox.showinfo("File Loaded Successfully", info_msg)
            
        else:
            self.current_file_path = None
            self.file_path_var.set("No file selected")
            self.status_var.set("File loading failed.")
            messagebox.showerror("Error Loading File", message)
        
        self.update_ui_state()
    
    def parse_time_windows(self) -> list:
        """Parse time windows from the entry field."""
        try:
            windows_str = self.time_windows_var.get().strip()
            windows = [int(w.strip()) for w in windows_str.split(',') if w.strip()]
            
            if not windows:
                raise ValueError("No time windows specified")
            
            # Validate windows are positive
            if any(w <= 0 for w in windows):
                raise ValueError("Time windows must be positive numbers")
            
            return sorted(windows)  # Sort for consistent display
            
        except ValueError as e:
            raise ValueError(f"Invalid time windows format: {str(e)}")
    
    def run_analysis_threaded(self):
        """Run analysis in a separate thread to prevent UI freezing."""
        threading.Thread(target=self.run_analysis, daemon=True).start()
    
    def run_analysis(self):
        """Run the duplicate detection analysis."""
        try:
            # Parse time windows
            time_windows = self.parse_time_windows()
            
            # Get analysis parameters
            similarity_threshold = self.similarity_var.get()
            exclude_resolved = self.exclude_resolved_var.get()
            
            # Update UI state
            self.root.after(0, lambda: self.set_analysis_running(True))
            
            # Get filtered data
            data = self.csv_parser.get_filtered_data(exclude_resolved)
            
            if data.empty:
                self.root.after(0, lambda: messagebox.showwarning(
                    "No Data", "No tickets to analyze after applying filters."))
                return
            
            # Create detector with progress callback
            self.duplicate_detector = DuplicateDetector(self.progress_callback)
            
            # Run analysis
            self.analysis_results = self.duplicate_detector.analyze(
                data, time_windows, similarity_threshold
            )
            
            # Update UI with results
            self.root.after(0, self.display_results)
            
        except ValueError as e:
            self.root.after(0, lambda: messagebox.showerror("Configuration Error", str(e)))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Analysis Error", f"An error occurred: {str(e)}"))
        finally:
            self.root.after(0, lambda: self.set_analysis_running(False))
    
    def progress_callback(self, message: str, current: int, total: int):
        """Handle progress updates from the analysis."""
        progress_percent = (current / total) * 100 if total > 0 else 0
        
        self.root.after(0, lambda: [
            self.status_var.set(message),
            self.progress_var.set(progress_percent)
        ])
    
    def set_analysis_running(self, running: bool):
        """Update UI state during analysis."""
        if running:
            self.analyze_button.configure(state='disabled', text="Analyzing...")
            self.progress_bar.pack(side='left', padx=(20, 0))
            self.status_var.set("Starting analysis...")
        else:
            self.analyze_button.configure(state='normal', text="Run Analysis")
            self.progress_bar.pack_forget()
            if not self.analysis_results:
                self.status_var.set("Analysis completed.")
        
        self.update_ui_state()
    
    def display_results(self):
        """Display analysis results in the tabbed interface."""
        # Clear existing tabs
        self.clear_results()
        
        if not self.analysis_results:
            self.status_var.set("No duplicates found.")
            return
        
        total_duplicates = sum(len(duplicates) for duplicates in self.analysis_results.values())
        
        if total_duplicates == 0:
            self.status_var.set("Analysis complete. No duplicates found.")
            return
        
        # Create tabs for each time window
        for time_window in sorted(self.analysis_results.keys()):
            duplicates = self.analysis_results[time_window]
            
            if duplicates:  # Only create tab if there are results
                tab_frame = ttk.Frame(self.results_notebook)
                tab_name = f"Within {time_window}h ({len(duplicates)})"
                self.results_notebook.add(tab_frame, text=tab_name)
                
                # Create results table for this tab
                self.create_results_table(tab_frame, duplicates)
        
        # Update status
        self.status_var.set(f"Analysis complete. Found {total_duplicates} potential duplicate pairs.")
        self.update_ui_state()
    
    def create_results_table(self, parent, duplicates):
        """Create a results table for a specific time window."""
        # Create frame for table and scrollbars
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Define columns
        columns = ('Site', 'Ticket 1', 'Description 1', 'Created 1', 'Ticket 2', 
                  'Description 2', 'Created 2', 'Time Diff', 'Similarity')
        
        # Create treeview
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # Configure column headings and widths
        column_widths = {'Site': 150, 'Ticket 1': 100, 'Description 1': 200, 'Created 1': 140,
                        'Ticket 2': 100, 'Description 2': 200, 'Created 2': 140, 
                        'Time Diff': 80, 'Similarity': 80}
        
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(tree, c, False))
            tree.column(col, width=column_widths.get(col, 100), minwidth=50)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack scrollbars and tree
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        tree.pack(side='left', fill='both', expand=True)
        
        # Populate table with data
        for duplicate in duplicates:
            values = (
                duplicate['site'],
                duplicate['ticket1_number'],
                duplicate['ticket1_description'][:50] + "..." if len(duplicate['ticket1_description']) > 50 else duplicate['ticket1_description'],
                duplicate['ticket1_created'],
                duplicate['ticket2_number'],
                duplicate['ticket2_description'][:50] + "..." if len(duplicate['ticket2_description']) > 50 else duplicate['ticket2_description'],
                duplicate['ticket2_created'],
                duplicate['time_difference_formatted'],
                f"{duplicate['similarity_score']}%"
            )
            tree.insert('', 'end', values=values)
    
    def sort_treeview(self, tree, col, reverse):
        """Sort treeview by column."""
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        
        # Sort numerically if possible, otherwise alphabetically
        try:
            # Try to sort as numbers (for similarity percentages)
            if col == 'Similarity':
                data.sort(key=lambda x: int(x[0].replace('%', '')), reverse=reverse)
            else:
                data.sort(key=lambda x: x[0], reverse=reverse)
        except ValueError:
            data.sort(key=lambda x: x[0], reverse=reverse)
        
        for index, (val, child) in enumerate(data):
            tree.move(child, '', index)
        
        # Update column heading to show sort direction
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))
    
    def clear_results(self):
        """Clear all result tabs."""
        for tab in self.results_notebook.tabs():
            self.results_notebook.forget(tab)
    
    def export_results(self):
        """Export analysis results to CSV or Excel."""
        if not self.analysis_results or not self.duplicate_detector:
            messagebox.showwarning("No Results", "No results to export. Please run analysis first.")
            return
        
        # Get export file path
        file_path = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Export data
            df = self.duplicate_detector.export_results()
            success, message = self.export_manager.export_data(df, file_path)
            
            if success:
                self.status_var.set(f"Results exported to {os.path.basename(file_path)}")
                messagebox.showinfo("Export Successful", message)
            else:
                messagebox.showerror("Export Failed", message)
                
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred during export: {str(e)}")
    
    def update_ui_state(self):
        """Update UI state based on current conditions."""
        file_loaded = self.current_file_path is not None
        results_available = bool(self.analysis_results)
        
        # Enable/disable analyze button
        self.analyze_button.configure(state='normal' if file_loaded else 'disabled')
        
        # Enable/disable export button
        self.export_button.configure(state='normal' if results_available else 'disabled')

def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = DuplicateTicketApp(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        root.destroy()

if __name__ == "__main__":
    main()