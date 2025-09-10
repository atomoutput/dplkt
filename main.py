#!/usr/bin/env python3
"""
Modern GUI for ServiceNow Duplicate Ticket Detection Tool
Enhanced user experience with intuitive design and better visual feedback
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from csv_parser import CSVParser
from duplicate_detector import DuplicateDetector
from export_manager import ExportManager

class DuplicateTicketApp:
    """Modern, intuitive GUI application for duplicate ticket detection."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("ServiceNow Duplicate Ticket Detection Tool")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 700)
        
        # Initialize components
        self.csv_parser = CSVParser(repair_callback=self.repair_progress_callback)
        self.duplicate_detector = None
        self.export_manager = ExportManager()
        
        # Data storage
        self.current_file_path = None
        self.analysis_results = {}
        
        # Create modern GUI
        self.setup_styles()
        self.create_widgets()
        self.create_layout()
        
        # Initial state
        self.update_ui_state()
    
    def setup_styles(self):
        """Configure modern ttk styles."""
        style = ttk.Style()
        
        # Use a modern theme if available
        available_themes = style.theme_names()
        if 'clam' in available_themes:
            style.theme_use('clam')
        elif 'alt' in available_themes:
            style.theme_use('alt')
        
        # Custom styles for modern look
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'))
        style.configure('Subtitle.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        style.configure('Info.TLabel', font=('Segoe UI', 10), foreground='#666666')
        style.configure('Success.TLabel', font=('Segoe UI', 10), foreground='#2e8b57')
        style.configure('Warning.TLabel', font=('Segoe UI', 10), foreground='#ff8c00')
        style.configure('Error.TLabel', font=('Segoe UI', 10), foreground='#dc143c')
        
        # Button styles
        style.configure('Primary.TButton', font=('Segoe UI', 11, 'bold'))
        style.configure('Secondary.TButton', font=('Segoe UI', 10))
        style.configure('Success.TButton', font=('Segoe UI', 10, 'bold'))
        
        # Frame styles
        style.configure('Card.TFrame', relief='raised', borderwidth=1)
        style.configure('Section.TLabelframe', font=('Segoe UI', 11, 'bold'), padding=(15, 10))
        style.configure('Modern.TLabelframe', font=('Segoe UI', 11, 'bold'), padding=(20, 15))
    
    def create_widgets(self):
        """Create all GUI widgets with modern design."""
        # Main container with padding
        self.main_frame = ttk.Frame(self.root, padding="20")
        
        # Header section
        self.create_header()
        
        # Main content area
        self.create_main_content()
        
        # Footer/Status section
        self.create_footer()
    
    def create_header(self):
        """Create the application header."""
        self.header_frame = ttk.Frame(self.main_frame)
        
        # Title and description
        self.title_label = ttk.Label(self.header_frame, 
                                    text="ServiceNow Duplicate Ticket Detection", 
                                    style='Title.TLabel')
        self.title_label.pack(anchor='w')
        
        self.subtitle_label = ttk.Label(self.header_frame,
                                       text="Intelligently detect and analyze duplicate tickets from CSV exports",
                                       style='Info.TLabel')
        self.subtitle_label.pack(anchor='w', pady=(5, 0))
        
        # Separator
        self.separator1 = ttk.Separator(self.header_frame, orient='horizontal')
        self.separator1.pack(fill='x', pady=(15, 0))
    
    def create_main_content(self):
        """Create the main content area with improved layout."""
        self.content_frame = ttk.Frame(self.main_frame)
        
        # Left panel - Input and Configuration
        self.left_panel = ttk.Frame(self.content_frame, width=400)
        self.left_panel.pack(side='left', fill='y', padx=(0, 20))
        self.left_panel.pack_propagate(False)  # Maintain fixed width
        
        # Right panel - Results and Analysis
        self.right_panel = ttk.Frame(self.content_frame)
        self.right_panel.pack(side='right', fill='both', expand=True)
        
        # Create left panel sections
        self.create_file_section()
        self.create_config_section()
        self.create_actions_section()
        
        # Create right panel sections
        self.create_results_section()
    
    def create_file_section(self):
        """Create the file input section."""
        self.file_section = ttk.LabelFrame(self.left_panel, text="📁 File Selection", 
                                          style='Modern.TLabelframe')
        self.file_section.pack(fill='x', pady=(0, 20))
        
        # File selection area
        self.file_select_frame = ttk.Frame(self.file_section)
        self.file_select_frame.pack(fill='x', pady=(0, 15))
        
        self.load_button = ttk.Button(self.file_select_frame, text="Select CSV File",
                                     command=self.load_file, style='Primary.TButton')
        self.load_button.pack(fill='x')
        
        # File info display
        self.file_info_frame = ttk.Frame(self.file_section)
        self.file_info_frame.pack(fill='x')
        
        self.file_path_var = tk.StringVar(value="No file selected")
        self.file_path_label = ttk.Label(self.file_info_frame, textvariable=self.file_path_var,
                                        style='Info.TLabel', wraplength=350)
        self.file_path_label.pack(anchor='w')
        
        # File stats (initially hidden)
        self.file_stats_frame = ttk.Frame(self.file_section)
        self.file_stats_var = tk.StringVar()
        self.file_stats_label = ttk.Label(self.file_stats_frame, textvariable=self.file_stats_var,
                                         style='Success.TLabel')
        self.file_stats_label.pack(anchor='w')
    
    def create_config_section(self):
        """Create the configuration section."""
        self.config_section = ttk.LabelFrame(self.left_panel, text="⚙️ Analysis Settings", 
                                            style='Modern.TLabelframe')
        self.config_section.pack(fill='x', pady=(0, 20))
        
        # Maximum timeframe with preset buttons
        ttk.Label(self.config_section, text="Maximum Timeframe (hours):", 
                 style='Subtitle.TLabel').pack(anchor='w', pady=(0, 8))
        
        # Preset buttons
        self.preset_frame = ttk.Frame(self.config_section)
        self.preset_frame.pack(fill='x', pady=(0, 10))
        
        presets = [
            ("Quick", "1"),
            ("Standard", "24"),
            ("Comprehensive", "72"),
            ("Extended", "168")
        ]
        
        for i, (name, value) in enumerate(presets):
            btn = ttk.Button(self.preset_frame, text=name,
                           command=lambda v=value: self.set_max_timeframe(v),
                           style='Secondary.TButton')
            btn.pack(side='left', padx=(0, 5) if i < len(presets)-1 else 0)
        
        # Custom timeframe entry with spinner
        self.timeframe_frame = ttk.Frame(self.config_section)
        self.timeframe_frame.pack(fill='x', pady=(0, 15))
        
        self.max_timeframe_var = tk.IntVar(value=72)
        self.max_timeframe_spinbox = ttk.Spinbox(self.timeframe_frame, 
                                                from_=1, to=8760,  # 1 hour to 1 year
                                                textvariable=self.max_timeframe_var,
                                                width=10, font=('Segoe UI', 10))
        self.max_timeframe_spinbox.pack(side='left')
        
        ttk.Label(self.timeframe_frame, text="hours (finds all duplicates within this timeframe)",
                 style='Info.TLabel').pack(side='left', padx=(10, 0))
        
        # Similarity threshold with visual feedback
        ttk.Label(self.config_section, text="Similarity Threshold:", 
                 style='Subtitle.TLabel').pack(anchor='w', pady=(0, 8))
        
        self.similarity_frame = ttk.Frame(self.config_section)
        self.similarity_frame.pack(fill='x', pady=(0, 15))
        
        self.similarity_var = tk.IntVar(value=85)
        self.similarity_scale = ttk.Scale(self.similarity_frame, from_=50, to=100,
                                         variable=self.similarity_var, orient='horizontal')
        self.similarity_scale.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        self.similarity_label_var = tk.StringVar(value="85%")
        self.similarity_label = ttk.Label(self.similarity_frame, textvariable=self.similarity_label_var,
                                         style='Header.TLabel', width=8)
        self.similarity_label.pack(side='right')
        self.similarity_scale.configure(command=self.update_similarity_label)
        
        # Similarity guidance
        self.similarity_guide_var = tk.StringVar(value="Balanced - Good for most cases")
        self.similarity_guide = ttk.Label(self.config_section, textvariable=self.similarity_guide_var,
                                         style='Info.TLabel')
        self.similarity_guide.pack(anchor='w', pady=(0, 15))
        
        # Options with better descriptions
        ttk.Label(self.config_section, text="Options:", 
                 style='Subtitle.TLabel').pack(anchor='w', pady=(0, 8))
        
        self.exclude_resolved_var = tk.BooleanVar(value=False)
        self.exclude_resolved_check = ttk.Checkbutton(self.config_section, 
                                                     text="Exclude resolved tickets",
                                                     variable=self.exclude_resolved_var)
        self.exclude_resolved_check.pack(anchor='w', pady=(0, 5))
        
        self.auto_repair_var = tk.BooleanVar(value=True)
        self.auto_repair_check = ttk.Checkbutton(self.config_section,
                                                text="Auto-repair corrupted CSV files",
                                                variable=self.auto_repair_var)
        self.auto_repair_check.pack(anchor='w', pady=(0, 15))
        
        # Enhanced Analysis Options
        ttk.Label(self.config_section, text="📊 Enhanced Analysis (Excel Only):", 
                 style='Subtitle.TLabel').pack(anchor='w', pady=(0, 8))
        
        # Enhanced analysis checkboxes
        self.enable_same_day_var = tk.BooleanVar(value=False)
        self.enable_same_day_check = ttk.Checkbutton(self.config_section,
                                                    text="Same-day duplicates (ignore description similarity)",
                                                    variable=self.enable_same_day_var)
        self.enable_same_day_check.pack(anchor='w', pady=(0, 3))
        
        self.enable_rapid_fire_var = tk.BooleanVar(value=False)
        self.enable_rapid_fire_check = ttk.Checkbutton(self.config_section,
                                                      text="Rapid-fire duplicates (15-60 minute windows)",
                                                      variable=self.enable_rapid_fire_var)
        self.enable_rapid_fire_check.pack(anchor='w', pady=(0, 3))
        
        self.enable_exact_match_var = tk.BooleanVar(value=False)
        self.enable_exact_match_check = ttk.Checkbutton(self.config_section,
                                                       text="Exact content matches (100% identical descriptions)",
                                                       variable=self.enable_exact_match_var)
        self.enable_exact_match_check.pack(anchor='w', pady=(0, 3))
        
        self.enable_category_patterns_var = tk.BooleanVar(value=False)
        self.enable_category_patterns_check = ttk.Checkbutton(self.config_section,
                                                             text="Category patterns (same category/subcategory per day)",
                                                             variable=self.enable_category_patterns_var)
        self.enable_category_patterns_check.pack(anchor='w')
    
    def create_actions_section(self):
        """Create the actions section."""
        self.actions_section = ttk.LabelFrame(self.left_panel, text="🚀 Actions", 
                                             style='Modern.TLabelframe')
        self.actions_section.pack(fill='x')
        
        # Main analyze button
        self.analyze_button = ttk.Button(self.actions_section, text="Run Analysis",
                                        command=self.run_analysis_threaded,
                                        style='Primary.TButton')
        self.analyze_button.pack(fill='x', pady=(0, 10))
        
        # Secondary actions
        self.secondary_frame = ttk.Frame(self.actions_section)
        self.secondary_frame.pack(fill='x')
        
        self.export_button = ttk.Button(self.secondary_frame, text="Export Results",
                                       command=self.export_results, style='Success.TButton')
        self.export_button.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        self.clear_button = ttk.Button(self.secondary_frame, text="Clear",
                                      command=self.clear_results, style='Secondary.TButton')
        self.clear_button.pack(side='right', fill='x', expand=True, padx=(5, 0))
    
    def create_results_section(self):
        """Create the results display section."""
        self.results_section = ttk.LabelFrame(self.right_panel, text="📊 Analysis Results", 
                                             style='Modern.TLabelframe')
        self.results_section.pack(fill='both', expand=True)
        
        # Results header with summary
        self.results_header = ttk.Frame(self.results_section)
        self.results_header.pack(fill='x', pady=(0, 15))
        
        self.results_summary_var = tk.StringVar(value="Run analysis to see results")
        self.results_summary = ttk.Label(self.results_header, textvariable=self.results_summary_var,
                                        style='Info.TLabel')
        self.results_summary.pack(anchor='w')
        
        # Tabbed results display
        self.results_notebook = ttk.Notebook(self.results_section)
        self.results_notebook.pack(fill='both', expand=True)
        
        # Welcome tab
        self.create_welcome_tab()
    
    def create_welcome_tab(self):
        """Create a welcome tab with instructions."""
        self.welcome_frame = ttk.Frame(self.results_notebook)
        self.results_notebook.add(self.welcome_frame, text="Welcome")
        
        # Center the welcome content
        self.welcome_content = ttk.Frame(self.welcome_frame)
        self.welcome_content.pack(expand=True)
        
        welcome_text = """
🎯 Welcome to the ServiceNow Duplicate Detection Tool

📋 Quick Start:
1. Click 'Select CSV File' to load your ServiceNow export
2. Set maximum timeframe (how far back to search for duplicates)
3. Adjust similarity threshold (85% recommended)
4. Click 'Run Analysis' to detect duplicates
5. Review results with time categories for easy grouping
6. Export findings when ready

💡 Tips:
• Quick (1h): Catches immediate duplicate submissions
• Standard (24h): Finds duplicates within a business day
• Comprehensive (72h): Includes weekend duplicates
• Extended (168h): Full week analysis
• Each duplicate pair appears only once (no time window overlap)

🛠️ Features:
• Maximum timeframe approach eliminates duplicate reporting
• Automatic time categorization (0-1h, 1-4h, 4-8h, etc.)
• Site-based grouping prevents cross-site matches
• Enhanced analysis options for Excel exports
• Automatic CSV repair for corrupted files
        """
        
        ttk.Label(self.welcome_content, text=welcome_text.strip(),
                 style='Info.TLabel', justify='left').pack(pady=50)
    
    def create_footer(self):
        """Create the footer with status and progress."""
        self.footer_frame = ttk.Frame(self.main_frame)
        
        # Separator
        self.separator2 = ttk.Separator(self.footer_frame, orient='horizontal')
        self.separator2.pack(fill='x', pady=(15, 10))
        
        # Status and progress area
        self.status_frame = ttk.Frame(self.footer_frame)
        self.status_frame.pack(fill='x')
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to analyze CSV files")
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_var,
                                     style='Info.TLabel')
        self.status_label.pack(side='left')
        
        # Progress bar (initially hidden)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.status_frame, variable=self.progress_var,
                                           mode='determinate', length=300)
    
    def create_layout(self):
        """Arrange widgets in the layout."""
        self.main_frame.pack(fill='both', expand=True)
        self.header_frame.pack(fill='x', pady=(0, 20))
        self.content_frame.pack(fill='both', expand=True)
        self.footer_frame.pack(fill='x', pady=(20, 0))
    
    def repair_progress_callback(self, message: str):
        """Handle repair progress updates."""
        self.status_var.set(f"Repair: {message}")
        self.root.update_idletasks()
    
    def set_max_timeframe(self, value):
        """Set maximum timeframe from preset."""
        self.max_timeframe_var.set(int(value))
    
    def update_similarity_label(self, value):
        """Update similarity threshold label with guidance."""
        threshold = int(float(value))
        self.similarity_label_var.set(f"{threshold}%")
        
        # Update guidance
        if threshold >= 95:
            self.similarity_guide_var.set("Very Strict - Only exact matches")
        elif threshold >= 90:
            self.similarity_guide_var.set("Strict - High precision, fewer matches")
        elif threshold >= 80:
            self.similarity_guide_var.set("Balanced - Good for most cases")
        elif threshold >= 70:
            self.similarity_guide_var.set("Relaxed - More matches, some false positives")
        else:
            self.similarity_guide_var.set("Very Relaxed - Many matches, review carefully")
    
    def load_file(self):
        """Load and validate CSV file with enhanced feedback."""
        file_path = filedialog.askopenfilename(
            title="Select ServiceNow CSV Export",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        # Show loading feedback
        self.status_var.set("Loading and validating CSV file...")
        self.load_button.configure(text="Loading...", state='disabled')
        self.root.update()
        
        # Load file
        success, message = self.csv_parser.load_and_validate(
            file_path, auto_repair=self.auto_repair_var.get()
        )
        
        if success:
            self.current_file_path = file_path
            filename = os.path.basename(file_path)
            self.file_path_var.set(f"✓ {filename}")
            
            # Show file statistics
            summary = self.csv_parser.get_data_summary()
            stats_text = f"📊 {summary['total_tickets']} tickets from {summary['unique_sites']} sites"
            if summary.get('was_repaired', False):
                stats_text += " (auto-repaired)"
            self.file_stats_var.set(stats_text)
            self.file_stats_frame.pack(fill='x', pady=(10, 0))
            
            # Show data info
            date_range = (summary['date_range']['earliest'].strftime('%Y-%m-%d') + 
                         " to " + summary['date_range']['latest'].strftime('%Y-%m-%d'))
            resolved_info = f", {summary['resolved_tickets']} resolved" if summary.get('resolved_tickets', 0) > 0 else ""
            
            self.status_var.set(f"Ready to analyze • {date_range}{resolved_info}")
            
            # Clear previous results
            self.clear_results()
            
        else:
            self.current_file_path = None
            self.file_path_var.set("❌ Failed to load file")
            self.file_stats_frame.pack_forget()
            self.status_var.set(f"Error: {message}")
        
        # Restore button
        self.load_button.configure(text="Select CSV File", state='normal')
        self.update_ui_state()
    
    def run_analysis_threaded(self):
        """Run analysis in a separate thread."""
        threading.Thread(target=self.run_analysis, daemon=True).start()
    
    def run_analysis(self):
        """Run the duplicate detection analysis."""
        try:
            # Parse configuration
            max_timeframe = self.max_timeframe_var.get()
            similarity_threshold = self.similarity_var.get()
            exclude_resolved = self.exclude_resolved_var.get()
            
            # Update UI
            self.root.after(0, lambda: self.set_analysis_running(True))
            
            # Get data
            data = self.csv_parser.get_filtered_data(exclude_resolved)
            
            if data.empty:
                self.root.after(0, lambda: messagebox.showwarning(
                    "No Data", "No tickets to analyze after applying filters."))
                return
            
            # Create detector and run enhanced analysis
            self.duplicate_detector = DuplicateDetector(self.progress_callback)
            
            # Check if any enhanced analysis is enabled
            enable_enhanced = any([
                self.enable_same_day_var.get(),
                self.enable_rapid_fire_var.get(),
                self.enable_exact_match_var.get(),
                self.enable_category_patterns_var.get()
            ])
            
            if enable_enhanced:
                self.analysis_results = self.duplicate_detector.analyze_enhanced(
                    data, max_timeframe, similarity_threshold,
                    enable_same_day=self.enable_same_day_var.get(),
                    enable_rapid_fire=self.enable_rapid_fire_var.get(),
                    enable_exact_match=self.enable_exact_match_var.get(),
                    enable_category_patterns=self.enable_category_patterns_var.get()
                )
                # Store results in new format for display
                self.fuzzy_results = self.analysis_results['fuzzy_matching']
            else:
                # Use new single timeframe analysis
                self.fuzzy_results = self.duplicate_detector.analyze(data, max_timeframe, similarity_threshold)
                self.analysis_results = {'fuzzy_matching': self.fuzzy_results}
            
            # Update UI with results
            self.root.after(0, self.display_results)
            
        except ValueError as e:
            self.root.after(0, lambda: messagebox.showerror("Configuration Error", 
                f"Please check your settings:\n{str(e)}"))
        except Exception as e:
            error_msg = f"Analysis failed with an unexpected error:\n{str(e)}\n\nPlease check your CSV file format and try again."
            self.root.after(0, lambda: messagebox.showerror("Analysis Error", error_msg))
        finally:
            self.root.after(0, lambda: self.set_analysis_running(False))
    
    def progress_callback(self, message: str, current: int, total: int):
        """Handle analysis progress updates."""
        progress_percent = (current / total) * 100 if total > 0 else 0
        
        self.root.after(0, lambda: [
            self.status_var.set(f"Analyzing: {message}"),
            self.progress_var.set(progress_percent)
        ])
    
    def set_analysis_running(self, running: bool):
        """Update UI state during analysis."""
        if running:
            self.analyze_button.configure(state='disabled', text="Analyzing...")
            self.progress_bar.pack(side='right', padx=(20, 0))
        else:
            self.analyze_button.configure(state='normal', text="Run Analysis")
            self.progress_bar.pack_forget()
        
        self.update_ui_state()
    
    def display_results(self):
        """Display analysis results with enhanced visualization."""
        # Clear existing result tabs (keep welcome)
        for tab_id in self.results_notebook.tabs()[1:]:  # Skip welcome tab
            self.results_notebook.forget(tab_id)
        
        if not self.analysis_results:
            self.results_summary_var.set("No results to display")
            return
        
        # Handle both enhanced and new single timeframe formats
        if hasattr(self, 'fuzzy_results'):
            # New format: list of duplicate pairs
            total_duplicates = len(self.fuzzy_results) if self.fuzzy_results else 0
        elif 'fuzzy_matching' in self.analysis_results:
            # Enhanced results format
            fuzzy_results = self.analysis_results['fuzzy_matching']
            if isinstance(fuzzy_results, list):
                total_duplicates = len(fuzzy_results)
            else:
                # Legacy multi-window format
                total_duplicates = sum(len(pairs) for pairs in fuzzy_results.values()) if fuzzy_results else 0
        else:
            # Legacy format (backward compatibility)
            total_duplicates = sum(len(pairs) for pairs in self.analysis_results.values())
        
        if total_duplicates == 0:
            self.results_summary_var.set("✅ No duplicate tickets found")
            self.status_var.set("Analysis complete - No duplicates detected")
            
            # Add a "no results" tab
            no_results_frame = ttk.Frame(self.results_notebook)
            self.results_notebook.add(no_results_frame, text="No Duplicates")
            
            ttk.Label(no_results_frame, 
                     text="🎉 Great! No duplicate tickets were found.\n\nThis suggests good ticket management practices.",
                     style='Success.TLabel', justify='center').pack(expand=True)
            return
        
        # Update summary
        affected_sites = len(set(dup['site'] for pairs in self.analysis_results.values() for dup in pairs))
        self.results_summary_var.set(f"🔍 Found {total_duplicates} potential duplicates across {affected_sites} sites")
        
        # Create result tabs
        for time_window in sorted(self.analysis_results.keys()):
            duplicates = self.analysis_results[time_window]
            
            if duplicates:
                tab_frame = ttk.Frame(self.results_notebook)
                tab_name = f"Within {time_window}h ({len(duplicates)})"
                self.results_notebook.add(tab_frame, text=tab_name)
                
                self.create_enhanced_results_table(tab_frame, duplicates, time_window)
        
        # Switch to first results tab
        if len(self.results_notebook.tabs()) > 1:
            self.results_notebook.select(1)
        
        self.status_var.set(f"Analysis complete - {total_duplicates} potential duplicates found")
        self.update_ui_state()
    
    def create_enhanced_results_table(self, parent, duplicates, time_window):
        """Create an enhanced results table with better formatting."""
        # Create main container
        container = ttk.Frame(parent, padding="10")
        container.pack(fill='both', expand=True)
        
        # Summary header for this time window
        stats = self.duplicate_detector.get_summary_stats()[time_window]
        summary_text = (f"📋 {stats['total_pairs']} duplicate pairs • "
                       f"🏢 {stats['affected_sites']} sites affected • "
                       f"📊 {stats['avg_similarity']:.1f}% average similarity")
        
        ttk.Label(container, text=summary_text, style='Info.TLabel').pack(anchor='w', pady=(0, 10))
        
        # Create table frame
        table_frame = ttk.Frame(container)
        table_frame.pack(fill='both', expand=True)
        
        # Define columns with better headers
        columns = ('Site', 'First Ticket', 'Description 1', 'Second Ticket', 
                  'Description 2', 'Time Gap', 'Similarity')
        
        # Create treeview
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # Configure columns with better widths
        column_configs = {
            'Site': 200,
            'First Ticket': 100,
            'Description 1': 250,
            'Second Ticket': 100,
            'Description 2': 250,
            'Time Gap': 80,
            'Similarity': 80
        }
        
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(tree, c, False))
            tree.column(col, width=column_configs[col], minwidth=60)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack components
        tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Populate with data
        for duplicate in duplicates:
            # Format descriptions with truncation
            desc1 = self.truncate_text(duplicate['ticket1_description'], 40)
            desc2 = self.truncate_text(duplicate['ticket2_description'], 40)
            
            values = (
                self.truncate_text(duplicate['site'], 25),
                duplicate['ticket1_number'],
                desc1,
                duplicate['ticket2_number'],
                desc2,
                duplicate['time_difference_formatted'],
                f"{duplicate['similarity_score']}%"
            )
            
            # Add color coding based on similarity
            item = tree.insert('', 'end', values=values)
            if duplicate['similarity_score'] >= 95:
                tree.set(item, 'Similarity', f"🔴 {duplicate['similarity_score']}%")
            elif duplicate['similarity_score'] >= 90:
                tree.set(item, 'Similarity', f"🟡 {duplicate['similarity_score']}%")
            else:
                tree.set(item, 'Similarity', f"🟢 {duplicate['similarity_score']}%")
    
    def truncate_text(self, text, max_length):
        """Truncate text with ellipsis if too long."""
        if len(text) <= max_length:
            return text
        return text[:max_length-3] + "..."
    
    def sort_treeview(self, tree, col, reverse):
        """Sort treeview by column."""
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        
        # Clean similarity column for sorting
        if col == 'Similarity':
            data = [(item[0].split()[-1].replace('%', ''), item[1]) for item in data]
            data.sort(key=lambda x: int(x[0]), reverse=reverse)
        else:
            data.sort(key=lambda x: x[0], reverse=reverse)
        
        for index, (val, child) in enumerate(data):
            tree.move(child, '', index)
        
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))
    
    def export_results(self):
        """Export results with user feedback."""
        if not self.analysis_results or not self.duplicate_detector:
            messagebox.showwarning("No Results", "No results to export. Please run analysis first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Analysis Results",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            self.status_var.set("Exporting results...")
            self.root.update_idletasks()
            
            # Check if we have enhanced results and if Excel export is chosen
            file_extension = os.path.splitext(file_path)[1].lower()
            enable_enhanced = any([
                self.enable_same_day_var.get(),
                self.enable_rapid_fire_var.get(),
                self.enable_exact_match_var.get(),
                self.enable_category_patterns_var.get()
            ])
            
            if enable_enhanced and file_extension in ['.xlsx', '.xls']:
                # Use enhanced export with multiple sheets
                success, message = self.export_manager.export_enhanced_data(
                    self.analysis_results, file_path,
                    enable_same_day=self.enable_same_day_var.get(),
                    enable_rapid_fire=self.enable_rapid_fire_var.get(),
                    enable_exact_match=self.enable_exact_match_var.get(),
                    enable_category_patterns=self.enable_category_patterns_var.get()
                )
            else:
                # Use standard export (CSV or single-sheet Excel)
                if hasattr(self, 'fuzzy_results') and self.fuzzy_results:
                    # New format: use the dedicated export method
                    df = self.duplicate_detector.export_results_new(self.fuzzy_results)
                elif 'fuzzy_matching' in self.analysis_results:
                    # Enhanced format: convert back to standard DataFrame
                    df = self.export_manager._convert_fuzzy_results_to_dataframe(
                        self.analysis_results['fuzzy_matching']
                    )
                else:
                    # Legacy format
                    df = self.duplicate_detector.export_results()
                
                success, message = self.export_manager.export_data(df, file_path)
            
            if success:
                self.status_var.set(f"✅ Results exported to {os.path.basename(file_path)}")
                messagebox.showinfo("Export Successful", f"Results saved to:\n{file_path}")
            else:
                self.status_var.set("❌ Export failed")
                messagebox.showerror("Export Failed", message)
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {str(e)}")
    
    def clear_results(self):
        """Clear all results."""
        # Clear result tabs (keep welcome)
        for tab_id in self.results_notebook.tabs()[1:]:
            self.results_notebook.forget(tab_id)
        
        # Switch back to welcome
        self.results_notebook.select(0)
        
        self.analysis_results = {}
        self.results_summary_var.set("Run analysis to see results")
        self.status_var.set("Results cleared")
        self.update_ui_state()
    
    def update_ui_state(self):
        """Update UI state based on current conditions."""
        file_loaded = self.current_file_path is not None
        results_available = bool(self.analysis_results)
        
        # Update button states
        self.analyze_button.configure(state='normal' if file_loaded else 'disabled')
        self.export_button.configure(state='normal' if results_available else 'disabled')
        self.clear_button.configure(state='normal' if results_available else 'disabled')

def main():
    """Main entry point for the GUI."""
    root = tk.Tk()
    app = DuplicateTicketApp(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        root.destroy()

if __name__ == "__main__":
    main()