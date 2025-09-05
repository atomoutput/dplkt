# ServiceNow Duplicate Ticket Detection Tool

A desktop application for automatically detecting potential duplicate tickets from ServiceNow CSV exports.

## Features

- **Automatic CSV Repair**: Fixes corrupted exports, encoding issues, and malformed data
- Load ServiceNow CSV exports with automatic validation
- Configurable similarity detection using fuzzy string matching
- Time-window based duplicate detection
- Site-based grouping to prevent cross-site comparisons
- Interactive tabbed results display (GUI mode)
- Command-line interface optimized for Termux/Android
- Export results to CSV/Excel formats

## Requirements

- Python 3.7+
- Tkinter (usually included with Python)

## Installation

1. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the application:
   ```bash
   # GUI mode (desktop/laptop)
   python main.py
   
   # CLI mode (recommended for Termux/Android)
   python cli_main.py current.csv
   ```

## Usage

### Command Line Interface (Recommended)

```bash
# Basic duplicate detection
python cli_main.py current.csv

# Advanced options
python cli_main.py current.csv -t "1,8,24" -s 90 -o results.csv --verbose

# Repair corrupted CSV files
python cli_main.py corrupted.csv --repair-only --create-backup

# Disable auto-repair if needed
python cli_main.py current.csv --no-auto-repair
```

### GUI Mode (Desktop)

1. Click "Load CSV File" and select your ServiceNow export
2. Configure analysis parameters:
   - Time windows (comma-separated hours, e.g., "1, 8, 24, 72")
   - Similarity threshold (50-100%)
   - Option to exclude resolved tickets
3. Click "Run Analysis" to detect duplicates
4. View results in tabbed interface, sorted by time windows
5. Export results using the "Export Results" button

## CSV Repair Features

The tool automatically detects and repairs common CSV corruption issues:

- **Encoding Problems**: Automatic detection and conversion (UTF-8, Windows-1252, ISO-8859-1)
- **Malformed Structure**: Fixes broken rows and inconsistent formatting
- **Empty Rows**: Removes completely empty rows
- **Duplicate Entries**: Removes exact duplicate rows
- **Character Issues**: Handles special characters and encoding artifacts

### Repair Options

- `--repair-only`: Only repair the file without running analysis
- `--no-auto-repair`: Disable automatic repair during normal analysis
- `--create-backup`: Create backup before repairing (default: enabled)
- `--overwrite-original`: Replace original file with repaired version
- `--encoding`: Specify target encoding (default: utf-8)

## CSV Format Requirements

The CSV file must contain these columns:
- Site
- Number
- Short description
- Created (format: DD-Mon-YYYY HH:MM:SS)
- Resolved (optional)