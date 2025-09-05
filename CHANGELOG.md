# Changelog

All notable changes to the ServiceNow Duplicate Ticket Detection Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-09-05

### Added
- Initial release of ServiceNow Duplicate Ticket Detection Tool
- Automatic CSV repair functionality for corrupted exports
- Intelligent duplicate detection using fuzzy string matching
- Time-window based analysis with configurable similarity thresholds
- Site-based grouping to prevent cross-site false positives
- Command-line interface optimized for Termux/Android environments
- GUI interface for desktop environments (Tkinter-based)
- Export results to CSV and Excel formats
- Comprehensive progress feedback and error handling
- Automatic encoding detection and conversion
- Robust CSV structure validation and cleanup
- Multiple time window analysis (1h, 8h, 24h, 72h configurable)
- Similarity threshold adjustment (50-100%)
- Option to exclude resolved tickets from analysis
- Backup creation during CSV repair operations
- Verbose output mode for detailed analysis results

### Technical Features
- Modular architecture with 6 core components
- Performance: 1,400-4,000 tickets/sec processing rate  
- Memory efficient: handles 2K+ tickets with low memory footprint
- Dual string similarity engines (fuzzywuzzy + difflib fallback)
- Automatic dependency fallbacks for constrained environments
- Cross-platform compatibility (Windows, macOS, Linux, Android/Termux)

### Documentation
- Comprehensive README with usage examples
- Quick test script for validation
- Requirements specification
- Installation and setup instructions