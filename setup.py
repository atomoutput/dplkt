#!/usr/bin/env python3
"""
Setup script for ServiceNow Duplicate Ticket Detection Tool
Handles installation, dependency management, and environment setup
"""

import os
import sys
import subprocess
from pathlib import Path

def run_command(command, description):
    """Run a shell command with error handling."""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"✅ {description} completed successfully")
            return True
        else:
            print(f"❌ {description} failed: {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ Error during {description}: {e}")
        return False

def check_python_version():
    """Check if Python version meets requirements."""
    version = sys.version_info
    if version.major == 3 and version.minor >= 7:
        print(f"✅ Python {version.major}.{version.minor}.{version.micro} meets requirements")
        return True
    else:
        print(f"❌ Python {version.major}.{version.minor}.{version.micro} is too old. Requires Python 3.7+")
        return False

def install_dependencies():
    """Install required Python packages."""
    requirements_file = Path(__file__).parent / "requirements.txt"
    
    if not requirements_file.exists():
        print("❌ requirements.txt not found")
        return False
    
    command = f"{sys.executable} -m pip install -r {requirements_file}"
    return run_command(command, "Installing Python dependencies")

def check_optional_dependencies():
    """Check for optional dependencies and suggest installation."""
    optional_deps = [
        ("chardet", "Better encoding detection for CSV repair"),
        ("fuzzywuzzy", "Enhanced string similarity matching"),
        ("python-Levenshtein", "Faster fuzzy string operations"),
        ("openpyxl", "Excel export functionality"),
    ]
    
    print("\n📋 Checking optional dependencies:")
    
    for package, description in optional_deps:
        try:
            __import__(package)
            print(f"✅ {package} - {description}")
        except ImportError:
            print(f"⚠️  {package} - {description} (optional, install for better performance)")

def create_sample_config():
    """Create a sample configuration file."""
    config_content = '''# ServiceNow Duplicate Ticket Detection Configuration
# Copy this file to config.py and modify as needed

# Default analysis parameters
DEFAULT_TIME_WINDOWS = [1, 8, 24, 72]  # Hours
DEFAULT_SIMILARITY_THRESHOLD = 85      # Percentage
DEFAULT_EXCLUDE_RESOLVED = False       # Boolean

# CSV repair settings
REPAIR_CREATE_BACKUP = True           # Create .bak files
REPAIR_TARGET_ENCODING = "utf-8"      # Target encoding for repairs

# Export settings
EXPORT_FORMAT = "csv"                 # Default: csv or xlsx
EXPORT_INCLUDE_SUMMARY = True         # Include summary sheet in Excel exports

# Performance tuning
MAX_MEMORY_USAGE_MB = 512             # Memory limit for large files
ENABLE_MULTITHREADING = True          # Use multiple threads for processing
'''
    
    config_path = Path(__file__).parent / "config_sample.py"
    with open(config_path, 'w') as f:
        f.write(config_content)
    
    print(f"✅ Created sample configuration: {config_path}")

def run_tests():
    """Run basic functionality tests."""
    print("\n🧪 Running basic tests...")
    
    test_script = Path(__file__).parent / "quick_test.sh"
    if test_script.exists():
        return run_command("bash quick_test.sh", "Running test suite")
    else:
        print("⚠️  Test script not found, skipping tests")
        return True

def main():
    """Main setup routine."""
    print("🚀 ServiceNow Duplicate Ticket Detection Tool - Setup")
    print("=" * 55)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install dependencies
    if not install_dependencies():
        print("❌ Failed to install dependencies. Please check your environment.")
        sys.exit(1)
    
    # Check optional dependencies
    check_optional_dependencies()
    
    # Create sample configuration
    create_sample_config()
    
    # Run tests
    if not run_tests():
        print("⚠️  Some tests failed, but installation can continue")
    
    print("\n🎉 Setup completed successfully!")
    print("\nNext steps:")
    print("1. Run: python cli_main.py --help")
    print("2. Test with: python cli_main.py your_file.csv")
    print("3. For GUI: python main.py")
    print("\nFor more information, see README.md")

if __name__ == "__main__":
    main()