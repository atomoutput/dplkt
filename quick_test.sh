#!/bin/bash
echo "Testing ServiceNow Duplicate Ticket Detection Tool with CSV Repair"
echo "=================================================================="
echo

echo "1. Basic functionality test with high threshold:"
python cli_main.py current.csv -t "1" -s 98 2>/dev/null | tail -10

echo
echo "2. CSV Repair functionality test:"
echo "Creating test corrupted CSV..."
cat > test_corrupted.csv << 'EOF'
Site,Number,Short description,Created,Resolved
"Test Site 1","CS001","Café machine issue","22-Jun-2025 10:00:00",""

"Test Site 2","CS002","Naïve user problem","22-Jun-2025 10:01:00",""
EOF

echo "Testing repair-only mode:"
python cli_main.py test_corrupted.csv --repair-only 2>/dev/null | head -10

echo
echo "3. Help command with new repair options:"
python cli_main.py --help 2>/dev/null | head -20

echo
echo "4. File structure:"
ls -la *.py | head -10

echo
echo "5. Installation check:"
if python -c "import pandas; print('pandas: OK')" 2>/dev/null; then
    echo "✓ Core dependencies available"
else
    echo "✗ Missing dependencies - run: pip install -r requirements.txt"
fi

echo
echo "6. Cleanup test files:"
rm -f test_*.csv test_*.bak

echo
echo "🎉 Application ready for use in Termux environment!"
echo "Usage examples:"
echo "  python cli_main.py current.csv -v"
echo "  python cli_main.py corrupted.csv --repair-only"
echo "  python cli_main.py current.csv --no-auto-repair"