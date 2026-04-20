#!/bin/bash
cd "$(dirname "$0")"
echo "====================================="
echo "  NARIX Customs Invoice Converter"
echo "====================================="
echo ""

# Install dependencies
echo "Installing dependencies..."
pip install flask pdfplumber openpyxl python-dateutil --break-system-packages -q

echo "Starting server..."
echo ""
echo "  ✓ Open in browser: http://localhost:5000"
echo ""
python app.py
