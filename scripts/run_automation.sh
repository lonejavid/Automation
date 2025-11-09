#!/bin/bash

# KFC Guyana Drive-Thru Automation Launcher
# Makes it super easy to run the automation

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
SRC_DIR="$PROJECT_ROOT/src"
APPS_DIR="$PROJECT_ROOT/apps"

export PYTHONPATH="$SRC_DIR:$PYTHONPATH"
cd "$PROJECT_ROOT" || exit 1

echo "=================================="
echo "🍗 KFC GUYANA AUTOMATION"
echo "=================================="
echo ""
echo "What would you like to do?"
echo ""
echo "1. Full Automation (Download + Process)"
echo "2. Download from HMECloud only"
echo "3. Process existing files"
echo "4. Launch Web Interface"
echo "5. Exit"
echo ""
read -p "Enter your choice (1-5): " choice

case $choice in
    1)
        echo ""
        echo "🚀 Starting full automation..."
        PYTHONPATH="$SRC_DIR" python3 scripts/full_automation.py
        ;;
    2)
        echo ""
        echo "📥 Starting HMECloud download..."
        PYTHONPATH="$SRC_DIR" python3 -m automation.hmecloud
        ;;
    3)
        echo ""
        echo "📊 Processing existing files..."
        PYTHONPATH="$SRC_DIR" python3 scripts/full_automation.py --skip-download
        ;;
    4)
        echo ""
        echo "💻 Launching web interface..."
        echo "Opening browser to http://localhost:8501"
        PYTHONPATH="$SRC_DIR" streamlit run "$APPS_DIR/app_integrated.py"
        ;;
    5)
        echo ""
        echo "👋 Goodbye!"
        exit 0
        ;;
    *)
        echo ""
        echo "❌ Invalid choice. Exiting."
        exit 1
        ;;
esac

echo ""
echo "=================================="
echo "✅ Done!"
echo "=================================="

