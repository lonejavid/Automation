#!/bin/bash

# KFC Guyana Drive-Thru Automation Launcher
# Makes it super easy to run the automation

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

cd /Users/sanctum/Desktop/Automation

case $choice in
    1)
        echo ""
        echo "🚀 Starting full automation..."
        python3 full_automation.py
        ;;
    2)
        echo ""
        echo "📥 Starting HMECloud download..."
        python3 hmecloud_automation.py
        ;;
    3)
        echo ""
        echo "📊 Processing existing files..."
        python3 complete_automation.py
        ;;
    4)
        echo ""
        echo "💻 Launching web interface..."
        echo "Opening browser to http://localhost:8501"
        streamlit run app_integrated.py
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

