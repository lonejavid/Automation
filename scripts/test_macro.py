#!/usr/bin/env python3
"""
Test script to run DT macro on latest downloaded file
"""

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

from automation.run_macro import process_downloaded_file

if __name__ == "__main__":
    print("\n" + "="*80)
    print("🧪 TESTING DT MACRO ON DOWNLOADED FILE")
    print("="*80)
    
    success = process_downloaded_file()
    
    if success:
        print("\n✅ Test completed successfully!")
    else:
        print("\n❌ Test failed!")
        sys.exit(1)

