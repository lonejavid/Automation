# 🚀 Complete HMECloud Automation Guide

## 🎯 Overview

This automation now handles **EVERYTHING** from login to final report:

1. ✅ Auto-login to HMECloud
2. ✅ Auto-download reports
3. ✅ Data transformation
4. ✅ Template updates
5. ✅ Pivot table refresh
6. ✅ Date updates

**Total time: ~5 minutes** (fully automated!)

---

## 📦 Installation

### First Time Setup:

```bash
cd /Users/sanctum/Desktop/Automation

# Install required packages
pip3 install -r requirements.txt

# Install ChromeDriver (for Selenium)
brew install chromedriver

# Or manually download from: https://chromedriver.chromium.org/
```

---

## 🚀 Usage Options

### Option 1: Full Automation (Recommended) ⭐

**One command does everything:**

```bash
python3 full_automation.py
```

This will:
- Login to HMECloud automatically
- Download all 6 stores
- Process all data
- Update template
- Done!

**Interactive mode** - lets you choose options:
- All stores vs single store
- Yesterday vs custom date

---

### Option 2: HMECloud Download Only

**Download reports from HMECloud:**

```bash
python3 hmecloud_automation.py
```

Features:
- Auto-login with saved credentials
- Interactive store selection
- Downloads to `downloads/` folder
- Browser window opens (you can see what's happening!)

---

### Option 3: Process Existing Files

**If you already have files downloaded:**

```bash
python3 complete_automation.py
```

Or:

```bash
python3 full_automation.py --skip-download
```

---

### Option 4: Web Interface (Streamlit)

**Beautiful web UI:**

```bash
streamlit run app_integrated.py
```

Then open browser to: http://localhost:8501

Features:
- Professional UI
- Real-time progress
- Store selection
- Date picker
- Full automation from web interface

---

## 🔐 Credentials

**Auto-configured:**
- Email: `doudit@hotmail.com`
- Password: `Kfcguy123!@#`

Stored in: `hmecloud_automation.py` (lines 18-19)

To change credentials, edit these lines in the script.

---

## 🏪 Stores

The automation knows about all 6 KFC stores:

1. (Ungrouped) 2 Vlissengen Road – KFC
2. (Ungrouped) 5 Mandela - KFC
3. (Ungrouped) Movie Towne - KFC
4. (Ungrouped) Giftland Mall - KFC
5. (Ungrouped) Sheriff Street - KFC
6. (Ungrouped) Providence - KFC

---

## 📁 File Structure

```
Automation/
├── full_automation.py          ⭐ COMPLETE automation
├── hmecloud_automation.py      🌐 HMECloud download
├── complete_automation.py      📊 Data processing
├── app_integrated.py           💻 Web interface
│
├── downloads/                  📥 Downloaded files
├── templates/                  📊 Excel template
│
├── transform_data.py           (module)
├── template_operations.py      (module)
└── requirements.txt
```

---

## ⚙️ How It Works

### Full Automation Flow:

```
1. Open Chrome browser
   ↓
2. Navigate to hmecloud.com
   ↓
3. Auto-login with credentials
   ↓
4. Navigate to Reports → Raw Car Data
   ↓
5. For each store:
   - Select store
   - Set date
   - View report
   - Export to Excel
   - Download
   ↓
6. Close browser
   ↓
7. Transform data (all files)
   ↓
8. Paste into template
   ↓
9. Copy formulas
   ↓
10. Refresh pivot tables
   ↓
11. Update dates
   ↓
12. Save & done! ✅
```

---

## 🎯 Daily Workflow

### Every Morning:

```bash
# One command:
python3 full_automation.py

# Press Enter (or select options)
# Wait 5 minutes
# Done!
```

**That's it!** Open the Excel template and click "Refresh All".

---

## ⏱️ Time Comparison

| Method | Time |
|--------|------|
| **Manual (old way)** | 40-50 minutes |
| **Semi-automated** | 17 minutes |
| **Full automation** | **5 minutes** |

**Time saved: 90%!** ⚡

---

## 🔧 Advanced Options

### Command Line Options:

```bash
# Interactive mode (default)
python3 full_automation.py

# Skip download, use existing files
python3 full_automation.py --skip-download

# Show help
python3 full_automation.py --help
```

### Python API:

```python
from hmecloud_automation import download_all_stores, download_single_store
from datetime import datetime, timedelta

# Download all stores for yesterday
download_all_stores()

# Download specific store
download_single_store("(Ungrouped) 5 Mandela - KFC")

# Custom date
yesterday = datetime.now() - timedelta(days=1)
download_all_stores(report_date=yesterday)
```

---

## 🐛 Troubleshooting

### "ChromeDriver not found"

Install ChromeDriver:
```bash
brew install chromedriver
```

Or download from: https://chromedriver.chromium.org/

### "Browser won't open"

Make sure Chrome is installed and ChromeDriver matches your Chrome version.

### "Login failed"

Check if HMECloud credentials have changed. Update in `hmecloud_automation.py`.

### "Download failed"

- Check internet connection
- Verify HMECloud is accessible
- Make sure `downloads/` folder exists

### "Processing failed"

- Make sure files downloaded successfully
- Check template path in `complete_automation.py`
- Verify Excel template exists

---

## 📝 Configuration

### Change Download Folder:

Edit `hmecloud_automation.py`:
```python
DOWNLOADS_FOLDER = "/your/path/here"
```

### Change Template Path:

Edit `complete_automation.py`:
```python
TEMPLATE_PATH = "/your/path/here/template.xlsx"
```

### Headless Mode (No Browser Window):

Edit `hmecloud_automation.py` line 56:
```python
options.add_argument("--headless")  # Uncomment this line
```

---

## 🎉 Features

### HMECloud Automation:
- ✅ Auto-login
- ✅ Store selection
- ✅ Date selection
- ✅ Report download
- ✅ Error handling
- ✅ Progress tracking

### Data Processing:
- ✅ Unlimited rows (not just 197!)
- ✅ Auto-store detection
- ✅ Column cleanup
- ✅ Formula concatenation
- ✅ Pivot table refresh
- ✅ Date updates
- ✅ Auto-backup

### User Interface:
- ✅ Command line (simple)
- ✅ Interactive mode (easy)
- ✅ Web interface (beautiful)
- ✅ Real-time progress
- ✅ Error messages

---

## 💡 Tips

1. **Run daily:** Set up as a scheduled task (cron job)
2. **Check logs:** Watch browser window to see what's happening
3. **Verify data:** Always check final report in Excel
4. **Keep backups:** Script auto-creates backups before modifying template

---

## 🚀 Quick Reference

| What | Command |
|------|---------|
| Full automation | `python3 full_automation.py` |
| Download only | `python3 hmecloud_automation.py` |
| Process only | `python3 complete_automation.py` |
| Web interface | `streamlit run app_integrated.py` |
| Help | `python3 full_automation.py --help` |

---

## 📞 Support

If something doesn't work:

1. Check error messages
2. Verify credentials
3. Check internet connection
4. Update ChromeDriver
5. Re-install packages: `pip3 install -r requirements.txt`

---

**Built for KFC Guyana 🍗**

*Last updated: November 5, 2024*

