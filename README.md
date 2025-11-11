# KFC Guyana Drive-Thru Automation Platform

## ⭐ COMPLETE END-TO-END AUTOMATION

### 🤖 NOW WITH AUTO-LOGIN & AUTO-DOWNLOAD!

**EVERYTHING IS AUTOMATED:**
- ✅ Auto-login to HMECloud
- ✅ Auto-download all reports  
- ✅ Auto-transform data
- ✅ Auto-update template
- ✅ Auto-refresh pivot tables

**Total time: ~5 minutes** (fully automated!)  
**vs Manual: 40+ minutes**  
**Time saved: 90%!** ⚡

---

## 🚀 Super Quick Start

### Option 1: Web Interface (Easiest) ⭐

Launch the user-friendly web interface:

```bash
streamlit run app.py
```

Then open your browser to `http://localhost:8501` and click **"Start Automation"** button.

The automation will open in a separate browser window and run automatically!

---

### Option 2: Command Line

```bash
PYTHONPATH=src python3 scripts/test_store_selection.py
```

**That's it!** The script will:
1. Login to HMECloud automatically
2. Download the report
3. Format it using DT macro
4. Process the data

---

## 📋 Usage Options

### Option 1: Easy Launcher (Recommended) ⭐

```
./scripts/run_automation.sh
```

Interactive menu with all options.

### Option 2: Full Automation

```
PYTHONPATH=src python3 scripts/full_automation.py
```

Downloads from HMECloud + processes everything.

### Option 3: HMECloud Download Only

```
PYTHONPATH=src python3 -m automation.hmecloud
```

Just download reports (interactive mode).

### Option 4: Process Existing Files

```
PYTHONPATH=src python3 scripts/full_automation.py --skip-download
```

If you already have downloaded files.

### Option 5: Web Interface

```
streamlit run apps/app_integrated.py
```

Beautiful web UI at http://localhost:8501

---

## ✨ Features

### HMECloud Integration:
- 🔐 **Auto-login** - Stored credentials, automatic authentication
- 📥 **Auto-download** - Downloads all 6 stores automatically
- 🏪 **Store selection** - Download all or single store
- 📅 **Date selection** - Yesterday (default) or custom date
- 🌐 **Visual browser** - See what's happening in real-time

### Data Processing:
- 🔄 **Data Transformation** - Replaces Ctrl+D macro, unlimited rows!
- 📋 **Auto-paste** - Pastes into Drive-Thru template  
- 📝 **Formula concatenation** - Copies formulas down automatically
- 📊 **Pivot table refresh** - Sets all pivot tables to refresh on open
- 📅 **Date updates** - Updates dates across all sheets
- 💾 **Auto-backup** - Creates backup before modifying template

### User Interface:
- 💻 **Multiple interfaces** - CLI, interactive, web UI
- 📊 **Progress tracking** - See what's happening
- ⚡ **Error handling** - Clear error messages
- 🎯 **Easy to use** - One command does everything

---

## 📁 Project Structure

```
Automation/
├── run_automation.sh           ⭐ Easy launcher
├── full_automation.py          🤖 Complete automation
├── hmecloud_automation.py      🌐 HMECloud login & download
├── complete_automation.py      📊 Data processing
├── app_integrated.py           💻 Web interface
│
├── transform_data.py           (transformation module)
├── template_operations.py      (template operations module)
│
├── downloads/                  📥 Downloaded files
├── templates/                  📊 Drive-Thru Excel template
│
├── AUTOMATION_GUIDE.md         📖 Complete documentation
├── QUICK_START.md              🚀 Quick reference
└── requirements.txt            📦 Dependencies
```

---

## 🎯 Daily Workflow

### Every Morning - ONE COMMAND:

```bash
python3 full_automation.py
```

1. Script logs into HMECloud ✅
2. Downloads all 6 stores ✅
3. Transforms data ✅
4. Updates template ✅
5. You open Excel and click "Refresh All" ✅
6. **Done!** ✅

**Total time: ~5 minutes** (mostly automated)

---

## 🔧 First Time Setup

### 1. Install Dependencies:

```bash
pip3 install -r requirements.txt
```

### 2. Install ChromeDriver:

```bash
brew install chromedriver
```

Or download from: https://chromedriver.chromium.org/

### 3. Run:

```bash
python3 full_automation.py
```

---

## 🔐 Credentials

**Pre-configured and ready to use:**
- Email: `doudit@hotmail.com`
- Password: `Kfcguy123!@#`

Stored in `hmecloud_automation.py` - change if needed.

---

## 🏪 Stores Configured

All 6 KFC Guyana stores are pre-configured:

1. (Ungrouped) 2 Vlissengen Road – KFC
2. (Ungrouped) 5 Mandela - KFC
3. (Ungrouped) Movie Towne - KFC
4. (Ungrouped) Giftland Mall - KFC
5. (Ungrouped) Sheriff Street - KFC
6. (Ungrouped) Providence - KFC

---

## ⏱️ Time Comparison

| Method | Time | Automation |
|--------|------|------------|
| **Manual (old)** | 40-50 min | 0% |
| **Semi-auto** | 17 min | 60% |
| **Full-auto (NEW!)** | **5 min** | **90%** ⚡ |

---

## 📖 Documentation

- **AUTOMATION_GUIDE.md** - Complete guide with all features
- **QUICK_START.md** - Quick reference
- **USAGE.md** - Original usage guide

---

## 🐛 Troubleshooting

### "ChromeDriver not found"
```bash
brew install chromedriver
```

### "Login failed"
Check credentials in `hmecloud_automation.py`

### "No files found"
Make sure downloads completed successfully

### More help:
See **AUTOMATION_GUIDE.md** for detailed troubleshooting

---

## 🎉 What's New

### Version 2.0 - Full Automation!

- ✅ HMECloud auto-login
- ✅ Automatic report downloads
- ✅ Interactive store/date selection
- ✅ Web interface option
- ✅ Easy launcher script
- ✅ Complete documentation
- ✅ Error handling & progress tracking

---

## 📞 Quick Reference

| Task | Command |
|------|---------|
| **Full automation** | `python3 full_automation.py` |
| **Easy launcher** | `./run_automation.sh` |
| **Download only** | `python3 hmecloud_automation.py` |
| **Process only** | `python3 complete_automation.py` |
| **Web UI** | `streamlit run app_integrated.py` |

---

**Built for KFC Guyana** 🍗  
*Version 2.0 - Now with complete HMECloud automation!*

