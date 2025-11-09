# 🚀 Setup Guide for Windows (HP Laptop)

Complete step-by-step guide to set up and run the KFC Guyana Drive-Thru Automation on a new Windows laptop.

---

## 📋 Prerequisites Checklist

Before you begin, make sure you have:
- ✅ Windows 10 or Windows 11
- ✅ Internet connection
- ✅ Google Chrome browser installed
- ✅ Microsoft Excel installed (for macro functionality)
- ✅ Git installed (or download ZIP from GitHub)

---

## 🔧 Step 1: Install Python

### Option A: Download from Python.org (Recommended)

1. Go to https://www.python.org/downloads/
2. Download **Python 3.9 or newer** (latest stable version)
3. Run the installer
4. **IMPORTANT:** Check the box "Add Python to PATH" during installation
5. Click "Install Now"
6. Wait for installation to complete

### Verify Python Installation

Open **Command Prompt** (Press `Win + R`, type `cmd`, press Enter) and run:

```bash
python --version
```

You should see something like: `Python 3.9.x` or `Python 3.11.x`

If you see an error, Python is not in your PATH. Reinstall Python and make sure to check "Add Python to PATH".

---

## 📥 Step 2: Download the Project

### Option A: Using Git (Recommended)

1. Open **Command Prompt**
2. Navigate to where you want to save the project (e.g., Desktop):
   ```bash
   cd Desktop
   ```
3. Clone the repository:
   ```bash
   git clone <your-repository-url>
   ```
4. Navigate into the project folder:
   ```bash
   cd Automation
   ```

### Option B: Download as ZIP

1. Download the project as ZIP from GitHub
2. Extract the ZIP file to your Desktop (or desired location)
3. Open **Command Prompt** and navigate to the extracted folder:
   ```bash
   cd Desktop\Automation
   ```

---

## 📦 Step 3: Install Project Dependencies

1. In **Command Prompt**, make sure you're in the project folder:
   ```bash
   cd Desktop\Automation
   ```
   (Or wherever you saved the project)

2. Install all required packages:
   ```bash
   pip install -r requirements.txt
   ```

   This will install:
   - `pandas` - Data processing
   - `openpyxl` - Excel file manipulation
   - `selenium` - Browser automation
   - `webdriver-manager` - Automatic ChromeDriver management
   - `streamlit` - Web interface (optional)
   - `xlwings` - Excel macro execution (optional, but recommended)

3. Wait for installation to complete (may take 2-5 minutes)

### Verify Installation

Test that packages are installed:
```bash
python -c "import selenium; print('Selenium installed!')"
python -c "import openpyxl; print('Openpyxl installed!')"
```

---

## 🌐 Step 4: Install ChromeDriver

The project uses `webdriver-manager` which **automatically downloads** ChromeDriver for you. However, you need Google Chrome installed.

### Install Google Chrome (if not already installed)

1. Go to https://www.google.com/chrome/
2. Download and install Chrome
3. Make sure Chrome is up to date

### Verify ChromeDriver (Automatic)

The `webdriver-manager` package will automatically download the correct ChromeDriver version when you first run the script. No manual setup needed!

---

## 🔐 Step 5: Configure Credentials

The project currently has credentials hardcoded. You can either:

### Option A: Use Default Credentials (Already Set)

The default credentials are already configured in the code:
- Username: `doudit@hotmail.com`
- Password: `Kfcguy123!@#`

If these are correct, you can skip this step.

### Option B: Update Credentials in Code

1. Open the file: `src\automation\hmecloud.py`
2. Find these lines (around line 33-34):
   ```python
   USERNAME = "doudit@hotmail.com"
   PASSWORD = "Kfcguy123!@#"
   ```
3. Replace with your credentials:
   ```python
   USERNAME = "your-email@example.com"
   PASSWORD = "your-password"
   ```
4. Save the file

---

## ✅ Step 6: Test the Setup

### Quick Test

Run a simple test to verify everything works:

```bash
python -m automation.hmecloud
```

Or use the test script:

```bash
python scripts\test_store_selection.py
```

**Expected behavior:**
- Chrome browser should open automatically
- Script will log into HMECloud
- Navigate to reports
- Select a store and date
- Download an Excel file

If you see any errors, check the troubleshooting section below.

---

## 🎯 Step 7: Run the Full Automation

### Option 1: Run Single Store Download

```bash
python scripts\test_store_selection.py
```

### Option 2: Run Full Automation (All Stores)

```bash
python scripts\full_automation.py
```

### Option 3: Process Existing Files Only

If you already have downloaded Excel files:

```bash
python scripts\full_automation.py --skip-download
```

---

## 📁 Project Structure

After setup, your project should look like this:

```
Automation/
├── src/
│   └── automation/
│       ├── hmecloud.py          # HME Cloud automation
│       ├── run_macro.py          # Excel macro runner
│       ├── transform_data.py     # Data transformation
│       └── ...
├── scripts/
│   ├── full_automation.py        # Complete automation
│   ├── test_store_selection.py   # Test single store
│   └── ...
├── data/
│   ├── downloads/                # Downloaded Excel files
│   └── templates/                # Excel templates
├── requirements.txt              # Python dependencies
├── README.md                     # Main documentation
└── SETUP_GUIDE.md               # This file
```

---

## 🐛 Troubleshooting

### Problem: "python is not recognized"

**Solution:**
- Python is not in your PATH
- Reinstall Python and check "Add Python to PATH"
- Or use `py` instead of `python`:
  ```bash
  py -m pip install -r requirements.txt
  ```

### Problem: "pip is not recognized"

**Solution:**
- Use `python -m pip` instead:
  ```bash
  python -m pip install -r requirements.txt
  ```

### Problem: "ChromeDriver not found" or Selenium errors

**Solution:**
- Make sure Google Chrome is installed and up to date
- The `webdriver-manager` should handle this automatically
- If issues persist, manually download ChromeDriver from https://chromedriver.chromium.org/

### Problem: "No module named 'xlwings'"

**Solution:**
- Install xlwings:
  ```bash
  pip install xlwings
  ```
- Note: xlwings requires Microsoft Excel to be installed

### Problem: "Login failed" or "Could not find username field"

**Solution:**
- Check your internet connection
- Verify credentials in `src\automation\hmecloud.py`
- HMECloud website might have changed - check if the site is accessible

### Problem: "Excel file not opening after macro"

**Solution:**
- Make sure Microsoft Excel is installed
- On Windows, Excel should open automatically
- If using xlwings, ensure Excel is properly installed

### Problem: "Permission denied" or file access errors

**Solution:**
- Make sure you have write permissions in the project folder
- Try running Command Prompt as Administrator
- Check if any Excel files are open (close them first)

---

## 🔄 Daily Usage

Once setup is complete, running the automation is simple:

### Daily Workflow

1. Open **Command Prompt**
2. Navigate to project folder:
   ```bash
   cd Desktop\Automation
   ```
3. Run the automation:
   ```bash
   python scripts\test_store_selection.py
   ```
4. Wait for the script to complete (5-10 minutes)
5. Excel file will be downloaded and processed automatically
6. Excel will open automatically when done

---

## 📝 Quick Reference Commands

| Task | Command |
|------|---------|
| **Install dependencies** | `pip install -r requirements.txt` |
| **Test single store** | `python scripts\test_store_selection.py` |
| **Full automation** | `python scripts\full_automation.py` |
| **Process existing files** | `python scripts\full_automation.py --skip-download` |
| **Check Python version** | `python --version` |
| **Check installed packages** | `pip list` |

---

## 🆘 Getting Help

If you encounter issues:

1. **Check the error message** - It usually tells you what's wrong
2. **Check this guide** - Review the troubleshooting section
3. **Check README.md** - More detailed documentation
4. **Verify all steps** - Make sure you completed all setup steps

---

## ✅ Setup Complete Checklist

Before running automation, verify:

- [ ] Python is installed and in PATH
- [ ] All dependencies installed (`pip install -r requirements.txt`)
- [ ] Google Chrome is installed
- [ ] Microsoft Excel is installed (for macro functionality)
- [ ] Credentials are configured correctly
- [ ] Test run completed successfully
- [ ] Project folder structure is correct

---

## 🎉 You're Ready!

Once all steps are complete, you can start using the automation. The script will:

1. ✅ Automatically log into HMECloud
2. ✅ Navigate to Raw Car Data Report
3. ✅ Select store and date
4. ✅ Download Excel file
5. ✅ Run DT macro on the file
6. ✅ Open Excel automatically when done

**Total time: ~5-10 minutes** (fully automated!)

---

**Need help?** Check the troubleshooting section or review the main README.md file.

