# 🚀 QUICK START GUIDE - Drive-Thru Automation

## ✅ What's Automated:

This automation handles **Steps 6-10** from your procedure:
- ✅ Data transformation (Ctrl+D macro replacement - unlimited rows!)
- ✅ Paste into Drive Thru template
- ✅ Concatenate formulas
- ✅ Refresh pivot tables
- ✅ Update dates

## 📋 Daily Workflow (2 Steps):

### STEP 1: Download Files Manually (15 minutes)

1. Go to https://hmecloud.com/
2. Login with your credentials
3. Go to Reports → Raw Car Data Report
4. For each of the 6 stores:
   - Select store
   - Select yesterday's date
   - Click "View Report"
   - Click "Export" → "Microsoft Excel"
5. Save all 6 files to: `/Users/sanctum/Desktop/Automation/downloads/`

### STEP 2: Run Automation (2 minutes)

```bash
cd /Users/sanctum/Desktop/Automation
python3 complete_automation.py
```

**That's it!** The script does everything else automatically.

---

## 📁 Folder Structure:

```
Automation/
├── downloads/                    ← Put your downloaded files here
│   ├── Raw Car Data Report (1).xlsx
│   ├── Raw Car Data Report (2).xlsx
│   └── ... (all 6 stores)
│
├── templates/
│   └── Drive Thru Optimization - KFC Guyana (16-10)-copy.xlsx
│
├── complete_automation.py        ⭐ RUN THIS!
├── transform_data.py            (transformation module)
└── template_operations.py       (template operations module)
```

---

## ⚙️ Configuration:

Edit `complete_automation.py` if needed:

```python
# Line 18-19: Folder locations
DOWNLOADS_FOLDER = "/Users/sanctum/Desktop/Automation/downloads"
TEMPLATE_PATH = "/Users/sanctum/Desktop/Automation/templates/..."

# Line 23: Yellow-headed columns (formulas to copy down)
FORMULA_COLUMNS = [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]

# Line 26-33: Where to update dates in each sheet
DATE_CONFIGS = {
    "Consol Wkly time trnd": "A1",  # Change cell reference as needed
    "Summary - Stores": "B2",        # etc.
    ...
}
```

---

## 📊 What the Script Does:

```
1. Finds all .xlsx files in downloads/
   ↓
2. Transforms each file (Ctrl+D replacement)
   - Works for UNLIMITED rows (not just 197!)
   - Extracts store name automatically
   - Removes problematic columns F, G, J, K
   ↓
3. Creates backup of template
   ↓
4. Pastes all data into AllStores (or Raw Data) sheet
   ↓
5. Copies formulas down (yellow-headed columns)
   ↓
6. Sets pivot tables to refresh on open
   ↓
7. Updates dates in 6 specified sheets
   ↓
8. Saves template
   ↓
✅ DONE!
```

---

## ⏱️ Time Savings:

**Manual (all steps):** 40-50 minutes

**With automation:**
- Manual download: 15 minutes
- Script runtime: 2 minutes
- **Total: 17 minutes**

**Time saved: 60%!**

---

## 🔧 First Time Setup:

1. Make sure Python libraries are installed:
```bash
pip3 install pandas openpyxl
```

2. Put template in `templates/` folder (already done ✅)

3. Create `downloads/` folder (already done ✅)

---

## 📝 Notes:

- **Formula columns:** Update `FORMULA_COLUMNS` list with actual yellow-headed column numbers
- **Date cells:** Update `DATE_CONFIGS` with exact cell references for each sheet
- **Backup:** Script automatically creates backup before modifying template
- **Pivot tables:** Will fully refresh when you open file in Excel

---

## ❓ Need Help?

If you need to find:
- Which columns have formulas → Open template, look for yellow headers
- Which cells need date updates → Check each of the 6 sheets

---

## 🎯 Bottom Line:

**Manual download (unavoidable) + Full automation (everything else) = Best solution!**

Just download the files, run the script, and you're done! 🎉



