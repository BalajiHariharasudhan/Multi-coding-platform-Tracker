# 🚀 Automated Competitive Programming Excel Tracker

An automated Python tool that syncs your accepted problem submissions from **LeetCode**, **Codeforces**, and **AtCoder** directly into your personal Excel daily tracker spreadsheet (`.xlsx`).

---

## ✨ Features

- **🌐 Multi-Platform Support**: Automatically queries:
  - **LeetCode** (via GraphQL API)
  - **Codeforces** (via Official REST API)
  - **AtCoder** (via Kenkoooo AtCoder Problems API)
  - **GeeksforGeeks** (via profile activity parsing)
- **🏷️ Automated Metadata Extraction**:
  - Exact submission date (`DD.MM.YY`)
  - Problem title (`<Problem Title> - <Platform>`)
  - Submission URL link
  - Difficulty level (`Easy` / `Medium` / `Hard`)
  - Topic tags (`Array`, `Dynamic Programming`, `Graph`, etc.)
- **🛡️ Duplicate Prevention**: Prevents double-entry by checking problem URLs and titles before writing.
- **📊 Excel Layout & Formula Preservation**: Keeps `=SUM()` total formulas, profile headers, and summary statistics (`COUNTIF`) completely intact.
- **🔒 Safe File Locking**: Safely detects if Microsoft Excel is open and logs warnings instead of crashing.
- **⏰ Windows Task Scheduler Integration**: Runs seamlessly in the background on a schedule (e.g. every 30 minutes).

---

## 📁 Repository Structure

```text
LeetCodeTracker/
├── config.py              # Active configuration file (handles & Excel paths)
├── config.example.py      # Template configuration file for new users
├── tracker.py             # Main multi-platform automation script
├── run_tracker.bat        # Windows Task Scheduler batch runner
├── requirements.txt       # Python dependencies
├── README.md              # Project documentation & setup guide
└── .gitignore             # Git ignore file
```

---

## 🛠️ Step-by-Step Setup Guide

### Step 1: Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/LeetCodeTracker.git
cd LeetCodeTracker
```

### Step 2: Install Dependencies

Make sure Python 3.8+ is installed. Then run:

```bash
pip install -r requirements.txt
```

### Step 3: Configure `config.py`

Create or edit `config.py` (you can copy `config.example.py`):

```python
# Handles / Usernames
LEETCODE_USERNAME = "balaji_harihara_sudhan"
CODEFORCES_HANDLE = "bala3507563"
ATCODER_HANDLE = "balaji_8567"
GFG_HANDLE = "bala35t75l"

# Path to your Excel Daily Tracker
EXCEL_FILE = r"C:\Users\YourName\Desktop\Balaji_Daily_Tracker.xlsx"
SHEET_NAME = "Daily Task"
START_ROW = 14
```

### Step 4: Run Initial Test

Close Excel if it's open, then execute:

```bash
python tracker.py
```

Check console output and `tracker_log.txt` to verify new problems are logged into your Excel workbook!

---

## ⏰ Step 5: Automate with Windows Task Scheduler

To make the tracker run automatically every 30 minutes:

1. Press <kbd>Win</kbd> + <kbd>R</kbd>, type **`taskschd.msc`**, and press **Enter**.
2. Click **Create Task...** on the right panel.
3. **General Tab**:
   - Name: `LeetCode Daily Tracker`
4. **Triggers Tab**:
   - Click **New...**
   - Begin the task: **On a schedule** -> **Daily**
   - Under *Advanced settings*:
     - Check **Repeat task every:** `30 minutes`
     - For a duration of: `Indefinitely` (or `1 day`)
5. **Actions Tab**:
   - Click **New...**
   - Action: **Start a program**
   - Program/script: `C:\LeetCodeTracker\run_tracker.bat`
   - Start in: `C:\LeetCodeTracker`
6. Click **OK** to save.

---

## 📄 License

Distributed under the MIT License. Feel free to use, modify, and share!
