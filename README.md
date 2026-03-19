# 🤖 CSV Report Automation

Automates the processing of CSV data files and generates structured Excel reports — including cleaned data, summary statistics, and a data quality log.

Built with Python and Pandas. No manual work required after setup.

---

## 📋 What It Does

1. **Loads** all CSV files from an input folder automatically
2. **Cleans** the data — removes duplicates, strips whitespace, drops empty rows
3. **Analyzes** the data — generates summary statistics for all numeric columns
4. **Exports** a timestamped Excel report with 3 sheets:
   - `Clean Data` — processed, ready-to-use dataset
   - `Summary` — descriptive statistics per numeric column
   - `Quality Report` — before/after metrics (rows removed, nulls, etc.)

---

## 📁 Project Structure

```
csv-report-automation/
├── automation.py          # Main script
├── requirements.txt       # Dependencies
├── data/
│   ├── input/             # Place your CSV files here
│   │   └── sales_january.csv   # Example dataset
│   └── output/            # Generated reports appear here
└── README.md
```

---

## ⚡ Quick Start

**1. Clone the repository**
```bash
git clone https://github.com/azahidgarcia/python-csv-report-automation.git
cd python-csv-report-automation
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Add your CSV files**
```
Place your .csv files inside the data/input/ folder.
```

**4. Run the script**
```bash
python automation.py
```

**5. Get your report**
```
Your Excel report will appear in data/output/ with a timestamp.
```

---

## 📊 Example Output

```
==================================================
  CSV Report Automation Script
  Started: 2024-01-22 10:30:00
==================================================
📂 Found 1 CSV file(s): ['sales_january.csv']
   ✓ Loaded sales_january.csv — 15 rows

🧹 Cleaning complete:
   Duplicates removed : 2
   Empty rows removed : 0
   Nulls remaining    : 1
   Final row count    : 13

✅ Report saved to: data/output/report_20240122_103001.xlsx

🎉 Done! Check the output folder for your report.
==================================================
```

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3.10+ | Core language |
| Pandas | Data processing and cleaning |
| OpenPyXL | Excel report generation |

---

## 🔧 Customization

You can modify `automation.py` to:
- Change input/output folder paths
- Add custom cleaning rules for your data
- Include additional sheets in the report
- Schedule automated runs with cron or Task Scheduler

---

## 👤 Author

**Azahid García** — Python Automation & Data Engineer

- 🌐 [Fiverr Profile](https://fiverr.com)
- 💼 [Upwork Profile](https://upwork.com)
- 🐙 [GitHub](https://github.com/azahidgarcia)

---

## 📄 License

MIT License — free to use and modify.
