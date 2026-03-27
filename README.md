markdown
# 🗄️ GSTR-1 Cross verifier

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![UI](https://img.shields.io/badge/UI-Tkinter%20(Win95)-lightgrey)

**GSTR-1 Cross verifier** is a desktop utility designed for accounting and finance teams. It automates the painful process of finding "sneaky" backdated or modified entries in live ERP systems after monthly GST (Goods & Services Tax) reports have already been locked and filed. 

Wrapped in a nostalgic, authentic Windows 95/98 graphical interface, this tool uses the power of `pandas` to process hundreds of thousands of rows in seconds, outputting a beautiful, color-coded Excel audit report.

---

## ✨ Features
* **🕰️ Authentic Retro UI:** Built with Tkinter using authentic Win95 hex codes (`#D4D0C8`), sunken borders, and block progress bars.
* **🧠 Smart Key Detection:** Automatically analyzes sheet headers (e.g., `b2b`, `cdnr`, `b2cs`) to mathematically determine the correct Primary Keys (Invoice Numbers, GSTINs, Place of Supply, etc.) without hardcoding.
* **🛡️ Intelligent Comparison:** 
  * Converts dates (`2025-04-01` vs `01-Apr-2025`) mathematically to prevent false positives.
  * Ignores floating-point number discrepancies (e.g., `5000` vs `5000.0`).
  * Case-insensitive string matching.
* **📊 Beautiful Excel Output:** Generates a borderless, zebra-striped Excel file with color-coded statuses (🟩 `New Backdated Entry` / 🟨 `Modified`).
* **🚀 Auto-Launch:** Automatically opens the generated Excel file upon completion.

---

## 🛠️ Installation (For Developers)

1. Clone the repository:
   ```bash
   git clone https://github.com/YourUsername/Retro-GST-Auditor.git
   cd Retro-GST-Auditor
   ```
2. Install the required dependencies:
   ```bash
   pip install pandas openpyxl numpy
   ```
3. Run the application:
   ```bash
   python GST_Compare_Tool.py
   ```

---

## 📦 Building a Standalone `.exe` (For End Users)

If you want to distribute this tool to your accounting team so they can use it without installing Python, you can package it into a single `.exe` file using **PyInstaller**.

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Build the executable:
   ```bash
   pyinstaller --onefile --windowed GST_Compare_Tool.py
   ```
3. Look inside the newly created `dist/` folder. You will find `GST_Compare_Tool.exe`. You can share this single file with your team!

---

## 📖 How to Use the Tool

1. **Step 1 (Past Reports):** Click "Browse" and select all the locked monthly GST Excel files you have previously downloaded (e.g., April to February). *You can select multiple files at once.*
2. **Step 2 (YTD Master):** Click "Browse" and select the currently exported Year-To-Date file from your live ERP system.
3. **Step 3 (Configuration):** 
   * The tool will automatically read the YTD file and generate a list of available sheets (ignoring instruction/help sheets). 
   * Check the boxes for the sheets you want to audit (e.g., `b2b`, `cdnr`).
   * Choose where you want to save the final Output file.
4. **Step 4 (Execution):** Click **Execute Audit**. The progress bar will fill up, and the final Excel report will automatically open showing exactly which rows were added or modified in past months!

---

## 🏗️ Under the Hood

The comparison engine relies on a Left Merge mechanism via `pandas`:
1. Concatenates all historical monthly files into a single snapshot DataFrame.
2. Performs a `pd.merge(YTD, Snapshot, how='left', indicator=True)`.
3. Flags rows marked as `left_only` as entirely new (backdated) entries.
4. Iterates through the remaining rows, comparing the `_old` (snapshot) columns against the current YTD columns to identify exact modifications.

---

## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
