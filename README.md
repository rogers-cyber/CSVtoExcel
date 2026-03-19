# CSVtoExcel – Desktop CSV to Excel Converter v1.0.0

CSVtoExcel v1.0.0 is a professional desktop tool for fast CSV to Excel (.xlsx) conversion with modern GUI, multi-file support, preview, and SQLite-based conversion history.

This version introduces **multi-threaded conversion**, **batch CSV handling**, **preview of first 100 rows**, **encoding selection**, and full **conversion history management**. Users can browse files or folders, drag & drop CSVs, preview data, convert to Excel, and manage past conversions in a responsive, styled PySide6 interface.

------------------------------------------------------------
WINDOWS DOWNLOAD (EXE)
------------------------------------------------------------

Download the latest Windows executable from GitHub Releases:

https://github.com/rogers-cyber/CSVtoExcel/releases

- No Python installation required
- Portable standalone executable
- Ready-to-run on Windows
- Optimized for multi-threaded batch CSV conversion

------------------------------------------------------------
DISTRIBUTION
------------------------------------------------------------

CSVtoExcel is a paid / commercial desktop utility.

This repository/documentation may include:

- Production-ready Python source code
- Compiled Windows executables
- Commercial licensing terms (see LICENSE / sales page)

Python is not required when using the compiled executable version.

------------------------------------------------------------
FEATURES
------------------------------------------------------------

CORE CAPABILITIES

- ⚡ Batch CSV → Excel conversion
- 👁 Preview first 100 rows of CSV before conversion
- 📝 Optional header row detection
- 🔤 Encoding selection (UTF-8, Latin-1, UTF-16, CP1252)
- 🕑 Multi-threaded processing for faster conversion
- 💾 Save Excel output in custom destination folder or file
- 🗃 Conversion history with SQLite database
- 🔄 Re-export previous conversions with original settings
- 🗑 Delete history records without removing files
- 🎨 Responsive UI with alternating row colors and hover highlight
- Drag & Drop CSV files or folders directly into the app
- Thread-safe, non-blocking operations

CONVERSION MODES

- Single CSV conversion  
  Convert one CSV to Excel with header and encoding options.

- Batch CSV conversion  
  Select multiple CSV files or a folder containing CSVs; processed in a separate thread for responsiveness.

HISTORY MANAGEMENT

- View last 500 conversions
- Open folder containing Excel output
- Re-export Excel from original CSV
- Delete history record without affecting files

UI & PREVIEW

- CSV preview table (first 100 rows by default)
- Sorting enabled on preview and history tables
- Adjustable preview row count (5–10000)
- Clean modern interface with toolbar-style buttons
- Status messages, progress bar, and log of processed files

------------------------------------------------------------
INSTALLATION (SOURCE CODE)
------------------------------------------------------------

1. Clone the repository:

```bash
git clone https://github.com/rogers-cyber/CSVtoExcel.git
```

2. Navigate to project directory:

```bash
cd CSVtoExcel
```

3. Install required dependencies:

```bash
pip install PySide6 xlsxwriter
```

4. Run the application:

```bash
python CSVtoExcel.py
```

------------------------------------------------------------
BUILD WINDOWS EXECUTABLE
------------------------------------------------------------

1. Install PyInstaller:

```bash
pip install pyinstaller
```

2. Build executable:

```bash
pyinstaller --onefile --windowed --name "CSVtoExcel" --icon=logo.ico CSVtoExcel.py
```

The compiled executable will appear in:

```
dist/CSVtoExcel.exe
```

------------------------------------------------------------
USAGE GUIDE
------------------------------------------------------------

1. Open CSV or Folder

- Click **Open CSV** to select one or more files.
- Click **Open Folder** to select a folder containing CSVs.

2. Preview

- Preview first 100 rows of the CSV (adjustable with spin box)
- Toggle **Has header row** if CSV has no header
- Choose file encoding (UTF-8, Latin-1, UTF-16, CP1252)

3. Choose Destination

- Optional: select folder or specific Excel filename
- Multi-file batch conversion will require a folder

4. Convert

- Click **Convert → Excel** to start conversion
- Progress bar and status messages will update in real-time

5. History

- Click **History** to view past conversions
- Re-export, open folder, or delete records
- All conversions are stored in SQLite database located in app data folder

6. Help

- Click **Help** for detailed user guide

7. Drag & Drop

- Drag CSV files or folders onto the main window to automatically load

------------------------------------------------------------
LOGGING & ERROR HANDLING
------------------------------------------------------------

- Status messages display current processing file
- Progress bar shows conversion progress
- Errors reported in message boxes
- Recovered conversions are recorded in history
- Robust CSV reading with encoding fallback (`errors="replace"`)

------------------------------------------------------------
REPOSITORY STRUCTURE
------------------------------------------------------------

CSVtoExcel/

├── CSVtoExcel.py  
├── logo.ico  
├── README.md  
├── LICENSE  
├── csv_to_excel_history.db (generated on first run)  

------------------------------------------------------------
DEPENDENCIES
------------------------------------------------------------

Python 3.10+  

Libraries used:

- PySide6
- xlsxwriter
- sqlite3
- datetime
- pathlib
- csv
- sys, os, traceback

------------------------------------------------------------
INTENDED USE
------------------------------------------------------------

Ideal for:

- Business users needing CSV → Excel conversion
- Batch CSV conversion workflows
- Data analysis preparation
- Keeping track of past conversions
- Anyone needing quick and clean Excel exports from CSV files

------------------------------------------------------------
ABOUT
------------------------------------------------------------

CSVtoExcel is developed by MateTools for professional offline productivity on Windows.

Website:

https://matetools.gumroad.com

© 2026 MateTools  
All rights reserved.

------------------------------------------------------------
LICENSE
------------------------------------------------------------

CSVtoExcel – Desktop CSV to Excel Converter v1.0.0 – License Agreement

Copyright (c) 2026 MateTools. All rights reserved.

This software is provided under a **single-user commercial license**. By using this software, you agree to the following terms:

1. License Grant
   - You are granted a non-exclusive, non-transferable license to use CSVtoExcel v1.0.0 for personal or commercial purposes.
   - You may install and use the software on any number of computers you personally control.

2. Restrictions
   - Do NOT resell, redistribute, or sublicense this software.
   - Do NOT modify, reverse-engineer, decompile, or attempt to derive the source code, except where permitted by law.
   - Do NOT claim ownership of this software.
   - Any unauthorized distribution, copying, or sharing of this software is strictly prohibited.

3. Ownership
   - All rights, title, and interest in CSVtoExcel v1.0.0 remain with MateTools.
   - No part of this software may be copied, reused, or incorporated into another product without prior written permission.

4. Support
   - Support, updates, and bug fixes are provided at the discretion of MateTools.
   - For questions, feature requests, or business licensing, contact: rogermodu@gmail.com

5. Disclaimer
   - CSVtoExcel v1.0.0 is provided "as-is" without warranty of any kind, express or implied.
   - MateTools is not liable for any damages, data loss, system outages, security incidents, or other consequences resulting from the use or misuse of this software.

---

By using CSVtoExcel v1.0.0, you acknowledge that you have read, understood, and agreed to this license.