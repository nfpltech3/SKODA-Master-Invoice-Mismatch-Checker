# SKODA Item Mismatch Checker v2

## Overview
The **SKODA Item Mismatch Checker** is a Python-based desktop application (built with `tkinter` and `pandas`) designed to automate the reconciliation of customs and invoice data for Nagarkot Forwarders Pvt. Ltd. 

It takes three inputs:
1. **Master Sheet**: The source of truth for Part Numbers, CTH, and Duty Rates.
2. **Item Report**: The internal system report containing Job numbers and currently assigned values.
3. **Extracted Invoice**: The structured output from various invoice parsers (e.g., standard Skoda AS, VAG Group VW/Audi, or Premium Sound Solutions).

It normalizes data (Countries, Quantities, Prices, Descriptions) and merges these files to detect mismatched fields, missing records, and duplicate occurrences. 

## Key Features
- **Multi-Format Support**: Automatically aliases varying headers (e.g., `Part No.`, `Article No` -> `Mat. NO.`) to support different invoice extractor formats out-of-the-box.
- **Advanced Normalization**: Intelligent parsing of European number formats (e.g., `2.236,90`), date formats, country codes (e.g., `TW` to `TAIWAN`), and stripping generic models/descriptions.
- **Robust Error Handling**: Silent logging to `mismatch_checker.log` without crashing the GUI.
- **Interactive UI**: Nagarkot-branded Tkinter interface with a double-click-to-edit Treeview for manual override of detected mismatches.
- **Excel Export**: Generates a clean, styled Excel `.xlsx` report of all discrepancies.

## Technical Stack
- **Python 3.10+**
- **Pandas**: Core engine for dataframe merging and conflict detection.
- **OpenPyXL**: For styling and saving the exported Excel reports.
- **Tkinter**: GUI framework.

## Project Structure
- `mismatch_checker.py`: The main application code containing business logic, GUI implementation, and data normalization loops.
- `USER_GUIDE.md`: The end-user instruction manual.
- `logo.png`: Branded logo used in the application header.
- `mismatch_checker.log`: Auto-generated log file tracking execution history and tracebacks.

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```bash
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. Run application
```bash
python mismatch_checker.py
```

---

### Build Executable (For Desktop Apps)

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Build using the included Spec file (Ensure you do not run mismatch_checker.py directly for builds):
```bash
pyinstaller Skoda_Mismatch_Checker.spec
```

3. Locate Executable:
The application will be generated in the `dist/` folder.

## Version History

| Version | Release Date | Key Enhancements & Features |
| :---: | :---: | :--- |
| **v2.0** | 2026-05-19 | **Advanced Logging:** Real-time log rotation (5MB cap) for file fingerprints and per-field match tracebacks.<br>**Audit Trails:** Added tracking for unmatched column aliases and dropped blank rows.<br>**Verbose Mode:** Added `VERBOSE_LOG` constant for deep trace debugging. |
| **v1.1** | 2026-05-16 | **Format Support:** Added header mapping for VW/Audi and Premium Sound Solutions invoices.<br>**Data Normalization:** Refined country extraction logic (e.g., Taiwan aliases). |
