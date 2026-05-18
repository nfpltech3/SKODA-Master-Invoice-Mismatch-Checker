# SKODA Item Mismatch Checker User Guide

## Introduction
The **SKODA Item Mismatch Checker** is an automated desktop audit application developed for Nagarkot Forwarders Pvt. Ltd. It conducts a comprehensive three-way reconciliation between your **SKODA Master Sheet**, internal **Item Report**, and **Extracted Invoice** data. The tool automatically normalizes regional differences (such as country names, numeric formats, and date layouts) and cross-references them to instantly flag and highlight discrepancies in CTH codes, quantities, descriptions, duty rates, and values.

---

## How to Use

### 1. Launching the App
1. Locate the compiled executable file: `Skoda_Mismatch_Checker.exe`.
2. Double-click to launch the application. The program will start in a clean, full-screen view.
3. Make sure that the `logo.png` image is present in the same directory as the executable if running outside the packed environment, though the compiled `.exe` bundles this asset automatically.

### 2. The Workflow (Step-by-Step)
1. **Load Master Sheet**: Click the **📂 Master Sheet** button and select your SKODA Master file.
   - *Constraint: The spreadsheet must contain a sheet named exactly `Master Sheet` with columns `PartNo`, `CTH1`, `Part Suffix`, `Basic Duty Rate`, `IGST Notn Sr No`, and `Description`.*
2. **Load Item Report**: Click the **📂 Item Report** button and select your internal Item Report file.
   - *Constraint: The data must reside in a sheet named `Sheet1` with columns `Job No`, `Job Date`, `Invoice No`, `Invoice Date`, `Model`, `Amount`, `CTH`, `Generic Description`, `IGST Notification SrNo`, and `Product Desc`.*
3. **Load Extracted Invoice**: Click the **📂 Extracted Invoice** button and select the invoice extraction file.
   - *Constraint: The file must have a sheet named `Sheet1` containing core fields such as `Invoice Number`, `Invoice Date`, `Mat. NO.`, `Country of Origin`, `Quantity`, `Total Price`, and `HS Code`. The tool automatically handles alternative column headers (e.g., `Part Number`, `COO`, `Qty`, `HS-CODE`, `VALUE OF GOODS`).*
   - *Note: The **▶ Run** button will remain grayed out and unclickable until all three files are successfully uploaded.*
4. **Execute Reconciliation**: Click the bright red **▶ Run** button. The button will shift to a dark red **⌛ Processing...** state while compiling your audit.
5. **Review Results**: Scroll through the Data Preview table to audit the discrepancies.
   - **Light Red (`#FFEBEE`) rows** denote active `MISMATCH` errors (e.g., mismatched CTHs, description differences, quantity or pricing issues).
   - **White (`#FFFFFF`) rows** denote `INFO` items (e.g., "CTH Match" warnings confirming that the chapter-level first 2 digits match even if master data is missing).
   - *Tip: The table columns are sized dynamically according to text length, allowing you to read all Remarks in full. You can also double-click on any cell to edit its value directly within the table before exporting!*
6. **Export Findings**: Click the **💾 Export Excel** button to download a formatted spreadsheet of the results. 
   - *Note: The export dialogue will automatically suggest a filename appended with a dynamic date and time stamp (e.g., `comparison_mismatch_report_20260518_112850.xlsx`) to prevent file overwrite.*
7. **Reset Application**: Click the gray **🔄 Clear All** button to safely clear the active memory, reset all file paths, and prepare the interface for the next audit.

---

## Interface Reference

| Control / Input | Description | Expected Format / Constraint |
| :--- | :--- | :--- |
| **📂 Master Sheet** | Opens selection window for the SKODA Master database. | `.xlsx` or `.xls` containing `Master Sheet` sheet |
| **📂 Item Report** | Opens selection window for the internal customs report. | `.xlsx` or `.xls` containing `Sheet1` sheet |
| **📂 Extracted Invoice** | Opens selection window for the parser-extracted invoice. | `.xlsx` or `.xls` containing `Sheet1` sheet |
| **▶ Run** | Begins the comparison engine. Muted gray when locked, bright red when active. | Enabled only after all 3 files are successfully loaded |
| **💾 Export Excel** | Saves the modified mismatch list to an styled Excel spreadsheet. | Generates formatted `.xlsx` with bold values and red highlights |
| **🔄 Clear All** | Wipes the current session data, file paths, and clears the table. | Resets the GUI to its clean, initial startup state |
| **Data Preview Grid** | Renders mismatches side-by-side. Supports double-click inline edits. | Segoe UI 10 font, styled color tags based on `Status` |

---

## Troubleshooting & Validations

If you see an error or incorrect behavior, please consult this troubleshooting guide:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **Missing required columns in [File] / [Sheet]** | The comparison engine failed because mandatory headers are absent. | Open the specific spreadsheet and verify the column headers exactly match the required formats described in Step 2. |
| **Master Sheet load failed** | The master database could not be loaded into the pandas engine. | Ensure the file is not corrupted, is not currently locked or open in Excel, and has a sheet named exactly `Master Sheet`. |
| **Item Report load failed** | The internal Item Report file could not be read. | Check that the data is not blank and is placed inside a sheet called `Sheet1`. |
| **Extracted Invoice load failed** | The parser-extracted invoice file is unreadable. | Verify that the document is a valid Excel workbook and is placed in `Sheet1`. |
| **Error during comparison** | An unexpected exception occurred during data processing. | The tool catches all crashes safely. Locate the `mismatch_checker.log` in the application folder to view the technical error trace, and share it with support. |
| **Export Error** | The tool could not save the final mismatch report. | Ensure the destination folder is not write-protected and that a previously exported report of the exact same name is not currently open in Microsoft Excel. |
