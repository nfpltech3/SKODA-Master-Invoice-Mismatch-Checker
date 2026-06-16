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
   - *Constraint: The spreadsheet must contain a sheet named exactly `Master Sheet` with columns `PartNo`, `CTH1`, `Basic Duty Rate`, `IGST Notn Sr No`.*
2. **Load Item Report**: Click the **📂 Item Report** button and select your internal Item Report file.
   - *Constraint: The data must reside in a sheet with columns `Job No`, `Job Date`, `Invoice No`, `Invoice Date`, `Model`, `Amount`, `CTH`, `IGST Notification SrNo`, `Product Desc`, `Quantity`, and `Country of Origin`. (The `Unit` column is optional and will be matched if present).*
3. **Load Extracted Invoice**: Click the **📂 Extracted Invoice** button and select the invoice extraction file.
   - *Constraint: The file must have a sheet containing core fields such as `Invoice Number`, `Invoice Date`, `Mat. NO.`, `Country of Origin`, `Quantity`, `Total Price`, and `Description`. (The `Price per PC` column is optional and will be matched against Unit if present). The tool automatically handles alternative column headers (e.g., `Part Number`, `COO`, `Qty`, `HS-CODE`, `VALUE OF GOODS`).*
   - *Note: The **▶ Run** button will remain grayed out and unclickable until all three files are successfully uploaded.*
4. **Execute Reconciliation**: Click the bright red **▶ Run** button. The button will shift to a dark red **⌛ Processing...** state while compiling your audit.
5. **Review Results**: Scroll through the Data Preview table to audit the discrepancies.
   - **Light Red (`#FFEBEE`) rows** denote active `MISMATCH` errors (e.g., mismatched CTHs, description differences, quantity or pricing issues).
   - **White (`#FFFFFF`) rows** denote `INFO` items. When zero mismatches are found, the table will populate with white rows listing every successfully verified **Job Number**.
   - *Note on Missing Master Items:* If a part is missing in the Master Sheet, the row will now automatically pull the CTH codes and Descriptions from both the Extracted Invoice and Item Report into dedicated side-by-side columns.
   - *Tip: The table columns are sized dynamically according to text length, allowing you to read all Remarks in full. You can also double-click on any cell to edit its value directly within the table before exporting!*
6. **Export Findings**: Click the **💾 Export Excel** button to download a formatted spreadsheet of the results. 
   - *Note: This button will be automatically disabled if there are zero mismatches to export.*
   - The exported Excel file includes **AutoFilters**, explicit cell borders, and dynamically sized columns with text wrapping so no descriptions are cut off.
   - *After exporting, a popup will ask if you want to immediately open the Excel file for review.*
7. **Reset Application**: Click the gray **🔄 Clear All** button to safely clear the active memory, reset all file paths, and prepare the interface for the next audit.

---

## Interface Reference

| Control / Input | Description | Expected Format / Constraint |
| :--- | :--- | :--- |
| **📂 Master Sheet** | Opens selection window for the SKODA Master database. | `.xlsx`, `.xls`, or `.csv` containing `Master Sheet` sheet |
| **📂 Item Report** | Opens selection window for the internal customs report. | `.xlsx`, `.xls`, or `.csv` containing `Sheet1` sheet |
| **📂 Extracted Invoice** | Opens selection window for the parser-extracted invoice. | `.xlsx`, `.xls`, or `.csv` containing `Sheet1` sheet |
| **▶ Run** | Begins the comparison engine. Muted gray when locked, bright red when active. | Enabled only after all 3 files are successfully loaded |
| **💾 Export Excel** | Saves the modified mismatch list to an styled Excel spreadsheet. | Generates formatted `.xlsx` with bold values and red highlights |
| **🔄 Clear All** | Wipes the current session data, file paths, and clears the table. | Resets the GUI to its clean, initial startup state |
| **Data Preview Grid** | Renders mismatches side-by-side. Supports double-click inline edits. | Segoe UI 10 font, styled color tags based on `Status` |

---

## What is Compared?

The comparison engine executes a strict set of checks across all three files:

### 1. Structural Alignment
The tool groups all items by their Part Number and ensures the exact number of rows align 1-to-1 between the Item Report and the Extracted Invoice.
- If a part appears more times in one file than the other, it flags a **Duplicate Occurrence**.
- If a part is entirely absent from one file, it is flagged as missing.
- If a part is not found in the Master Sheet, it is flagged under **Master Reference**.

### 2. Master Sheet vs. Item Report
For every matched part, the tool validates customs reference data:
- **CTH Code:** Master `CTH1` vs. Item `CTH`
- **Basic Duty Rate**
- **IGST Notification SrNo**
- **Generic Description:** Master `Part Suffix` vs. Item `Generic Description`

### 3. Extracted Invoice vs. Item Report
The tool normalizes formats (stripping symbols, converting aliases like 'TW' to 'TAIWAN', and formatting dates) before comparing:
- **Description:** Normalized Invoice Description vs. Item `Product Desc`
- **Invoice Number**
- **Invoice Date**
- **Country of Origin**
- **Quantity**
- **Price / Amount:** Invoice `Total Price` vs. Item `Amount`
- **Unit:** Normalizes Invoice price suffixes (e.g., `/PC`, `/KG`) to match Item Report units (`PCS`, `KGS`).

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
