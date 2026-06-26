# SKODA Item Mismatch Checker: Comparison Reference List

This document provides a complete list of all the comparisons the tool makes, including the exact Mismatch Field it triggers and the Remark it writes for each scenario.

### 1. Structural & Row Count Mismatches
These mismatches occur when comparing the physical rows between the three files to make sure everything lines up 1-to-1.

| What is being Compared | Mismatch Field | Exact Remark Written |
| :--- | :--- | :--- |
| **Missing in all sources** <br>*(In Item Report, but missing in both Invoice and Master)* | `Model` | `"Model found in Item Report but missing in both Extracted Invoice and Master Sheet"` |
| **Orphaned in Item Report** <br>*(Part exists in Invoice, but Item Report has more line-items for it)* | `Duplicate Occurrence` | `"Item Report contains {X} row(s) for this Model, but Extracted Invoice contains only {Y} row(s)."` |
| **Missing in Invoice** <br>*(Part is in Item Report, but completely absent from Invoice)* | `Duplicate Occurrence` | `"Model found in Item Report but missing in Extracted Invoice"` |
| **Orphaned in Invoice** <br>*(Part exists in Item Report, but Invoice has more line-items for it)* | `Duplicate Occurrence` | `"Extracted Invoice contains {Y} row(s) for this Model, but Item Report contains only {X} row(s)."` |
| **Missing in Item Report** <br>*(Part is in Invoice, but completely absent from Item Report)* | `Duplicate Occurrence` | `"Model found in Extracted Invoice but missing in Item Report"` |
| **Missing in Master Sheet** <br>*(Part exists in Invoice/Item Report, but is not in Master)* | `Master Reference` | `"Model missing in Master Sheet"` |
| **Conflicting Master Data** <br>*(Master Sheet has the same part listed twice with different Duty/CTH rates)* | `PartNo` | `"Duplicate Part Number in Master Sheet with conflicting reference values"` |

---

### 2. Master Sheet vs. Item Report Mismatches
Once the rows are aligned, the tool cross-references the customs data between the Master Sheet and the Item Report.

| What is being Compared | Mismatch Field | Exact Remark Written |
| :--- | :--- | :--- |
| **CTH Code** <br>*(Master `CTH1` vs. Item `CTH`)* | `CTH` | `"CTH mismatch between Master Sheet and Item Report"` |
| **Basic Duty Rate** <br>*(Master `Basic Duty Rate` vs. Item `Basic Duty Rate`)* | `Basic Duty Rate` | `"Basic Duty Rate mismatch between Master Sheet & Item Report"` |
| **IGST Notification** <br>*(Master `IGST Notn Sr No` vs. Item `IGST Notification SrNo`)* | `IGST Notification SrNo` | `"IGST Notification SrNo mismatch between Master Sheet & Item Report"` |
| **Generic Description** <br>*(Master `Part Suffix` vs. Item `Generic Description`)* | `Generic Description` | `"Gen. Desc. does not match in Item Report & Master Sheet."` |

---

### 3. Extracted Invoice vs. Item Report Mismatches
The tool then compares the physical invoice data against the Item Report data (applying intelligent normalization for dates, countries, and prices).

| What is being Compared | Mismatch Field | Exact Remark Written |
| :--- | :--- | :--- |
| **Part Description** <br>*(Invoice `Description` vs. Item `Product Desc`)* | `Description` | `"Description mismatch"` |
| **Invoice Number** <br>*(Invoice `Invoice Number` vs. Item `Invoice No`)* | `Invoice Number` | `"Invoice Number mismatch"` |
| **Invoice Date** <br>*(Invoice `Invoice Date` vs. Item `Invoice Date`)* | `Invoice Date` | `"Invoice Date mismatch"` |
| **Country of Origin** <br>*(Invoice `Country of Origin` vs. Item `Country of Origin`)* | `Country of Origin` | `"Country of Origin mismatch"` |
| **Quantity** <br>*(Invoice `Quantity` vs. Item `Quantity`)* | `Quantity` | `"Quantity mismatch"` |
| **Price / Amount** <br>*(Invoice `Total Price` vs. Item `Amount`)* | `Amount` | `"Amount & Total Price mismatch"` |
| **Currency** <br>*(Invoice `Currency` vs. Item `Amount`)* | `Currency` | `"Currency does not match in Item Report & Invoice"` |
| **Unit** <br>*(Invoice `Price per PC` suffix vs. Item `Unit`)* | `Unit` | `"Unit does not match in Item Report & Invoice"` |

---

### 4. Special Information Rows
These are not errors, but are captured by the tool for auditing and UI purposes.

| Context | Mismatch Field | Exact Remark Written |
| :--- | :--- | :--- |
| **Targeted CTH Flagged** <br>*(Item has CTH starting with 74, 76, 40111010, 85261000, 9401)* | `CTH Alert` | `"CTH Code Detected"` *(Highlighted in bold font)* |
| **Zero Mismatches Found** <br>*(Outputs a success row for every verified Job Number)* | `Success` | `"Successfully verified. No mismatches found."` *(Highlighted in white)* |
