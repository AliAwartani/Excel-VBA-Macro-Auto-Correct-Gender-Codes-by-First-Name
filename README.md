# Gender Correction by First Name - Excel VBA Macro

## Overview
This Excel VBA macro automatically corrects gender codes in a dataset based on the majority gender associated with each first name. It is designed to clean and standardize gender data in large Excel files, while preserving original data and highlighting any corrections made.  

The macro intelligently detects relevant columns by their headers and handles exceptions such as titles or prefixes that should be ignored.

---

## Features
- **Automatic Column Detection:** Finds columns by header names (`FIRSTNAM`, `GENDER`, `TITLE`, `TITLECODE`) so the macro works on varying spreadsheet structures.
- **Destination Column:** Inserts a new "Corrected" gender column next to the original source column.
- **Gender Frequency Dictionary:** Builds a dictionary of first names and counts of each gender code.
- **Majority-Rule Correction:** Assigns gender based on the most frequent code for a first name.
- **Ignore List:** Skips rows with specified titles or prefixes (e.g., `Messrs.`, `Rev.`, `Captain.`) to avoid incorrect corrections.
- **Highlight Changes:** Highlights corrected cells in the source column for easy review.
- **Optimized Performance:** Temporarily disables screen updating, events, automatic calculation, and alerts for faster execution on large datasets.

---

## Requirements
- Microsoft Excel with VBA support.
- Dataset with headers (assumes first row contains headers).
- Gender codes in the dataset:  
  - `1` = Male  
  - `2` = Female  
  - `0` or blank = Unknown  

---

## Installation
1. Open your Excel workbook.
2. Press `ALT + F11` to open the VBA editor.
3. Insert a new module: `Insert > Module`.
4. Copy and paste the macro code into the module.
5. Close the VBA editor and save your workbook as a macro-enabled file (`.xlsm`).

---

## Usage
1. Open the workbook containing your data.
2. Ensure headers match the expected names (`FIRSTNAM`, `GENDER`, `TITLE`, `TITLECODE`) or modify the macro to match your headers.
3. Press `ALT + F8`, select `CorrectGendersByFirstName_AutoFindCols`, and click **Run**.
4. The macro will:  
   - Insert a new “Corrected” column next to the original gender column.  
   - Calculate gender frequencies by first name.  
   - Update the new column with corrected gender codes.  
   - Highlight any cells in the original column that were corrected.  
5. Review the changes and save the workbook.

---

## Customization
- **Headers:** Change the header names in the `FindColumnByHeader` calls if your dataset uses different labels.
- **Ignore List:** Update the `ignoreArr` array to add or remove titles and prefixes to skip.
- **Highlight Color:** Modify `RGB(255, 255, 0)` to change the highlight color for corrected cells.

---

## Error Handling
- If a required header is not found, the macro will show an error message and exit.
- Any runtime errors restore Excel settings and alert the user.

---

## Notes
- The macro assumes the first row contains headers and data starts at row 3.
- Corrections are based purely on majority gender per first name, which may not always reflect individual correctness.
- Always keep a backup of your original data before running macros.

---

## License
This code is provided as-is, without warranty. You are free to use and modify it for personal or commercial purposes.

