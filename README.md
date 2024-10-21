## **Script for Comparing Formulas Between Spreadsheets**

### **General Description**
This script was created by Cadu Segatto and is used to compare formulas between two Excel spreadsheets: a matrix spreadsheet (with reference data) and an CHECK spreadsheet (with data to be verified). The script identifies discrepancies between the corresponding formulas and generates a detailed report in an Excel file, listing the differences found.

### **How ​​it works**

1. **Loading Spreadsheets**:

The script loads two spreadsheets:

- **Matrix**: Reference spreadsheet containing the expected formulas.

- **Check**: Spreadsheet where the formulas are verified against the Matrix.

2. **Column Mapping**:

The script uses a mapping dictionary that defines which columns of one spreadsheet correspond to the columns of the other. This ensures that the correct formulas are compared, even if the columns in the two sheets are in different positions.

3. **Formula Comparison**:

For each mapped column, the script compares the formulas between the two sheets. If the formula in the Matrix does not have the `=` sign at the beginning, it is automatically added to standardize the comparison.

4. **Difference Report**:

If differences are found in the formulas between the Matrix and Check sheets, the discrepancies are stored in a list. The report is then generated and exported to a separate Excel file containing:
- The tab (sheet) where the difference was found.
- The cell and formula from the Matrix sheet.
- The corresponding cell and formula from the Check sheet.

### **Functions**

#### 1. **compare_formulas()**
- **Parameters**:
- `ws_matriz`: Tab of the Matrix sheet.
- `ws_check`: Check spreadsheet tab.
- `column_mapping`: Dictionary with the mapping of columns between the two spreadsheets.
- `row_matriz`: Number of the initial row in the Matrix spreadsheet.
- `row_check`: Number of the initial row in the Check spreadsheet.
- `differences`: List where the discrepancies will be stored.
- `sheet_name`: Name of the tab where the comparison is being made.
- **Description**: Compares the formulas between two mapped columns from different spreadsheets and, if there are differences, stores them in the `differences` list.

#### 2. **compare_formulas_in_sheet()**
- **Parameters**:
- `ws_matriz`: Tab of the Matrix spreadsheet.
- `wb_check`: Object representing the Check spreadsheet.
- `sheet_name`: Name of the tab where the comparison is being made.
- `row_budget`: Budget line number in the Matrix.
- `column_mapping`: Dictionary with the column mapping.
- `row_matriz`: Starting line for verification in the Matrix.
- `row_check`: Starting line for verification in Check.
- `differences`: List where the discrepancies will be stored.
- **Description**: Performs the comparison of formulas for a specific tab, using the mapping of defined columns and rows.

#### 3. **main()**
- **Description**: Main function that executes the script. It loads the files, applies the comparison of formulas in the specified tabs and generates a final report of the discrepancies in an Excel file.

### **Steps to Run the Script**

1. Make sure the required libraries are installed:
- `xlwings`
- `pandas`

2. Open the script and adjust the paths of the Excel files you want to compare:
- **Matrix**: `Source_Formulas.xlsx`
- **Check**: `Check.xlsx`

3. Run the script. It will loop through the mapped columns in the tabs contained in the sheets.
-	If something was changed, you will need to check if the tabs name is the same.

4. After execution, the difference report will be saved in the following location:
- `C:\\Users\\your_user\\Desktop\\Formula Checks\\comparison_formulas.xlsx`

### **Script Output**

- **Excel Report**: An Excel file will be generated with the differences found, including the tab, cell, and formulas that differ from the two spreadsheets.

---

### **Notes**

- Make sure that the spreadsheets are in the correct format and that the column mapping is up to date, as columns may vary between different versions of the spreadsheets. - The script currently ignores cells with free text (not containing formulas) that are mapped in the column mapping dictionary.
