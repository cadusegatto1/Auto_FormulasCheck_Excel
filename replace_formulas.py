## Created by Cadu Segatto
# This script automates the transfer of formulas between two Excel spreadsheets

import xlwings as xw

# Load files
# Change the file paths to your actual files
wb_matriz = xw.Book('C:\\Users\\your_user\\Desktop\\Formula Checks\\Matriz_Ficticia.xlsx')  # Change the path and filename here
ws_matriz = wb_matriz.sheets['Planilha_Matriz']  # Change the sheet name if needed

# Opening the template file
wb_namer = xw.Book('C:\\Users\\your_user\\Desktop\\Formula Checks\\Modelo_Ficticio.xlsx')  # Change the path and filename here

def apply_formulas(ws_matriz, ws_namer, column_mapping, row_matriz, row_namer):
    # Iterate through each mapped column
    for col_matriz, col_namer in column_mapping.items():
        # Construct the cell references
        ref_matriz = f'{col_matriz}{row_matriz}'
        ref_namer = f'{col_namer}{row_namer}'

        # Collects the formula from the source cell
        formula = ws_matriz.range(ref_matriz).formula

        # Ensure formula starts with '='
        if not formula.startswith("="):
            formula = "=" + formula
        
        # Apply formula to target cell
        ws_namer.range(ref_namer).formula = formula

def apply_budget_formulas(ws_matriz, ws_namer, row_source):
    # Original formulas from matriz
    formula_total_budget = ws_matriz.range(f'C{row_source}').formula  # Formula to Total Budget
    formula_budget_used = ws_matriz.range(f'D{row_source}').formula   # Formula to Budget Used
    formula_remaining_budget = ws_matriz.range(f'E{row_source}').formula  # Formula to Remaining Budget

    # Apply formulas to the target cells
    ws_namer.range('D3').formula = '=' + formula_total_budget
    ws_namer.range('E3').formula = '=' + formula_budget_used
    ws_namer.range('F3').formula = '=' + formula_remaining_budget

def apply_formulas_to_sheet(ws_matriz, wb_namer, sheet_name, row_budget, column_mapping, row_matriz, row_namer):
    print(f"Running {sheet_name}")
    ws_namer = wb_namer.sheets[sheet_name]
    
    # Apply budget formulas based on the formulas' source line
    apply_budget_formulas(ws_matriz, ws_namer, row_budget)
    
    # Apply column mapping
    apply_formulas(ws_matriz, ws_namer, column_mapping, row_matriz, row_namer)

# Column Mapping: [Source Column in Matriz, Target Column in Fictitious Calculator]
# Change these mappings according to your Excel file structure
def main():
    column_mapping = {
        'F': 'G',  # 'Fictitious Metric 1' - change as needed
        'G': 'H',  # 'Fictitious Metric 2' - change as needed
        'H': 'I',  # 'Fictitious Metric 3' - change as needed
        'I': 'J',  # 'Fictitious Metric 4' - change as needed
        'K': 'L',  # 'Fictitious Metric 5' - change as needed
    }

    # Applying formulas to all fictitious sheets
    apply_formulas_to_sheet(ws_matriz, wb_namer, 'FICTICIO 1', row_budget=3, column_mapping=column_mapping, row_matriz=3, row_namer=8)
    apply_formulas_to_sheet(ws_matriz, wb_namer, 'FICTICIO 2', row_budget=4, column_mapping=column_mapping, row_matriz=4, row_namer=8)
    apply_formulas_to_sheet(ws_matriz, wb_namer, 'FICTICIO 3', row_budget=5, column_mapping=column_mapping, row_matriz=5, row_namer=8)
    apply_formulas_to_sheet(ws_matriz, wb_namer, 'FICTICIO 4', row_budget=6, column_mapping=column_mapping, row_matriz=6, row_namer=8)
    apply_formulas_to_sheet(ws_matriz, wb_namer, 'FICTICIO 5', row_budget=7, column_mapping=column_mapping, row_matriz=7, row_namer=8)

    print("Fórmulas aplicadas com sucesso!")

if __name__ == "__main__":
    main()