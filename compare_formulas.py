import xlwings as xw
import pandas as pd

# Function to ensure formulas start with '='
def ensure_formula(formula):
    return '=' + formula if formula and not formula.startswith('=') else formula

# General function to compare formulas between two worksheets
def compare_formulas(ws_matriz, ws_check, cell_pairs, differences, sheet_name):
    for ref_matriz, ref_check in cell_pairs:
        # Get the formula and ensure it has the '=' sign
        formula_matriz = ensure_formula(ws_matriz.range(ref_matriz).formula)
        formula_check = ensure_formula(ws_check.range(ref_check).formula)

        # Compare formulas and log differences if they exist
        if formula_matriz != formula_check:
            differences.append({
                'Sheet': sheet_name,
                'Matriz Cell': ref_matriz,
                'Matriz Formula': formula_matriz,
                'Check Cell': ref_check,
                'Check Formula': formula_check
            })

# Function to apply formula comparison for both budget and column mapping
def compare_formulas_in_sheet(ws_matriz, wb_check, sheet_name, column_mapping, row_matriz, row_check, differences):
    print(f"Comparando {sheet_name}")
    ws_check = wb_check.sheets[sheet_name]

    # Define cell pairs for budget formula comparison
    budget_cell_pairs = [
        (f'A{row_matriz}', 'B2'),  # Total Budget
        (f'B{row_matriz}', 'C2'),  # Budget Used
        (f'C{row_matriz}', 'D2')   # Remaining Budget
    ]
    # Compare budget formulas
    compare_formulas(ws_matriz, ws_check, budget_cell_pairs, differences, sheet_name)

    # Create cell pairs for column mapping comparison
    column_cell_pairs = [(f'{col_matriz}{row_matriz}', f'{col_check}{row_check}') 
                         for col_matriz, col_check in column_mapping.items()]
    # Compare column formulas
    compare_formulas(ws_matriz, ws_check, column_cell_pairs, differences, sheet_name)

# Main function to perform the comparison and save the report
def main():
    # Load fictional files
    wb_matriz = xw.Book('C:\\Users\\your_user\\Desktop\\Formula Checks\\Matriz_Ficticia.xlsx')  # Change the path and filename here
    ws_matriz = wb_matriz.sheets['Planilha1']  # Change the sheet name if needed

    wb_check = xw.Book('C:\\Users\\your_user\\Desktop\\Formula Checks\\Verificacao_Ficticia.xlsx')  # Change the path and filename here

    # Column Mapping: [Fictitious Matrix Columns, Fictitious Check Columns]
    column_mapping = {
        'A': 'B',  # 'Total Budget' - change as needed
        'B': 'C',  # 'Budget Used' - change as needed
        'C': 'D',  # 'Remaining Budget' - change as needed
        'D': 'E',  # 'Churn Rate' - change as needed
        'E': 'F',  # 'Growth Rate' - change as needed
    }

    # List to store differences
    differences = []

    # Apply comparison to all specified sheets
    sheets_to_compare = {
        'ABAS 1': 3, 
        'ABAS 2': 4, 
        'ABAS 3': 5
    }

    for sheet_name, row_matriz in sheets_to_compare.items():
        compare_formulas_in_sheet(ws_matriz, wb_check, sheet_name, column_mapping, row_matriz, row_check=2, differences=differences)

    # Create a DataFrame with the differences
    df_differences = pd.DataFrame(differences)
    
    # Save the report to a separate Excel file
    output_file = 'C:\\Users\\your_user\\Desktop\\Formula Checks\\comparacao_formulas.xlsx'  # Change the path and filename here
    df_differences.to_excel(output_file, index=False)

    print(f"Relatório gerado com sucesso em: {output_file}")

# Executing the main function
if __name__ == "__main__":
    main()
