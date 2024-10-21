import xlwings as xw
import pandas as pd

# Função para garantir que as fórmulas comecem com '='
def ensure_formula(formula):
    return '=' + formula if formula and not formula.startswith('=') else formula

# Função geral para comparar fórmulas entre duas planilhas
def compare_formulas(ws_matriz, ws_check, cell_pairs, differences, sheet_name):
    for ref_matriz, ref_check in cell_pairs:
        # Obter e garantir que a fórmula tenha o sinal '='
        formula_matriz = ensure_formula(ws_matriz.range(ref_matriz).formula)
        formula_check = ensure_formula(ws_check.range(ref_check).formula)

        # Comparar fórmulas e registrar diferenças, se houver
        if formula_matriz != formula_check:
            differences.append({
                'Sheet': sheet_name,
                'Matriz Cell': ref_matriz,
                'Matriz Formula': formula_matriz,
                'Check Cell': ref_check,
                'Check Formula': formula_check
            })

# Função para aplicar a comparação de fórmulas para ambos mapeamentos de colunas e linhas de orçamento
def compare_formulas_in_sheet(ws_matriz, wb_check, sheet_name, column_mapping, row_matriz, row_check, differences):
    print(f"Comparando {sheet_name}")
    ws_check = wb_check.sheets[sheet_name]

    # Definir pares de células para comparação de fórmulas de orçamento
    budget_cell_pairs = [
        (f'A{row_matriz}', 'B2'),  # Orçamento Total
        (f'B{row_matriz}', 'C2'),  # Orçamento Usado
        (f'C{row_matriz}', 'D2')   # Orçamento Restante
    ]
    # Comparar fórmulas de orçamento
    compare_formulas(ws_matriz, ws_check, budget_cell_pairs, differences, sheet_name)

    # Criar pares de células para comparação de mapeamento de colunas
    column_cell_pairs = [(f'{col_matriz}{row_matriz}', f'{col_check}{row_check}') 
                         for col_matriz, col_check in column_mapping.items()]
    # Comparar fórmulas de colunas
    compare_formulas(ws_matriz, ws_check, column_cell_pairs, differences, sheet_name)

# Função principal para realizar a comparação e salvar o relatório
def main():
    # Carregar arquivos fictícios
    wb_matriz = xw.Book('C:\\Users\\your_user\\Desktop\\Formula Checks\\Matriz Ficticia.xlsx')
    ws_matriz = wb_matriz.sheets['Planilha1']

    wb_check = xw.Book('C:\\Users\\your_user\\Desktop\\Formula Checks\\Verificacao Ficticia.xlsx')

    # Mapeamento de Colunas: [Colunas da Matriz Ficticia , Colunas da Verificacao Ficticia]
    column_mapping = {
        'A': 'B',  # 'Total Orçamento'
        'B': 'C',  # 'Orçamento Usado'
        'C': 'D',  # 'Orçamento Restante'
        'D': 'E',  # 'Churn Rate'
        'E': 'F',  # 'Taxa de Crescimento'
    }

    # Lista para armazenar diferenças
    differences = []

    # Aplicar comparação a todas as abas
    sheets_to_compare = {
        'ABAS 1': 3, 'ABAS 2': 4, 'ABAS 3': 5
    }

    for sheet_name, row_matriz in sheets_to_compare.items():
        compare_formulas_in_sheet(ws_matriz, wb_check, sheet_name, column_mapping, row_matriz, row_check=2, differences=differences)

    # Criar um DataFrame com as diferenças
    df_differences = pd.DataFrame(differences)
    
    # Salvar o relatório em um arquivo Excel separado
    output_file = 'C:\\Users\\your_user\\Desktop\\Formula Checks\\comparacao_formulas.xlsx'
    df_differences.to_excel(output_file, index=False)

    print(f"Relatório gerado com sucesso em: {output_file}")

# Executando a função principal
if __name__ == "__main__":
    main()
