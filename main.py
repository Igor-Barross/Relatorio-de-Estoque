"""
Script de automação de relatório de estoque

Objetivo:
    Ler um relatório de estoque em Excel do sistema, filtrar informações relevantes
    (Produto, Estoque, Custo de Reposição) e gerar uma nova planilha com colunas
    adicionais para facilitar o controle físico do estoque, calcular diferenças e
    valores totais para reposição.

Funcionalidades:
    - Filtra produtos e remove linhas desnecessárias.
    - Remove códigos do produto, deixando apenas o nome.
    - Converte valores de texto em números no Excel.
    - Cria colunas adicionais: Estoque Físico, Diferença e Total.
    - Aplica fórmulas para calcular diferença entre estoque físico e do sistema,
      e o custo total de reposição.
    - Formata a planilha final com bordas, alinhamento e largura de colunas.

Entrada:
    - arquivo Excel do sistema: "relatorio.xls"

Saída:
    - nova planilha Excel formatada: "estoque horizonte.xlsx"

Bibliotecas utilizadas:
    - pandas
    - openpyxl
    - xlrd2
"""


import xlrd2
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font


def filtrar_coluna(df, coluna, filtro):
    """
    Essa função retira valores indesejados de uma coluna do DataFrame.
    'df' -> DataFrame original
    'coluna' -> Nome da coluna que será filtrada (string)
    'filtro' -> String ou regex a ser excluída
    """

    # Retira células vazias na coluna escolhida
    df = df[df[coluna].notna()]

    # Remove valores que contenham o filtro
    df = df[~df[coluna].astype(str).str.contains(filtro, na=False, case=False)]

    return df


def remover_codigo_produto(sheet, max_rows, coluna):
    """
    Remove o código inicial que aparece antes do nome do produto em um relatório Excel.
    
    Parâmetros:
    - sheet: objeto da planilha (openpyxl worksheet)
    - max_rows: número total de linhas da planilha
    - coluna: letra da coluna onde os produtos estão (ex: 'A')
    """

    for row in range(2, max_rows + 1):
        celula = sheet[f"{coluna}{row}"]

        if celula.value and isinstance(celula.value, str):
            partes = celula.value.split('-', 1)

            if len(partes) > 1:
                    celula.value = partes[1] 


def converter_para_float(sheet, max_rows, coluna):
    """
    Converte valores do tipo string para float em uma coluna específica da planilha.
    Substitui vírgula por ponto e aplica formato numérico.
    """
    
    for row in range(2, max_rows + 1):
        celula = sheet[f"{coluna}{row}"]

        if isinstance(celula.value, str):
            sheet[f"{coluna}{row}"].value = float(celula.value.replace(",", "."))
            # sheet[f"{coluna}{row}"].number_format = '0,0'


def converter_para_int_ou_limpar(sheet, max_rows, coluna):
    """
    Converte strings para inteiros em uma coluna, limpa valores inválidos 
    e aplica formato numérico.
    """
    
    for row in range(2, max_rows + 1):
        celula = sheet[f"{coluna}{row}"]

        if isinstance(celula.value, str):
            try:
                # tenta converter o texto em número
                numero = int(celula.value)
                celula.value = numero
            except:
                # se não der pra converter, deixa em branco
                celula.value = None
        
        if celula.value is not None:
            # deixa a célula vazia mas já com formato numérico
            celula.number_format = 'General'


def aplicar_formato_moeda(sheet, max_rows, coluna):
    for row in range(2, max_rows + 1):
        sheet[f"{coluna}{row}"].number_format = '"R$"* #,##0.00_-'


df = pd.read_excel(
    "relatorio.xls", 
    engine="xlrd",
    skiprows=4
)

# Filtros para linhas e palavras que não vão para o relatório final

# Filtro para Coluna PRODUTO
df = filtrar_coluna(df, "Produto", "Produto")
df = filtrar_coluna(df, "Produto", "Total geral")
df = filtrar_coluna(df, "Produto", "Filtro:")
df = filtrar_coluna(df, "Produto", "FELIANA ALIMENTOS LTDA")

# Filtro para Coluna ESTOQUE
df = filtrar_coluna(df, "Estoque", "Estoque")

# Filtro para Coluna CST. REP (custo de reposição)
df = filtrar_coluna(df, "Cst. Rep.", "Cst. Rep.")

# Colunas principais do relatório do sistema
df = df[["Produto", "Estoque", "Cst. Rep."]]

# Criação das colunas para nova tabela
df.insert(1, "Estoque Físico", " ")
df.insert(3, "Diferença", " ")
df.insert(5, "Total", " ")

arquivo_saida = "Estoque horizonte.xlsx"
df.to_excel(arquivo_saida, index=False)

workbook = load_workbook(arquivo_saida)
sheet = workbook.active
max_rows = sheet.max_row
max_cols = sheet.max_column

# Remove códigos que aparecem antes do nome do produto na coluna especificada
# Mantém apenas o nome do produto

remover_codigo_produto(sheet, max_rows, 'A')

# Converte valores da coluna de texto para float
# Troca vírgula por ponto e aplica formatação numérica

converter_para_float(sheet, max_rows, 'C') # COLUNA ESTOQUE SISTEMA
converter_para_float(sheet, max_rows, 'E') # COLUNA CST. REP

# Converte valores da coluna para inteiro quando possível
# Se não puder converter, limpa a célula
# Garante que a coluna fique pronta para cálculos no Excel

converter_para_int_ou_limpar(sheet, max_rows, 'B') # COLUNA ESTOQUE

# Aplica formatação de moeda para colunas de valor

aplicar_formato_moeda(sheet, max_rows, 'E')
aplicar_formato_moeda(sheet, max_rows, 'F')

# Cria fórmulas no Excel para:

# 1. Diferença entre estoque físico e estoque do sistema
for row in range(2, max_rows + 1):
    sheet[f'D{row}'] = f"=B{row} - C{row}"

# 2. Valor total de reposição baseado na diferença e custo unitário
for row in range(2, max_rows + 1):
    sheet[f'F{row}'] = f"=D{row} * E{row}"


# Estilização de toda planilha

borda_fina = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Aplicação das bordas em toda planilha
for row in sheet.iter_rows(min_row=1, max_row=max_rows, min_col=1, max_col=max_cols):
    for celula in row:
        celula.border = borda_fina

# Redimensionamento das Colunas
sheet.column_dimensions['A'].width = 55 # COLUNA PRODUTO
sheet.column_dimensions['B'].width = 15 # COLUNA ESTOQUE FISICO
sheet.column_dimensions['C'].width = 13 # COLUNA ESTOQUE SISTEMA
sheet.column_dimensions['D'].width = 10 # COLUNA DIFERENÇA
sheet.column_dimensions['E'].width = 13 # COLUNA CST. REP
sheet.column_dimensions['F'].width = 13 # COLUNA TOTAL

# Redimensionamento do cabeçario
for row in range(1, max_rows + 1):
    sheet.row_dimensions[row].height = 18

# Alinhamento horizontal e vertical de todas as linhas do cabeçario
align_center = Alignment(horizontal="center", vertical="center")

for celula in sheet[1]:
    celula.alignment = align_center
    celula.font = Font(bold=True)

# Alinhamento para valores das colunas
for row in range(2, max_rows + 1):
    sheet[f'A{row}'].alignment = Alignment(horizontal="left", vertical="center") # COLUNA PRODUTO
    sheet[f'B{row}'].alignment = align_center # COLUNA ESTOQUE FISICO
    sheet[f'C{row}'].alignment = align_center # COLUNA ESTOQUE SISTEMA
    sheet[f'D{row}'].alignment = align_center # COLUNA DIFERENÇA
    sheet[f'E{row}'].alignment = Alignment(horizontal="right", vertical="center") # COLUNA CST. REP
    sheet[f'F{row}'].alignment = Alignment(horizontal="right", vertical="center") # COLUNA TOTAL

workbook.save(arquivo_saida)
print('Relatório gerado com sucesso!')