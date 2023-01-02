
from openpyxl import Workbook, load_workbook



planilha = load_workbook("Text.xlsx")

aba_ativa = planilha.active
# I
for celula in aba_ativa["E"]:
     if celula.value == 35:
        linha = celula.row
        aba_ativa[f"E{linha}"] = " "

planilha.save("Text.xlsx")



#
#
# sheet_obj = planilha.active
#
# row = sheet_obj.max_row
# column = sheet_obj.max_column
#
# print("Total Rows:", row)
# print("Total Columns:", column)
#
# # SELECIONADO PRIMEIRA COLUNA
# for i in range(1, row + 1):
#     cell_obj = sheet_obj.cell(row = i, column = 1)
#     print(cell_obj.value)
#
# # SELECIONANDO PRIMEIRA LINHA
# for i in range(1, column + 1):
#     cell_obj = sheet_obj.cell(row = 1, column = i)
#     print(cell_obj.value, end = " | ")













