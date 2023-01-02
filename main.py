from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


# from pandas  as pd

# Cria Novo workbook
wb = Workbook()

# Seleciona a active Sheet
ws1 = wb.active

# Rename it

ws1.title = 'my test'

# Escreve alguns dados

for col in range(1,5):
    for row in range(1,6):
        letter = get_column_letter(col)
        ws1[letter + str(row)] = letter + str(row)





# Cria nova sheet

ws2 = wb.create_sheet(title="Ok")

ws2["C1"] = "OK"

# Salva arquivo (Se não colocar o caminho complete, ele salva
# na mesma pasta do scritp.

wb.save('Text.xlsx')






# Acesse Sheet 1
ws = wb['Sheet1']
# Acesse Sheet 2
ws = wb['Sheet2']
# Use a aba ativa quando o arquivo foi carregado
ws = wb.active


# Adiciona valor a uma célula
ws['A1'] = "Any Value"
# Adiciona fórmula a uma célula
ws['C7'] = '=SUM(C2:C6)'
# Adiciona linha no final da tabela (append)
# Use uma lista onde cada valor vai para uma célula da linha
to_append = ['A1', 2, 'C1', 4, 'end']
#Append
ws.append(to_append)
#Salve
wb.save('Test.xlsx')


