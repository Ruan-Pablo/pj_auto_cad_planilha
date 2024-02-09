import pyautogui 
import openpyxl

# entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informações de um campo e colar no correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    print(linha[0].value)