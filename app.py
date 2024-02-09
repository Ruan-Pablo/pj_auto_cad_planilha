import pyautogui 
import openpyxl
import pyperclip
from time import sleep

# entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informações de um campo e colar no correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_prod = linha[0].value 
    pyperclip.copy(nome_prod)
    pyautogui.click(1347,317,duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1241,435, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1182,595, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    cod_prod = linha[3].value
    pyperclip.copy(cod_prod)
    pyautogui.click(1215,690, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1231,802, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1252,902, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    # PRÓXIMA PÁGINA
    pyautogui.click(1173,980, duration=1)
    sleep(2)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1261,339, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    qtd = linha[7].value
    pyperclip.copy(qtd)
    pyautogui.click(1186,446, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(1260,561, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1262,670, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    tamanho  = linha[10].value
    pyautogui.click(1297,788,duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(1297,818,duration=1)
    elif tamanho == "Médio":
        pyautogui.click(1297,843,duration=1)
    else:
        pyautogui.click(1297,867,duration=1)
    material= linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1297,880, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    
    # proxima página
    pyautogui.click(1172,953,duration=1)
    sleep(2)

    fabricante  = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1210,364, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1211,471, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    obs = linha[14].value
    pyperclip.copy(obs)
    pyautogui.click(1216,601, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    cod_barras = linha[15].value
    pyperclip.copy(cod_barras)
    pyautogui.click(1211,743, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')
    local_armazem = linha[16].value
    pyperclip.copy(local_armazem)
    pyautogui.click(1215,862, duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('esc')

    # concluir
    pyautogui.click(1190,927, duration=1)
    # confimar
    pyautogui.click(1661,203,duration=1)
    sleep(2)
    # add mais um 
    pyautogui.click(1428,646,duration=1)