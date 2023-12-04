
import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(743,225,duration=1)
    pyautogui.hotkey('command','v')

    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(746,294,duration=1)
    pyautogui.hotkey('command','v')
    
    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(749,403,duration=1)
    pyautogui.hotkey('command','v')

    # Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(748,472,duration=1)
    pyautogui.hotkey('command','v')

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(748,541,duration=1)
    pyautogui.hotkey('command','v')

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(748,613,duration=1)
    pyautogui.hotkey('command','v')

    # Botão próximo
    pyautogui.click(749,661,duration=1)
    sleep(3)

    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(747,247,duration=1)
    pyautogui.hotkey('command','v')

    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(743,316,duration=1)
    pyautogui.hotkey('command','v')

    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(743,385,duration=1)
    pyautogui.hotkey('command','v')
    
    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(744,451,duration=1)
    pyautogui.hotkey('command','v')

    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(747,524,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(748,524,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(748,541,duration=1)
    else:
        pyautogui.click(743,558,duration=1)

    # material    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(746,589,duration=1)
    pyautogui.hotkey('command','v')
    # Botão próximo
    pyautogui.click(763,642,duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(749,278,duration=1)
    pyautogui.hotkey('command','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(748,347,duration=1)
    pyautogui.hotkey('command','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(745,418,duration=1)
    pyautogui.hotkey('command','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(748,520,duration=1)
    pyautogui.hotkey('command','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(750,588,duration=1)
    pyautogui.hotkey('command','v')

    # Botão concluir
    pyautogui.click(750,641,duration=1)
    # Botão confirmar inclusão
    pyautogui.click(1179,212,duration=1)
    # Botão confirmação 2
    #pyautogui.click(1886,191,duration=1)
    # iniciar cadastro novamente
    pyautogui.click(984,450,duration=1)

# Repetir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# clicar em ok, para finalizar o processo
# Clicar no ok mais um vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"
