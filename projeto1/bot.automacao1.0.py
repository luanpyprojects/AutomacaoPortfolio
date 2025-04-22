#Automação

import pandas as pd
import time
import pyautogui
import unicodedata

# Passo 1 - Importar dados
tabela = pd.read_excel('produtos_ficticios.xlsx')

# Passo 2 - Limpar dados
def limpar_texto(texto):
    texto = unicodedata.normalize('NFD', texto)  # Normaliza a string
    texto = texto.encode('ascii', 'ignore').decode('utf-8')  # Remove caracteres especiais
    texto = texto.replace('ç', 'c').replace('Ç', 'C')  # Substitui ç e Ç por c e C
    return texto

# Aplicar a limpeza em todas as células do DataFrame
for coluna in tabela.columns:
    tabela[coluna] = tabela[coluna].apply(lambda x: limpar_texto(str(x)) if isinstance(x, str) else x)

# Exibir a tabela limpa
print(tabela)

# Passo 3 - Configurar pausa do PyAutoGUI
pyautogui.PAUSE = 0.25  # Define o tempo de pausa entre cada ação do pyautogui


#Passo 2- Entrar no chrome
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
time.sleep(0.5)


#Passo3- Entrar no site de cadastro
pyautogui.write('https://cadastro-produtos-devaprender.netlify.app/')
pyautogui.press('enter')

#Passo4- Estabelecer sequência

for linha in tabela.index:
    time.sleep(1.5)
    pyautogui.press('tab')

    #Nome
    nome=str(tabela.loc[linha,'Nome do produto'])
    pyautogui.write(nome)
    pyautogui.press('tab')

    #Descrição
    descricao=str(tabela.loc[linha,'Descrição'])
    pyautogui.write(descricao)
    pyautogui.press('tab')

    #Categoria
    categoria=str(tabela.loc[linha,'Categoria'])
    pyautogui.write(categoria)
    pyautogui.press('tab')
    
    #Codigo
    codigo=str(tabela.loc[linha,'Código do produto'])
    pyautogui.write(codigo)
    pyautogui.press('tab')

    #Peso
    peso=str(tabela.loc[linha,'Peso (em kg)'])
    pyautogui.write(peso)
    pyautogui.press('tab')
    
    #Dimensoes
    dimensoes=str(tabela.loc[linha,'Dimensões (L x A x P)'])
    pyautogui.write(dimensoes)
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(1.5)

    #Preço
    pyautogui.click(x=319, y=345)
    preco=str(tabela.loc[linha,'Preço'])
    pyautogui.write(preco)
    pyautogui.press('tab')

    #Quantidade em estoque
    qtd=str(tabela.loc[linha,'Quantidade em estoque'])
    pyautogui.write(qtd)
    pyautogui.press('tab')

    #Data de validade
    validade=str(tabela.loc[linha,'Data de validade'])
    pyautogui.write(validade)
    pyautogui.press('tab')

    #Cor
    cor=str(tabela.loc[linha,'Cor'])
    pyautogui.write(cor)
    pyautogui.press('tab')

    #Tamanho
    tamanho=str(tabela.loc[linha,'Tamanho'])
    pyautogui.write(tamanho)
    pyautogui.press('tab')

    #Material
    material=str(tabela.loc[linha,'Material'])
    pyautogui.write(material)
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(1.5)

    #Fabricante
    pyautogui.click(x=319, y=345)
    fabricante=str(tabela.loc[linha,'Fabricante'])
    pyautogui.write(fabricante)
    pyautogui.press('tab')

    #País de origem
    origem=str(tabela.loc[linha,'País de origem'])
    pyautogui.write(origem)
    pyautogui.press('tab')

    #Observações
    obs=str(tabela.loc[linha,'Observações'])
    pyautogui.write(obs)
    pyautogui.press('tab')

    #Código de barras
    codigo2=str(tabela.loc[linha,'Código de barras'])
    pyautogui.write(codigo2)
    pyautogui.press('tab')

    #Localização no armazém 
    loc=str(tabela.loc[linha,'Localização no armazém'])
    pyautogui.write(loc)
    pyautogui.press('tab')
    pyautogui.press('enter')

    #restart
    time.sleep(1.5)
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.press('tab')
    pyautogui.press('enter')

      

