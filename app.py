import openpyxl
import pyperclip #mantém os acentos das palavras
import pyautogui
import time #preenche os dados 
# entrar na planillha

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx') #esse comando abre a planilha
sheet_produtos = workbook['Produtos'] #aqui seleciona a página desejada, no caso desse projeto a produtos
# copiar informações de um campo e colar num campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2): #esse for acessa as linhas 
  #Nome produto
  nome_produto = linha[0].value
  pyperclip.copy(nome_produto) 
  pyautogui.click(846,236,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Descrição
  descricao = linha[1].value 
  pyperclip.copy(descricao)
  pyautogui.click(823,307,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Categoria
  categoria = linha[2].value 
  pyperclip.copy(categoria)
  pyautogui.click(827,427,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Codigo produto
  codigo_produto = linha[3].value 
  pyperclip.copy(codigo_produto)
  pyautogui.click(822,503,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Peso
  peso = linha[4].value 
  pyperclip.copy(peso)
  pyautogui.click(819,580,duration=1)
  pyautogui.hotkey('ctrl','v')

  #Dimensões
  dimensoes = linha[5].value 
  pyperclip.copy(dimensoes)
  pyautogui.click(824,666,duration=1)
  pyautogui.hotkey('ctrl','v')
  
  #Botão próximo
  pyautogui.click(852,707,duration=1)
  time.sleep(3)

  #Preco
  preco = linha[6].value 
  pyperclip.copy(preco)
  pyautogui.click(824,251,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Quantidade
  quantidade = linha[7].value
  pyperclip.copy(quantidade)
  pyautogui.click(827,327,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Validade
  validade = linha[8].value
  pyperclip.copy(validade)
  pyautogui.click(821,404,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Cor
  cor = linha[9].value
  pyperclip.copy(cor)
  pyautogui.click(824,484,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Tamanho
  tamanho = linha[10].value
  pyautogui.click(825,554,duration=1)
  if tamanho == 'Pequeno':
    pyautogui.click(866,594,duration=1)
  elif tamanho =='Médio':
    pyautogui.click(859,612,duration=1)
  else:
    pyautogui.click(867,649,duration=1)

  #Material
  material = linha[11].value
  pyperclip.copy(material)
  pyautogui.click(822,640,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Botão proximo
  pyautogui.click(858,693,duration=1)
  time.sleep(3)

  #Fabricante
  fabricante = linha[12].value
  pyperclip.copy(fabricante)
  pyautogui.click(823,269,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Pais
  pais_origem = linha[13].value
  pyperclip.copy(pais_origem)
  pyautogui.click(828,344,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Observações
  observacoes = linha[14].value
  pyperclip.copy(observacoes)
  pyautogui.click(827,429,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Codigo de barras
  codigo_barras = linha[15].value
  pyperclip.copy(codigo_barras)
  pyautogui.click(832,545,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')

  #Localizacao
  localizacao_armazem = linha[16].value
  pyperclip.copy(localizacao_armazem)
  pyautogui.click(850,618,duration=1)
  time.sleep(0.5)
  pyautogui.hotkey('ctrl','v')
  
  pyautogui.click(850,678,duration=1)

  pyautogui.click(1209,192,duration=1)

  pyautogui.click(1025,471,duration=1)

  





  


