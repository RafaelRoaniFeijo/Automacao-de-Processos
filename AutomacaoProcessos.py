#!/usr/bin/env python
# coding: utf-8

# In[19]:


#Biblioteca que serve para automatizar tarefas e processos
#Automatiza o mouse, o teclado e a tela do computador
get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[20]:


import pyautogui
import pyperclip
import time
pyautogui.PAUSE = 1 #espera 1s entre os comandos

#pyautogui.hotkey -> conjunto de teclas
#pyautogui.write -> escrever um texto
#pyautogui.press -> apertar uma tecla

#Passo 1: Entrar no sistema da empresa (link do drive como exemplo)
#abrir aba no navegador, escrever o link e dar o enter
#comandos = pyautogui.o que quer fazer
#hotkey: conjunto de teclas (ctrl t abre uma aba no vavegador)
pyautogui.hotkey("ctrl", "t")

#se o navegador não estiver aberto
#pyautogui.press("win")
#pyautogui.write("firefox")
#pyautogui.press("enter")

#escrever o link

#copia o link e cola no navegador
pyperclip.copy("https://drive.google.com/drive/folders/1B3h6h-R1Mw_BhP2uYqkdeGmYL-DTeCYc")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
#demora alguns segundos para carregar por causa da internet
time.sleep(5) #só pausa o codigo nesse momento

#tempo/tem que esperar a tela carregar, coloca o pause no inicio
#tem alguns caracteres especiais, o pyperclip permite isso
#se tivesse que logar, pegar o mouse, escrever o login e senha

#Passo 2: Navegar no sistema e encontrar a base de dados(entrar na pasta exportar)
#dar duplo clique na pasta exportar com o MOUSE com a posição da tela
pyautogui.click(x=946, y=397, clicks = 2) #texto entre aspas, numeros não -
time.sleep(2) #tempo para o navegador carregar

#Passo 3: Exportar/Fazer Download da Base de Dados

#fazer download do arquivo: clica no arquivo, clica nos tres pontos, clica no download
pyautogui.click(x=403, y=381)
pyautogui.click(x=1154, y=167)
pyautogui.click(x=1044, y=548)
time.sleep(5) #esperar o download





# In[21]:


#Passo 4: Importar a base de dados para o Python
import pandas as pd #apelido que deu pro pandas
tabela = pd.read_excel(r"D:\Vendas - Dez.xlsx") #se tivesse mais de uma aba, colocaria " " , sheets = n° aba
#pode ler várias bases de dados e armazena no tabela
#sempre coloca r antes de um caminho do pc, para dizer que n tem caractere especial
display(tabela) #print em outros locais

#pd.+tab mostra todas oções


# In[22]:


#Passo 5: Calcular os indicadores
#faturamento = soma da coluna valor final
faturamento = tabela["Valor Final"].sum() #se quiser contar .count, media .average, se quiser somar .sum
#quantidade total
quantidade = tabela["Quantidade"].sum()

print(faturamento)


# In[23]:


#Passo 6: Enviar um email para diretoria com o relatório
#abrir email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://outlook.live.com/mail/0/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)


#clica no escrever
pyautogui.click(x=196, y=145)
time.sleep(2)
#clica no email destinatario
pyautogui.write("rafaelfeijo@acad.ftec.com.br")
pyautogui.press("tab") #para reconhecer o email
#escrever assunto
pyautogui.press("tab")
pyautogui.write("AutomacaoProcessos")
pyautogui.press("tab")
#escrever o email
#texto em varias linhas, f quer dizer que quer formatar e quer dizer que pode colocar variaveis dinamicas
#milhar e casas decimais
#separacao de milhar é virgula
texto = f"""
Prezados, bom dia
O faturamento de ontem foi de R${faturamento:,.2f}
A quantidade de produtos é de {quantidade:,} 

abs
"""
pyperclip.copy(texto) #se tiver caracter especial
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter") #manda email
#depois rodar tudo de uma vez


# In[24]:


#codigo para descobrir a posição no monitor
time.sleep(5)
pyautogui.position()
#coloca a posição onde tem a pasta que quer


# In[ ]:




