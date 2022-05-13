# criar o navegador
from selenium import webdriver
# localizar elementos (os itens de um site)
from selenium.webdriver.common.by import By
# permite clicar teclas no teclado
from selenium.webdriver.common.keys import Keys

import pandas as pd

navegador = webdriver.Chrome()

# Passo 1 - Entrar no google
navegador.get(r'https://www.google.com.br/')

# Passo 2 - Pesquisar a cotação do dolar
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dólar', Keys.ENTER)  #.click para clicar
# navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# Passo 3 - Pegar a cotação do dolar
cotacao_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(f'Cotação do dólar -> R${cotacao_dolar}\n')

# Passo 4 - Pesquisar e pegando a cotação do euro
navegador.get(r'https://www.google.com.br/')
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro', Keys.ENTER)

cotacao_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(f'Cotação do euro -> R${cotacao_euro}\n')

# Passo 5 - Pesquisar e pegando a cotação do ouro
navegador.get(r'https://www.melhorcambio.com/ouro-hoje')

cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
print(f'Cotação do ouro -> R${cotacao_ouro}')
cotacao_ouro = cotacao_ouro.replace(',', '.')
print(f'Cotação do ouro -> R${cotacao_ouro}\n\n')

# Passo 6 - Atualizar a minha base de dados com as novas cotações
# Atualizar na coluna de cotação aonde a moeda for dolar, colocar dolar e assim vai

tabela = pd.read_excel("Produtos.xlsx")

print(tabela,'\n\n')

# atualizar a cotação de acordo com a moeda correspondente
# dolar
# as linhas onde a coluna 'Moeda' = 'Dólar"
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)

#euro
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)

#ouro
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

print(tabela,'\n\n')

# atualizar o preço de compra -> preco original x cotação
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

# atualizar o preco de venda -> preco de compra * margem
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

# criar uma tabela nova
# tabela['Preço de Venda Atualzado'] = tabela['Preço de Compra'] * tabela['Margem']

print(tabela,'\n\n')

# Agora vamos exportar a nova base de preços atualizada (editar arquivo original ou salvar um novo)

tabela.to_excel('Produtos Com Index.xlsx')  # colocando o mesmo nome edita a tabela original
tabela.to_excel('Produtos Sem Index.xlsx', index=0)

navegador.quit()

print('\n\n **** PROGRAMA ENCERRADO **** ')
