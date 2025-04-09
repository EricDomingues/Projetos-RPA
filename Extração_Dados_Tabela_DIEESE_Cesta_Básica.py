# Esse script entra no site do DIEESE extrai a base de valor médio da cesta básica por cidade capital, trata a tabela criando uma coluna de data, cidades e faz um proc para obter a região e uf pertencente a cidade capital com um período de um ano em relação ao mês atual.

import pandas as pd
import os
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

timestart = datetime.now().strftime("%d.%m.%Y_%Hh.%M")

origem = "C://Automacao_tabela_cesta"

# Caminho de Origem do Arquivo
caminho_casos_rodados = f"{origem}//0_CASOS_RODADOS"
caminho_input = f"{origem}//1_INPUT"
caminho_output = f"{origem}//2_OUTPUT" 
caminho_tabelas_auxiliares = f"{origem}//3_Tabelas_Auxiliares"


########################################################### DOWNLOAD DO EXCEL DO SITE DIESE

caminho_download = 'C:\\Automacao_tabela_cesta\\1_INPUT' # TIVE QUE COLOCAR DOIS CAMINHOS PQ POR ALGUM MOTIVO O OPTION DA SELENIUM SÓ ACEITA DIRETÓRIO COM "\"

# Configurando o Local Padrão para Armazenar do Site
preferences = {
    "download.default_directory": caminho_download,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}

chrome_options = Options()
chrome_options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome(options=chrome_options)

driver.get("https://www.dieese.org.br/cesta/")

espera = WebDriverWait(driver, 10)

data_inicio = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="produtoDataInicial"]')))
data_final = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="produtoDataFinal"]')))


###################### Pegando sempre um período de 1 ano
data_final_num = datetime.now() - relativedelta(months=1) ## Excluindo o mês atual
data_final_formatada = data_final_num.strftime('%m%Y') 

data_inicio_num = datetime.now() - relativedelta(months=14) ## Limitando para trazer 14 meses
data_inicio_formatada = data_inicio_num.strftime('%m%Y') 

data_inicio.send_keys(data_inicio_formatada)
data_final.send_keys(data_final_formatada)
###########################################################


### BOTÃO EXTRAIR
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="produtoForm"]/div[3]/p[3]/input[1]'))).click()

### DOWNLOAD DO ITEM
espera.until(EC.presence_of_element_located((By.XPATH, '//input[@value="Exporta Excel"]'))).click()

time.sleep(3)

driver.close()

###########################################################


########################################################### TRATAMENTO DA TABELA DO DIESE

tab_diese = pd.read_excel(f'{caminho_input}//exporta.xls')
tab_diese = tab_diese.drop(tab_diese.tail(2).index)

tab_diese.to_excel(f'{caminho_output}//tab_diese.xlsx', index=False, header=None)

shutil.move(f'{caminho_input}//exporta.xls', caminho_casos_rodados) ## Move o arquivo baixado para pasta de casos rodados

# Renomeando o Arquivo
tab_diese_novo = f'{caminho_casos_rodados}//tab_diese_input_{timestart}.xls'
os.rename(f'{caminho_casos_rodados}//exporta.xls', tab_diese_novo)

tab_diese_principal = pd.read_excel(f'{caminho_output}//tab_diese.xlsx')
tab_diese_principal = tab_diese_principal.rename(columns={'Unnamed: 0': 'data_referencia'})

#########################################################################################


############################################################################# FINALIZAÇÃO


tab_diese_principal = tab_diese_principal.melt(id_vars=['data_referencia'], var_name='cidade', value_name='valor_cesta')
tab_diese_principal = tab_diese_principal.dropna() # Removendo linhas nulas

tab_regiao_capital = pd.read_excel(f'{caminho_tabelas_auxiliares}//tabela_capitais.xlsx')

## JOIN entre a tabela principal e uma auxiliar de para de detalhe da cidade, região, UF, ...
tab_diese_principal = tab_diese_principal.merge(tab_regiao_capital[['cidade', 'regiao', 'uf', 'dsc_uf']], on='cidade', how='left')

## Convertendo "data_referencia" para o formato de data
tab_diese_principal['data_referencia'] = pd.to_datetime(tab_diese_principal['data_referencia']).dt.date

## Condicional para trazer sempre 1 ano de base
if max(tab_diese_principal['data_referencia']).strftime('%m%Y') != data_final_formatada:
    print("O site do DIESE não esta com a base de dados atualizada")
else:
    tab_diese_principal = tab_diese_principal[tab_diese_principal['data_referencia'] > pd.to_datetime(data_inicio_num.strftime('%Y-%m-01')).date()]
    print("O site do DIESE esta com a base de dados atualizada")


tab_diese_principal.to_excel(f'{caminho_output}//tab_diese_{timestart}.xlsx', index=False)

## Excluindo arquivo antigo
os.remove(f'{caminho_output}//tab_diese.xlsx')
