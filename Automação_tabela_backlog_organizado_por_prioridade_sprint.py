# Este script em python foi realizado para automatizar uma planilha de backlog colocando os casos que devem ser tratados com priorida com base na criação de uma coluna de sprint

import pandas as pd
import numpy as np
from datetime import date, timedelta

caminho = 'C:/Users/Dell/OneDrive/Área de Trabalho/Backlog_vr'

input = f'{caminho}/01 - INPUT'
output = f'{caminho}/02 - OUTPUT'
auxiliar = f'{caminho}/03 - TABELA AUXILIAR/tab_teste.xlsx'

tab_auxiliar = pd.read_excel(auxiliar, sheet_name='teste')

tab_auxiliar_pendente = tab_auxiliar[tab_auxiliar['status'] == 'Pendente']
tab_auxiliar_pendente = tab_auxiliar_pendente.dropna(subset= ['esforço'])
tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço']


###########################################################################################################################
############################################################################################## Tratamento horas, dias e mês

tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço_dias'].str.replace('dias', '', regex=False)

tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço_dias'].str.replace('dia', '', regex=False)

tab_auxiliar_pendente['esforço_dias'] = np.where(
    tab_auxiliar_pendente['esforço_dias'].str.contains('hora', na=False),
    0.041,
    tab_auxiliar_pendente['esforço_dias']
)

def meses_esforço(valor):    
    if "meses" in str(valor):  # Verifica se a célula contém 'meses'
        valor = str(valor).replace("meses", "")  # Remove 'meses'
        return float(valor) * 20  # Multiplica por 20 dias úteis
    return valor  # Retorna o valor original se não contém 'semanas' 

def mes_esforço(valor):
    if "mês" in str(valor):  # Verifica se a célula contém 'mês'
        valor = str(valor).replace("mês", "")  # Remove 'mês'
        return float(valor) * 20  # Multiplica por 20 dias úteis
    return valor  # Retorna o valor original se não contém 'semanas'

def semanas_esforço(valor):  
    if "semanas" in str(valor):  # Verifica se a célula contém 'semanas'
        valor = str(valor).replace("semanas", "")  # Remove 'semanas'
        return float(valor) * 5  # Multiplica por 5 dias úteis
    return valor  # Retorna o valor original se não contém 'semanas'  
  
def semana_esforço(valor):
    if "semana" in str(valor):  # Verifica se a célula contém 'semana'
        valor = str(valor).replace("semana", "")  # Remove 'semana'
        return float(valor) * 5  # Multiplica por 5 dias úteis
    return valor  # Retorna o valor original se não contém 'semana'

# Função mês   
tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço_dias'].apply(meses_esforço)
tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço_dias'].apply(mes_esforço)
tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço_dias'].apply(semanas_esforço)
tab_auxiliar_pendente['esforço_dias'] = tab_auxiliar_pendente['esforço_dias'].apply(semana_esforço)

# Converte a coluna inteira para float
tab_auxiliar_pendente['esforço_dias'] = pd.to_numeric(tab_auxiliar_pendente['esforço_dias'], errors='coerce')

###########################################################################################################################


###########################################################################################################################
###################################################################### Montando sprint com base em soma acumulativa de dias

# Função corrigida para obter sextas-feiras
def proximas_sextas(numero_sextas):
    hoje = date.today()
    dias_para_sexta = (4 - hoje.weekday()) % 7
    primeira_sexta = hoje + timedelta(days=dias_para_sexta)
    return [primeira_sexta + timedelta(weeks=i) for i in range(numero_sextas)]


# Ordenando os dados e agrupando
df = tab_auxiliar_pendente.sort_values(by='esforço_dias', ascending=True)

# Priorizar casos de 'mensuração' e 'deu merda prod'
df['prioridade'] = np.where(
    df['grupo'].isin(['mensuração', 'merda em prod']), 
    1, 
    0
)

# Reordenar priorizando os casos marcados como prioridade
df = df.sort_values(by=['prioridade', 'esforço_dias'], ascending=[False, True])

# Soma cumulativa
df['soma_acumulativa'] = df['esforço_dias'].cumsum()

# Determinando grupos com base na soma
df['grupo2'] = (df['soma_acumulativa'] // 6).astype(int)

# Obter o número de grupos
numero_grupos = df['grupo2'].max() + 1

# Obter as próximas sextas-feiras
sextas = proximas_sextas(numero_grupos)

# Atribuir as sextas-feiras aos grupos
df['sprint'] = df['grupo2'].apply(lambda x: sextas[x])

###########################################################################################################################


df = df.drop(columns=['esforço_dias', 'prioridade', 'soma_acumulativa', 'grupo2'])

df.head(30)
