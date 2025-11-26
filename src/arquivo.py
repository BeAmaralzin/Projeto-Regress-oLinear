import pandas as pd
from sklearn.linear_model import LinearRegression
from dateutil.relativedelta import relativedelta
import numpy as np
import openpyxl
import sys 
import statsmodels.api as sm  # Nova biblioteca para o SARIMA

#Configurações Iniciais

nome_arquivo = r'C:/Users/Livia Bontempo/OneDrive/amaral/DADOS.xlsx'
aba_dados_limpos = 'dados_para_analise' 
aba_original = 'Plan1' 
ano_para_prever = 2025


#Carregar e Preparar os Dados

try:
    df = pd.read_excel(nome_arquivo, sheet_name=aba_dados_limpos)
except Exception as e:
    print(f"Erro ao ler a aba '{aba_dados_limpos}': {e}")
    print("Verifique se o nome do arquivo e da aba estão corretos.")
    sys.exit() 

df.columns = df.columns.astype(str).str.strip()
print(f"Colunas detectadas na aba '{aba_dados_limpos}': {list(df.columns)}")

if 'Data' in df.columns:
    date_col = 'Data'
else:
    for c in df.columns:
        cl = c.lower()
        if ('data' in cl) or ('mes' in cl) or ('mês' in cl) or ('month' in cl) or ('date' in cl):
            date_col = c
            break

if not date_col:
    print(f"Erro: não foi possível encontrar uma coluna de datas na aba '{aba_dados_limpos}'.")
    print("Colunas encontradas:", list(df.columns))
    print("Por favor renomeie a coluna de datas para 'Data' ou verifique a aba selecionada.")
    sys.exit()

#converter usando o formato  '%m/%Y' 
try:
    # tenta conversão estrita primeiro
    df['Data'] = pd.to_datetime(df[date_col], format='%m/%Y', errors='raise')
except Exception:
    # fallback: parser flexível (coerce para NaT em valores inválidos)
    df['Data'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    if df['Data'].isna().all():
        print(f"Erro: não foi possível converter a coluna '{date_col}' para datas.")
        print("Verifique os valores e o formato (ex: '08/2025' ou '08/25').")
        sys.exit()
    else:
        print(f"A coluna '{date_col}' foi convertida para datas usando um parser flexível. Alguns valores podem ter sido definidos como NaT.")

df = df.sort_values(by='Data')
df = df.dropna(subset=['QNT'])

# Preparar os Dados para o Modelo de Série Temporal 


df = df.set_index('Data')
y = df['QNT']
# --- 4. Treinar o Modelo SARIMA ---

print("Treinando o Modelo SARIMA")
print("Isso pode levar alguns segundos...")

# SAZONALIDADE
try:
    model = sm.tsa.statespace.SARIMAX(
        y,
        order=(1, 1, 1),
        seasonal_order=(1, 1, 1, 12),
        enforce_stationarity=False,
        enforce_invertibility=False
    )
    
    # Treina o modelo
    results = model.fit(disp=False) 
    
    print("Modelo Treinado com Sucesso")

except Exception as e:
    print(f"Erro ao treinar o modelo SARIMA: {e}")
    print("Verifique se há dados suficientes para a análise (pelo menos 2 ciclos sazonais, ex: 24 meses).")
    sys.exit()

ultima_data = y.index.max()

mes_final_desejado = 12 
ultimo_mes_dados = ultima_data.month
num_previsoes = mes_final_desejado - ultimo_mes_dados

if num_previsoes <= 0:
    print(f"\nOs dados já estão completos até Dezembro de {ano_para_prever}.")
    sys.exit()

print(f"\nCalculando {num_previsoes} Previsões para {ano_para_prever}")

forecast_object = results.get_forecast(steps=num_previsoes)
previsoes_series = forecast_object.predicted_mean

previsoes_finais = []

for proxima_data, previsao_qnt in previsoes_series.items():
    previsao_qnt_arredondada = round(previsao_qnt)
    
    previsoes_finais.append({
        'data': proxima_data,
        'mes_nome': proxima_data.strftime('%B'), 
        'previsao': previsao_qnt_arredondada
    })
    
    print(f"Previsão para {proxima_data.strftime('%m/%Y')}: {previsao_qnt_arredondada}")

#Escrever os Resultados

aba_alvo = 'graficos'
coluna_alvo = 5  
linhas_alvo = [11, 12, 13, 14] 

print(f"\nEscrevendo previsões na planilha '{aba_alvo}'")

if len(previsoes_finais) != len(linhas_alvo):
    print(f"Erro: O script gerou {len(previsoes_finais)} previsões, mas você especificou {len(linhas_alvo)} linhas.")
    print("Por favor, ajuste o número de meses a prever ou as 'linhas_alvo' no script.")
    sys.exit()

try:
    workbook = openpyxl.load_workbook(nome_arquivo)

    if aba_alvo in workbook.sheetnames:
        sheet = workbook[aba_alvo]
    else:
        print(f"Aviso: Aba '{aba_alvo}' não encontrada. Criando uma nova...")
        sheet = workbook.create_sheet(title=aba_alvo)


    for i, previsao in enumerate(previsoes_finais):

        linha_atual = linhas_alvo[i]
        valor_previsao = previsao['previsao']        
        mes_nome = previsao['mes_nome']        
        sheet.cell(row=linha_atual, column=coluna_alvo).value = valor_previsao        
        print(f"Valor '{valor_previsao}' (para {mes_nome}) salvo na célula E{linha_atual}")

    workbook.save(nome_arquivo)
    
    print(f"\nSucesso! Todas as {len(previsoes_finais)} previsões foram salvas em '{nome_arquivo}'.")

except Exception as e:
    print(f"\nOcorreu um erro ao tentar escrever na planilha: {e}")
    print("As previsões NÃO foram salvas.")