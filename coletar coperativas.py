#%%
import pandas as pd
import numpy as np

def processar_postos_bcb(caminho_arquivo="202606POSTOS.xlsx", caminho_saida="dados_postos_municipios.xlsx"):
    print(f"Lendo o arquivo {caminho_arquivo}...")
    
    # Passo 1: Leitura do arquivo pulando as 9 primeiras linhas (header=9) para pegar os nomes das colunas
    df = pd.read_excel(caminho_arquivo, header=9)
    df.columns = df.columns.str.strip() # Remove espaços vazios dos nomes das colunas
    
    # Passo 2: Limpeza da chave IBGE
    # Removemos linhas que não possuem código de município e garantimos que é uma string de 7 dígitos
    df = df.dropna(subset=['MUNICIPIO IBGE'])
    df['Código IBGE (7 dígitos)'] = df['MUNICIPIO IBGE'].astype(int).astype(str)
    
    # Passo 3: Classificação
    # Cria uma coluna booleana (1 ou 0) se a palavra 'Cooperativa' existir no segmento
    df['É_Cooperativa'] = df['SEGMENTO'].str.contains("Cooperativa", na=False, case=False).astype(int)
    df['Total_Postos'] = 1 # Flag unitária para somarmos no groupby
    
    # Passo 4: Agregação por Município
    print("Agrupando dados por município...")
    agrupado = df.groupby('Código IBGE (7 dígitos)').agg(
        Total_Postos_BCB=('Total_Postos', 'sum'),
        Postos_Cooperativas=('É_Cooperativa', 'sum')
    ).reset_index()
    
    # Calcula os postos que são bancos tradicionais por diferença matemática
    agrupado['Postos_Bancos_Tradicionais'] = agrupado['Total_Postos_BCB'] - agrupado['Postos_Cooperativas']
    
    # Passo 5: Exportação e Diagnóstico
    agrupado.to_excel(caminho_saida, index=False)
    
    print("\n--- Diagnóstico ---")
    print(f"Total de municípios atendidos no BCB (Postos): {len(agrupado)}")
    print(f"Municípios com presença de Cooperativas: {(agrupado['Postos_Cooperativas'] > 0).sum()}")
    print(f"\nArquivo salvo com sucesso em: {caminho_saida}")

if __name__ == "__main__":
    processar_postos_bcb()
# %%
