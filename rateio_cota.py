#%%
import pandas as pd
import numpy as np

def realizar_rateio_cotas():
    print("A iniciar o rateio top-down de cotas...")
    
    # 1. Carregar os resultados do IPCM (nível municipal)
    print("  -> A ler os resultados do IPCM...")
    ipcm_df = pd.read_excel('ipcm_resultado.xlsx', sheet_name='IPCM')
    
    # 2. Carregar os dados Macro (nível estadual)
    print("  -> A ler o ficheiro macro de cotas...")
    macro_df = pd.read_excel('InteractiveSheet_2026-05-19_06_02_38.xlsx', sheet_name='Página1')
    
    # 3. Filtrar o trimestre mais recente com dados válidos por UF
    # Como as cotas por UF são trimestrais, usamos os dados onde o valor > 0
    macro_valido = macro_df[macro_df['cotas_ativas_uf_total'] > 0]
    ultimo_periodo = macro_valido['ano_mes'].max()
    print(f"  -> Período macro utilizado para a distribuição: {ultimo_periodo}")
    
    # Guardar apenas a relação: Estado -> Cotas Totais daquele período
    cotas_por_estado = macro_valido[macro_valido['ano_mes'] == ultimo_periodo][['estado', 'cotas_ativas_uf_total']]
    
    # 4. Juntar os dados macro à nossa base de municípios
    # Fazemos um 'left join' ligando a coluna UF com a coluna 'estado'
    df_final = ipcm_df.merge(cotas_por_estado, left_on='UF', right_on='estado', how='left')
    
    # 5. Calcular a SOMA do IPCM Geral de cada estado
    soma_ipcm_uf = df_final.groupby('UF')['IPCM__geral'].sum().reset_index()
    soma_ipcm_uf = soma_ipcm_uf.rename(columns={'IPCM__geral': 'Soma_IPCM_Estado'})
    
    # Adicionar esta soma ao nosso dataframe principal
    df_final = df_final.merge(soma_ipcm_uf, on='UF', how='left')
    
    # 6. O Cálculo do Rateio (O peso de cada município dentro do seu estado)
    df_final['peso_municipio_no_estado'] = df_final['IPCM__geral'] / df_final['Soma_IPCM_Estado']
    
    # A estimativa final de cotas (arredondada a zero casas decimais para não termos "meias cotas")
    df_final['cotas_estimadas_municipio'] = (df_final['peso_municipio_no_estado'] * df_final['cotas_ativas_uf_total']).round(0)
    
    # 7. Limpar, Organizar e Guardar o resultado
    colunas_exportacao = [
        'Código IBGE (7 dígitos)', 'Município', 'UF', 'População estimada', 
        'IPCM__geral', 'Soma_IPCM_Estado', 'cotas_ativas_uf_total', 
        'peso_municipio_no_estado', 'cotas_estimadas_municipio'
    ]
    
    # Ordenar por estado e, dentro de cada estado, pelos municípios com mais cotas
    resultado = df_final[colunas_exportacao].sort_values(by=['UF', 'cotas_estimadas_municipio'], ascending=[True, False])
    
    caminho_saida = 'ipcm_cotas_estimadas.xlsx'
    resultado.to_excel(caminho_saida, index=False)
    
    print(f"\nSucesso! O ficheiro '{caminho_saida}' foi gerado.")
    print("\nTop 5 municípios com maior volume estimado de cotas:")
    top_5 = resultado[['Município', 'UF', 'cotas_estimadas_municipio']].sort_values(by='cotas_estimadas_municipio', ascending=False).head(5)
    print(top_5.to_string(index=False))

if __name__ == "__main__":
    realizar_rateio_cotas()
# %%
