#%% !/usr/bin/env python3
"""
Cruza TODAS as fontes do IPCM pelo Código IBGE de 7 dígitos, calcula indicadores e gera a base integrada.

Entradas (aba "Dados" de cada arquivo, geradas pelos scripts de coleta), todas na mesma pasta:
    dados_ibge_municipios.xlsx      <- coletar_ibge.py        (PIB, VAB, população, domicílios)
    dados_estban_municipios.xlsx    <- coletar_estban.py      (crédito, depósitos, agências)
    dados_frota_municipios.xlsx     <- coletar_frota.py       (frota por grupo de veículo)
    dados_caged_municipios.xlsx     <- coletar_caged.py       (emprego CLT: movimentação, estoque estimado)
    dados_pendentes_municipios.xlsx <- coletar_pendentes.py   (Censo: área/densidade/urbanização; CEMPRE: massa
                                                               salarial; frota por idade)
    dados_postos_municipios.xlsx    <- coletar_postos_bcb.py  (Cooperativas e postos do BCB)
Saídas:
    base_integrada.xlsx  (abas: Dados, Verificação, Alertas, Cobertura, Estatísticas, Dicionário, detalhes)
    base_integrada.csv   (a mesma aba Dados, UTF-8, separador ',' e decimal '.' - para BI / dashboard)

Requisitos: pip install pandas openpyxl
"""
import os
import re

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ARQ = {"ibge": "dados_ibge_municipios.xlsx", "estban": "dados_estban_municipios.xlsx",
       "frota": "dados_frota_municipios.xlsx", "caged": "dados_caged_municipios.xlsx",
       "pend": "dados_pendentes_municipios.xlsx"}
SAIDA_XLSX, SAIDA_CSV = "base_integrada.xlsx", "base_integrada.csv"
INCLUIR_ABAS_DETALHE = True        # copia "Frota por tipo" e "Movimentação mensal" (CAGED) para a saída
SALARIO_MINIMO = {2019: 998, 2020: 1045, 2021: 1100, 2022: 1212, 2023: 1320, 2024: 1412, 2025: 1518, 2026: 1621}
CH = "Código IBGE (7 dígitos)"

ausentes = []                       # colunas esperadas que não vieram nos arquivos (aviso no final)


def le(chave, aba="Dados", **kw):
    if not os.path.exists(ARQ[chave]):
        raise FileNotFoundError(f"Arquivo '{ARQ[chave]}' não encontrado na pasta atual ({os.getcwd()}). "
                                "Rode antes o script de coleta correspondente.")
    return pd.read_excel(ARQ[chave], sheet_name=aba, **kw)


def unico(df, nome):
    if not df[CH].is_unique:
        dup = df[df[CH].duplicated(keep=False)][CH].unique()[:10]
        raise ValueError(f"{nome}: código IBGE repetido (ex.: {list(dup)}). Corrija a coleta antes de cruzar.")


def acha(df, prefixo):
    c = [x for x in df.columns if x.startswith(prefixo)]
    if len(c) != 1:
        raise KeyError(f"Coluna que começa com '{prefixo}' não encontrada de forma única: {c}")
    return c[0]


def col(df, nome):
    """Coluna se existir; senão série vazia (e registra o aviso)."""
    if nome in df.columns:
        return pd.to_numeric(df[nome], errors="coerce")
    if nome not in ausentes:
        ausentes.append(nome)
    return pd.Series(np.nan, index=df.index)


def razao(a, b, mult=1.0):
    """a / b; denominador <= 0 ou vazio gera vazio (nunca infinito)."""
    return (a / b.where(b > 0)) * mult


# ----------------------------------------------------------------------------- leitura
print("Lendo as fontes...")
ib = le("ibge")
es = le("estban", dtype={"Período (AAAAMM)": str})
fr = le("frota", dtype={"Período (AAAAMM)": str})
cg = le("caged")
pe = le("pend", dtype={"Mês da frota por idade (AAAAMM)": str})
fontes = {"IBGE": ib, "ESTBAN": es, "Frota": fr, "CAGED": cg, "Pendentes": pe}
for nome, d in fontes.items():
    unico(d, nome)
    print(f"  {nome:10s}: {len(d):>5,} municípios, {d.shape[1]:>2} colunas".replace(",", "."))

ano_pib = int(ib["Ano do PIB"].iloc[0])
ano_vab = ib["Ano do VAB"].iloc[0]
ano_est = int(ib["Ano da estimativa"].iloc[0])
metodo_pop = str(ib["Método da população (ano do PIB)"].iloc[0])
per_es = str(es["Período (AAAAMM)"].dropna().iloc[0])
per_fr = str(fr["Período (AAAAMM)"].dropna().iloc[0])
comp_cg = str(int(cg[acha(cg, "Competência final")].dropna().iloc[0]))
ini_cg = str(int(cg[acha(cg, "Início da janela")].dropna().iloc[0]))
ano_cg = int(comp_cg[:4])
ano_cem = int(pe["Ano do CEMPRE"].dropna().iloc[0]) if "Ano do CEMPRE" in pe.columns and pe["Ano do CEMPRE"].notna().any() else None
mes_idade = str(pe["Mês da frota por idade (AAAAMM)"].dropna().iloc[0]) if "Mês da frota por idade (AAAAMM)" in pe.columns and pe["Mês da frota por idade (AAAAMM)"].notna().any() else None
_m = re.search(r"/(\d{4})", acha(cg, "Estoque CLT RAIS"))
ano_rais = int(_m.group(1)) if _m else None

periodos = pd.DataFrame([
    ("IBGE", "PIB total e PIB per capita", str(ano_pib)),
    ("IBGE", "VAB por setor", str(int(ano_vab)) if pd.notna(ano_vab) else "indisponível"),
    ("IBGE", "População estimada (última)", str(ano_est)),
    ("IBGE", "População no ano do PIB", f"{ano_pib} ({metodo_pop})"),
    ("IBGE", "Domicílios, área, densidade, urbanização 2022; crescimento 2010-2022", "Censo 2022 / 2010"),
    ("IBGE (CEMPRE)", "Massa salarial formal", str(ano_cem) if ano_cem else "indisponível"),
    ("BCB (ESTBAN)", "Crédito, depósitos, agências", per_es),
    ("SENATRAN", "Frota por tipo", per_fr),
    ("SENATRAN", "Frota por ano de fabricação (idade)", mes_idade or "indisponível"),
    ("MTE (Novo CAGED)", "Fluxos de emprego (12 meses)", f"{ini_cg} a {comp_cg}"),
    ("MTE (RAIS + CAGED)", "Estoque CLT estimado", f"RAIS {ano_rais} + saldo até {comp_cg}" if ano_rais else "indisponível (sem RAIS)"),
], columns=["Fonte", "Informação", "Período de referência"])
print("\nPeríodos de referência (cada fonte usa o dado mais recente disponível):")
print(periodos.to_string(index=False))
print("  Observação: razões entre fontes de períodos diferentes (ex.: crédito/PIB) servem para COMPARAR municípios, não como valor absoluto.")

# ----------------------------------------------------------------------------- padroniza colunas
es_o = es.copy()
es = es.drop(columns=["Município", "UF"]).rename(columns={"Período (AAAAMM)": "ESTBAN período (AAAAMM)"})
fr = fr.drop(columns=["Município", "UF"]).rename(columns={"Período (AAAAMM)": "Frota período (AAAAMM)"})
cg = cg.drop(columns=["Município", "UF"]).rename(columns={
    acha(cg, "Competência final"): "CAGED competência final (AAAAMM)",
    acha(cg, "Início da janela"): "CAGED início da janela de 12 meses (AAAAMM)",
    acha(cg, "Estoque CLT RAIS"): "Estoque CLT base RAIS (31/12 do ano-base)",
    acha(cg, "Admissões"): "Admissões CAGED (12m)", acha(cg, "Desligamentos"): "Desligamentos CAGED (12m)",
    acha(cg, "Saldo ("): "Saldo CAGED (12m)", acha(cg, "Saldo acumulado"): "Saldo CAGED acumulado desde a base",
    acha(cg, "Estoque CLT estimado"): "Estoque CLT estimado (competência final)",
    acha(cg, "Salário médio de admissão"): "Salário médio de admissão 12m (R$)",
    acha(cg, "Salário médio de desligamento"): "Salário médio de desligamento 12m (R$)"})
pe = pe.drop(columns=["Município", "UF"], errors="ignore")

# ----------------------------------------------------------------------------- cruzamento (base = IBGE)
df = ib.copy()
for d in (es, fr, cg, pe):
    df = df.merge(d, on=CH, how="left", validate="one_to_one")
assert len(df) == len(ib), "O cruzamento alterou o número de municípios."

df["Tem dado ESTBAN"] = df["ESTBAN período (AAAAMM)"].notna()
df["Tem dado frota"] = df["Frota período (AAAAMM)"].notna()
df["Tem dado CAGED"] = df["CAGED competência final (AAAAMM)"].notna()
df["Tem dado CEMPRE"] = col(df, "Pessoal ocupado assalariado (CEMPRE)").notna()
df["Tem dado Censo"] = col(df, "Área territorial (km²)").notna()
df["Tem dado frota por idade"] = col(df, "Idade média da frota (anos)").notna()
df["Agências (ESTBAN)"] = df["Agências (ESTBAN)"].fillna(0)      # sem registro no ESTBAN = sem agência

# -----------------------------------------------------------------------------
# INTEGRAÇÃO DOS POSTOS DO BCB (COOPERATIVAS E TRADICIONAIS)
# -----------------------------------------------------------------------------
print("\nA integrar base de cooperativas e postos do BCB...")
if os.path.exists("dados_postos_municipios.xlsx"):
    df_postos = pd.read_excel("dados_postos_municipios.xlsx")
    
    # CORREÇÃO: Lê o tipo de dado ('dtype') da base principal e aplica aos postos
    df_postos[CH] = df_postos[CH].astype(df[CH].dtype) 
    
    df = df.merge(df_postos, on=CH, how="left")
    
    # Preenche com 0 os municípios que não têm postos
    colunas_postos = ['Total_Postos_BCB', 'Postos_Cooperativas', 'Postos_Bancos_Tradicionais']
    for c in colunas_postos:
        if c in df.columns:
            df[c] = df[c].fillna(0)
    
    # Cria a coluna unificada somando agências originais com os novos postos
    df['Agências e Postos (Total)'] = df['Agências (ESTBAN)'] + df['Total_Postos_BCB']
    print(f"  Integração concluída. {(df['Agências e Postos (Total)'] > 0).sum()} municípios com infraestrutura de crédito.")
else:
    print("  Aviso: 'dados_postos_municipios.xlsx' não encontrado. Usando apenas agências do ESTBAN.")
    df['Agências e Postos (Total)'] = df['Agências (ESTBAN)']
# -----------------------------------------------------------------------------
# ----------------------------------------------------------------------------- indicadores
pop = col(df, "População (ano do PIB)")          # ano do PIB (não usada nos indicadores recentes)
pop_rec = col(df, "População estimada")             # última estimativa do IBGE (dados de 2025-2026)
pop_cs = col(df, "População Censo 2022")
pib_mil = col(df, "PIB total (R$ mil)")
pib_r = pib_mil * 1000
cred = col(df, "Saldo de operações de crédito (R$)")
dep = col(df, "Total depósitos à vista + a prazo (R$)")
ag = col(df, "Agências (ESTBAN)")
ag_postos = col(df, "Agências e Postos (Total)") # NOVA VARIÁVEL
vab = (col(df, "VAB Agropecuária (R$ mil)") + col(df, "VAB Indústria (R$ mil)")
       + col(df, "VAB Serviços total (R$ mil)"))
frota = col(df, "Total da frota")
est = col(df, "Estoque CLT estimado (competência final)")
est_base = col(df, "Estoque CLT base RAIS (31/12 do ano-base)")
sal_anual = col(df, "Salários e outras remunerações (R$ mil/ano)")
sal_exc = col(df, "Salários exc. administração pública (R$ mil/ano)")
assal = col(df, "Pessoal ocupado assalariado (CEMPRE)")
dom = col(df, "Domicílios particulares permanentes")

I = {}   # nome -> (série, fonte, descrição, formato)

def ind(nome, serie, fonte, desc, fmt):
    I[nome] = (serie.replace([np.inf, -np.inf], np.nan), fonte, desc, fmt)

# crédito e sistema bancário
ind("Crédito / PIB", razao(cred, pib_r), "ESTBAN ÷ IBGE", "Saldo de crédito sobre o PIB (PIB em R$)", "0.0%")
ind("Depósitos / PIB", razao(dep, pib_r), "ESTBAN ÷ IBGE", "Depósitos à vista + a prazo sobre o PIB", "0.0%")
ind("Crédito per capita (R$)", razao(cred, pop_rec), "ESTBAN ÷ IBGE", "Saldo de crédito por habitante", "#,##0.00")
ind("Depósitos per capita (R$)", razao(dep, pop_rec), "ESTBAN ÷ IBGE", "Depósitos à vista + a prazo por habitante", "#,##0.00")

# INDICADORES ATUALIZADOS PARA INCLUIR POSTOS DE COOPERATIVAS
ind("Habitantes por agência ou posto", razao(pop_rec, ag_postos), "ESTBAN/BCB ÷ IBGE", "População por agência ou posto (vazio se não há)", "#,##0")
ind("Agências e postos por 100 mil hab.", razao(ag_postos, pop_rec, 100_000), "ESTBAN/BCB ÷ IBGE", "Agências e postos por 100 mil habitantes", "0.0")

# crédito por modalidade (verbetes 161-169 do ESTBAN)
imob = col(df, "Crédito imobiliário (R$)")
rural = col(df, "Crédito rural (custeio e investimento) (R$)") + col(df, "Crédito agroindustrial (R$)")
ind("Crédito imobiliário per capita (R$)", razao(imob, pop_rec), "ESTBAN ÷ IBGE", "Financiamentos imobiliários (verbete 169) por habitante; concentrado nas sedes (Caixa em Brasília)", "#,##0.00")
ind("Crédito rural e agroindustrial per capita (R$)", razao(rural, pop_rec), "ESTBAN ÷ IBGE", "Financiamentos rurais (163) e agroindustriais (167) por habitante", "#,##0.00")
ind("Crédito rural e agroind. / VAB agropecuária", razao(rural, col(df, "VAB Agropecuária (R$ mil)") * 1000), "ESTBAN ÷ IBGE", "Crédito rural e agroindustrial sobre o VAB agropecuário (VAB de ano anterior)", "0.0%")
ind("% do crédito que é imobiliário", razao(imob, cred), "ESTBAN", "Participação dos financiamentos imobiliários no verbete 160", "0.0%")
# frota
ind("Veículos por mil hab.", razao(frota, pop_rec, 1000), "Frota ÷ IBGE", "Frota total por mil habitantes", "#,##0.0")
ind("Automóveis e utilitários por mil hab.", razao(col(df, "Automóveis e utilitários"), pop_rec, 1000), "Frota ÷ IBGE", "Automóveis, caminhonetes, camionetas e utilitários por mil habitantes", "#,##0.0")
ind("Motocicletas por mil hab.", razao(col(df, "Motocicletas"), pop_rec, 1000), "Frota ÷ IBGE", "Motocicletas e motonetas por mil habitantes", "#,##0.0")
ind("Motos por automóvel", razao(col(df, "Motocicletas"), col(df, "Automóveis e utilitários")), "Frota", "Razão motocicletas / automóveis e utilitários", "0.00")
ind("Caminhões, tratores e reboques por mil hab.", razao(col(df, "Caminhões, tratores e reboques"), pop_rec, 1000), "Frota ÷ IBGE", "Frota de carga e agrícola por mil habitantes", "#,##0.0")
ind("Veículos por domicílio", razao(frota, dom), "Frota ÷ IBGE", "Frota total por domicílio (Censo 2022)", "0.00")
ind("Frota com 10+ anos por mil hab.", razao(col(df, "Frota com 10 anos ou mais (unidades)"), pop_rec, 1000), "Frota idade ÷ IBGE", "Frota renovável (10+ anos) por mil habitantes", "#,##0.0")
ind("Frota (arquivo por idade) / Frota (arquivo por tipo)", razao(col(df, "Frota total (arquivo por ano de fabricação)"), frota), "Conferência", "Deve ficar próximo de 1: compara os dois arquivos da SENATRAN", "0.00")
# emprego e renda
ind("Estoque CLT por 100 hab.", razao(est, pop_rec, 100), "CAGED/RAIS ÷ IBGE", "Empregos CLT estimados por 100 habitantes", "0.0")
ind("Crescimento do emprego CLT em 12 meses", razao(col(df, "Saldo CAGED (12m)"), est_base), "CAGED ÷ RAIS", "Saldo dos últimos 12 meses sobre o estoque base RAIS (aproximação: a base tem mais de 12 meses)", "0.0%")
ind("Rotatividade CAGED", razao((col(df, "Admissões CAGED (12m)") + col(df, "Desligamentos CAGED (12m)")) / 2, est_base), "CAGED ÷ RAIS", "Média de admissões e desligamentos (12 meses) sobre o estoque base RAIS", "0.0%")
if ano_cg in SALARIO_MINIMO:
    sm = SALARIO_MINIMO[ano_cg]
    ind("Salário de admissão (em salários mínimos)", col(df, "Salário médio de admissão 12m (R$)") / sm, "CAGED", f"Salário médio de admissão (12 meses) ÷ salário mínimo de {ano_cg} (R$ {sm}); a janela atravessa dois valores de mínimo, então é aproximação", "0.00")
ind("Empregos CLT por R$ milhão de PIB", razao(est, pib_mil / 1000), "CAGED/RAIS ÷ IBGE", "Empregos CLT estimados por R$ milhão de PIB", "#,##0.0")
ind("Massa salarial per capita (R$/mês)", razao(sal_anual * 1000 / 12, pop_rec), "CEMPRE ÷ IBGE", "Salários e outras remunerações (todos os assalariados formais) por habitante, por mês", "#,##0.00")
ind("Massa salarial exc. adm. pública per capita (R$/mês)", razao(sal_exc * 1000 / 12, pop_rec), "CEMPRE ÷ IBGE", "Massa salarial sem administração pública, por habitante, por mês", "#,##0.00")
ind("Massa salarial / PIB", razao(sal_anual, pib_mil), "CEMPRE ÷ IBGE", "Salários anuais sobre o PIB", "0.0%")
ind("Assalariados formais por 100 hab.", razao(assal, pop_rec, 100), "CEMPRE ÷ IBGE", "Pessoal assalariado (CEMPRE) por 100 habitantes", "0.0")
ind("Estoque CLT estimado / assalariados CEMPRE", razao(est, assal), "Conferência", "Compara estoque CLT estimado com assalariados do CEMPRE (que inclui estatutários)", "0.00")
# estrutura produtiva e demografia
ind("% VAB Agropecuária", razao(col(df, "VAB Agropecuária (R$ mil)"), vab), "IBGE", "Participação no VAB (agro + indústria + serviços)", "0.0%")
ind("% VAB Indústria", razao(col(df, "VAB Indústria (R$ mil)"), vab), "IBGE", "Participação no VAB", "0.0%")
ind("% VAB Serviços (inclui adm. pública)", razao(col(df, "VAB Serviços total (R$ mil)"), vab), "IBGE", "Participação no VAB", "0.0%")
ind("Habitantes por domicílio", razao(pop_cs.fillna(pop_rec), dom), "IBGE", "População do Censo 2022 ÷ domicílios do Censo 2022", "0.00")

vazios = [n for n, (s, *_r) in I.items() if s.notna().sum() == 0]
if vazios:
    print(f"  Indicadores sem nenhum valor (dados de origem ausentes), removidos: {vazios}")
for n, (s, *_r) in I.items():
    if n not in vazios:
        df[n] = s
ind_ok = {n: v for n, v in I.items() if n not in vazios}

# ----------------------------------------------------------------------------- formatos e dicionário
def formato(c):
    if c in ind_ok:
        return ind_ok[c][3]
    cl = c.lower()
    if c.startswith(("Código", "Ano ", "CAGED competência", "CAGED início")) or c == "Ano":
        return "0"
    if "período" in cl or "(aaaamm)" in cl or c.startswith("Tem dado"):
        return "@"
    if c.startswith("Crescimento populacional anual"):
        return "0.00%"
    if c.startswith(("%", "Grau de urbanização", "Crescimento populacional")):
        return "0.0%"
    if "salários mínimos" in cl or c.startswith("Idade mediana"):
        return "0.00" if "salários mínimos" in cl else "0"
    if c.startswith("Idade"):
        return "0.0"
    if "(r$)" in cl and "mil" not in cl:
        return "#,##0.00"
    if c.startswith(("Área", "Densidade")):
        return "#,##0.00"
    return "#,##0"


origem = {}
for nome, d in [("IBGE", ib), ("ESTBAN", es), ("Frota", fr), ("CAGED", cg), ("Pendentes (Censo/CEMPRE/frota idade)", pe)]:
    for c in d.columns:
        origem.setdefault(c, nome)
dic = []
for c in df.columns:
    if c in ind_ok:
        dic.append((c, f"Calculado ({ind_ok[c][1]})", ind_ok[c][2]))
    elif c.startswith("Tem dado"):
        dic.append((c, "Cruzamento", "VERDADEIRO se o município tem registro na fonte"))
    else:
        dic.append((c, origem.get(c, "IBGE"), "Coluna original da fonte"))
dic = pd.DataFrame(dic, columns=["Coluna", "Fonte", "Descrição"])

# ----------------------------------------------------------------------------- verificação (totais fonte x base)
def soma(d, c):
    return float(pd.to_numeric(d[c], errors="coerce").sum()) if c in d.columns else np.nan


base_cod = set(df[CH])
checks = []
for fonte, d, cols in [
    ("IBGE", ib, ["PIB total (R$ mil)", "População estimada"]),
    ("ESTBAN", es, ["Saldo de operações de crédito (R$)", "Total depósitos à vista + a prazo (R$)", "Agências (ESTBAN)"]),
    ("Frota", fr, ["Total da frota", "Motocicletas"]),
    ("CAGED", cg, ["Admissões CAGED (12m)", "Desligamentos CAGED (12m)"]),
    ("Pendentes", pe, ["População Censo 2022", "Salários e outras remunerações (R$ mil/ano)",
                       "Frota total (arquivo por ano de fabricação)"]),
]:
    fora = d[~d[CH].isin(base_cod)]
    for c in cols:
        if c not in d.columns:
            checks.append({"Fonte": fonte, "Métrica": c, "Municípios na fonte": len(d),
                           "Municípios fora da base IBGE": len(fora), "Total na fonte": np.nan,
                           "Total na base integrada": np.nan, "Diferença": np.nan, "Situação": "Coluna ausente"})
            continue
        v_f, v_b = soma(d, c), soma(df, c)
        checks.append({"Fonte": fonte, "Métrica": c, "Municípios na fonte": len(d),
                       "Municípios fora da base IBGE": len(fora), "Total na fonte": v_f,
                       "Total na base integrada": v_b, "Diferença": v_f - v_b,
                       "Situação": "OK" if abs(v_f - v_b) < 0.5 else "Diferença (municípios fora da base?)"})
ver = pd.DataFrame(checks)
print("\nVerificação (total na fonte x total na base):")
print(ver[["Fonte", "Métrica", "Municípios fora da base IBGE", "Diferença", "Situação"]].to_string(index=False))
for fonte, d in fontes.items():
    f = d[~d[CH].isin(base_cod)][CH].tolist()
    if f:
        print(f"  {fonte}: {len(f)} código(s) fora da base IBGE: {f[:10]}")

# ----------------------------------------------------------------------------- alertas de qualidade
regras = [("Crédito / PIB acima de 300%", "Crédito / PIB", lambda s: s > 3),
          ("Veículos por mil hab. acima de 1.500", "Veículos por mil hab.", lambda s: s > 1500),
          ("Estoque CLT por 100 hab. acima de 80", "Estoque CLT por 100 hab.", lambda s: s > 80),
          ("Idade média da frota acima de 30 anos", "Idade média da frota (anos)", lambda s: s > 30),
          ("Salário médio de admissão abaixo de 0,8 SM", "Salário de admissão (em salários mínimos)", lambda s: s < 0.8),
          ("Frota (idade)/(tipo) fora de 0,9–1,1", "Frota (arquivo por idade) / Frota (arquivo por tipo)", lambda s: (s < 0.9) | (s > 1.1)),
          ("Massa salarial / PIB acima de 100%", "Massa salarial / PIB", lambda s: s > 1)]
alertas = []
for nome, c, f in regras:
    if c not in df.columns:
        continue
    s = df[c]
    m = f(s).fillna(False)
    ex = df.loc[m].assign(_v=s[m]).sort_values("_v", ascending=False).head(5)
    exemplos = "; ".join(f"{r['Município']}-{r['UF']} ({r['_v']:.2f})" for _, r in ex.iterrows())
    alertas.append({"Regra": nome, "Municípios": int(m.sum()), "Exemplos (até 5)": exemplos})
alertas = pd.DataFrame(alertas, columns=["Regra", "Municípios", "Exemplos (até 5)"])
print("\nAlertas de qualidade:")
print(alertas[["Regra", "Municípios"]].to_string(index=False) if len(alertas) else "  (nenhuma regra aplicável)")

# ----------------------------------------------------------------------------- cobertura e estatísticas
cobertura = pd.DataFrame({"Coluna": df.columns, "Preenchidos": [int(df[c].notna().sum()) for c in df.columns]})
cobertura["% preenchido"] = cobertura["Preenchidos"] / len(df)
cobertura["Fonte"] = cobertura["Coluna"].map(dict(zip(dic["Coluna"], dic["Fonte"])))
print("\nColunas com menos de 50% de preenchimento:")
baixas = cobertura[cobertura["% preenchido"] < 0.5]
print(baixas[["Coluna", "% preenchido"]].to_string(index=False) if len(baixas) else "  (nenhuma)")
num_cols = [c for c in df.columns if df[c].dtype.kind in "if" and not c.startswith(("Código", "Ano", "CAGED competência", "CAGED início"))]
est_desc = df[num_cols].describe(percentiles=[.01, .25, .5, .75, .99]).T.reset_index().rename(columns={"index": "Coluna"})

# ----------------------------------------------------------------------------- saída
FONT = Font(name="Arial", size=10)
HEAD = (Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"),
        Alignment(wrap_text=True, vertical="center", horizontal="center"))


def cabecalho(ws, tab, largura=18, altura=48):
    for j, c in enumerate(tab.columns, 1):
        h = ws.cell(row=1, column=j)
        h.font, h.fill, h.alignment = HEAD
        ws.column_dimensions[get_column_letter(j)].width = 32 if c == "Município" else largura
    ws.row_dimensions[1].height = altura


def estiliza(ws, tab, formatos=None, largura=18, freeze="D2", altura=48):
    cabecalho(ws, tab, largura, altura)
    for j, c in enumerate(tab.columns, 1):
        fm = (formatos or {}).get(c)
        for i in range(2, len(tab) + 2):
            cel = ws.cell(row=i, column=j)
            cel.font = FONT
            if fm:
                cel.number_format = fm
    if freeze:
        ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions


formatos = {c: formato(c) for c in df.columns}
print("\nGravando o Excel (pode levar alguns instantes)...")
with pd.ExcelWriter(SAIDA_XLSX, engine="openpyxl") as w:
    df.to_excel(w, sheet_name="Dados", index=False)
    estiliza(w.sheets["Dados"], df, formatos)
    tabs = [("Verificação", ver, {"Total na fonte": "#,##0.00", "Total na base integrada": "#,##0.00",
                                 "Diferença": "#,##0.00"}, 28),
            ("Alertas", alertas, None, 60),
            ("Cobertura", cobertura, {"% preenchido": "0.0%", "Preenchidos": "#,##0"}, 40),
            ("Estatísticas", est_desc, {c: "#,##0.00" for c in est_desc.columns if c != "Coluna"}, 22),
            ("Períodos", periodos, None, 60),
            ("Dicionário", dic, None, 70)]
    for nome, tab, fm, larg in tabs:
        tab.to_excel(w, sheet_name=nome, index=False)
        estiliza(w.sheets[nome], tab, fm, largura=larg, freeze="A2", altura=26)
    if INCLUIR_ABAS_DETALHE:
        for chave, aba, nome in [("frota", "Frota por tipo", "Frota por tipo"),
                                 ("caged", "Movimentação mensal", "CAGED mensal")]:
            try:
                d = le(chave, aba)
            except Exception as e:
                print(f"  Aba '{aba}' não incluída: {e}")
                continue
            d.to_excel(w, sheet_name=nome, index=False)
            cabecalho(w.sheets[nome], d)
            w.sheets[nome].freeze_panes = "E2"
    notas = pd.DataFrame({"Notas": [
        "Base integrada de 5 arquivos de coleta; chave = Código IBGE de 7 dígitos; a base de municípios é a do IBGE/SIDRA (left join).",
        "Cada fonte usa o dado MAIS RECENTE disponível; os períodos estão na aba 'Períodos'. Razões entre fontes de períodos diferentes "
        "(ex.: crédito/PIB, massa salarial/PIB) comparam municípios entre si, mas não são valores absolutos; evite usá-las como nível.",
        "O VAB setorial é de um ano anterior ao do PIB: o IBGE publicou o PIB 2022-2023 sem a abertura por setor.",
        "Per capita de dados recentes usa a última população estimada do IBGE; indicadores de domicílio usam o Censo 2022.",
        "Colunas 'Tem dado ...' indicam se o município aparece em cada fonte. Ausência no ESTBAN costuma significar sem agência (agências = 0); "
        "nas demais fontes a ausência fica VAZIA (não é zero).",
        "O código '999999' do CAGED (movimentos sem município) não existe na tabela do IBGE (provavelmente 'município ignorado') e não entra na base; "
        "o CAGED/API do IBGE pode listar 1 município a mais que o SIDRA (possível município recém-criado), que também fica de fora.",
        "Estoque de emprego CLT é estimativa (RAIS do ano anterior + saldo do Novo CAGED). Salário do CAGED refere-se a admitidos/desligados.",
        "Massa salarial do CEMPRE é de todos os assalariados formais, inclusive administração pública; use as colunas 'exc. administração pública' "
        "para o setor privado.",
        "Crédito e depósitos do ESTBAN são contabilizados no município da agência; razões altas em polos regionais refletem essa concentração.",
        "Idade da frota refere-se à frota total (o arquivo da SENATRAN não separa por tipo).",
        "Veja 'Verificação' (totais fonte x base), 'Alertas' (valores suspeitos), 'Cobertura' (preenchimento por coluna), "
        "'Estatísticas' (distribuição, útil para normalização por percentil) e 'Dicionário'.",
    ]})
    notas.to_excel(w, sheet_name="Fontes e notas", index=False)
    estiliza(w.sheets["Fontes e notas"], notas, largura=170, freeze="A2", altura=24)

df.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig")
if ausentes:
    print(f"\nAviso: colunas esperadas que não vieram nos arquivos de entrada (indicadores dependentes ficam vazios): {ausentes}")
print(f"\nSalvo: {SAIDA_XLSX} e {SAIDA_CSV} | {len(df):,} municípios x {df.shape[1]} colunas".replace(",", "."))
# %%