#%% !/usr/bin/env python3
"""
Coleta os dados pendentes do IPCM (fontes públicas) e salva em XLSX:

  1. CENSO (IBGE/SIDRA)  - área, densidade, população 2022 e 2010, crescimento populacional e urbanização
  2. CEMPRE (IBGE/SIDRA) - massa salarial formal (salários e outras remunerações), pessoal assalariado,
                           salário médio mensal; também separa a administração pública
  3. FROTA POR IDADE (SENATRAN) - idade média da frota e % da frota com 10+ / 15+ anos (renovável)

Os IDs de variáveis/categorias do SIDRA são DESCOBERTOS pelos metadados da API (por nome), e não fixados
no código; cada bloco imprime o que encontrou. Cada bloco roda isolado: se um falhar, os outros seguem.

Requisitos: pip install requests pandas openpyxl
Uso:        python coletar_pendentes.py
"""
import difflib
import io
import re
import time
import unicodedata

import numpy as np
import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ANO_REF = None                 # None = ano do mês mais recente da frota (idade = ano do arquivo - ano de fabricação)
ANO_CEMPRE = None              # None = último ano do CEMPRE publicado (massa salarial)
MES_FROTA = None               # None = mês mais recente da frota por ano de fabricação
ANO_CENSO_ATUAL, ANO_CENSO_ANT = 2022, 2010
SAIDA = "dados_pendentes_municipios.xlsx"

SIDRA = "https://apisidra.ibge.gov.br/values"
AGREGADOS = "https://servicodados.ibge.gov.br/api/v3/agregados"
URL_IBGE_MUN = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios"
HEADERS = {"User-Agent": "Mozilla/5.0 (coleta-pendentes)"}
BASE_SENATRAN = "https://www.gov.br/transportes/pt-br/assuntos/transito"
URLS_FROTA_IDADE = {
    (2021, 12): BASE_SENATRAN + "/arquivos-senatran/estatisticas/renavam/2021/dezembro/"
                                "f_frota_por_uf_municipio_ano_fab_mod_dezembro_2021.xlsx",
}
MESES = {1: "janeiro", 2: "fevereiro", 3: "marco", 4: "abril", 5: "maio", 6: "junho", 7: "julho",
         8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}

# Correções manuais de pareamento (UF, nome SENATRAN normalizado) -> código IBGE de 7 dígitos
DE_PARA_MANUAL = {
}
UFS = {"ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM", "BAHIA": "BA", "CEARA": "CE",
       "DISTRITO_FEDERAL": "DF", "ESPIRITO_SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA",
       "MATO_GROSSO": "MT", "MATO_GROSSO_DO_SUL": "MS", "MINAS_GERAIS": "MG", "PARA": "PA",
       "PARAIBA": "PB", "PARANA": "PR", "PERNAMBUCO": "PE", "PIAUI": "PI", "RIO_DE_JANEIRO": "RJ",
       "RIO_GRANDE_DO_NORTE": "RN", "RIO_GRANDE_DO_SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR",
       "SANTA_CATARINA": "SC", "SAO_PAULO": "SP", "SERGIPE": "SE", "TOCANTINS": "TO"}


# ------------------------------------------------------------------ utilitários
def norm(txt):
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]+", "_", t.upper()).strip("_")


def txt(s):
    """minúsculas, sem acento, espaços simples (para casar nomes de variáveis/categorias)."""
    t = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", t).strip()


def get(url, tentativas=4, **kw):
    for i in range(tentativas):
        try:
            r = requests.get(url, headers=HEADERS, timeout=kw.pop("timeout", 180))
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == tentativas - 1:
                raise RuntimeError(f"Falha em {url}: {e}")
            time.sleep(2 ** i)


def num(s):
    return pd.to_numeric(s.astype(str).str.strip().str.replace(",", ".", regex=False), errors="coerce")


# ------------------------------------------------------------------ SIDRA: metadados e consulta
def anos_disponiveis(tab):
    try:
        per = get(f"{AGREGADOS}/{tab}/periodos").json()
        return sorted({int(str(p["id"])[:4]) for p in per})
    except Exception:
        return []


def meta(tab):
    return get(f"{AGREGADOS}/{tab}/metadados").json()


def acha_vars(m, padroes):
    """padroes: {chave: regex}. Casa com nome (e unidade) da variável, sem acentos. Devolve {chave: (id, nome)}."""
    out = {}
    for chave, pad in padroes.items():
        for v in m["variaveis"]:
            if re.search(pad, txt(v["nome"]) + " | " + txt(v.get("unidade") or "")):
                out[chave] = (v["id"], v["nome"])
                break
    return out


def caminho_classes(m, escolhas=None):
    """Monta '/c<classe>/<categoria>' para TODAS as classificações da tabela.
    escolhas: {regex_nome_da_classe: regex_nome_da_categoria}; as demais usam 'Total'."""
    escolhas = escolhas or {}
    partes = []
    for cl in m.get("classificacoes", []):
        cats = cl.get("categorias", [])
        pad = next((p for rc, p in escolhas.items() if re.search(rc, txt(cl["nome"]))), r"^total$")
        cat = next((c for c in cats if re.search(pad, txt(c["nome"]))), None)
        if cat is None and pad == r"^total$" and cats:
            cat = cats[0]
            print(f"    (classificação '{cl['nome']}' sem categoria Total; usando '{cat['nome']}')")
        if cat is None:
            raise RuntimeError(f"Categoria /{pad}/ não encontrada na classificação '{cl['nome']}'.")
        partes.append(f"/c{cl['id']}/{cat['id']}")
    return "".join(partes)


def sidra(caminho):
    dados = get(f"{SIDRA}/{caminho}").json()
    cab, linhas = dados[0], dados[1:]
    if not linhas:
        return pd.DataFrame()
    df = pd.DataFrame(linhas)
    ren = {}
    for k, rot in cab.items():
        if rot.startswith("Município (Código)"):
            ren[k] = "cod"
        elif rot.startswith("Variável (Código)"):
            ren[k] = "cod_var"
        elif rot.startswith("Ano (Código)"):
            ren[k] = "ano"
        elif k == "V":
            ren[k] = "valor"
    df = df.rename(columns=ren)
    df["valor"] = num(df["valor"])
    df["cod"] = df["cod"].astype(int)
    return df


def valores(tab, m, vars_ids, periodo, escolhas=None):
    """Consulta N6 (municípios) para as variáveis e o período; devolve DataFrame largo (cod x variável)."""
    v = ",".join(str(i) for i in vars_ids)
    d = sidra(f"t/{tab}/n6/all/v/{v}/p/{periodo}{caminho_classes(m, escolhas)}")
    if d.empty:
        raise RuntimeError(f"Tabela {tab}: consulta sem valores (período {periodo}).")
    return d.pivot_table(index="cod", columns="cod_var", values="valor", aggfunc="first")


# ------------------------------------------------------------------ IBGE: base de municípios
def tabela_ibge():
    r = get(URL_IBGE_MUN, timeout=120)
    linhas = []
    for m in r.json():
        uf = (m.get("microrregiao") or {}).get("mesorregiao", {}).get("UF", {}).get("sigla") \
            or m["regiao-imediata"]["regiao-intermediaria"]["UF"]["sigla"]
        linhas.append({"codigo_ibge_7": m["id"], "municipio": m["nome"], "uf": uf,
                       "chave": f"{uf}|{norm(m['nome'])}"})
    return pd.DataFrame(linhas)


# ================================================================== BLOCO 1: CENSO
def bloco_censo():
    fm = {}
    out = pd.DataFrame()

    # --- 2022: população, área, densidade (tabela 4714) ---
    m = meta(4714)
    v = acha_vars(m, {"pop": r"^populacao residente", "area": r"area da unidade territorial",
                      "dens": r"densidade demografica"})
    print("  4714 ->", {k: f"{i} ({n})" for k, (i, n) in v.items()})
    faltam = {"pop", "area", "dens"} - set(v)
    if faltam:
        raise RuntimeError(f"Tabela 4714: variáveis não encontradas: {sorted(faltam)}")
    w = valores(4714, m, [i for i, _ in v.values()], ANO_CENSO_ATUAL)
    w.columns = [{str(i): k for k, (i, _) in v.items()}[str(c)] for c in w.columns]
    out["População Censo 2022"] = w["pop"]
    out["Área territorial (km²)"] = w["area"]
    out["Densidade demográfica 2022 (hab/km²)"] = w["dens"]
    fm.update({"População Censo 2022": "#,##0", "Área territorial (km²)": "#,##0.00",
               "Densidade demográfica 2022 (hab/km²)": "#,##0.00"})

    # --- 2010: população total e urbana (tabela 202) ---
    try:
        m2 = meta(202)
        v2 = acha_vars(m2, {"pop": r"^populacao residente"})
        vid = v2["pop"][0]
        tot = valores(202, m2, [vid], ANO_CENSO_ANT, {"situacao": r"^total$", "sexo": r"^total$"})
        urb = valores(202, m2, [vid], ANO_CENSO_ANT, {"situacao": r"^urbana$", "sexo": r"^total$"})
        out["População Censo 2010"] = tot.iloc[:, 0]
        out["População urbana Censo 2010"] = urb.iloc[:, 0]
        fm.update({"População Censo 2010": "#,##0", "População urbana Censo 2010": "#,##0"})
        print("  202 -> população 2010 total e urbana OK")
    except Exception as e:
        print(f"  Aviso: população 2010 (tabela 202) indisponível: {e}")

    # --- 2022: população urbana (tabela 9923) ---
    try:
        m3 = meta(9923)
        v3 = acha_vars(m3, {"pop": r"^populacao residente(?!.*%)(?!.*percent)"})
        vid3 = v3["pop"][0]
        urb22 = valores(9923, m3, [vid3], ANO_CENSO_ATUAL, {"situacao": r"^urbana$"})
        out["População urbana Censo 2022"] = urb22.iloc[:, 0]
        fm["População urbana Censo 2022"] = "#,##0"
        print("  9923 -> população urbana 2022 OK")
    except Exception as e:
        print(f"  Aviso: população urbana 2022 (tabela 9923) indisponível: {e}")

    # --- derivados ---
    if "População Censo 2010" in out:
        p10, p22 = out["População Censo 2010"], out["População Censo 2022"]
        out["Crescimento populacional 2010-2022"] = (p22 / p10.where(p10 > 0)) - 1
        out["Crescimento populacional anual 2010-2022 (CAGR)"] = (p22 / p10.where(p10 > 0)) ** (1 / 12) - 1
        out["Grau de urbanização 2010"] = out["População urbana Censo 2010"] / p10.where(p10 > 0)
        fm.update({"Crescimento populacional 2010-2022": "0.0%",
                   "Crescimento populacional anual 2010-2022 (CAGR)": "0.00%", "Grau de urbanização 2010": "0.0%"})
    if "População urbana Censo 2022" in out:
        p22 = out["População Censo 2022"]
        out["Grau de urbanização 2022"] = out["População urbana Censo 2022"] / p22.where(p22 > 0)
        fm["Grau de urbanização 2022"] = "0.0%"
    out.index.name = "codigo_ibge_7"
    return out.reset_index(), fm


# ================================================================== BLOCO 2: CEMPRE
def bloco_cempre():
    fm = {}
    padroes = {"empresas": r"^numero de empresas", "ul": r"unidades locais",
               "po": r"pessoal ocupado total", "assal": r"pessoal ocupado assalariado",
               "sal": r"salarios e outras remuneracoes",
               "sm_sm": r"salario medio mensal.*(minimo)", "sm_rs": r"salario medio mensal.*(r\$|reais)"}
    escolhida = None
    for tab in (6450, 1685):
        try:
            m = meta(tab)
        except Exception as e:
            print(f"  CEMPRE {tab}: metadados indisponíveis ({e})")
            continue
        if "N6" not in str(m.get("nivelTerritorial", "")):
            print(f"  CEMPRE {tab}: sem nível municipal; ignorada.")
            continue
        v = acha_vars(m, padroes)
        if {"assal", "sal"} <= set(v):
            escolhida = (tab, m, v)
            break
    if not escolhida:
        raise RuntimeError("Nenhuma tabela do CEMPRE com pessoal assalariado e salários por município.")
    tab, m, v = escolhida
    print(f"  CEMPRE tabela {tab}; variáveis:", {k: f"{i} ({n})" for k, (i, n) in v.items()})
    rev = {str(i): k for k, (i, _) in v.items()}
    ids = [i for i, _ in v.values()]

    anos_cem = anos_disponiveis(tab)
    ano_cem = ANO_CEMPRE or (max(anos_cem) if anos_cem else None)
    print(f"  Anos do CEMPRE publicados: {anos_cem[-4:] if anos_cem else '?'} | usando: {ano_cem or 'last'}")
    periodo_cem = ano_cem or "last"
    w = valores(tab, m, ids, periodo_cem)
    w.columns = [rev[str(c)] for c in w.columns]
    rot = {"empresas": "Empresas atuantes (CEMPRE)", "ul": "Unidades locais (CEMPRE)",
           "po": "Pessoal ocupado total (CEMPRE)", "assal": "Pessoal ocupado assalariado (CEMPRE)",
           "sal": "Salários e outras remunerações (R$ mil/ano)",
           "sm_sm": "Salário médio mensal (salários mínimos)", "sm_rs": "Salário médio mensal (R$)"}
    out = pd.DataFrame(index=w.index)
    out["Ano do CEMPRE"] = ano_cem
    fm["Ano do CEMPRE"] = "0"
    for k, nome in rot.items():
        if k in w:
            out[nome] = w[k]
            fm[nome] = "#,##0.00" if k.startswith("sm") else "#,##0"
    out["Massa salarial mensal média (R$ mil)"] = out["Salários e outras remunerações (R$ mil/ano)"] / 12
    fm["Massa salarial mensal média (R$ mil)"] = "#,##0"

    # --- administração pública (seção O) para separar a massa salarial do setor privado ---
    try:
        esc = {r"cnae|atividade|secao|classific": r"administracao publica"}
        wa = valores(tab, m, [v["assal"][0], v["sal"][0]], periodo_cem, esc)
        wa.columns = [rev[str(c)] for c in wa.columns]
        out["Assalariados na administração pública"] = wa["assal"]
        out["Salários na administração pública (R$ mil/ano)"] = wa["sal"]
        out["Salários exc. administração pública (R$ mil/ano)"] = (
            out["Salários e outras remunerações (R$ mil/ano)"] - wa["sal"])
        out["% da massa salarial na administração pública"] = wa["sal"] / out[
            "Salários e outras remunerações (R$ mil/ano)"].where(
            out["Salários e outras remunerações (R$ mil/ano)"] > 0)
        fm.update({"Assalariados na administração pública": "#,##0",
                   "Salários na administração pública (R$ mil/ano)": "#,##0",
                   "Salários exc. administração pública (R$ mil/ano)": "#,##0",
                   "% da massa salarial na administração pública": "0.0%"})
        print("  CEMPRE: administração pública separada OK")
    except Exception as e:
        print(f"  Aviso: não foi possível separar a administração pública ({e})")
    out.index.name = "codigo_ibge_7"
    return out.reset_index(), fm


# ================================================================== BLOCO 3: FROTA POR IDADE
MES_NUM = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
           "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}


def links_da_pagina(ano):
    """[(ano, mes, texto_normalizado, url)] - cada link associado ao título '(Mês de AAAA)' mais próximo acima dele."""
    html = get(f"{BASE_SENATRAN}/conteudo-Senatran/frota-de-veiculos-{ano}", timeout=120).text
    titulos = []
    for m in re.finditer(r"\(([A-Za-zÀ-ÿ]+) de (20\d\d)\)", html):
        mes = MES_NUM.get(norm(m.group(1)).lower().replace("_", ""))
        if mes:
            titulos.append((m.start(), mes, int(m.group(2))))
    saida = []
    for a_ in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, flags=re.S | re.I):
        ant = [t for t in titulos if t[0] < a_.start()]
        if not ant:
            continue
        url = a_.group(1)
        url = url if url.startswith("http") else "https://www.gov.br" + url
        saida.append((ant[-1][2], ant[-1][1], norm(re.sub(r"<[^>]+>", "", a_.group(2))), url))
    return saida


def acha_url_frota_idade(ano=None, mes=None):
    """(ano, mes, url) do arquivo 'Quantidade de Veículos por UF Município Ano de Fabricação Modelo'; sem ano/mês = o mais recente."""
    if ano and mes and (ano, mes) in URLS_FROTA_IDADE:
        return ano, mes, URLS_FROTA_IDADE[(ano, mes)]
    from datetime import date
    anos = [ano] if ano else [date.today().year, date.today().year - 1]
    cand = []
    for a_ in anos:
        try:
            cand += [c for c in links_da_pagina(a_) if "ANO_DE_FABRICACAO" in c[2]]
        except Exception as e:
            print(f"  Aviso: página de {a_} indisponível ({e})")
    if mes:
        cand = [c for c in cand if c[1] == mes]
    if not cand:
        raise RuntimeError("Link da frota por ano de fabricação não encontrado na página da SENATRAN.")
    a_, m_, _, url = max(cand, key=lambda c: (c[0], c[1]))
    return a_, m_, re.sub(r"/@@download/file$", "", url)


def baixa_xlsx(url):
    for u in (url, url + "/@@download/file"):
        try:
            r = get(u, tentativas=2, timeout=300)
        except RuntimeError as e:
            print("   ", e)
            continue
        if r.content.startswith(b"PK"):
            return r.content
        print(f"    {u}: resposta não é um .xlsx")
    raise RuntimeError("Não foi possível baixar a planilha de frota por ano de fabricação.")


def le_planilha(conteudo):
    xls = pd.read_excel(io.BytesIO(conteudo), sheet_name=None, header=None, dtype=object)
    for nome, bruto in xls.items():
        for i in range(min(len(bruto), 30)):
            cel = [norm(x) for x in bruto.iloc[i].tolist() if pd.notna(x)]
            if "UF" in cel and any(c.startswith("MUNICIPIO") for c in cel):
                df = bruto.iloc[i + 1:].copy()
                df.columns = [norm(x) if pd.notna(x) else f"SEM_NOME_{j}" for j, x in enumerate(bruto.iloc[i])]
                df = df.loc[:, ~df.columns.duplicated()].dropna(how="all")
                print(f"  Aba '{nome}', cabeçalho na linha {i + 1}; colunas: {list(df.columns)[:12]}")
                return df
    raise RuntimeError("Cabeçalho (UF / MUNICIPIO) não encontrado na planilha.")


def para_longo(df):
    """Devolve DataFrame (uf, mun, ano, qtd), aceitando formato longo (coluna de ano) ou largo (uma coluna por ano)."""
    col_mun = next(c for c in df.columns if c.startswith("MUNICIPIO"))
    anos_wide = [c for c in df.columns if re.fullmatch(r"(ANO_)?(19|20)\d\d", c)]
    if len(anos_wide) >= 5:
        print(f"  Formato largo: {len(anos_wide)} colunas de ano")
        lg = df.melt(id_vars=["UF", col_mun], value_vars=anos_wide, var_name="ano", value_name="qtd")
        lg["ano"] = lg["ano"].map(lambda x: int(re.search(r"(19|20)\d\d", x).group()))
    else:
        cand = [c for c in df.columns if c.startswith("ANO")]
        col_ano = next((c for c in cand if "FAB" in c), cand[0] if cand else None)
        if col_ano is None:
            raise RuntimeError(f"Coluna de ano de fabricação não encontrada. Colunas: {list(df.columns)}")
        numericas = [c for c in df.columns if c not in ("UF", col_mun, col_ano)
                     and num(df[c]).notna().mean() > 0.9]
        pref = [c for c in numericas if re.search(r"QTD|QUANT|TOTAL|FROTA|VEICUL", c)]
        col_qtd = (pref or numericas or [None])[0]
        if col_qtd is None:
            raise RuntimeError(f"Coluna de quantidade não encontrada. Colunas: {list(df.columns)}")
        print(f"  Formato longo: ano='{col_ano}', quantidade='{col_qtd}'")
        lg = pd.DataFrame({"UF": df["UF"], col_mun: df[col_mun], "ano": num(df[col_ano]), "qtd": df[col_qtd]})
    lg = lg.rename(columns={"UF": "uf", col_mun: "mun"})
    lg["qtd"] = num(lg["qtd"]).fillna(0)
    lg["uf"] = lg["uf"].astype(str).str.strip().map(lambda x: x.upper() if len(x) == 2 else UFS.get(norm(x), x))
    lg["mun"] = lg["mun"].astype(str).str.strip()
    lg = lg[~lg["mun"].map(norm).isin(["TOTAL", "TOTAIS", "NAN"])]
    return lg


def parear(uniq, ibge):
    """uniq: DataFrame(uf, mun). Devolve com codigo_ibge_7 (exato -> manual -> aproximado)."""
    u = uniq.copy()
    u["chave"] = u["uf"] + "|" + u["mun"].map(norm)
    u = u.merge(ibge[["chave", "codigo_ibge_7"]].drop_duplicates("chave"), on="chave", how="left")
    for idx in u.index[u["codigo_ibge_7"].isna()]:
        k = (u.at[idx, "uf"], norm(u.at[idx, "mun"]))
        if k in DE_PARA_MANUAL:
            u.at[idx, "codigo_ibge_7"] = int(DE_PARA_MANUAL[k])
    usados = set(u["codigo_ibge_7"].dropna().astype(int))
    for idx in u.index[u["codigo_ibge_7"].isna()]:
        cand = ibge[(ibge["uf"] == u.at[idx, "uf"]) & (~ibge["codigo_ibge_7"].isin(usados))].copy()
        if cand.empty:
            continue
        cand["n"] = cand["municipio"].map(norm)
        par = difflib.get_close_matches(norm(u.at[idx, "mun"]), cand["n"].tolist(), n=1, cutoff=0.86)
        if par:
            cod = int(cand[cand["n"] == par[0]].iloc[0]["codigo_ibge_7"])
            u.at[idx, "codigo_ibge_7"] = cod
            usados.add(cod)
    return u


def bloco_frota_idade(ibge):
    fm = {}
    ano_f, mes_f, url = acha_url_frota_idade(ANO_REF, MES_FROTA)
    ref = ano_f
    print(f"  Mês da frota por idade: {mes_f:02d}/{ano_f}\n  Baixando:", url)
    lg = para_longo(le_planilha(baixa_xlsx(url)))
    g = lg.groupby(["uf", "mun", "ano"], as_index=False)["qtd"].sum()
    valido = g["ano"].between(1900, ref + 1)
    g["idade"] = (ref - g["ano"]).clip(lower=0)
    print(f"  {len(g):,} linhas (município x ano). Sem ano válido: "
          f"{g.loc[~valido, 'qtd'].sum() / max(g['qtd'].sum(), 1):.2%} da frota".replace(",", "."))

    # pareia por (UF, nome) e agrega por município
    pares = parear(g[["uf", "mun"]].drop_duplicates(), ibge)
    g = g.merge(pares[["uf", "mun", "codigo_ibge_7"]], on=["uf", "mun"], how="left")
    sem = g[g["codigo_ibge_7"].isna()]
    print(f"  Municípios pareados: {pares['codigo_ibge_7'].notna().sum()} | sem par: {int(pares['codigo_ibge_7'].isna().sum())}"
          f" ({sem['qtd'].sum() / max(g['qtd'].sum(), 1):.3%} da frota)")
    g = g[g["codigo_ibge_7"].notna()].copy()
    g["codigo_ibge_7"] = g["codigo_ibge_7"].astype(int)
    gv = g[g["ano"].between(1900, ref + 1)].copy()
    gv["faixa"] = pd.cut(gv["idade"], [-1, 5, 10, 15, 20, 500],
                         labels=["ate5", "6a10", "11a15", "16a20", "mais20"])

    tot = g.groupby("codigo_ibge_7")["qtd"].sum()
    tv = gv.groupby("codigo_ibge_7")["qtd"].sum()
    out = pd.DataFrame({"Frota total (arquivo por ano de fabricação)": tot})
    out["Mês da frota por idade (AAAAMM)"] = f"{ano_f}{mes_f:02d}"
    out["Idade média da frota (anos)"] = (gv["idade"] * gv["qtd"]).groupby(gv["codigo_ibge_7"]).sum() / tv.where(tv > 0)
    gs = gv.sort_values(["codigo_ibge_7", "idade"]).copy()          # mediana ponderada, sem groupby.apply
    gs["acum"] = gs.groupby("codigo_ibge_7")["qtd"].cumsum()
    gs["soma"] = gs.groupby("codigo_ibge_7")["qtd"].transform("sum")
    out["Idade mediana da frota (anos)"] = gs[gs["acum"] >= gs["soma"] / 2].groupby("codigo_ibge_7")["idade"].first()
    faixas = gv.pivot_table(index="codigo_ibge_7", columns="faixa", values="qtd", aggfunc="sum",
                            observed=False).fillna(0)
    for c, rot in [("ate5", "% da frota com até 5 anos"), ("6a10", "% da frota com 6 a 10 anos"),
                   ("11a15", "% da frota com 11 a 15 anos"), ("16a20", "% da frota com 16 a 20 anos"),
                   ("mais20", "% da frota com mais de 20 anos")]:
        out[rot] = faixas[c] / tv.where(tv > 0)
        fm[rot] = "0.0%"
    ge10 = gv[gv["idade"] >= 10].groupby("codigo_ibge_7")["qtd"].sum()
    ge15 = gv[gv["idade"] >= 15].groupby("codigo_ibge_7")["qtd"].sum()
    out["% da frota com 10 anos ou mais (renovável)"] = (ge10 / tv.where(tv > 0)).reindex(out.index).fillna(0)
    out["% da frota com 15 anos ou mais"] = (ge15 / tv.where(tv > 0)).reindex(out.index).fillna(0)
    out["Frota com 10 anos ou mais (unidades)"] = ge10.reindex(out.index).fillna(0)
    out["% da frota sem ano de fabricação válido"] = 1 - (tv / tot.where(tot > 0)).reindex(out.index)
    fm.update({"Mês da frota por idade (AAAAMM)": "@", "Frota total (arquivo por ano de fabricação)": "#,##0", "Idade média da frota (anos)": "0.0",
               "Idade mediana da frota (anos)": "0", "% da frota com 10 anos ou mais (renovável)": "0.0%",
               "% da frota com 15 anos ou mais": "0.0%", "Frota com 10 anos ou mais (unidades)": "#,##0",
               "% da frota sem ano de fabricação válido": "0.0%"})
    out.index.name = "codigo_ibge_7"
    return out.reset_index(), fm


# ================================================================== saída
def salva(base, blocos, status, caminho):
    df = base.copy()
    fmts = {}
    for d, fm in blocos:
        df = df.merge(d, on="codigo_ibge_7", how="left", validate="one_to_one")
        fmts.update(fm)
    df = df.rename(columns={"codigo_ibge_7": "Código IBGE (7 dígitos)", "municipio": "Município", "uf": "UF"})
    df = df.drop(columns=["chave"], errors="ignore")
    fonte = Font(name="Arial", size=10)
    head = (Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"),
            Alignment(wrap_text=True, vertical="center", horizontal="center"))
    st = pd.DataFrame(status)
    notas = pd.DataFrame({"Notas": [
        "Chave: Código IBGE de 7 dígitos; base de municípios = API de localidades do IBGE (pode ter 1 município a mais que o SIDRA).",
        f"Censo: tabelas SIDRA 4714 (2022), 202 (2010) e 9923 (urbana 2022). Crescimento = pop. 2022 / pop. 2010 - 1; CAGR em 12 anos.",
        "CEMPRE (ano na coluna 'Ano do CEMPRE'): massa salarial = 'Salários e outras remunerações' (anual, R$ mil), de TODOS os assalariados formais, "
        "inclusive administração pública; as colunas 'exc. administração pública' isolam o setor privado.",
        "CEMPRE tem ocultações por sigilo estatístico em municípios muito pequenos (valores vazios).",
        "Frota por idade: arquivo SENATRAN 'UF, Município, Ano de Fabricação e Modelo' (mês na coluna 'Mês da frota por idade'); "
        "idade = ano do arquivo - ano de fabricação. O arquivo NÃO separa por tipo de veículo: a idade vale para a frota total.",
        "'Frota renovável' = parcela com 10+ anos (critério convencional; ajuste conforme o segmento).",
        "Valores vazios = município não encontrado/oculto na fonte (não são zero).",
    ]})
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Dados", index=False)
        ws = w.sheets["Dados"]
        for j, c in enumerate(df.columns, 1):
            h = ws.cell(row=1, column=j)
            h.font, h.fill, h.alignment = head
            ws.column_dimensions[get_column_letter(j)].width = 30 if c == "Município" else 20
            fm = fmts.get(c) or ("0" if c.startswith("Código") else None)
            for i in range(2, len(df) + 2):
                cel = ws.cell(row=i, column=j)
                cel.font = fonte
                if fm:
                    cel.number_format = fm
        ws.row_dimensions[1].height = 55
        ws.freeze_panes = "D2"
        ws.auto_filter.ref = ws.dimensions
        for nome, tab, larg in [("Diagnóstico", st, 40), ("Fontes e notas", notas, 150)]:
            tab.to_excel(w, sheet_name=nome, index=False)
            s = w.sheets[nome]
            for j in range(1, len(tab.columns) + 1):
                s.cell(row=1, column=j).font, s.cell(row=1, column=j).fill, s.cell(row=1, column=j).alignment = head
                s.column_dimensions[get_column_letter(j)].width = larg
                for i in range(2, len(tab) + 2):
                    s.cell(row=i, column=j).font = fonte
    print(f"\nArquivo salvo: {caminho} ({len(df)} municípios x {df.shape[1]} colunas)")


def main():
    print("Baixando a lista de municípios do IBGE...")
    ibge = tabela_ibge()
    base = ibge[["codigo_ibge_7", "municipio", "uf"]].copy()
    blocos, status = [], []
    for nome, fn in [("Censo (área, densidade, crescimento, urbanização)", bloco_censo),
                     ("CEMPRE (massa salarial formal)", bloco_cempre),
                     ("Frota por idade (SENATRAN)", lambda: bloco_frota_idade(ibge))]:
        print(f"\n== {nome}")
        try:
            d, fm = fn()
            blocos.append((d, fm))
            cols = [c for c in d.columns if c != "codigo_ibge_7"]
            cobertura = int(d[cols].notna().any(axis=1).sum())
            status.append({"Bloco": nome, "Situação": "OK", "Municípios com dado": cobertura,
                           "Colunas": len(cols), "Detalhe": ""})
            print(f"  OK: {cobertura} municípios, {len(cols)} colunas")
        except Exception as e:
            status.append({"Bloco": nome, "Situação": "ERRO", "Municípios com dado": 0, "Colunas": 0,
                           "Detalhe": str(e)[:300]})
            print(f"  ERRO neste bloco: {e}")
    print("\nResumo:", [(s["Bloco"].split(" (")[0], s["Situação"]) for s in status])
    salva(base, blocos, status, SAIDA)


if __name__ == "__main__":
    main()
# %%
