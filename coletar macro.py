#%%!/usr/bin/env python3
"""
Coleta os dados da VISÃO MACRO (nível UF / nacional) do IPCM e salva em XLSX + CSVs.

Blocos (cada um roda isolado; falha em um não derruba os outros; resultado na aba "Diagnóstico"):
  bcb     Cotas ativas (BCB, OData PANORAMA_DE_CONSORCIOS): série nacional, por segmento e por UF
  share   Market share por administradora (BCB, arquivo AAAAMMConsorcios.zip)  [endereço do ZIP a confirmar]
  sgs     Selic (432 meta, 4390 acumulada no mês) e dólar (série 1) - SGS/BCB
  ibge    UF: código, população e PIB total / per capita (SIDRA)
  trends  Google Trends: "consórcio", "consórcio de carros", "financiamento de veículos"
  fipe    Tabela FIPE: preços de uma CESTA de modelos + histórico (API v2 parallelum)
  abac    ABAC: indicadores setoriais extraídos de páginas públicas (SEMIAUTOMÁTICO - conferir)

Uso:
    python coletar_macro.py                 # todos os blocos
    python coletar_macro.py bcb sgs ibge    # só alguns

Requisitos: pip install requests pandas openpyxl
Opcionais : pip install pytrends            (Trends; o pacote foi arquivado em abr/2025 e pode falhar com HTTP 429)
Saídas    : dados_macro.xlsx, pasta macro_csv/ (um CSV por aba)
"""
import io
import os
import re
import sys
import time
import unicodedata
import zipfile
from datetime import date, datetime

import numpy as np
import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================================ CONFIGURAÇÃO
INICIO_ANO = 2009                      # início das séries (o Panorama de consórcios começa em 2009)
SAIDA_XLSX, PASTA_CSV = "dados_macro.xlsx", "macro_csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (coleta-macro)"}

# BCB - Panorama de consórcios (OData)
OLINDA = "https://olinda.bcb.gov.br/olinda/servico/PANORAMA_DE_CONSORCIOS/versao/v1/odata"
IDS_NACIONAL = [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,      # grupos ativos e cotas ativas por segmento
                28, 31, 34, 54, 55, 56, 57, 58, 60]                      # contempladas e comercializadas (12 meses)
IDS_UF = list(range(99, 126))                                           # cotas ativas por estado (99 a 125)

# BCB - base de dados de consórcios (market share por administradora)
URLS_ZIP_CONSORCIOS = [   # candidatos (o endereço exato NÃO foi confirmado); {p} = AAAAMM
    "https://www.bcb.gov.br/content/estabilidadefinanceira/consorciobd/{p}Consorcios.zip",
    "https://www.bcb.gov.br/content/estabilidadefinanceira/consorcios/{p}Consorcios.zip",
    "https://www.bcb.gov.br/fis/consorcios/port/arquivos/{p}Consorcios.zip",
    "https://www.bcb.gov.br/fis/consorcios/Arquivos/{p}Consorcios.zip",
]
PAGINAS_BCB_BD = ["https://www.bcb.gov.br/estabilidadefinanceira/consorciobd",
                  "https://www.bcb.gov.br/fis/consorcios/port/consorcio_banco_de_dados.asp?idpai=consorcio&frame=1"]
PASTA_ZIPS_MANUAL = "consorcios_bcb_zips"      # se baixar os ZIPs na mão, coloque aqui (AAAAMMConsorcios.zip)
MESES_SHARE = 12                               # quantos meses recentes baixar para o market share

# SGS
SERIES_SGS = {"selic_meta_pct_aa": 432, "selic_acum_mes_pct": 4390, "dolar_venda_brl": 1}

# Google Trends
TERMOS_TRENDS = ["consórcio", "consórcio de carros", "financiamento de veículos"]
TRENDS_POR_UF = False                          # True = também interesse por UF (27 requisições; sujeito a bloqueio)
PASTA_TRENDS_CSV = "trends_csv"                # alternativa manual: CSVs exportados em trends.google.com

# FIPE
FIPE_BASE = "https://fipe.parallelum.com.br/api/v2"
FIPE_TOKEN = os.environ.get("FIPE_TOKEN", "")  # opcional (1.000 req/dia com token gratuito)
FIPE_MAX_REQ = 900
FIPE_MAX_VERSOES = 2                           # versões (modelos) por item da cesta
FIPE_CESTA = {                                 # tipo -> [(regex da marca, regex do modelo)]
    "cars": [("fiat", r"strada"), ("fiat", r"mobi"), ("gm|chevrolet", r"onix"), ("hyundai", r"hb20"),
             ("vw|volks", r"polo"), ("vw|volks", r"t-cross"), ("toyota", r"corolla"), ("jeep", r"compass"),
             ("renault", r"kwid"), ("honda", r"hr-v")],
    "motorcycles": [("honda", r"cg 160"), ("honda", r"biz"), ("yamaha", r"fazer"), ("honda", r"pcx")],
    "trucks": [("vw|volks", r"delivery"), ("mercedes", r"accelo"), ("iveco", r"daily")],
}

# ABAC
ABAC_SEMENTES = ["https://blog.abac.org.br/drops-de-mercado/", "https://abac.org.br/sistema-downloads/"]
ABAC_MAX_PAGINAS = 40

UFS = {  # sigla -> nome (para identificar as métricas por UF)
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas", "BA": "Bahia", "CE": "Ceará",
    "DF": "Distrito Federal", "ES": "Espírito Santo", "GO": "Goiás", "MA": "Maranhão", "MT": "Mato Grosso",
    "MS": "Mato Grosso do Sul", "MG": "Minas Gerais", "PA": "Pará", "PB": "Paraíba", "PR": "Paraná",
    "PE": "Pernambuco", "PI": "Piauí", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte",
    "RS": "Rio Grande do Sul", "RO": "Rondônia", "RR": "Roraima", "SC": "Santa Catarina", "SP": "São Paulo",
    "SE": "Sergipe", "TO": "Tocantins"}


# ============================================================================ utilitários
def norm(t):
    s = unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").upper()


def slug(t):
    return norm(t).lower()


def http_get(url, params=None, headers=None, tentativas=4, timeout=90):
    h = dict(HEADERS)
    h.update(headers or {})
    ultimo = None
    for i in range(tentativas):
        try:
            r = requests.get(url, params=params, headers=h, timeout=timeout)
            if r.status_code in (429, 500, 502, 503, 504):
                raise requests.HTTPError(f"HTTP {r.status_code}")
            return r
        except requests.RequestException as e:
            ultimo = e
            time.sleep(min(2 ** (i + 1), 30))
    raise RuntimeError(f"Falha em {url}: {ultimo}")


def num_br(s):
    """'1.234,56' / 'R$ 1.234,56' -> 1234.56 ; aceita número."""
    if s is None:
        return np.nan
    if isinstance(s, (int, float)):
        return float(s)
    t = re.sub(r"[^\d,.\-]", "", str(s))
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):          # '1.000' / '12.345.678' = milhar (formato brasileiro)
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return np.nan


def meses_ym(ini, fim):
    out, a, m = [], int(ini[:4]), int(ini[4:])
    while f"{a}{m:02d}" <= fim:
        out.append(f"{a}{m:02d}")
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)
    return out


# ============================================================================ BLOCO bcb
def uf_da_metrica(nome):
    n = str(nome)
    for sig in UFS:
        if re.search(rf"(?<![A-Za-z]){sig}(?![A-Za-z])", n):
            return sig
    nn = norm(n)
    for sig, nom in sorted(UFS.items(), key=lambda kv: -len(kv[1])):    # nomes longos primeiro (MATO_GROSSO_DO_SUL)
        if norm(nom) in nn:
            return sig
    return None


def odata_metricas(ym):
    url = f"{OLINDA}/Metricas(DataBase=@DataBase)"
    params = {"@DataBase": ym, "$format": "json", "$top": 200}
    r = http_get(url, params=params)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code} para DataBase={ym}")
    js = r.json()
    if js.get("@odata.nextLink"):
        raise RuntimeError(f"Resposta truncada para {ym} (nextLink); aumente $top.")
    return js.get("value", [])


def bloco_bcb():
    hoje = date.today()
    fim = f"{hoje.year}{hoje.month:02d}"
    linhas, publicados = [], []
    print(f"  Consultando o Panorama de Consórcios de {INICIO_ANO}01 a {fim} (uma requisição por mês)...")
    for ym in meses_ym(f"{INICIO_ANO}01", fim):
        vals = odata_metricas(ym)
        if vals:
            publicados.append(ym)
            for v in vals:
                linhas.append({"data_base": int(v["DataBase"]), "id_metrica": int(v["IdMetrica"]),
                               "grupo": v.get("Grupo"), "metrica": v.get("Metrica"),
                               "valor": pd.to_numeric(v.get("Valor"), errors="coerce"), "unidade": v.get("Unidade")})
        time.sleep(0.15)
    if not linhas:
        raise RuntimeError("Nenhum mês com dados no Panorama de Consórcios.")
    longo = pd.DataFrame(linhas)
    longo["valor_unidades"] = np.where(longo["unidade"].astype(str).str.lower() == "mil",
                                       longo["valor"] * 1000, longo["valor"])
    print(f"  {len(publicados)} meses com dados: {publicados[0]} a {publicados[-1]}; {longo['id_metrica'].nunique()} métricas")

    # --- série nacional (wide) ---
    nac = longo[longo["id_metrica"].isin(IDS_NACIONAL)].copy()
    nac["coluna"] = nac["id_metrica"].astype(str).str.zfill(3) + "_" + nac["metrica"].map(slug)
    nacional = nac.pivot_table(index="data_base", columns="coluna", values="valor_unidades", aggfunc="first").reset_index()
    # --- cotas ativas por UF (long) ---
    uf = longo[longo["id_metrica"].isin(IDS_UF)].copy()
    uf["uf"] = uf["metrica"].map(uf_da_metrica)
    sem_uf = uf[uf["uf"].isna()]["metrica"].unique()
    if len(sem_uf):
        print(f"  Aviso: {len(sem_uf)} métricas por UF sem sigla identificada (ex.: {list(sem_uf)[:3]})")
    uf_ok = uf[uf["uf"].notna()][["data_base", "uf", "id_metrica", "metrica", "grupo", "valor_unidades"]].rename(
        columns={"valor_unidades": "cotas_ativas"})
    uf_ok["ano"] = uf_ok["data_base"] // 100
    uf_ok["mes"] = uf_ok["data_base"] % 100
    print(f"  UF: {uf_ok['uf'].nunique()} UFs, {uf_ok['data_base'].nunique()} datas-base "
          f"(meses: {sorted(int(m) for m in uf_ok['mes'].unique())})")
    return {"Cotas_Nacional_Mensal": nacional, "Cotas_UF": uf_ok.sort_values(["data_base", "uf"]),
            "BCB_Metricas_Longo": longo}


# ============================================================================ BLOCO share
def _links_zip_na_pagina():
    achados = []
    for pag in PAGINAS_BCB_BD:
        try:
            html = http_get(pag, tentativas=2).text
        except Exception:
            continue
        for h in re.findall(r'href="([^"]+\.zip)"', html, flags=re.I):
            if re.search(r"\d{6}.*consorc|consorc.*\d{6}", h, flags=re.I):
                achados.append(h if h.startswith("http") else "https://www.bcb.gov.br" + h)
    return achados


def baixa_zip_consorcios(ym, links_pagina):
    os.makedirs(PASTA_ZIPS_MANUAL, exist_ok=True)
    local = os.path.join(PASTA_ZIPS_MANUAL, f"{ym}Consorcios.zip")
    if os.path.exists(local) and zipfile.is_zipfile(local):
        return open(local, "rb").read()
    cands = [l for l in links_pagina if ym in l] + [u.format(p=ym) for u in URLS_ZIP_CONSORCIOS]
    for u in cands:
        try:
            r = http_get(u, tentativas=1, timeout=120)
        except Exception:
            continue
        if r.status_code == 200 and r.content[:2] == b"PK":
            open(local, "wb").write(r.content)
            return r.content
    return None


def le_segmentos(zbytes):
    with zipfile.ZipFile(io.BytesIO(zbytes)) as z:
        nome = next((n for n in z.namelist() if re.search(r"segmentos?_?consolidados", n, flags=re.I)
                     and n.lower().endswith(".csv")), None)
        if nome is None:
            raise RuntimeError(f"'Segmentos_Consolidados.csv' não encontrado no ZIP: {z.namelist()}")
        raw = z.read(nome)
    for enc in ("cp1252", "utf-8"):
        try:
            return pd.read_csv(io.StringIO(raw.decode(enc)), sep=";", decimal=",", dtype=str)
        except UnicodeDecodeError:
            continue
    raise RuntimeError("Codificação do CSV não reconhecida.")


def bloco_share():
    hoje = date.today()
    ult = f"{hoje.year}{hoje.month:02d}"
    cand = list(reversed(meses_ym(f"{hoje.year - 3}01", ult)))
    links = _links_zip_na_pagina()
    if links:
        print(f"  Links de ZIP encontrados na página do BCB: {len(links)} (ex.: {links[0]})")
    frames, usados, tentados = [], [], 0
    for ym in cand:
        if len(usados) >= MESES_SHARE:
            break
        tentados += 1
        zb = baixa_zip_consorcios(ym, links)
        if zb is None:
            if tentados > 8 and not usados:
                break                                   # nada baixou: evita 36 tentativas inúteis
            continue
        df = le_segmentos(zb)
        df.columns = [slug(c) for c in df.columns]
        if not usados:
            print("  Colunas do Segmentos_Consolidados:", list(df.columns))
        df["data_base_arquivo"] = int(ym)
        frames.append(df)
        usados.append(ym)
    if not frames:
        raise RuntimeError(
            "Não consegui baixar nenhum 'AAAAMMConsorcios.zip'. Baixe manualmente em "
            "https://www.bcb.gov.br/estabilidadefinanceira/consorciobd (Dados Consolidados) e salve na pasta "
            f"'{PASTA_ZIPS_MANUAL}/' com o nome AAAAMMConsorcios.zip; depois rode de novo (bloco 'share').")
    bruto = pd.concat(frames, ignore_index=True)
    cols = list(bruto.columns)
    c_cnpj = next((c for c in cols if "cnpj" in c), None)
    c_nome = next((c for c in cols if re.search(r"administradora|nome|razao", c) and "cnpj" not in c), None)
    c_seg = next((c for c in cols if "segmento" in c), None)
    c_ativas = [c for c in cols if re.search(r"cota", c) and re.search(r"ativ", c)
                and not re.search(r"contempl|desist|exclu|nao|inadimpl|comercial|quita", c)]
    print(f"  Colunas escolhidas -> cnpj: {c_cnpj} | nome: {c_nome} | segmento: {c_seg} | cotas ativas: {c_ativas}")
    if not (c_nome and c_seg and c_ativas):
        print("  ATENÇÃO: não identifiquei todas as colunas; a aba 'Share_Bruto' traz o arquivo original para ajuste.")
        return {"Share_Bruto": bruto}
    base = pd.DataFrame({"data_base": bruto["data_base_arquivo"], "cnpj": bruto[c_cnpj] if c_cnpj else None,
                         "administradora": bruto[c_nome].str.strip(), "segmento": bruto[c_seg].str.strip(),
                         "cotas_ativas": bruto[c_ativas[0]].map(num_br)})
    tot = base.groupby(["data_base", "segmento"])["cotas_ativas"].transform("sum")
    base["market_share_segmento"] = base["cotas_ativas"] / tot.where(tot > 0)
    tot_g = base.groupby("data_base")["cotas_ativas"].transform("sum")
    base["market_share_total"] = base["cotas_ativas"] / tot_g.where(tot_g > 0)
    top = (base.groupby(["data_base", "administradora"], as_index=False)["cotas_ativas"].sum()
           .sort_values(["data_base", "cotas_ativas"], ascending=[True, False]))
    top["share_total"] = top["cotas_ativas"] / top.groupby("data_base")["cotas_ativas"].transform("sum")
    return {"Share_Administradora": base, "Share_Total_Adm": top, "Share_Bruto": bruto}


# ============================================================================ BLOCO sgs
def sgs_serie(codigo, ini, fim, janela_anos=5):
    """Baixa uma série SGS em janelas (a API recusa janelas muito longas)."""
    partes, a = [], ini.year
    while a <= fim.year:
        di, df_ = date(a, 1, 1), min(date(a + janela_anos - 1, 12, 31), fim)
        r = http_get(f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados",
                     params={"formato": "json", "dataInicial": di.strftime("%d/%m/%Y"),
                             "dataFinal": df_.strftime("%d/%m/%Y")})
        if r.status_code == 200 and r.text.strip().startswith("["):
            partes += r.json()
        elif r.status_code not in (200, 404):
            raise RuntimeError(f"SGS {codigo}: HTTP {r.status_code} ({di:%Y}-{df_:%Y})")
        a += janela_anos
    df = pd.DataFrame(partes)
    if df.empty:
        return df
    df["data"] = pd.to_datetime(df["data"], format="%d/%m/%Y")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    return df.drop_duplicates("data").sort_values("data").reset_index(drop=True)


def bloco_sgs():
    ini, fim = date(INICIO_ANO, 1, 1), date.today()
    dfs = {}
    for nome, cod in SERIES_SGS.items():
        d = sgs_serie(cod, ini, fim)
        print(f"  SGS {cod} ({nome}): {len(d)} observações, {d['data'].min():%Y-%m-%d} a {d['data'].max():%Y-%m-%d}")
        dfs[nome] = d
    # mensal
    mensal = None
    def junta(m, serie):
        return serie if m is None else m.merge(serie, on="mes", how="outer")
    for nome, d in dfs.items():
        d = d.assign(mes=d["data"].dt.to_period("M").dt.to_timestamp())
        if nome == "selic_acum_mes_pct":
            s = d.groupby("mes", as_index=False)["valor"].last().rename(columns={"valor": nome})
        elif nome == "selic_meta_pct_aa":
            s = d.groupby("mes", as_index=False).agg(selic_meta_media_pct_aa=("valor", "mean"),
                                                      selic_meta_fim_mes_pct_aa=("valor", "last"))
        else:
            s = d.groupby("mes", as_index=False).agg(dolar_media=("valor", "mean"), dolar_fim_mes=("valor", "last"),
                                                      dolar_max=("valor", "max"), dolar_min=("valor", "min"))
        mensal = junta(mensal, s)
    mensal = mensal.sort_values("mes").reset_index(drop=True)
    mensal["dolar_var_12m"] = mensal["dolar_media"].pct_change(12)
    diario = dfs["dolar_venda_brl"].rename(columns={"valor": "dolar_venda_brl"})
    return {"Selic_Dolar_Mensal": mensal, "Dolar_Diario": diario[["data", "dolar_venda_brl"]]}


# ============================================================================ BLOCO ibge
SIDRA = "https://apisidra.ibge.gov.br/values"
AGREG = "https://servicodados.ibge.gov.br/api/v3/agregados"


def sidra(caminho):
    d = http_get(f"{SIDRA}/{caminho}", timeout=180).json()
    cab, linhas = d[0], d[1:]
    if not linhas:
        return pd.DataFrame()
    df = pd.DataFrame(linhas)
    ren = {}
    for k, rot in cab.items():
        if rot.endswith("(Código)") and rot.startswith(("Unidade da Federação", "Município", "Brasil")):
            ren[k] = "cod_local"
        elif rot in ("Unidade da Federação", "Município"):
            ren[k] = "nome_local"
        elif rot.startswith("Variável (Código)"):
            ren[k] = "cod_var"
        elif rot == "Variável":
            ren[k] = "nome_var"
        elif rot.startswith("Ano (Código)"):
            ren[k] = "ano"
        elif k == "V":
            ren[k] = "valor"
    df = df.rename(columns=ren)
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df["ano"] = df["ano"].astype(int)
    return df


def bloco_ibge():
    est = http_get("https://servicodados.ibge.gov.br/api/v1/localidades/estados?orderBy=nome").json()
    cad = pd.DataFrame([{"cod_uf": e["id"], "uf": e["sigla"], "nome_uf": e["nome"],
                         "regiao_sigla": e["regiao"]["sigla"], "regiao": e["regiao"]["nome"]} for e in est])
    print(f"  Cód_UF: {len(cad)} UFs")

    # população por UF (estimativas, todos os anos) + Censo 2022
    pop = sidra("t/6579/n3/all/v/9324/p/all")
    pop = pop.rename(columns={"valor": "populacao"})[["cod_local", "ano", "populacao"]].astype({"cod_local": int})
    pop["fonte_pop"] = "estimativa IBGE (6579)"
    cen = sidra("t/4714/n3/all/v/93/p/2022")
    if not cen.empty and 2022 not in set(pop["ano"]):
        c = cen.rename(columns={"valor": "populacao"})[["cod_local", "ano", "populacao"]].astype({"cod_local": int})
        c["fonte_pop"] = "Censo 2022 (4714)"
        pop = pd.concat([pop, c], ignore_index=True)
    print(f"  População UF: anos {pop['ano'].min()}-{pop['ano'].max()}")

    # PIB por UF: tenta N3 direto; senão soma dos municípios
    try:
        pib = sidra("t/5938/n3/all/v/37/p/all")
    except Exception:
        pib = pd.DataFrame()
    origem = "SIDRA 5938 (N3)"
    if pib.empty or pib["valor"].notna().sum() == 0:
        print("  PIB N3 indisponível; somando PIB municipal por UF (pode levar alguns minutos)...")
        origem = "soma do PIB municipal (5938, N6)"
        anos = sorted({int(p["id"][:4]) for p in http_get(f"{AGREG}/5938/periodos").json()})
        partes = []
        for a in anos:
            d = sidra(f"t/5938/n6/all/v/37/p/{a}")
            if d.empty:
                continue
            d["uf"] = d["nome_local"].str.rsplit(" - ", n=1).str[1]
            partes.append(d.groupby(["uf", "ano"], as_index=False)["valor"].sum())
        pib = pd.concat(partes).merge(cad[["uf", "cod_uf"]], on="uf").rename(columns={"cod_uf": "cod_local"})
    pib = pib.rename(columns={"valor": "pib_mil_reais"})[["cod_local", "ano", "pib_mil_reais"]].astype({"cod_local": int})
    pib["fonte_pib"] = origem
    painel = (pop.merge(pib, on=["cod_local", "ano"], how="outer")
                 .merge(cad, left_on="cod_local", right_on="cod_uf", how="left").drop(columns="cod_local"))
    painel["pib_per_capita_reais"] = (painel["pib_mil_reais"] * 1000 / painel["populacao"]).round(2)
    painel = painel.sort_values(["uf", "ano"]).reset_index(drop=True)
    print(f"  PIB UF ({origem}): anos {int(pib['ano'].min())}-{int(pib['ano'].max())}")
    return {"UF_Cadastro": cad, "UF_Pop_PIB_Anual": painel}


# ============================================================================ BLOCO trends
def trends_pytrends():
    from pytrends.request import TrendReq        # ImportError -> cai no fallback de CSV
    pt = TrendReq(hl="pt-BR", tz=180, timeout=(10, 30))
    tf = f"{INICIO_ANO}-01-01 {date.today():%Y-%m-%d}"
    for i in range(5):
        try:
            pt.build_payload(TERMOS_TRENDS, cat=0, timeframe=tf, geo="BR", gprop="")
            df = pt.interest_over_time()
            break
        except Exception as e:
            espera = 15 * (2 ** i)
            print(f"    Trends falhou ({e}); nova tentativa em {espera}s")
            time.sleep(espera)
    else:
        raise RuntimeError("Google Trends recusou as requisições (429/bloqueio).")
    df = df.reset_index().rename(columns={"date": "data"})
    out = {"Trends_Brasil": df}
    if TRENDS_POR_UF:
        linhas = []
        for sig in UFS:
            time.sleep(8)
            pt.build_payload(TERMOS_TRENDS, timeframe=tf, geo=f"BR-{sig}")
            d = pt.interest_over_time().reset_index().rename(columns={"date": "data"})
            d["uf"] = sig
            linhas.append(d)
        out["Trends_UF"] = pd.concat(linhas, ignore_index=True)
    return out


def trends_csv():
    if not os.path.isdir(PASTA_TRENDS_CSV):
        raise RuntimeError(f"Pasta '{PASTA_TRENDS_CSV}' não existe.")
    arqs = [f for f in os.listdir(PASTA_TRENDS_CSV) if f.lower().endswith(".csv")]
    if not arqs:
        raise RuntimeError(f"Nenhum CSV em '{PASTA_TRENDS_CSV}'.")
    dfs = []
    for a in arqs:
        txt = open(os.path.join(PASTA_TRENDS_CSV, a), encoding="utf-8-sig").read().splitlines()
        i = next((k for k, l in enumerate(txt) if re.match(r"^(M[êe]s|Semana|Dia|Month|Week|Day)[,;]", l)), None)
        if i is None:
            print(f"    Aviso: cabeçalho não reconhecido em {a}")
            continue
        sep = ";" if txt[i].count(";") > txt[i].count(",") else ","
        d = pd.read_csv(io.StringIO("\n".join(txt[i:])), sep=sep)
        d = d.rename(columns={d.columns[0]: "data"})
        d.columns = [re.sub(r":\s*\(.*\)$", "", c).strip() for c in d.columns]
        d["data"] = pd.to_datetime(d["data"], errors="coerce")
        for c in d.columns[1:]:
            d[c] = pd.to_numeric(d[c].astype(str).str.replace("<1", "0.5"), errors="coerce")
        dfs.append(d.dropna(subset=["data"]).set_index("data"))
    if not dfs:
        raise RuntimeError("Nenhum CSV de Trends pôde ser lido.")
    df = pd.concat(dfs, axis=1)
    df = df.loc[:, ~df.columns.duplicated()].reset_index()
    return {"Trends_Brasil": df}


def bloco_trends():
    try:
        out = trends_pytrends()
        print("  Google Trends via pytrends OK")
    except Exception as e:
        print(f"  pytrends indisponível/bloqueado ({e}); tentando CSVs exportados em '{PASTA_TRENDS_CSV}/'...")
        try:
            out = trends_csv()
        except Exception as e2:
            raise RuntimeError(
                f"{e2} | Como obter: em trends.google.com/trends/explore escolha 'Brasil', período 'Desde 2004', "
                f"digite os 3 termos juntos (para ficarem na mesma escala), clique em 'Fazer download (CSV)' e salve em "
                f"'{PASTA_TRENDS_CSV}/'. Depois rode: python coletar_macro.py trends") from e2
    t = out["Trends_Brasil"].copy()
    cols = [c for c in t.columns if c not in ("data", "isPartial")]
    t["mes"] = pd.to_datetime(t["data"]).dt.to_period("M").dt.to_timestamp()
    out["Trends_Mensal"] = t.groupby("mes", as_index=False)[cols].mean()
    for c in cols:
        out["Trends_Mensal"][f"{slug(c)}_var_12m"] = out["Trends_Mensal"][c].pct_change(12)
    print(f"  Trends: {len(out['Trends_Mensal'])} meses; termos: {cols}")
    return out


# ============================================================================ BLOCO fipe
class Fipe:
    def __init__(self):
        self.n = 0
        self.h = {"accept": "application/json"}
        if FIPE_TOKEN:
            self.h["X-Subscription-Token"] = FIPE_TOKEN

    def get(self, caminho, params=None):
        if self.n >= FIPE_MAX_REQ:
            raise RuntimeError(f"Orçamento de {FIPE_MAX_REQ} requisições FIPE esgotado (rode de novo amanhã).")
        self.n += 1
        r = http_get(f"{FIPE_BASE}/{caminho}", params=params, headers=self.h, tentativas=3)
        if r.status_code == 429:
            raise RuntimeError("FIPE: limite de requisições atingido (HTTP 429). Use FIPE_TOKEN ou tente amanhã.")
        if r.status_code != 200:
            raise RuntimeError(f"FIPE HTTP {r.status_code} em {caminho}")
        return r.json()


def _lista_historico(js):
    """Procura, na resposta do endpoint /history, a lista de {mês, preço} (formato tolerante)."""
    if isinstance(js, list):
        cand = js
    else:
        cand = next((v for v in js.values() if isinstance(v, list) and v and isinstance(v[0], dict)), [])
    saida = []
    for it in cand:
        k_preco = next((k for k in it if re.search(r"price|valor|preco", k, flags=re.I)), None)
        k_mes = next((k for k in it if re.search(r"month|mes|reference|refer", k, flags=re.I)), None)
        if k_preco:
            saida.append({"referencia": it.get(k_mes), "preco": num_br(it[k_preco])})
    return saida


def bloco_fipe():
    api = Fipe()
    refs = api.get("references")
    ref_atual = refs[0] if refs else {}
    print(f"  Referência mais recente da FIPE: {ref_atual}")
    cod_ref = str(ref_atual.get("code", "")) or None
    par = {"reference": cod_ref} if cod_ref else None
    precos, hist = [], []
    for tipo, itens in FIPE_CESTA.items():
        marcas = api.get(f"{tipo}/brands", par)
        for rx_marca, rx_modelo in itens:
            bm = [m for m in marcas if re.search(rx_marca, str(m.get("name", "")), flags=re.I)]
            for b in bm[:1]:
                modelos = api.get(f"{tipo}/brands/{b['code']}/models", par)
                modelos = modelos.get("models", modelos) if isinstance(modelos, dict) else modelos
                sel = [m for m in modelos if re.search(rx_modelo, str(m.get("name", "")), flags=re.I)][:FIPE_MAX_VERSOES]
                for m in sel:
                    anos = api.get(f"{tipo}/brands/{b['code']}/models/{m['code']}/years", par)
                    if not anos:
                        continue
                    zero = [a for a in anos if str(a["code"]).startswith("32000")]
                    ano = (zero or anos)[0]
                    try:
                        d = api.get(f"{tipo}/brands/{b['code']}/models/{m['code']}/years/{ano['code']}", par)
                    except RuntimeError as e:
                        print("    ", e)
                        continue
                    precos.append({"tipo": tipo, "marca": d.get("brand", b["name"]), "modelo": d.get("model", m["name"]),
                                   "ano_modelo": d.get("modelYear"), "combustivel": d.get("fuel"),
                                   "codigo_fipe": d.get("codeFipe"), "ano_id": ano["code"],
                                   "referencia": d.get("referenceMonth", ref_atual.get("month")),
                                   "preco": num_br(d.get("price"))})
                    if d.get("codeFipe"):
                        try:
                            h = api.get(f"{tipo}/{d['codeFipe']}/years/{ano['code']}/history")
                            for it in _lista_historico(h):
                                it.update({"tipo": tipo, "codigo_fipe": d["codeFipe"], "modelo": d.get("model"),
                                           "ano_id": ano["code"]})
                                hist.append(it)
                        except RuntimeError as e:
                            print("    histórico:", e)
    print(f"  FIPE: {len(precos)} modelos, {len(hist)} pontos de histórico; {api.n} requisições usadas")
    if not precos:
        raise RuntimeError("Nenhum preço FIPE coletado (verifique a cesta e a conexão).")
    out = {"FIPE_Precos": pd.DataFrame(precos)}
    if hist:
        h = pd.DataFrame(hist)
        out["FIPE_Historico"] = h
        h["idx"] = h.groupby("codigo_fipe")["preco"].transform(lambda s: s / s.iloc[-1] * 100 if len(s) else s)
    return out


# ============================================================================ BLOCO abac
PAD_NUM = r"(\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?)\s*(milh[ãa]o|milh[õo]es|mil|bilh[ãa]o|bilh[õo]es)"
INDICADORES_ABAC = {"participantes_ativos": r"participantes? ativos?|consorciados? ativos?",
                    "vendas_de_cotas": r"vendas? de cotas|adesões",
                    "cotas_contempladas": r"contempla[çc][õo]es|cotas contempladas",
                    "negocios_reais": r"neg[óo]cios|volume de cr[ée]dito"}
MULT = {"mil": 1e3, "milhão": 1e6, "milhao": 1e6, "milhões": 1e6, "milhoes": 1e6, "bilhão": 1e9, "bilhao": 1e9,
        "bilhões": 1e9, "bilhoes": 1e9}
MES_PT = {"janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
          "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}


def _texto(html):
    t = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t)


def periodo_do_texto(titulo, url):
    base = f"{titulo} {url}".lower()
    m = re.search(r"(janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)[\s_/-]*(?:de[\s_/-]*)?(20\d\d)", base)
    if m:
        return f"{int(m.group(2))}{MES_PT[m.group(1).replace('ç', 'c') if m.group(1) not in MES_PT else m.group(1)]:02d}"
    return None


def bloco_abac():
    vistos, fila, paginas = set(), list(ABAC_SEMENTES), []
    pdfs = set()
    while fila and len(vistos) < ABAC_MAX_PAGINAS:
        u = fila.pop(0)
        if u in vistos:
            continue
        vistos.add(u)
        try:
            r = http_get(u, tentativas=2)
        except Exception as e:
            print(f"    ABAC: falha em {u} ({e})")
            continue
        if r.status_code != 200:
            continue
        html = r.text
        paginas.append((u, html))
        for h in re.findall(r'href="([^"#]+)"', html):
            full = h if h.startswith("http") else requests.compat.urljoin(u, h)
            if "abac.org.br" not in full:
                continue
            if full.lower().endswith(".pdf"):
                pdfs.add(full)
            elif re.search(r"dados-economicos|materia-de-|boletim|/page/\d+|[?&]paged?=\d+", full, flags=re.I):
                fila.append(full)
    print(f"  ABAC: {len(paginas)} páginas lidas, {len(pdfs)} PDFs encontrados")
    linhas = []
    for u, html in paginas:
        txt = _texto(html)
        titulo = (re.search(r"(?is)<title>(.*?)</title>", html) or [None, ""])[1].strip()
        per = periodo_do_texto(titulo, u)
        if not per:
            continue
        for ind, rx in INDICADORES_ABAC.items():
            for m in re.finditer(rx, txt, flags=re.I):
                trecho = txt[max(0, m.start() - 160): m.end() + 160]
                n = re.search(PAD_NUM, trecho, flags=re.I)
                if n:
                    linhas.append({"periodo": per, "indicador": ind, "valor": num_br(n.group(1)) * MULT[n.group(2).lower()],
                                   "unidade_original": n.group(2), "trecho": trecho.strip(), "url": u})
                    break
    if not linhas:
        raise RuntimeError("Nenhum indicador ABAC extraído automaticamente. Use a aba 'ABAC_Manual' e os PDFs listados.")
    ext = pd.DataFrame(linhas).drop_duplicates(["periodo", "indicador", "url"])
    modelo = pd.DataFrame({"periodo_AAAAMM": [], "segmento": [], "indicador": [], "valor": [], "unidade": [], "fonte_url": []})
    print(f"  ABAC: {len(ext)} valores extraídos (CONFERIR com o trecho/URL; extração é heurística)")
    return {"ABAC_Extraido": ext.sort_values(["periodo", "indicador"]),
            "ABAC_PDFs": pd.DataFrame({"url_pdf": sorted(pdfs)}), "ABAC_Manual": modelo}


# ============================================================================ painéis
def paineis(res):
    out = {}
    nac = res.get("Cotas_Nacional_Mensal")
    mensal = None
    if nac is not None:
        mensal = nac.copy()
        mensal["mes"] = pd.to_datetime(mensal["data_base"].astype(str) + "01", format="%Y%m%d")
    for chave, col in [("Selic_Dolar_Mensal", "mes"), ("Trends_Mensal", "mes")]:
        t = res.get(chave)
        if t is None:
            continue
        mensal = t if mensal is None else mensal.merge(t, on="mes", how="outer")
    if mensal is not None:
        out["Painel_Mensal"] = mensal.sort_values("mes").reset_index(drop=True)
    uf = res.get("Cotas_UF")
    ibge = res.get("UF_Pop_PIB_Anual")
    if uf is not None and ibge is not None:
        p = uf.merge(ibge.drop(columns=["fonte_pop", "fonte_pib"], errors="ignore"), on=["uf", "ano"], how="left")
        p = p.sort_values(["uf", "data_base"])
        for c in ["populacao", "pib_mil_reais", "pib_per_capita_reais"]:
            p[c + "_ffill"] = p.groupby("uf")[c].ffill()          # último ano disponível (PIB defasa ~2-3 anos)
        out["Painel_UF_Trimestral"] = p
    return out


# ============================================================================ saída
HEAD = (Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"),
        Alignment(wrap_text=True, vertical="center", horizontal="center"))


def salva(res, diag, caminho):
    os.makedirs(PASTA_CSV, exist_ok=True)
    notas = pd.DataFrame({"Notas": [
        "Cotas ativas: BCB, serviço OData PANORAMA_DE_CONSORCIOS (Metricas por DataBase=AAAAMM). Valores em 'mil' foram convertidos para unidades (coluna valor_unidades).",
        "Por UF: métricas 99-125 (somente cotas ativas por estado); o BCB não publica o cruzamento segmento x UF nesse serviço. Frequência: trimestral (mar/jun/set/dez).",
        "O serviço OData não traz administradora; o market share vem do arquivo AAAAMMConsorcios.zip (Segmentos_Consolidados). O endereço do ZIP não foi confirmado.",
        "Selic: SGS 432 (meta, % a.a.) e 4390 (acumulada no mês, % a.m.); dólar: SGS 1 (venda, diário). Agregação mensal: média e último valor.",
        "IBGE por UF: população (SIDRA 6579; Censo 2022 na tabela 4714) e PIB (5938). O PIB tem defasagem de 2-3 anos; o painel traz colunas '_ffill'.",
        "Google Trends: índice relativo (0-100) por requisição; os 3 termos foram pedidos juntos (mesma escala). Não é volume de buscas. O último mês pode estar incompleto.",
        "FIPE: preços de uma CESTA de modelos (não o catálogo inteiro) e histórico por código FIPE; a API gratuita limita requisições diárias.",
        "ABAC: sem API/planilha pública; indicadores extraídos por heurística de páginas públicas. CONFERIR com o trecho e a URL antes de usar.",
        "Dados de períodos diferentes: cada série tem seu último mês publicado; veja 'Diagnóstico'.",
    ]})
    abas = dict(res)
    abas["Diagnóstico"] = pd.DataFrame(diag)
    abas["Fontes e notas"] = notas
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        for nome, df in abas.items():
            n = nome[:31]
            d = df.copy()
            for c in d.columns:
                if str(d[c].dtype).startswith("datetime"):
                    d[c] = d[c].dt.strftime("%Y-%m-%d")
            d.to_excel(w, sheet_name=n, index=False)
            ws = w.sheets[n]
            for j, c in enumerate(d.columns, 1):
                h = ws.cell(row=1, column=j)
                h.font, h.fill, h.alignment = HEAD
                ws.column_dimensions[get_column_letter(j)].width = 120 if nome == "Fontes e notas" else (60 if c in ("trecho", "url", "detalhe") else 18)
            ws.freeze_panes = "A2"
            if len(d):
                ws.auto_filter.ref = ws.dimensions
            df.to_csv(os.path.join(PASTA_CSV, f"{n}.csv"), index=False, encoding="utf-8-sig")
    print(f"\nArquivo salvo: {caminho} ({len(abas)} abas) | CSVs em '{PASTA_CSV}/'")


BLOCOS = [("bcb", "BCB - cotas ativas (nacional, segmento, UF)", bloco_bcb),
          ("share", "BCB - market share por administradora", bloco_share),
          ("sgs", "BCB/SGS - Selic e dólar", bloco_sgs),
          ("ibge", "IBGE - UF (código, população, PIB)", bloco_ibge),
          ("trends", "Google Trends", bloco_trends),
          ("fipe", "Tabela FIPE", bloco_fipe),
          ("abac", "ABAC - estatísticas setoriais", bloco_abac)]


def main(selecionados=None):
    res, diag = {}, []
    for chave, nome, fn in BLOCOS:
        if selecionados and chave not in selecionados:
            continue
        print(f"\n== {nome}")
        try:
            r = fn()
            res.update(r)
            diag.append({"bloco": chave, "descricao": nome, "situacao": "OK",
                         "abas": ", ".join(r), "linhas": sum(len(v) for v in r.values()), "detalhe": ""})
            print(f"  OK: {', '.join(f'{k} ({len(v)})' for k, v in r.items())}")
        except Exception as e:
            diag.append({"bloco": chave, "descricao": nome, "situacao": "ERRO", "abas": "", "linhas": 0,
                         "detalhe": str(e)[:600]})
            print(f"  ERRO neste bloco: {e}")
    res.update(paineis(res))
    print("\nResumo:", [(d["bloco"], d["situacao"]) for d in diag])
    salva(res, diag, SAIDA_XLSX)


if __name__ == "__main__":
    sel = {a for a in sys.argv[1:] if a in {b[0] for b in BLOCOS}}
    main(sel or None)
# %%
