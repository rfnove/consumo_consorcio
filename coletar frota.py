#%%!/usr/bin/env python3
"""
Coleta a frota de veículos por município (SENATRAN / Ministério dos Transportes) e salva em XLSX.

Fonte: planilha mensal "Frota por Município e Tipo" (RENAVAM), publicada em
       https://www.gov.br/transportes/pt-br/assuntos/transito/conteudo-Senatran/frota-de-veiculos-AAAA

Saída (por município, para o mês de referência):
  - Automóveis e utilitários   (automóvel + caminhonete + camioneta + utilitário)
  - Motocicletas               (motocicleta + motoneta; ciclomotor em coluna separada)
  - Caminhões, tratores e reboques (caminhão + caminhão-trator + trator + reboque + semirreboque)

Requisitos: pip install requests pandas openpyxl xlrd      (xlrd é necessário para arquivos .xls)
Uso:        python coletar_frota.py            (usa ANO_REF / MES_REF abaixo)
            python coletar_frota.py 202112     (mês específico, AAAAMM)
"""
import difflib
import io
import re
import sys
import unicodedata

from datetime import date

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ANO_REF, MES_REF = None, None     # None = MÊS MAIS RECENTE publicado no site da SENATRAN
SAIDA = "dados_frota_municipios.xlsx"
HEADERS = {"User-Agent": "Mozilla/5.0 (coleta-frota)"}
URL_IBGE_MUN = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios"
BASE = "https://www.gov.br/transportes/pt-br/assuntos/transito"
PAGINA_ANO = BASE + "/conteudo-Senatran/frota-de-veiculos-{ano}"

# Link direto conhecido (dezembro/2021). Para outros meses o script procura o link na página do ano.
URLS_CONHECIDAS = {
    (2021, 12): BASE + "/arquivos-senatran/estatisticas/renavam/2021/dezembro/frota-munic-modelo-dezembro-2021.xls",
}
MESES = {1: "janeiro", 2: "fevereiro", 3: "marco", 4: "abril", 5: "maio", 6: "junho", 7: "julho",
         8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}

# Tipos de veículo (regex sobre o nome normalizado da coluna) -> grupos pedidos
# Grupos pedidos (regex sobre o nome normalizado da coluna da planilha)
GRUPOS = {
    "automoveis_utilitarios": [r"^AUTOMOVEL$", r"^CAMINHONETE$", r"^CAMIONETA$", r"^UTILITARIO$"],
    "motocicletas": [r"^MOTOCICLETA$", r"^MOTONETA$"],
    "caminhoes_tratores_reboques": [r"^CAMINHAO$", r"^CAMINHAO_TRATOR$", r"^TRATOR(_|$)",
                                    r"^REBOQUE$", r"^SEMI_?REBOQUE$"],
}
ROTULOS = {
    "AUTOMOVEL": "Automóvel", "BONDE": "Bonde", "CAMINHAO": "Caminhão", "CAMINHAO_TRATOR": "Caminhão-trator",
    "CAMINHONETE": "Caminhonete", "CAMIONETA": "Camioneta", "CHASSI_PLATAF": "Chassi/plataforma",
    "CHASSI_PLATAFORMA": "Chassi/plataforma", "CICLOMOTOR": "Ciclomotor", "MICRO_ONIBUS": "Micro-ônibus",
    "MOTOCICLETA": "Motocicleta", "MOTONETA": "Motoneta", "ONIBUS": "Ônibus", "QUADRICICLO": "Quadriciclo",
    "REBOQUE": "Reboque", "SEMI_REBOQUE": "Semirreboque", "SEMIREBOQUE": "Semirreboque",
    "SIDE_CAR": "Side-car", "OUTROS": "Outros", "TRATOR_ESTEIRA": "Trator de esteira",
    "TRATOR_RODAS": "Trator de rodas", "TRICICLO": "Triciclo", "UTILITARIO": "Utilitário",
}
# Correções manuais de pareamento: (UF, nome da SENATRAN normalizado) -> código IBGE de 7 dígitos.
# O script imprime, para cada município sem par, a chave exata para colar aqui. Exemplo:
#   ("SP", "SAO_PAULO"): 3550308,
DE_PARA_MANUAL = {
}

UFS = {"ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM", "BAHIA": "BA", "CEARA": "CE",
       "DISTRITO_FEDERAL": "DF", "ESPIRITO_SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA",
       "MATO_GROSSO": "MT", "MATO_GROSSO_DO_SUL": "MS", "MINAS_GERAIS": "MG", "PARA": "PA",
       "PARAIBA": "PB", "PARANA": "PR", "PERNAMBUCO": "PE", "PIAUI": "PI", "RIO_DE_JANEIRO": "RJ",
       "RIO_GRANDE_DO_NORTE": "RN", "RIO_GRANDE_DO_SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR",
       "SANTA_CATARINA": "SC", "SAO_PAULO": "SP", "SERGIPE": "SE", "TOCANTINS": "TO"}


def norm(txt):
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]+", "_", t.upper()).strip("_")


MES_NUM = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
           "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}


def links_da_pagina(ano):
    """Lê a página do ano e devolve [(ano, mes, texto_do_link, url)], associando cada link ao título
    '(Mês de AAAA)' mais próximo acima dele."""
    r = requests.get(PAGINA_ANO.format(ano=ano), headers=HEADERS, timeout=120)
    r.raise_for_status()
    html = r.text
    titulos = []
    for m in re.finditer(r"\(([A-Za-zÀ-ÿ]+) de (20\d\d)\)", html):
        mes = MES_NUM.get(norm(m.group(1)).lower().replace("_", ""))
        if mes:
            titulos.append((m.start(), mes, int(m.group(2))))
    saida = []
    for a_ in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, flags=re.S | re.I):
        anteriores = [t for t in titulos if t[0] < a_.start()]
        if not anteriores:
            continue
        _, mes, ano_t = anteriores[-1]
        texto = re.sub(r"<[^>]+>", "", a_.group(2)).strip()
        url = a_.group(1)
        url = url if url.startswith("http") else "https://www.gov.br" + url
        saida.append((ano_t, mes, norm(texto), url))
    return saida


def acha_url(ano=None, mes=None):
    """Devolve (ano, mes, url) do arquivo 'Frota por Município e Tipo'. Sem ano/mês: o mais recente publicado."""
    if ano and mes and (ano, mes) in URLS_CONHECIDAS:
        return ano, mes, URLS_CONHECIDAS[(ano, mes)]
    anos = [ano] if ano else [date.today().year, date.today().year - 1]
    cand = []
    for a_ in anos:
        try:
            cand += [c for c in links_da_pagina(a_) if "FROTA_POR_MUNICIPIO_E_TIPO" in c[2]]
        except Exception as e:
            print(f"  Aviso: página de {a_} indisponível ({e})")
    if mes:
        cand = [c for c in cand if c[1] == mes]
    if not cand:
        raise RuntimeError("Link de 'Frota por Município e Tipo' não encontrado na página da SENATRAN.")
    a_, m_, _, url = max(cand, key=lambda c: (c[0], c[1]))
    return a_, m_, url


def baixa(url):
    r = requests.get(url, headers=HEADERS, timeout=300)
    if r.status_code != 200 or b"<html" in r.content[:300].lower():
        raise RuntimeError(f"Falha ao baixar {url} (HTTP {r.status_code}).")
    return r.content


def le_frota(conteudo):
    try:
        bruto = pd.read_excel(io.BytesIO(conteudo), header=None, dtype=object)
    except ImportError as e:
        raise RuntimeError("Para ler arquivos .xls instale o xlrd:  pip install xlrd") from e
    # localiza a linha de cabeçalho (contém UF e MUNICIPIO)
    for i in range(min(len(bruto), 30)):
        cel = [norm(x) for x in bruto.iloc[i].tolist() if pd.notna(x)]
        if "UF" in cel and "MUNICIPIO" in cel:
            cab = i
            break
    else:
        raise RuntimeError("Cabeçalho (UF / MUNICIPIO) não encontrado na planilha.")
    df = bruto.iloc[cab + 1:].copy()
    df.columns = [norm(x) if pd.notna(x) else f"SEM_NOME_{j}" for j, x in enumerate(bruto.iloc[cab])]
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(subset=["MUNICIPIO"])
    df = df[~df["MUNICIPIO"].map(norm).isin(["TOTAL", "TOTAIS"])]
    return df


def tabela_ibge():
    r = requests.get(URL_IBGE_MUN, headers=HEADERS, timeout=120)
    r.raise_for_status()
    linhas = []
    for m in r.json():
        uf = (m.get("microrregiao") or {}).get("mesorregiao", {}).get("UF", {}).get("sigla") \
            or m["regiao-imediata"]["regiao-intermediaria"]["UF"]["sigla"]
        linhas.append({"codigo_ibge_7": m["id"], "nome_ibge": m["nome"], "uf": uf,
                       "chave": f"{uf}|{norm(m['nome'])}"})
    return pd.DataFrame(linhas)


def num(s):
    return pd.to_numeric(s.astype(str).str.strip().str.replace(".", "", regex=False)
                         .str.replace(",", ".", regex=False), errors="coerce")


def coleta(ano=ANO_REF, mes=MES_REF):
    print("Procurando arquivo de frota da SENATRAN...")
    ano, mes, url = acha_url(ano, mes)
    print(f"  Mês mais recente/escolhido: {mes:02d}/{ano}")
    print("  Baixando:", url)
    raw = le_frota(baixa(url))
    print("  Colunas da planilha:", [c for c in raw.columns])

    # --- todos os tipos de veículo presentes na planilha (colunas numéricas, exceto UF/MUNICIPIO/TOTAL) ---
    ignora = {"UF", "MUNICIPIO", "TOTAL", "TOTAIS"}
    tipos = [c for c in raw.columns
             if c not in ignora and not c.startswith(("COD", "SEM_NOME")) and num(raw[c]).notna().mean() > 0.9]
    print("  Tipos de veículo encontrados:", tipos)

    df = pd.DataFrame({"uf_raw": raw["UF"].astype(str).str.strip(),
                       "municipio_senatran": raw["MUNICIPIO"].astype(str).str.strip()})
    df["uf"] = df["uf_raw"].map(lambda x: x.upper() if len(x) == 2 else UFS.get(norm(x), x))
    tipo_cols = {}                                   # nome na planilha -> coluna no DataFrame
    for c in tipos:
        tipo_cols[c] = "t_" + c.lower()
        df[tipo_cols[c]] = num(raw[c]).fillna(0)
    soma_tipos = df[list(tipo_cols.values())].sum(axis=1)
    tot_col = next((c for c in raw.columns if c == "TOTAL"), None)
    df["total_frota"] = num(raw[tot_col]) if tot_col else soma_tipos
    if tot_col:
        dif = (df["total_frota"] - soma_tipos).abs()
        print(f"  Conferência: TOTAL difere da soma dos tipos em {int((dif > 0).sum())} linhas "
              f"(diferença máxima: {int(dif.max()) if len(dif) else 0}).")
    df = df.reset_index(drop=True)

    # --- grupos pedidos ---
    usados = set()
    for grupo, pads in GRUPOS.items():
        cols_g = [tipo_cols[c] for c in tipos if any(re.search(p, c) for p in pads)]
        usados |= set(cols_g)
        df[grupo] = df[cols_g].sum(axis=1) if cols_g else 0
        print(f"  {grupo} = {[c[2:].upper() for c in cols_g]}")
        if not cols_g:
            print(f"  ATENÇÃO: nenhuma coluna encontrada para {grupo}.")
    df["demais_tipos"] = df[[c for c in tipo_cols.values() if c not in usados]].sum(axis=1)
    tipos_info = [(tipo_cols[c], ROTULOS.get(c, c.replace("_", " ").title())) for c in tipos]

    # --- código IBGE: nome exato (UF + nome) e depois aproximado ---
    print("Baixando municípios do IBGE para obter o código...")
    ibge = tabela_ibge()
    df["chave"] = df["uf"] + "|" + df["municipio_senatran"].map(norm)
    # se a planilha já trouxer o código IBGE, usa direto
    col_cod = next((c for c in raw.columns if c.startswith("COD") and
                    num(raw[c]).isin(ibge["codigo_ibge_7"]).mean() > 0.9), None)
    if col_cod:
        df["codigo_ibge_7"] = num(raw[col_cod]).astype("Int64")
        df = df.merge(ibge[["codigo_ibge_7", "nome_ibge"]], how="left", on="codigo_ibge_7")
        aprox = pd.DataFrame(columns=["uf", "municipio_senatran", "municipio_ibge", "codigo_ibge_7"])
    else:
        df = df.merge(ibge[["chave", "codigo_ibge_7", "nome_ibge"]].drop_duplicates("chave"),
                      how="left", on="chave")
        for idx in df.index[df["codigo_ibge_7"].isna()]:           # correções manuais primeiro
            k = (df.at[idx, "uf"], norm(df.at[idx, "municipio_senatran"]))
            if k in DE_PARA_MANUAL:
                lin = ibge[ibge["codigo_ibge_7"] == int(DE_PARA_MANUAL[k])]
                if len(lin):
                    df.at[idx, "codigo_ibge_7"] = int(DE_PARA_MANUAL[k])
                    df.at[idx, "nome_ibge"] = lin.iloc[0]["nome_ibge"]
        usados = set(df["codigo_ibge_7"].dropna().astype(int))
        aprox = []
        for idx in df.index[df["codigo_ibge_7"].isna()]:
            uf, nome = df.at[idx, "uf"], norm(df.at[idx, "municipio_senatran"])
            cand = ibge[(ibge["uf"] == uf) & (~ibge["codigo_ibge_7"].isin(usados))].copy()
            if cand.empty:
                continue
            cand["n"] = cand["nome_ibge"].map(norm)
            par = difflib.get_close_matches(nome, cand["n"].tolist(), n=1, cutoff=0.86)
            if par:
                lin = cand[cand["n"] == par[0]].iloc[0]
                df.at[idx, "codigo_ibge_7"], df.at[idx, "nome_ibge"] = lin["codigo_ibge_7"], lin["nome_ibge"]
                usados.add(int(lin["codigo_ibge_7"]))
                aprox.append({"uf": uf, "municipio_senatran": df.at[idx, "municipio_senatran"],
                              "municipio_ibge": lin["nome_ibge"], "codigo_ibge_7": int(lin["codigo_ibge_7"])})
        aprox = pd.DataFrame(aprox, columns=["uf", "municipio_senatran", "municipio_ibge", "codigo_ibge_7"])
    sem = df[df["codigo_ibge_7"].isna()][["uf", "municipio_senatran", "total_frota"]].copy()
    print(f"  Pareados: {df['codigo_ibge_7'].notna().sum()} | aproximados: {len(aprox)} | sem par: {len(sem)}")
    if len(sem):
        sem = df[df["codigo_ibge_7"].isna()][["uf", "municipio_senatran", "total_frota"]].copy()
    sem["total_frota"] = sem["total_frota"].fillna(0)  # Preenche NaN com 0 antes do print/soma
    print(f"  Pareados: {df['codigo_ibge_7'].notna().sum()} | aproximados: {len(aprox)} | sem par: {len(sem)}")
    if len(sem):
        total_sem_cod = sem["total_frota"].sum()
        pct = (total_sem_cod / df["total_frota"].sum() * 100) if df["total_frota"].sum() else 0
        print(f"  Frota sem código IBGE: {int(total_sem_cod):,} veículos ({pct:.3f}% do total).".replace(",", "."))
        print("  Sem par (chave para DE_PARA_MANUAL | veículos):")
        for _, r in sem.sort_values("total_frota", ascending=False).head(20).iterrows():
            print(f"     ({r['uf']!r}, {norm(r['municipio_senatran'])!r}) | {int(r['total_frota'])}")
    falt = ibge[~ibge["codigo_ibge_7"].isin(df["codigo_ibge_7"].dropna().astype(int))]
    if len(falt):
        print(f"  Municípios do IBGE sem linha na frota ({len(falt)}):",
              ", ".join(f"{r.nome_ibge}-{r.uf} ({r.codigo_ibge_7})" for r in falt.head(20).itertuples()))

    df["periodo"] = f"{ano}{mes:02d}"
    df["municipio"] = df["nome_ibge"].fillna(df["municipio_senatran"])
    pareado = df[df["codigo_ibge_7"].notna()].copy()
    pareado["codigo_ibge_7"] = pareado["codigo_ibge_7"].astype(int)
    dup = pareado["codigo_ibge_7"].duplicated(keep=False).sum()
    if dup:
        print(f"  ATENÇÃO: {dup} linhas com código IBGE repetido; somando por município.")
        soma = [c for c in pareado.columns if pareado[c].dtype.kind in "if"]
        pareado = (pareado.groupby(["codigo_ibge_7", "municipio", "uf", "periodo"], as_index=False)[soma]
                   .sum())
    pareado = pareado.sort_values("codigo_ibge_7").reset_index(drop=True)
    return pareado, aprox, sem, url, tipos_info


def salva_xlsx(df, aprox, sem, url, tipos_info, caminho):
    fonte = Font(name="Arial", size=10)
    cab_f = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cab_fill = PatternFill("solid", fgColor="1F4E78")

    def folha(ws, colunas, freeze="D2"):  # colunas: [(titulo, chave, formato)]
        for j, (t, _, _) in enumerate(colunas, 1):
            c = ws.cell(row=1, column=j, value=t)
            c.font, c.fill = cab_f, cab_fill
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            ws.column_dimensions[get_column_letter(j)].width = 34 if t == "Município" else 18
        ws.row_dimensions[1].height = 45
        for i, row in enumerate(df.to_dict("records"), start=2):
            for j, (_, k, fmt) in enumerate(colunas, 1):
                v = row.get(k)
                v = None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
                c = ws.cell(row=i, column=j, value=v)
                c.font, c.number_format = fonte, fmt
        ws.freeze_panes = freeze
        ws.auto_filter.ref = ws.dimensions

    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    folha(ws, [
        ("Código IBGE (7 dígitos)", "codigo_ibge_7", "0"),
        ("Município", "municipio", "@"),
        ("UF", "uf", "@"),
        ("Período (AAAAMM)", "periodo", "@"),
        ("Automóveis e utilitários", "automoveis_utilitarios", "#,##0"),
        ("Motocicletas", "motocicletas", "#,##0"),
        ("Caminhões, tratores e reboques", "caminhoes_tratores_reboques", "#,##0"),
        ("Demais tipos (ônibus, ciclomotor etc.)", "demais_tipos", "#,##0"),
        ("Total da frota", "total_frota", "#,##0"),
    ])
    wd = wb.create_sheet("Frota por tipo")
    folha(wd, [("Código IBGE (7 dígitos)", "codigo_ibge_7", "0"), ("Município", "municipio", "@"),
               ("UF", "uf", "@"), ("Período (AAAAMM)", "periodo", "@")]
          + [(rotulo, col, "#,##0") for col, rotulo in tipos_info]
          + [("Total da frota", "total_frota", "#,##0")], freeze="E2")

    def tabela(nome, tab):
        if tab is None or not len(tab):
            return
        w = wb.create_sheet(nome)
        for j, c in enumerate(tab.columns, 1):
            h = w.cell(row=1, column=j, value=c)
            h.font, h.fill = cab_f, cab_fill
            w.column_dimensions[get_column_letter(j)].width = 32
        for i, row in enumerate(tab.itertuples(index=False), start=2):
            for j, v in enumerate(row, 1):
                v = None if (isinstance(v, float) and pd.isna(v)) else v
                w.cell(row=i, column=j, value=v).font = fonte
    tabela("Pareamento aproximado", aprox)
    tabela("Sem código IBGE", sem)

    nt = wb.create_sheet("Fontes e notas")
    notas = [
        "Fonte: SENATRAN / Ministério dos Transportes - Frota de veículos por município e tipo (RENAVAM).",
        f"Arquivo: {url}",
        f"Período: {df['periodo'].iloc[0]} (posição da frota no mês de referência).",
        "Automóveis e utilitários = automóvel + caminhonete + camioneta + utilitário.",
        "Motocicletas = motocicleta + motoneta (ciclomotor fica apenas na aba 'Frota por tipo').",
        "Caminhões, tratores e reboques = caminhão + caminhão-trator + trator (esteira e rodas) + reboque + semirreboque.",
        "Demais tipos = total menos os três grupos acima (ônibus, micro-ônibus, ciclomotor, bonde, quadriciclo, triciclo etc.). A aba 'Frota por tipo' traz TODOS os tipos da planilha, um por coluna.",
        "A frota é contada no município de registro do veículo (licenciamento), que pode diferir do de uso.",
        "O código IBGE foi obtido por pareamento UF + nome com a API de localidades do IBGE (ver abas de pareamento).",
        "Frota registrada inclui veículos antigos que podem não estar mais em circulação (não há baixa automática).",
    ]
    for i, t in enumerate(notas, 1):
        nt.cell(row=i, column=1, value=t).font = fonte
    nt.column_dimensions["A"].width = 130
    wb.save(caminho)
    print(f"Arquivo salvo: {caminho} ({len(df)} municípios)")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if re.fullmatch(r"\d{6}", a)]
    ano, mes = (int(args[0][:4]), int(args[0][4:])) if args else (ANO_REF, MES_REF)
    salva_xlsx(*coleta(ano, mes), SAIDA)
# %%
