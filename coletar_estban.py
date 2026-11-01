#%%!/usr/bin/env python3
"""
Coleta dados do ESTBAN (Estatística Bancária Mensal por Município - Banco Central)
e salva em XLSX.

Saída (por município, para um mês de referência):
  - Saldo de operações de crédito (verbete 160)
  - Depósitos à vista e a prazo
  - Quantidade de agências

Requisitos: pip install requests pandas openpyxl
Uso:        python coletar_estban.py              (usa o mês mais recente disponível)
            python coletar_estban.py 202312       (mês específico, AAAAMM)
"""
import difflib
import io
import re
import sys
import zipfile
import unicodedata
from datetime import date

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_BCB = "https://www.bcb.gov.br/content/estatisticas/estatistica_bancaria_estban"


def urls_estban(nivel, periodo):
    """URLs candidatas (o BCB já mudou a extensão ao longo dos anos). nivel: municipio | agencia."""
    sufixo = "ESTBAN" if nivel == "municipio" else "ESTBAN_AG"
    return [f"{BASE_BCB}/{nivel}/{periodo}_{sufixo}{ext}" for ext in (".csv.zip", ".ZIP", ".csv")]


URL_IBGE_MUN = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios"
HEADERS = {"User-Agent": "Mozilla/5.0 (coleta-estban)"}
SAIDA = "dados_estban_municipios.xlsx"

# Ano de referência: usa dezembro desse ano. None = MÊS MAIS RECENTE DISPONÍVEL (padrão atual do projeto).
ANO_REF = None

# Padrões (regex, aplicados aos nomes de coluna normalizados) dos verbetes usados.
# Se quiser forçar colunas específicas, preencha com a lista de nomes exatos.
PAD_CREDITO = r"^VERBETE_160(_|$)"
PAD_POUPANCA = r"^VERBETE_420(_|$)"   # 420 = depósitos de poupança (documentado)
PAD_VISTA = r"VISTA"                  # verbetes 4xx com "VISTA" no nome (401 governos, 410/411 setor privado)
CODIGOS_VISTA = r"^VERBETE_(401|410|411)(_|$)"   # plano B, por código ESTBAN
PAD_PRAZO = r"DEPOSITOS_A_PRAZO"


def abrevia(cols):
    """Mostra só os códigos dos verbetes de cada coluna (nomes agregados do BCB são enormes)."""
    if isinstance(cols, str):
        cols = [cols]
    return [", ".join(re.findall(r"VERBETE_(\d{3})", c)) or c for c in (cols or [])]


def norm(txt):
    """Maiúsculas, sem acentos e com '_' no lugar de símbolos."""
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]+", "_", t.upper()).strip("_")


def baixa_arquivo(nivel, periodo):
    """Tenta cada URL candidata; devolve os bytes (ZIP ou CSV) ou None."""
    for url in urls_estban(nivel, periodo):
        try:
            r = requests.get(url, headers=HEADERS, timeout=180)
        except requests.RequestException as ex:
            print(f"  Erro de conexão em {url}: {ex}")
            continue
        inicio = r.content[:200].lower()
        if r.status_code == 200 and (r.content.startswith(b"PK") or
                                     (b"<html" not in inicio and b"<!doctype" not in inicio)):
            return r.content
        print(f"  {url} -> HTTP {r.status_code}")
    return None


def conta_agencias(periodo):
    """Conta agências distintas por município a partir do arquivo por agência."""
    z = baixa_arquivo("agencia", periodo)
    if z is None:
        print("  Não foi possível baixar o arquivo por agência.")
        return None
    df = le_estban(z)
    cod = next(c for c in df.columns if c.startswith("CODMUN"))
    df["codigo_bcb"] = df[cod].astype(str).str.strip()
    chaves = ["codigo_bcb"] + [c for c in ("CNPJ", "AGENCIA") if c in df.columns]
    df = df.drop_duplicates(subset=chaves)
    return df.groupby("codigo_bcb").size().rename("agencias").reset_index()


def acha_periodo(periodo=None):
    """Devolve (periodo, bytes_zip). Sem período: tenta do mês atual para trás."""
    if periodo:
        z = baixa_arquivo("municipio", periodo)
        if z is None:
            raise RuntimeError(f"ESTBAN {periodo} não encontrado no site do BCB.")
        return periodo, z
    hoje = date.today()
    ano, mes = hoje.year, hoje.month
    for _ in range(18):
        p = f"{ano}{mes:02d}"
        print(f"  Tentando {p}...")
        z = baixa_arquivo("municipio", p)
        if z:
            return p, z
        mes -= 1
        if mes == 0:
            ano, mes = ano - 1, 12
    raise RuntimeError("Nenhum arquivo ESTBAN encontrado nos últimos 18 meses.")


def le_estban(zbytes):
    if zbytes.startswith(b"PK"):
        with zipfile.ZipFile(io.BytesIO(zbytes)) as z:
            nome = next(n for n in z.namelist() if n.lower().endswith(".csv"))
            texto = z.read(nome).decode("latin-1")
    else:
        texto = zbytes.decode("latin-1")
    linhas = texto.splitlines()
    ini = next(i for i, l in enumerate(linhas)
               if "CODMUN" in norm(l) or l.startswith("#DATA_BASE"))
    df = pd.read_csv(io.StringIO("\n".join(linhas[ini:])), sep=";", decimal=",",
                     dtype=str, na_values=["", " "])
    df.columns = [norm(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    return df.dropna(how="all")


def to_num(s):
    s = s.astype(str).str.strip()
    return pd.to_numeric(s.str.replace(",", ".", regex=False), errors="coerce")


def tabela_ibge():
    r = requests.get(URL_IBGE_MUN, headers=HEADERS, timeout=120)
    r.raise_for_status()
    linhas = []
    for m in r.json():
        uf = (m.get("microrregiao") or {}).get("mesorregiao", {}).get("UF", {}).get("sigla") \
            or m["regiao-imediata"]["regiao-intermediaria"]["UF"]["sigla"]
        linhas.append({"codigo_ibge_7": m["id"], "nome_ibge": m["nome"],
                       "uf": uf, "chave": f"{uf}|{norm(m['nome'])}"})
    return pd.DataFrame(linhas)


def coleta(periodo=None):
    print("Procurando arquivo ESTBAN no site do BCB...")
    periodo, zbytes = acha_periodo(periodo)
    print(f"Período: {periodo}")
    raw = le_estban(zbytes)

    col_cod = next(c for c in raw.columns if c.startswith("CODMUN"))
    col_uf = next(c for c in raw.columns if c == "UF")
    col_nome = next(c for c in ["MUNICIPIO", "NOME_MUNICIPIO", "NOMEMUN"] if c in raw.columns)
    verbetes = [c for c in raw.columns if c.startswith("VERBETE_")]
    if not verbetes:
        raise RuntimeError(f"Sem colunas de verbete. Colunas lidas: {list(raw.columns)}")

    for c in verbetes + [c for c in raw.columns if c.startswith("AGEN_")]:
        raw[c] = to_num(raw[c])

    c_cred = [c for c in verbetes if re.search(PAD_CREDITO, c)]
    print("  Verbetes de passivo (4xx) presentes no arquivo:")
    for c in [c for c in verbetes if c.startswith("VERBETE_4")]:
        print("    ", abrevia(c)[0])
    def tem_verbete_entre(c, ini, fim):
        return any(ini <= int(x) <= fim for x in re.findall(r"VERBETE_(\d{3})", c))
    # à vista = verbetes 401 a 419 (no arquivo podem vir juntos numa coluna agregada ou separados)
    c_vista = [c for c in verbetes if tem_verbete_entre(c, 401, 419)]
    c_prazo = [c for c in verbetes if re.search(PAD_PRAZO, c)]   # 432 = depósitos a prazo
    c_poup = [c for c in verbetes if re.search(PAD_POUPANCA, c)]

    def cols_verbete(*codigos):
        return [c for c in verbetes if any(int(x) in codigos for x in re.findall(r"VERBETE_(\d{3})", c))]
    mod = {"credito_emprestimos": cols_verbete(161), "credito_financiamentos": cols_verbete(162),
           "credito_rural": cols_verbete(163), "credito_agroindustrial": cols_verbete(167),
           "credito_imobiliario": cols_verbete(169)}
    print("  Crédito por modalidade (verbetes):", {k: abrevia(v) for k, v in mod.items()})
    print("  Verbetes usados -> crédito:", abrevia(c_cred))
    print("                     à vista:", abrevia(c_vista))
    print("                     a prazo:", abrevia(c_prazo))
    print("                     poupança:", abrevia(c_poup))
    if not (c_cred and c_vista and c_prazo):
        print("  ATENÇÃO: algum verbete não foi encontrado. Veja a aba 'ESTBAN completo' "
              "e ajuste PAD_CREDITO/PAD_VISTA/PAD_PRAZO no topo do script.")

    c_ag = next((c for c in ["AGEN_PROCESSADAS", "AGEN_ESPERADAS"] if c in raw.columns), None)
    ag_df = None
    if c_ag is None:
        print("  Arquivo por município sem coluna de agências; contando pelo arquivo por agência...")
        ag_df = conta_agencias(periodo)

    raw["_cod"] = raw[col_cod].astype(str).str.strip()
    raw["_uf"] = raw[col_uf].astype(str).str.strip().str.upper()
    raw["_nome"] = raw[col_nome].astype(str).str.strip()

    # Soma por município (funciona mesmo se o arquivo vier com mais de uma linha por município)
    agg = {c: "sum" for c in verbetes}
    if c_ag:
        agg[c_ag] = "sum"
    g = raw.groupby(["_cod", "_uf", "_nome"], as_index=False).agg(agg)

    out = pd.DataFrame({
        "codigo_bcb": g["_cod"], "uf": g["_uf"], "municipio_estban": g["_nome"],
        "saldo_credito": g[c_cred].sum(axis=1, min_count=1) if c_cred else None,
        "depositos_vista": g[c_vista].sum(axis=1, min_count=1) if c_vista else None,
        "depositos_prazo": g[c_prazo].sum(axis=1, min_count=1) if c_prazo else None,
        "depositos_poupanca": g[c_poup].sum(axis=1, min_count=1) if c_poup else None,
        **{k: (g[v].sum(axis=1, min_count=1) if v else None) for k, v in mod.items()},
        "agencias": g[c_ag] if c_ag else None,
    })

    if c_ag is None and ag_df is not None:
        out = out.drop(columns="agencias").merge(ag_df, on="codigo_bcb", how="left")
        c_ag = "contagem no arquivo por agência (CNPJ + AGENCIA distintos)"

    # --- Código IBGE: usa direto se já for IBGE; senão cruza por UF + nome ---
    print("Baixando tabela de municípios do IBGE para pareamento de códigos...")
    ibge = tabela_ibge()
    ibge_ids = set(ibge["codigo_ibge_7"])
    aprox = None
    if out["codigo_bcb"].str.fullmatch(r"\d{7}").all() and \
            out["codigo_bcb"].astype(int).isin(ibge_ids).mean() > 0.98:
        out["codigo_ibge_7"] = out["codigo_bcb"].astype(int)
        out = out.merge(ibge[["codigo_ibge_7", "nome_ibge"]], how="left", on="codigo_ibge_7")
    else:
        out["chave"] = out["uf"] + "|" + out["municipio_estban"].map(norm)
        out = out.merge(ibge[["chave", "codigo_ibge_7", "nome_ibge"]], how="left", on="chave")
        out = out.drop(columns="chave")
        out["pareamento"] = out["codigo_ibge_7"].notna().map({True: "nome exato", False: None})
        # Nomes diferentes entre BCB e IBGE: tenta correspondência aproximada dentro da mesma UF
        usados_ibge = set(out["codigo_ibge_7"].dropna().astype(int))
        aprox = []
        for idx in out.index[out["codigo_ibge_7"].isna()]:
            uf, nome = out.at[idx, "uf"], norm(out.at[idx, "municipio_estban"])
            cand = ibge[(ibge["uf"] == uf) & (~ibge["codigo_ibge_7"].isin(usados_ibge))].copy()
            cand["n"] = cand["nome_ibge"].map(norm)
            par = difflib.get_close_matches(nome, cand["n"].tolist(), n=1, cutoff=0.86)
            if par:
                lin = cand[cand["n"] == par[0]].iloc[0]
                out.at[idx, "codigo_ibge_7"] = lin["codigo_ibge_7"]
                out.at[idx, "nome_ibge"] = lin["nome_ibge"]
                out.at[idx, "pareamento"] = "aproximado"
                usados_ibge.add(int(lin["codigo_ibge_7"]))
                aprox.append({"uf": uf, "municipio_estban": out.at[idx, "municipio_estban"],
                              "municipio_ibge": lin["nome_ibge"],
                              "codigo_ibge_7": int(lin["codigo_ibge_7"])})
        aprox = pd.DataFrame(aprox, columns=["uf", "municipio_estban", "municipio_ibge", "codigo_ibge_7"])
        print(f"  Pareamento aproximado aplicado a {len(aprox)} municípios (conferir aba 'Pareamento aproximado').")
    sem = out[out["codigo_ibge_7"].isna()].copy()
    print(f"  Municípios pareados: {out['codigo_ibge_7'].notna().sum()} | sem par: {len(sem)}")
    out["total_depositos"] = out[["depositos_vista", "depositos_prazo"]].sum(axis=1, min_count=1)
    out["periodo"] = periodo
    out["municipio"] = out["nome_ibge"].fillna(out["municipio_estban"])
    out = out[out["codigo_ibge_7"].notna()].copy()
    out["codigo_ibge_7"] = out["codigo_ibge_7"].astype(int)
    out = out.sort_values("codigo_ibge_7").reset_index(drop=True)

    completo = g.rename(columns={"_cod": "codigo_bcb", "_uf": "uf", "_nome": "municipio_estban"})
    # inclui o código IBGE na aba de verbetes (permite cruzar crédito por modalidade com as outras fontes)
    completo = completo.merge(out[["codigo_bcb", "uf", "codigo_ibge_7"]], on=["codigo_bcb", "uf"], how="left")
    completo.insert(0, "codigo_ibge_7", completo.pop("codigo_ibge_7"))
    return out, sem[["codigo_bcb", "uf", "municipio_estban"]], completo, \
        {"credito": c_cred, "vista": c_vista, "prazo": c_prazo, "agencias": c_ag,
         "aprox": aprox}


def salva_xlsx(df, sem, completo, usados, caminho):
    fonte = Font(name="Arial", size=10)
    cab_f = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cab_fill = PatternFill("solid", fgColor="1F4E78")

    def cabecalho(ws, titulos):
        for j, t in enumerate(titulos, 1):
            c = ws.cell(row=1, column=j, value=t)
            c.font, c.fill = cab_f, cab_fill
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.row_dimensions[1].height = 40

    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    cols = [
        ("Código IBGE (7 dígitos)", "codigo_ibge_7", "0"),
        ("Município", "municipio", "@"),
        ("UF", "uf", "@"),
        ("Período (AAAAMM)", "periodo", "@"),
        ("Saldo de operações de crédito (R$)", "saldo_credito", "#,##0.00"),
        ("Crédito: empréstimos e títulos descontados (R$)", "credito_emprestimos", "#,##0.00"),
        ("Crédito: financiamentos (R$)", "credito_financiamentos", "#,##0.00"),
        ("Crédito rural (custeio e investimento) (R$)", "credito_rural", "#,##0.00"),
        ("Crédito agroindustrial (R$)", "credito_agroindustrial", "#,##0.00"),
        ("Crédito imobiliário (R$)", "credito_imobiliario", "#,##0.00"),
        ("Depósitos à vista (R$)", "depositos_vista", "#,##0.00"),
        ("Depósitos a prazo (R$)", "depositos_prazo", "#,##0.00"),
        ("Depósitos de poupança (R$)", "depositos_poupanca", "#,##0.00"),
        ("Total depósitos à vista + a prazo (R$)", "total_depositos", "#,##0.00"),
        ("Agências (ESTBAN)", "agencias", "#,##0"),
    ]
    cabecalho(ws, [c[0] for c in cols])
    for i, row in enumerate(df.to_dict("records"), start=2):
        for j, (t, k, fmt) in enumerate(cols, 1):
            v = row[k]
            v = None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.number_format = fonte, fmt
    for j, (t, _, _) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = 34 if t == "Município" else 20
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    # Aba com todos os verbetes (para recalcular com outro critério)
    wc = wb.create_sheet("ESTBAN completo")
    cabecalho(wc, list(completo.columns))
    for i, row in enumerate(completo.itertuples(index=False), start=2):
        for j, v in enumerate(row, 1):
            v = None if (isinstance(v, float) and pd.isna(v)) else v
            wc.cell(row=i, column=j, value=v).font = fonte

    wv = wb.create_sheet("Verbetes")      # dicionário: códigos de verbete -> nome da coluna
    cabecalho(wv, ["Códigos dos verbetes", "Nome da coluna em 'ESTBAN completo'"])
    for i, c in enumerate([c for c in completo.columns if c.startswith("VERBETE_")], start=2):
        wv.cell(row=i, column=1, value=abrevia(c)[0]).font = fonte
        wv.cell(row=i, column=2, value=c).font = fonte
    wv.column_dimensions["A"].width = 28
    wv.column_dimensions["B"].width = 120

    ap = usados.get("aprox")
    if ap is not None and len(ap):
        wa = wb.create_sheet("Pareamento aproximado")
        cabecalho(wa, list(ap.columns))
        for i, row in enumerate(ap.itertuples(index=False), start=2):
            for j, v in enumerate(row, 1):
                wa.cell(row=i, column=j, value=v).font = fonte
        for j in range(1, 5):
            wa.column_dimensions[get_column_letter(j)].width = 32

    if len(sem):
        wn = wb.create_sheet("Sem código IBGE")
        cabecalho(wn, list(sem.columns))
        for i, row in enumerate(sem.itertuples(index=False), start=2):
            for j, v in enumerate(row, 1):
                wn.cell(row=i, column=j, value=v).font = fonte

    nt = wb.create_sheet("Fontes e notas")
    notas = [
        "Fonte: Banco Central do Brasil - ESTBAN (Estatística Bancária Mensal e por Município).",
        f"Arquivo: {urls_estban('municipio', df['periodo'].iloc[0])[0]} (BCB; uma linha por instituição e município, somadas aqui por município)",
        "Valores em R$ (reais, conforme o arquivo do BCB). Saldos de fim do mês de referência (estoque).",
        "Poupança (verbete 420) fica em coluna separada e NÃO entra no total 'à vista + a prazo'.",
        "Crédito por modalidade: 161 empréstimos e títulos descontados; 162 financiamentos; 163 financiamentos rurais (custeio/investimento); 167 agroindustriais; 169 imobiliários.",
        "ATENÇÃO: o verbete 160 soma, por instituição, valores bem acima da carteira de crédito divulgada pelo BCB e concentra operações na sede dos bancos (Brasília, Osasco). Use para comparação relativa, não como estoque de crédito local.",
        f"Crédito = verbete(s) {abrevia(usados['credito'])} (160 - Operações de crédito).",
        f"Depósitos à vista = verbete(s) {abrevia(usados['vista'])}. No arquivo do BCB essa coluna agrega os verbetes 401 a 419 "
        "(à vista de governos e do setor privado, mais rubricas como depósitos judiciais, obrigatórios, para investimentos e demais depósitos).",
        f"Depósitos a prazo = verbete(s) {abrevia(usados['prazo'])} (432 - Depósitos a prazo).",
        f"Agências = coluna {usados['agencias']} (agências que enviaram dados no mês).",
        "O ESTBAN NÃO traz postos de atendimento bancário (PAB/PAE); para isso é preciso outra base do BCB (cadastro de agências e postos).",
        "O ESTBAN usa código de município do BCB; o código IBGE foi obtido por pareamento (UF + nome) com a API de localidades do IBGE.",
        "Banco/município com menos de 3 instituições pode ter dados omitidos pelo BCB por sigilo.",
    ]
    for i, t in enumerate(notas, 1):
        nt.cell(row=i, column=1, value=t).font = fonte
    nt.column_dimensions["A"].width = 140
    wb.save(caminho)
    print(f"Arquivo salvo: {caminho} ({len(df)} municípios)")


if __name__ == "__main__":
    # Aceita só AAAAMM; ignora argumentos do Jupyter (ex.: --f=kernel.json)
    args = [a for a in sys.argv[1:] if re.fullmatch(r"\d{6}", a)]
    per = args[0] if args else (f"{ANO_REF}12" if ANO_REF else None)
    salva_xlsx(*coleta(per), SAIDA)
# %%
