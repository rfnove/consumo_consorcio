#%%!/usr/bin/env python3
"""
Coleta o Novo CAGED (MTE / PDET) por município, com os DADOS MAIS RECENTES, e salva em XLSX.

Fonte: microdados públicos no FTP do PDET
    ftp://ftp.mtps.gov.br/pdet/microdados/NOVO CAGED/AAAA/AAAAMM/CAGED{MOV,FOR,EXC}AAAAMM.7z
        MOV = movimentações no prazo | FOR = fora do prazo (competências anteriores) | EXC = exclusões (peso -1)

O que o script faz sozinho:
  1. descobre o ÚLTIMO mês publicado do Novo CAGED (competência final);
  2. descobre o ÚLTIMO ano da RAIS (arquivo de estabelecimentos) para usar como base do estoque;
  3. lê todos os meses desde 01/(ano da RAIS + 1) até a competência final e calcula:
       - admissões, desligamentos, saldo e salários médios dos ÚLTIMOS 12 MESES;
       - estoque CLT estimado na competência final = vínculos CLT da RAIS em 31/12 + saldo acumulado do CAGED.

Atenção: pode baixar de 15 a 25 meses de arquivos (~55 MB cada MOV) -> alguns GB no cache e bastante tempo.
Os últimos meses ainda podem ser revisados pelo MTE.

Requisitos: pip install requests pandas openpyxl py7zr
Uso:        python coletar_caged.py                 (último mês disponível)
            python coletar_caged.py 202608          (competência final específica)
"""
import os
import re
import shutil
import sys
import tempfile
from ftplib import FTP, error_perm

import pandas as pd
import requests
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

JANELA = 12                 # meses usados nos fluxos (admissões, desligamentos, saldo, salários)
CALCULAR_ESTOQUE = True     # baixa a RAIS mais recente como base do estoque
SALARIO_MAX = 100_000       # salários acima disso (R$/mês) são ignorados nas médias (erros de digitação)
CACHE_DIR = "cache_caged"
SAIDA = "dados_caged_municipios.xlsx"

FTP_HOST = "ftp.mtps.gov.br"
DIR_CAGED = "pdet/microdados/NOVO CAGED"
DIR_RAIS = "pdet/microdados/RAIS"
URL_IBGE_MUN = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios"
HEADERS = {"User-Agent": "Mozilla/5.0 (coleta-caged)"}


def slug(c):
    t = unicodedata.normalize("NFKD", str(c)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")


def soma_meses(ym, k):
    a, m = int(ym[:4]), int(ym[4:])
    t = a * 12 + (m - 1) + k
    return f"{t // 12}{t % 12 + 1:02d}"


def meses(ini, fim):
    out, ym = [], ini
    while ym <= fim:
        out.append(ym)
        ym = soma_meses(ym, 1)
    return out


# ------------------------------------------------------------------ FTP / arquivos
def conecta():
    ftp = FTP(FTP_HOST, timeout=120)
    ftp.encoding = "latin-1"  # Adicione esta linha para corrigir a leitura do diretório
    ftp.login()
    ftp.set_pasv(True)
    return ftp

def lista(ftp, pasta):
    ftp.cwd("/")
    ftp.cwd(pasta)
    return [os.path.basename(n.rstrip("/")) for n in ftp.nlst()]


def ultimo_mes_caged(ftp):
    anos = sorted([n for n in lista(ftp, DIR_CAGED) if re.fullmatch(r"20\d\d", n)], reverse=True)
    for a in anos[:3]:
        for ym in sorted([n for n in lista(ftp, f"{DIR_CAGED}/{a}") if re.fullmatch(rf"{a}(0[1-9]|1[0-2])", n)],
                         reverse=True):
            if any(n.upper().startswith("CAGEDMOV") for n in lista(ftp, f"{DIR_CAGED}/{a}/{ym}")):
                return ym
    raise RuntimeError("Nenhum mês do Novo CAGED encontrado no FTP.")


def ultimo_ano_rais(ftp):
    anos = sorted([n for n in lista(ftp, DIR_RAIS) if re.fullmatch(r"20\d\d", n)], reverse=True)
    for a in anos[:5]:
        arqs = [n for n in lista(ftp, f"{DIR_RAIS}/{a}") if re.search(r"(?i)estab|estb", n) and n.lower().endswith(".7z")]
        if arqs:
            return int(a), arqs[0]
    raise RuntimeError("Nenhum arquivo de estabelecimentos da RAIS encontrado no FTP.")


def baixa(ftp, pasta, nome, destino):
    if os.path.exists(destino) and os.path.getsize(destino) > 0:
        return True
    tmp = destino + ".parte"
    try:
        ftp.cwd("/")
        ftp.cwd(pasta)
        with open(tmp, "wb") as f:
            ftp.retrbinary(f"RETR {nome}", f.write, blocksize=1 << 20)
        os.replace(tmp, destino)
        return True
    except error_perm as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        print(f"    (não encontrado: {pasta}/{nome} -> {e})")
        return False


def extrai(arq7z, pasta_tmp):
    try:
        import py7zr
    except ImportError as e:
        raise RuntimeError("Instale o py7zr para ler os arquivos .7z:  pip install py7zr") from e
    with py7zr.SevenZipFile(arq7z, "r") as z:
        z.extractall(path=pasta_tmp)
    arqs = [os.path.join(pasta_tmp, n) for n in os.listdir(pasta_tmp) if n.lower().endswith((".txt", ".csv"))]
    if not arqs:
        raise RuntimeError(f"Nenhum .txt dentro de {arq7z}")
    return arqs[0]


# ------------------------------------------------------------------ processamento CAGED
def agrega_caged(txt, peso, ini, fim):
    """Lê um arquivo CAGED em blocos; devolve agregados por (competência, município) dentro de [ini, fim]."""
    alvo = {"competenciamov", "municipio", "saldomovimentacao", "salario"}
    partes = []
    for ch in pd.read_csv(txt, sep=";", decimal=",", encoding="utf-8", chunksize=500_000,
                          usecols=lambda c: slug(c) in alvo, low_memory=False):
        ch.columns = [slug(c) for c in ch.columns]
        ch = ch[(ch["competenciamov"] >= int(ini)) & (ch["competenciamov"] <= int(fim))]
        if ch.empty:
            continue
        sal = pd.to_numeric(ch["salario"], errors="coerce")
        valido = (sal > 0) & (sal <= SALARIO_MAX)
        adm = (ch["saldomovimentacao"] == 1) * peso
        des = (ch["saldomovimentacao"] == -1) * peso
        partes.append(pd.DataFrame({
            "competencia": ch["competenciamov"], "cod6": ch["municipio"],
            "admissoes": adm, "desligamentos": des,
            "adm_sal_n": adm * valido, "adm_sal_soma": adm * valido * sal.fillna(0),
            "des_sal_n": des * valido, "des_sal_soma": des * valido * sal.fillna(0),
        }).groupby(["competencia", "cod6"], as_index=False).sum())
    if not partes:
        return None
    return pd.concat(partes).groupby(["competencia", "cod6"], as_index=False).sum()


def coleta_caged(ftp, ini, fim):
    os.makedirs(CACHE_DIR, exist_ok=True)
    acumulado = []
    for ym in meses(ini, fim):
        for tipo, peso in (("MOV", 1), ("FOR", 1), ("EXC", -1)):
            nome = f"CAGED{tipo}{ym}.7z"
            destino = os.path.join(CACHE_DIR, nome)
            print(f"  {nome} ...")
            try:
                ok = baixa(ftp, f"{DIR_CAGED}/{ym[:4]}/{ym}", nome, destino)
            except Exception as e:
                print(f"    erro ({e}); reconectando...")
                ftp = conecta()
                ok = baixa(ftp, f"{DIR_CAGED}/{ym[:4]}/{ym}", nome, destino)
            if not ok:
                continue
            tmp = tempfile.mkdtemp(prefix="caged_")
            try:
                g = agrega_caged(extrai(destino, tmp), peso, ini, fim)
                if g is not None:
                    acumulado.append(g)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
    if not acumulado:
        raise RuntimeError("Nenhum dado do Novo CAGED foi lido.")
    return pd.concat(acumulado).groupby(["competencia", "cod6"], as_index=False).sum()


# ------------------------------------------------------------------ RAIS (base do estoque)
def detecta_encoding(caminho):
    with open(caminho, "rb") as f:
        amostra = f.read(200_000)
    try:
        amostra.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError as e:
        return "utf-8" if e.start >= len(amostra) - 4 else "latin-1"


def estoque_rais(ftp, ano_base, nome):
    """Soma de vínculos celetistas ativos em 31/12 por município (RAIS - estabelecimentos)."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    destino = os.path.join(CACHE_DIR, f"RAIS{ano_base}_{nome}")
    print(f"  RAIS {ano_base}: {nome} ...")
    baixa(ftp, f"{DIR_RAIS}/{ano_base}", nome, destino)
    tmp = tempfile.mkdtemp(prefix="rais_")
    try:
        txt = extrai(destino, tmp)
        enc = detecta_encoding(txt)
        cab = pd.read_csv(txt, sep=";", encoding=enc, nrows=0).columns.tolist()
        slugs = {c: slug(c) for c in cab}
        c_mun = next((c for c, sl in slugs.items() if re.fullmatch(r"munic\w*pio", sl)), None)
        c_clt = next((c for c, sl in slugs.items() if re.fullmatch(r"qtd_v\w*nculos_clt", sl)), None)
        c_ativos = next((c for c, sl in slugs.items() if re.fullmatch(r"qtd_v\w*nculos_ativos", sl)), None)
        print(f"  Colunas da RAIS: município={c_mun!r}, CLT={c_clt!r}, ativos={c_ativos!r}")
        if c_clt is None and c_ativos is not None:
            print("  ATENÇÃO: sem coluna de vínculos CLT; usando 'Qtd Vínculos Ativos' (inclui estatutários).")
            c_clt = c_ativos
        if c_mun is None or c_clt is None:
            raise RuntimeError(f"Colunas de município/vínculos não identificadas. Cabeçalho: {cab}")
        partes = []
        for ch in pd.read_csv(txt, sep=";", decimal=",", encoding=enc, chunksize=500_000,
                              usecols=[c_mun, c_clt], low_memory=False):
            ch.columns = ["cod6", "estoque_base"]
            partes.append(ch.groupby("cod6", as_index=False)["estoque_base"].sum())
        return pd.concat(partes).groupby("cod6", as_index=False)["estoque_base"].sum()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def tabela_ibge():
    r = requests.get(URL_IBGE_MUN, headers=HEADERS, timeout=120)
    r.raise_for_status()
    linhas = []
    for m in r.json():
        uf = (m.get("microrregiao") or {}).get("mesorregiao", {}).get("UF", {}).get("sigla") \
            or m["regiao-imediata"]["regiao-intermediaria"]["UF"]["sigla"]
        linhas.append({"codigo_ibge_7": m["id"], "cod6": m["id"] // 10, "municipio": m["nome"], "uf": uf})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------ montagem
def monta(fim=None):
    ftp = conecta()
    fim = fim or ultimo_mes_caged(ftp)
    ini_janela = soma_meses(fim, -(JANELA - 1))
    base_rais, ano_rais, ini_estoque = None, None, None
    rais_info = None
    if CALCULAR_ESTOQUE:
        try:
            rais_info = ultimo_ano_rais(ftp)
            ano_rais = rais_info[0]
            ini_estoque = f"{ano_rais + 1}01"
        except Exception as e:
            print(f"  ATENÇÃO: RAIS indisponível ({e}); sem estoque.")
    ini = min(ini_janela, ini_estoque) if ini_estoque else ini_janela
    print(f"Competência final: {fim} | janela de fluxos: {ini_janela}-{fim} | RAIS base: {ano_rais} | "
          f"meses a ler: {len(meses(ini, fim))}")

    mov = coleta_caged(ftp, ini, fim)
    ibge = tabela_ibge()
    fora = set(mov["cod6"]) - set(ibge["cod6"])
    if fora:
        sub_f = mov[mov["cod6"].isin(fora)].copy()
        sub_f["mov"] = sub_f["admissoes"] + sub_f["desligamentos"]
        top = sub_f.groupby("cod6")["mov"].sum().sort_values(ascending=False).head(5)
        print(f"  Aviso: {len(fora)} código(s) de município fora da tabela do IBGE ({int(sub_f['mov'].sum())} movimentos); "
              "ignorados. Principais: " + ", ".join(f"{int(c)} ({int(v)} mov.)" for c, v in top.items()))

    mm = mov.merge(ibge, on="cod6", how="inner")
    mm["saldo"] = mm["admissoes"] - mm["desligamentos"]
    mm["sal_adm_medio"] = mm["adm_sal_soma"] / mm["adm_sal_n"].where(mm["adm_sal_n"] > 0)
    mm["sal_des_medio"] = mm["des_sal_soma"] / mm["des_sal_n"].where(mm["des_sal_n"] > 0)

    soma = ["admissoes", "desligamentos", "adm_sal_n", "adm_sal_soma", "des_sal_n", "des_sal_soma"]
    j12 = mov[mov["competencia"] >= int(ini_janela)].groupby("cod6", as_index=False)[soma].sum()
    dados = ibge.merge(j12, on="cod6", how="left")
    for c in soma:
        dados[c] = dados[c].fillna(0)          # município sem movimentação formal = 0
    dados["saldo_12m"] = dados["admissoes"] - dados["desligamentos"]
    dados["sal_adm_medio"] = dados["adm_sal_soma"] / dados["adm_sal_n"].where(dados["adm_sal_n"] > 0)
    dados["sal_des_medio"] = dados["des_sal_soma"] / dados["des_sal_n"].where(dados["des_sal_n"] > 0)

    dados["estoque_base_rais"] = pd.NA
    dados["saldo_acumulado"] = pd.NA
    dados["estoque_estimado"] = pd.NA
    if rais_info:
        try:
            print(f"Baixando a base de estoque (RAIS {ano_rais})...")
            base = estoque_rais(ftp, *rais_info)
            acum = mov[mov["competencia"] >= int(ini_estoque)].assign(s=lambda d: d["admissoes"] - d["desligamentos"])
            acum = acum.groupby("cod6", as_index=False)["s"].sum().rename(columns={"s": "saldo_acumulado"})
            dados = dados.drop(columns=["estoque_base_rais", "saldo_acumulado", "estoque_estimado"]) \
                .merge(base, on="cod6", how="left").merge(acum, on="cod6", how="left")
            dados["estoque_base"] = dados["estoque_base"].fillna(0)
            dados["saldo_acumulado"] = dados["saldo_acumulado"].fillna(0)
            dados = dados.rename(columns={"estoque_base": "estoque_base_rais"})
            dados["estoque_estimado"] = dados["estoque_base_rais"] + dados["saldo_acumulado"]
        except Exception as e:
            print(f"  ATENÇÃO: não foi possível montar o estoque ({e}). Colunas de estoque ficam vazias.")
    try:
        ftp.quit()
    except Exception:
        pass
    dados["competencia_final"] = fim
    dados["inicio_janela"] = ini_janela
    mm["competencia_final"] = fim
    meta = {"fim": fim, "ini_janela": ini_janela, "ano_rais": ano_rais, "ini_estoque": ini_estoque}
    return dados.sort_values("codigo_ibge_7").reset_index(drop=True), \
        mm.sort_values(["codigo_ibge_7", "competencia"]).reset_index(drop=True), meta


def salva_xlsx(dados, mm, meta, caminho):
    fonte = Font(name="Arial", size=10)
    cab_f = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cab_fill = PatternFill("solid", fgColor="1F4E78")
    fim, ij, ar = meta["fim"], meta["ini_janela"], meta["ano_rais"]

    def folha(ws, tab, colunas, freeze="C2"):
        for j, (t, _, _) in enumerate(colunas, 1):
            c = ws.cell(row=1, column=j, value=t)
            c.font, c.fill = cab_f, cab_fill
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            ws.column_dimensions[get_column_letter(j)].width = 30 if t == "Município" else 18
        ws.row_dimensions[1].height = 48
        for i, row in enumerate(tab.to_dict("records"), start=2):
            for j, (_, k, fmt) in enumerate(colunas, 1):
                v = row.get(k)
                if v is pd.NA or (isinstance(v, float) and pd.isna(v)):
                    v = None
                c = ws.cell(row=i, column=j, value=v)
                c.font, c.number_format = fonte, fmt
        ws.freeze_panes = freeze
        ws.auto_filter.ref = ws.dimensions

    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    folha(ws, dados, [
        ("Código IBGE (7 dígitos)", "codigo_ibge_7", "0"), ("Município", "municipio", "@"), ("UF", "uf", "@"),
        ("Competência final (AAAAMM)", "competencia_final", "0"),
        ("Início da janela de 12 meses (AAAAMM)", "inicio_janela", "0"),
        (f"Estoque CLT RAIS 31/12/{ar} (base)", "estoque_base_rais", "#,##0"),
        (f"Admissões (12 meses até {fim})", "admissoes", "#,##0"),
        (f"Desligamentos (12 meses até {fim})", "desligamentos", "#,##0"),
        (f"Saldo (12 meses até {fim})", "saldo_12m", "#,##0;-#,##0"),
        (f"Saldo acumulado desde 31/12/{ar}", "saldo_acumulado", "#,##0;-#,##0"),
        (f"Estoque CLT estimado em {fim} (base RAIS + saldo CAGED)", "estoque_estimado", "#,##0"),
        (f"Salário médio de admissão (12 meses até {fim}) (R$)", "sal_adm_medio", "#,##0.00"),
        (f"Salário médio de desligamento (12 meses até {fim}) (R$)", "sal_des_medio", "#,##0.00"),
    ])
    wm = wb.create_sheet("Movimentação mensal")
    folha(wm, mm, [
        ("Código IBGE (7 dígitos)", "codigo_ibge_7", "0"), ("Município", "municipio", "@"), ("UF", "uf", "@"),
        ("Competência (AAAAMM)", "competencia", "0"),
        ("Admissões", "admissoes", "#,##0"), ("Desligamentos", "desligamentos", "#,##0"),
        ("Saldo", "saldo", "#,##0;-#,##0"),
        ("Salário médio de admissão (R$)", "sal_adm_medio", "#,##0.00"),
        ("Salário médio de desligamento (R$)", "sal_des_medio", "#,##0.00"),
    ], freeze="E2")

    nt = wb.create_sheet("Fontes e notas")
    notas = [
        "Fonte: MTE / PDET - microdados públicos do Novo CAGED (ftp://ftp.mtps.gov.br/pdet/microdados/NOVO CAGED/) e RAIS (estabelecimentos).",
        f"Competência final = último mês publicado ({fim}). Fluxos = janela móvel de {JANELA} meses ({ij} a {fim}); os meses mais recentes "
        "ainda podem ser revisados pelo MTE.",
        "Movimentação = CAGEDMOV (no prazo) + CAGEDFOR (fora do prazo) - CAGEDEXC (exclusões), todos os arquivos publicados até a competência final.",
        "Saldo = admissões - desligamentos (apenas vínculos celetistas).",
        f"Salário médio de admissão/desligamento = média do salário contratual (R$/mês) dos movimentos, ignorando valores <= 0 ou > R$ {SALARIO_MAX:,}. "
        "Atenção: a janela de 12 meses atravessa dois valores de salário mínimo.",
        "ATENÇÃO - o CAGED só informa o salário de quem foi admitido/desligado no período, não o salário médio de todos os empregados; para esse "
        "indicador use o CEMPRE (coletar_pendentes.py).",
        f"Estoque estimado = vínculos CLT ativos em 31/12/{ar} (RAIS, estabelecimentos) + saldo do CAGED desde 01/{ar + 1 if ar else ''} até {fim}. "
        "É aproximação: RAIS e CAGED têm coberturas e cortes diferentes e o erro se acumula com o número de meses.",
        "Municípios sem nenhuma movimentação aparecem com 0 admissões/desligamentos e salário vazio.",
        "O CAGED não cobre estatutários, militares nem informais. Local do movimento = município do estabelecimento empregador.",
    ]
    for i, t in enumerate(notas, 1):
        nt.cell(row=i, column=1, value=t).font = fonte
    nt.column_dimensions["A"].width = 160
    wb.save(caminho)
    print(f"Arquivo salvo: {caminho} ({len(dados)} municípios)")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if re.fullmatch(r"20\d\d(0[1-9]|1[0-2])", a)]
    salva_xlsx(*monta(args[0] if args else None), SAIDA)
# %%
