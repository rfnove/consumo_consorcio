#%%!/usr/bin/env python3
"""
Coleta dados municipais do IBGE (API SIDRA) usando o ANO MAIS RECENTE de cada informação e salva em XLSX.

  - PIB total e PIB per capita : último ano publicado (tabela 5938)
  - VAB por setor (agro, indústria, serviços, adm. pública): último ano com os 4 setores preenchidos.
      (O IBGE publicou o PIB 2022-2023 SEM a abertura setorial; por isso o VAB costuma estar defasado em relação ao PIB.)
  - População estimada          : última estimativa (tabela 6579) + população no ano do PIB (interpolada se faltar o ano)
  - Domicílios                  : Censo 2022 (tabela 4712)

Cada informação traz sua própria coluna de ano. Requisitos: pip install requests pandas openpyxl
"""
import re
import time

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SIDRA = "https://apisidra.ibge.gov.br/values"
AGREGADOS = "https://servicodados.ibge.gov.br/api/v3/agregados"
SAIDA = "dados_ibge_municipios.xlsx"
ANO_PIB = None        # None = último ano publicado; ou fixe, ex.: 2021
HEADERS = {"User-Agent": "Mozilla/5.0 (coleta-ibge)"}

PIB_TABELA = 5938
PIB_VARS = {
    37:   "pib_total_mil_reais",
    513:  "vab_agropecuaria_mil_reais",
    517:  "vab_industria_mil_reais",
    6575: "vab_servicos_exc_adm_mil_reais",
    525:  "vab_adm_publica_mil_reais",
}
POP_TABELA, POP_VAR = 6579, 9324
CENSO_TABELA, CENSO_POP_VAR = 4714, 93          # população recenseada 2022 (usada na interpolação)
DOM_TABELA, DOM_VAR = 4712, 381


def get(url, tentativas=4, timeout=120):
    for i in range(tentativas):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == tentativas - 1:
                raise RuntimeError(f"Falha ao consultar {url}: {e}")
            time.sleep(2 ** i)


def consulta_sidra(caminho):
    """Consulta a API SIDRA e devolve DataFrame com colunas renomeadas."""
    dados = get(f"{SIDRA}/{caminho}", timeout=180).json()
    cab, linhas = dados[0], dados[1:]
    if not linhas:
        return pd.DataFrame()
    df = pd.DataFrame(linhas)
    ren = {}
    for k, rot in cab.items():
        if rot.startswith("Município (Código)"):
            ren[k] = "cod_municipio"
        elif rot == "Município":
            ren[k] = "municipio_uf"
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
    df["cod_municipio"] = df["cod_municipio"].astype(int)
    if "ano" in df:
        df["ano"] = df["ano"].astype(int)
    return df


def anos_disponiveis(tabela):
    """Anos (períodos anuais) publicados numa tabela, pela API de agregados."""
    try:
        per = get(f"{AGREGADOS}/{tabela}/periodos").json()
        return sorted({int(str(p["id"])[:4]) for p in per})
    except Exception:
        return []


def variavel_per_capita():
    """Procura, nos metadados da tabela do PIB, a variável de PIB per capita (id e nome)."""
    try:
        m = get(f"{AGREGADOS}/{PIB_TABELA}/metadados").json()
        for v in m["variaveis"]:
            if "per capita" in v["nome"].lower():
                return v["id"], v["nome"]
    except Exception:
        pass
    return None, None


def serie_pop(ano):
    """População por município no ano (estimativa da tabela 6579; ou Censo 2022 para 2022)."""
    d = consulta_sidra(f"t/{POP_TABELA}/n6/all/v/{POP_VAR}/p/{ano}")
    if d.empty and ano == 2022:
        d = consulta_sidra(f"t/{CENSO_TABELA}/n6/all/v/{CENSO_POP_VAR}/p/2022")
    return d.set_index("cod_municipio")["valor"] if not d.empty else None


def pop_no_ano(ano, disponiveis):
    """População do ano pedido: exata se existir; senão interpolação linear entre os anos vizinhos."""
    s = serie_pop(ano) if ano in disponiveis or ano == 2022 else None
    if s is not None:
        return s, f"estimativa/censo {ano}"
    base = sorted(set(disponiveis) | {2022})
    abaixo = [a for a in base if a < ano]
    acima = [a for a in base if a > ano]
    a0 = abaixo[-1] if abaixo else None
    a1 = acima[0] if acima else None
    s0 = serie_pop(a0) if a0 else None
    s1 = serie_pop(a1) if a1 else None
    if s0 is not None and s1 is not None:
        w = (ano - a0) / (a1 - a0)
        return s0 * (1 - w) + s1 * w, f"interpolação {a0}-{a1}"
    if s0 is not None:
        return s0, f"estimativa {a0} (mais próxima)"
    if s1 is not None:
        return s1, f"estimativa {a1} (mais próxima)"
    return None, "indisponível"


def coleta():
    # ---------------- PIB total e per capita (último ano) ----------------
    print("Baixando PIB total (tabela 5938)...")
    ano_pib = ANO_PIB or int(consulta_sidra(f"t/{PIB_TABELA}/n6/all/v/37/p/last")["ano"].max())
    id_pc, nome_pc = variavel_per_capita()
    vars_pib = [37] + ([id_pc] if id_pc else [])
    pib = consulta_sidra(f"t/{PIB_TABELA}/n6/all/v/{','.join(map(str, vars_pib))}/p/{ano_pib}")
    print(f"  Ano do PIB: {ano_pib} | PIB per capita do IBGE: {nome_pc or 'não encontrado (será calculado)'}")
    base = pib[pib["cod_var"] == "37"].set_index("cod_municipio")["valor"].rename("pib_total_mil_reais").to_frame()
    nomes = pib.drop_duplicates("cod_municipio").set_index("cod_municipio")["municipio_uf"]
    base.insert(0, "municipio", nomes.str.rsplit(" - ", n=1).str[0])
    base.insert(1, "uf", nomes.str.rsplit(" - ", n=1).str[1])
    base["ano_pib"] = ano_pib
    if id_pc:
        pc = pib[pib["cod_var"] == str(id_pc)].set_index("cod_municipio")["valor"]
        base["pib_per_capita_ibge"] = pc

    # ---------------- VAB setorial (último ano completo) ----------------
    print("Procurando o último ano com VAB setorial completo...")
    vab_ids = [v for v in PIB_VARS if v != 37]
    vab = None
    ano_vab = None
    for ano in range(ano_pib, ano_pib - 8, -1):
        tmp = consulta_sidra(f"t/{PIB_TABELA}/n6/all/v/{','.join(map(str, vab_ids))}/p/{ano}")
        if tmp.empty:
            continue
        validos = tmp.groupby("cod_var")["valor"].count().reindex([str(v) for v in vab_ids]).fillna(0)
        print(f"  Ano {ano}: valores válidos -> {validos.astype(int).to_dict()}")
        if (validos > 0.5 * tmp["cod_municipio"].nunique()).all():
            vab, ano_vab = tmp, ano
            break
    if vab is not None:
        w = vab.pivot_table(index="cod_municipio", columns="cod_var", values="valor", aggfunc="first")
        w.columns = [PIB_VARS[int(c)] for c in w.columns]
        base = base.join(w)
    else:
        print("  ATENÇÃO: nenhum ano com VAB setorial completo; colunas de VAB ficam vazias.")
        for c in PIB_VARS.values():
            if c != "pib_total_mil_reais":
                base[c] = None
    base["ano_vab"] = ano_vab
    print(f"  Ano do VAB setorial: {ano_vab}")

    # ---------------- População ----------------
    print("Baixando população estimada (tabela 6579)...")
    disp = anos_disponiveis(POP_TABELA)
    pop_ult = consulta_sidra(f"t/{POP_TABELA}/n6/all/v/{POP_VAR}/p/last")
    base["populacao_estimada"] = pop_ult.set_index("cod_municipio")["valor"]
    base["ano_populacao"] = int(pop_ult["ano"].max())
    pop_p, metodo = pop_no_ano(ano_pib, disp)
    base["populacao_ano_pib"] = pop_p
    base["metodo_pop_ano_pib"] = metodo
    print(f"  Última estimativa: {base['ano_populacao'].iloc[0]} | população no ano do PIB: {metodo}")

    # ---------------- Domicílios (Censo 2022) ----------------
    print("Baixando domicílios do Censo (tabela 4712)...")
    dom = consulta_sidra(f"t/{DOM_TABELA}/n6/all/v/{DOM_VAR}/p/last")
    if not dom.empty:
        print("  Variável retornada:", dom["nome_var"].iloc[0], "| ano:", dom["ano"].iloc[0])
        base["domicilios_particulares_permanentes"] = dom.set_index("cod_municipio")["valor"]
        base["ano_censo"] = int(dom["ano"].max())
    else:
        base["domicilios_particulares_permanentes"] = None
        base["ano_censo"] = None

    base = base.reset_index().rename(columns={"cod_municipio": "codigo_ibge_7"})
    base["codigo_ibge_6"] = base["codigo_ibge_7"] // 10
    pc_calc = base["pib_total_mil_reais"] * 1000 / base["populacao_ano_pib"]
    base["pib_per_capita_reais"] = (base["pib_per_capita_ibge"] if "pib_per_capita_ibge" in base else pc_calc)
    base["pib_per_capita_reais"] = base["pib_per_capita_reais"].fillna(pc_calc).round(2)
    base["vab_servicos_total_mil_reais"] = base[
        ["vab_servicos_exc_adm_mil_reais", "vab_adm_publica_mil_reais"]].sum(axis=1, min_count=1)
    return base.sort_values("codigo_ibge_7").reset_index(drop=True)


def salva_xlsx(df, caminho):
    cols = [
        ("Código IBGE (7 dígitos)", "codigo_ibge_7", "0"),
        ("Código IBGE (6 dígitos)", "codigo_ibge_6", "0"),
        ("Município", "municipio", "@"),
        ("UF", "uf", "@"),
        ("População estimada", "populacao_estimada", "#,##0"),
        ("Ano da estimativa", "ano_populacao", "0"),
        ("População (ano do PIB)", "populacao_ano_pib", "#,##0"),
        ("Método da população (ano do PIB)", "metodo_pop_ano_pib", "@"),
        ("Ano do PIB", "ano_pib", "0"),
        ("PIB total (R$ mil)", "pib_total_mil_reais", "#,##0"),
        ("PIB per capita (R$)", "pib_per_capita_reais", "#,##0.00"),
        ("Ano do VAB", "ano_vab", "0"),
        ("VAB Agropecuária (R$ mil)", "vab_agropecuaria_mil_reais", "#,##0"),
        ("VAB Indústria (R$ mil)", "vab_industria_mil_reais", "#,##0"),
        ("VAB Serviços exc. adm. pública (R$ mil)", "vab_servicos_exc_adm_mil_reais", "#,##0"),
        ("VAB Adm. pública (R$ mil)", "vab_adm_publica_mil_reais", "#,##0"),
        ("VAB Serviços total (R$ mil)", "vab_servicos_total_mil_reais", "#,##0"),
        ("Domicílios particulares permanentes", "domicilios_particulares_permanentes", "#,##0"),
        ("Ano do Censo", "ano_censo", "0"),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    fonte = Font(name="Arial", size=10)
    cab_f = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cab_fill = PatternFill("solid", fgColor="1F4E78")
    for j, (titulo, _, _) in enumerate(cols, 1):
        c = ws.cell(row=1, column=j, value=titulo)
        c.font, c.fill = cab_f, cab_fill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = 30 if titulo == "Município" else 18
    ws.row_dimensions[1].height = 45
    for i, row in enumerate(df.to_dict("records"), start=2):
        for j, (_, chave, fmt) in enumerate(cols, 1):
            v = row[chave]
            if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NA:
                v = None
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.number_format = fonte, fmt
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = ws.dimensions

    nt = wb.create_sheet("Fontes e notas")
    a_pib, a_vab = int(df["ano_pib"].iloc[0]), df["ano_vab"].iloc[0]
    notas = [
        "Fonte: IBGE - API SIDRA (apisidra.ibge.gov.br). Cada informação usa o ano mais recente disponível; veja as colunas de ano.",
        f"PIB total e per capita: tabela {PIB_TABELA}, ano {a_pib}, valores correntes em R$ mil (per capita em R$).",
        f"VAB setorial: ano {int(a_vab) if pd.notna(a_vab) else 'indisponível'}. O IBGE publicou o PIB 2022-2023 sem a abertura do VAB por setor "
        "(reformulação do sistema de contas); por isso o VAB é de um ano anterior ao do PIB. Use as participações setoriais como perfil estrutural.",
        f"População estimada: tabela {POP_TABELA}, variável {POP_VAR}. 'População (ano do PIB)' vem da estimativa do próprio ano, "
        "ou é interpolada linearmente entre os anos vizinhos (inclui o Censo 2022) quando o ano não existe (coluna 'Método').",
        f"Domicílios: tabela {DOM_TABELA}, variável {DOM_VAR}, Censo 2022 (domicílios particulares permanentes ocupados).",
        "PIB per capita: usa a variável do IBGE quando existe; senão, PIB x 1.000 / população do ano do PIB.",
        "VAB Serviços total = Serviços (exclusive adm. pública) + Adm., defesa, educação e saúde públicas.",
        "Código de 6 dígitos = código de 7 dígitos sem o dígito verificador.",
    ]
    for i, t in enumerate(notas, 1):
        nt.cell(row=i, column=1, value=t).font = fonte
    nt.column_dimensions["A"].width = 150
    wb.save(caminho)
    print(f"\nArquivo salvo: {caminho} ({len(df)} municípios)")


if __name__ == "__main__":
    salva_xlsx(coleta(), SAIDA)
# %%
