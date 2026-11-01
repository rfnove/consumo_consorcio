#%%!/usr/bin/env python3
"""
IPCM - Índice de Potencial de Consórcio Municipal (versão de viabilidade).

Lê a aba "Dados" do base_integrada.xlsx e calcula, por município e por segmento
(geral, imóveis, auto/moto, pesados/agro):

  * 5 pilares (0-100): massa salarial, demografia, frota, restrição de crédito, PIB per capita
  * IPCM (0-100) = média ponderada dos pilares (pesos configuráveis por segmento)
  * Quadrante "Oceano Azul" = demanda alta + crédito restrito (filtro, não soma)
  * rankings nacional e por UF, sinalizadores de qualidade (sede bancária, sem agência, município pequeno)

Método (por variável): log nas variáveis de tamanho -> limite nos percentis 1% e 99% -> posto percentual (0-1).
O posto percentual evita que São Paulo/Brasília dominem a escala. Variáveis "inversas" (menos = mais potencial)
usam o posto invertido, com os empates no extremo recebendo a nota máxima (município sem agência = restrição máxima).
Pilar = média ponderada das variáveis disponíveis (exige >= 50% do peso disponível).

Saídas: ipcm_resultado.xlsx, ipcm_resultado.csv, ipcm_dashboard.json
Requisitos: pip install pandas numpy openpyxl
Uso:        python calcular_ipcm.py [caminho_da_base.xlsx]
"""
import json
import re
import sys

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ARQ_ENTRADA = "base_integrada.xlsx"
SAIDA_XLSX, SAIDA_CSV, SAIDA_JSON = "ipcm_resultado.xlsx", "ipcm_resultado.csv", "ipcm_dashboard.json"

# ============================================================================ PARÂMETROS
PILARES = ["massa", "demografia", "frota", "credito", "pib"]
ROTULOS = {"massa": "Massa salarial", "demografia": "Demografia", "frota": "Frota alvo",
           "credito": "Restrição de crédito", "pib": "PIB per capita"}
SEGMENTOS = ["geral", "imoveis", "auto_moto", "pesados_agro"]
PESOS = {   # pontos (normalizados na hora); AJUSTE AQUI
    "geral":        {"massa": 25, "demografia": 15, "frota": 20, "credito": 25, "pib": 15},
    "imoveis":      {"massa": 30, "demografia": 20, "frota": 5,  "credito": 25, "pib": 20},
    "auto_moto":    {"massa": 25, "demografia": 10, "frota": 35, "credito": 20, "pib": 10},
    "pesados_agro": {"massa": 15, "demografia": 5,  "frota": 35, "credito": 25, "pib": 20},
}
HQ_QUANTIL = 0.99          # crédito per capita acima deste percentil (entre municípios com agência) = sede bancária
HQ_MANUAL = [5300108, 3534401]   # Brasília e Osasco (sedes de BB/Caixa/BRB e Bradesco)
MIN_POP_OCEANO = 10_000    # população mínima para entrar no quadrante Oceano Azul
Q_DEMANDA, Q_RESTRICAO = 0.60, 0.60   # percentis mínimos de demanda e de restrição para o Oceano Azul
MIN_PILARES = 4            # pilares mínimos para calcular o IPCM


def v(col, w, log=False, inv=False, na=None, hq=False):
    """na: None | 'zero' (sem dado = 0) | 'max' (sem dado = nota máxima); hq: neutraliza em sedes bancárias."""
    return {"col": col, "w": w, "log": log, "inv": inv, "na": na, "hq": hq}


_MASSA = [v("Salários e outras remunerações (R$ mil/ano)", .40, log=True),
          v("Massa salarial per capita (R$/mês)", .40),
          v("Salário de admissão (em salários mínimos)", .20)]
_DEMO = [v("População estimada", .50, log=True), v("Densidade demográfica 2022 (hab/km²)", .20, log=True),
         v("Grau de urbanização 2022", .15), v("Crescimento populacional anual 2010-2022 (CAGR)", .15)]
_FROTA_GERAL = [v("Total da frota", .30, log=True), v("Veículos por mil hab.", .30),
                v("Frota com 10 anos ou mais (unidades)", .25, log=True),
                v("% da frota com 10 anos ou mais (renovável)", .15)]
_CRED_BASE = [v("Agências por 100 mil hab.", .45, inv=True, na="zero"),
              v("Habitantes por agência", .20, na="max")]


def _cred(extra_col):
    return _CRED_BASE + [v(extra_col, .35, log=True, inv=True, na="zero", hq=True)]


VARIAVEIS = {
    "geral": {"massa": _MASSA, "demografia": _DEMO, "frota": _FROTA_GERAL, "credito": _cred("Crédito per capita (R$)"),
              "pib": [v("PIB per capita (R$)", 1.0, log=True)]},
    "imoveis": {"massa": _MASSA, "demografia": _DEMO, "frota": _FROTA_GERAL,
                "credito": _cred("Crédito imobiliário per capita (R$)"),
                "pib": [v("PIB per capita (R$)", .70, log=True), v("Crescimento populacional anual 2010-2022 (CAGR)", .30)]},
    "auto_moto": {"massa": _MASSA, "demografia": _DEMO,
                  "frota": [v("Automóveis e utilitários", .20, log=True), v("Motocicletas", .20, log=True),
                            v("Automóveis e utilitários por mil hab.", .15), v("Motocicletas por mil hab.", .15),
                            v("Frota com 10 anos ou mais (unidades)", .20, log=True),
                            v("% da frota com 10 anos ou mais (renovável)", .10)],
                  "credito": _cred("Crédito per capita (R$)"), "pib": [v("PIB per capita (R$)", 1.0, log=True)]},
    "pesados_agro": {"massa": _MASSA, "demografia": _DEMO,
                     "frota": [v("Caminhões, tratores e reboques", .40, log=True),
                               v("Caminhões, tratores e reboques por mil hab.", .40),
                               v("% da frota com 10 anos ou mais (renovável)", .20)],
                     "credito": _cred("Crédito rural e agroindustrial per capita (R$)"),
                     "pib": [v("PIB per capita (R$)", .50, log=True), v("% VAB Agropecuária", .30),
                             v("% VAB Indústria", .20)]},
}
COL_COD, COL_NOME, COL_UF, COL_POP = "Código IBGE (7 dígitos)", "Município", "UF", "População estimada"


# ============================================================================ funções
def posto(s, log, inv):
    x = np.log1p(s.clip(lower=0)) if log else s.astype(float)
    lo, hi = x.quantile(.01), x.quantile(.99)
    x = x.clip(lo, hi)
    return x.rank(pct=True, ascending=False, method="max") if inv else x.rank(pct=True, method="average")


def pilar_score(d, itens, hq):
    cols, pesos = [], []
    for it in itens:
        if it["col"] not in d.columns:
            print(f"    aviso: coluna ausente '{it['col']}' (ignorada)")
            continue
        s = d[it["col"]].astype(float)
        if it["na"] == "zero":
            s = s.fillna(0)
        if it["hq"]:
            s = s.mask(hq)
        p = posto(s, it["log"], it["inv"])
        if it["na"] == "max":
            p = p.fillna(1.0)
        cols.append(p)
        pesos.append(it["w"])
    if not cols:
        return pd.Series(np.nan, index=d.index)
    m = pd.concat(cols, axis=1)
    w = np.array(pesos)
    disp = m.notna().values * w
    num = (m.fillna(0).values * w).sum(axis=1)
    den = disp.sum(axis=1)
    out = np.where(den >= 0.5 * w.sum(), num / np.where(den > 0, den, 1), np.nan)
    return pd.Series(out * 100, index=d.index)


def combina(pil, pesos):
    w = pd.Series({k: pesos[k] for k in PILARES}, dtype=float)
    m = pil[PILARES]
    disp = (m.notna() * w).sum(axis=1)
    num = (m.fillna(0) * w).sum(axis=1)
    sc = (num / disp.where(disp > 0))
    return sc.where(m.notna().sum(axis=1) >= MIN_PILARES)


# ============================================================================ cálculo
def calcula(d):
    d = d.copy()
    n = len(d)
    # --- sedes bancárias e municípios sem agência ---
    tem = d["Tem dado ESTBAN"].astype(bool)
    cpc = d["Crédito per capita (R$)"]
    corte = cpc[tem & (cpc > 0)].quantile(HQ_QUANTIL)
    hq = (tem & (cpc > corte)) | d[COL_COD].isin(HQ_MANUAL)
    d["sede_bancaria"] = hq
    d["sem_agencia"] = d["Agências (ESTBAN)"].fillna(0) <= 0
    d["municipio_pequeno"] = d[COL_POP] < MIN_POP_OCEANO
    print(f"  Sedes bancárias neutralizadas no crédito: {int(hq.sum())} (corte p{int(HQ_QUANTIL*100)} = R$ {corte:,.0f} per capita)".replace(",", "."))
    print(f"  Municípios sem agência: {int(d['sem_agencia'].sum())} ({d['sem_agencia'].mean():.0%}) -> restrição máxima")

    res = d[[COL_COD, COL_NOME, COL_UF, COL_POP]].copy()
    res["sede_bancaria"], res["sem_agencia"], res["municipio_pequeno"] = d["sede_bancaria"], d["sem_agencia"], d["municipio_pequeno"]
    detalhe = {}
    for seg in SEGMENTOS:
        print(f"  Segmento: {seg}")
        pil = pd.DataFrame({p: pilar_score(d, VARIAVEIS[seg][p], hq) for p in PILARES})
        for p in PILARES:
            res[f"{p}__{seg}"] = pil[p].round(1)
        res[f"IPCM__{seg}"] = combina(pil, PESOS[seg]).round(1)
        # demanda = massa, demografia e frota (pesos do segmento); restrição = pilar de crédito
        wd = {k: PESOS[seg][k] for k in ("massa", "demografia", "frota")}
        dem = (pil[list(wd)].fillna(0) * pd.Series(wd)).sum(axis=1) / (pil[list(wd)].notna() * pd.Series(wd)).sum(axis=1)
        res[f"demanda__{seg}"] = dem.round(1)
        base_q = ~d["municipio_pequeno"] & ~hq
        td, tr = dem[base_q].quantile(Q_DEMANDA), pil["credito"][base_q].quantile(Q_RESTRICAO)
        alta_d, alta_r = dem >= td, pil["credito"] >= tr
        quad = np.select([alta_d & alta_r, alta_d & ~alta_r, ~alta_d & alta_r], ["Oceano azul", "Saturado", "Restrito sem demanda"], "Baixo potencial")
        quad = pd.Series(quad, index=d.index).mask(dem.isna() | pil["credito"].isna(), "Sem dados")
        quad = quad.mask(hq, "Sede bancária (crédito não comparável)")
        res[f"quadrante__{seg}"] = quad
        res[f"oceano_azul__{seg}"] = ((quad == "Oceano azul") & ~d["municipio_pequeno"]).values
        res[f"score_oceano__{seg}"] = np.where(res[f"oceano_azul__{seg}"], np.sqrt(dem * pil["credito"]).round(1), np.nan)
        res[f"rank_nacional__{seg}"] = res[f"IPCM__{seg}"].rank(ascending=False, method="min")
        res[f"rank_uf__{seg}"] = res.groupby(COL_UF)[f"IPCM__{seg}"].rank(ascending=False, method="min")
        detalhe[seg] = (td, tr)
        print(f"    limiares: demanda >= {td:.1f} | restrição >= {tr:.1f} | oceano azul: {int(res[f'oceano_azul__{seg}'].sum())} municípios")
    return res, detalhe


def sensibilidade(d_pil_cols, res):
    """Quanto o ranking geral muda ao variar o peso do crédito (demais pesos reescalados)."""
    base = PESOS["geral"]
    pil = pd.DataFrame({p: res[f"{p}__geral"] for p in PILARES})
    ref = combina(pil, base)
    top_ref = set(ref.nlargest(100).index)
    linhas = []
    for wc in [0, 10, 25, 40, 60]:
        resto = sum(v_ for k, v_ in base.items() if k != "credito")
        pes = {k: (v_ / resto) * (100 - wc) for k, v_ in base.items() if k != "credito"}
        pes["credito"] = wc
        s = combina(pil, pes)
        linhas.append({"peso do crédito": wc, "Spearman com o ranking padrão": round(s.corr(ref, method="spearman"), 3),
                       "municípios do top 100 padrão que continuam no top 100": len(top_ref & set(s.nlargest(100).index))})
    return pd.DataFrame(linhas)


def validacao(d, res):
    out = {}
    pil_g = res[[f"{p}__geral" for p in PILARES]].rename(columns=lambda c: ROTULOS[c.split("__")[0]])
    out["Correlação entre pilares (Spearman, geral)"] = pil_g.corr(method="spearman").round(2).reset_index().rename(columns={"index": "pilar"})
    ref = {"População": COL_POP, "PIB total": "PIB total (R$ mil)", "Frota total": "Total da frota",
           "Massa salarial (nível)": "Salários e outras remunerações (R$ mil/ano)"}
    out["IPCM geral x tamanho do município (Spearman)"] = pd.DataFrame(
        [{"variável": k, "Spearman com IPCM geral": round(res["IPCM__geral"].corr(d[c], method="spearman"), 3)} for k, c in ref.items()])
    out["Sensibilidade ao peso do crédito (geral)"] = sensibilidade(None, res)
    dist = res[[f"IPCM__{s}" for s in SEGMENTOS]].describe(percentiles=[.1, .5, .9]).T.round(1).reset_index().rename(columns={"index": "indicador"})
    out["Distribuição do IPCM"] = dist
    cap = ["SP", "RJ", "DF", "MG", "BA", "CE", "PE", "RS", "PR", "AM", "PA", "GO"]
    top = res.nsmallest(100, "rank_nacional__geral")
    out["Top 100 geral: composição"] = pd.DataFrame({
        "indicador": ["% com menos de 50 mil hab.", "% em municípios sem agência", "mediana de população", "UFs distintas"],
        "valor": [round((top[COL_POP] < 50_000).mean() * 100, 1), round(top["sem_agencia"].mean() * 100, 1),
                  int(top[COL_POP].median()), top[COL_UF].nunique()]})
    return out


# ============================================================================ saída
HEAD = (Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"),
        Alignment(wrap_text=True, vertical="center", horizontal="center"))


def salva(res, d, val, detalhe, caminho):
    cols_id = [COL_COD, COL_NOME, COL_UF, COL_POP, "sede_bancaria", "sem_agencia", "municipio_pequeno"]
    abas = {"IPCM": res}
    for seg in SEGMENTOS:
        c = cols_id[:4] + [f"IPCM__{seg}", f"rank_nacional__{seg}", f"rank_uf__{seg}"] + [f"{p}__{seg}" for p in PILARES] + \
            [f"demanda__{seg}", f"quadrante__{seg}", f"score_oceano__{seg}"]
        if seg == "geral":
            abas["Ranking_Geral"] = res.sort_values(f"rank_nacional__{seg}")[c].head(300)
        abas[f"OceanoAzul_{seg}"[:31]] = res[res[f"oceano_azul__{seg}"]].sort_values(f"score_oceano__{seg}", ascending=False)[c]
    param = [{"segmento": s, **{ROTULOS[p]: PESOS[s][p] for p in PILARES}, "limiar demanda": round(detalhe[s][0], 1),
              "limiar restrição": round(detalhe[s][1], 1)} for s in SEGMENTOS]
    abas["Parâmetros"] = pd.DataFrame(param)
    linhas = []
    for seg in SEGMENTOS:
        for p in PILARES:
            for it in VARIAVEIS[seg][p]:
                linhas.append({"segmento": seg, "pilar": ROTULOS[p], "variável": it["col"], "peso no pilar": it["w"],
                               "log": it["log"], "invertida (menos = mais potencial)": it["inv"],
                               "sem dado": {None: "ignorada", "zero": "=0", "max": "nota máxima"}[it["na"]],
                               "neutralizada em sede bancária": it["hq"]})
    abas["Variáveis"] = pd.DataFrame(linhas)
    for nome, t in val.items():
        abas[("Val_" + nome)[:31]] = t
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        for nome, t in abas.items():
            nome = re.sub(r"[:\\/?*\[\]]", "-", nome)[:31]          # caracteres inválidos em nome de aba
            t.to_excel(w, sheet_name=nome, index=False)
            ws = w.sheets[nome]
            for j, c in enumerate(t.columns, 1):
                h = ws.cell(row=1, column=j)
                h.font, h.fill, h.alignment = HEAD
                ws.column_dimensions[get_column_letter(j)].width = 28 if c in (COL_NOME, "variável") else 16
            ws.freeze_panes = "A2"
            ws.row_dimensions[1].height = 40
    res.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig")
    reg = []
    for _, r in res.iterrows():
        o = {"cod": int(r[COL_COD]), "nome": r[COL_NOME], "uf": r[COL_UF], "pop": int(r[COL_POP]) if pd.notna(r[COL_POP]) else None,
             "sede_bancaria": bool(r["sede_bancaria"]), "sem_agencia": bool(r["sem_agencia"])}
        for seg in SEGMENTOS:
            o[seg] = {"ipcm": None if pd.isna(r[f"IPCM__{seg}"]) else float(r[f"IPCM__{seg}"]),
                      "pilares": {p: (None if pd.isna(r[f"{p}__{seg}"]) else float(r[f"{p}__{seg}"])) for p in PILARES},
                      "demanda": None if pd.isna(r[f"demanda__{seg}"]) else float(r[f"demanda__{seg}"]),
                      "quadrante": r[f"quadrante__{seg}"], "oceano_azul": bool(r[f"oceano_azul__{seg}"])}
        reg.append(o)
    json.dump({"pesos_padrao": PESOS, "rotulos": ROTULOS, "limiares": {s: {"demanda": float(detalhe[s][0]), "restricao": float(detalhe[s][1])} for s in SEGMENTOS},
               "municipios": reg}, open(SAIDA_JSON, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print(f"\nSalvo: {caminho}, {SAIDA_CSV}, {SAIDA_JSON}")


def main(arq):
    print(f"Lendo {arq} ...")
    d = pd.read_excel(arq, sheet_name="Dados")
    print(f"  {len(d)} municípios, {d.shape[1]} colunas")
    res, detalhe = calcula(d)
    val = validacao(d, res)
    for nome, t in val.items():
        print(f"\n{nome}:\n{t.to_string(index=False)}")
    g = res.nsmallest(15, "rank_nacional__geral")[[COL_NOME, COL_UF, COL_POP, "IPCM__geral", "demanda__geral", "credito__geral", "quadrante__geral"]]
    print("\nTop 15 IPCM geral:\n" + g.to_string(index=False))
    oa = res[res["oceano_azul__geral"]].nlargest(15, "score_oceano__geral")[[COL_NOME, COL_UF, COL_POP, "IPCM__geral", "demanda__geral", "credito__geral", "score_oceano__geral"]]
    print("\nTop 15 Oceano Azul (geral):\n" + oa.to_string(index=False))
    salva(res, d, val, detalhe, SAIDA_XLSX)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a.lower().endswith(".xlsx")]
    main(args[0] if args else ARQ_ENTRADA)

# %%
