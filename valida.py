#%%
import coletar_estban as ce
per, z = ce.acha_periodo("202606")
raw = ce.le_estban(z)
print(raw.shape, [c for c in raw.columns if c.startswith("VERBETE_16")])
col = next(c for c in raw.columns if c.startswith("VERBETE_160"))
raw[col] = ce.to_num(raw[col])
inst = next(c for c in raw.columns if c.startswith("NOME_INST"))
print((raw.groupby(inst)[col].sum().sort_values(ascending=False).head(10) / 1e9).round(0))   # R$ bilhões
print("linhas duplicadas por município+CNPJ:", raw.duplicated(["CODMUN", "CNPJ"]).sum())
# %%
