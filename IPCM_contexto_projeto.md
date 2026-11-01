# IPCM Intelligence — Contexto do Projeto (para uso como conhecimento de projeto)

> **Como usar este documento:** ele resume o escopo, as decisões, os dados, os scripts e o estado atual do projeto. Leia-o inteiro antes de responder qualquer pergunta sobre o projeto. Quando houver conflito entre este documento e uma informação nova dada pelo usuário na conversa, vale a informação nova (e avise que este documento está desatualizado).
>
> **Última atualização:** 24/10/2026. Decisão vigente: **usar o dado mais recente disponível de cada fonte** (substitui a decisão anterior de fixar 2021).

---

## 1. Instruções de trabalho para o assistente

- **Idioma:** português do Brasil.
- **Ambiente do usuário:** roda os scripts em Python, geralmente em **Jupyter** (por isso scripts não devem depender de `sys.argv` sem filtrar argumentos do kernel).
- **Ambiente do assistente:** o sandbox do assistente **não tem acesso à internet**. Os scripts de coleta **nunca foram executados contra as fontes reais pelo assistente**; foram testados com dados simulados e depois corrigidos a partir das saídas que o usuário colou. Nunca afirmar que algo "funciona com os dados reais" sem uma saída do usuário que comprove.
- **Padrão de trabalho que funcionou:** o assistente escreve o script → o usuário roda → cola a saída → o assistente corrige. Sempre pedir as linhas de diagnóstico que o script imprime.
- **Honestidade sobre incerteza:** vários detalhes (IDs de variáveis, nomes de colunas, URLs) vieram de memória ou de documentação de terceiros. Marcar o que não foi confirmado.
- **Preferências observadas:** respostas diretas, tabelas para comparações, código completo e pronto para rodar, indicação clara do que mudou e do que o usuário precisa conferir.

---

## 2. O projeto

**IPCM Intelligence** (Plataforma Integrada de Inteligência de Mercado e Predição de Consórcios). Objetivo: apoiar decisões de alocação de vendas, marketing e expansão de rede no setor de consórcios, usando **exclusivamente dados públicos**.

### Duas camadas

| Camada | Escopo | Estado |
|---|---|---|
| **Fase 1 — Macro (UF)** | Análise descritiva regional, previsão de cotas ativas (3–12 meses), cenários Mínimo/Base/Máximo por MAE, ranking de expansão por UF | **Já desenvolvida anteriormente; será reaproveitada.** Os dados e o código dessa fase serão enviados pelo usuário. |
| **Fase 2 — Micro (município), IPCM** | Score 0–100 por município, radar dos 5 pilares, ranking "Oceano Azul", simulador de peso do crédito, filtro por segmento, exportação CSV | **Foco atual:** provar a viabilidade do potencial de consumo municipal. |

### Os 5 pilares do IPCM

1. **Massa salarial formal** (CAGED/MTE) — poder de compra do emprego formal.
2. **Dinâmica demográfica** (IBGE) — população e adensamento urbano.
3. **Frota alvo renovável** (SENATRAN) — leves, motos, pesados/agro com potencial de substituição.
4. **Restrição de crédito tradicional** (ESTBAN/BACEN) — volume de crédito e presença de agências; **quanto maior a retração do crédito, maior a pontuação**.
5. **PIB per capita** (IBGE).

### Segmentos de produto (filtro)
Consórcio Geral, Imóveis, Auto/Moto, Pesados/Agro. Cada segmento deve ter pesos próprios dos pilares.

### Stack prevista
Pipeline em Python (e SQL) unificando as bases pelo **Código IBGE (6 e 7 dígitos)**; dashboard em HTML5/CSS3 (dark mode estilo Databricks), JavaScript puro e Chart.js (radar e rankings).

---

## 3. Conclusão sobre viabilidade (fase 2)

**Veredito preliminar: viável como índice relativo de potencial.** Dá para ranquear municípios por segmento com critério defensável. **Não** é possível afirmar quantidade absoluta de cotas ou vendas por município com dados públicos, porque não existe dado público de cotas por município.

### Cobertura dos pilares com os dados atuais

| Pilar | Situação | Lacuna |
|---|---|---|
| PIB per capita | Pronto | — |
| Demografia | Quase pronto | Falta **área/densidade** (Censo 2022, SIDRA tabela 4714 — a confirmar) |
| Restrição de crédito | Pronto | Medir **em relação ao tamanho da economia** (ver §9) |
| Frota | Pronta por tipo | Falta **idade da frota** (para "renovável"); RENAVAM cobre só parte do maquinário agrícola |
| Massa salarial | Parcial | CAGED traz salário de **admissão**, não a remuneração média de todos os empregados; a RAIS (vínculos) completaria |

Segmentos: **Auto/Moto** e **Pesados/Agro** bem sustentados; **Imóveis** é o mais fraco (só domicílios, população, renda e crédito; faltam crescimento populacional 2010→2022, urbanização e crédito imobiliário).

### Plano de validação (a executar)

1. Calcular o IPCM por município e segmento.
2. Agregar por UF e comparar com as **cotas ativas por UF por habitante** da fase 1.
3. Se o ranking de UFs pelo índice acompanhar o ranking real, o índice captura algo do mercado; se não, revisar pilares/pesos.
4. Limite: são 27 pontos (ou 5, se a série for regional); valida o agregado, **não** município a município.

### Ponte entre as fases (rateio top-down)
Distribuir a projeção de cotas ativas de cada UF entre seus municípios proporcionalmente ao potencial do IPCM no segmento → meta estimada de cotas por município coerente com a previsão estadual. **Premissa a declarar explicitamente:** as cotas se distribuem dentro da UF na mesma proporção do potencial.

---

## 4. Dados públicos coletados (pipeline atual)

**Decisão (24/10/2026):** cada fonte usa o **dado mais recente disponível**, descoberto pelo próprio script (não há ano fixo no código). Cada planilha traz colunas com o período de cada informação, e o `cruzar_dados.py` gera uma aba **"Períodos"**. Consequência: as fontes têm períodos diferentes; razões como crédito/PIB servem para **comparar municípios entre si** (o IPCM usa percentis), não como valor absoluto.

Chave de integração: **Código IBGE de 7 dígitos** (`Código IBGE (7 dígitos)`); o de 6 dígitos é o de 7 sem o dígito verificador.

**Estado de execução:** os 5 coletores e o `cruzar_dados.py` foram **reescritos/ajustados para "mais recente" e testados só com dados simulados**; ainda **não foram executados na versão atual contra as fontes reais**. Versões anteriores (ano 2021) rodaram e foram confirmadas pelo usuário (IBGE, ESTBAN, frota, CAGED e `coletar_pendentes.py`).

### Períodos esperados (a confirmar na execução)

| Informação | Período esperado | Observação |
|---|---|---|
| PIB total e per capita | 2023 | Último publicado (dez/2025) |
| VAB por setor | **2021** | O IBGE publicou o PIB 2022-2023 **sem a abertura do VAB por setor** (reformulação do sistema de contas). Confirmado por saída do usuário (0/5570 valores em 2022 e 2023) e pelo IBGE |
| População estimada | Última da tabela 6579 (2025 ou 2026) | |
| População no ano do PIB | 2023, **interpolada** entre Censo 2022 e estimativa 2024 | Não há estimativa 2023 na tabela 6579 |
| Crédito/depósitos (ESTBAN) | Mês mais recente (provavelmente 2 meses de defasagem) | Script sonda os últimos 18 meses |
| Frota por tipo e por idade | **Setembro/2026** (publicado) | Páginas `frota-de-veiculos-AAAA` do Ministério dos Transportes |
| Novo CAGED | Último mês publicado; fluxos = 12 meses | |
| Estoque CLT | RAIS mais recente + saldo CAGED acumulado até o último mês | |
| CEMPRE (massa salarial) | Último ano da tabela | |
| Censo | 2022 (e 2010 para crescimento) | |

### 4.1 IBGE — `coletar_ibge.py` → `dados_ibge_municipios.xlsx`

- **PIB e VAB:** SIDRA tabela **5938**; variáveis 37 (PIB), 513 (VAB agro), 517 (indústria), 6575 (serviços exc. adm. pública), 525 (adm. pública) — **IDs confirmados**. O script procura nos metadados a variável de PIB per capita; se não achar, calcula (PIB × 1000 ÷ população). Colunas de ano separadas: **Ano do PIB** e **Ano do VAB**.
- **População:** tabela **6579**, variável 9324 (última estimativa) e população no ano do PIB (exata se existir; senão **interpolação linear** entre anos vizinhos, incluindo Censo 2022 da tabela 4714). Coluna "Método da população (ano do PIB)".
- **Domicílios:** tabela **4712**, variável 381 = "Domicílios particulares permanentes **ocupados**" (Censo 2022).
- Cuidado: o VAB é de ano anterior ao PIB; usar como perfil estrutural (participações).

### 4.2 ESTBAN/BCB — `coletar_estban.py` → `dados_estban_municipios.xlsx`

- `ANO_REF = None` → **mês mais recente disponível** (ou passar AAAAMM).
- URL: `https://www.bcb.gov.br/content/estatisticas/estatistica_bancaria_estban/municipio/{AAAAMM}_ESTBAN.{csv.zip|ZIP|csv}` (o endereço antigo `www4.bcb.gov.br/fis/cosif/...` dá 404). Em dez/2021 o `.ZIP` funcionou.
- Formato: `;`, Latin-1, duas linhas de título; **uma linha por instituição e município** (soma por município); valores em **R$**.
- **`CODMUN` é código do BCB** → pareamento por UF + nome, com aproximado (`difflib`, corte 0,86) e abas de conferência.
- Verbetes: **160** crédito; **420** poupança (coluna separada); **432** a prazo; **à vista = qualquer coluna que contenha verbetes de 401 a 419** (em dez/2021 vinha numa única coluna agregada 401–419; o script cobre também colunas separadas). A aba "ESTBAN completo" inclui o código IBGE e a aba "Verbetes" lista os códigos (para identificar crédito imobiliário e rural nos verbetes 16x — **pendente**).
- Agências: `AGEN_PROCESSADAS`. Sem postos (PAB/PAE).
- Em aberto: confirmar se o ESTBAN inclui **cooperativas de crédito**.
- Resultado real anterior (dez/2021): 3.144 municípios com código IBGE (11 aproximados), 21 sem par.

### 4.3 SENATRAN (frota) — `coletar_frota.py` → `dados_frota_municipios.xlsx`

- O script lê a página `https://www.gov.br/transportes/pt-br/assuntos/transito/conteudo-Senatran/frota-de-veiculos-AAAA`, associa cada link ao título "(Mês de AAAA)" e escolhe o **mês mais recente** do arquivo "Frota por Município e Tipo". Os **nomes dos arquivos variam de mês a mês** (e há erros como "Maro" para março), por isso não se monta a URL a partir de padrão.
- Formatos: dez/2021 era `.xls` (requer `xlrd`); 2026 são `.xlsx`.
- Planilha traz só UF e nome → pareamento por nome + `DE_PARA_MANUAL`.
- Tipos detectados dinamicamente. **Grupos (decisão do assistente, não oficial):** *Automóveis e utilitários* = automóvel + caminhonete + camioneta + utilitário; *Motocicletas* = motocicleta + motoneta; *Caminhões, tratores e reboques* = caminhão + caminhão-trator + trator (esteira e rodas) + reboque + semirreboque; *Demais tipos* = restante.
- Resultado real anterior (dez/2021): TOTAL = soma dos tipos (0 divergências); 5.559 municípios pareados (31 aproximados), 13 sem par.
- Frota é contada no município de **registro**; inclui veículos antigos sem baixa.

### 4.4 Novo CAGED — `coletar_caged.py` → `dados_caged_municipios.xlsx`

- FTP `ftp.mtps.gov.br`, pasta `pdet/microdados/NOVO CAGED/AAAA/AAAAMM/` (`CAGEDMOV`, `CAGEDFOR`, `CAGEDEXC`; EXC com peso −1). Requer `py7zr`; cache em `cache_caged/`.
- **Lógica atual:** descobre o último mês publicado (competência final) e o último ano da RAIS (arquivo `RAIS_ESTAB_PUB.7z`, pasta `pdet/microdados/RAIS/AAAA`). Lê **todos os meses desde 01/(ano RAIS + 1) até a competência final** (pode ser ~20 meses × 3 arquivos: vários GB e bastante tempo).
- **Fluxos = janela móvel de 12 meses** (admissões, desligamentos, saldo, salários médios de admissão/desligamento). **Estoque CLT estimado = vínculos CLT da RAIS em 31/12 + saldo acumulado do CAGED até a competência final.** Meses recentes podem ser revisados pelo MTE.
- **Confirmado por execução anterior:** colunas da RAIS `Município` e `Qtd Vínculos CLT`; **código `999999`** (~275 mil movimentos em 2021) não existe na tabela do IBGE (provavelmente "município ignorado"; não confirmado em documentação) e é descartado.
- O CAGED só traz salário de admitidos/desligados, **não** o salário médio de todos os empregados (isso vem do CEMPRE).
- Cobre só CLT; local = estabelecimento. A API de localidades do IBGE lista 5.571 municípios (um a mais que o SIDRA); decisão do usuário: ignorar.

### 4.5 Pendentes — `coletar_pendentes.py` → `dados_pendentes_municipios.xlsx`

Três blocos independentes (cada um registra OK/ERRO na aba "Diagnóstico"); IDs do SIDRA são **descobertos pelos metadados** (por nome):
- **Censo:** tabela 4714 (pop. 2022, área km², densidade), 202 (pop. 2010 total e urbana), 9923 (pop. urbana 2022) → crescimento 2010-2022, CAGR, grau de urbanização.
- **CEMPRE** (tabela 6450 ou 1685): pessoal assalariado, salários e outras remunerações (R$ mil/**ano**), salário médio mensal, **último ano publicado**; separa a **administração pública** (massa salarial exc. adm. pública). A massa do CEMPRE inclui todos os assalariados formais, inclusive estatutários; em muitos municípios pequenos a prefeitura domina.
- **Frota por idade (SENATRAN):** arquivo "UF, Município, Ano de Fabricação e Modelo" do **mês mais recente**; idade média/mediana, faixas, % com 10+ e 15+ anos ("renovável"). **A idade vale para a frota total** (o arquivo não separa por tipo).
- Versão anterior (ano 2021) rodou e foi considerada OK pelo usuário.

### 4.6 Cruzamento — `cruzar_dados.py` → `base_integrada.xlsx` e `base_integrada.csv`

- **5 entradas** (as 5 planilhas acima); base = IBGE (left join por código de 7 dígitos, `validate="one_to_one"`); colunas `Tem dado ...` por fonte; agências ausentes = 0; demais ausentes ficam **vazios**.
- **Abas:** Dados, Verificação (totais fonte × base), Alertas (valores suspeitos), Cobertura (% preenchido), Estatísticas (distribuição), **Períodos**, Dicionário, Frota por tipo, CAGED mensal, Fontes e notas. O CSV (`,` e `.`) é para BI/dashboard.
- Per capita de dados recentes usa a **última população estimada**; domicílios usam o Censo 2022; salário de admissão em salários mínimos usa o mínimo do ano da competência final (**2026 = R$ 1.621**, Decreto 12.797/2025; a janela de 12 meses atravessa dois valores, então é aproximação).
- Indicadores (≈30): crédito/PIB e depósitos/PIB; crédito e depósitos per capita; habitantes por agência; agências por 100 mil hab.; veículos, automóveis, motos e carga por mil hab.; motos por automóvel; veículos por domicílio; frota renovável por mil hab.; estoque CLT por 100 hab.; crescimento e rotatividade do emprego; salário de admissão em SM; empregos por R$ milhão de PIB; massa salarial per capita (total e exc. adm. pública); massa salarial/PIB; assalariados por 100 hab.; participação setorial do VAB; habitantes por domicílio; duas conferências de consistência.
- **Estado:** testado só com dados simulados; **nunca executado com os dados reais** (o usuário vai rodar e enviar a saída).

## 5. Esquema da base integrada (colunas originais)

- **IBGE:** Código IBGE (7 e 6 dígitos), Município, UF, População estimada, Ano da estimativa, População (ano do PIB), Método da população (ano do PIB), **Ano do PIB**, PIB total (R$ mil), PIB per capita (R$), **Ano do VAB**, VAB Agropecuária, VAB Indústria, VAB Serviços exc. adm. pública, VAB Adm. pública, VAB Serviços total (R$ mil), Domicílios particulares permanentes, Ano do Censo.
- **ESTBAN:** ESTBAN período (AAAAMM), Saldo de operações de crédito (R$), Depósitos à vista, a prazo e de poupança (R$), Total depósitos à vista + a prazo (R$), Agências (ESTBAN).
- **Frota:** Frota período (AAAAMM), Automóveis e utilitários, Motocicletas, Caminhões/tratores/reboques, Demais tipos, Total da frota.
- **CAGED:** CAGED competência final e início da janela (AAAAMM), Estoque CLT base RAIS (31/12 do ano-base), Admissões, Desligamentos e Saldo (12m), Saldo acumulado desde a base, Estoque CLT estimado (competência final), Salário médio de admissão e de desligamento 12m (R$).
- **Pendentes (Censo/CEMPRE/frota idade):** População Censo 2022 e 2010, população urbana, área (km²), densidade, crescimento 2010-2022 (total e CAGR), grau de urbanização 2010 e 2022; Ano do CEMPRE, empresas/unidades locais/pessoal ocupado e assalariado, salários e outras remunerações (R$ mil/ano), salário médio mensal (R$ e SM), massa salarial mensal média, administração pública (assalariados, salários, % da massa); Mês da frota por idade, idade média e mediana, faixas de idade, % com 10+ e 15+ anos, frota com 10+ anos (unidades).

Unidades: PIB e VAB em **R$ mil**; ESTBAN em **R$**; salários e outras remunerações do CEMPRE em **R$ mil por ano**; frota em unidades; salários do CAGED em R$/mês.

---

## 6. Dado de consórcios (necessário para a fase 1 e para a validação)

Pesquisa do assistente sobre dados **públicos** de consórcios (resumo; nada foi baixado):
- O BCB passou a divulgar dados consolidados do Sistema de Consórcios no **Portal de Dados Abertos**, substituindo o relatório anual (Panorama); está prevista periodicidade **trimestral**.
- Séries antigas do **SGS** (cotas ativas por segmento e por tipo de administradora; cotistas por região e por faixa de renda) são **anuais**, a partir de **31/12/2010**.
- **Não confirmado:** existência de recorte por **UF** e de frequência **mensal** em fonte pública.
- Implicação: a previsão de 3–12 meses com cone de MAE exige série mensal; com série anual (~16 pontos) o modelo é estatisticamente frágil. Alternativas: dados mensais da fase 1, outra fonte (ex.: ABAC — não verificada) ou horizonte anual/trimestral.

> O usuário informou que a fase 1 já usou dados próprios (que serão enviados) e que o projeto como um todo deve usar **somente dados públicos**.

---

## 7. Decisões já tomadas

| Decisão | Motivo |
|---|---|
| **Dado mais recente de cada fonte** (24/10/2026) | Mapa comercial com dados de 2021 envelheceria; substitui a decisão anterior de fixar 2021 por consistência |
| VAB setorial fica em 2021 | IBGE não publicou a abertura setorial em 2022-2023 |
| Base de municípios = IBGE (SIDRA) | Chave única; município extra do CAGED/API é ignorado |
| Valores calculados como números (não fórmulas) | Leitura correta no pandas |
| Pareamento ESTBAN/frota por UF + nome, com aproximado e de-para manual | ESTBAN usa código BCB; frota não traz código |
| Poupança fora do total de depósitos | Verbete 420 separado |
| Estoque CLT **estimado** (RAIS + saldo CAGED) | CAGED não informa estoque |
| Massa salarial vem do **CEMPRE** (não do CAGED) | CAGED só tem salário de admitidos |
| Frota ausente/CAGED ausente = vazio; agências ausentes = 0 | Evitar distorção nos indicadores |
| Projeto acadêmico: sem estimar/projetar dados para 2026 | Complexidade fora do escopo |

## 8. Pendências

**Dados**
1. Rodar na ordem os 5 coletores na versão "mais recente" e enviar as saídas: `coletar_ibge.py` → `coletar_estban.py` → `coletar_frota.py` → `coletar_caged.py` (lento) → `coletar_pendentes.py`.
2. Rodar `cruzar_dados.py` e revisar "Verificação", "Alertas", "Cobertura" e "Períodos".
3. Conferir as abas de pareamento aproximado e "Sem código IBGE" (ESTBAN e frota); preencher `DE_PARA_MANUAL` se relevante.
4. Identificar na aba "Verbetes" do ESTBAN quais são **crédito imobiliário** e **rural** (verbetes 16x) para os segmentos Imóveis e Pesados/Agro.
5. Indicadores para **Imóveis**: crédito imobiliário (item 4) e demais variáveis; é o segmento com menos cobertura.
6. Decidir como usar a **frota por idade** por segmento (vale só para a frota total).

**Modelagem**
7. Escrever `calcular_ipcm.py`: normalização por **percentil**, pesos configuráveis por segmento, quadrante Oceano Azul, exportação CSV/JSON.
8. Receber dados/código da fase 1; verificar granularidade da série de cotas; desenhar e executar a validação por UF (§3).
9. Protótipo do dashboard em HTML (radar, ranking, slider de peso do crédito, filtro por segmento, exportação CSV).

## 9. Recomendações de desenho do score (a aplicar)

1. **Pilar de crédito = crédito relativo, não volume.** Como "menos crédito = mais pontos", o volume absoluto daria nota alta de restrição a qualquer cidade pequena. Usar crédito/PIB, crédito per capita e agências por 100 mil habitantes (já calculados).
2. **"Oceano Azul" é combinação (demanda alta × restrição alta), não soma.** Tratá-lo como filtro/quadrante sobre o score. Uma média ponderada deixa um pilar forte compensar um fraco.
3. **Correlação entre pilares:** PIB per capita, massa salarial e crédito se correlacionam; checar redundância antes de fixar pesos (matriz de correlação).
4. **Normalização por percentil** e tratamento de outliers (capitais, polos de locadoras e sedes).
5. **Declarar premissas** ao apresentar: score é índice relativo de potencial; rateio top-down assume proporcionalidade; estoque CLT é estimativa.

## 10. Limitações a lembrar ao interpretar os dados

- **ESTBAN** contabiliza crédito/depósito no município da **agência**; cidades sem agência aparecem sem crédito mesmo com moradores atendidos na vizinhança.
- **Frota** no município de **registro**; inclui veículos antigos sem baixa.
- **CAGED** só celetistas; salário refere-se a admitidos/desligados; local = estabelecimento.
- **Hipótese de negócio não testada:** "menor crédito bancário ⇒ maior demanda por consórcio".
- Municípios muito pequenos: números pequenos geram razões instáveis.

---

## 11. Inventário de arquivos

| Arquivo | Função |
|---|---|
| `coletar_ibge.py` | PIB, VAB, população, domicílios (SIDRA), anos mais recentes |
| `coletar_estban.py` | Crédito, depósitos, agências (BCB) |
| `coletar_frota.py` | Frota por município e tipo (SENATRAN) |
| `coletar_caged.py` | Fluxos de 12 meses e estoque estimado de emprego CLT (MTE) |
| `coletar_pendentes.py` | Censo (área, densidade, urbanização), CEMPRE (massa salarial), frota por idade |
| `cruzar_dados.py` | Cruza as 5 planilhas, calcula indicadores, verifica totais |
| `dados_*_municipios.xlsx` | Saídas de cada coleta |
| `base_integrada.xlsx` | Base final com abas Dados, Verificação, Dicionário e detalhes |
| `cache_caged/` | Arquivos baixados do FTP (pode ocupar GB) |

**Dependências:** `requests pandas openpyxl xlrd py7zr`.

**Ordem de execução:** `coletar_ibge.py` → `coletar_estban.py` → `coletar_frota.py` → `coletar_caged.py` → `coletar_pendentes.py` → `cruzar_dados.py`.


---

## 12. Visão macro (UF / nacional) — coleta do zero com dados públicos — `coletar_macro.py`

**Situação:** os dados da fase 1 anterior não foram localizados; a coleta foi refeita do zero em um único script, com blocos independentes (cada um registra OK/ERRO na aba "Diagnóstico"). Saídas: `dados_macro.xlsx` e pasta `macro_csv/`. **Testado só com respostas simuladas; ainda não executado contra as fontes reais.**

| Dado pedido | Fonte e método | Estado / limites |
|---|---|---|
| Série nacional de cotas ativas (mensal) | BCB, serviço OData **PANORAMA_DE_CONSORCIOS**: `Metricas(DataBase=@DataBase)` com `@DataBase=AAAAMM`, `$format=json`, `$top=200`. 125 métricas por mês (campos DataBase, IdMetrica, Grupo, Metrica, Valor, Unidade). IDs: 10 = cotas ativas total; 11 imóveis, 12 veículos pesados, 13 automóveis, 14 motocicletas, 15 outros bens móveis duráveis, 16 serviços, 17–21 subcategorias; 9 = grupos ativos; 28/31/34 contempladas; 54–60 comercializadas (12 meses) | Valores em **mil** (convertidos para unidades). Último período observado em doc de terceiros: 202606; 202609 vazio. Início real da série não confirmado (o script consulta desde 2009 e ignora meses vazios) |
| Série regional por UF (trimestral) | Mesmas métricas **99 a 125** (somente cotas ativas por estado). A sigla da UF é extraída do nome da métrica | Não há cruzamento segmento × UF no OData. Os nomes das métricas por UF não foram vistos (a identificação da UF pode precisar de ajuste) |
| Market share por administradora | Arquivo `AAAAMMConsorcios.zip` (BCB, Base de Dados de Consórcios): contém `AAAAMMSegmentos_Consolidados.csv` (por administradora, segmentos 1–6) e `AAAAMMBens_Imoveis_Grupos.csv`; CSV `;` decimal `,` cp1252. "Dados por UF" (por administradora × UF × segmento) é trimestral | **O endereço de download do ZIP não foi confirmado.** O script tenta links da página e 4 padrões; se falhar, usar pasta `consorcios_bcb_zips/` com ZIPs baixados à mão. Nomes das colunas do CSV não vistos (detecção por heurística) |
| Selic e dólar | SGS: **432** (meta Selic, % a.a.), **4390** (Selic acumulada no mês, % a.m.), **1** (dólar venda, diário). Janelas de 5 anos (a API recusa janelas muito longas, HTTP 406) | Códigos confirmados em fontes de terceiros; mensal = média e último valor |
| Cód_UF, população e PIB por UF | API de localidades (estados); SIDRA 6579 (estimativas, N3), 4714 (Censo 2022), 5938 (PIB; tenta N3, senão soma municipal) | PIB com defasagem de 2–3 anos; painel traz colunas `_ffill` |
| Google Trends | `pytrends` (**arquivado em abr/2025**; pode falhar com 429) ou CSV exportado manualmente de trends.google.com para a pasta `trends_csv/`. Os 3 termos são pedidos **juntos** (mesma escala 0–100) | A API oficial do Google é alpha por solicitação. Índice relativo, não volume. Últimos meses podem ser parciais |
| Tabela FIPE | API v2 `https://fipe.parallelum.com.br/api/v2` (`/references`, `/{cars\|motorcycles\|trucks}/brands/.../models/.../years/...`, `/{tipo}/{codigoFipe}/years/{ano}/history`). Limite gratuito: 1.000 req/dia com token (variável `FIPE_TOKEN`) | Coleta uma **CESTA** configurável de modelos (não o catálogo inteiro, inviável no limite) e o histórico por código FIPE. Formato exato da resposta de `/history` não confirmado (parser tolerante) |
| ABAC | Boletins e matérias mensais em páginas/PDF (`abac.org.br/sistema-downloads/...`, `blog.abac.org.br/drops-de-mercado/...`). **Não há API nem planilha pública** | Extração **heurística** de participantes ativos, vendas de cotas, contempladas e negócios; sempre conferir pelo trecho/URL. Lista de PDFs e aba `ABAC_Manual` para preenchimento manual |

**Painéis gerados:** `Painel_Mensal` (cotas nacionais + Selic + dólar + Trends, por mês) e `Painel_UF_Trimestral` (cotas por UF + população + PIB + PIB per capita, com último valor disponível).

**Limites estruturais a lembrar ao modelar:**
- O OData agrega (sem administradora, sem cruzamento segmento × UF); series regionais são **trimestrais**.
- Série nacional "mensal" pode ter lacunas em meses sem publicação.
- Séries de períodos diferentes (PIB 2023 × cotas 2026); trabalhar com variações e defasagens.
- ABAC "participantes ativos" (pessoas) ≠ BCB "cotas ativas" (cotas); não misturar sem explicitar.


---

## 13. IPCM calculado na base real (teste de viabilidade) — `calcular_ipcm.py`

**Base usada:** `base_integrada` com 5.570 municípios e 113 colunas (ESTBAN 202606, frota 202609, CAGED 202509–202608, PIB 2023, VAB 2021, CEMPRE 2021, Censo 2022). O estoque CLT continua vazio na base (CAGED não foi reexecutado com a correção da RAIS); **o IPCM não depende dele**.

**Método:** por variável, log nas de tamanho → limites nos percentis 1% e 99% → posto percentual; variáveis inversas (crédito, agências) com empate no extremo = nota máxima (município sem agência = restrição máxima); pilar = média ponderada das variáveis disponíveis; IPCM = média ponderada dos 5 pilares (pesos por segmento, configuráveis); **Oceano Azul = demanda ≥ p60 e restrição ≥ p60** (filtro, com população ≥ 10 mil e sem sede bancária). Sedes bancárias (crédito per capita > p99 entre municípios com agência + Brasília e Osasco; 29 municípios) têm as variáveis monetárias de crédito neutralizadas.

**Pesos iniciais (arbitrários, a validar):** geral 25/15/20/25/15; imóveis 30/20/5/25/20; auto-moto 25/10/35/20/10; pesados-agro 15/5/35/25/20 (massa/demografia/frota/crédito/PIB p.c.).

**Resultados observados:**
- IPCM de 16,5 a 89,7 (média 53,5); correlação de postos com população 0,28 (não é só tamanho).
- Correlação entre pilares: massa×frota 0,76; massa×PIB p.c. 0,76; **restrição de crédito × demanda: −0,56 a −0,64** (cidades pequenas têm menos banco e menos demanda; o Oceano Azul é uma combinação rara).
- Oceano Azul: 189 municípios (geral), 151 (imóveis), 191 (auto/moto), 200 (pesados/agro); concentração em SP, RJ, MG, SC (e PR no agro).
- **Fragilidade:** 48% dos municípios (2.673) não têm agência no ESTBAN, então o pilar de crédito vira quase binário (100 para todos eles). 53% do top 100 geral são municípios sem agência. Variar o peso do crédito de 25 para 10 ou 40 mantém só ~50 dos 100 primeiros (Spearman 0,89 e 0,76); com peso 60, Spearman 0,36.

**Riscos de validade a tratar:**
1. **Cooperativas de crédito (Sicredi, Sicoob etc.) não estão no ESTBAN**; "sem agência bancária" ≠ "sem crédito". A restrição está superestimada em municípios atendidos só por cooperativas. Melhoria: coletar agências de cooperativas (cadastro do BCB; fonte não verificada).
2. O IPCM mede **propensão/atratividade**, não **volume**. Considerar exibir "tamanho de mercado" (massa salarial absoluta, frota) separado do score.
3. Sem dado de cotas por município, a validação possível é agregada por UF (§3), ainda **pendente** (depende da série de cotas por UF; ver §12).
4. A massa salarial "exc. administração pública" só existe em 49% dos municípios (sigilo estatístico do CEMPRE em municípios pequenos); o pilar usa a massa salarial total (100% de cobertura).

**Crédito por modalidade (ESTBAN 202606):** verbetes 161 (empréstimos e títulos descontados, R$ 3,40 tri), 162 (financiamentos, 1,21), 163 (rural, 1,23), 167 (agroindustrial, 0,02), 169 (imobiliário, 3,09); os 5 somam R$ 8,93 tri contra R$ 12,73 tri do verbete 160 (outros componentes não exportados). Brasília concentra 51% do imobiliário e 52% do rural (sede de Caixa e BB). Totais por instituição ficam ~1,6 a 2× acima das carteiras divulgadas; usar só para comparação relativa.
