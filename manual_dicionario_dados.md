# Dicionário de Dados — Dashboard de Acompanhamento de Obra

> **Como ler este documento:** cada valor exibido no app é rastreado até sua origem exata no arquivo Excel.
> O app lê dois arquivos: o **EAP** (arquivo de medição, ex: `EAP_MEDICAO_4_...xlsx`) e **`data/registros_projetos.xlsx`** (gerado internamente pelo app).
> As duas abas lidas do EAP são **`CONFIG`** e **`EAP DE MEDIÇÃO`**.

---

## Cabeçalho Global (aparece em todas as abas)

| Campo exibido | Origem |
|---|---|
| Título "Construção — Nova Sede da Justiça Federal" | Fixo no código (`tab_contrato`) |
| Subtítulo "Juazeiro do Norte / CE · Consórcio Juazeiro do Norte · CNPJ 62.009.288/0001-12" | Fixo no código |
| Badge **"Nª Medição"** (canto superior direito) | Aba `CONFIG` — célula que define `medicao_atual` (lida por `carregar_dados`) |
| Rodapé "Dados extraídos das abas CONFIG e EAP DE MEDIÇÃO" | Fixo no código |

---

## Aba: Visão Geral

### KPIs (3 cartões no topo)

#### KPI 1 — Avanço Físico Acumulado
| Campo | Origem |
|---|---|
| **Valor principal** (ex: `23,45%`) | Aba `EAP DE MEDIÇÃO` — linha de totais (linha 270 na planilha), coluna de `%_acumulado` da medição atual. Internamente: `df_totais["total_pct_acum"]` para a medição atual |
| Sub-label "Mês atual: X%" | Aba `EAP DE MEDIÇÃO` — mesma linha de totais, coluna de `%_mes` da medição atual. Internamente: `df_totais["total_pct_mes"]` |
| Sub-label "Medições anteriores: X%" | Aba `EAP DE MEDIÇÃO` — mesma linha de totais, coluna `%_acumulado` da medição **anterior** (N-1). Internamente: `df_totais["total_pct_acum"]` para `medicao == atual-1` |
| Badge "Obra no prazo / Atenção / Atraso (IDP = X)" | IDP vem da aba `CONFIG` — célula **M27** |

#### KPI 2 — Valor Medido Acumulado
| Campo | Origem |
|---|---|
| **Valor principal** (ex: `R$ 1.234.567,89`) | Aba `EAP DE MEDIÇÃO` — linha de totais, coluna de `valor_medido_acumulado` da medição atual. Internamente: `df_totais["total_medido_acum"]` |
| Sub-label "X% do contrato executado" | Calculado: `valor_medido_acumulado ÷ valor_contrato` |
| Badge "Nª Medição" | Aba `CONFIG` — `medicao_atual` |

#### KPI 3 — Saldo Contratual (3 colunas internas: Total, Obras, Projetos)
| Campo | Origem |
|---|---|
| **Saldo Obras** | Calculado: `valor_contrato − valor_medido_acumulado`. Onde `valor_contrato` = soma dos itens de nível 1 da aba `EAP DE MEDIÇÃO` (coluna `preco_total_rs`) |
| **% restante Obras** | `saldo_obras ÷ valor_contrato` |
| **Saldo Projetos** | Arquivo interno `data/registros_projetos.xlsx` — aba `Resumo`, campo `montante_total` (última entrada) menos soma dos `valor` com `status = Pago` da aba `Pagamentos` |
| **% restante Projetos** | `saldo_projetos ÷ montante_projetos` |
| **Total (Obras + Proj.)** | Soma dos dois saldos acima |

### Gráfico: Avanço por Etapa (card esquerdo)
| Campo | Origem |
|---|---|
| Nome de cada etapa | Aba `CONFIG` — tabela `eap_orcamento`, campo `descricao` (itens de nível 1, sem ponto no número do item) |
| Barra de progresso (% realizado) | Aba `EAP DE MEDIÇÃO` — coluna `pct_acumulado` para cada item de nível 1, medição atual |
| Risco (Alto/Médio/Normal/Baixo) | Calculado no código com base em `pct_acumulado` e `saldo_a_medir ÷ valor_contrato` |

### Gráfico: Situação da Obra — IDP e Análise (card direito)
| Campo | Origem |
|---|---|
| Velocímetro IDP | Aba `CONFIG` — célula **M27** |
| Texto de análise narrativa | Gerado automaticamente pelo código com base no valor do IDP |
| "Obra no prazo / Atenção / Atraso" | Calculado: IDP ≥ 1,0 → no prazo; IDP ≥ 0,9 → atenção; IDP < 0,9 → atraso |

### Gráfico: Curva S (parte inferior)
| Campo | Origem |
|---|---|
| Linha **Planejado** (acumulado) | Aba `CONFIG` — tabela `cronograma_mensal`, coluna `valor_planejado_acum`, dividido pelo `valor_contrato` |
| Linha **Realizado** (acumulado) | Aba `EAP DE MEDIÇÃO` — `df_totais["total_medido_acum"]` para cada medição, dividido pelo `valor_contrato` |

---

## Aba: Avanço Físico

### Tabela: Avanço por Etapa (coluna esquerda)
| Coluna | Origem |
|---|---|
| **Nº** | Número do item (nível 1) da aba `EAP DE MEDIÇÃO` |
| **Etapa** | Aba `CONFIG` — `eap_orcamento.descricao` |
| **Valor Total** | Aba `EAP DE MEDIÇÃO` — coluna `preco_total_rs` do item nível 1 |
| **Medido Acum.** | Calculado: `preco_total_rs × pct_acumulado` (medição atual) — aba `EAP DE MEDIÇÃO` |
| **Saldo a Medir** | Calculado: `Valor Total − Medido Acumulado` |
| **Risco** | Calculado no código (ver legenda no próprio card) |

### Gráfico: Comparativo Realizado vs Planejado (coluna direita, superior)
| Campo | Origem |
|---|---|
| Barras **Planejado** (por medição) | Aba `CONFIG` — `cronograma_mensal.valor_planejado` ÷ `valor_contrato` |
| Barras **Realizado** (por medição) | Aba `EAP DE MEDIÇÃO` — `df_totais["total_pct_mes"]` para cada medição |
| Label do eixo X (período) | Aba `EAP DE MEDIÇÃO` — coluna `data_ref` do `df_totais` |

### Gráfico: % Realizado por Etapa (coluna direita, inferior)
| Campo | Origem |
|---|---|
| Barras horizontais | Aba `EAP DE MEDIÇÃO` — `pct_acumulado` por item de nível 1, medição atual |

### Tabela: Serviços com Evolução na Nª Medição (inferior)
| Coluna | Origem |
|---|---|
| **Item** | Número do item — aba `EAP DE MEDIÇÃO`, coluna `item` |
| **Serviço / Descrição** | Aba `EAP DE MEDIÇÃO`, coluna `descricao` |
| **Status** (Concluído / Em andamento) | Calculado: `pct_acumulado ≥ 100%` → Concluído |
| **Avanço Mês** (ex: `+2,50%`) | Aba `EAP DE MEDIÇÃO` — coluna `pct_mes`, medição atual (itens com `pct_mes > 0`) |
| **% Acumulado** | Aba `EAP DE MEDIÇÃO` — coluna `pct_acumulado`, medição atual |

---

## Aba: Execução Financeira

### Cards de Resumo (3 no topo)
| Campo | Origem |
|---|---|
| **Valor do Contrato** | Aba `EAP DE MEDIÇÃO` — soma de `preco_total_rs` dos itens de nível 1 (sem ponto no item) |
| **Executado Acumulado** | Aba `EAP DE MEDIÇÃO` — `df_totais["total_medido_acum"]` da medição atual |
| **Saldo Contratual** | Calculado: `Valor do Contrato − Executado Acumulado` |

### Tabela: Histórico de Medições (coluna esquerda)
| Coluna | Origem |
|---|---|
| **Med.** | Número sequencial da medição |
| **Período** | Aba `EAP DE MEDIÇÃO` — `df_totais["data_ref"]` (data de referência de cada medição) |
| **% Mês** | Aba `EAP DE MEDIÇÃO` — `df_totais["total_pct_mes"]` |
| **% Acum.** | Aba `EAP DE MEDIÇÃO` — `df_totais["total_pct_acum"]` |
| **Valor Medido** | Aba `EAP DE MEDIÇÃO` — `df_totais["total_medido"]` |
| **Acumulado** | Aba `EAP DE MEDIÇÃO` — `df_totais["total_medido_acum"]` |
| Rodapé "Saldo contratual" | Calculado: `valor_contrato − total_medido_acum` (medição atual) |

### Gráfico: Evolução Financeira (coluna direita)
| Campo | Origem |
|---|---|
| Linha acumulada realizada | Aba `EAP DE MEDIÇÃO` — `df_totais["total_medido_acum"]` por medição |

### Tabela: Composição por Etapa (inferior)
| Coluna | Origem |
|---|---|
| **Etapa** | Aba `CONFIG` — `eap_orcamento.descricao` |
| **Valor contratual** | Aba `EAP DE MEDIÇÃO` — `preco_total_rs` por item nível 1 |
| **Executado acumulado** | `preco_total_rs × pct_acumulado` (medição atual) |
| **Saldo** | `Valor contratual − Executado acumulado` |

---

## Aba: Prazos

### Card: Progresso Temporal (barras de progresso)
| Campo | Origem |
|---|---|
| **Vigência do Contrato** — dias decorridos/total | Aba `CONFIG` — `vigencia.dias_decorridos` e `vigencia.prazo_total_dias` |
| **Data de fim da Vigência** | Aba `CONFIG` — `vigencia.fim` |
| **Prazo de Execução da Obra** — dias decorridos/total | Aba `CONFIG` — `execucao.dias_decorridos` e `execucao.prazo_total_dias` |
| **Data de fim da Execução** | Aba `CONFIG` — `execucao.fim` |
| **Análise de Tendência · IDP** | IDP: aba `CONFIG` — célula **M27**. Texto narrativo gerado automaticamente |
| **Tendência de término** | Aba `CONFIG` — `tendencia.nova_data_termino` |

### Card: Timeline / Linha do Tempo
| Linha | Origem |
|---|---|
| Prazo de Vigência | Aba `CONFIG` — `vigencia.fim` |
| Vigência Restante | Calculado: `vigencia.prazo_total_dias − vigencia.dias_decorridos` |
| Prazo de Execução | Aba `CONFIG` — `execucao.fim` |
| Execução Restante | Calculado: `execucao.prazo_total_dias − execucao.dias_decorridos` |
| IDP | Aba `CONFIG` — célula **M27** |
| Status | Calculado com base no IDP |

### Card: Análise de Ritmo
| Campo | Origem |
|---|---|
| Ritmo médio realizado (% por medição) | Calculado: `avanco_fisico_acumulado ÷ numero_de_medicoes` |
| Ritmo necessário para concluir no prazo | Calculado: `(100% − avanco_acumulado) ÷ medicoes_restantes` |
| Medições restantes | Aba `CONFIG` — calculado com base em `execucao.prazo_total_dias` e `data_ref` da última medição |

---

## Aba: Dados do Contrato

### Card: Identificação do Contrato
> Todos os campos abaixo são **fixos no código** (`tab_contrato` → dicionário `CONTRATO`), não vêm do Excel.

| Campo | Valor fixo no código |
|---|---|
| Objeto | "Construção da Nova Sede da Justiça Federal em Juazeiro do Norte / CE" |
| Nº do Contrato | "25/2025" |
| Concorrência Pública | "90001/2025" |
| CNO da Obra | "90.025.93927/72" |
| Empresa Contratada | "Consórcio Juazeiro do Norte" |
| CNPJ | "62.009.288/0001-12" |
| Situação do Contrato | "Vigente" (fixo) |
| Paralisações | "Nenhuma registrada" (fixo) |

### Card: Valores e Prazos
| Campo | Origem |
|---|---|
| **Valor Inicial do Contrato** | Fixo no código: `R$ 28.566.749,33` |
| **Valor Atual do Contrato** | Fixo no código: `R$ 28.566.749,33` |
| Aditivos de Valor | Fixo: "Nenhum" |
| Assinatura do Contrato | Fixo: "04/08/2025" |
| Ordem de Serviço | Fixo: "10/12/2025" |
| **Prazo de Execução** | Parte dinâmica: duração em meses vem de aba `CONFIG` — `tendencia.nova_duracao_meses`; data de término vem de `tendencia.nova_data_termino` |
| Prazo de Vigência | Fixo: "40 meses · Término em 04/12/2028" |
| Aditivos de Prazo | Fixo: "Nenhum" |

### Card: Endereço da Obra
| Campo | Origem |
|---|---|
| Endereço completo | Fixo no código |

---

## Aba: Upload EAP — Subaba Obras

> Esta subaba **não exibe dados** calculados — é apenas o mecanismo de importação.

| Campo | Descrição |
|---|---|
| Arquivo carregado | Arquivo `.xlsx` ou `.xlsm` enviado pelo usuário |
| Medição detectada | Lida da aba `CONFIG` do arquivo enviado (`medicao_atual`) |
| Medição atual no sistema | Lida da aba `CONFIG` do arquivo salvo em `data/eap_atual.xlsx` |
| Destino de gravação | `data/eap_atual.xlsx` (arquivo padrão) + `data/eap_medidaN.xlsx` (cópia por número de medição) |

---

## Aba: Upload EAP — Subaba Projetos

> Esta subaba usa um arquivo Excel **interno gerado pelo próprio app**: `data/registros_projetos.xlsx`, com duas abas: **Resumo** e **Pagamentos**. Não tem relação direta com as abas CONFIG ou EAP.

### Editor: Resumo por Medição
| Coluna | O que é | Origem / Edição |
|---|---|---|
| **Med.** | Número da medição | Editável manualmente |
| **Data** | Data do registro | Editável — seletor de data |
| **Avanço %** | Média dos `% Avanço` de todos os pagamentos daquela medição | Calculado automaticamente ao carregar; editável |
| **Montante (R$)** | Valor total contratado dos projetos | Editável manualmente |
| **Saldo (R$)** | `Montante − soma dos Valores de Pagamento` da medição | Calculado automaticamente ao carregar; editável |
| **Status** | Status geral da medição | Editável — selectbox (Pago / Pendente / Nao Pago) |
| **Observações** | Texto livre | Editável manualmente |

### Editor: Pagamentos por Etapa
| Coluna | O que é | Origem / Edição |
|---|---|---|
| **Med.** | Número da medição à qual o pagamento pertence | Editável |
| **Etapa / Projeto** | Nome da etapa ou projeto | Editável (texto livre) |
| **% Avanço** | Percentual de avanço desta etapa nesta medição | Editável |
| **Valor (R$)** | Valor pago ou a pagar nesta etapa | Editável |
| **Status** | Status do pagamento | Editável — selectbox (Pago / Pendente / Nao Pago) |

### Como o Saldo Contratual de Projetos é calculado (KPI na Visão Geral)
```
montante_total  →  último valor salvo no Resumo com montante > 0
total_pago      →  soma dos Valores com Status = "Pago" na aba Pagamentos
saldo_projetos  =  montante_total - total_pago
pct_restante    =  saldo_projetos / montante_total
```

---

## Mapeamento: Colunas do `df_totais` (aba `EAP DE MEDIÇÃO`)

O app lê a aba `EAP DE MEDIÇÃO` e constrói um DataFrame de totais (`df_totais`) com uma linha por medição:

| Coluna interna | Conteúdo | Linha/Coluna aproximada na planilha |
|---|---|---|
| `medicao` | Número da medição (1, 2, 3…) | Linha de totais, agrupada por medição |
| `total_pct_mes` | % medido no mês corrente | Linha 270 — coluna `%_mes` da medição N |
| `total_pct_acum` | % medido acumulado | Linha 270 — coluna `%_acumulado` da medição N |
| `total_medido` | Valor R$ medido no mês | Linha 271 — coluna `valor_medido` da medição N |
| `total_medido_acum` | Valor R$ medido acumulado | Linha 271 — coluna `valor_acumulado` da medição N |
| `data_ref` | Data de referência da medição | Linha de cabeçalho da medição N na EAP |

## Mapeamento: Chaves da aba `CONFIG`

| Dado usado no app | Localização na aba CONFIG |
|---|---|
| `medicao_atual` | Célula que indica o número da medição corrente |
| `idp.valor` | Célula **M27** |
| `vigencia.inicio` | Data de início da vigência |
| `vigencia.fim` | Data de fim da vigência |
| `vigencia.dias_decorridos` | Dias decorridos desde início da vigência |
| `vigencia.prazo_total_dias` | Prazo total em dias da vigência |
| `execucao.inicio` | Data de início da execução (Ordem de Serviço) |
| `execucao.fim` | Data de fim da execução |
| `execucao.dias_decorridos` | Dias decorridos desde a OS |
| `execucao.prazo_total_dias` | Prazo total em dias da execução |
| `tendencia.nova_data_termino` | Data projetada de término (com base no IDP) |
| `tendencia.nova_duracao_dias` | Duração projetada em dias |
| `tendencia.nova_duracao_meses` | Duração projetada em meses |
| `eap_orcamento` | Tabela com itens de nível 1: `item`, `descricao`, `valor_rs`, `pct_concluido` |
| `cronograma_mensal` | Tabela mês a mês: `mes`, `valor_planejado`, `valor_planejado_acum` |

---

*Gerado automaticamente com base na leitura do código-fonte `app.py`.*
