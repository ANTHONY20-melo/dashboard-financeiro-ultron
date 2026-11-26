# 📊 Guia Power BI — Dashboard Financeiro Novembro/2026

Pacote pronto para montar o dashboard no Power BI Desktop a partir dos CSVs
gerados de `dados.json` (mesma fonte do dashboard HTML).

## Arquivos do pacote

| Arquivo | Conteúdo | Linhas |
|---|---|---|
| `receitas.csv` | Receitas com bruto / dízimo / líquido + total | 6 |
| `despesas.csv` | Prioridades + Pessoais + Quitáveis + Extra (roupas), com Status | 27 |
| `investimentos.csv` | Depósitos mensais e acumulado da reserva | 4 |
| `resumo.csv` | KPIs prontos para os cartões | 10 |
| `medidas-dax.txt` | Medidas DAX prontas para copiar | — |

Formato: UTF-8 com BOM, separador `;` e decimais com vírgula (padrão pt-BR).

---

## 1. Importar as tabelas

No **Power BI Desktop**:

1. **Página Inicial → Obter Dados → Texto/CSV**
2. Importe os 4 CSVs (um de cada vez).
3. Na pré-visualização, confira se **Origem do Arquivo = 65001: Unicode (UTF-8)** e o separador `;` foi detectado.
4. Se os números aparecerem como texto, clique no ícone de tipo da coluna `Valor` e escolha **Usar Localidade... → Decimal fixo → Português (Brasil)**.
5. Repita para `receitas.csv`, `despesas.csv`, `investimentos.csv` e `resumo.csv`.

> Não é preciso criar relacionamentos entre as tabelas — cada CSV é independente.

## 2. Criar as medidas DAX

Clique com o direito na tabela `resumo` → **Nova Medida** e cole as medidas do
arquivo `medidas-dax.txt`. Elas calculam dinamicamente a partir de
`despesas` e `receitas` (não dependem da coluna fixa do `resumo.csv`).

## 3. Montar a página (layout sugerido)

### Cartões de KPI (topo — 6 cartões)
| Cartão | Medida/Campo |
|---|---|
| Receita Líquida Prevista | `[Receita Líquida Total]` |
| Contas de Novembro | `[Contas de Novembro]` |
| Sobra Prevista | `[Sobra Prevista]` |
| Dívidas Quitáveis | `[Total Quitáveis]` |
| Rescisão Corrigida | `[Rescisão Corrigida]` |
| Reserva 3 Meses | `[Reserva Acumulada]` |

Dica: Formatar cartão → Rótulo de dados → **Dinâmico pt-BR, 2 casas**, cor de
fundo `#111A2E`, cantos arredondados 10px.

### Gráficos
1. **Rosca (Donut) — "Composição das receitas"**
   - `despesas`→ não; use `receitas[Descrição]` + `receitas[Valor Bruto]`
   - Legenda: Descrição · Valores: Valor Bruto (exclua a linha "TOTAL PREVISTO" com um filtro de visual: Descrição ≠ "TOTAL PREVISTO")
2. **Rosca — "Para onde vai o líquido"**
   - Gráfico de pizza com `despesas[Categoria]` + `despesas[Valor]` filtrado em **Prioridade, Pessoal, Extra** + medida `[Sobra Prevista]` como "Sobra"
   - Alternativa simples: usar `[Contas de Novembro]` vs `[Sobra Prevista]` em gráfico de barras empilhadas 100%
3. **Barras horizontais — "Despesas pessoais"**
   - `despesas[Item]` + `despesas[Valor]`, filtro Categoria = "Pessoal", ordenar decrescente por valor
4. **Barras — "Dívidas quitáveis"**
   - `despesas[Item]` + `despesas[Valor]`, filtro Categoria = "Quitável"
5. **Linha — "Reserva acumulada"**
   - `investimentos[Mês]` (eixo) + `investimentos[Acumulado]` (valores) + `investimentos[Depósito]` (colunas) → usar **Gráfico de Linhas e Colunas Empilhadas**

### Tabela — "Contas a pagar"
- `despesas[Item]`, `[Categoria]`, `[Valor]`, `[Status]`
- **Formatação condicional** no Status: "A confirmar" = amarelo, "Quitado" = verde
- Ativar **Totais** na formatação

### Segmentações (slicers)
- `despesas[Categoria]` → estilo Lista ou Menu Suspenso
- `despesas[Status]` → botões

## 4. Coluna calculada opcional (marcar pago no PBIX)

No modo Tabela, na `despesas`, crie a coluna:

```
Quitado Manual = NO
```

E use o recurso **"Digitação rápida/entrada manual"** (ou edite os dados no
Power Query) — ou simplesmente marque o Status com valores `Quitado` /
`Pendente` conforme for pagando e clique em **Atualizar**.

## 5. Estilo visual recomendado (tema dark)

Crie o tema em **Exibição → Temas → Procurar temas** ou aplique manualmente:

| Elemento | Valor |
|---|---|
| Plano de fundo da página | `#0B1220` |
| Plano de fundo dos visuais | `#111A2E` |
| Borda dos visuais | `#24304F` |
| Texto primário | `#E6ECF7` |
| Texto secundário | `#93A3C4` |
| Verde (positivo) | `#22C55E` |
| Vermelho (despesa) | `#EF4444` |
| Amarelo (atenção) | `#F59E0B` |
| Azul | `#3B82F6` |
| Roxo (quitáveis) | `#A78BFA` |
| Ciano (sobra) | `#22D3EE` |

## 6. Regenerar os CSVs

Se o `dados.json` mudar (ex.: Brasilcard/C6 confirmados), regenere:

```powershell
python dashboard-financeiro\powerbi\gerar_csvs.py
```

E no Power BI: **Página Inicial → Atualizar**.
