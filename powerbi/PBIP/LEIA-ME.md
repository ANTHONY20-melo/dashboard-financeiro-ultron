# 📁 PBIP — Dashboard Financeiro (Novembro/2026)

Projeto **Power BI Desktop** no formato *Power BI Project* (texto aberto).

## Abrir

Dê duplo clique em **`DashboardFinanceiro.pbip`** (ou abra pelo Power BI Desktop:
Arquivo → Abrir relatório → Procurar → este arquivo).

Ao abrir, o Power BI carrega automaticamente:
- **Modelo:** 4 tabelas importadas dos CSVs (`receitas`, `despesas`, `investimentos`, `resumo`)
- **Medidas:** 15 medidas DAX prontas
- **Relatório:** página "Novembro 2026" com 6 cartões KPI, 2 roscas, 2 barras,
  1 linha (reserva), 1 tabela de contas e 1 segmentador por categoria

## Se os dados vierem vazios

O modelo importa os CSVs por caminho absoluto. Se você mover a pasta do
projeto, atualize o caminho em cada `tables/*.tmdl` (propriedade `source`)
ou clique em **Página Inicial → Atualizar** após abrir.

## Regenerar visuais

```powershell
python gerar_visuais.py
```

## Estrutura

```
PBIP/
├── DashboardFinanceiro.pbip                ← ABRIR ESTE
├── DashboardFinanceiro.SemanticModel/      ← modelo (TMDL)
│   ├── definition.pbism
│   └── definition/
│       ├── database.tmdl
│       ├── model.tmdl
│       └── tables/ (receitas, despesas, investimentos, resumo)
├── DashboardFinanceiro.Report/             ← relatório (PBIR)
│   ├── definition.pbir
│   └── definition/
│       ├── report.json / version.json
│       └── pages/mainPage/ (page.json + 13 visuais)
└── gerar_visuais.py                        ← gerador dos visuais
```
