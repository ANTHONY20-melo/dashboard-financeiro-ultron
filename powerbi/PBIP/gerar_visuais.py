# -*- coding: utf-8 -*-
"""Gera os visuais PBIR (visual.json) da página mainPage do projeto PBIP.

Estrutura de cada visual:
    pages/mainPage/visuals/<nome>/visual.json

Uso:
    python gerar_visuais.py
"""
import json
import pathlib

BASE = pathlib.Path(__file__).resolve().parent
VISUAIS = BASE / "DashboardFinanceiro.Report" / "definition" / "pages" / "mainPage" / "visuals"

SCHEMA_VC = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.6.0/schema.json"
SCHEMA_FILTER = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/filterConfiguration/1.2.0/schema.json"


def proj_coluna(entidade: str, coluna: str, ativo: bool = True) -> dict:
    return {
        "field": {
            "Column": {
                "Expression": {"SourceRef": {"Entity": entidade}},
                "Property": coluna,
            }
        },
        "queryRef": f"{entidade}.{coluna}",
        "nativeQueryRef": coluna,
        "active": ativo,
    }


def proj_medida(entidade: str, medida: str, ativo: bool = True) -> dict:
    return {
        "field": {
            "Measure": {
                "Expression": {"SourceRef": {"Entity": entidade}},
                "Property": medida,
            }
        },
        "queryRef": f"{entidade}.{medida}",
        "nativeQueryRef": medida,
        "active": ativo,
    }


def filtro_in(tabela: str, coluna: str, valores: list[str]) -> dict:
    """Filtro categórico 'In' (padrão do schema filterConfiguration 1.2.0)."""
    return {
        "name": f"f{tabela}{coluna}{len(valores)}",
        "type": "Categorical",
        "filter": {
            "Version": 2,
            "From": [{"Name": "d", "Entity": tabela, "Type": 0}],
            "Where": [
                {
                    "Condition": {
                        "In": {
                            "Expressions": [
                                {
                                    "Column": {
                                        "Expression": {"SourceRef": {"Source": "d"}},
                                        "Property": coluna,
                                    }
                                }
                            ],
                            "Values": [
                                [{"Literal": {"Value": f"'{v}'"}}] for v in valores
                            ],
                        }
                    }
                }
            ],
        },
    }


def visual(nome: str, tipo: str, pos: dict, roles: dict, filtro: dict | None = None) -> None:
    corpo: dict = {
        "$schema": SCHEMA_VC,
        "name": nome,
        "position": {**pos, "z": pos.get("z", 1000), "tabOrder": pos.get("tabOrder", 0)},
        "visual": {"visualType": tipo, "query": {"queryState": roles}},
    }
    if filtro:
        corpo["filterConfig"] = {"$schema": SCHEMA_FILTER, "filters": [filtro]}
    pasta = VISUAIS / nome
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / "visual.json"
    destino.write_text(
        json.dumps(corpo, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"[OK] {nome} ({tipo})")


def main() -> None:
    if VISUAIS.exists():
        import shutil

        shutil.rmtree(VISUAIS)
    VISUAIS.mkdir(parents=True)

    # ---------- 6 cartões KPI (linha superior) ----------
    kpis = [
        ("kpiReceitaLiquida", "Receita Líquida Total"),
        ("kpiContasNovembro", "Contas de Novembro"),
        ("kpiSobraPrevista", "Sobra Prevista"),
        ("kpiQuitaveis", "Total Quitáveis"),
        ("kpiRescisao", "Rescisão Corrigida"),
        ("kpiReserva", "Reserva Acumulada"),
    ]
    for i, (nome, medida) in enumerate(kpis):
        visual(
            nome,
            "card",
            {"x": 8 + i * 198, "y": 16, "width": 188, "height": 96, "z": 1000 + i},
            {"Values": {"projections": [proj_medida("resumo", medida)]}},
        )

    # ---------- segmentação por categoria ----------
    visual(
        "slicerCategoria",
        "slicer",
        {"x": 8, "y": 128, "width": 190, "height": 84, "z": 1006},
        {"Category": {"projections": [proj_coluna("despesas", "Categoria")]}},
    )

    # ---------- rosca: composição das receitas ----------
    visual(
        "donutReceitas",
        "donutChart",
        {"x": 208, "y": 128, "width": 300, "height": 236, "z": 1007},
        {
            "Category": {"projections": [proj_coluna("receitas", "Descrição")]},
            "Values": {"projections": [proj_medida("resumo", "Receita Bruta Total")]},
        },
    )

    # ---------- rosca: para onde vai o líquido ----------
    visual(
        "donutUsoLiquido",
        "donutChart",
        {"x": 518, "y": 128, "width": 300, "height": 236, "z": 1008},
        {
            "Category": {"projections": [proj_coluna("despesas", "Categoria")]},
            "Values": {"projections": [proj_medida("resumo", "Total Despesas")]},
        },
        filtro_in("despesas", "Categoria", ["Prioridade", "Pessoal", "Extra"]),
    )

    # ---------- barras: despesas pessoais por item ----------
    visual(
        "barrasPessoais",
        "barChart",
        {"x": 828, "y": 128, "width": 444, "height": 236, "z": 1009},
        {
            "Category": {"projections": [proj_coluna("despesas", "Item")]},
            "Y": {"projections": [proj_medida("resumo", "Total Despesas")]},
        },
        filtro_in("despesas", "Categoria", ["Pessoal"]),
    )

    # ---------- barras: dívidas quitáveis ----------
    visual(
        "barrasQuitaveis",
        "barChart",
        {"x": 8, "y": 378, "width": 300, "height": 334, "z": 1010},
        {
            "Category": {"projections": [proj_coluna("despesas", "Item")]},
            "Y": {"projections": [proj_medida("resumo", "Total Despesas")]},
        },
        filtro_in("despesas", "Categoria", ["Quitável"]),
    )

    # ---------- linha: reserva mensal ----------
    visual(
        "linhaReserva",
        "lineChart",
        {"x": 318, "y": 378, "width": 500, "height": 334, "z": 1011},
        {
            "Category": {"projections": [proj_coluna("investimentos", "Mês")]},
            "Values": {
                "projections": [
                    proj_medida("resumo", "Depósito Mensal"),
                    proj_medida("resumo", "Reserva Acumulada"),
                ]
            },
        },
    )

    # ---------- tabela: contas a pagar ----------
    visual(
        "tabelaContas",
        "tableEx",
        {"x": 828, "y": 378, "width": 444, "height": 334, "z": 1012},
        {
            "Values": {
                "projections": [
                    proj_coluna("despesas", "Categoria", False),
                    proj_coluna("despesas", "Item", False),
                    proj_coluna("despesas", "Valor", False),
                    proj_coluna("despesas", "Status", False),
                ]
            }
        },
    )


if __name__ == "__main__":
    main()
