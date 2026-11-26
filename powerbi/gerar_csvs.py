# -*- coding: utf-8 -*-
"""Gera CSVs para importação no Power BI a partir de dados.json.

Formato: UTF-8 com BOM, delimitador ponto-e-vírgula, decimais com vírgula
(padron pt-BR — o Power BI detecta com locale "Português (Brasil)").

Uso:
    python gerar_csvs.py
"""
import csv
import json
import pathlib

BASE = pathlib.Path(__file__).resolve().parent
DADOS = BASE.parent / "dados.json"

BRL = lambda v: "" if v is None else f"{v:.2f}".replace(".", ",")


def salvar(nome: str, cabecalho: list[str], linhas: list[list]) -> None:
    destino = BASE / nome
    with destino.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        w.writerow(cabecalho)
        w.writerows(linhas)
    print(f"[OK] {destino.name}: {len(linhas)} linhas")


def main() -> None:
    d = json.loads(DADOS.read_text(encoding="utf-8"))

    # ---------- receitas.csv ----------
    receitas = []
    for r in d["receitas"]:
        receitas.append([r["desc"], BRL(r["bruto"]), BRL(r["dizimo"]), BRL(r["liquido"]), r["obs"]])
    receitas.append(["TOTAL PREVISTO", BRL(d["totalReceitas"]["bruto"]), BRL(d["totalReceitas"]["dizimo"]), BRL(d["totalReceitas"]["liquido"]), ""])
    salvar("receitas.csv", ["Descrição", "Valor Bruto", "Dízimo", "Líquido", "Observação"], receitas)

    # ---------- despesas.csv ----------
    despesas = []
    for p in d["prioridades"]:
        despesas.append(["Prioridade", p["desc"], BRL(p["valor"]), p["obs"], "Pendente"])
    for p in d["pessoais"]:
        status = "A confirmar" if p["valor"] is None else "Pendente"
        despesas.append(["Pessoal", p["desc"], BRL(p["valor"]), p["obs"], status])
    for q in d["quitaveis"]:
        status = "A confirmar" if q["valor"] is None else "Pendente"
        despesas.append(["Quitável", q["desc"], BRL(q["valor"]), q["obs"], status])
    despesas.append(["Extra", d["comprasPrevistas"][0]["desc"], BRL(d["comprasPrevistas"][0]["valor"]), "Previsto no documento", "Pendente"])
    salvar("despesas.csv", ["Categoria", "Item", "Valor", "Observação", "Status"], despesas)

    # ---------- investimentos.csv ----------
    acumulado = 0.0
    invest = []
    for i in d["investimentos"]:
        dep = float(i["deposito"] or 0)
        acumulado += dep
        invest.append([i["mes"], BRL(i["sobra"]), BRL(i["deposito"]) if dep else "", BRL(acumulado)])
    salvar("investimentos.csv", ["Mês", "Sobra Líquida", "Depósito", "Acumulado"], invest)

    # ---------- resumo.csv (KPIs prontos para cards) ----------
    liquido = float(d["totalReceitas"]["liquido"])
    contas = float(d["totalContasNovembro"])
    roupas = float(d["comprasPrevistas"][0]["valor"])
    sobra = liquido - contas - roupas
    quitaveis = float(d["subtotalQuitaveis"])
    resumo = [
        ["Receita Bruta", BRL(d["totalReceitas"]["bruto"]), "Total bruto das receitas"],
        ["Dízimo (10%)", BRL(d["totalReceitas"]["dizimo"]), "Descontado no recebimento"],
        ["Receita Líquida", BRL(liquido), "Após dízimo"],
        ["Contas de Novembro", BRL(contas), "Prioridades + pessoais (sem Brasilcard)"],
        ["Roupas (previsto)", BRL(roupas), "Compra prevista"],
        ["Sobra Prevista", BRL(sobra), "Líquido - contas - roupas"],
        ["Dívidas Quitáveis", BRL(quitaveis), "PAN x3 + BMG (C6 a confirmar)"],
        ["Sobra após quitar tudo", BRL(sobra - quitaveis), "Se quitar todas as quitáveis"],
        ["Rescisão Corrigida", BRL(d["rescisaoCorrigida"]), "Valor correto (não R$ 8.364,58)"],
        ["Reserva 3 meses", "450,00", "3 depósitos de R$ 150 (dez/26 a fev/27)"],
    ]
    salvar("resumo.csv", ["Indicador", "Valor", "Descrição"], resumo)


if __name__ == "__main__":
    main()
