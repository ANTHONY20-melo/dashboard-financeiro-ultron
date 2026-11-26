# -*- coding: utf-8 -*-
"""Gerador do Dashboard Financeiro — Novembro/2026 + planejamento anual.

Lê dados.json (fonte da verdade extraída do PDF DOC-20261110-WA0109.pdf),
injeta no template.html e escreve o dashboard HTML standalone (offline,
com Chart.js embutido).

Uso:
    python gerar_dashboard.py            # gera o HTML
    python gerar_dashboard.py --abrir    # gera e abre no navegador
"""
import json
import pathlib
import shutil
import sys
import webbrowser

BASE = pathlib.Path(__file__).resolve().parent
DADOS = BASE / "dados.json"
TEMPLATE = BASE / "template.html"
CHART_JS = BASE / "vendor" / "chart.umd.min.js"
SAIDA = BASE / "dashboard-financeiro-nov-2026.html"
SAIDA_GENERICA = BASE / "dashboard-financeiro.html"

TOKEN = "__DADOS_JSON__"
TOKEN_CHART = "<!--CHARTJS-->"
CDN_CHART = "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"


def montar_html() -> str:
    dados = json.loads(DADOS.read_text(encoding="utf-8"))
    template = TEMPLATE.read_text(encoding="utf-8")
    if TOKEN not in template:
        raise SystemExit(f"ERRO: token {TOKEN} não encontrado no template.html")

    html = template.replace(TOKEN, json.dumps(dados, ensure_ascii=False))

    if TOKEN_CHART in html:
        if CHART_JS.exists():
            js = CHART_JS.read_text(encoding="utf-8")
            html = html.replace(TOKEN_CHART, f"<script>{js}</script>")
        else:
            html = html.replace(TOKEN_CHART, f'<script src="{CDN_CHART}"></script>')

    return html


def gerar() -> pathlib.Path:
    html = montar_html()
    SAIDA.write_text(html, encoding="utf-8")
    shutil.copyfile(SAIDA, SAIDA_GENERICA)
    return SAIDA


def main() -> None:
    saida = gerar()
    tamanho = saida.stat().st_size / 1024
    print(f"OK  {saida}  ({tamanho:.1f} KB)")
    print(f"OK  {SAIDA_GENERICA}")
    if "--abrir" in sys.argv:
        webbrowser.open(saida.as_uri())


if __name__ == "__main__":
    main()
