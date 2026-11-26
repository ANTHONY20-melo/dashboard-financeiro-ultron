# -*- coding: utf-8 -*-
"""Gera tudo: dashboard HTML (offline) + planilha Excel + CSV para Power BI.

Uso:
    python gerar_tudo.py
    python gerar_tudo.py --abrir   # abre o dashboard no navegador
"""
import sys
import webbrowser

from gerar_dashboard import gerar as gerar_dashboard
from gerar_planilha import gerar as gerar_planilha


def main() -> None:
    html = gerar_dashboard()
    xlsx, csv = gerar_planilha()

    print()
    print("=" * 66)
    print("  DASHBOARD FINANCEIRO — ENTREGAS GERADAS")
    print("=" * 66)
    for rotulo, arquivo in (
        ("Dashboard HTML", html),
        ("Planilha Excel", xlsx),
        ("CSV Power BI  ", csv),
    ):
        print(f"  {rotulo}: {arquivo}  ({arquivo.stat().st_size / 1024:.1f} KB)")
    print("=" * 66)

    if "--abrir" in sys.argv:
        webbrowser.open(html.as_uri())


if __name__ == "__main__":
    main()
