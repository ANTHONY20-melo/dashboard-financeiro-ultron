# -*- coding: utf-8 -*-
"""Gera a planilha Excel do controle financeiro (2026-2027) a partir de dados.json.

Abas:
  - Resumo Anual   : 12 meses x 2 anos, com fórmulas de sobra e totais
  - Nov 2026       : detalhamento completo do documento base
  - Lançamentos    : tabela livre (com validações) para o dia a dia
  - Investimentos  : projeção da reserva mensal
  - Instruções     : como usar

Uso:
    python gerar_planilha.py
"""
import json
import pathlib
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

BASE = pathlib.Path(__file__).resolve().parent
DADOS = BASE / "dados.json"
SAIDA = BASE / "planilha-financeira-2026-2027.xlsx"

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

AZUL = "1D4ED8"
AZUL_ESCURO = "0F172A"
CINZA = "F1F5F9"
CINZA2 = "E2E8F0"
VERDE = "16A34A"
VERMELHO = "DC2626"
AMARELO = "B45309"
BRANCO = "FFFFFF"

MOEDA = 'R$ #,##0.00'
BORDA_FINA = Border(*[Side(style="thin", color=CINZA2)] * 4)


def titulo(ws, texto: str, ncols: int, linha: int = 1) -> None:
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=ncols)
    c = ws.cell(row=linha, column=1, value=texto)
    c.font = Font(bold=True, size=15, color=BRANCO)
    c.fill = PatternFill("solid", fgColor=AZUL_ESCURO)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[linha].height = 30


def secao(ws, texto: str, ncols: int, linha: int) -> None:
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=ncols)
    c = ws.cell(row=linha, column=1, value=texto)
    c.font = Font(bold=True, size=11, color=BRANCO)
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[linha].height = 22


def cabecalho(ws, linha: int, colunas: list[str]) -> None:
    for i, nome in enumerate(colunas, start=1):
        c = ws.cell(row=linha, column=i, value=nome)
        c.font = Font(bold=True, size=10, color=BRANCO)
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA_FINA
    ws.row_dimensions[linha].height = 26


def larguras(ws, mapa: dict[str, int]) -> None:
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


def moeda(c) -> None:
    c.number_format = MOEDA
    c.alignment = Alignment(horizontal="right")


def zebra(ws, linha_ini: int, linha_fim: int, ncols: int) -> None:
    for r in range(linha_ini, linha_fim + 1):
        if (r - linha_ini) % 2 == 1:
            for c in range(1, ncols + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=CINZA)


def preparar_impressao(ws, area: str, orientacao: str = "landscape") -> None:
    """Deixa a aba pronta para imprimir/exportar em PDF sem dezenas de páginas."""
    ws.print_area = area
    ws.page_setup.orientation = orientacao
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "&A · página &P de &N"
    ws.oddFooter.center.size = 8
    ws.oddFooter.center.color = "64748B"


def aba_resumo(wb: Workbook, dados: dict) -> None:
    ws = wb.active
    ws.title = "Resumo Anual"
    titulo(ws, "Dashboard Financeiro · Planejamento Anual 2026-2027", 7)
    ws["A2"] = f"Fonte: {dados.get('fonte', '—')} · gerado em {date.today():%d/%m/%Y}"
    ws["A2"].font = Font(size=9, color="64748B")

    cabecalho(ws, 4, ["Ano", "Mês", "Receita (R$)", "Despesa (R$)", "Sobra (R$)", "Status", "Observação"])

    linha = 5
    for ano, meses in sorted(dados.get("planejamento", {}).items()):
        for i, nome in enumerate(MESES):
            chave = f"{i + 1:02d}"
            reg = meses.get(chave, {})
            ws.cell(row=linha, column=1, value=int(ano)).alignment = Alignment(horizontal="center")
            ws.cell(row=linha, column=2, value=nome)
            rec = ws.cell(row=linha, column=3, value=reg.get("receita"))
            des = ws.cell(row=linha, column=4, value=reg.get("despesa"))
            sob = ws.cell(row=linha, column=5, value=f'=IF(COUNT(C{linha}:D{linha})=0,"",C{linha}-D{linha})')
            st = ws.cell(row=linha, column=6, value=f'=IF(E{linha}="","—",IF(E{linha}>=0,"Positivo","Negativo"))')
            obs = ws.cell(row=linha, column=7, value=reg.get("obs", ""))
            for c in (rec, des, sob):
                moeda(c)
            st.alignment = Alignment(horizontal="center")
            obs.alignment = Alignment(wrap_text=True, vertical="top")
            for c in range(1, 8):
                ws.cell(row=linha, column=c).border = BORDA_FINA
            linha += 1

    fim = linha - 1
    zebra(ws, 5, fim, 7)

    # formatação condicional da sobra
    ws.conditional_formatting.add(
        f"E5:E{fim}",
        CellIsRule(operator="lessThan", formula=["0"],
                   font=Font(color=VERMELHO, bold=True),
                   fill=PatternFill("solid", fgColor="FEE2E2")),
    )
    ws.conditional_formatting.add(
        f"E5:E{fim}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"],
                   font=Font(color=VERDE, bold=True)),
    )

    # totais por ano
    r = fim + 2
    secao(ws, "Totais por ano", 7, r)
    cabecalho(ws, r + 1, ["Ano", "Receita", "Despesa", "Sobra", "Meses preenchidos", "", ""])
    anos = sorted(dados.get("planejamento", {}).keys())
    for j, ano in enumerate(anos):
        rr = r + 2 + j
        ws.cell(row=rr, column=1, value=int(ano)).alignment = Alignment(horizontal="center")
        ws.cell(row=rr, column=2, value=f'=SUMIF($A$5:$A${fim},A{rr},$C$5:$C${fim})')
        ws.cell(row=rr, column=3, value=f'=SUMIF($A$5:$A${fim},A{rr},$D$5:$D${fim})')
        ws.cell(row=rr, column=4, value=f'=B{rr}-C{rr}')
        ws.cell(row=rr, column=5, value=f'=COUNTIFS($A$5:$A${fim},A{rr},$C$5:$C${fim},"<>")+COUNTIFS($A$5:$A${fim},A{rr},$D$5:$D${fim},"<>")-COUNTIFS($A$5:$A${fim},A{rr},$C$5:$C${fim},"<>",$D$5:$D${fim},"<>")')
        for c in (2, 3, 4):
            moeda(ws.cell(row=rr, column=c))
        ws.cell(row=rr, column=5).alignment = Alignment(horizontal="center")

    # indicadores laterais
    ws["I4"] = "Indicadores"
    ws["I4"].font = Font(bold=True, size=11, color=BRANCO)
    ws["I4"].fill = PatternFill("solid", fgColor=AZUL)
    ws.merge_cells("I4:J4")
    indicadores = [
        ("Receita total", f"=SUM(B{r + 2}:B{r + 1 + len(anos)})"),
        ("Despesa total", f"=SUM(C{r + 2}:C{r + 1 + len(anos)})"),
        ("Sobra total", f"=SUM(D{r + 2}:D{r + 1 + len(anos)})"),
        ("Meses preenchidos", f"=SUM(E{r + 2}:E{r + 1 + len(anos)})"),
        ("Reserva prevista/ano", (dados.get("metaReservaMensal") or 0) * 12),
        ("Meta mensal de reserva", dados.get("metaReservaMensal") or 0),
    ]
    for k, (rot, val) in enumerate(indicadores):
        rr = 5 + k
        a = ws.cell(row=rr, column=9, value=rot)
        a.font = Font(size=9, color="475569")
        b = ws.cell(row=rr, column=10, value=val)
        b.font = Font(bold=True)
        if "Meses" in rot:
            b.alignment = Alignment(horizontal="center")
        else:
            moeda(b)
        a.border = BORDA_FINA
        b.border = BORDA_FINA

    larguras(ws, {"A": 8, "B": 13, "C": 15, "D": 15, "E": 15, "F": 12, "G": 52,
                  "H": 2, "I": 22, "J": 16})
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A4:G{fim}"
    preparar_impressao(ws, "A1:J33")


def aba_novembro(wb: Workbook, dados: dict) -> None:
    ws = wb.create_sheet("Nov 2026")
    titulo(ws, "Detalhamento — Novembro/2026", 5)
    ws["A2"] = "Documento base: " + dados.get("fonte", "—")
    ws["A2"].font = Font(size=9, color="64748B")

    r = 6
    secao(ws, "1. Receitas — descontar 10% (dízimo) no recebimento", 5, r)
    r += 1
    cabecalho(ws, r, ["Descrição", "Valor bruto", "Dízimo 10%", "Líquido", "Recebido?"])
    r += 1
    ini_rec = r
    for rec in dados["receitas"]:
        ws.cell(row=r, column=1, value=rec["desc"])
        moeda(ws.cell(row=r, column=2, value=rec["bruto"]))
        moeda(ws.cell(row=r, column=3, value=rec["dizimo"]))
        moeda(ws.cell(row=r, column=4, value=rec["liquido"]))
        ws.cell(row=r, column=5, value="Não").alignment = Alignment(horizontal="center")
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDA_FINA
        r += 1
    fim_rec = r - 1
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    for col in (2, 3, 4):
        c = ws.cell(row=r, column=col, value=f"=SUM({get_column_letter(col)}{ini_rec}:{get_column_letter(col)}{fim_rec})")
        c.font = Font(bold=True, color=VERDE)
        moeda(c)
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="DCFCE7")
    total_receitas_row = r

    def bloco(titulo_txt: str, itens: list[dict], col_valor: str = "valor",
              col_obs: str = "obs", col_status: str = "Pago?") -> tuple[int, int, int]:
        nonlocal r
        r += 1
        secao(ws, titulo_txt, 5, r)
        r += 1
        cabecalho(ws, r, ["Item", "Valor", "Observação", col_status])
        r += 1
        ini = r
        for it in itens:
            ws.cell(row=r, column=1, value=it["desc"])
            v = it.get(col_valor)
            if v is None:
                c = ws.cell(row=r, column=2, value="a confirmar")
                c.font = Font(italic=True, color=AMARELO)
                c.alignment = Alignment(horizontal="right")
            else:
                moeda(ws.cell(row=r, column=2, value=v))
            ws.cell(row=r, column=3, value=it.get(col_obs, "")).alignment = Alignment(wrap_text=True, vertical="top", horizontal="left", indent=1)
            ws.cell(row=r, column=4, value="Não").alignment = Alignment(horizontal="center")
            for c in range(1, 5):
                ws.cell(row=r, column=c).border = BORDA_FINA
            r += 1
        fim = r - 1
        ws.cell(row=r, column=1, value="SUBTOTAL").font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=f"=SUM(B{ini}:B{fim})")
        c.font = Font(bold=True, color=VERDE)
        moeda(c)
        for c2 in range(1, 5):
            ws.cell(row=r, column=c2).fill = PatternFill("solid", fgColor="DCFCE7")
        return r, ini, fim

    sub_pri, pri_ini, pri_fim = bloco("2.1 Prioridades de novembro", dados["prioridades"])
    sub_pes, pes_ini, pes_fim = bloco("2.2 Despesas pessoais / parcelas", dados["pessoais"])
    sub_qui, qui_ini, qui_fim = bloco("2.3 Dívidas quitáveis (opcional)", dados["quitaveis"], col_status="Quitado?")

    # reserva
    r += 1
    secao(ws, "3. Reserva e investimentos", 5, r)
    r += 1
    cabecalho(ws, r, ["Mês", "Sobra líquida", "Depósito Nubank", "Acumulado", ""])
    r += 1
    ini_inv = r
    acum = 0
    for inv in dados["investimentos"]:
        ws.cell(row=r, column=1, value=inv["mes"])
        if inv.get("sobra") is None:
            ws.cell(row=r, column=2, value="—").alignment = Alignment(horizontal="right")
        else:
            moeda(ws.cell(row=r, column=2, value=inv["sobra"]))
        dep = inv.get("deposito")
        if dep is None:
            ws.cell(row=r, column=3, value="—").alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=4, value="—").alignment = Alignment(horizontal="right")
        else:
            acum += dep
            moeda(ws.cell(row=r, column=3, value=dep))
            moeda(ws.cell(row=r, column=4, value=acum))
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDA_FINA
        r += 1

    # resumo compacto no topo (linhas 3-4)
    rotulos = ["Receita líquida", "Contas (prio + pessoais)", "Roupas", "Sobra prevista", "Quitáveis (opcional)"]
    for i, rot in enumerate(rotulos, start=1):
        c = ws.cell(row=3, column=i, value=rot)
        c.font = Font(size=8, bold=True, color="475569")
        c.fill = PatternFill("solid", fgColor=CINZA)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA_FINA
    valores = [
        f"=D{total_receitas_row}",
        f"=B{sub_pri}+B{sub_pes}",
        dados["comprasPrevistas"][0]["valor"],
        "=A4-B4-C4",
        f"=B{sub_qui}",
    ]
    for i, val in enumerate(valores, start=1):
        c = ws.cell(row=4, column=i, value=val)
        c.font = Font(bold=True, size=11, color=VERDE if i == 4 else "111111")
        moeda(c)
        c.fill = PatternFill("solid", fgColor="DCFCE7" if i == 4 else BRANCO)
        c.border = BORDA_FINA
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 24

    # validações Sim/Não
    dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E{ini_rec}:E{fim_rec}")
    dv.add(f"D{pri_ini}:D{pri_fim}")
    dv.add(f"D{pes_ini}:D{pes_fim}")
    dv.add(f"D{qui_ini}:D{qui_fim}")

    larguras(ws, {"A": 40, "B": 16, "C": 60, "D": 14, "E": 11})
    ws.sheet_view.showGridLines = False
    preparar_impressao(ws, f"A1:E{r}", "portrait")


def aba_lancamentos(wb: Workbook) -> None:
    ws = wb.create_sheet("Lançamentos")
    titulo(ws, "Lançamentos do dia a dia", 7)
    ws["A2"] = "Preencha uma linha por movimento. Use os filtros do cabeçalho para separar por mês, tipo ou status."
    ws["A2"].font = Font(size=9, color="64748B")
    cabecalho(ws, 4, ["Data", "Tipo", "Categoria", "Descrição", "Valor (R$)", "Status", "Observação"])

    for r in range(5, 305):
        ws.cell(row=r, column=1).number_format = "DD/MM/AAAA"
        moeda(ws.cell(row=r, column=5))
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = BORDA_FINA

    tab = Table(displayName="Lancamentos", ref="A4:G304")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    dv_tipo = DataValidation(type="list", formula1='"Receita,Despesa,Investimento"', allow_blank=True)
    dv_status = DataValidation(type="list", formula1='"Pago,Pendente,A confirmar"', allow_blank=True)
    ws.add_data_validation(dv_tipo)
    ws.add_data_validation(dv_status)
    dv_tipo.add("B5:B304")
    dv_status.add("F5:F304")

    larguras(ws, {"A": 12, "B": 14, "C": 20, "D": 44, "E": 15, "F": 14, "G": 40})
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    preparar_impressao(ws, "A1:G20")


def aba_investimentos(wb: Workbook, dados: dict) -> None:
    ws = wb.create_sheet("Investimentos")
    titulo(ws, "Reserva / investimento mensal", 4)
    meta = dados.get("metaReservaMensal", 0)
    ws["A2"] = f"Meta: R$ {meta:,.2f} por mês na caixinha".replace(",", "X").replace(".", ",").replace("X", ".")
    ws["A2"].font = Font(size=9, color="64748B")
    cabecalho(ws, 4, ["Mês", "Sobra líquida", "Depósito (R$)", "Acumulado (R$)"])

    r = 5
    acum = 0.0
    for inv in dados.get("investimentos", []):
        ws.cell(row=r, column=1, value=inv["mes"])
        if inv.get("sobra") is None:
            ws.cell(row=r, column=2, value="—").alignment = Alignment(horizontal="right")
        else:
            moeda(ws.cell(row=r, column=2, value=inv["sobra"]))
        dep = inv.get("deposito")
        if dep is None:
            ws.cell(row=r, column=3, value="—").alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=4, value="—").alignment = Alignment(horizontal="right")
        else:
            acum += dep
            moeda(ws.cell(row=r, column=3, value=dep))
            moeda(ws.cell(row=r, column=4, value=f"=SUM($C$5:C{r})"))
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDA_FINA
        r += 1

    # reserva futura: 12 meses de meta a partir do fim da lista
    r += 1
    secao(ws, "Projeção — 12 meses de meta", 4, r)
    r += 1
    cabecalho(ws, r, ["Mês", "Sobra líquida", "Depósito (R$)", "Acumulado (R$)"])
    r += 1
    for k in range(12):
        ws.cell(row=r, column=1, value=f"Mês +{k + 1}")
        ws.cell(row=r, column=2, value="—").alignment = Alignment(horizontal="right")
        moeda(ws.cell(row=r, column=3, value=meta))
        moeda(ws.cell(row=r, column=4, value=f"=SUM($C$5:C{r})"))
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDA_FINA
        r += 1

    larguras(ws, {"A": 14, "B": 16, "C": 16, "D": 18})
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    preparar_impressao(ws, f"A1:D{r - 1}", "portrait")


def aba_instrucoes(wb: Workbook) -> None:
    ws = wb.create_sheet("Instruções")
    titulo(ws, "Como usar esta planilha", 2)
    linhas = [
        ("", ""),
        ("1. Resumo Anual", "Visão dos 12 meses de 2026 e 2027. Preencha receita e despesa de cada mês; a sobra e os status são calculados automaticamente."),
        ("2. Nov 2026", "Detalhamento do documento base (receitas, prioridades, despesas pessoais, dívidas quitáveis e reserva). Marque Sim/Não nas colunas de status."),
        ("3. Lançamentos", "Use no dia a dia: data, tipo (Receita/Despesa/Investimento), categoria, descrição, valor e status. As listas já têm validação."),
        ("4. Investimentos", "Projeção da reserva mensal na caixinha. O acumulado é calculado por fórmula."),
        ("", ""),
        ("Regras de ouro", "• Descontar 10% de dízimo no recebimento."),
        ("", "• Pedir valor atualizado com desconto antes de quitar qualquer empréstimo."),
        ("", "• Confirmar Brasilcard e C6 no extrato antes de pagar."),
        ("", "• Acumular reserva antes de usar o dinheiro na entrada do carro (projeto 2027)."),
        ("", ""),
        ("Dica", "Para levar os lançamentos do dashboard HTML para cá, use o botão 'Exportar CSV' e cole/importe na aba Lançamentos."),
        ("", ""),
        ("Atenção", "Esta planilha é um documento de planejamento pessoal. Confirme os valores finais com os extratos bancários."),
    ]
    for i, (titulo_linha, texto) in enumerate(linhas, start=3):
        a = ws.cell(row=i, column=1, value=titulo_linha)
        a.font = Font(bold=True, color=AZUL)
        b = ws.cell(row=i, column=2, value=texto)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        b.font = Font(size=10)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 110
    ws.sheet_view.showGridLines = False
    preparar_impressao(ws, "A1:B18", "portrait")


def gerar_csv_anual(dados: dict) -> pathlib.Path:
    """CSV do planejamento anual (Power BI / Excel), padrão pt-BR."""
    pasta = BASE / "powerbi"
    pasta.mkdir(exist_ok=True)
    saida = pasta / "planejamento_anual.csv"
    linhas = ["Ano;Mês;Receita;Despesa;Sobra;Observação"]
    for ano, meses in sorted(dados.get("planejamento", {}).items()):
        for i, nome in enumerate(MESES):
            reg = meses.get(f"{i + 1:02d}", {})
            rec = reg.get("receita")
            des = reg.get("despesa")
            sob = None if (rec is None and des is None) else (rec or 0) - (des or 0)
            conv = lambda v: "" if v is None else f"{v:.2f}".replace(".", ",")
            linhas.append(f'{ano};{nome};{conv(rec)};{conv(des)};{conv(sob)};{reg.get("obs", "")}')
    saida.write_text("\ufeff" + "\r\n".join(linhas) + "\r\n", encoding="utf-8")
    return saida


def gerar() -> tuple[pathlib.Path, pathlib.Path]:
    dados = json.loads(DADOS.read_text(encoding="utf-8"))
    wb = Workbook()
    aba_resumo(wb, dados)
    aba_novembro(wb, dados)
    aba_lancamentos(wb)
    aba_investimentos(wb, dados)
    aba_instrucoes(wb)
    wb.save(SAIDA)
    csv = gerar_csv_anual(dados)
    return SAIDA, csv


if __name__ == "__main__":
    xlsx, csv = gerar()
    print(f"OK  {xlsx}  ({xlsx.stat().st_size / 1024:.1f} KB)")
    print(f"OK  {csv}")
