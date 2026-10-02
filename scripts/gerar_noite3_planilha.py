"""Gera o modelo de planilha de orçamento da Noite 3 (.ods).

Uma aba por cenário orçado (cada equipe orça 2). Colunas fixas:
Item · Especificação · Qtd · Unid. · Preço unit. · Total · Loja · Link
"""
import os

from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import (
    Style, TableColumnProperties, TableCellProperties, TextProperties,
    ParagraphProperties,
)
from odf.number import CurrencyStyle, CurrencySymbol, Number, Text as NText
from odf.table import Table, TableColumn, TableRow, TableCell, CoveredTableCell
from odf.text import P

OUTDIR = "/home/nedel/curso"
VERDE = "#0B5C5C"
LARANJA = "#E88B1A"
CINZA_CLARO = "#EEF3F3"

# Linhas sugeridas: (item, especificação de exemplo, unidade)
ITENS = [
    ("Cabo de rede", "Ex.: Cat5e ou Cat6, por metro", "m"),
    ("Conector RJ45", "Mesma categoria do cabo", "un"),
    ("Capa protetora do conector", "Opcional", "un"),
    ("Roteador", "Ex.: dual band, portas Gigabit", "un"),
    ("Switch", "Se precisar de mais portas", "un"),
    ("Repetidor / kit mesh / ponto de acesso", "Se precisar de cobertura", "un"),
    ("Canaleta / eletroduto", "Por onde o cabo vai passar", "m"),
    ("Tomada / ponto de rede (keystone)", "Se for embutir na parede", "un"),
    ("", "", ""),
    ("", "", ""),
    ("Frete", "Ou escrever \"retirada na loja\"", "-"),
    ("Mão de obra", "Por ponto instalado", "ponto"),
]


def build():
    doc = OpenDocumentSpreadsheet()

    moeda = CurrencyStyle(name="moeda", volatile="true")
    moeda.addElement(CurrencySymbol(language="pt", country="BR", text="R$"))
    moeda.addElement(NText(text=" "))
    moeda.addElement(Number(decimalplaces="2", minintegerdigits="1", grouping="true"))
    doc.automaticstyles.addElement(moeda)

    def cell_style(name, bg=None, bold=False, color=None, size="11pt", money=False, wrap=False, border=True):
        st = Style(name=name, family="table-cell", **({"datastylename": "moeda"} if money else {}))
        props = {}
        if bg:
            props["backgroundcolor"] = bg
        if border:
            props["border"] = "0.5pt solid #999999"
        if wrap:
            props["wrapoption"] = "wrap"
        props["verticalalign"] = "middle"
        st.addElement(TableCellProperties(**props))
        tp = {"fontsize": size}
        if bold:
            tp["fontweight"] = "bold"
        if color:
            tp["color"] = color
        st.addElement(TextProperties(**tp))
        doc.automaticstyles.addElement(st)
        return st

    s_titulo = cell_style("titulo", bold=True, color=VERDE, size="16pt", border=False)
    s_kicker = cell_style("kicker", bold=True, color=LARANJA, size="10pt", border=False)
    s_rotulo = cell_style("rotulo", bold=True, bg=CINZA_CLARO)
    s_cab = cell_style("cab", bg=VERDE, bold=True, color="#FFFFFF")
    s_txt = cell_style("txt")
    s_dica = cell_style("dica", color="#888888")
    s_num = cell_style("num")
    s_money = cell_style("money", money=True)
    s_total_rot = cell_style("totalrot", bg=LARANJA, bold=True, color="#FFFFFF", size="13pt")
    s_total = cell_style("total", bg=LARANJA, bold=True, color="#FFFFFF", size="13pt", money=True)
    s_just = cell_style("just", wrap=True)
    s_aviso = cell_style("aviso", color="#C5362B", bold=True, border=False)

    larguras = ["5.5cm", "6cm", "1.6cm", "1.6cm", "3cm", "3.2cm", "4cm", "7cm"]
    col_styles = []
    for i, w in enumerate(larguras):
        cs = Style(name=f"col{i}", family="table-column")
        cs.addElement(TableColumnProperties(columnwidth=w))
        doc.automaticstyles.addElement(cs)
        col_styles.append(cs)

    def txt(value, style):
        c = TableCell(stylename=style, valuetype="string")
        c.addElement(P(text=value))
        return c

    def empty(style=None):
        return TableCell(stylename=style) if style else TableCell()

    def row(*cells):
        r = TableRow()
        for c in cells:
            r.addElement(c)
        return r

    def merged(value, style, span):
        c = TableCell(stylename=style, valuetype="string", numbercolumnsspanned=str(span))
        c.addElement(P(text=value))
        return [c] + [CoveredTableCell() for _ in range(span - 1)]

    for aba in ("Cenário A", "Cenário B"):
        t = Table(name=aba)
        for cs in col_styles:
            t.addElement(TableColumn(stylename=cs))

        t.addElement(row(txt("PRÁTICA · NOITE 3 · REDES DE COMPUTADORES", s_kicker)))
        t.addElement(row(txt("Orçamento de instalação de rede", s_titulo)))
        t.addElement(row(empty()))
        t.addElement(row(txt("Equipe:", s_rotulo), *merged("", s_txt, 3)))
        t.addElement(row(txt("Cenário nº / cliente:", s_rotulo), *merged("", s_txt, 3)))
        t.addElement(row(txt("Meio escolhido:", s_rotulo), *merged("Ex.: cabo Cat6 até o escritório + Wi-Fi na loja", s_dica, 3)))
        t.addElement(row(empty()))

        cab = ["Item", "Especificação", "Qtd", "Unid.", "Preço unit.", "Total", "Loja", "Link"]
        t.addElement(row(*[txt(h, s_cab) for h in cab]))

        primeira = 9  # linha (1-based) do primeiro item
        for i, (item, espec, unid) in enumerate(ITENS):
            n = primeira + i
            total = TableCell(stylename=s_money, valuetype="currency", currency="BRL",
                              formula=f"of:=[.C{n}]*[.E{n}]")
            t.addElement(row(
                txt(item, s_txt) if item else empty(s_txt),
                txt(espec, s_dica) if espec else empty(s_txt),
                empty(s_num),
                txt(unid, s_txt) if unid else empty(s_txt),
                empty(s_money),
                total,
                empty(s_txt),
                empty(s_txt),
            ))
        ultima = primeira + len(ITENS) - 1

        total = TableCell(stylename=s_total, valuetype="currency", currency="BRL",
                          formula=f"of:=SUM([.F{primeira}:.F{ultima}])")
        t.addElement(row(empty(), empty(), empty(), empty(), txt("TOTAL", s_total_rot), total))
        t.addElement(row(empty()))

        t.addElement(row(txt("Suposições (metros, distâncias...):", s_rotulo), *merged("", s_just, 7)))
        t.addElement(row(empty()))
        just_row = TableRow()
        just_row.addElement(txt("Por que escolhemos essa solução:", s_rotulo))
        for c in merged("", s_just, 7):
            just_row.addElement(c)
        t.addElement(just_row)
        t.addElement(row(empty()))
        t.addElement(row(txt("Orçamento sem link e sem justificativa não vale ponto.", s_aviso)))

        doc.spreadsheet.addElement(t)

    path = os.path.join(OUTDIR, "Planilha - Redes 3 - Modelo de orcamento.ods")
    doc.save(path)
    return path


if __name__ == "__main__":
    print("  →", build())
