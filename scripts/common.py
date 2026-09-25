"""Helpers reutilizáveis para gerar aulas do curso de Redes.

Uso:
    from common import COR_PRIMARIA, add_heading, add_para, add_bullet, add_table, add_hr
    from common import build_slide, add_title_block, add_text_block, PALETA
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.shapes import MSO_SHAPE

# ============================================================
# PALETA — cores do curso
# ============================================================
COR_PRIMARIA = RGBColor(0x0B, 0x5C, 0x5C)   # verde-petróleo (título)
COR_ALERTA   = RGBColor(0xC5, 0x36, 0x2B)   # vermelho de alerta
COR_DESTAQUE = RGBColor(0xE8, 0x8B, 0x1A)   # laranja
COR_CINZA    = RGBColor(0x55, 0x55, 0x55)
COR_SUAVE    = RGBColor(0x88, 0x88, 0x88)


class PPalette:
    P = PRGBColor(0x0B, 0x5C, 0x5C)     # primária
    D = PRGBColor(0xE8, 0x8B, 0x1A)     # destaque (laranja)
    W = PRGBColor(0xFF, 0xFF, 0xFF)     # branco
    G = PRGBColor(0x55, 0x55, 0x55)     # cinza texto
    K = PRGBColor(0x1A, 0x1A, 0x1A)     # preto texto
    R = PRGBColor(0xC5, 0x36, 0x2B)     # alerta
    LBG = PRGBColor(0xF4, 0xF7, 0xF7)   # fundo claro
    SOFT_RED = PRGBColor(0xFF, 0xF2, 0xEF)  # fundo suave alerta


# ============================================================
# HELPERS DOCX
# ============================================================
def set_cell_shading(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def add_heading(doc, text, level=1, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text.upper() if level == 1 else text)
    run.bold = True
    run.font.size = Pt(15 if level == 1 else 12)
    run.font.color.rgb = color or COR_PRIMARIA
    return p


def add_para(doc, text, bold=False, italic=False, size=11, color=None, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if align is not None:
        p.alignment = align
    return p


def add_bullet(doc, text, size=11):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p


def add_table(doc, rows_data, header=True, col_widths=None):
    n_cols = len(rows_data[0])
    table = doc.add_table(rows=len(rows_data), cols=n_cols)
    table.style = 'Light Grid Accent 1'
    for i, row_data in enumerate(rows_data):
        for j, val in enumerate(row_data):
            cell = table.rows[i].cells[j]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            if header and i == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                set_cell_shading(cell, '0B5C5C')
    if col_widths:
        for row in table.rows:
            for idx, w in enumerate(col_widths):
                row.cells[idx].width = Cm(w)
    return table


def add_hr(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '888888')
    pBdr.append(bottom)
    pPr.append(pBdr)


def setup_page(doc, margem=Cm(2.0)):
    for section in doc.sections:
        section.left_margin = margem
        section.right_margin = margem
        section.top_margin = margem
        section.bottom_margin = margem


def cabecalho_documento(doc, titulo, subtitulo):
    p = doc.add_paragraph()
    r = p.add_run(titulo)
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = COR_PRIMARIA
    p = doc.add_paragraph()
    r = p.add_run(subtitulo)
    r.font.size = Pt(10)
    r.italic = True
    r.font.color.rgb = COR_CINZA
    add_hr(doc)


# ============================================================
# HELPERS PPTX
# ============================================================
def new_presentation():
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    return prs


def add_slide(prs):
    BLANK = prs.slide_layouts[6]
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = PPalette.LBG
    bg.line.fill.background()
    return s


def add_text(s, text, left, top, width, height, size=18, bold=False, color=None,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    tb = s.shapes.add_textbox(PInches(left), PInches(top), PInches(width), PInches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size = PPt(size)
        r.font.bold = bold
        r.font.italic = italic
        if color is not None:
            r.font.color.rgb = color
    return tb


def add_bar(s, top=0.6, color=None):
    c = color or PPalette.P
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PInches(0.7), PInches(top),
                             PInches(11.9), PInches(0.06))
    bar.fill.solid()
    bar.fill.fore_color.rgb = c
    bar.line.fill.background()


def add_footer(s, text):
    add_text(s, text, 0.7, 7.05, 11.9, 0.3, size=10, color=PPalette.G, italic=True)


def add_title_block(s, kicker, title, subtitle=None):
    add_text(s, kicker, 0.7, 0.4, 11.9, 0.35, size=13, bold=True, color=PPalette.D)
    add_text(s, title, 0.7, 0.75, 11.9, 0.9, size=34, bold=True, color=PPalette.P)
    add_bar(s, 1.7)
    if subtitle:
        add_text(s, subtitle, 0.7, 1.85, 11.9, 0.6, size=17, color=PPalette.G, italic=True)


def add_slide_cover(prs, kicker, big_line1, big_line2, subtitle, footer_line):
    s = add_slide(prs)
    hero = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, PInches(4.8))
    hero.fill.solid()
    hero.fill.fore_color.rgb = PPalette.P
    hero.line.fill.background()
    add_text(s, kicker, 0.8, 1.0, 12, 0.6, size=18, bold=True, color=PPalette.D)
    add_text(s, big_line1, 0.8, 1.6, 12, 1.1, size=64, bold=True, color=PPalette.W)
    add_text(s, big_line2, 0.8, 2.6, 12, 1.1, size=64, bold=True, color=PPalette.W)
    add_text(s, subtitle, 0.8, 3.8, 12, 0.6, size=22, color=PPalette.W, italic=True)
    add_text(s, footer_line, 0.8, 6.6, 12, 0.5, size=14, color=PPalette.G)
    return s


def add_table_slide(s, rows, top=2.8, left=0.7, width=11.9, height=3.8, col_widths=None, size=14):
    n_cols = len(rows[0])
    tbl_shape = s.shapes.add_table(len(rows), n_cols, PInches(left), PInches(top),
                                    PInches(width), PInches(height))
    tbl = tbl_shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            tbl.columns[i].width = PInches(w)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = str(val)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = PPt(size)
                    run.font.bold = (i == 0)
                    run.font.color.rgb = PPalette.W if i == 0 else PPalette.K
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = PPalette.P
    return tbl
