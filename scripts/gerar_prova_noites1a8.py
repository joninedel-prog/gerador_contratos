"""Prova de Redes — conteúdo das Noites 1 a 8. 15 questões objetivas, nível fácil.

Gera:
  curso/Prova - Redes - Noites 1 a 8.docx
  curso/Gabarito - Redes - Prova Noites 1 a 8.docx
"""
import os

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(BASE)
BRASAO = os.path.join(BASE, "assets", "brasao_sc.jpeg")

# (noite, enunciado, [alternativas], índice da correta, explicação curta)
QUESTOES = [
    (1, "O que é o Wi-Fi?",
     ["A própria internet.",
      "O caminho sem fio entre o aparelho e o roteador.",
      "O plano de dados que a operadora vende.",
      "Um programa que deixa a internet mais rápida."],
     1, "Wi-Fi não é internet: é só o caminho até o roteador."),
    (1, "O computador da recepção imprime na impressora da secretaria, mas a internet da escola caiu. "
        "Esses dois aparelhos continuam formando uma rede?",
     ["Sim. Rede é aparelho ligado trocando informação, não precisa de internet.",
      "Não. Sem internet não existe rede.",
      "Só se a impressora tiver Wi-Fi.",
      "Só se o computador tiver dados móveis."],
     0, "Rede vem primeiro, internet vem depois."),
    (2, "Dos equipamentos da instalação (ONU, modem, roteador, switch, cabo), qual deles cria a rede "
        "da casa e distribui a conexão para os aparelhos?",
     ["O cabo de rede.", "A ONU.", "O switch.", "O roteador."],
     3, "Cinco equipamentos, e só um cria rede: o roteador."),
    (2, "Em qual porta do roteador vai o cabo que vem do provedor (da ONU ou do modem)?",
     ["Em qualquer porta, tanto faz.",
      "Na porta WAN (ou Internet), que geralmente tem outra cor.",
      "Na porta LAN número 1.",
      "Na entrada da fonte de energia."],
     1, "Cabo do provedor numa porta LAN é o erro número 1: conecta, mas nada abre."),
    (2, "Um roteador doméstico tem duas senhas diferentes. Quais são?",
     ["A senha do Wi-Fi e a senha de administrador (pra entrar no painel).",
      "A senha do banco e a senha do e-mail.",
      "A senha do provedor e a senha do celular.",
      "Ele tem só uma senha."],
     0, "Wi-Fi pra usar a internet; administrador pra mexer no roteador."),
    (3, "Quais são os três meios por onde o sinal de rede pode viajar?",
     ["Tomada, antena e satélite.",
      "Wi-Fi, Bluetooth e 4G.",
      "Roteador, switch e modem.",
      "Cabo, fibra e ondas de rádio."],
     3, "Só existem três meios: cabo, fibra e onda. Não existe um quarto."),
    (3, "Qual é o comprimento máximo recomendado para um cabo de rede (par trançado) sem repetidor?",
     ["10 metros.", "1 quilômetro.", "100 metros.", "Não tem limite."],
     2, "Passou de 100 m, a escolha é fibra."),
    (4, "Qual é o nome do conector que vai na ponta do cabo de rede?",
     ["RJ11.", "USB.", "HDMI.", "RJ45."],
     3, "RJ45, colocado com o alicate crimpador. O RJ11 é o do telefone."),
    (4, "No padrão T568B, que a turma usa na crimpagem, qual cor vai no pino 1?",
     ["Branco do laranja (laranja/branco).", "Azul.", "Marrom.", "Verde."],
     0, "Começa laranja, azul no meio, marrom no fim."),
    (5, "No testador de cabo, um dos LEDs não acende. O que isso quer dizer?",
     ["O cabo está perfeito.",
      "A bateria do testador está cheia.",
      "Um dos fios não está chegando (pino aberto). O cabo tem defeito.",
      "O cabo é de fibra."],
     2, "LED faltando = fio que não encosta. Cabo reprovado não se instala."),
    (5, "Pra achar o defeito sem gastar peça à toa, o técnico deve:",
     ["Trocar tudo de uma vez: cabo, roteador e computador.",
      "Trocar uma coisa por vez e testar de novo.",
      "Desligar e ligar até funcionar.",
      "Comprar um roteador novo."],
     1, "Olhar, testar, isolar (uma coisa por vez) e resolver."),
    (6, "Sobre os endereços de um aparelho na rede, qual frase está certa?",
     ["O MAC muda toda vez que o aparelho liga.",
      "O IP vem gravado de fábrica.",
      "O MAC é como o CPF (vem de fábrica) e o IP é como o número do apartamento (quem dá é o roteador).",
      "IP e MAC são a mesma coisa."],
     2, "MAC é o CPF. IP é o número do apartamento."),
    (7, "O que faz o DNS?",
     ["Funciona como uma agenda: transforma o nome do site no número (IP) do site.",
      "Deixa o Wi-Fi mais forte.",
      "Bloqueia vírus.",
      "Testa o cabo de rede."],
     0, "O DNS é a agenda da internet."),
    (7, "Um site mostra o cadeado (https) na barra de endereço. Isso garante, sozinho, que o site é "
        "mesmo do seu banco?",
     ["Sim. Com cadeado é sempre seguro.",
      "Não. O golpista também pode ter cadeado: é preciso conferir o dono do endereço.",
      "Sim, desde que o site seja bonito.",
      "Só se for aberto pelo celular."],
     1, "Cadeado mais dono certo. Um sem o outro não vale."),
    (8, "Ao configurar o Wi-Fi do roteador do cliente, qual segurança deve ser escolhida?",
     ["Rede aberta, sem senha.",
      "WEP.",
      "WPA2 (ou WPA2/WPA3).",
      "Tanto faz, todas são iguais."],
     2, "Nunca aberta nem WEP. WPA2 é a tranca."),
]

LETRAS = "abcd"


def _run(p, txt, bold=False, size=11, italic=False):
    r = p.add_run(txt)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = "Arial"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    return r


def sem_bordas(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{b}")
        e.set(qn("w:val"), "nil")
        borders.append(e)
    tbl_pr.append(borders)


def novo_doc():
    doc = Document()
    for s in doc.sections:
        s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.left_margin = s.right_margin = Cm(2.0)
        s.top_margin = s.bottom_margin = Cm(1.6)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    return doc


def cabecalho(doc, titulo):
    t = doc.add_table(rows=1, cols=2)
    sem_bordas(t)
    for cell, w in zip(t.rows[0].cells, (Cm(3.2), Cm(13.8))):
        cell.width = w
    c_logo, c_txt = t.rows[0].cells
    c_logo.paragraphs[0].add_run().add_picture(BRASAO, width=Cm(2.2))
    linhas = [
        ("ESTADO DE SANTA CATARINA", True),
        ("Secretaria de Estado da Educação", False),
        ("Centro de Educação de Jovens e Adultos – CEJA Itapiranga – SC", False),
        ("Curso FIC Montador e Reparador de Computadores · Redes de Computadores", False),
        ("Professor: Joni César Eidt", False),
    ]
    for i, (txt, b) in enumerate(linhas):
        p = c_txt.paragraphs[0] if i == 0 else c_txt.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        _run(p, txt, bold=b, size=10)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(6)
    _run(p, titulo, bold=True, size=14)


def campos_aluno(doc):
    for txt in ("Nome: ______________________________________________________________",
                "Data: ____/____/2026                                   Nota: ________"):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        _run(p, txt)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    _run(p, "Marque com X a alternativa correta. Cada questão tem só uma resposta certa. "
            "Pode usar lápis ou caneta. Leia com calma: as questões são sobre o que fizemos em aula.",
         italic=True, size=10)


def build_prova():
    doc = novo_doc()
    cabecalho(doc, "AVALIAÇÃO — REDES DE COMPUTADORES (Noites 1 a 8)")
    campos_aluno(doc)

    for n, (_, enun, alts, _, _) in enumerate(QUESTOES, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        _run(p, f"{n}. ", bold=True)
        _run(p, enun)
        for i, alt in enumerate(alts):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.8)
            p.paragraph_format.space_after = Pt(0)
            if i < len(alts) - 1:
                p.paragraph_format.keep_with_next = True
            _run(p, f"(   ) {LETRAS[i]}) {alt}")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    _run(p, "Bom trabalho!", italic=True, size=10)

    path = os.path.join(RAIZ, "Prova - Redes - Noites 1 a 8.docx")
    doc.save(path)
    print("OK:", path)


def build_gabarito():
    doc = novo_doc()
    cabecalho(doc, "GABARITO — AVALIAÇÃO DE REDES (Noites 1 a 8)")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    _run(p, "Uso do professor. 15 questões. Sugestão: cada acerto vale a nota total dividida por 15 "
            "(ex.: nota 10 → 0,67 por questão).", italic=True, size=10)

    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    for c, txt in zip(t.rows[0].cells, ("Nº", "Resposta", "Noite", "Por quê")):
        c.text = ""
        _run(c.paragraphs[0], txt, bold=True, size=10)
    for n, (noite, _, alts, certa, expl) in enumerate(QUESTOES, 1):
        row = t.add_row().cells
        vals = (str(n), f"{LETRAS[certa]}) {alts[certa]}", str(noite), expl)
        for c, v in zip(row, vals):
            c.text = ""
            _run(c.paragraphs[0], v, size=10, bold=(c is row[0]))
    for row in t.rows:
        for c, w in zip(row.cells, (Cm(1.0), Cm(7.2), Cm(1.3), Cm(7.5))):
            c.width = w

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    _run(p, "Sequência pra correção rápida: ", bold=True, size=11)
    _run(p, "  ".join(f"{n}-{LETRAS[q[3]]}" for n, q in enumerate(QUESTOES, 1)), size=11)

    path = os.path.join(RAIZ, "Gabarito - Redes - Prova Noites 1 a 8.docx")
    doc.save(path)
    print("OK:", path)


if __name__ == "__main__":
    build_prova()
    build_gabarito()
