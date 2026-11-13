"""Gera o Plano Mensal 4 de Informática (curso de Redes) — 16/11/2026 a 15/12/2026.

Mesmo layout do modelo SED/CEJA do Plano Mensal 3. Conteúdo: servidores Linux
(Ubuntu Server 24.04) nas máquinas antigas preparadas pelo professor, ligados à
rede montada com os roteadores configurados nas noites anteriores. Encontros às
sextas: 27/11, 04/12 e 11/12 (20/11 é feriado — Dia da Consciência Negra).
"""
import os

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTDIR = "/home/nedel/curso"
BRASAO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "brasao_sc.jpeg")
FONTE = "Arial"
ARQUIVO = "Plano Mensal 4 - Informatica - NOVEMBRO-DEZEMBRO 2026.docx"

CAMPOS = [
    ("COMPONENTE CURRICULAR:", "Informática — Curso FIC em Montador e Reparador de Computadores"),
    ("PROFESSOR:", "Joni César Eidt"),
    ("TEMA GERADOR:", "Projeto de Vida: Construindo o Futuro"),
    ("RECORTE DO COMPONENTE:", "Servir — o computador que trabalha para a rede: servidores com Linux"),
    ("TURMA:", "Ensino Médio – Noturno (Multisseriado)"),
    ("QUESTÃO MOTRIZ:", "Um escritório pequeno quer guardar os arquivos num lugar só, acessar de qualquer "
                        "computador e ter uma página interna para os funcionários. Como transformar um "
                        "computador antigo no servidor que resolve isso?"),
]

OBJETIVOS = [
    "Diferenciar o computador que usa a rede (cliente) do computador que oferece serviços a ela (servidor), "
    "reconhecendo exemplos no cotidiano;",
    "Reconhecer o sistema operacional Linux e a distribuição Ubuntu Server como opção gratuita para "
    "reaproveitar computadores antigos;",
    "Operar o servidor pelo terminal: entrar no sistema, navegar entre pastas, criar, copiar, mover e "
    "apagar arquivos e consultar a ajuda dos comandos;",
    "Criar usuários e grupos e aplicar permissões, entendendo quem pode ler, alterar ou executar cada arquivo;",
    "Identificar e fixar o endereço de rede do servidor, relacionando-o à reserva de endereço feita no "
    "roteador, e testar a comunicação entre as máquinas;",
    "Instalar e atualizar programas pelo gerenciador de pacotes;",
    "Acessar e administrar o servidor a distância por conexão segura (SSH);",
    "Configurar uma pasta compartilhada na rede, acessível a partir de computadores com Windows e de celulares;",
    "Publicar uma página interna em servidor web e acessá-la pelos aparelhos da rede local;",
    "Documentar o servidor entregue em ficha técnica e reconhecer a administração de pequenos servidores "
    "como frente de prestação de serviço.",
]

INTERACAO_INTRO = ("O componente segue integrado ao tema gerador pelo viés do trabalho como projeto de "
                   "vida, e dialoga com as demais áreas nos seguintes pontos:")

INTERACAO = [
    ("LINGUAGENS — ", "leitura de comandos e de mensagens do sistema em inglês técnico, escrita de "
                      "instruções passo a passo para outra pessoa executar e produção da página interna "
                      "e da ficha técnica do servidor (EM13LGG701 e EM13LGG703)."),
    ("MATEMÁTICA — ", "unidades de armazenamento, cálculo de espaço em disco ocupado e livre, e leitura "
                      "do sistema de permissões por combinação de números (EM13MAT101 e EM13MAT405)."),
    ("CIÊNCIAS HUMANAS — ", "software livre e trabalho colaborativo, reaproveitamento de equipamentos e "
                            "lixo eletrônico, e responsabilidade de quem administra dados de outras "
                            "pessoas (EM13CHS106 e EM13CHS304)."),
    ("CIÊNCIAS DA NATUREZA — ", "consumo de energia de um computador ligado o tempo todo e descarte "
                                "adequado de componentes eletrônicos (EM13CNT309)."),
]

HABILIDADE_IF = ("(EMIFFTP05) Selecionar e mobilizar intencionalmente recursos criativos para resolver "
                 "problemas reais relacionados à produtividade, à colaboração e/ou à comunicação.")

CONTEUDOS = [
    "Cliente e servidor: o que é um serviço de rede; exemplos do cotidiano; por que um computador antigo "
    "serve de servidor; Linux, software livre e Ubuntu Server 24.04.",
    "Terminal do Linux: entrada no sistema; estrutura de pastas; comandos de navegação e de manipulação "
    "de arquivos (pwd, ls, cd, mkdir, cp, mv, rm, cat, nano); ajuda dos comandos; desligar e reiniciar "
    "com segurança.",
    "Usuários e permissões: administrador (sudo); criação de usuários e grupos; dono do arquivo; "
    "permissões de leitura, escrita e execução.",
    "Servidor na rede: endereço de rede do servidor (ip a); endereço fixo e reserva por MAC no roteador; "
    "teste de comunicação (ping); instalação e atualização de programas (apt).",
    "Acesso remoto: conexão segura (SSH) a partir do Windows e do celular; senha forte para o servidor.",
    "Serviços: pasta compartilhada na rede (Samba) com usuários e permissões; servidor web (Nginx) com "
    "página interna; verificação do funcionamento a partir de outros aparelhos.",
    "Registro do serviço: ficha técnica do servidor (endereço, usuários, serviços ativos e como acessar).",
]

METODOLOGIA = [
    "Predominância de atividade prática (mais de sessenta por cento do encontro) e exposição dialogada "
    "em blocos curtos, sempre seguidos de prática no terminal;",
    "A situação-problema antecede o conteúdo: cada encontro parte de uma demanda de cliente — guardar "
    "arquivos num lugar só, acessar sem ir até a máquina, ter uma página interna;",
    "Trabalho em equipes, uma por servidor, com rodízio de quem digita e de quem confere, usando as "
    "máquinas antigas já preparadas com Ubuntu Server e ligadas aos roteadores configurados na unidade "
    "anterior;",
    "Cartões de comandos impressos e passo a passo por serviço, com espaço para anotar o resultado obtido;",
    "Verificação objetiva permanente: o comando funciona ou devolve erro, a pasta aparece ou não aparece "
    "no outro computador, a página abre ou não abre no celular;",
    "Leitura guiada das mensagens de erro como parte do método de diagnóstico;",
    "Conferência cruzada entre equipes: uma equipe acessa a pasta e a página do servidor da outra apenas "
    "com a ficha técnica entregue;",
    "Retomada nos minutos iniciais de cada encontro e atendimento individual nos minutos finais;",
    "Encontros previstos às sextas-feiras 27/11, 04/12 e 11/12 (20/11 é feriado), com o último encontro "
    "dedicado à entrega e apresentação do servidor da equipe.",
]

RECURSOS = [
    "Laboratório de informática e projetor;",
    "Computadores antigos reaproveitados com Ubuntu Server 24.04 instalado pelo professor;",
    "Roteadores configurados nas unidades anteriores, cabos de rede e régua de tomadas;",
    "Computadores do laboratório com Windows e celulares dos estudantes como clientes da rede, com "
    "aplicativo de terminal SSH;",
    "Material impresso produzido pelo professor: roteiros, folhas do aluno, cartões de comandos, passo "
    "a passo de cada serviço e modelo de ficha técnica do servidor.",
]

AVALIACAO = [
    "Portfólio profissional: ficha técnica do servidor entregue pela equipe;",
    "Produto verificável: servidor em funcionamento com acesso remoto, pasta compartilhada com "
    "permissões corretas e página interna, validado por outra equipe somente com a ficha técnica;",
    "Atividades práticas de aula, com ficha de acompanhamento por encontro (comandos executados, "
    "usuários criados, testes de acesso);",
    "Avaliação prática integradora no encontro de 11/12: entrega e apresentação do servidor da equipe "
    "como se fosse para o cliente;",
    "Participação e frequência, com atenção permanente ao risco de evasão;",
    "Recomposição imediata das aprendizagens não atingidas, com repetição do procedimento até a "
    "execução correta.",
]


def _run(p, texto, bold=False):
    r = p.add_run(texto)
    r.bold = bold
    r.font.name = FONTE
    r.font.size = Pt(12)
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONTE)
    return r


def _para(doc, space_before=0, space_after=0, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.0
    p.alignment = align
    return p


def campo(doc, rotulo, valor):
    p = _para(doc, space_after=10)
    _run(p, rotulo, bold=True)
    _run(p, "  " + valor)


def titulo_secao(doc, texto, space_before=10):
    p = _para(doc, space_before=space_before)
    p.paragraph_format.keep_with_next = True
    _run(p, texto, bold=True)


def bullet(doc, texto, negrito=None):
    p = _para(doc)
    pf = p.paragraph_format
    pf.left_indent = Cm(1.05)
    pf.first_line_indent = Cm(-0.55)
    pf.tab_stops.add_tab_stop(Cm(1.05))
    _run(p, "●\t")
    if negrito:
        _run(p, negrito, bold=True)
    _run(p, texto)


def sem_bordas(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for lado in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{lado}')
        b.set(qn('w:val'), 'nil')
        borders.append(b)
    tbl_pr.append(borders)


def cabecalho(doc):
    t = doc.add_table(rows=1, cols=2)
    sem_bordas(t)
    larguras = (Cm(4.3), Cm(12.6))
    for cell, w in zip(t.rows[0].cells, larguras):
        cell.width = w
    c_logo, c_txt = t.rows[0].cells
    p = c_logo.paragraphs[0]
    p.paragraph_format.left_indent = Cm(1.25)
    p.add_run().add_picture(BRASAO, width=Cm(2.5))
    linhas = [
        ("ESTADO DE SANTA CATARINA", True),
        ("Secretaria de Estado da Educação", False),
        ("Centro de Educação de Jovens e Adultos – CEJA", False),
        ("Itapiranga – SC Fone: 3472 5032", False),
    ]
    for i, (txt, b) in enumerate(linhas):
        p = c_txt.paragraphs[0] if i == 0 else c_txt.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(10) if i == 0 else Pt(0)
        _run(p, txt, bold=b)


def gerar():
    doc = Document()
    for s in doc.sections:
        s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.left_margin = Cm(2.2)
        s.right_margin = Cm(2.1)
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
    normal = doc.styles['Normal']
    normal.font.name = FONTE
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONTE)

    cabecalho(doc)
    _para(doc, space_after=10)

    p = _para(doc, space_after=10)
    _run(p, "PLANO MENSAL", bold=True)
    for rotulo, valor in CAMPOS:
        campo(doc, rotulo, valor)

    titulo_secao(doc, "OBJETIVOS E METAS:", space_before=0)
    for t in OBJETIVOS:
        bullet(doc, t)

    titulo_secao(doc, "INTERAÇÃO COM AS ÁREAS DO CONHECIMENTO:")
    p = _para(doc)
    _run(p, INTERACAO_INTRO)
    for area, texto in INTERACAO:
        bullet(doc, texto, negrito=area)

    titulo_secao(doc, "HABILIDADE DO ITINERÁRIO FORMATIVO:")
    p = _para(doc)
    _run(p, HABILIDADE_IF)

    titulo_secao(doc, "CONTEÚDOS DO PERÍODO:")
    for t in CONTEUDOS:
        bullet(doc, t)

    titulo_secao(doc, "CAMINHO METODOLÓGICO:")
    for t in METODOLOGIA:
        bullet(doc, t)

    titulo_secao(doc, "RECURSOS:")
    for t in RECURSOS:
        bullet(doc, t)

    _para(doc)
    titulo_secao(doc, "AVALIAÇÃO E MONITORAMENTO:")
    for t in AVALIACAO:
        bullet(doc, t)

    path = os.path.join(OUTDIR, ARQUIVO)
    doc.save(path)
    print(path)


if __name__ == "__main__":
    gerar()
