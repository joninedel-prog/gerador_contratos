"""Gera os 4 arquivos da Noite 1 do curso de Redes de Computadores.

Saída (na pasta /home/nedel/curso/):
- Roteiro - Redes 1 - Rede nao e internet.docx
- Aula - Redes 1 - Rede nao e internet.pptx
- Folha do Aluno - Redes 1 - Separar as palavras.docx
- Resumo do Aluno - Redes 1 - Rede nao e internet.docx
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.shapes import MSO_SHAPE

OUTDIR = "/home/nedel/curso"

# ============================================================
# PALETA E ESTILO (para lembrar o visual da Folha do Aluno atual)
# ============================================================
COR_PRIMARIA = RGBColor(0x0B, 0x5C, 0x5C)   # verde-petróleo (tom do curso)
COR_ALERTA   = RGBColor(0xC5, 0x36, 0x2B)   # vermelho de alerta
COR_DESTAQUE = RGBColor(0xE8, 0x8B, 0x1A)   # laranja (como o WAN)
COR_CINZA    = RGBColor(0x55, 0x55, 0x55)


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


def add_run(paragraph, text, bold=False, color=None, size=None, italic=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if size is not None:
        run.font.size = Pt(size)
    return run


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


# ============================================================
# ROTEIRO — guia do professor
# ============================================================
def build_roteiro():
    doc = Document()

    for section in doc.sections:
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    # Cabeçalho
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = title.add_run("REDE NÃO É INTERNET")
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = COR_PRIMARIA

    sub = doc.add_paragraph()
    r = sub.add_run("Roteiro do professor · Redes de Computadores · Noite 1 · Curso FIC · CEJA Itapiranga – SC")
    r.font.size = Pt(10)
    r.font.color.rgb = COR_CINZA
    r.italic = True
    add_hr(doc)

    # A IDEIA DESTA NOITE
    add_heading(doc, "A IDEIA DESTA NOITE")
    add_para(doc, "Esta é a primeira noite. É a noite em que a turma vai parar de misturar palavras que todo mundo mistura.")
    add_para(doc, "\"Estou sem internet\" quase nunca é \"estou sem internet\". É sem Wi-Fi, é sem dados móveis, é rede sem sinal do provedor, é aparelho no modo avião. Cada uma dessas frases pede uma resposta diferente do técnico.")
    add_para(doc, "Se o aluno sai daqui separando quatro palavras — rede, internet, Wi-Fi e dados móveis — a noite valeu. E o resto do curso fica muito mais fácil, porque toda aula depois desta parte deste chão.")

    add_heading(doc, "O que se resolve nesta noite", level=2)
    add_para(doc, "Quatro palavras, uma para cada situação. E o ícone do celular, que já conta a maior parte da história antes de qualquer chamado começar.")

    add_heading(doc, "Esta é a primeira de dez noites", level=2)
    add_para(doc, "Hoje não tem equipamento na mão, não tem cabo pra ligar. Hoje é conversa dirigida. Se você tiver um roteador na mesa, deixa lá desligado — vai ganhar peso na Noite 2.")

    add_hr(doc)

    # ANTES DA AULA
    add_heading(doc, "ANTES DA AULA")

    add_heading(doc, "Material da mesa", level=2)
    add_bullet(doc, "Um roteador (opcional, desligado, só como cenário)")
    add_bullet(doc, "Um celular do professor, com Wi-Fi que dá pra ligar e desligar")
    add_bullet(doc, "Uma régua de tomadas (se for usar o roteador)")
    add_bullet(doc, "Quadro branco + canetão de duas cores")

    add_heading(doc, "Prepare antes", level=2)
    add_bullet(doc, "Confira que o Wi-Fi da sala está funcionando (ou o hotspot do seu celular)")
    add_bullet(doc, "Confira que dá pra desligar o Wi-Fi do seu celular e ainda ficar com dados móveis")
    add_bullet(doc, "Se puder, chegue com o roteador com uma etiqueta visível (SSID, senha) — vai usar na Noite 2")

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "A folha do aluno, uma por pessoa")
    add_bullet(doc, "O resumo do aluno, uma por pessoa (para levar para casa)")

    add_hr(doc)

    # QUADRO DE TEMPO
    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Abertura e apresentação", "20 min", "Quem é você, o curso, e a pergunta-provocação"],
        ["2. Rede é conexão entre coisas", "30 min", "Conceito raiz, sem internet ainda"],
        ["3. Internet é a rede das redes", "25 min", "O que muda quando a rede fica grande"],
        ["INTERVALO", "15 min", ""],
        ["4. Wi-Fi, cabo e dados móveis", "35 min", "Os três caminhos até a rede"],
        ["5. O que os ícones do celular contam", "30 min", "Reconhecimento visual"],
        ["6. Prática: separar as palavras", "40 min", "Cenários reais em equipe"],
        ["7. Fechamento e tarefa", "15 min", "Consolidação"],
    ], col_widths=[6, 2, 8.5])
    add_para(doc, "Se atrasar: corte o Bloco 1 para 10 minutos. Não corte o Bloco 6 — a prática é o que fixa.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — ABERTURA · 20 MIN")

    add_heading(doc, "Você, primeiro", level=2)
    add_para(doc, "Dois minutos falando quem é você e o que faz de dia. Não muito. A turma quer ouvir isso mas não quer aula sobre isso.")

    add_heading(doc, "O curso, em três frases", level=2)
    add_bullet(doc, "Dez noites. Uma por semana.")
    add_bullet(doc, "No fim, o aluno consegue instalar uma rede numa casa, entender por que travou quando trava, e falar com o cliente sem inventar.")
    add_bullet(doc, "Não é aula de teoria. É aula de trabalho de campo.")

    add_heading(doc, "A pergunta que abre o curso", level=2)
    add_para(doc, "Escreva no quadro:")
    add_para(doc, "\"Você diz que está sem internet. O que é que você está sem, exatamente?\"", bold=True)
    add_para(doc, "Deixe a pergunta parada por uns segundos. Alguém vai responder \"sem Wi-Fi\". Outra pessoa vai responder \"sem sinal\". Outra vai dizer \"sem crédito\". Todas essas respostas são certas — e diferentes. Escreva cada resposta no canto do quadro.")
    add_para(doc, "\"Repare que a gente já disse a palavra 'internet' de quatro jeitos diferentes, e ninguém tá errado. Cada um resolve num lugar diferente. Hoje a gente vai colocar nome nessas quatro coisas.\"", italic=True)

    add_heading(doc, "O acordo da noite", level=2)
    add_para(doc, "\"Hoje ninguém precisa saber nada de rede. Se você acha que já sabe, ótimo: presta atenção porque no fim do curso vai instalar rede na casa dos outros e vai precisar explicar. Se você acha que não sabe, melhor ainda: essa aula é seu ponto de partida.\"", italic=True)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — REDE É CONEXÃO ENTRE COISAS · 30 MIN")

    add_heading(doc, "A definição enxuta", level=2)
    add_para(doc, "Escreva no quadro:")
    add_para(doc, "REDE = duas ou mais coisas ligadas trocando informação", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Só isso. Duas caixas registradoras do mercado ligadas ao mesmo computador do caixa: é uma rede. Dois celulares ligados por Bluetooth trocando foto: é uma rede. A tomada da parede ligada à geladeira? Não. Porque a tomada não troca informação com a geladeira. Só entrega energia.\"", italic=True)

    add_heading(doc, "Três exemplos de rede sem internet — 15 min", level=2)
    add_para(doc, "Um por vez, com a turma junto.")

    add_heading(doc, "Exemplo 1: a rede da secretaria", level=2)
    add_para(doc, "\"Aqui na escola tem uma rede que liga o computador da secretaria à impressora. Você manda imprimir daqui, sai lá. Isso não usa internet. Se a internet da escola cair, ainda dá pra imprimir.\"", italic=True)

    add_heading(doc, "Exemplo 2: a rede do mercado", level=2)
    add_para(doc, "\"No mercado, quando o cartão passa, o valor vai da máquina do caixa até o computador do gerente. Isso é rede. Depois disso, sim, o computador do gerente conversa com o banco pela internet. Mas o pedaço de dentro do mercado é rede sem internet.\"", italic=True)

    add_heading(doc, "Exemplo 3: o Bluetooth", level=2)
    add_para(doc, "\"Quando você conecta um fone Bluetooth no celular, você fez uma rede de dois. Sem internet nenhuma. Prova: no modo avião, o Bluetooth continua funcionando.\"", italic=True)

    add_heading(doc, "Pergunte pra sala — 10 min", level=2)
    add_para(doc, "\"Alguém aqui já viu uma rede funcionando sem internet? Onde?\"", italic=True)
    add_para(doc, "Deixe as respostas virem. Muita gente vai lembrar de LAN house de jogo, sistema de PDV do trabalho, câmera de segurança que grava direto num HD. Aceite tudo. Escreva no quadro os que aparecerem.")

    add_heading(doc, "A frase-âncora", level=2)
    add_para(doc, "\"Rede é conexão entre coisas. Internet é uma rede específica, muito grande. A rede vem primeiro. A internet vem depois.\"", bold=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — INTERNET É A REDE DAS REDES · 25 MIN")

    add_heading(doc, "O nome já conta", level=2)
    add_para(doc, "\"Inter + net. Entre redes. A internet é o que acontece quando várias redes de vários lugares se conectam entre si.\"", italic=True)

    add_heading(doc, "Desenhe no quadro — 10 min", level=2)
    add_para(doc, "Três círculos separados no quadro. Cada um representa uma rede pequena (a de uma casa, a de um mercado, a de uma escola).")
    add_para(doc, "\"Cada círculo desses funciona sozinho. A internet aparece quando alguém conecta esses três círculos entre si. Aí a rede da sua casa consegue conversar com a rede do YouTube, que fica lá em algum lugar do mundo.\"", italic=True)
    add_para(doc, "Traça linhas conectando os círculos passando por um ponto central. Escreva \"provedor\" no ponto central.")
    add_para(doc, "\"O provedor é quem conecta o círculo da sua casa ao mundo. Sem ele, sua casa continua tendo rede — só não fala com ninguém fora.\"", italic=True)

    add_heading(doc, "O que muda quando é internet — 10 min", level=2)
    add_table(doc, [
        ["Coisa", "Sem internet", "Com internet"],
        ["Onde a informação vai", "Dentro da casa ou local", "Pode ir pro outro lado do mundo"],
        ["Quem cuida da estrada", "Você (e o provedor da rede local)", "O provedor da internet"],
        ["Quando cai", "Você conserta", "Pode ser você OU o provedor"],
        ["Custa mensalidade?", "Não (a rede em si)", "Sim (o acesso à internet)"],
    ], col_widths=[5, 5.5, 6])

    add_heading(doc, "A pegadinha do vocabulário — 5 min", level=2)
    add_para(doc, "\"O cliente diz: 'não tenho internet'. Você pergunta: 'o Wi-Fi conecta?'. Ele diz: 'conecta'. Aí você sabe: a rede da casa dele está funcionando. O problema é da internet — que é o pedaço que sai da casa dele. Isso é meia batalha ganha.\"", italic=True)

    add_hr(doc)
    add_heading(doc, "INTERVALO · 15 MIN", color=COR_CINZA)
    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — WI-FI, CABO E DADOS MÓVEIS · 35 MIN")

    add_heading(doc, "O ponto que precisa ficar claro", level=2)
    add_para(doc, "\"Wi-Fi não é internet. Cabo de rede não é internet. Dados móveis é o mais próximo, mas mesmo assim não é a mesma coisa. Todos eles são caminhos até uma rede. Se a rede não tiver internet, o caminho não muda nada.\"", italic=True)

    add_heading(doc, "Os três caminhos, um por um — 20 min", level=2)

    add_heading(doc, "Wi-Fi — 7 min", level=2)
    add_para(doc, "\"Wi-Fi é jeito de conectar sem fio até um roteador. Só isso. É o pedaço do caminho que não tem fio.\"", italic=True)
    add_bullet(doc, "Wi-Fi liga o celular ao roteador")
    add_bullet(doc, "O roteador é que liga na internet (se estiver ligado)")
    add_para(doc, "\"Se você conecta no Wi-Fi da escola e a escola não tá pagando a internet, o Wi-Fi continua conectando — e nada abre. Aí no celular aparece aquele triângulo com ponto de exclamação. A gente volta nesse ícone.\"", italic=True)

    add_heading(doc, "Cabo de rede — 7 min", level=2)
    add_para(doc, "\"É a mesma coisa que o Wi-Fi, só que com fio. Você espeta o cabo entre o computador e o roteador. Mais estável, mais rápido, sem interferência. É o que o técnico usa quando o Wi-Fi não dá conta.\"", italic=True)
    add_para(doc, "\"O cabo por si só também não é internet. É o caminho. Se o roteador do outro lado não tiver internet chegando, o cabo tá lá, mas nada abre.\"", italic=True)

    add_heading(doc, "Dados móveis — 6 min", level=2)
    add_para(doc, "\"Aqui é diferente. Quando você usa dados móveis, seu celular tá conectado à antena da operadora. E essa antena tá ligada na internet da operadora. Então é o único dos três em que 'ligou = tem internet', porque a operadora vende a internet junto.\"", italic=True)
    add_para(doc, "\"Mas mesmo aqui tem separação. Se você tá em lugar sem sinal, os dados móveis não pegam. Se você acabou o pacote, o sinal tá lá mas a internet foi bloqueada. E se a antena da operadora tiver problema, você tem sinal e não tem internet.\"", italic=True)

    add_heading(doc, "O quadro dos três caminhos — 15 min", level=2)
    add_table(doc, [
        ["Caminho", "O que faz", "Cria internet?", "Depende de quê pra ter internet"],
        ["Wi-Fi", "Liga o aparelho ao roteador da casa", "Não", "Do roteador estar conectado ao provedor"],
        ["Cabo de rede", "Liga o aparelho ao roteador da casa", "Não", "Do roteador estar conectado ao provedor"],
        ["Dados móveis", "Liga o celular à antena da operadora", "Não, mas costuma vir junto", "Da operadora ter sinal e o pacote estar ativo"],
    ], col_widths=[3.2, 4.5, 3, 5.8])
    add_para(doc, "\"Nenhum dos três é internet. Todos os três são caminhos. A internet é o que corre por dentro do caminho — quando corre.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Uma demonstração se der — 5 min", level=2)
    add_para(doc, "Pegue seu celular. Mostre pra turma.")
    add_bullet(doc, "Desligue os dados móveis, mantenha o Wi-Fi ligado no Wi-Fi da sala. Abra o WhatsApp: funciona. Por quê? Porque o roteador tem internet.")
    add_bullet(doc, "Agora desligue o Wi-Fi e ligue os dados móveis. Funciona também. Por quê? Porque a operadora tem internet.")
    add_bullet(doc, "Ligue os dois. Funciona. O celular escolhe o Wi-Fi porque não gasta o pacote.")
    add_bullet(doc, "Ligue o modo avião. Não funciona. Nenhum caminho tá aberto.")
    add_para(doc, "\"O celular tá o tempo todo escolhendo por qual caminho vai. Você só precisa saber que os caminhos são separados.\"", italic=True)

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — O QUE OS ÍCONES DO CELULAR CONTAM · 30 MIN")

    add_heading(doc, "A tese", level=2)
    add_para(doc, "\"Os ícones no canto de cima do celular são um relatório em miniatura. Se você aprender a ler, resolve metade do chamado antes de perguntar qualquer coisa ao cliente.\"", italic=True)

    add_heading(doc, "Os quatro ícones que importam — 20 min", level=2)

    add_heading(doc, "Sinal de barrinhas (celular)", level=2)
    add_para(doc, "\"Quantas barras têm de sinal da antena da operadora. Zero barra: você tá em lugar sem cobertura. Uma barra: sinal fraco. Se aparecer letra do lado (4G, 5G) é porque os dados móveis estão ativados. Se não aparecer letra, o celular pega ligação mas não passa internet pela operadora.\"", italic=True)

    add_heading(doc, "Sinal de Wi-Fi (arco de curvas)", level=2)
    add_para(doc, "\"Diz que o celular tá conectado a algum Wi-Fi. Quantas curvinhas acesas = quão perto do roteador. Se o Wi-Fi tá aceso, o caminho tá aberto até o roteador — só isso. Não diz nada se a internet do lado de lá está funcionando.\"", italic=True)

    add_heading(doc, "O triângulo com exclamação (ou o Wi-Fi com um X)", level=2)
    add_para(doc, "\"Isso aqui é ouro. Significa: 'estou conectado, mas não estou conseguindo abrir nada da internet'. É o caso mais comum de chamado de casa. Rede da casa funcionando, internet do provedor não.\"", italic=True, color=COR_ALERTA)

    add_heading(doc, "Avião", level=2)
    add_para(doc, "\"Modo avião. Desliga tudo: Wi-Fi, dados móveis, Bluetooth (às vezes), sinal de ligação. Se o cliente ligou pra você reclamando de sinal, primeira pergunta: 'tá com o aviãozinho ligado?'\"", italic=True)

    add_heading(doc, "O jogo do ícone — 10 min", level=2)
    add_para(doc, "Fale uma situação, a turma diz que ícone o celular do cliente provavelmente está mostrando.")
    add_bullet(doc, "\"Cliente diz: 'liguei no Wi-Fi de casa mas o WhatsApp não abre'.\" → Wi-Fi aceso + triângulo com exclamação")
    add_bullet(doc, "\"Cliente diz: 'saí de casa e o WhatsApp continua funcionando'.\" → Sem Wi-Fi + barras com 4G/5G")
    add_bullet(doc, "\"Cliente diz: 'meu celular tá sem nada'.\" → Ou modo avião, ou lugar sem sinal, ou desligado. Uma pergunta resolve.")
    add_bullet(doc, "\"Cliente diz: 'tô conectado no Wi-Fi do vizinho e não abre nada'.\" → Wi-Fi aceso + triângulo. Rede do vizinho sem internet, ou senha errada travou.")
    add_para(doc, "\"Se o cliente conseguir mandar foto do ícone antes de você sair de casa, você já sabe o que vai encontrar.\"", italic=True)

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — PRÁTICA: SEPARAR AS PALAVRAS · 40 MIN")

    add_heading(doc, "Divide a turma em equipes de 3 ou 4", level=2)
    add_para(doc, "Cada equipe recebe a folha do aluno. Todos vão preencher individualmente e depois discutir em equipe. Você circula ajudando.")

    add_heading(doc, "Depois de 15 minutos, faça o exercício aberto — 25 min", level=2)
    add_para(doc, "Escolha 6 cenários da folha e discuta com a turma inteira. Não dê resposta direto. Deixe as equipes debaterem. Só corrija se ninguém chegar.")

    add_heading(doc, "Cenários pra ter na manga", level=2)

    add_para(doc, "1. Wi-Fi funciona mas WhatsApp não abre", bold=True)
    add_para(doc, "Rede local: OK. Internet: problema. O caminho até o roteador tá aberto, mas o roteador não tá recebendo internet do provedor. Ou o provedor caiu, ou o cabo do provedor tá solto, ou é problema mais fundo.")

    add_para(doc, "2. A impressora não imprime, mas a internet funciona", bold=True)
    add_para(doc, "Isso é rede local, sem envolver internet. A impressora e o computador não estão se enxergando na rede da casa. Não vale chamar o provedor.")

    add_para(doc, "3. Só o filho consegue conectar no Wi-Fi. Ninguém mais.", bold=True)
    add_para(doc, "Rede local. Ou o roteador só tem endereço pra um aparelho por vez (raro), ou as senhas dos outros aparelhos estão erradas, ou o roteador só reconhece MAC autorizado (menos comum).")

    add_para(doc, "4. Saiu de casa e o Instagram parou no celular dele", bold=True)
    add_para(doc, "Ele saiu do alcance do Wi-Fi de casa e os dados móveis dele estão desligados ou sem pacote. O caminho Wi-Fi acabou e o caminho dados móveis não abriu.")

    add_para(doc, "5. A TV conecta na rede, mas Netflix diz que não tem internet", bold=True)
    add_para(doc, "Rede local: OK (TV enxerga o roteador). Internet: problema. Mesma coisa do caso 1, mas percebido pela TV.")

    add_para(doc, "6. \"A internet tá lenta só de dia\"", bold=True)
    add_para(doc, "Wi-Fi disputado (muitos vizinhos na mesma faixa 2,4 GHz), ou plano de internet limitado com muito consumo, ou rede móvel congestionada. Precisa perguntar mais.")

    add_para(doc, "\"Ninguém aqui hoje precisa resolver o problema. Hoje é só separar: o problema tá na rede ou na internet? Se tá na rede da casa, é você que resolve. Se tá antes da casa, é o provedor. Só isso já muda o rumo do chamado.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — FECHAMENTO E TAREFA · 15 MIN")

    add_heading(doc, "Quiz de saída", level=2)
    add_para(doc, "Escreva as perguntas no quadro. Quem quiser responde em voz alta. Não pressione ninguém.")
    add_bullet(doc, "Rede e internet: qual vem primeiro?")
    add_bullet(doc, "Cite uma rede que funciona sem internet.")
    add_bullet(doc, "Wi-Fi é a mesma coisa que internet?")
    add_bullet(doc, "O celular tá com o ícone de Wi-Fi aceso e o triângulo de exclamação. O que isso quer dizer?")
    add_bullet(doc, "Cliente sem sinal nenhum: primeira pergunta que você faz?")

    add_heading(doc, "A tarefa", level=2)
    add_para(doc, "\"Nesta semana, vocês vão prestar atenção nos ícones do próprio celular. Duas vezes por dia, olhe o canto de cima. Anote o que apareceu: sinal, dados móveis, Wi-Fi, exclamação, avião. Traga anotado, mesmo que seja quatro linhas de caderno. A gente compara na próxima.\"", italic=True)
    add_para(doc, "\"E se em algum momento você ficou 'sem internet' esta semana, tente descobrir por qual caminho. Foi Wi-Fi? Foi dados móveis? Um dos dois, os dois? Anote e traga.\"", italic=True)

    add_heading(doc, "Anuncie a próxima", level=2)
    add_para(doc, "\"Na próxima noite a gente vai atrás dos equipamentos. Vocês vão colocar a mão no roteador, no cabo, e vão saber o nome de cada peça. Se tiver um roteador em casa, tirem uma foto da traseira e do embaixo dele — vamos usar.\"", italic=True)

    add_hr(doc)

    # SE PERGUNTAREM
    add_heading(doc, "SE PERGUNTAREM")

    add_para(doc, "\"O 5G é internet melhor?\"", bold=True)
    add_para(doc, "> \"É rede móvel mais rápida e com menos atraso. Se a rede da operadora ainda não tá boa na sua região, você vai ter 5G no ícone e velocidade ruim. Tecnologia nova não conserta operadora ruim.\"", italic=True)

    add_para(doc, "\"Dá pra usar internet sem provedor?\"", bold=True)
    add_para(doc, "> \"Do jeito comum, não. Ou você paga um provedor de internet fixa, ou paga uma operadora de celular pelos dados móveis. Existem redes comunitárias e conexões via satélite, mas alguém tá pagando por algum caminho. Internet nunca é grátis.\"", italic=True)

    add_para(doc, "\"Meu Wi-Fi tá lento, é internet ruim?\"", bold=True)
    add_para(doc, "> \"Pode ser. Ou pode ser Wi-Fi ruim: distância do roteador, parede grossa, vizinho no mesmo canal. Na Noite 3 a gente ataca isso.\"", italic=True)

    add_para(doc, "\"E se eu tô conectado no Wi-Fi de graça do estabelecimento?\"", bold=True)
    add_para(doc, "> \"Aí o estabelecimento paga a internet, e você usa o caminho Wi-Fi deles. Se cair, o técnico é o deles, não o seu.\"", italic=True)

    add_para(doc, "\"Roteador é a mesma coisa que Wi-Fi?\"", bold=True)
    add_para(doc, "> \"O roteador é o aparelho. O Wi-Fi é o jeito sem fio de conectar nele. Mesmo roteador pode ter Wi-Fi ligado ou desligado, e continua sendo roteador. Isso fica mais claro na Noite 2.\"", italic=True)

    add_hr(doc)

    # DEPOIS DA AULA
    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "Quem participou da conversa; quem ficou calado a noite inteira? (Anote pra puxar mais na Noite 2)")
    add_bullet(doc, "Quantos conseguiram separar Wi-Fi de internet no Bloco 6 sem hesitar")
    add_bullet(doc, "Que analogias funcionaram melhor com essa turma específica — vai reaproveitar")
    add_bullet(doc, "Se alguém trouxe uma dúvida boa que ficou no ar, use ela pra abrir a Noite 2")

    add_hr(doc)

    # DESCRITIVO
    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Introduzi a disciplina de Redes de Computadores, apresentando a estrutura do curso e os objetivos formativos previstos para as dez semanas. Trabalhei com os estudantes a diferenciação conceitual entre rede e internet, evidenciando que rede é o termo genérico para conexão entre dispositivos, enquanto internet é uma rede específica de escala global formada pela interconexão de redes menores. Apresentei exemplos cotidianos de redes que operam sem acesso à internet e discuti as diferenças funcionais entre os caminhos de conexão mais utilizados: rede sem fio Wi-Fi, rede cabeada e rede de dados móveis. Trabalhei ainda o reconhecimento dos indicadores visuais presentes na interface de dispositivos móveis (barras de sinal, símbolo de Wi-Fi, indicador de conectividade limitada e modo avião) como recurso de diagnóstico preliminar. A avaliação ocorreu por meio de participação oral em atividade prática de análise de cenários e por instrumento escrito individual.")

    path = os.path.join(OUTDIR, "Roteiro - Redes 1 - Rede nao e internet.docx")
    doc.save(path)
    return path


# ============================================================
# SLIDES — aula
# ============================================================
def build_slides():
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)

    BLANK = prs.slide_layouts[6]

    P = PRGBColor(0x0B, 0x5C, 0x5C)     # verde-petróleo
    D = PRGBColor(0xE8, 0x8B, 0x1A)     # laranja
    W = PRGBColor(0xFF, 0xFF, 0xFF)
    G = PRGBColor(0x55, 0x55, 0x55)
    K = PRGBColor(0x1A, 0x1A, 0x1A)
    R = PRGBColor(0xC5, 0x36, 0x2B)
    LBG = PRGBColor(0xF4, 0xF7, 0xF7)   # fundo claro

    def add_slide():
        s = prs.slides.add_slide(BLANK)
        # fundo claro
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = LBG
        bg.line.fill.background()
        return s

    def add_text(s, text, left, top, width, height, size=18, bold=False, color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
        tb = s.shapes.add_textbox(PInches(left), PInches(top), PInches(width), PInches(height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = anchor
        # texto pode ser lista de linhas ou str
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
        c = color or P
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PInches(0.7), PInches(top), PInches(11.9), PInches(0.06))
        bar.fill.solid()
        bar.fill.fore_color.rgb = c
        bar.line.fill.background()

    def add_footer(s, text="Redes de Computadores · Noite 1"):
        add_text(s, text, 0.7, 7.05, 11.9, 0.3, size=10, color=G, italic=True)

    def add_title_block(s, kicker, title, subtitle=None):
        add_text(s, kicker, 0.7, 0.4, 11.9, 0.35, size=13, bold=True, color=D)
        add_text(s, title, 0.7, 0.75, 11.9, 0.9, size=34, bold=True, color=P)
        add_bar(s, 1.7)
        if subtitle:
            add_text(s, subtitle, 0.7, 1.85, 11.9, 0.6, size=17, color=G, italic=True)

    # ----- SLIDE 1: capa -----
    s = add_slide()
    hero = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, PInches(4.8))
    hero.fill.solid()
    hero.fill.fore_color.rgb = P
    hero.line.fill.background()
    add_text(s, "NOITE 1", 0.8, 1.0, 12, 0.6, size=18, bold=True, color=D)
    add_text(s, "REDE NÃO É", 0.8, 1.6, 12, 1.1, size=64, bold=True, color=W)
    add_text(s, "INTERNET", 0.8, 2.6, 12, 1.1, size=64, bold=True, color=W)
    add_text(s, "Separar as palavras que o dia a dia mistura", 0.8, 3.8, 12, 0.6, size=22, color=W, italic=True)
    add_text(s, "Redes de Computadores · Curso FIC · CEJA Itapiranga – SC", 0.8, 6.6, 12, 0.5, size=14, color=G)

    # ----- SLIDE 2: A pergunta da noite -----
    s = add_slide()
    add_title_block(s, "A PERGUNTA DA NOITE", "Você diz que está sem internet.", "O que é que você está sem, exatamente?")
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(0.7), PInches(3.0), PInches(11.9), PInches(3.5))
    box.fill.solid()
    box.fill.fore_color.rgb = W
    box.line.color.rgb = P
    box.line.width = Emu(20000)
    add_text(s, [
        "As respostas que aparecem:",
        "",
        "• \"Estou sem Wi-Fi\"",
        "• \"Estou sem sinal\"",
        "• \"Estou sem crédito\"",
        "• \"O aparelho não conecta em nada\"",
    ], 1.1, 3.25, 11.5, 3.2, size=20, color=K)
    add_text(s, "Cada uma dessas frases resolve num lugar diferente.", 0.7, 6.6, 11.9, 0.4, size=14, color=D, italic=True, bold=True)
    add_footer(s)

    # ----- SLIDE 3: O acordo da noite -----
    s = add_slide()
    add_title_block(s, "O ACORDO DA NOITE", "Hoje ninguém precisa saber nada.")
    add_text(s, [
        "Se você já sabe",
        "vai precisar explicar pra cliente depois do curso.",
        "",
        "Se você não sabe",
        "essa aula é o seu ponto de partida.",
        "",
        "Não é aula de teoria. É aula de trabalho de campo.",
    ], 0.7, 2.6, 11.9, 4.4, size=22, color=K)
    add_footer(s)

    # ----- SLIDE 4: A definição -----
    s = add_slide()
    add_title_block(s, "A DEFINIÇÃO", "REDE", "Duas ou mais coisas ligadas trocando informação")
    add_text(s, "Só isso.", 0.7, 3.0, 11.9, 0.6, size=32, bold=True, color=D)
    boxes = [
        ("SÃO REDE", ["Duas caixas do mercado ligadas no computador do caixa", "Dois celulares por Bluetooth trocando foto", "Impressora e computador da secretaria"], P),
        ("NÃO É REDE", ["A tomada e a geladeira (só entrega energia, não troca informação)"], R),
    ]
    x = 0.7
    for label, items, color in boxes:
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(x), PInches(4.0), PInches(6.1), PInches(2.7))
        card.fill.solid()
        card.fill.fore_color.rgb = W
        card.line.color.rgb = color
        card.line.width = Emu(25000)
        add_text(s, label, x + 0.25, 4.15, 5.6, 0.4, size=14, bold=True, color=color)
        add_text(s, ["• " + it for it in items], x + 0.25, 4.55, 5.6, 2.0, size=14, color=K)
        x += 6.3
    add_footer(s)

    # ----- SLIDE 5: Três redes sem internet -----
    s = add_slide()
    add_title_block(s, "TRÊS EXEMPLOS", "Redes que funcionam SEM internet")
    items = [
        ("1", "A impressora da secretaria", "Se a internet da escola cair, ainda dá pra imprimir."),
        ("2", "O caixa do mercado", "O valor vai da máquina do caixa pro computador do gerente. Só depois disso o gerente conversa com o banco pela internet."),
        ("3", "O Bluetooth", "Fone Bluetooth no celular é uma rede de dois. Funciona até no modo avião."),
    ]
    y = 2.6
    for num, tit, desc in items:
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, PInches(0.7), PInches(y), PInches(0.9), PInches(0.9))
        circ.fill.solid()
        circ.fill.fore_color.rgb = P
        circ.line.fill.background()
        add_text(s, num, 0.7, y + 0.05, 0.9, 0.9, size=32, bold=True, color=W, align=PP_ALIGN.CENTER)
        add_text(s, tit, 1.8, y, 11, 0.5, size=20, bold=True, color=P)
        add_text(s, desc, 1.8, y + 0.5, 11, 0.9, size=15, color=K)
        y += 1.4
    add_footer(s)

    # ----- SLIDE 6: A frase-âncora -----
    s = add_slide()
    add_title_block(s, "PARA GUARDAR", "A frase que resume esta noite")
    add_text(s, "Rede vem primeiro.", 0.7, 3.0, 11.9, 0.9, size=48, bold=True, color=P, align=PP_ALIGN.CENTER)
    add_text(s, "Internet vem depois.", 0.7, 4.0, 11.9, 0.9, size=48, bold=True, color=P, align=PP_ALIGN.CENTER)
    add_text(s, "Internet é uma rede específica — muito grande.", 0.7, 5.5, 11.9, 0.6, size=22, color=G, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s)

    # ----- SLIDE 7: Internet = entre redes -----
    s = add_slide()
    add_title_block(s, "O NOME JÁ CONTA", "INTER + NET = entre redes")
    add_text(s, [
        "Várias redes locais no mundo inteiro.",
        "Conectadas entre si através de provedores.",
        "",
        "Se o provedor cai, cada rede local continua funcionando por dentro —",
        "só não fala com o mundo lá fora.",
    ], 0.7, 3.0, 11.9, 3.5, size=22, color=K)
    add_footer(s)

    # ----- SLIDE 8: O que muda -----
    s = add_slide()
    add_title_block(s, "O QUE MUDA", "Rede local × Internet")
    rows = [
        ["Coisa", "Sem internet", "Com internet"],
        ["Onde a informação vai", "Dentro da casa/local", "Pro outro lado do mundo"],
        ["Quem cuida da estrada", "Você (rede local)", "O provedor da internet"],
        ["Quando cai", "Você conserta", "Pode ser você OU o provedor"],
        ["Custa mensalidade?", "A rede em si, não", "O acesso à internet, sim"],
    ]
    tbl = s.shapes.add_table(len(rows), 3, PInches(0.7), PInches(2.8), PInches(11.9), PInches(3.8)).table
    tbl.columns[0].width = PInches(3.3)
    tbl.columns[1].width = PInches(4.3)
    tbl.columns[2].width = PInches(4.3)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = val
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = PPt(15)
                    run.font.bold = (i == 0)
                    run.font.color.rgb = W if i == 0 else K
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = P
    add_footer(s)

    # ----- SLIDE 9: Pegadinha do vocabulário -----
    s = add_slide()
    add_title_block(s, "A PEGADINHA DO VOCABULÁRIO", "Como descobrir onde tá o problema em 3 perguntas")
    add_text(s, "CLIENTE:", 0.7, 2.8, 11.9, 0.4, size=14, bold=True, color=G)
    add_text(s, "\"Não tenho internet.\"", 0.7, 3.2, 11.9, 0.6, size=26, color=K, italic=True)
    add_text(s, "VOCÊ:", 0.7, 4.0, 11.9, 0.4, size=14, bold=True, color=D)
    add_text(s, "\"O Wi-Fi conecta?\"", 0.7, 4.4, 11.9, 0.6, size=26, color=K, italic=True)
    add_text(s, "CLIENTE:", 0.7, 5.2, 11.9, 0.4, size=14, bold=True, color=G)
    add_text(s, "\"Conecta.\"", 0.7, 5.6, 11.9, 0.6, size=26, color=K, italic=True)
    add_text(s, "→ Rede local funcionando. Problema é da internet do provedor.", 0.7, 6.4, 11.9, 0.5, size=18, bold=True, color=P)
    add_footer(s)

    # ----- SLIDE 10: Os três caminhos -----
    s = add_slide()
    add_title_block(s, "OS TRÊS CAMINHOS", "Como o aparelho chega até uma rede")
    caminhos = [("WI-FI", "Sem fio até o roteador"), ("CABO", "Com fio até o roteador"), ("DADOS MÓVEIS", "Até a antena da operadora")]
    x = 0.7
    for tit, desc in caminhos:
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(x), PInches(3.0), PInches(4.0), PInches(2.8))
        card.fill.solid()
        card.fill.fore_color.rgb = P
        card.line.fill.background()
        add_text(s, tit, x, 3.4, 4.0, 0.6, size=24, bold=True, color=W, align=PP_ALIGN.CENTER)
        add_text(s, desc, x, 4.3, 4.0, 1.5, size=15, color=W, align=PP_ALIGN.CENTER, italic=True)
        x += 4.1
    add_text(s, "Nenhum dos três É internet.", 0.7, 6.1, 11.9, 0.5, size=22, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_text(s, "Todos são caminhos até uma rede.", 0.7, 6.55, 11.9, 0.4, size=18, color=G, align=PP_ALIGN.CENTER, italic=True)
    add_footer(s)

    # ----- SLIDE 11: Wi-Fi -----
    s = add_slide()
    add_title_block(s, "UM DE CADA VEZ · 1", "WI-FI", "Sem fio até o roteador. Só isso.")
    add_text(s, [
        "O que faz",
        "  Conecta o celular ou notebook ao roteador, sem fio",
        "",
        "Cria internet?",
        "  Não. É só o caminho.",
        "",
        "Quando falha",
        "  Wi-Fi conecta e nada abre → apareceu o triângulo com exclamação",
        "  A rede da casa tá OK, mas o roteador não tá recebendo internet",
    ], 0.7, 2.8, 11.9, 4.0, size=18, color=K)
    add_footer(s)

    # ----- SLIDE 12: Cabo -----
    s = add_slide()
    add_title_block(s, "UM DE CADA VEZ · 2", "CABO DE REDE", "Mesma coisa que Wi-Fi, com fio.")
    add_text(s, [
        "O que faz",
        "  Conecta o computador ao roteador, com fio",
        "",
        "Vantagens",
        "  Mais estável, mais rápido, sem interferência",
        "  É o que o técnico usa quando o Wi-Fi não dá conta",
        "",
        "Quando falha",
        "  Se o roteador do outro lado não tem internet, o cabo tá lá, mas nada abre",
    ], 0.7, 2.8, 11.9, 4.0, size=18, color=K)
    add_footer(s)

    # ----- SLIDE 13: Dados móveis -----
    s = add_slide()
    add_title_block(s, "UM DE CADA VEZ · 3", "DADOS MÓVEIS", "O caminho que costuma vir com internet junto.")
    add_text(s, [
        "O que faz",
        "  Conecta o celular à antena da operadora",
        "",
        "Por que é diferente",
        "  A operadora vende a internet junto com o sinal",
        "  Único dos três em que 'ligou = tem internet' costuma valer",
        "",
        "Mas mesmo aqui tem separação",
        "  Sem sinal → não pega. Sem pacote → sinal fica, internet foi. Operadora com defeito → sinal e nada abre.",
    ], 0.7, 2.8, 11.9, 4.0, size=17, color=K)
    add_footer(s)

    # ----- SLIDE 14: Quadro dos três caminhos -----
    s = add_slide()
    add_title_block(s, "DE UMA OLHADA", "O quadro dos três caminhos")
    rows = [
        ["Caminho", "O que faz", "Cria internet?", "Depende de quê"],
        ["Wi-Fi", "Liga o aparelho ao roteador", "Não", "Do roteador ter internet do provedor"],
        ["Cabo de rede", "Liga o aparelho ao roteador", "Não", "Do roteador ter internet do provedor"],
        ["Dados móveis", "Liga o celular à antena", "Não, mas vem junto", "Da operadora ter sinal e pacote ativo"],
    ]
    tbl = s.shapes.add_table(len(rows), 4, PInches(0.7), PInches(2.8), PInches(11.9), PInches(3.6)).table
    tbl.columns[0].width = PInches(2.4)
    tbl.columns[1].width = PInches(3.5)
    tbl.columns[2].width = PInches(2.4)
    tbl.columns[3].width = PInches(3.6)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = val
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = PPt(14)
                    run.font.bold = (i == 0)
                    run.font.color.rgb = W if i == 0 else K
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = P
    add_text(s, "Nenhum dos três é internet. Todos são caminhos.", 0.7, 6.5, 11.9, 0.4, size=18, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s)

    # ----- SLIDE 15: A demonstração -----
    s = add_slide()
    add_title_block(s, "DEMONSTRAÇÃO", "Um celular, quatro cenários")
    add_text(s, [
        "1. Só Wi-Fi ligado → funciona (o roteador tem internet)",
        "2. Só dados móveis → funciona (a operadora tem internet)",
        "3. Wi-Fi + dados móveis → o celular escolhe o Wi-Fi (não gasta pacote)",
        "4. Modo avião → nada funciona (todos os caminhos fechados)",
    ], 0.7, 3.0, 11.9, 3.0, size=20, color=K)
    add_text(s, "O celular tá o tempo todo escolhendo por qual caminho vai.", 0.7, 6.1, 11.9, 0.4, size=18, bold=True, color=P, align=PP_ALIGN.CENTER)
    add_text(s, "Você só precisa saber que os caminhos são separados.", 0.7, 6.5, 11.9, 0.4, size=16, color=G, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s)

    # ----- SLIDE 16: Os ícones do celular -----
    s = add_slide()
    add_title_block(s, "OS ÍCONES DO CELULAR", "Um relatório em miniatura no canto de cima")
    add_text(s, [
        "Se você aprender a ler os quatro ícones que importam,",
        "resolve metade do chamado antes de perguntar qualquer coisa ao cliente.",
    ], 0.7, 3.3, 11.9, 2.0, size=22, color=K, align=PP_ALIGN.CENTER)
    add_text(s, "Nas próximas quatro telas: um ícone por vez.", 0.7, 5.8, 11.9, 0.5, size=16, color=G, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s)

    # ----- SLIDE 17: Barras de sinal -----
    s = add_slide()
    add_title_block(s, "ÍCONE 1", "BARRAS DE SINAL", "Sinal da antena da operadora")
    add_text(s, [
        "Quantas barras acesas",
        "  Força do sinal da operadora chegando no celular",
        "",
        "Tem letra do lado (4G, 5G)",
        "  Dados móveis estão ligados",
        "",
        "Sem letra do lado",
        "  Pega ligação, mas não passa internet pela operadora",
    ], 0.7, 2.8, 11.9, 4.0, size=18, color=K)
    add_footer(s)

    # ----- SLIDE 18: Wi-Fi -----
    s = add_slide()
    add_title_block(s, "ÍCONE 2", "SINAL DE WI-FI (arco de curvas)", "Conectado num Wi-Fi")
    add_text(s, [
        "O que ele diz",
        "  O celular está conectado a algum Wi-Fi",
        "",
        "Quantas curvinhas",
        "  Quão perto do roteador",
        "",
        "O que ele NÃO diz",
        "  Nada sobre se a internet do outro lado do roteador está funcionando",
    ], 0.7, 2.8, 11.9, 4.0, size=18, color=K)
    add_footer(s)

    # ----- SLIDE 19: Triângulo -----
    s = add_slide()
    add_title_block(s, "ÍCONE 3", "TRIÂNGULO COM EXCLAMAÇÃO", "O ícone mais importante deste curso")
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(0.7), PInches(3.0), PInches(11.9), PInches(2.5))
    box.fill.solid()
    box.fill.fore_color.rgb = PRGBColor(0xFF, 0xF2, 0xEF)
    box.line.color.rgb = R
    box.line.width = Emu(30000)
    add_text(s, "\"Estou conectado, mas não consigo abrir nada.\"", 0.9, 3.3, 11.5, 0.7, size=24, bold=True, color=R, italic=True)
    add_text(s, [
        "É o caso mais comum de chamado de casa.",
        "Rede da casa: funcionando.",
        "Internet do provedor: com problema.",
    ], 0.9, 4.1, 11.5, 1.4, size=17, color=K)
    add_text(s, "Se o cliente mandar foto desse ícone antes de você sair, você já sabe o que vai encontrar.", 0.7, 5.8, 11.9, 0.5, size=15, color=G, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s)

    # ----- SLIDE 20: Avião -----
    s = add_slide()
    add_title_block(s, "ÍCONE 4", "MODO AVIÃO", "Desliga tudo de uma vez")
    add_text(s, [
        "O que ele desliga",
        "  Wi-Fi, dados móveis, sinal de ligação, Bluetooth (às vezes)",
        "",
        "Quando aparece por engano",
        "  Cliente encostou no botão sem querer",
        "  Aparelho ligou em modo avião por rotina de bateria",
        "",
        "A primeira pergunta pro cliente 'sem sinal'",
        "  \"Você tá com o aviãozinho ligado?\"",
    ], 0.7, 2.8, 11.9, 4.0, size=17, color=K)
    add_footer(s)

    # ----- SLIDE 21: Jogo do ícone -----
    s = add_slide()
    add_title_block(s, "JOGO DO ÍCONE", "Qual ícone o celular do cliente está mostrando?")
    add_text(s, [
        "1. \"Liguei no Wi-Fi de casa mas o WhatsApp não abre.\"",
        "     → Wi-Fi aceso + triângulo com exclamação",
        "",
        "2. \"Saí de casa e o WhatsApp continua funcionando.\"",
        "     → Sem Wi-Fi + barras com 4G ou 5G",
        "",
        "3. \"Meu celular tá sem nada.\"",
        "     → Modo avião, lugar sem sinal, ou desligado. Uma pergunta resolve.",
        "",
        "4. \"Tô no Wi-Fi do vizinho e não abre nada.\"",
        "     → Wi-Fi aceso + triângulo. Rede do vizinho sem internet.",
    ], 0.7, 2.8, 11.9, 4.2, size=15, color=K)
    add_footer(s)

    # ----- SLIDE 22: Prática -----
    s = add_slide()
    add_title_block(s, "PRÁTICA · 40 MIN", "Separar as palavras")
    add_text(s, [
        "1. Equipes de 3 ou 4 pessoas",
        "",
        "2. Cada aluno preenche a folha individualmente (15 min)",
        "",
        "3. Discussão em equipe (10 min)",
        "",
        "4. Rodada aberta: professor escolhe 6 cenários e a turma debate (15 min)",
        "",
        "Regra da prática",
        "  Hoje ninguém resolve nada.",
        "  Hoje é só separar: o problema tá na rede ou na internet?",
    ], 0.7, 2.8, 11.9, 4.2, size=17, color=K)
    add_footer(s)

    # ----- SLIDE 23: O resumo -----
    s = add_slide()
    add_title_block(s, "O RESUMO DA NOITE", "Quatro ideias que ficam")
    add_text(s, [
        "• Rede é conexão entre coisas.",
        "• Internet é a rede das redes.",
        "• Wi-Fi, cabo e dados móveis são caminhos até uma rede.",
        "• Os ícones do celular contam a história antes de qualquer pergunta.",
    ], 0.7, 3.0, 11.9, 3.6, size=22, color=K)
    add_footer(s)

    # ----- SLIDE 24: Fechando -----
    s = add_slide()
    add_title_block(s, "FECHANDO", "Cinco perguntas para conferir")
    add_text(s, [
        "1. Rede e internet: qual vem primeiro?",
        "2. Cite uma rede que funciona sem internet.",
        "3. Wi-Fi é a mesma coisa que internet?",
        "4. Ícone: Wi-Fi aceso + triângulo com exclamação. O que quer dizer?",
        "5. Cliente sem sinal nenhum. Qual é a sua primeira pergunta?",
    ], 0.7, 2.8, 11.9, 4.0, size=18, color=K)
    add_footer(s)

    # ----- SLIDE 25: O que fica desta noite -----
    s = add_slide()
    add_title_block(s, "O QUE FICA DESTA NOITE", "Cinco coisas para levar")
    add_text(s, [
        "• Rede é conexão entre coisas. Nem toda rede é internet.",
        "• Internet é a rede das redes. Precisa de provedor.",
        "• Wi-Fi e cabo são caminhos até o roteador. Dados móveis é caminho até a antena da operadora.",
        "• Só dados móveis costuma vir com internet junto — os outros dois dependem do roteador estar conectado.",
        "• Quando o cliente diz \"sem internet\", separe: é rede da casa, ou é internet do provedor?",
    ], 0.7, 2.8, 11.9, 4.2, size=16, color=K)
    add_footer(s)

    # ----- SLIDE 26: Para a próxima noite -----
    s = add_slide()
    add_title_block(s, "PARA A PRÓXIMA NOITE", "Duas tarefas simples")
    add_text(s, [
        "• Preste atenção nos ícones do seu celular esta semana",
        "     — anote duas vezes por dia o que apareceu",
        "",
        "• Se em algum momento você ficou \"sem internet\",",
        "     tente descobrir por qual caminho falhou (Wi-Fi? dados móveis?)",
        "",
        "• Se tiver um roteador em casa,",
        "     tire uma foto da traseira e do embaixo dele",
        "",
        "Na próxima noite: os cinco equipamentos de rede, com a mão neles.",
    ], 0.7, 2.8, 11.9, 4.2, size=17, color=K)
    add_footer(s)

    path = os.path.join(OUTDIR, "Aula - Redes 1 - Rede nao e internet.pptx")
    prs.save(path)
    return path


# ============================================================
# FOLHA DO ALUNO
# ============================================================
def build_folha():
    doc = Document()

    for section in doc.sections:
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)

    # Cabeçalho
    p = doc.add_paragraph()
    r = p.add_run("SEPARAR AS PALAVRAS")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = COR_PRIMARIA

    p = doc.add_paragraph()
    r = p.add_run("Redes de Computadores · Noite 1 · Curso FIC · CEJA Itapiranga")
    r.font.size = Pt(10)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    # linha
    add_hr(doc)

    # Nome / data
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("Nome: ")
    r.font.size = Pt(11)
    r.bold = True
    r = p.add_run("_________________________________________________     Data: ____ / ____ / ______")
    r.font.size = Pt(11)

    # 1
    add_heading(doc, "1  Rede ou internet?", level=2)
    add_para(doc, "Marque com R se for rede que funciona sem internet, com I se depender da internet.", italic=True, size=10, color=COR_CINZA)
    doc.add_paragraph()

    situacoes = [
        "O computador do caixa manda o valor pro computador do gerente do mercado",
        "Você abre o Instagram no celular pelos dados móveis",
        "Você imprime da secretaria na impressora da secretaria",
        "Você manda uma foto por Bluetooth pro colega ao lado",
        "Você assiste vídeo do YouTube conectado no Wi-Fi de casa",
        "Uma câmera de segurança grava direto num HD, dentro do escritório",
        "O caixa eletrônico do banco confere seu saldo",
        "Dois notebooks trocam arquivos ligados por um cabo direto",
    ]
    for sit in situacoes:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run("(       )   ")
        r.font.size = Pt(11)
        r.bold = True
        r = p.add_run(sit)
        r.font.size = Pt(11)

    # 2
    add_heading(doc, "2  Os três caminhos", level=2)
    add_para(doc, "Preencha a tabela com o que você lembra da aula.", italic=True, size=10, color=COR_CINZA)

    tbl = add_table(doc, [
        ["Caminho", "O que faz?", "Cria internet?", "Depende de quê pra ter internet"],
        ["Wi-Fi", "", "", ""],
        ["Cabo de rede", "", "", ""],
        ["Dados móveis", "", "", ""],
    ], col_widths=[3, 4.5, 3, 5.8])

    # Aumenta altura das linhas (linhas em branco pra preencher à mão)
    for row in tbl.rows[1:]:
        row.height = Cm(1.4)

    # 3
    add_heading(doc, "3  Escreva com suas palavras", level=2)
    p = doc.add_paragraph()
    r = p.add_run("Rede e internet são a mesma coisa? Por quê?")
    r.font.size = Pt(11)
    r.bold = True
    for _ in range(4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("_" * 90)
        r.font.size = Pt(11)
        r.font.color.rgb = COR_CINZA

    # Página 2
    doc.add_page_break()

    p = doc.add_paragraph()
    r = p.add_run("OS ÍCONES DO CELULAR")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = COR_PRIMARIA

    p = doc.add_paragraph()
    r = p.add_run("Análise de cenários · Noite 1 · 2 de 2")
    r.font.size = Pt(10)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    add_hr(doc)

    # 4 ícones - tabela ligar
    add_heading(doc, "4  O que cada ícone quer dizer", level=2)
    add_para(doc, "Ligue cada ícone à situação correta escrevendo o número da situação no espaço do ícone.", italic=True, size=10, color=COR_CINZA)

    add_table(doc, [
        ["Ícone", "Nº", "Situações"],
        ["Barras de sinal do celular", "____", "1. Estou conectado num Wi-Fi mas não abre nada"],
        ["Símbolo de Wi-Fi (arco)", "____", "2. Estou sem qualquer sinal, meu celular tá isolado"],
        ["Triângulo com exclamação", "____", "3. Estou pegando sinal da antena da operadora"],
        ["Avião", "____", "4. Estou perto de um roteador e conectado nele"],
    ], col_widths=[6, 1.8, 8.5])

    # 5 cinco cenários
    add_heading(doc, "5  Cinco cenários — onde está o problema?", level=2)
    add_para(doc, "Marque com X. Não precisa resolver o problema, só apontar de que lado ele está.", italic=True, size=10, color=COR_CINZA)

    tbl = add_table(doc, [
        ["Cenário", "Rede da casa", "Internet do provedor", "Rede móvel / operadora", "Outra coisa"],
        ["1. Wi-Fi conecta mas WhatsApp não abre", "", "", "", ""],
        ["2. Impressora não imprime, mas a internet funciona", "", "", "", ""],
        ["3. Saiu de casa e o Instagram parou no celular", "", "", "", ""],
        ["4. TV conecta na rede, mas Netflix diz 'sem internet'", "", "", "", ""],
        ["5. Vizinhos da rua toda dizem que estão sem internet", "", "", "", ""],
    ], col_widths=[7, 2.4, 3, 3.2, 2.4])
    for row in tbl.rows[1:]:
        row.height = Cm(1.0)

    # tarefa
    doc.add_paragraph()
    add_heading(doc, "TAREFA DA SEMANA", level=2, color=COR_DESTAQUE)
    add_bullet(doc, "Anote duas vezes por dia o que aparece no canto de cima do seu celular")
    add_bullet(doc, "Se em algum momento ficar \"sem internet\", tente descobrir por qual caminho falhou")
    add_bullet(doc, "Traga anotado — vamos comparar na próxima aula")

    # rodapé
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Traga esta folha preenchida · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Folha do Aluno - Redes 1 - Separar as palavras.docx")
    doc.save(path)
    return path


# ============================================================
# RESUMO DO ALUNO
# ============================================================
def build_resumo():
    doc = Document()

    for section in doc.sections:
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)

    # Cabeçalho
    p = doc.add_paragraph()
    r = p.add_run("REDE NÃO É INTERNET")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = COR_PRIMARIA

    p = doc.add_paragraph()
    r = p.add_run("Resumo do aluno · Noite 1 · Redes de Computadores · Curso FIC · CEJA Itapiranga")
    r.font.size = Pt(10)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    add_hr(doc)

    # Ideia principal
    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Todo dia a gente mistura quatro palavras: rede, internet, Wi-Fi e dados móveis. Cada uma resolve num lugar diferente. Quando o cliente diz \"estou sem internet\", ele pode estar sem qualquer uma das quatro. Separar é o primeiro passo de qualquer atendimento.")

    # Rede
    add_heading(doc, "O que é REDE", level=2)
    add_para(doc, "Duas ou mais coisas ligadas trocando informação. Só isso.")
    add_bullet(doc, "Impressora da secretaria com o computador da recepção — é rede.")
    add_bullet(doc, "Máquina do caixa e computador do gerente do mercado — é rede.")
    add_bullet(doc, "Fone Bluetooth ligado no celular — é rede de dois.")
    add_para(doc, "Rede não precisa de internet para funcionar.", bold=True, color=COR_PRIMARIA)

    # Internet
    add_heading(doc, "O que é INTERNET", level=2)
    add_para(doc, "A rede das redes. Várias redes locais conectadas entre si, no mundo inteiro. Cada casa entra na internet através de um provedor.")
    add_para(doc, "Se o provedor cai, a rede da casa continua funcionando por dentro — só não fala com o mundo.", bold=True, color=COR_PRIMARIA)

    # Três caminhos
    add_heading(doc, "Os três caminhos até uma rede", level=2)
    add_table(doc, [
        ["Caminho", "Como funciona"],
        ["Wi-Fi", "Sem fio até o roteador da casa. Não é internet — é só o caminho."],
        ["Cabo de rede", "Com fio até o roteador. Mais estável, mais rápido."],
        ["Dados móveis", "Até a antena da operadora. Costuma vir com internet junto, porque a operadora vende as duas coisas."],
    ], col_widths=[3.5, 12])

    # Ícones
    add_heading(doc, "Os ícones do celular contam a história", level=2)
    add_table(doc, [
        ["Ícone", "O que quer dizer"],
        ["Barras de sinal", "Força do sinal da operadora chegando no celular"],
        ["4G / 5G ao lado das barras", "Dados móveis estão ativados"],
        ["Arco de Wi-Fi", "Conectado num Wi-Fi (não diz se a internet do outro lado funciona)"],
        ["Triângulo com exclamação", "Conectado, mas nada abre. Provedor com problema, quase sempre."],
        ["Avião", "Modo avião — nenhum caminho aberto"],
    ], col_widths=[5, 10.5])

    # Como pensar
    add_heading(doc, "Como pensar num chamado", level=2)
    add_bullet(doc, "1. O aparelho tem sinal de algum caminho aceso? (Wi-Fi, dados móveis)")
    add_bullet(doc, "2. Se sim, apareceu o triângulo com exclamação?")
    add_bullet(doc, "3. Se apareceu, o problema é da internet — não da rede da casa.")
    add_bullet(doc, "4. Se não apareceu e ainda assim não abre nada, pode ser configuração do aparelho.")

    # Três frases pra guardar
    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Rede vem primeiro. Internet vem depois.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Wi-Fi não é internet — é o caminho até o roteador.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Quando o cliente diz 'sem internet', pergunte primeiro qual caminho falhou.\"", bold=True, color=COR_PRIMARIA)

    # Se testar
    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Cite uma rede que funciona sem internet.")
    add_bullet(doc, "Você tá com Wi-Fi aceso e o triângulo com exclamação apareceu. Onde tá o problema?")
    add_bullet(doc, "Seu vizinho tá sem internet e você não. Por que isso é possível?")

    # rodapé
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Este material é seu. Guarde e volte quando precisar. · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Resumo do Aluno - Redes 1 - Rede nao e internet.docx")
    doc.save(path)
    return path


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    print("Gerando Roteiro...")
    print("  →", build_roteiro())
    print("Gerando Slides...")
    print("  →", build_slides())
    print("Gerando Folha do Aluno...")
    print("  →", build_folha())
    print("Gerando Resumo do Aluno...")
    print("  →", build_resumo())
    print("\nPronto.")
