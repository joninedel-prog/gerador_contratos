"""Gera os 4 arquivos da Noite 3 do curso de Redes.

Noite 3 — Por onde o sinal viaja: cabo, fibra e ondas
Padrão prática-pesada: 105 min de atividade em 210 min totais.
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from common import (
    COR_PRIMARIA, COR_ALERTA, COR_DESTAQUE, COR_CINZA,
    add_heading, add_para, add_bullet, add_table, add_hr,
    setup_page, cabecalho_documento,
    PPalette, new_presentation, add_slide, add_text, add_title_block,
    add_slide_cover, add_footer, add_table_slide, add_bar,
)
from pptx.util import Inches as PInches, Pt as PPt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUTDIR = "/home/nedel/curso"
NOITE = "Noite 3"
TITULO_CURTO = "Por onde o sinal viaja"
FOOTER = f"Redes de Computadores · {NOITE}"


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "POR ONDE O SINAL VIAJA",
        f"Roteiro do professor · Redes de Computadores · {NOITE} · Curso FIC · CEJA Itapiranga – SC",
    )

    # A IDEIA
    add_heading(doc, "A IDEIA DESTA NOITE")
    add_para(doc, "Na Noite 2 os alunos viram os cinco equipamentos. Faltou uma coisa: por onde a informação viaja entre eles. Cabo par trançado, fibra óptica e onda de rádio — cada um resolve num lugar diferente. O técnico que não sabe escolher enfia cabo caro onde não precisa e passa Wi-Fi onde tinha que passar cabo.")
    add_para(doc, "Hoje eles reconhecem os três meios que vão usar no dia a dia, aprendem a ler a capa de um cabo Ethernet e escolhem o meio certo pra quatro cenários de instalação. Na noite que vem, o alicate entra e cada um monta o próprio cabo.")

    add_heading(doc, "O que se resolve nesta noite", level=2)
    add_para(doc, "Três meios de transmissão. A capa de um cabo lida como quem lê remédio. E uma leitura de Wi-Fi da região usando o celular do aluno.")

    add_heading(doc, "Esta é a terceira de quinze noites", level=2)
    add_para(doc, "A aula é 50% mão-na-massa. Se você terminar cedo, tem a CAIXA DE FERRAMENTAS DIDÁTICAS no fim deste roteiro com atividades extras. Não corte prática pra ganhar tempo — corte exposição.")

    add_hr(doc)

    # ANTES DA AULA
    add_heading(doc, "ANTES DA AULA")

    add_heading(doc, "Material da mesa", level=2)
    add_bullet(doc, "4 a 6 cabos de rede prontos (categorias diferentes se possível — Cat5e, Cat6, e um blindado se tiver)")
    add_bullet(doc, "1 cabo coaxial de antena (pra comparar)")
    add_bullet(doc, "1 cabo de telefone com RJ11 (pra comparar — quase todo mundo confunde com RJ45)")
    add_bullet(doc, "1 cabo de energia (pra comparar contra cabo de dados)")
    add_bullet(doc, "Se der, 1 pedaço curto de fibra óptica — ou uma foto/vídeo pra passar no projetor")
    add_bullet(doc, "Alicate crimpador na mesa, só pra mostrar (a crimpagem é na Noite 4)")
    add_bullet(doc, "1 roteador ligado com Wi-Fi ativo — pra a prática de Wi-Fi Scan")

    add_heading(doc, "Prepare antes", level=2)
    add_bullet(doc, "Instale no seu celular o app \"Wi-Fi Analyzer\" (Android, gratuito). Se você usa iPhone, deixe um Android emprestado — o iOS não deixa ver canais.")
    add_bullet(doc, "Confira que dá pra ver várias redes na sala (rede da escola + vizinhos + hotspot do seu celular como plano B).")
    add_bullet(doc, "Como plano B, crie um hotspot no seu celular com nome específico (ex.: \"REDE_TESTE_PROF\") pra usar como referência conhecida no scan.")
    add_bullet(doc, "Confira que os cabos da mesa estão íntegros e legíveis (a marcação da capa é o que a turma vai ler).")

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "Folha do aluno (3 páginas), uma por pessoa")
    add_bullet(doc, "Resumo do aluno, uma por pessoa (pra levar pra casa)")

    add_hr(doc)

    # QUADRO DE TEMPO
    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Retomada + tarefa da Noite 2", "10 min", "Jogo relâmpago dos 5 equipamentos, conferir fotos"],
        ["2. Três meios, três decisões", "25 min", "Como a informação viaja: elétron, luz, onda"],
        ["3. Cabo de rede em detalhe", "25 min", "Par trançado, categorias, RJ45, ler a capa"],
        ["4. PRÁTICA — Identificar cabos", "35 min", "Equipes preenchem ficha de identificação"],
        ["INTERVALO", "15 min", ""],
        ["5. Fibra e ondas", "20 min", "Como funcionam, onde usar cada uma"],
        ["6. PRÁTICA — Wi-Fi Scan", "25 min", "Turma vê as redes da região com o celular"],
        ["7. PRÁTICA — 4 estudos de caso", "25 min", "Escolher o meio certo pra cada cenário"],
        ["8. Duelo relâmpago", "15 min", "Time A × Time B — 10 perguntas rápidas"],
        ["9. Fechamento + tarefa", "15 min", "Quiz de saída e anúncio da Noite 4"],
    ], col_widths=[6, 2, 8.5])
    add_para(doc, "105 min de prática ativa (50% do total). Se atrasar: corte o Bloco 8 antes do Bloco 6 ou 7. Nunca corte prática.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — RETOMADA + TAREFA · 10 MIN")

    add_heading(doc, "Jogo relâmpago dos 5 equipamentos — 5 min", level=2)
    add_para(doc, "Fale a frase, a turma responde qual equipamento é. Sem levantar a mão. Quem souber, fala.")
    add_bullet(doc, "\"Não cria rede, só tem uma porta de saída\" → Modem")
    add_bullet(doc, "\"Cria rede e distribui pra casa toda\" → Roteador")
    add_bullet(doc, "\"Aumenta o número de portas de cabo\" → Switch")
    add_bullet(doc, "\"Leva Wi-Fi onde o roteador não alcança\" → Ponto de acesso")
    add_bullet(doc, "\"Onde o cabo entra no computador\" → Placa de rede")

    add_heading(doc, "Conferir a tarefa — 5 min", level=2)
    add_para(doc, "\"Quem trouxe foto do roteador? Quem trouxe alguma anotação da capa de cabo?\"", italic=True)
    add_para(doc, "Não faça auditoria. Só valide que quem trouxe seja reconhecido, e diga: \"quem não trouxe hoje ainda dá tempo — o cabo vai voltar em duas noites.\"", italic=True)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — TRÊS MEIOS, TRÊS DECISÕES · 25 MIN")

    add_heading(doc, "A frase que abre o bloco", level=2)
    add_para(doc, "\"Todo equipamento de rede que a gente viu se conecta com outro por um dos três meios. Não existe um quarto. Vocês vão passar a vida escolhendo entre esses três.\"", italic=True)

    add_heading(doc, "Os três no quadro — 10 min", level=2)
    add_para(doc, "Escreva grande:")
    add_bullet(doc, "CABO — pulsos elétricos correndo dentro de fio de cobre")
    add_bullet(doc, "FIBRA — pulsos de luz correndo dentro de vidro")
    add_bullet(doc, "ONDA — sinal de rádio no ar")
    add_para(doc, "\"Elétron, luz, onda. Três jeitos de fazer a mesma coisa: mover informação de um ponto pra outro. Cada um resolve um problema diferente. Vocês vão ver por quê agora.\"", italic=True)

    add_heading(doc, "Comparação em uma tabela — 10 min", level=2)
    add_table(doc, [
        ["Item", "Cabo (par trançado)", "Fibra", "Onda (Wi-Fi/celular)"],
        ["Como funciona", "Pulsos elétricos no cobre", "Pulsos de luz no vidro", "Sinal de rádio no ar"],
        ["Velocidade", "Alta", "Muito alta", "Depende — pode ser rápido"],
        ["Alcance", "Até 100m sem repetir", "Quilômetros sem repetir", "Wi-Fi: metros. Celular: km."],
        ["Interferência", "Sofre", "Não sofre", "Sofre muito"],
        ["Custo por metro", "Barato", "Caro, mas caiu", "Sem \"por metro\" — é sinal"],
        ["Quem instala", "Você (Noite 4)", "Equipe especializada", "Não instala — configura"],
        ["Onde é comum", "Do roteador aos aparelhos fixos", "Da rua até a casa", "Do roteador aos móveis"],
    ], col_widths=[3, 4, 4, 4.5])

    add_heading(doc, "A decisão em três perguntas — 5 min", level=2)
    add_para(doc, "\"Vocês vão escolher assim, no chamado:\"", italic=True)
    add_bullet(doc, "1. Vai passar mais de 100 metros? → Fibra")
    add_bullet(doc, "2. O aparelho fica parado no mesmo lugar? → Cabo (mais estável)")
    add_bullet(doc, "3. O aparelho se mexe (celular, notebook)? → Wi-Fi (é o único jeito)")
    add_para(doc, "\"Não tem regra que caiba em 3 perguntas sempre. Mas essa aqui vai valer em 80% dos casos que vocês vão ver.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — CABO DE REDE EM DETALHE · 25 MIN")

    add_heading(doc, "O nome completo — 3 min", level=2)
    add_para(doc, "\"O cabo que a gente chama de cabo de rede tem nome técnico: par trançado. Vamos ver por que trançado.\"", italic=True)

    add_heading(doc, "Por que trançado — 5 min", level=2)
    add_para(doc, "Pegue um pedaço de cabo cortado (se puder abrir a ponta de um). Mostre.")
    add_para(doc, "\"Dentro tem 8 fios, em 4 pares. Cada par é trançado sobre si mesmo. A trança protege o sinal contra a interferência do próximo par e do ambiente. Não é enfeite. Se destrançar demais na ponta ao crimpar, o cabo perde qualidade — e vocês vão sentir isso na Noite 4.\"", italic=True)

    add_heading(doc, "Categorias — 10 min", level=2)
    add_table(doc, [
        ["Categoria", "O que suporta", "Onde a gente vê", "Preço"],
        ["Cat5e", "Até 1 Gbps, até 100m", "Casa comum, obra antiga", "Mais barato"],
        ["Cat6", "1 Gbps sem esforço, começa 10 Gbps curto", "Instalação nova, escritório", "Um pouco mais caro"],
        ["Cat6a", "10 Gbps até 100m", "Data center, prédio corporativo", "Bem mais caro"],
        ["Cat7, Cat8", "Especializado", "Você quase nunca vai ver", "Caro e exigente"],
    ], col_widths=[2.8, 5.5, 5, 3])
    add_para(doc, "\"Pra 99% dos chamados de casa, Cat5e ou Cat6 resolve. Não venda Cat6a pra ninguém que só quer internet doméstica — não faz diferença nenhuma pro cliente e você fica com fama de empurrar produto.\"", italic=True, color=COR_PRIMARIA)

    add_heading(doc, "Blindado × não blindado — 3 min", level=2)
    add_bullet(doc, "Não blindado (UTP): o comum. Barato, mais fino, mais fácil de crimpar.")
    add_bullet(doc, "Blindado (STP, FTP): tem folha metálica protegendo. Usa em ambiente com muita interferência (indústria, perto de motor grande).")
    add_para(doc, "\"Se o cliente não tem torno mecânico no quintal, ele não precisa de blindado.\"", italic=True)

    add_heading(doc, "Como ler a capa — 4 min", level=2)
    add_para(doc, "Pegue um cabo qualquer. Mostre a impressão na capa.")
    add_bullet(doc, "Marca (Furukawa, Nexans, marcas genéricas)")
    add_bullet(doc, "Categoria (CAT5E, CAT6)")
    add_bullet(doc, "Tipo (UTP, FTP)")
    add_bullet(doc, "Certificação (CMR = uso interno, CMX = uso externo, etc.)")
    add_bullet(doc, "Metragem impressa a cada metro (dá pra medir sem trena)")
    add_para(doc, "\"A metragem impressa é o segredo do técnico que economiza tempo. Você lê o número no início do cabo e no fim, subtrai, e sabe o comprimento. Muito melhor que esticar o cabo pra medir.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · IDENTIFICAR CABOS · 35 MIN")

    add_heading(doc, "Instrução — 3 min", level=2)
    add_para(doc, "\"Vocês vão dividir em equipes de 3 ou 4. Cada equipe vai passar pelas mesas (ou receber os cabos), identificando o que cada cabo é. Não vale abrir cabo. Vale olhar, ler a capa e comparar com o quadro que a gente montou.\"", italic=True)

    add_heading(doc, "Como distribuir — escolha", level=2)
    add_bullet(doc, "Se tiver várias mesas: cada mesa com 2-3 cabos, equipes rodam a cada 8 min")
    add_bullet(doc, "Se for uma mesa só: equipes vão em rodízio de 6 min por vez")
    add_bullet(doc, "Se tiver poucos cabos: divida em conjuntos e faça circular")

    add_heading(doc, "O que a ficha pede — 27 min de trabalho", level=2)
    add_para(doc, "Para cada cabo:")
    add_bullet(doc, "1. Tipo (par trançado, coaxial, telefone, energia) — anotar como reconheceu")
    add_bullet(doc, "2. Categoria (se par trançado)")
    add_bullet(doc, "3. Blindagem (UTP ou blindado)")
    add_bullet(doc, "4. Estado do conector (bom, danificado, faltando)")
    add_bullet(doc, "5. Comprimento aproximado (pela marca ou pela vista)")
    add_bullet(doc, "6. Onde ele seria adequado usar (2-3 opções)")

    add_heading(doc, "Você circula", level=2)
    add_para(doc, "Não corrige direto. Faz pergunta:")
    add_bullet(doc, "\"Como você sabe que esse é par trançado e não coaxial?\"")
    add_bullet(doc, "\"Se destrançar demais essa ponta, o que acontece?\"")
    add_bullet(doc, "\"Que categoria você usaria pra uma casa de 3 quartos?\"")
    add_bullet(doc, "\"Esse aqui é RJ45 ou RJ11? Como você separou?\"")

    add_heading(doc, "Fechamento do bloco — 5 min", level=2)
    add_para(doc, "Cada equipe apresenta 1 cabo pra turma inteira (30 segundos por equipe). Sorteio na hora.")

    add_hr(doc)
    add_heading(doc, "INTERVALO · 15 MIN", color=COR_CINZA)
    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — FIBRA E ONDAS · 20 MIN")

    add_heading(doc, "Fibra — 8 min", level=2)
    add_para(doc, "\"Vocês já viram fibra na Noite 2 — o cabo que vem do poste e entra na caixinha. Agora vamos entender por dentro.\"", italic=True)
    add_bullet(doc, "Um fio de vidro fininho, do tamanho de um cabelo")
    add_bullet(doc, "Um LED ou laser manda pulsos de luz por dentro")
    add_bullet(doc, "Do outro lado, um sensor lê os pulsos")
    add_bullet(doc, "Como é luz e não elétron, não sofre interferência de motor, transformador ou raio")
    add_bullet(doc, "Distância enorme: dá pra passar 40 km sem repetir o sinal")
    add_bullet(doc, "Frágil: dobra demais e quebra. Emenda precisa de máquina.")
    add_para(doc, "\"Vocês vão trabalhar do roteador da casa pra dentro. Emenda de fibra é serviço do provedor. Mas vão ver muito cabo de fibra chegando na casa — precisam reconhecer.\"", italic=True)

    add_heading(doc, "Ondas — 10 min", level=2)
    add_para(doc, "Escreva no quadro:")
    add_bullet(doc, "Wi-Fi: onda curta, alcance de metros")
    add_bullet(doc, "Bluetooth: onda mais curta, alcance de metros também")
    add_bullet(doc, "Celular (4G/5G): onda que a antena da operadora manda, alcance de quilômetros")
    add_bullet(doc, "Satélite: onda que vem do espaço, alcance planetário")
    add_para(doc, "\"Todas essas quatro são a mesma coisa: onda de rádio. O que muda é a frequência e a potência.\"", italic=True)
    add_para(doc, "Frequência = quão rápido a onda oscila. Duas frequências que vocês precisam conhecer:")
    add_bullet(doc, "2,4 GHz — atravessa parede melhor, tem interferência de micro-ondas e fone, é disputada")
    add_bullet(doc, "5 GHz — mais rápida, atravessa parede pior, é menos disputada")
    add_para(doc, "\"Vocês vão ver as duas ao vivo na próxima prática.\"", italic=True)

    add_heading(doc, "Um exemplo pra fixar — 2 min", level=2)
    add_para(doc, "\"A câmera Wi-Fi de casa e o rádio FM do carro operam por ondas de rádio, mas em frequências completamente diferentes. Rádio FM é 88 a 108 MHz. Roteador é 2.400 MHz. Mesma família, potências e usos completamente diferentes.\"", italic=True)

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — PRÁTICA · WI-FI SCAN COM CELULAR · 25 MIN")

    add_heading(doc, "Instrução — 5 min", level=2)
    add_para(doc, "\"Todo mundo com celular Android vai instalar um app — Wi-Fi Analyzer (grátis). Quem tem iPhone faz dupla com quem tem Android. Se ninguém tiver celular, temos o meu como plano B.\"", italic=True)
    add_para(doc, "\"O app escaneia as redes Wi-Fi ao redor e mostra num gráfico. Cada rede aparece como uma \"montanha\". Quanto mais alta, mais forte. A posição na horizontal mostra em qual canal ela está.\"", italic=True)

    add_heading(doc, "A atividade — 15 min", level=2)
    add_para(doc, "Duplas. Cada dupla escreve na folha do aluno:")
    add_bullet(doc, "1. Quantas redes vocês estão vendo? (Total)")
    add_bullet(doc, "2. Qual é a mais forte? (Nome e força em dBm)")
    add_bullet(doc, "3. Qual é a mais fraca?")
    add_bullet(doc, "4. Quantas estão no canal 1? No canal 6? No canal 11? (São os 3 canais principais do 2,4 GHz)")
    add_bullet(doc, "5. Tem alguma rede em 5 GHz? (Aparecem em canais maiores, tipo 36, 40, 149)")
    add_bullet(doc, "6. Qual rede é a da escola? Que canal ela está usando?")
    add_para(doc, "Depois, a dupla anda pela sala. Vê como a força muda quando aproxima ou afasta do roteador. Anota.")

    add_heading(doc, "Discussão coletiva — 5 min", level=2)
    add_para(doc, "\"Por que tem tanta rede? O que atrapalha se todo mundo tá no canal 6?\"", italic=True)
    add_para(doc, "Aceite as respostas. Confirme: no 2,4 GHz só existem 3 canais que não se atrapalham (1, 6 e 11). Se dez roteadores estão todos no canal 6, todos ficam mais lentos. Isso volta com força na Noite 10 (cobertura de Wi-Fi).")

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — PRÁTICA · 4 ESTUDOS DE CASO · 25 MIN")

    add_heading(doc, "Instrução — 2 min", level=2)
    add_para(doc, "\"Cada equipe vai receber quatro cenários. Pra cada um: qual meio usar (cabo, fibra, Wi-Fi), qual categoria de cabo se for cabo, e por quê. Sem inventar equipamento que ainda não vimos.\"", italic=True)

    add_heading(doc, "Os quatro cenários — 15 min de trabalho", level=2)

    add_para(doc, "Cenário 1 — Casa térrea, 80 m², sala com roteador", bold=True)
    add_para(doc, "Cliente quer: TV com Netflix na sala (2 m do roteador); notebook do filho no quarto (10 m com uma parede); Wi-Fi no quintal pra celular. Como conectar cada aparelho?")

    add_para(doc, "Cenário 2 — Sobrado, roteador no térreo, escritório em cima", bold=True)
    add_para(doc, "Cliente reclama que o Wi-Fi do escritório é ruim. Fica 8 m acima e uma laje entre. Ele quer PC parado no escritório. Qual solução?")

    add_para(doc, "Cenário 3 — Comércio pequeno com 4 caixas registradoras", bold=True)
    add_para(doc, "Loja de 200 m². Sinal do provedor entra na sala de estoque, no fundo. As 4 caixas ficam na frente. Cliente quer conexão estável (não pode cair no meio da venda). Rede sem fio ou cabo? Se cabo, que categoria? Quantos metros?")

    add_para(doc, "Cenário 4 — Galpão de 40 m × 30 m, sem paredes internas", bold=True)
    add_para(doc, "Escritório numa ponta, área de estoque no resto. Precisa de Wi-Fi cobrindo tudo. Como resolver? (Dica: uma solução só não resolve — precisa combinar meios.)")

    add_heading(doc, "Apresentação — 8 min", level=2)
    add_para(doc, "Cada equipe apresenta 1 cenário (sorteio). Turma pergunta. Você fecha com \"resposta que costuma dar certo\" — não é ÚNICA, mas é comum.")

    add_para(doc, "Respostas típicas (pra você conferir):", bold=True, color=COR_PRIMARIA)
    add_bullet(doc, "Cenário 1: TV com cabo (mais estável), notebook com Wi-Fi, quintal precisa ponto de acesso extra")
    add_bullet(doc, "Cenário 2: passar cabo até o escritório (é PC parado); se não puder passar cabo, ponto de acesso ligado por cabo até onde o cabo alcança")
    add_bullet(doc, "Cenário 3: cabo Cat5e ou Cat6, ~30-40 m dependendo da planta; passar por eletroduto embutido se possível")
    add_bullet(doc, "Cenário 4: cabo Cat6 do escritório até um ou dois pontos de acesso no meio do galpão; Wi-Fi único não cobre 40 m com aparelho doméstico")

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — DUELO RELÂMPAGO · 15 MIN")
    add_para(doc, "Divide a turma em dois times. Alterna. Pergunta rápida, resposta em 5 segundos. Ponto pra quem acertar. Empate = pergunta bônus.")

    add_heading(doc, "12 perguntas de reserva", level=2)
    add_bullet(doc, "Cabo Ethernet Cat5e aguenta até quantos metros? — 100 m")
    add_bullet(doc, "Fibra sofre interferência de motor? — Não")
    add_bullet(doc, "2,4 GHz atravessa parede melhor ou pior que 5 GHz? — Melhor")
    add_bullet(doc, "Qual o conector do cabo par trançado? — RJ45")
    add_bullet(doc, "Qual o conector do cabo de telefone? — RJ11")
    add_bullet(doc, "Fibra é feita de metal? — Não, vidro")
    add_bullet(doc, "Se o cabo destrançar demais na ponta, o que perde? — Qualidade do sinal")
    add_bullet(doc, "Cat6a ou Cat5e é mais adequado pra casa comum? — Cat5e (o outro é overkill)")
    add_bullet(doc, "Cabo blindado é pra que ambiente? — Muita interferência elétrica")
    add_bullet(doc, "Wi-Fi passa por parede de concreto igual passa por madeira? — Não, perde mais no concreto")
    add_bullet(doc, "Onde a metragem está impressa no cabo? — Na capa, a cada metro")
    add_bullet(doc, "Emenda de fibra em residência: quem faz? — Provedor / técnico especializado")

    add_hr(doc)

    # BLOCO 9
    add_heading(doc, "BLOCO 9 — FECHAMENTO E TAREFA · 15 MIN")

    add_heading(doc, "Quiz de saída", level=2)
    add_bullet(doc, "Como você reconhece um cabo par trançado sem abrir a capa?")
    add_bullet(doc, "Qual é a diferença entre Cat5e e Cat6, na prática?")
    add_bullet(doc, "Fibra ou cabo pra atender uma sala 8 m de distância?")
    add_bullet(doc, "Wi-Fi de 5 GHz é melhor ou pior que 2,4 GHz? Depende do quê?")
    add_bullet(doc, "Quando o cliente diz \"o vizinho tá roubando meu Wi-Fi\" — o que a gente aprendeu hoje ajuda a entender?")

    add_heading(doc, "A tarefa", level=2)
    add_para(doc, "\"Nessa semana, prestem atenção em três cabos que vocês encontrarem em casa, no trabalho ou na rua. Fotografem a capa se der. Anotem:\"", italic=True)
    add_bullet(doc, "Que tipo de cabo é")
    add_bullet(doc, "Marca e categoria (se conseguirem ler)")
    add_bullet(doc, "Onde ele estava ligado")
    add_para(doc, "\"Bônus: se acharem um cabo com o conector quebrado, tragam ele. Vai virar exemplo na Noite 4.\"", italic=True, bold=True)

    add_heading(doc, "Anuncie a próxima", level=2)
    add_para(doc, "\"Na próxima noite, o alicate entra em cena. Cada um vai crimpar seu próprio cabo, do zero. Vamos crimpar, testar e ver quem fez direito. Peço: usem manga curta ou dobrem a manga, porque a bancada suja o braço.\"", italic=True)

    add_hr(doc)

    # SE PERGUNTAREM
    add_heading(doc, "SE PERGUNTAREM")

    add_para(doc, "\"Cabo verde é diferente do azul?\"", bold=True)
    add_para(doc, "> \"A cor da capa é só pra identificação visual, não muda a função. Empresa às vezes usa azul pra dados e amarelo pra telefonia — é convenção interna. Mas o cabo por dentro é igual.\"", italic=True)

    add_para(doc, "\"Por que dizem que Cat7 é melhor?\"", bold=True)
    add_para(doc, "> \"Aguenta velocidade maior, tem blindagem mais complexa. Mas em rede doméstica ninguém aproveita. É mais caro pra pouco ganho. Só usa em ambiente muito específico.\"", italic=True)

    add_para(doc, "\"Fibra tem no supermercado?\"", bold=True)
    add_para(doc, "> \"Existe fibra caseira pré-conectorizada, mas o cliente comum não precisa mexer com fibra. Se você precisou emendar fibra em casa, algo tá muito errado ou tá fora do trabalho de campo padrão.\"", italic=True)

    add_para(doc, "\"E a fibra que a operadora traz, tem cor?\"", bold=True)
    add_para(doc, "> \"A capa quase sempre é preta ou branca, e é bem mais fina que o par trançado. Mais leve também. Se você tocar, sente a diferença.\"", italic=True)

    add_para(doc, "\"Onda de rádio faz mal?\"", bold=True)
    add_para(doc, "> \"A potência que a gente lida (roteador, celular) é muito baixa. Estudos até hoje não mostram dano em uso normal. Mas essa pergunta chega no atendimento — respondam com calma, sem debate. Diga que sabem porque é uma pergunta comum.\"", italic=True)

    add_para(doc, "\"Cabo cor de rosa existe?\"", bold=True)
    add_para(doc, "> \"Existe cabo de rede em todas as cores possíveis. A função é a mesma. Serve pra o técnico identificar visualmente quando tem muito cabo no mesmo lugar.\"", italic=True)

    add_hr(doc)

    # DEPOIS DA AULA
    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "Quem separou cabo par trançado de coaxial sem hesitar? Quem ainda titubeou? (anote pra reforçar na Noite 4)")
    add_bullet(doc, "Quem instalou o Wi-Fi Analyzer e se envolveu com o scan? (esses ajudam a puxar a Noite 10 depois)")
    add_bullet(doc, "Se a apresentação dos cenários funcionou; se sim, repita nas próximas")
    add_bullet(doc, "Se algum caso real apareceu na conversa, guarde pra reaproveitar")

    add_hr(doc)

    # DESCRITIVO
    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Trabalhei os três meios de transmissão utilizados em redes de computadores — condutor metálico (par trançado), fibra óptica e propagação de sinais eletromagnéticos (redes sem fio) — com identificação de aplicações e limitações de cada um. Detalhei características do cabo par trançado (categorias Cat5e, Cat6 e Cat6a; blindagem; leitura de etiqueta e marcação; conector RJ45), com atividade prática de identificação de cabos em bancada. Apresentei fundamentos de fibra óptica e de propagação de ondas de rádio em faixas de 2,4 GHz e 5 GHz. Coordenei atividade prática de análise do espectro Wi-Fi local com uso de aplicativo em dispositivo móvel, incluindo identificação de redes vizinhas, canais em uso e níveis de sinal. Encerrei a aula com atividade de análise de quatro cenários de instalação, exigindo escolha justificada do meio de transmissão apropriado.")

    add_hr(doc)

    # CAIXA DE FERRAMENTAS DIDÁTICAS
    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo (ou pra guardar pra outra noite):", italic=True, color=COR_CINZA)

    add_para(doc, "1. Cabo cego", bold=True)
    add_para(doc, "Pegue um cabo bom e esconda um defeito (fio partido dentro, conector solto, decapado errado). A turma tenta descobrir só olhando e mexendo, sem tester. Ensina o olhar experiente. 15-20 min.")

    add_para(doc, "2. Novela do vizinho", bold=True)
    add_para(doc, "Conte um caso real (invente se precisar) de casa com Wi-Fi ruim porque 8 vizinhos estavam no canal 6. A turma sugere solução. Prepara terreno pra Noite 10. 10-15 min.")

    add_para(doc, "3. Tour Wi-Fi da escola", bold=True)
    add_para(doc, "Turma anda pela escola com o app aberto, anotando a força do sinal em cada lugar. Volta com o \"mapa de sinal\" da escola. Prepara terreno pra cobertura de casa. 20-25 min.")

    add_para(doc, "4. Três histórias verdadeiras", bold=True)
    add_para(doc, "Reserve três casos que você viu no trabalho e conte 5 min cada. Ancora a teoria em situação real e alonga a aula naturalmente. 15-20 min.")

    add_para(doc, "5. Prova cega de conectores", bold=True)
    add_para(doc, "Cabos com aparência quase igual mas conectores diferentes (RJ45 × RJ11 × conectores modulares antigos). A turma bate olho e classifica em 30 segundos por cabo. 10 min.")

    path = os.path.join(OUTDIR, "Roteiro - Redes 3 - Por onde o sinal viaja.docx")
    doc.save(path)
    return path


# ============================================================
# SLIDES
# ============================================================
def build_slides():
    prs = new_presentation()
    P, D, W, G, K, R, LBG = PPalette.P, PPalette.D, PPalette.W, PPalette.G, PPalette.K, PPalette.R, PPalette.LBG

    # 1: Capa
    add_slide_cover(prs, "NOITE 3", "POR ONDE", "O SINAL VIAJA",
                    "Cabo, fibra e ondas",
                    "Redes de Computadores · Curso FIC · CEJA Itapiranga – SC")

    # 2: Retomada
    s = add_slide(prs)
    add_title_block(s, "RETOMADA", "Os 5 equipamentos da Noite 2", "Jogo relâmpago: eu digo a frase, você diz o equipamento")
    add_text(s, [
        "\"Não cria rede, só tem uma porta de saída\"   →   Modem",
        "\"Cria rede e distribui pra casa toda\"   →   Roteador",
        "\"Aumenta o número de portas de cabo\"   →   Switch",
        "\"Leva Wi-Fi onde o roteador não alcança\"   →   Ponto de acesso",
        "\"Onde o cabo entra no computador\"   →   Placa de rede",
    ], 0.7, 3.0, 11.9, 3.5, size=18, color=K)
    add_footer(s, FOOTER)

    # 3: A pergunta da noite
    s = add_slide(prs)
    add_title_block(s, "A PERGUNTA DA NOITE", "Como a informação sai de um equipamento e chega no outro?")
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(0.7), PInches(3.0), PInches(11.9), PInches(3.5))
    box.fill.solid()
    box.fill.fore_color.rgb = W
    box.line.color.rgb = P
    box.line.width = Emu(20000)
    add_text(s, [
        "Existem só três meios de transmissão.",
        "Todo equipamento de rede que a gente viu",
        "se conecta com outro por um deles.",
        "",
        "Não existe um quarto.",
    ], 1.1, 3.3, 11.5, 3.0, size=22, color=K)
    add_footer(s, FOOTER)

    # 4: Os três meios
    s = add_slide(prs)
    add_title_block(s, "OS TRÊS MEIOS", "Cabo · Fibra · Onda")
    meios = [
        ("CABO", "Pulsos elétricos correndo dentro de fio de cobre", P),
        ("FIBRA", "Pulsos de luz correndo dentro de vidro", P),
        ("ONDA", "Sinal de rádio no ar", P),
    ]
    x = 0.7
    for tit, desc, color in meios:
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(x), PInches(3.0), PInches(4.0), PInches(3.0))
        card.fill.solid()
        card.fill.fore_color.rgb = color
        card.line.fill.background()
        add_text(s, tit, x, 3.6, 4.0, 0.7, size=30, bold=True, color=W, align=PP_ALIGN.CENTER)
        add_text(s, desc, x + 0.2, 4.6, 3.6, 1.5, size=14, color=W, align=PP_ALIGN.CENTER, italic=True)
        x += 4.1
    add_text(s, "Elétron. Luz. Onda. Três jeitos de mover a mesma informação.", 0.7, 6.4, 11.9, 0.5, size=17, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 5: Tabela comparativa
    s = add_slide(prs)
    add_title_block(s, "DE UMA OLHADA", "O que muda entre os três")
    rows = [
        ["Item", "Cabo (par trançado)", "Fibra", "Onda (Wi-Fi/celular)"],
        ["Alcance", "Até 100 m", "Km sem repetir", "Wi-Fi: metros. Celular: km"],
        ["Velocidade", "Alta", "Muito alta", "Depende"],
        ["Interferência", "Sofre", "Não sofre", "Sofre muito"],
        ["Custo", "Barato por metro", "Caro, mas caiu", "Sem 'por metro'"],
        ["Quem instala", "Você (Noite 4)", "Equipe do provedor", "Não instala — configura"],
    ]
    add_table_slide(s, rows, top=2.8, height=3.8, col_widths=[2.4, 3.2, 2.8, 3.5], size=13)
    add_footer(s, FOOTER)

    # 6: A decisão em 3 perguntas
    s = add_slide(prs)
    add_title_block(s, "A REGRA DE BOLSO", "Como escolher o meio no chamado")
    perguntas = [
        ("1", "Vai passar mais de 100 metros?", "→ Fibra"),
        ("2", "O aparelho fica parado no mesmo lugar?", "→ Cabo (mais estável)"),
        ("3", "O aparelho se mexe (celular, notebook)?", "→ Wi-Fi (é o único jeito)"),
    ]
    y = 2.8
    for num, perg, resp in perguntas:
        circ = s.shapes.add_shape(MSO_SHAPE.OVAL, PInches(0.7), PInches(y), PInches(0.9), PInches(0.9))
        circ.fill.solid()
        circ.fill.fore_color.rgb = P
        circ.line.fill.background()
        add_text(s, num, 0.7, y + 0.05, 0.9, 0.9, size=32, bold=True, color=W, align=PP_ALIGN.CENTER)
        add_text(s, perg, 1.9, y + 0.05, 11, 0.5, size=20, color=K)
        add_text(s, resp, 1.9, y + 0.55, 11, 0.5, size=18, bold=True, color=D, italic=True)
        y += 1.2
    add_text(s, "Vale em 80% dos casos.", 0.7, 6.5, 11.9, 0.4, size=16, color=G, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 7: Cabo = par trançado
    s = add_slide(prs)
    add_title_block(s, "O CABO EM DETALHE", "Nome técnico: PAR TRANÇADO", "Por que trançado?")
    add_text(s, [
        "Dentro tem 8 fios, em 4 pares",
        "  Cada par é trançado sobre si mesmo",
        "",
        "A trança protege o sinal",
        "  Contra a interferência do próximo par e do ambiente externo",
        "",
        "Não é enfeite",
        "  Se destrançar demais na ponta ao crimpar, o cabo perde qualidade",
    ], 0.7, 2.8, 11.9, 4.0, size=18, color=K)
    add_footer(s, FOOTER)

    # 8: Categorias
    s = add_slide(prs)
    add_title_block(s, "CATEGORIAS", "Cat5e, Cat6, Cat6a — o que muda")
    rows = [
        ["Categoria", "O que suporta", "Onde a gente vê", "Preço"],
        ["Cat5e", "Até 1 Gbps, 100 m", "Casa comum, obra antiga", "Mais barato"],
        ["Cat6", "1 Gbps sem esforço", "Instalação nova, escritório", "Um pouco mais caro"],
        ["Cat6a", "10 Gbps até 100 m", "Data center, prédio corporativo", "Bem mais caro"],
        ["Cat7, Cat8", "Especializado", "Quase nunca vê", "Caro"],
    ]
    add_table_slide(s, rows, top=2.8, height=3.5, col_widths=[2.4, 3.2, 4, 2.3])
    add_text(s, "Pra 99% dos chamados: Cat5e ou Cat6. Não venda Cat6a pra casa.", 0.7, 6.6, 11.9, 0.4, size=16, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 9: UTP × blindado
    s = add_slide(prs)
    add_title_block(s, "BLINDAGEM", "UTP × Blindado — quando cada um faz sentido")
    add_text(s, [
        "UTP (não blindado)",
        "  O comum. Barato. Mais fino. Mais fácil de crimpar.",
        "",
        "STP / FTP (blindado)",
        "  Tem folha metálica protegendo. Usa em ambiente com muita interferência",
        "  (indústria, perto de motor grande)",
    ], 0.7, 2.8, 11.9, 2.5, size=18, color=K)
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(0.7), PInches(5.5), PInches(11.9), PInches(1.0))
    box.fill.solid()
    box.fill.fore_color.rgb = W
    box.line.color.rgb = D
    box.line.width = Emu(20000)
    add_text(s, "Regra prática: se o cliente não tem torno mecânico no quintal, não precisa de blindado.", 0.9, 5.75, 11.5, 0.5, size=16, italic=True, color=K, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 10: Como ler a capa
    s = add_slide(prs)
    add_title_block(s, "A CAPA CONTA A HISTÓRIA", "Ler como quem lê rótulo de remédio")
    add_text(s, [
        "•  Marca (Furukawa, Nexans, marcas genéricas)",
        "•  Categoria (CAT5E, CAT6)",
        "•  Tipo (UTP, FTP)",
        "•  Certificação (CMR = uso interno, CMX = uso externo)",
        "•  Metragem impressa a cada metro",
    ], 0.7, 2.8, 11.9, 3.0, size=19, color=K)
    add_text(s, "O truque do técnico: subtrair a marca do início pela do fim. Sabe o comprimento sem esticar o cabo.", 0.7, 6.4, 11.9, 0.5, size=16, italic=True, color=D, bold=True, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 11: PRÁTICA 1 — Identificar cabos
    s = add_slide(prs)
    add_title_block(s, "PRÁTICA · 35 MIN", "Identificar cabos em bancada")
    add_text(s, [
        "1. Equipes de 3 ou 4 pessoas",
        "",
        "2. Cada equipe recebe (ou roda pelas) mesas com cabos",
        "",
        "3. Para cada cabo, anota na ficha:",
        "     — Tipo (par trançado? coaxial? telefone? energia?)",
        "     — Categoria (se par trançado)",
        "     — Blindagem, estado do conector, comprimento",
        "     — Onde ele seria adequado usar",
        "",
        "4. No fim, cada equipe apresenta 1 cabo pra turma inteira",
    ], 0.7, 2.8, 11.9, 4.0, size=16, color=K)
    add_footer(s, FOOTER)

    # 12: Fibra em detalhe
    s = add_slide(prs)
    add_title_block(s, "FIBRA ÓPTICA", "Um fio de vidro do tamanho de um cabelo")
    add_text(s, [
        "•  Um LED ou laser manda pulsos de luz por dentro",
        "•  Do outro lado, um sensor lê os pulsos",
        "•  Como é luz e não elétron, não sofre interferência",
        "•  Distância enorme: 40 km sem repetir",
        "•  Frágil: dobra demais e quebra. Emenda precisa de máquina.",
    ], 0.7, 2.8, 11.9, 3.5, size=19, color=K)
    add_text(s, "Vocês trabalham do roteador da casa pra dentro. Emenda de fibra é serviço do provedor.", 0.7, 6.4, 11.9, 0.5, size=15, italic=True, color=G, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 13: Ondas de rádio
    s = add_slide(prs)
    add_title_block(s, "ONDAS DE RÁDIO", "Wi-Fi, Bluetooth, celular e satélite — mesma família")
    add_text(s, [
        "•  Wi-Fi        onda curta, alcance de metros",
        "•  Bluetooth    onda mais curta, alcance de metros",
        "•  Celular      onda da antena da operadora, alcance de km",
        "•  Satélite     onda que vem do espaço, alcance planetário",
    ], 0.7, 2.8, 11.9, 2.5, size=18, color=K)
    add_text(s, "Todas são a mesma coisa: onda de rádio. O que muda é frequência e potência.", 0.7, 5.5, 11.9, 0.5, size=17, bold=True, color=P, align=PP_ALIGN.CENTER, italic=True)
    add_footer(s, FOOTER)

    # 14: 2,4 vs 5 GHz
    s = add_slide(prs)
    add_title_block(s, "AS DUAS FAIXAS DO WI-FI", "2,4 GHz × 5 GHz")
    rows = [
        ["", "2,4 GHz", "5 GHz"],
        ["Alcance", "Maior", "Menor"],
        ["Velocidade", "Menor", "Maior"],
        ["Atravessa parede", "Melhor", "Com dificuldade"],
        ["Disputa com vizinhos", "Muita", "Bem menos"],
        ["Aparelho antigo enxerga?", "Sim", "Nem sempre"],
    ]
    add_table_slide(s, rows, top=2.8, height=3.8, col_widths=[3.8, 4, 4.1], size=15)
    add_footer(s, FOOTER)

    # 15: PRÁTICA 2 — Wi-Fi Scan
    s = add_slide(prs)
    add_title_block(s, "PRÁTICA · 25 MIN", "Wi-Fi Scan com o próprio celular")
    add_text(s, [
        "App: Wi-Fi Analyzer (Android, grátis)",
        "",
        "Em duplas, anotem:",
        "     — Quantas redes vocês estão vendo",
        "     — A mais forte e a mais fraca (nome e força em dBm)",
        "     — Quantas no canal 1, no 6, no 11",
        "     — Tem alguma em 5 GHz?",
        "     — Qual é a rede da escola? Que canal?",
        "",
        "Depois: andem pela sala e vejam como a força muda.",
    ], 0.7, 2.8, 11.9, 4.2, size=15, color=K)
    add_footer(s, FOOTER)

    # 16: PRÁTICA 3 — Cenários
    s = add_slide(prs)
    add_title_block(s, "PRÁTICA · 25 MIN", "Escolher o meio pra 4 cenários reais")
    add_text(s, [
        "Cada equipe recebe os quatro cenários.",
        "",
        "Para cada um, respondem:",
        "     •  Qual meio usar (cabo, fibra, Wi-Fi)?",
        "     •  Se cabo, qual categoria?",
        "     •  Por quê?",
        "",
        "Depois, cada equipe apresenta 1 cenário (sorteio).",
    ], 0.7, 2.8, 11.9, 4.0, size=17, color=K)
    add_footer(s, FOOTER)

    # 17: Cenário 1
    s = add_slide(prs)
    add_title_block(s, "CENÁRIO 1", "Casa térrea, 80 m², sala com roteador")
    add_text(s, [
        "Cliente quer:",
        "",
        "     •  TV com Netflix na sala (2 m do roteador)",
        "     •  Notebook do filho no quarto (10 m, uma parede)",
        "     •  Wi-Fi no quintal pra celular",
        "",
        "Como conectar cada aparelho?",
    ], 0.7, 2.8, 11.9, 4.0, size=20, color=K)
    add_footer(s, FOOTER)

    # 18: Cenário 2
    s = add_slide(prs)
    add_title_block(s, "CENÁRIO 2", "Sobrado — roteador no térreo, escritório em cima")
    add_text(s, [
        "Cliente reclama:",
        "     •  Wi-Fi do escritório é ruim",
        "     •  Fica 8 m acima e uma laje entre",
        "     •  Ele quer PC parado no escritório",
        "",
        "Qual solução?",
    ], 0.7, 2.8, 11.9, 4.0, size=20, color=K)
    add_footer(s, FOOTER)

    # 19: Cenário 3
    s = add_slide(prs)
    add_title_block(s, "CENÁRIO 3", "Comércio pequeno com 4 caixas registradoras")
    add_text(s, [
        "Loja de 200 m².",
        "     •  Sinal do provedor entra na sala de estoque (fundo)",
        "     •  4 caixas ficam na frente da loja",
        "     •  Não pode cair no meio da venda",
        "",
        "Rede sem fio ou cabo? Se cabo, categoria? Metros?",
    ], 0.7, 2.8, 11.9, 4.0, size=19, color=K)
    add_footer(s, FOOTER)

    # 20: Cenário 4
    s = add_slide(prs)
    add_title_block(s, "CENÁRIO 4", "Galpão 40 × 30 m, sem paredes internas")
    add_text(s, [
        "     •  Escritório numa ponta",
        "     •  Área de estoque no resto",
        "     •  Precisa Wi-Fi cobrindo tudo",
        "",
        "Como resolver?",
        "",
        "Dica: uma solução só não resolve. Precisa combinar meios.",
    ], 0.7, 2.8, 11.9, 4.0, size=19, color=K)
    add_footer(s, FOOTER)

    # 21: Duelo relâmpago
    s = add_slide(prs)
    add_title_block(s, "DUELO RELÂMPAGO · 15 MIN", "Time A × Time B")
    add_text(s, [
        "•  Turma dividida em dois times",
        "•  Pergunta rápida, resposta em 5 segundos",
        "•  Alterna entre os times",
        "•  Ponto pra quem acertar",
        "•  Empate = pergunta bônus",
        "",
        "Regra: se a resposta demorar mais de 5 segundos, passa pro outro time.",
    ], 0.7, 2.8, 11.9, 4.0, size=19, color=K)
    add_footer(s, FOOTER)

    # 22: Resumo da noite
    s = add_slide(prs)
    add_title_block(s, "O RESUMO DA NOITE", "Quatro ideias que ficam")
    add_text(s, [
        "•  Existem só três meios de transmissão: cabo, fibra e ondas",
        "•  Cabo é padrão pra conexão fixa e curta. Fibra pra longa distância. Onda pra quem se mexe.",
        "•  Cat5e ou Cat6 resolve a casa comum. Não venda Cat6a sem motivo.",
        "•  A capa do cabo conta tudo — marca, categoria, tipo e metragem.",
    ], 0.7, 3.0, 11.9, 3.5, size=19, color=K)
    add_footer(s, FOOTER)

    # 23: Fechando
    s = add_slide(prs)
    add_title_block(s, "FECHANDO", "Cinco perguntas para conferir")
    add_text(s, [
        "1. Como você reconhece um cabo par trançado sem abrir a capa?",
        "2. Diferença entre Cat5e e Cat6, na prática?",
        "3. Fibra ou cabo pra uma sala 8 m de distância?",
        "4. 5 GHz é melhor ou pior que 2,4 GHz? Depende do quê?",
        "5. Cliente diz \"vizinho tá roubando meu Wi-Fi\". O que a aula ajuda a entender?",
    ], 0.7, 2.8, 11.9, 4.0, size=17, color=K)
    add_footer(s, FOOTER)

    # 24: O que fica
    s = add_slide(prs)
    add_title_block(s, "O QUE FICA DESTA NOITE", "Cinco coisas para levar")
    add_text(s, [
        "•  Três meios: cabo, fibra, ondas. Cada um resolve num lugar.",
        "•  Cabo par trançado: 8 fios, 4 pares. Não destrance na ponta.",
        "•  Cat5e ou Cat6 pra casa. Blindado só onde tem interferência de indústria.",
        "•  Wi-Fi tem duas faixas: 2,4 GHz (alcance) e 5 GHz (velocidade).",
        "•  A capa do cabo tem tudo que você precisa — marca, categoria e a metragem.",
    ], 0.7, 2.8, 11.9, 4.2, size=17, color=K)
    add_footer(s, FOOTER)

    # 25: Para a próxima noite
    s = add_slide(prs)
    add_title_block(s, "PARA A PRÓXIMA NOITE", "O alicate entra em cena")
    add_text(s, [
        "Tarefa desta semana:",
        "     •  Prestem atenção em três cabos que encontrarem em casa, trabalho ou rua",
        "     •  Fotografem a capa se der",
        "     •  Anotem tipo, marca, categoria e onde ele estava ligado",
        "",
        "Bônus:",
        "     Se acharem um cabo com conector quebrado, tragam.",
        "     Vai virar exemplo na Noite 4.",
        "",
        "Na Noite 4: cada um crimpa o próprio cabo. Manga curta ou dobrada — a bancada suja.",
    ], 0.7, 2.8, 11.9, 4.2, size=15, color=K)
    add_footer(s, FOOTER)

    path = os.path.join(OUTDIR, "Aula - Redes 3 - Por onde o sinal viaja.pptx")
    prs.save(path)
    return path


# ============================================================
# FOLHA DO ALUNO (3 páginas)
# ============================================================
def build_folha():
    doc = Document()
    setup_page(doc, margem=Cm(1.8))

    # PÁGINA 1 — Identificação de cabos
    cabecalho_documento(
        doc,
        "IDENTIFICAÇÃO DE CABOS",
        f"Redes de Computadores · {NOITE} · Curso FIC · CEJA Itapiranga · 1 de 3",
    )

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("Nome: ")
    r.font.size = Pt(11)
    r.bold = True
    r = p.add_run("_________________________________________________     Data: ____ / ____ / ______")
    r.font.size = Pt(11)

    add_para(doc, "Preencha uma ficha por cabo que passar pela sua equipe. Não abra os cabos — só olhe, leia a capa e compare com o que aprendeu na aula.", italic=True, size=10, color=COR_CINZA)

    for i in range(1, 5):
        add_heading(doc, f"CABO {i}", level=2)
        tbl = add_table(doc, [
            ["Item", "Sua resposta"],
            ["Tipo (par trançado, coaxial, telefone, energia)", ""],
            ["Como você reconheceu o tipo?", ""],
            ["Categoria (se par trançado)", ""],
            ["Blindagem (UTP ou blindado)", ""],
            ["Estado do conector", ""],
            ["Comprimento aproximado", ""],
            ["Onde seria adequado usar (2-3 opções)", ""],
        ], col_widths=[7, 10.5])
        for row in tbl.rows[1:]:
            row.height = Cm(0.7)
        doc.add_paragraph()

    # PÁGINA 2 — Wi-Fi Scan
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "WI-FI SCAN COM O CELULAR",
        f"Análise do espectro Wi-Fi · {NOITE} · 2 de 3",
    )

    add_para(doc, "Em duplas, com o app Wi-Fi Analyzer aberto no celular Android, preencha o que ver.", italic=True, size=10, color=COR_CINZA)

    add_heading(doc, "Dupla", level=2)
    p = doc.add_paragraph()
    r = p.add_run("Nomes: ")
    r.font.size = Pt(11)
    r.bold = True
    r = p.add_run("_________________________________________________________________________")
    r.font.size = Pt(11)

    add_heading(doc, "1  Panorama do que você vê", level=2)
    tbl = add_table(doc, [
        ["Item", "Valor"],
        ["Quantas redes no total?", ""],
        ["Nome da rede mais forte (SSID)", ""],
        ["Força dela (dBm)", ""],
        ["Nome da rede mais fraca", ""],
        ["Força dela (dBm)", ""],
        ["Quantas redes no canal 1?", ""],
        ["Quantas no canal 6?", ""],
        ["Quantas no canal 11?", ""],
        ["Tem alguma rede em 5 GHz? (canais >30)", ""],
        ["Qual é a rede da escola? Que canal?", ""],
    ], col_widths=[9, 8.5])
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)

    add_heading(doc, "2  Andando pela sala", level=2)
    add_para(doc, "Anote como a força do sinal da rede da escola muda quando você se afasta do roteador. Vá até um ponto distante, veja o número, volte perto do roteador, veja de novo. Anote a diferença.", italic=True, size=10, color=COR_CINZA)
    for _ in range(3):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("_" * 100)
        r.font.size = Pt(11)
        r.font.color.rgb = COR_CINZA

    add_heading(doc, "3  Discussão", level=2)
    add_para(doc, "Por que tem tantas redes no ar? O que atrapalha se todas estiverem no mesmo canal? Escreva com suas palavras.", size=11, bold=True)
    for _ in range(4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("_" * 100)
        r.font.size = Pt(11)
        r.font.color.rgb = COR_CINZA

    # PÁGINA 3 — Cenários + tarefa
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "ESCOLHER O MEIO CERTO",
        f"Estudos de caso · {NOITE} · 3 de 3",
    )

    add_para(doc, "Para cada cenário, sua equipe responde: qual meio (cabo, fibra ou Wi-Fi), qual categoria (se cabo), e uma linha justificando.", italic=True, size=10, color=COR_CINZA)

    cenarios = [
        ("Cenário 1 — Casa térrea, 80 m², sala com roteador",
         "TV Netflix na sala (2 m). Notebook no quarto (10 m, uma parede). Wi-Fi no quintal pra celular."),
        ("Cenário 2 — Sobrado, roteador no térreo, escritório em cima",
         "Wi-Fi do escritório ruim, 8 m acima, uma laje entre. PC parado no escritório."),
        ("Cenário 3 — Comércio pequeno, 4 caixas registradoras",
         "Loja 200 m². Provedor entra no fundo. Caixas na frente. Não pode cair na venda."),
        ("Cenário 4 — Galpão 40 × 30 m, sem paredes internas",
         "Escritório numa ponta. Estoque no resto. Wi-Fi precisa cobrir tudo."),
    ]
    for tit, desc in cenarios:
        add_heading(doc, tit, level=2)
        add_para(doc, desc, italic=True, size=10, color=COR_CINZA)
        p = doc.add_paragraph()
        r = p.add_run("Meio escolhido: ")
        r.font.size = Pt(11)
        r.bold = True
        r = p.add_run("____________________________________________________________")
        r.font.size = Pt(11)
        p = doc.add_paragraph()
        r = p.add_run("Categoria (se cabo): ")
        r.font.size = Pt(11)
        r.bold = True
        r = p.add_run("____________________________________________________")
        r.font.size = Pt(11)
        p = doc.add_paragraph()
        r = p.add_run("Por quê? ")
        r.font.size = Pt(11)
        r.bold = True
        for _ in range(2):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run("_" * 100)
            r.font.size = Pt(11)
            r.font.color.rgb = COR_CINZA

    doc.add_paragraph()
    add_heading(doc, "TAREFA DA SEMANA", level=2, color=COR_DESTAQUE)
    add_bullet(doc, "Fotografe (ou anote) três cabos que encontrar em casa, no trabalho ou na rua")
    add_bullet(doc, "Para cada um: tipo, marca (se der pra ler), categoria e onde ele estava ligado")
    add_bullet(doc, "Bônus: se achar um cabo com conector quebrado, traga — vai virar exemplo na próxima aula")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Traga esta folha preenchida · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Folha do Aluno - Redes 3 - Por onde o sinal viaja.docx")
    doc.save(path)
    return path


# ============================================================
# RESUMO DO ALUNO
# ============================================================
def build_resumo():
    doc = Document()
    setup_page(doc, margem=Cm(1.8))

    cabecalho_documento(
        doc,
        "POR ONDE O SINAL VIAJA",
        f"Resumo do aluno · {NOITE} · Redes de Computadores · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Existem só três meios de transmissão. Todo equipamento de rede se conecta com outro por um deles: cabo, fibra ou onda. Não existe um quarto. Escolher o meio certo é metade do trabalho — o técnico que não sabe escolher enfia cabo caro onde não precisa e passa Wi-Fi onde tinha que ter cabo.")

    add_heading(doc, "Os três meios", level=2)
    add_table(doc, [
        ["Meio", "Como funciona", "Alcance", "Onde é comum"],
        ["Cabo (par trançado)", "Pulsos elétricos no cobre", "Até 100 m", "Do roteador aos aparelhos fixos"],
        ["Fibra", "Pulsos de luz no vidro", "Km sem repetir", "Da rua até a casa"],
        ["Onda (Wi-Fi/celular)", "Sinal de rádio no ar", "Wi-Fi: metros. Celular: km.", "Aparelhos que se mexem"],
    ], col_widths=[3.5, 4, 3.5, 5])

    add_heading(doc, "A regra de bolso", level=2)
    add_bullet(doc, "Vai passar mais de 100 metros? → Fibra")
    add_bullet(doc, "Aparelho fica parado no mesmo lugar? → Cabo (mais estável)")
    add_bullet(doc, "Aparelho se mexe (celular, notebook)? → Wi-Fi (é o único jeito)")

    add_heading(doc, "Cabo de rede — o que saber", level=2)
    add_para(doc, "Nome técnico: par trançado. Tem 8 fios em 4 pares. Cada par é trançado sobre si mesmo pra proteger contra interferência. Se destrançar demais na ponta, o cabo perde qualidade.")

    add_heading(doc, "Categorias de cabo", level=2)
    add_table(doc, [
        ["Categoria", "Quando usar"],
        ["Cat5e", "Casa comum, obra antiga. Resolve 99% dos chamados."],
        ["Cat6", "Instalação nova, escritório. Um pouco mais caro, mas vale."],
        ["Cat6a", "Data center, prédio corporativo. Não venda pra casa comum."],
    ], col_widths=[3, 12.5])

    add_heading(doc, "Blindagem", level=2)
    add_para(doc, "UTP (não blindado) é o comum. Cabo blindado (STP, FTP) só faz sentido em ambiente com muita interferência elétrica (indústria, perto de motor grande). Se o cliente não tem torno mecânico no quintal, ele não precisa de blindado.")

    add_heading(doc, "Como ler a capa do cabo", level=2)
    add_bullet(doc, "Marca (Furukawa, Nexans, marcas genéricas)")
    add_bullet(doc, "Categoria (CAT5E, CAT6)")
    add_bullet(doc, "Tipo (UTP, FTP)")
    add_bullet(doc, "Certificação (CMR = uso interno, CMX = uso externo)")
    add_bullet(doc, "Metragem impressa a cada metro — subtraia o número do início pelo do fim e você sabe o comprimento sem esticar o cabo.")

    add_heading(doc, "Fibra em 3 frases", level=2)
    add_bullet(doc, "É vidro fininho, do tamanho de um cabelo. Manda luz por dentro.")
    add_bullet(doc, "Não sofre interferência de motor, transformador nem raio.")
    add_bullet(doc, "Frágil. Emenda é serviço do provedor, não do técnico de campo.")

    add_heading(doc, "Ondas de rádio", level=2)
    add_para(doc, "Wi-Fi, Bluetooth, celular e satélite são todas ondas de rádio — mudam só a frequência e a potência. Duas faixas de Wi-Fi que você precisa conhecer:")
    add_table(doc, [
        ["", "2,4 GHz", "5 GHz"],
        ["Alcance", "Maior", "Menor"],
        ["Velocidade", "Menor", "Maior"],
        ["Atravessa parede", "Melhor", "Com dificuldade"],
        ["Disputa com vizinhos", "Muita", "Menos"],
    ], col_widths=[4.5, 5.5, 5.5])

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Existem só três meios: cabo, fibra e onda. Não existe um quarto.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Cat5e ou Cat6 resolve a casa comum. Cat6a é overkill.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"A metragem impressa na capa é o que evita esticar cabo pra medir.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Como você reconhece um cabo par trançado sem abrir a capa?")
    add_bullet(doc, "Você precisa cobrir um galpão de 40 × 30 m. Uma solução só resolve?")
    add_bullet(doc, "Cliente reclama de Wi-Fi lento em uma rua cheia de casas. Que suspeita você tem?")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Este material é seu. Guarde e volte quando precisar. · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Resumo do Aluno - Redes 3 - Por onde o sinal viaja.docx")
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
