"""Gera os 4 arquivos da Noite 2 reformulada — Como a internet chega + os 5 equipamentos.

Reformulação do material original mantendo as frases-âncora que funcionaram
(caixinha, cinco equipamentos e só um cria rede, porta WAN erro nº 1, vizinho sem internet)
mas com o padrão prática-pesada: 50%+ do tempo em atividade do aluno,
role-play cliente-técnico, duelo relâmpago, e CAIXA DE FERRAMENTAS DIDÁTICAS.

Cobertura Wi-Fi movida para Noite 10 (só menção rápida aqui).
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
NOITE = "Noite 2"
FOOTER = f"Redes de Computadores · {NOITE}"


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "COMO A INTERNET CHEGA + OS 5 EQUIPAMENTOS",
        f"Roteiro do professor · Redes de Computadores · {NOITE} · Curso FIC · CEJA Itapiranga – SC",
    )

    # A IDEIA
    add_heading(doc, "A IDEIA DESTA NOITE")
    add_para(doc, "Na noite passada eles aprenderam a separar as palavras — rede não é internet. Hoje eles percorrem o caminho.")
    add_para(doc, "O cliente aponta para a caixinha na parede e diz que \"a internet está aí dentro\". Naquela caixinha não tem internet nenhuma — tem a ponta de um caminho muito longo. Esta noite é sobre esse caminho, e sobre os equipamentos que ficam na parte dele que é trabalho de vocês.")

    add_heading(doc, "O que se resolve nesta noite", level=2)
    add_para(doc, "Cinco equipamentos, e só um deles cria rede.")
    add_para(doc, "Quem confunde modem com roteador, ou acha que switch dá internet, vai errar orçamento e vai errar diagnóstico. E o erro mais caro da instalação doméstica — o cabo do provedor na porta errada — nasce exatamente dessa confusão.")

    add_heading(doc, "Esta é a segunda de quinze noites", level=2)
    add_para(doc, "Hoje ainda não se monta nada. Hoje se reconhece, se nomeia e se anota. E se atende: metade da aula é vocês respondendo a chamados fictícios em duplas. A montagem começa depois que todo mundo souber o nome das peças.")

    add_hr(doc)

    # ANTES DA AULA
    add_heading(doc, "ANTES DA AULA")

    add_heading(doc, "O material da mesa", level=2)
    add_table(doc, [
        ["Item", "Para quê"],
        ["Roteador (quantos houver)", "Achar a porta WAN, contar as LAN, ler a etiqueta"],
        ["Switch (se tiver)", "Comparar com o roteador e ver que não tem WAN"],
        ["Modem ou caixinha de fibra, se houver", "Mostrar a entrada da fibra"],
        ["Cabos de rede prontos", "Ligar e ver as luzes acenderem"],
        ["Régua de tomadas", "Para ligar os equipamentos na bancada"],
        ["Cartões de chamado (do bloco 5)", "Papel A5 com um chamado escrito em cada"],
    ], col_widths=[6, 10.5])
    add_para(doc, "Se você só tiver um roteador, a prática vira rodízio: cada equipe passa pela bancada por vez, e as demais preenchem a ficha com o que já viram.", italic=True, color=COR_CINZA)

    add_heading(doc, "Prepare antes", level=2)
    add_bullet(doc, "Ligue os equipamentos uma vez, sozinho, e confira quais luzes acendem")
    add_bullet(doc, "Se houver senha de administração trocada, não abra o painel hoje — isso é assunto da Noite 9")
    add_bullet(doc, "Tire uma foto da traseira de cada roteador, para o caso de a prática não fluir")
    add_bullet(doc, "Prepare os cartões de chamado do Bloco 5 (10 chamados, um por cartão) — modelos no fim deste roteiro")

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "A folha do aluno (3 páginas), uma por pessoa")
    add_bullet(doc, "O resumo do aluno, uma por pessoa (pra levar pra casa)")
    add_bullet(doc, "A ficha de identificação de equipamento faz parte da folha (página 3), uma por equipe")

    add_hr(doc)

    # QUADRO DE TEMPO
    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Retomada", "10 min", "O que ficou da Noite 1 e conferência de tarefa"],
        ["2. O caminho completo", "25 min", "Seis etapas, do servidor até o aparelho"],
        ["3. Os cinco equipamentos", "35 min", "Um de cada vez, com o objeto na mão"],
        ["INTERVALO", "15 min", ""],
        ["4. PRÁTICA — Identificação e ficha", "45 min", "Portas, etiquetas, luzes, apresentação por equipe"],
        ["5. PRÁTICA — Role-play cliente/técnico", "30 min", "Duplas atendem chamados e apontam o equipamento"],
        ["6. O método do técnico + cenários", "20 min", "A sequência de verificação em ação"],
        ["7. Duelo relâmpago", "15 min", "Time A × Time B — 12 perguntas"],
        ["8. Fechamento e tarefa", "15 min", "Consolidação e anúncio da Noite 3"],
    ], col_widths=[6, 2, 8.5])
    add_para(doc, "110 min de prática ativa (52% do total). Cobertura de Wi-Fi (posicionamento, canais) foi movida pra Noite 10 — aqui só menção rápida.", italic=True, color=COR_CINZA)
    add_para(doc, "Se atrasar: corte o Bloco 7 primeiro; se ainda precisar, corte o Bloco 6 pela metade. NUNCA corte a prática (blocos 4 e 5).", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — RETOMADA · 10 MIN")
    add_para(doc, "Não explique. Pergunte.")
    add_bullet(doc, "Wi-Fi é a mesma coisa que internet?")
    add_bullet(doc, "Dá para ter rede sem internet?")
    add_bullet(doc, "O que os dados móveis têm de diferente?")
    add_bullet(doc, "O que significa o ícone com sinal de exclamação?")

    add_heading(doc, "Se alguém disser a frase certa", level=2)
    add_para(doc, "— \"Wi-Fi é o caminho sem fio até o roteador\" — comemore. Repita alto. Vale mais vindo deles.")

    add_heading(doc, "Confira a tarefa da Noite 1", level=2)
    add_para(doc, "Pergunte quem trouxe os dados do roteador de casa ou as anotações dos ícones do celular. Anote no quadro duas ou três marcas diferentes que apareceram.")
    add_para(doc, "\"Repare que tem marca diferente aí. Muda o nome, muda a cara, mas todos fazem a mesma coisa. Hoje vocês vão entender qual coisa é essa.\"", italic=True)
    add_para(doc, "Se pouca gente trouxe, não faça disso um problema. Siga — os equipamentos da mesa resolvem.")

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — O CAMINHO COMPLETO · 25 MIN")

    add_heading(doc, "Abra pela cena do cliente", level=2)
    add_para(doc, "\"O cliente aponta para a caixinha na parede e diz: 'a internet está aí dentro'. Hoje a gente vai descobrir que ali não tem internet nenhuma. Ali tem a ponta de um caminho que começa do outro lado do mundo.\"", italic=True)

    add_heading(doc, "Desenhe no quadro — 12 min", level=2)
    add_para(doc, "Seis caixas, da esquerda para a direita:")
    add_para(doc, "servidor → backbone → provedor → a rua → sua casa → aparelho", bold=True, color=COR_PRIMARIA)
    add_para(doc, "Percorra uma por uma:")

    add_para(doc, "1. O servidor", bold=True)
    add_para(doc, "\"É um computador num prédio, ligado na tomada, guardando o site que você abriu. Não é nuvem, não é mágica: é máquina.\"", italic=True)

    add_para(doc, "2. O backbone", bold=True)
    add_para(doc, "\"As fibras de longa distância, que ligam cidades, estados e países. É por onde passa o volume grande. Boa parte disso está no fundo do mar.\"", italic=True)

    add_para(doc, "3. O provedor", bold=True)
    add_para(doc, "\"A empresa da sua cidade. Ela compra acesso a esse caminho grande e revende para os moradores. É ela que aparece na sua fatura.\"", italic=True)

    add_para(doc, "4. A rua", bold=True)
    add_para(doc, "\"Da central do provedor sai uma fibra grossa, com muitos fios dentro. Ela percorre a cidade pelos postes. Em caixas de emenda, essa fibra é aberta e dividida. De lá sai uma fibra fina só sua, que atravessa até a sua parede. Essa última parte o pessoal chama de drop.\"", italic=True)

    add_para(doc, "5. A sua casa", bold=True)
    add_para(doc, "\"A caixinha recebe a fibra. O roteador distribui.\"", italic=True)

    add_para(doc, "6. O aparelho", bold=True)
    add_para(doc, "\"Por cabo ou por Wi-Fi, chega até você.\"", italic=True)

    add_heading(doc, "A divisão que importa — 8 min", level=2)
    add_para(doc, "Trace uma linha no quadro entre a etapa 3 e a 4:")
    add_para(doc, "\"Etapas 1, 2 e 3 são com o provedor. Quando o problema está aí, não adianta reiniciar roteador nenhum. Etapas 4, 5 e 6 são o seu trabalho. Saber de que lado da linha está o problema é metade do atendimento.\"", italic=True, color=COR_PRIMARIA)

    add_para(doc, "E o exemplo que fixa:")
    add_para(doc, "\"É por isso que o vizinho pode estar sem internet e você não: cada casa tem o seu drop, saindo da mesma caixa no poste.\"", italic=True, bold=True)

    add_heading(doc, "O tamanho da fibra — 5 min", level=2)
    add_para(doc, "Se você tiver um pedaço de fibra, passe pela turma.")
    add_para(doc, "\"Fibra é fina como um fio de cabelo, e é vidro. Dobrou demais, quebrou. Por isso a emenda precisa de equipamento próprio e quem faz é a equipe do provedor. Vocês não vão emendar fibra — vocês vão trabalhar do roteador para dentro.\"", italic=True)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — OS CINCO EQUIPAMENTOS · 35 MIN")
    add_para(doc, "Com o objeto na mão. Cada equipamento circula pela turma enquanto você fala.")

    add_heading(doc, "Os cinco, no quadro", level=2)
    add_table(doc, [
        ["Equipamento", "O que faz, em uma frase"],
        ["Modem", "Converte o sinal do provedor no sinal que a rede entende"],
        ["Roteador", "Cria a rede da casa e decide para onde cada informação vai"],
        ["Switch", "Aumenta o número de portas de cabo. Não cria rede"],
        ["Ponto de acesso", "Leva o sinal sem fio onde o roteador não alcança"],
        ["Placa de rede", "Onde o cabo entra no computador"],
    ], col_widths=[4, 12.5])

    add_heading(doc, "Um de cada vez — 6 min para cada", level=2)

    add_para(doc, "Modem", bold=True)
    add_para(doc, "\"É o tradutor. Pega a luz da fibra, ou o sinal do cabo, e entrega algo que o roteador entende. Tem uma saída só. Sozinho, ele atende um aparelho e mais nada. Não cria rede.\"", italic=True)

    add_para(doc, "Roteador", bold=True)
    add_para(doc, "\"É o distribuidor, e é ele que cria a rede da casa. Recebe uma conexão e divide entre todos os aparelhos, por cabo e sem fio. Ele guarda quem é quem na rede e manda cada resposta para o aparelho certo. E atenção: ele tem duas senhas diferentes — a da rede sem fio e a de administração do aparelho.\"", italic=True)

    add_para(doc, "Switch", bold=True)
    add_para(doc, "\"É um multiplicador de portas. E só isso. Transforma uma porta em quatro, oito, dezesseis. Não cria rede, não distribui internet sozinho e não tem Wi-Fi. Você usa quando faltou porta no roteador.\"", italic=True)
    add_para(doc, "Mostre os dois lado a lado:")
    add_para(doc, "\"Olhem a traseira dos dois. O roteador tem uma porta diferente das outras. O switch tem todas iguais. Isso já conta o que cada um faz.\"", italic=True, color=COR_PRIMARIA)

    add_para(doc, "Ponto de acesso", bold=True)
    add_para(doc, "\"Cria mais um ponto de Wi-Fi ligado na mesma rede da casa. O ideal é levar cabo até ele: assim o sinal não perde força. O repetidor faz parecido, mas sem cabo — repete o que recebe pelo ar, e perde velocidade no caminho. Casa comprida, dois andares, galpão nos fundos: é aqui que entra.\"", italic=True)

    add_para(doc, "Placa de rede", bold=True)
    add_para(doc, "\"Vocês já viram na montagem: é onde o cabo entra, na traseira do gabinete. Tem duas luzinhas. Uma acesa fixa diz que tem cabo com sinal; a outra pisca quando há tráfego. Se queimar, resolve com uma placa USB, que custa pouco.\"", italic=True)

    add_heading(doc, "O quadro que fecha o bloco", level=2)
    add_table(doc, [
        ["Equipamento", "Cria rede?", "Tem Wi-Fi?"],
        ["Modem", "Não", "Não"],
        ["Roteador", "Sim", "Quase sempre"],
        ["Switch", "Não", "Não"],
        ["Ponto de acesso", "Não", "Sim"],
        ["Placa de rede", "Não", "Depende"],
    ], col_widths=[5, 4, 7.5])
    add_para(doc, "\"Só um deles cria rede. Guardem isso.\"", italic=True, bold=True, color=COR_PRIMARIA)

    add_heading(doc, "E a caixinha do provedor", level=2)
    add_para(doc, "\"Hoje o provedor entrega uma caixa só, que faz três coisas juntas: recebe a fibra, distribui a rede e emite o Wi-Fi. É mais barato e o técnico instala mais rápido. O problema é que, se ela der defeito, para tudo de uma vez — e o Wi-Fi dela costuma ser fraco. Muita casa resolve colocando um roteador melhor depois dela.\"", italic=True)

    add_hr(doc)
    add_heading(doc, "INTERVALO · 15 MIN", color=COR_CINZA)
    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · IDENTIFICAÇÃO E FICHA · 45 MIN")

    add_heading(doc, "O que cada equipe faz", level=2)
    add_bullet(doc, "Identificar cada equipamento na mesa e dizer em voz alta o que ele faz")
    add_bullet(doc, "Achar a porta WAN de cada roteador e explicar como reconheceu")
    add_bullet(doc, "Contar as portas LAN e anotar")
    add_bullet(doc, "Ler a etiqueta e preencher a ficha (na página 3 da folha do aluno): nome de rede, endereço de configuração, modelo, série")
    add_bullet(doc, "Ligar o roteador e observar as luzes acendendo, uma por uma")
    add_bullet(doc, "Trocar de bancada com outra equipe e conferir o que o colega anotou")
    add_bullet(doc, "Ao fim: cada equipe apresenta 1 equipamento em 30 segundos pra turma toda")
    add_para(doc, "\"Ninguém configura nada hoje. Hoje é reconhecer, nomear e anotar.\"", italic=True, color=COR_PRIMARIA)

    add_heading(doc, "A porta WAN — o erro número um", level=2)
    add_para(doc, "Chame a turma inteira ao redor de um roteador antes de soltar as equipes:")
    add_para(doc, "\"Repare: uma porta é separada das outras, quase sempre de outra cor, e costuma vir escrito WAN ou Internet. É por ela que entra o cabo do provedor. As outras, do mesmo tamanho, são para os aparelhos da casa.\"", italic=True)
    add_para(doc, "E o alerta:")
    add_para(doc, "\"Se você ligar o cabo do provedor numa porta LAN, tudo vai conectar na rede normalmente — e nada vai abrir. Parece defeito, e não é. É o erro mais comum de quem instala roteador.\"", italic=True, color=COR_ALERTA, bold=True)

    add_heading(doc, "A etiqueta embaixo", level=2)
    add_table(doc, [
        ["O que tem na etiqueta", "Para que serve"],
        ["Nome da rede (SSID) de fábrica", "É o nome que aparece na lista antes de alguém trocar"],
        ["Senha do Wi-Fi de fábrica", "Conectar na primeira vez"],
        ["Usuário e senha de administração", "Entram na configuração. São outra senha"],
        ["Endereço de configuração", "O número que se digita no navegador"],
        ["Modelo e número de série", "O que o suporte do provedor vai pedir"],
    ], col_widths=[6.5, 10])
    add_para(doc, "\"Tire uma foto da etiqueta antes de pendurar o aparelho na parede. Quem já instalou sabe por quê.\"", italic=True, color=COR_PRIMARIA)

    add_heading(doc, "As luzes", level=2)
    add_table(doc, [
        ["A luz", "Acesa fixa", "Piscando", "Apagada"],
        ["Power", "Ligado", "—", "Sem energia ou fonte com defeito"],
        ["Fibra / Internet", "Sinal chegando", "Tentando conectar", "Sem sinal do provedor"],
        ["WLAN / Wi-Fi", "Wi-Fi ligado", "Tráfego passando", "Wi-Fi desligado no aparelho"],
        ["LAN 1 a 4", "Cabo conectado", "Dados passando", "Nada ligado nessa porta"],
    ], col_widths=[3.5, 3.8, 3.8, 5.4])
    add_para(doc, "\"Luz de fibra apagada: nem adianta mexer no computador. O problema está antes do roteador.\"", italic=True, color=COR_PRIMARIA, bold=True)
    add_para(doc, "O nome das luzes muda de marca para marca. A lógica é sempre essa.", italic=True, color=COR_CINZA)

    add_heading(doc, "Enquanto circulam", level=2)
    add_para(doc, "Passe de equipe em equipe e faça uma pergunta por bancada. Não corrija de imediato: deixe o colega da equipe corrigir primeiro.")

    add_hr(doc)

    # BLOCO 5 - NOVO Role-play
    add_heading(doc, "BLOCO 5 — PRÁTICA · ROLE-PLAY CLIENTE/TÉCNICO · 30 MIN")

    add_heading(doc, "A ideia — 2 min", level=2)
    add_para(doc, "\"Vocês vão fazer duplas. Um faz o cliente, outro faz o técnico. Eu vou dar um cartão de chamado pra dupla. O técnico não pode ver o cartão. Ele tem que descobrir qual equipamento está com problema perguntando pro cliente.\"", italic=True)

    add_heading(doc, "Rodada 1 — 12 min", level=2)
    add_para(doc, "Cada dupla recebe 1 cartão. O cliente lê em silêncio, sem mostrar. Técnico faz 3-4 perguntas e responde: qual equipamento provavelmente está com problema? O que ele investigaria primeiro?")
    add_para(doc, "Você circula. Não corrige na hora — deixa o técnico chegar (ou não).")

    add_heading(doc, "Troca de papéis — 12 min", level=2)
    add_para(doc, "As duplas trocam papéis e recebem outro cartão. Repetem o exercício.")

    add_heading(doc, "Discussão coletiva — 4 min", level=2)
    add_para(doc, "3 ou 4 duplas contam qual chamado receberam e qual foi o diagnóstico. Você fecha com o \"pulo do gato\" — o detalhe que separa o técnico bom.")

    add_heading(doc, "Os 10 cartões de chamado", level=2)
    add_para(doc, "Escreva cada um em uma tirinha de papel. Diga ao cliente pra representar o cliente de verdade — pessoa comum, sem vocabulário técnico.", italic=True, color=COR_CINZA)

    add_para(doc, "1. \"Instalei um switch em casa mas nada conecta na internet.\"", bold=True)
    add_para(doc, "Diagnóstico: switch NÃO cria rede. Precisa estar ligado a um roteador (que tem internet). Cliente ligou o switch direto no cabo do provedor.")

    add_para(doc, "2. \"O Wi-Fi tá aceso no meu celular mas nada abre.\"", bold=True)
    add_para(doc, "Diagnóstico: rede local funciona, internet do provedor não. Luz de fibra do roteador provavelmente apagada. Etapas 1-3 do caminho.")

    add_para(doc, "3. \"Comprei um roteador novo. Não sei em que porta ligar o cabo que o técnico do provedor deixou.\"", bold=True)
    add_para(doc, "Diagnóstico: porta WAN. Reconhece pela cor diferente e pela palavra WAN ou Internet impressa.")

    add_para(doc, "4. \"A impressora do meu escritório não imprime, mas a internet funciona.\"", bold=True)
    add_para(doc, "Diagnóstico: problema é na rede local entre computador e impressora. Não envolve internet nem provedor.")

    add_para(doc, "5. \"A luz de fibra do meu roteador tá apagada. Reiniciei três vezes.\"", bold=True)
    add_para(doc, "Diagnóstico: problema antes do roteador. Cabo do provedor, poste ou provedor mesmo. Chamar o provedor.")

    add_para(doc, "6. \"Meu vizinho me disse que a internet dele funciona. A minha caiu. Não é do provedor?\"", bold=True)
    add_para(doc, "Diagnóstico: cada casa tem seu drop saindo da mesma caixa no poste. Vizinho ter internet não garante que o provedor tá 100% — mas indica que o problema pode ser da drop dele especificamente.")

    add_para(doc, "7. \"Instalei um roteador novo, o Wi-Fi conecta mas nenhum aparelho abre página.\"", bold=True)
    add_para(doc, "Diagnóstico: cabo do provedor foi ligado na porta LAN em vez da WAN. Erro nº 1.")

    add_para(doc, "8. \"Meu computador não conecta na rede por cabo. A luzinha da placa não pisca.\"", bold=True)
    add_para(doc, "Diagnóstico: placa de rede não tá recebendo sinal. Verificar cabo, verificar porta do roteador, verificar se placa não queimou.")

    add_para(doc, "9. \"O escritório fica atrás de duas paredes. O Wi-Fi lá é péssimo.\"", bold=True)
    add_para(doc, "Diagnóstico: problema de alcance. Solução mais robusta é ponto de acesso ligado por cabo. Repetidor perde velocidade.")

    add_para(doc, "10. \"Meu técnico disse que precisa trocar o modem porque a internet tá lenta.\"", bold=True)
    add_para(doc, "Diagnóstico: modem só converte sinal — se a internet tá lenta, pode ser Wi-Fi ruim, plano baixo ou problema do provedor. Trocar modem raramente resolve lentidão. Alertar cliente.")

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — O MÉTODO DO TÉCNICO + CENÁRIOS · 20 MIN")

    add_heading(doc, "A sequência — 8 min", level=2)
    add_para(doc, "Sempre na mesma ordem. Método evita retrabalho e evita troca desnecessária.")
    add_bullet(doc, "1. A luz de fibra ou internet está acesa? Se não, o problema é antes do roteador")
    add_bullet(doc, "2. Os cabos estão firmes, com clique? Cabo meio solto é campeão de chamado")
    add_bullet(doc, "3. O cabo do provedor está na porta WAN, e não numa LAN?")
    add_bullet(doc, "4. Outro aparelho conecta? Se só um falha, o problema é no aparelho")
    add_bullet(doc, "5. O aparelho conecta na rede mas não abre página? Aí é sinal do provedor")
    add_bullet(doc, "6. Desligar da tomada por dez segundos e ligar resolveu? Registre que resolveu")
    add_para(doc, "\"Registrar que resolveu importa tanto quanto resolver. Se acontecer de novo em uma semana, o registro é o que mostra que não é coincidência.\"", italic=True, color=COR_PRIMARIA)

    add_heading(doc, "Encene um chamado — 12 min", level=2)
    add_para(doc, "Escolha dois alunos: um é o cliente, outro é o técnico. Você sussurra o defeito para o cliente (por exemplo: \"cabo do provedor numa porta LAN\"). O técnico precisa achar fazendo perguntas na ordem do método.")
    add_para(doc, "A turma assiste e ajuda quando o técnico trava. Funciona muito melhor do que explicar.")
    add_para(doc, "Se der tempo, faça 2 encenações — muda o defeito e muda a dupla.")

    add_para(doc, "Defeitos sugeridos pra sussurrar:")
    add_bullet(doc, "\"Cabo do provedor ligado na porta LAN 3\"")
    add_bullet(doc, "\"Fonte do roteador tá com mau contato — power fica piscando\"")
    add_bullet(doc, "\"Cabo do computador tá firmando na placa mas não no roteador\"")
    add_bullet(doc, "\"Modem funciona, mas o roteador ligado nele foi resetado e perdeu configuração\"")

    add_hr(doc)

    # BLOCO 7 NOVO
    add_heading(doc, "BLOCO 7 — DUELO RELÂMPAGO · 15 MIN")
    add_para(doc, "Divide a turma em dois times. Alterna. Resposta em 5 segundos.")

    add_heading(doc, "12 perguntas de reserva", level=2)
    add_bullet(doc, "Modem cria rede? — Não")
    add_bullet(doc, "Qual dos 5 equipamentos cria rede? — Roteador")
    add_bullet(doc, "Switch tem Wi-Fi? — Não")
    add_bullet(doc, "Ponto de acesso é diferente de repetidor em quê? — Ponto de acesso liga por cabo, repetidor pega no ar")
    add_bullet(doc, "Qual é a diferença visual entre roteador e switch olhando a traseira? — Roteador tem uma porta diferente (WAN)")
    add_bullet(doc, "Onde a etiqueta com dados de fábrica geralmente fica? — Embaixo do aparelho")
    add_bullet(doc, "Quantas senhas diferentes um roteador doméstico tem? — Duas (Wi-Fi e administração)")
    add_bullet(doc, "Cliente diz \"luz de fibra apagada\". Onde está o problema? — Antes do roteador (provedor)")
    add_bullet(doc, "Etapas 1, 2 e 3 do caminho são responsabilidade de quem? — Provedor")
    add_bullet(doc, "Etapa 5 (a sua casa) é responsabilidade de quem? — Do técnico de campo")
    add_bullet(doc, "Cabo do provedor entra em qual porta? — WAN (ou Internet)")
    add_bullet(doc, "Qual é o erro nº 1 de quem instala roteador? — Ligar o cabo do provedor numa porta LAN")

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — FECHAMENTO E TAREFA · 15 MIN")

    add_heading(doc, "Quiz de saída", level=2)
    add_bullet(doc, "Qual é a diferença entre modem e roteador?")
    add_bullet(doc, "O switch cria rede? O que ele faz então?")
    add_bullet(doc, "Como você reconhece a porta WAN num roteador que nunca viu?")
    add_bullet(doc, "O que acontece se o cabo do provedor for ligado numa porta LAN?")
    add_bullet(doc, "A luz de fibra está apagada. Onde está o problema?")

    add_heading(doc, "A tarefa", level=2)
    add_para(doc, "\"Procurem um cabo de rede em casa e leiam o que está escrito na capa. Anotem a categoria, o fabricante e o número da metragem. Se não acharem cabo de rede, sirva qualquer cabo: o da TV, o de energia. Tragam anotado — a gente vai comparar as capas na aula.\"", italic=True)
    add_para(doc, "\"E quem não achar nenhum, sem problema: vai ter cabo de sobra na mesa.\"", italic=True)

    add_heading(doc, "Anuncie a próxima", level=2)
    add_para(doc, "\"Na próxima noite a gente vai ver por onde o sinal viaja: cabo, fibra e sem fio. E vocês vão aprender a ler a capa de um cabo como quem lê um rótulo de remédio.\"", italic=True)

    add_hr(doc)

    # SE PERGUNTAREM
    add_heading(doc, "SE PERGUNTAREM")

    add_para(doc, "\"Posso usar dois roteadores na mesma casa?\"", bold=True)
    add_para(doc, "> \"Pode, e é comum. Mas o segundo precisa ser configurado direito, senão cria duas redes separadas e os aparelhos ficam trocando de uma para outra. A gente vê isso na Noite 8 (planejar instalação).\"", italic=True)

    add_para(doc, "\"Repetidor presta?\"", bold=True)
    add_para(doc, "> \"Resolve quando não dá para passar cabo. Mas perde velocidade, porque ele conversa com o roteador e com o aparelho ao mesmo tempo. Se der para passar cabo até um ponto de acesso, fica muito melhor.\"", italic=True)

    add_para(doc, "\"Quanto custa um roteador bom?\"", bold=True)
    add_para(doc, "> \"Varia bastante. O que muda o preço é o número de antenas, se tem as duas faixas e a capacidade de aguentar muitos aparelhos. Na Noite 15 (projeto final) a gente compara modelos de verdade.\"", italic=True)

    add_para(doc, "\"O roteador do provedor é ruim?\"", bold=True)
    add_para(doc, "> \"Não é ruim, é básico. Atende bem um apartamento pequeno. Em casa grande, costuma faltar cobertura — e aí entra o serviço de vocês.\"", italic=True)

    add_para(doc, "\"Dá para esconder o nome da rede?\"", bold=True)
    add_para(doc, "> \"Dá, e a gente vê isso na Noite 11 (segurança). Mas adianta menos do que as pessoas pensam.\"", italic=True)

    add_hr(doc)

    # DEPOIS DA AULA
    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "Quem conseguiu achar a porta WAN sozinho, sem dica")
    add_bullet(doc, "Quem confundiu switch com roteador mesmo depois do bloco — precisa de retomada")
    add_bullet(doc, "Se a encenação do chamado funcionou; se sim, repita nas próximas noites")
    add_bullet(doc, "Quantos trouxeram a tarefa de casa, para calibrar o que pedir da próxima vez")
    add_bullet(doc, "Quais chamados do role-play (Bloco 5) mais fizeram efeito — guarde pra usar nas Noites 9 e 14")

    add_hr(doc)

    # DESCRITIVO
    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Trabalhei a infraestrutura de acesso à internet, percorrendo com os estudantes o trajeto da informação desde o servidor de destino até o equipamento do usuário, com identificação das etapas sob responsabilidade do provedor e das etapas sob responsabilidade do técnico de campo. Apresentei os equipamentos que compõem uma rede — modem, roteador, comutador, ponto de acesso e interface de rede — com manuseio dos equipamentos pelos estudantes e diferenciação funcional entre eles. Os estudantes identificaram a porta de entrada do provedor, as portas de distribuição, os dados da etiqueta de identificação e o significado dos indicadores luminosos, registrando as informações em ficha técnica. Coordenei atividade de simulação de atendimento em duplas (cliente/técnico) com cartões de chamado, exigindo do estudante que identificasse o equipamento responsável pela falha e propusesse a primeira verificação. Apresentei ainda a sequência padronizada de verificação para atendimento e realizei encenação coletiva de resolução de chamado.")

    add_hr(doc)

    # CAIXA DE FERRAMENTAS DIDÁTICAS
    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou pra guardar pra usar em outra noite:", italic=True, color=COR_CINZA)

    add_para(doc, "1. Autopsia de roteador", bold=True)
    add_para(doc, "Se você tiver um roteador antigo sem uso, abra ele na frente da turma. Mostre: antenas por dentro, placa, bateria de backup (se tiver), etiqueta interna. Vira memória permanente. 15-20 min.")

    add_para(doc, "2. Caça ao SSID pela rua", bold=True)
    add_para(doc, "Turma anda pela quadra da escola com o celular anotando os nomes de rede que aparecem. Volta e compara. Introduz conceito de canais Wi-Fi (Noite 10) e mostra que rede é onipresente. 25 min.")

    add_para(doc, "3. Etiqueta às cegas", bold=True)
    add_para(doc, "Você lê pra turma o que tem numa etiqueta (sem mostrar). Turma tem que apontar qual é o modelo, qual é a senha de Wi-Fi, qual é a de administração. Fixa a leitura de etiqueta. 10 min.")

    add_para(doc, "4. Sabotagem do roteador", bold=True)
    add_para(doc, "Você chega antes da aula e sabota (levemente) o roteador da mesa — troca a WAN de lugar, deixa cabo meio solto, desliga o Wi-Fi. Turma tem que descobrir. Ótima porta pra Noite 14 (diagnóstico). 20 min.")

    add_para(doc, "5. Compre este", bold=True)
    add_para(doc, "Traga catálogos ou fotos de 3 roteadores de preços diferentes. Turma escolhe qual comprar pra 3 cenários (casa pequena, casa grande, comércio). Prepara terreno pra orçamento na Noite 15. 15-20 min.")

    add_para(doc, "6. Testemunho real", bold=True)
    add_para(doc, "Se você conhece um técnico ativo, convide-o pra passar 15 min contando 2 chamados marcantes. O relato de campo vale por horas de teoria. Combine antes.")

    path = os.path.join(OUTDIR, "Roteiro - Redes 2 - Como a internet chega.docx")
    doc.save(path)
    return path


# ============================================================
# SLIDES
# ============================================================
def build_slides():
    prs = new_presentation()
    P, D, W, G, K, R, LBG = PPalette.P, PPalette.D, PPalette.W, PPalette.G, PPalette.K, PPalette.R, PPalette.LBG

    # 1: Capa
    add_slide_cover(prs, "NOITE 2", "COMO A INTERNET", "CHEGA ATÉ VOCÊ",
                    "Do cabo da rua até a tela do celular",
                    "Redes de Computadores · Curso FIC · CEJA Itapiranga – SC")

    # 2: Retomada
    s = add_slide(prs)
    add_title_block(s, "RETOMADA", "O que ficou da noite passada")
    add_text(s, [
        "1.  Wi-Fi é a mesma coisa que internet?",
        "2.  Dá para ter rede sem internet?",
        "3.  O que os dados móveis têm de diferente?",
        "4.  O que significa o ícone com sinal de exclamação?",
        "5.  Quem trouxe a ficha do roteador preenchida?",
    ], 0.7, 2.8, 11.9, 3.5, size=20, color=K)
    add_text(s, "Cinco minutos. Se a turma travar em alguma, retome ali mesmo antes de seguir.", 0.7, 6.5, 11.9, 0.4, size=14, italic=True, color=G, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 3: A pergunta da noite
    s = add_slide(prs)
    add_title_block(s, "A PERGUNTA DA NOITE", "De onde vem essa internet?")
    add_text(s, [
        "O cliente aponta para a caixinha na parede",
        "e diz: \"a internet está aí dentro\".",
    ], 0.7, 2.8, 11.9, 1.5, size=22, color=K, italic=True)
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(0.7), PInches(4.2), PInches(11.9), PInches(2.3))
    box.fill.solid()
    box.fill.fore_color.rgb = W
    box.line.color.rgb = D
    box.line.width = Emu(25000)
    add_text(s, [
        "A verdade",
        "Naquela caixinha não tem internet nenhuma.",
        "Tem só a ponta de um caminho muito longo.",
    ], 1.0, 4.4, 11.5, 2.0, size=20, color=K)
    add_footer(s, FOOTER)

    # 4: O caminho completo
    s = add_slide(prs)
    add_title_block(s, "O CAMINHO COMPLETO", "Seis etapas, sempre as mesmas")
    etapas = ["SERVIDOR", "BACKBONE", "PROVEDOR", "A RUA", "SUA CASA", "APARELHO"]
    descs = ["Máquina que guarda o site", "Fibras de longa distância", "Central na sua cidade",
             "Fibra no poste até você", "Caixinha e roteador", "Cabo ou Wi-Fi"]
    x = 0.7
    w = 1.95
    for i, (etapa, desc) in enumerate(zip(etapas, descs)):
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(x), PInches(2.9), PInches(w), PInches(2.6))
        card.fill.solid()
        card.fill.fore_color.rgb = P
        card.line.fill.background()
        add_text(s, str(i+1), x, 3.0, w, 0.5, size=14, bold=True, color=D, align=PP_ALIGN.CENTER)
        add_text(s, etapa, x, 3.5, w, 0.5, size=14, bold=True, color=W, align=PP_ALIGN.CENTER)
        add_text(s, desc, x + 0.05, 4.2, w - 0.1, 1.2, size=10, color=W, align=PP_ALIGN.CENTER, italic=True)
        x += w + 0.05
    add_text(s, "Etapas 1 a 3 são com o PROVEDOR    ·    Etapas 4, 5, 6 são o SEU TRABALHO", 0.7, 6.3, 11.9, 0.4, size=15, bold=True, color=P, align=PP_ALIGN.CENTER)
    add_text(s, "Saber de que lado está o problema é metade do atendimento.", 0.7, 6.7, 11.9, 0.4, size=13, italic=True, color=G, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 5: A divisão que importa
    s = add_slide(prs)
    add_title_block(s, "A DIVISÃO QUE IMPORTA", "Onde o seu trabalho começa e termina")
    add_text(s, "É por isso que o vizinho pode estar sem internet e você não:", 0.7, 2.8, 11.9, 0.5, size=20, italic=True, color=K, align=PP_ALIGN.CENTER)
    add_text(s, "cada casa tem o seu drop, saindo da mesma caixa no poste.", 0.7, 3.3, 11.9, 0.5, size=20, italic=True, color=K, align=PP_ALIGN.CENTER)
    add_text(s, [
        "•  Etapas 1, 2, 3 (servidor, backbone, provedor)",
        "     Quando o problema é aqui, não adianta reiniciar roteador nenhum",
        "",
        "•  Etapas 4, 5, 6 (rua, casa, aparelho)",
        "     É onde o técnico de campo atua",
    ], 0.7, 4.5, 11.9, 2.5, size=17, color=K)
    add_footer(s, FOOTER)

    # 6: Os 5 equipamentos
    s = add_slide(prs)
    add_title_block(s, "OS BICHOS", "Quem é quem numa rede")
    equip = [
        ("MODEM", "Converte o sinal do provedor no sinal que a rede entende"),
        ("ROTEADOR", "Cria a rede da casa e decide para onde cada informação vai"),
        ("SWITCH", "Aumenta o número de portas de cabo. Não cria rede"),
        ("PONTO DE ACESSO", "Leva o sinal sem fio onde o roteador não alcança"),
        ("PLACA DE REDE", "Onde o cabo entra no computador"),
    ]
    y = 2.7
    for tit, desc in equip:
        add_text(s, tit, 0.7, y, 3.5, 0.5, size=17, bold=True, color=P)
        add_text(s, desc, 4.3, y, 8.3, 0.5, size=15, color=K)
        y += 0.75
    add_text(s, "Cinco nomes. Nas próximas telas, um de cada vez.", 0.7, 6.6, 11.9, 0.4, size=14, italic=True, color=G, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 7: MODEM
    s = add_slide(prs)
    add_title_block(s, "UM DE CADA VEZ · 1", "MODEM", "O tradutor. Converte o sinal que vem da rua.")
    add_text(s, [
        "O que faz",
        "  Pega a luz da fibra (ou o sinal do cabo) e entrega algo que o roteador entende",
        "",
        "Quantas portas",
        "  Uma só, de saída. Ele não distribui nada",
        "",
        "Cria rede?",
        "  Não. Sozinho, ele atende um aparelho e só",
    ], 0.7, 2.8, 11.9, 3.8, size=17, color=K)
    add_text(s, "Modem traduz. Quem distribui é o roteador.", 0.7, 6.6, 11.9, 0.4, size=15, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 8: ROTEADOR
    s = add_slide(prs)
    add_title_block(s, "UM DE CADA VEZ · 2", "ROTEADOR", "O distribuidor. É ele que cria a rede da casa.")
    add_text(s, [
        "O que faz",
        "  Recebe uma conexão e divide entre todos os aparelhos, por cabo e sem fio",
        "",
        "Como decide",
        "  Guarda quem é quem na rede e manda cada resposta para o aparelho certo",
        "",
        "Tem senha",
        "  Duas: a da rede sem fio e a de administração do aparelho. SÃO DIFERENTES.",
    ], 0.7, 2.8, 11.9, 3.8, size=17, color=K)
    add_text(s, "Sem roteador, um aparelho só. Com roteador, a casa inteira.", 0.7, 6.6, 11.9, 0.4, size=15, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 9: SWITCH
    s = add_slide(prs)
    add_title_block(s, "UM DE CADA VEZ · 3", "SWITCH", "Um multiplicador de portas. E só isso.")
    add_text(s, [
        "O que faz",
        "  Transforma uma porta de cabo em quatro, oito, dezesseis",
        "",
        "O que NÃO faz",
        "  Não cria rede, não distribui internet sozinho, não tem Wi-Fi",
        "",
        "Como diferenciar do roteador",
        "  Olhem a traseira: roteador tem uma porta diferente (WAN). Switch tem todas iguais.",
    ], 0.7, 2.8, 11.9, 3.8, size=16, color=K)
    add_text(s, "Switch sem roteador não dá internet a ninguém. É extensão de porta.", 0.7, 6.6, 11.9, 0.4, size=14, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 10: PONTO DE ACESSO
    s = add_slide(prs)
    add_title_block(s, "UM DE CADA VEZ · 4", "PONTO DE ACESSO", "Leva o sinal sem fio onde o roteador não alcança.")
    add_text(s, [
        "O que faz",
        "  Cria mais um ponto de Wi-Fi, ligado na mesma rede da casa",
        "",
        "Como se liga",
        "  O ideal é por cabo até o roteador. Assim o sinal não perde força.",
        "",
        "O repetidor",
        "  Faz parecido, mas sem cabo — repete o que recebe pelo ar, e perde velocidade",
    ], 0.7, 2.8, 11.9, 3.8, size=16, color=K)
    add_text(s, "Levar cabo até um ponto de acesso resolve melhor que empilhar repetidor.", 0.7, 6.6, 11.9, 0.4, size=14, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 11: PLACA DE REDE
    s = add_slide(prs)
    add_title_block(s, "UM DE CADA VEZ · 5", "PLACA DE REDE", "A porta de entrada da rede dentro do computador.")
    add_text(s, [
        "Onde fica",
        "  Integrada na placa-mãe, na traseira do gabinete",
        "",
        "As luzinhas",
        "  Uma acesa fixa = cabo com sinal. A outra pisca = tráfego passando.",
        "",
        "Se queimar",
        "  Resolve com uma placa de rede USB, que custa pouco.",
    ], 0.7, 2.8, 11.9, 3.8, size=17, color=K)
    add_text(s, "A luzinha piscando na traseira é a primeira coisa a olhar num chamado.", 0.7, 6.6, 11.9, 0.4, size=14, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 12: O quadro dos 5
    s = add_slide(prs)
    add_title_block(s, "DE UMA OLHADA", "O quadro dos cinco")
    rows = [
        ["Equipamento", "Cria rede?", "Tem Wi-Fi?", "Para que serve"],
        ["Modem", "Não", "Não", "Traduzir o sinal do provedor"],
        ["Roteador", "Sim", "Quase sempre", "Distribuir para a casa toda"],
        ["Switch", "Não", "Não", "Aumentar o número de portas"],
        ["Ponto de acesso", "Não", "Sim", "Ampliar a cobertura sem fio"],
        ["Placa de rede", "Não", "Depende", "Conectar a máquina à rede"],
    ]
    add_table_slide(s, rows, top=2.8, height=3.8, col_widths=[3.2, 2.2, 2.5, 4.0], size=14)
    add_text(s, "Só um deles cria rede. Guardem isso.", 0.7, 6.7, 11.9, 0.4, size=18, bold=True, color=D, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 13: O erro número um
    s = add_slide(prs)
    add_title_block(s, "O ERRO NÚMERO UM", "A porta que é diferente das outras")
    add_text(s, [
        "WAN         LAN     LAN     LAN     LAN",
        "entrada     saídas para os aparelhos da casa",
        "provedor",
    ], 0.7, 2.8, 11.9, 1.5, size=22, bold=True, color=P, align=PP_ALIGN.CENTER)
    add_text(s, [
        "Como reconhecer",
        "  Fica separada das outras, quase sempre de outra cor, e escrito WAN ou Internet",
        "",
        "O que acontece se errar",
        "  Tudo conecta na rede normalmente, mas nada abre. Parece defeito e não é.",
    ], 0.7, 4.5, 11.9, 2.0, size=15, color=K)
    add_text(s, "Confundir essas duas é o erro mais comum de quem instala roteador.", 0.7, 6.6, 11.9, 0.4, size=14, bold=True, color=R, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 14: Onde está escrito
    s = add_slide(prs)
    add_title_block(s, "ONDE ESTÁ ESCRITO", "A etiqueta embaixo do aparelho")
    rows = [
        ["O que tem", "Para que serve"],
        ["Nome de rede (SSID) de fábrica", "Nome que aparece antes de alguém trocar"],
        ["Senha do Wi-Fi de fábrica", "Conectar na primeira vez"],
        ["Usuário e senha de administração", "Entram na configuração — OUTRA senha"],
        ["Endereço de configuração", "Número que se digita no navegador"],
        ["Modelo e número de série", "O que o suporte do provedor vai pedir"],
    ]
    add_table_slide(s, rows, top=2.8, height=3.5, col_widths=[5, 6.9], size=14)
    add_text(s, "Tire uma foto da etiqueta antes de pendurar o aparelho na parede.", 0.7, 6.6, 11.9, 0.4, size=15, bold=True, color=D, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 15: As luzes
    s = add_slide(prs)
    add_title_block(s, "SEM ABRIR NADA", "As luzes do painel já contam a história")
    rows = [
        ["A luz", "Acesa fixa", "Piscando", "Apagada"],
        ["Power", "Ligado", "—", "Sem energia"],
        ["Fibra / Internet", "Sinal chegando", "Tentando conectar", "Sem sinal do provedor"],
        ["WLAN / Wi-Fi", "Wi-Fi ligado", "Tráfego passando", "Wi-Fi desligado"],
        ["LAN 1 a 4", "Cabo conectado", "Dados passando", "Nada ligado na porta"],
    ]
    add_table_slide(s, rows, top=2.8, height=3.5, col_widths=[3, 3, 3, 2.9], size=13)
    add_text(s, "Luz de fibra apagada: nem adianta mexer no computador. O problema está antes.", 0.7, 6.6, 11.9, 0.4, size=14, bold=True, color=R, italic=True, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 16: PRÁTICA 1 - Identificação
    s = add_slide(prs)
    add_title_block(s, "PRÁTICA · 45 MIN", "Identificação e ficha")
    add_text(s, [
        "1.  Identificar cada equipamento e dizer o que faz",
        "2.  Achar a porta WAN de cada roteador (e explicar como reconheceu)",
        "3.  Contar as portas LAN e anotar",
        "4.  Ler a etiqueta e preencher a ficha (página 3 da folha)",
        "5.  Ligar o roteador e observar as luzes acendendo",
        "6.  Trocar de bancada e conferir o que a outra equipe anotou",
        "7.  Apresentar 1 equipamento em 30 segundos pra turma",
    ], 0.7, 2.8, 11.9, 4.0, size=16, color=K)
    add_text(s, "Hoje ninguém configura nada. Hoje é reconhecer, nomear e anotar.", 0.7, 6.7, 11.9, 0.4, size=15, italic=True, color=D, bold=True, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 17: PRÁTICA 2 - Role-play
    s = add_slide(prs)
    add_title_block(s, "PRÁTICA · 30 MIN", "Role-play: Cliente × Técnico")
    add_text(s, [
        "Duplas. Um é o cliente, outro é o técnico.",
        "",
        "•  Dupla recebe um cartão de chamado (o técnico NÃO vê)",
        "•  Cliente lê e representa o cliente de verdade — sem vocabulário técnico",
        "•  Técnico faz 3-4 perguntas, aponta o equipamento com problema",
        "     e diz o que investigaria primeiro",
        "",
        "•  Depois: troca de papéis, novo cartão",
        "",
        "•  Discussão coletiva: 3 duplas contam qual foi o chamado",
    ], 0.7, 2.8, 11.9, 4.2, size=15, color=K)
    add_footer(s, FOOTER)

    # 18: Método do técnico
    s = add_slide(prs)
    add_title_block(s, "O MÉTODO", "O que o técnico verifica ao chegar")
    add_text(s, [
        "1.  A luz de fibra ou internet está acesa? Se não, o problema é antes do roteador",
        "2.  Os cabos estão firmes, com clique? Cabo meio solto é campeão de chamado",
        "3.  O cabo do provedor está na porta WAN, e não numa LAN?",
        "4.  Outro aparelho conecta? Se só um falha, o problema é no aparelho",
        "5.  Conecta na rede mas não abre página? Aí é sinal do provedor",
        "6.  Desligar da tomada por 10 segundos e ligar resolveu? Registre que resolveu.",
    ], 0.7, 2.8, 11.9, 4.0, size=14, color=K)
    add_text(s, "Sempre na mesma ordem. Método evita retrabalho e troca desnecessária.", 0.7, 6.7, 11.9, 0.4, size=14, italic=True, color=G, align=PP_ALIGN.CENTER)
    add_footer(s, FOOTER)

    # 19: Encene um chamado
    s = add_slide(prs)
    add_title_block(s, "ENCENAR UM CHAMADO", "Cliente × Técnico ao vivo, turma inteira assiste")
    add_text(s, [
        "•  Professor escolhe 2 alunos: 1 cliente, 1 técnico",
        "•  Professor sussurra o defeito só pro cliente",
        "•  Técnico faz perguntas na ordem do método",
        "•  Turma assiste e ajuda quando o técnico trava",
        "",
        "Defeitos que rolam bem:",
        "     •  Cabo do provedor na porta LAN 3",
        "     •  Cabo do computador solto no roteador",
        "     •  Roteador foi resetado e perdeu configuração",
    ], 0.7, 2.8, 11.9, 4.2, size=15, color=K)
    add_footer(s, FOOTER)

    # 20: Duelo
    s = add_slide(prs)
    add_title_block(s, "DUELO RELÂMPAGO · 15 MIN", "Time A × Time B")
    add_text(s, [
        "As perguntas de hoje giram em torno de:",
        "",
        "     •  Diferença entre os 5 equipamentos",
        "     •  Reconhecer a porta WAN",
        "     •  Ler as luzes do painel",
        "     •  Etapas 1-3 × etapas 4-6",
        "     •  Cenários de chamado",
        "",
        "Resposta em 5 segundos. Alterna entre times.",
    ], 0.7, 2.8, 11.9, 4.2, size=17, color=K)
    add_footer(s, FOOTER)

    # 21: Resumo
    s = add_slide(prs)
    add_title_block(s, "O RESUMO DA NOITE", "Não tem internet dentro da caixinha")
    add_text(s, [
        "•  Ela é só a ponta de um caminho que começa do outro lado do mundo.",
        "",
        "•  Seu trabalho começa no poste da rua e termina na tela do cliente.",
        "",
        "•  Tudo que vem antes disso é com o provedor.",
    ], 0.7, 3.0, 11.9, 3.5, size=20, color=K)
    add_footer(s, FOOTER)

    # 22: Fechando
    s = add_slide(prs)
    add_title_block(s, "FECHANDO", "Seis perguntas para conferir")
    add_text(s, [
        "1.  Qual é a diferença entre modem e roteador?",
        "2.  O switch cria rede? O que ele faz então?",
        "3.  Como você reconhece a porta WAN num roteador que nunca viu?",
        "4.  O que acontece se o cabo do provedor for ligado numa porta LAN?",
        "5.  Por que o Wi-Fi do quarto costuma ser pior que o da sala?",
        "6.  A luz de fibra está apagada. Onde está o problema?",
    ], 0.7, 2.8, 11.9, 4.0, size=16, color=K)
    add_footer(s, FOOTER)

    # 23: O que fica
    s = add_slide(prs)
    add_title_block(s, "O QUE FICA DESTA NOITE", "Cinco coisas para levar")
    add_text(s, [
        "•  A internet vem de longe e passa por seis etapas até você",
        "•  Só o roteador cria rede — os outros ajudam",
        "•  A porta WAN é diferente, e trocar ela é o erro mais comum",
        "•  A etiqueta embaixo tem tudo que o suporte vai pedir",
        "•  As luzes do painel dizem onde está o problema antes de você abrir nada",
    ], 0.7, 3.0, 11.9, 3.5, size=17, color=K)
    add_footer(s, FOOTER)

    # 24: Para a próxima
    s = add_slide(prs)
    add_title_block(s, "PARA A PRÓXIMA NOITE", "Por onde o sinal viaja")
    add_text(s, [
        "Duas tarefas simples:",
        "",
        "•  Procurar um cabo de rede em casa e ler o que está escrito na capa",
        "•  Anotar a categoria, o fabricante e o número da metragem",
        "•  Se não achar cabo de rede, sirva qualquer cabo (TV, energia)",
        "•  Trazer anotado — vamos comparar as capas na aula",
        "",
        "Quem não achar nenhum, sem problema — vai ter cabo de sobra na mesa.",
    ], 0.7, 2.8, 11.9, 4.2, size=15, color=K)
    add_footer(s, FOOTER)

    path = os.path.join(OUTDIR, "Aula - Redes 2 - Como a internet chega.pptx")
    prs.save(path)
    return path


# ============================================================
# FOLHA DO ALUNO (3 páginas)
# ============================================================
def build_folha():
    doc = Document()
    setup_page(doc, margem=Cm(1.8))

    # PÁGINA 1
    cabecalho_documento(
        doc,
        "COMO A INTERNET CHEGA + 5 EQUIPAMENTOS",
        f"Redes de Computadores · {NOITE} · Curso FIC · CEJA Itapiranga · 1 de 3",
    )

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("Nome: ")
    r.font.size = Pt(11)
    r.bold = True
    r = p.add_run("_________________________________________________     Data: ____ / ____ / ______")
    r.font.size = Pt(11)

    add_heading(doc, "1  O caminho da internet até você", level=2)
    add_para(doc, "Numere de 1 a 6, na ordem em que a informação passa, do servidor até o seu celular.", italic=True, size=10, color=COR_CINZA)
    itens_caminho = [
        "O roteador da sua casa",
        "A fibra no poste da sua rua",
        "O servidor que guarda o site",
        "O seu celular ou computador",
        "O provedor da sua cidade",
        "As fibras de longa distância",
    ]
    for item in itens_caminho:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("(       )   ")
        r.font.size = Pt(11)
        r.bold = True
        r = p.add_run(item)
        r.font.size = Pt(11)

    add_heading(doc, "2  Os cinco equipamentos", level=2)
    add_para(doc, "Preencha durante a aula. Só um deles cria rede.", italic=True, size=10, color=COR_CINZA)

    tbl = add_table(doc, [
        ["Equipamento", "Cria rede?", "Tem Wi-Fi?", "Para que serve"],
        ["Modem", "", "", ""],
        ["Roteador", "", "", ""],
        ["Switch", "", "", ""],
        ["Ponto de acesso", "", "", ""],
        ["Placa de rede", "", "", ""],
    ], col_widths=[3.5, 2.5, 2.5, 8])
    for row in tbl.rows[1:]:
        row.height = Cm(1.1)

    add_heading(doc, "3  A porta que é diferente das outras", level=2)
    add_para(doc, "Escreva embaixo de cada porta o que entra ou sai nela.", italic=True, size=10, color=COR_CINZA)
    p = doc.add_paragraph()
    r = p.add_run("[  WAN  ]   [  LAN  ]   [  LAN  ]   [  LAN  ]   [  LAN  ]")
    r.font.size = Pt(14)
    r.bold = True
    r.font.color.rgb = COR_PRIMARIA
    for _ in range(3):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("_" * 100)
        r.font.size = Pt(11)
        r.font.color.rgb = COR_CINZA
    add_para(doc, "Ligar o cabo do provedor na porta errada: tudo conecta, e nada abre.", bold=True, italic=True, color=COR_ALERTA, size=11)

    add_heading(doc, "4  Escreva com suas palavras", level=2)
    p = doc.add_paragraph()
    r = p.add_run("Qual é a diferença entre o modem e o roteador?")
    r.font.size = Pt(11)
    r.bold = True
    for _ in range(5):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("_" * 100)
        r.font.size = Pt(11)
        r.font.color.rgb = COR_CINZA

    # PÁGINA 2 — Role-play
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "CHAMADOS PARA VOCÊ ATENDER",
        f"Role-play cliente/técnico · {NOITE} · 2 de 3",
    )

    add_para(doc, "Nesta atividade você faz o papel de técnico. Sua dupla recebe um cartão com o problema do cliente e vai te contar como se fosse um cliente de verdade. Sua missão: descobrir qual equipamento provavelmente está com problema, fazendo 3-4 perguntas.", italic=True, size=10, color=COR_CINZA)

    cenarios_folha = [
        "\"Instalei um switch em casa mas nada conecta na internet.\"",
        "\"O Wi-Fi tá aceso no meu celular mas nada abre.\"",
        "\"Comprei um roteador novo. Não sei em que porta ligar o cabo que o técnico deixou.\"",
        "\"A impressora do meu escritório não imprime, mas a internet funciona.\"",
        "\"A luz de fibra do meu roteador tá apagada. Reiniciei três vezes.\"",
    ]
    for i, cen in enumerate(cenarios_folha, 1):
        add_heading(doc, f"Chamado {i}", level=2)
        add_para(doc, cen, italic=True, color=COR_PRIMARIA, size=11)
        p = doc.add_paragraph()
        r = p.add_run("Que perguntas você faria pro cliente?")
        r.font.size = Pt(10)
        r.bold = True
        for _ in range(2):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run("_" * 100)
            r.font.size = Pt(11)
            r.font.color.rgb = COR_CINZA
        p = doc.add_paragraph()
        r = p.add_run("Qual equipamento suspeito? Por quê?")
        r.font.size = Pt(10)
        r.bold = True
        for _ in range(2):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run("_" * 100)
            r.font.size = Pt(11)
            r.font.color.rgb = COR_CINZA

    # PÁGINA 3 — Ficha de identificação + tarefa
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "FICHA DE IDENTIFICAÇÃO",
        f"Equipamento de rede · {NOITE} · 3 de 3",
    )

    add_para(doc, "Como preencher: quase tudo está na etiqueta embaixo do aparelho. Não mexa em configuração nenhuma — hoje é só olhar, contar e anotar. NÃO anote senha de Wi-Fi nesta folha.", italic=True, size=10, color=COR_ALERTA, bold=True)

    add_heading(doc, "1  O aparelho da bancada", level=2)
    tbl = add_table(doc, [
        ["Item", "Sua resposta"],
        ["Marca", ""],
        ["Modelo", ""],
        ["Nome de rede (SSID) de fábrica", ""],
        ["Endereço de configuração", ""],
        ["Quantas portas LAN?", ""],
        ["A porta WAN tem outra cor? Qual?", ""],
        ["Tem antena? Quantas?", ""],
    ], col_widths=[7, 10.5])
    for row in tbl.rows[1:]:
        row.height = Cm(0.8)

    add_heading(doc, "2  As luzes do painel", level=2)
    add_para(doc, "Ligue o aparelho e marque o que cada luz fez.", italic=True, size=10, color=COR_CINZA)
    tbl = add_table(doc, [
        ["A luz", "Acesa fixa", "Piscando", "Apagada"],
        ["Power", "☐", "☐", "☐"],
        ["Fibra / Internet", "☐", "☐", "☐"],
        ["WLAN / Wi-Fi", "☐", "☐", "☐"],
        ["LAN 1 a 4", "☐", "☐", "☐"],
    ], col_widths=[4, 4.5, 4.5, 4.5])
    for row in tbl.rows[1:]:
        row.height = Cm(0.8)

    add_heading(doc, "3  Duas perguntas", level=2)
    p = doc.add_paragraph()
    r = p.add_run("1. Como você reconheceu a porta WAN sem ninguém falar?")
    r.font.size = Pt(11)
    r.bold = True
    for _ in range(2):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("_" * 100)
        r.font.size = Pt(11)
        r.font.color.rgb = COR_CINZA
    p = doc.add_paragraph()
    r = p.add_run("2. A luz de fibra está apagada. Onde está o problema, e por quê?")
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
    add_bullet(doc, "Procure um cabo de rede em casa (ou qualquer cabo, se não achar) e anote o que está escrito na capa")
    add_bullet(doc, "Categoria, fabricante e a metragem (o número impresso a cada metro)")
    add_bullet(doc, "Traga anotado — vamos comparar as capas na Noite 3")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Traga esta folha preenchida · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Folha do Aluno - Redes 2 - Como a internet chega.docx")
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
        "COMO A INTERNET CHEGA + OS 5 EQUIPAMENTOS",
        f"Resumo do aluno · {NOITE} · Redes de Computadores · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Não tem internet dentro da caixinha. Ela é só a ponta de um caminho que começa do outro lado do mundo. Seu trabalho começa no poste da rua e termina na tela do cliente. Tudo que vem antes disso é com o provedor.")

    add_heading(doc, "O caminho da internet — 6 etapas", level=2)
    add_table(doc, [
        ["Etapa", "O que é", "De quem é o problema"],
        ["1. Servidor", "O computador que guarda o site", "Provedor"],
        ["2. Backbone", "Fibras de longa distância", "Provedor"],
        ["3. Provedor", "A empresa da sua cidade", "Provedor"],
        ["4. A rua", "Fibra no poste + drop até a casa", "Você"],
        ["5. Sua casa", "Caixinha e roteador", "Você"],
        ["6. Aparelho", "Cabo ou Wi-Fi até você", "Você"],
    ], col_widths=[3, 8.5, 4])
    add_para(doc, "É por isso que o vizinho pode estar sem internet e você não: cada casa tem o seu drop.", italic=True, color=COR_DESTAQUE, bold=True)

    add_heading(doc, "Os 5 equipamentos — só um cria rede", level=2)
    add_table(doc, [
        ["Equipamento", "Cria rede?", "Tem Wi-Fi?", "Para que serve"],
        ["Modem", "Não", "Não", "Traduz o sinal do provedor"],
        ["Roteador", "Sim", "Quase sempre", "Distribui pra casa toda"],
        ["Switch", "Não", "Não", "Aumenta o número de portas"],
        ["Ponto de acesso", "Não", "Sim", "Amplia a cobertura sem fio"],
        ["Placa de rede", "Não", "Depende", "Conecta a máquina à rede"],
    ], col_widths=[3.5, 2.5, 2.5, 7])

    add_heading(doc, "A porta WAN — o erro nº 1", level=2)
    add_para(doc, "No roteador, uma porta é diferente das outras: quase sempre de outra cor, escrita WAN ou Internet. É por ela que entra o cabo do provedor. As outras (todas do mesmo tamanho) são LAN — para os aparelhos da casa.")
    add_para(doc, "Se você ligar o cabo do provedor numa porta LAN, tudo conecta na rede mas nada abre. Parece defeito, e não é. É o erro mais comum.", bold=True, color=COR_ALERTA)

    add_heading(doc, "A etiqueta embaixo do roteador", level=2)
    add_bullet(doc, "Nome de rede (SSID) e senha do Wi-Fi de fábrica")
    add_bullet(doc, "Usuário e senha de administração — SÃO OUTRAS senhas")
    add_bullet(doc, "Endereço de configuração (número pra digitar no navegador)")
    add_bullet(doc, "Modelo e número de série (o suporte vai pedir)")
    add_para(doc, "Tire uma foto da etiqueta antes de pendurar o aparelho na parede.", italic=True, color=COR_PRIMARIA, bold=True)

    add_heading(doc, "As luzes do painel", level=2)
    add_table(doc, [
        ["A luz", "Acesa fixa", "Apagada"],
        ["Power", "Ligado", "Sem energia"],
        ["Fibra / Internet", "Sinal chegando", "Sem sinal do provedor"],
        ["Wi-Fi", "Wi-Fi ligado", "Wi-Fi desligado"],
        ["LAN 1-4", "Cabo conectado", "Nada ligado"],
    ], col_widths=[4, 5.5, 6])
    add_para(doc, "Luz de fibra apagada = problema antes do roteador. Nem adianta mexer no computador.", italic=True, color=COR_ALERTA, bold=True)

    add_heading(doc, "O método do técnico", level=2)
    add_bullet(doc, "1. A luz de fibra está acesa? Se não, o problema é antes do roteador.")
    add_bullet(doc, "2. Os cabos estão firmes, com clique?")
    add_bullet(doc, "3. O cabo do provedor está na porta WAN, e não numa LAN?")
    add_bullet(doc, "4. Outro aparelho conecta? Se só um falha, o problema é no aparelho.")
    add_bullet(doc, "5. Conecta na rede mas não abre página? Aí é sinal do provedor.")
    add_bullet(doc, "6. Desligar da tomada por 10 segundos resolveu? Registre.")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Não tem internet dentro da caixinha — ela é só a ponta de um caminho.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Cinco equipamentos, e só um deles cria rede: o roteador.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"A porta WAN é o erro número um de quem instala roteador.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Qual é a diferença entre modem e roteador, em uma frase?")
    add_bullet(doc, "Você vê um roteador que nunca viu. Como reconhece qual é a porta WAN?")
    add_bullet(doc, "Cliente diz 'a luz vermelha da fibra tá apagada'. Onde tá o problema?")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Este material é seu. Guarde e volte quando precisar. · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Resumo do Aluno - Redes 2 - Como a internet chega.docx")
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
