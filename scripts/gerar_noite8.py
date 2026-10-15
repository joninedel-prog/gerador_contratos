"""Gera os 4 arquivos da Noite 8 do curso de Redes.

Noite 8 — Configurando o roteador: da caixa à entrega
Padrão prática-pesada, sem role-play e sem duelo. Parte da tarefa da Noite 7
(etiqueta do roteador de casa). Na Noite 6 a turma só entrou no painel, criou a
senha de administrador e olhou a lista de aparelhos. Hoje cada equipe configura
de verdade um roteador de fábrica: nome e senha do Wi-Fi, endereço do painel,
faixa de endereços e reserva de IP, a agenda (DNS) entregue pelo roteador,
rede de convidados, cópia de segurança com restauração e ficha de entrega.
No fim as equipes trocam de roteador e testam se a ficha da outra funciona.
Não precisa de internet.
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
    add_slide_cover, add_footer, add_table_slide,
)
from pptx.enum.text import PP_ALIGN

OUTDIR = "/home/nedel/curso"
NOITE = "Noite 8"
FOOTER = f"Redes de Computadores · {NOITE}"


# O que se configura e por quê (a entrega)
ENTREGA = [
    ("Senha de administrador", "Quem tem ela manda no roteador. De fábrica, qualquer um entra"),
    ("Nome do Wi-Fi", "O cliente acha a rede dele. \"TP-Link_3F2A\" não diz de quem é"),
    ("Senha do Wi-Fi + WPA2", "A senha da etiqueta qualquer visita lê. WPA2 (ou WPA3) é a tranca"),
    ("Endereço do painel", "Não pode bater com o endereço da ONU do provedor"),
    ("Faixa de endereços e reserva", "A impressora e a câmera ficam sempre no mesmo número"),
    ("A agenda (DNS)", "O roteador diz pros aparelhos qual agenda usar"),
    ("Rede de convidados", "A visita navega sem enxergar os aparelhos da casa"),
    ("Cópia de segurança", "Se o roteador resetar, tudo volta em 1 minuto"),
]

# Onde fica cada coisa nos TP-Link da bancada (o nome muda com o firmware)
MENUS = [
    ("Nome e senha do Wi-Fi", "Sem fio (Wireless) → Configurações · ou Básico → Sem fio"),
    ("Tipo de segurança", "Sem fio → Segurança sem fio · WPA2-PSK, criptografia AES"),
    ("Endereço do painel", "Rede → LAN · ou Avançado → Rede → LAN"),
    ("Faixa de endereços (DHCP)", "DHCP → Configurações DHCP · ou Avançado → Rede → Servidor DHCP"),
    ("Reserva de endereço", "DHCP → Reserva de endereço · ou na mesma tela, \"Reserva\""),
    ("A agenda (DNS)", "Na tela do DHCP: \"DNS primário\" e \"DNS secundário\""),
    ("Rede de convidados", "Rede de convidados (Guest Network) · nem todo modelo tem"),
    ("Cópia de segurança", "Ferramentas do sistema → Backup e restauração"),
]


def plano_equipe(n):
    """Plano de endereços de cada equipe (o painel muda pra 192.168.1X.1)."""
    rede = f"192.168.{10 + n}"
    return {
        "painel": f"{rede}.1",
        "faixa": f"{rede}.100 a {rede}.150",
        "reserva": f"{rede}.50",
    }


# Ficha de entrega: os campos
FICHA = [
    "Cliente / local",
    "Marca e modelo do roteador",
    "Endereço do painel",
    "Senha de administrador",
    "Nome do Wi-Fi",
    "Senha do Wi-Fi e segurança (WPA2/WPA3)",
    "Rede de convidados: nome e senha",
    "Faixa de endereços (DHCP)",
    "Reservas (aparelho → endereço)",
    "Agenda (DNS) entregue pelo roteador",
    "Cópia de segurança salva em",
    "Configurado por / data",
]


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "CONFIGURANDO O ROTEADOR: DA CAIXA À ENTREGA",
        f"Roteiro do professor · Redes de Computadores · {NOITE} · Curso FIC · CEJA Itapiranga – SC",
    )

    add_heading(doc, "A IDEIA DESTA NOITE")
    add_para(doc, "Na Noite 6 a turma entrou no painel, criou a senha de administrador e só olhou. Hoje mexe de verdade. Cada equipe pega um roteador de fábrica e deixa ele pronto pra entregar ao cliente: nome, senha, endereços, a agenda (DNS), rede de convidados e cópia de segurança. No fim, outra equipe recebe o roteador com a ficha de entrega e testa se dá pra usar só com o papel.")
    add_para(doc, "É o serviço que mais vai aparecer pra eles: roteador novo na caixa, roteador que resetou, roteador que \"ninguém sabe a senha\". Quem sabe configurar e documentar cobra pelo serviço.")
    add_para(doc, "A frase da noite: \"Roteador de fábrica funciona. Roteador configurado é serviço entregue.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "O que se resolve nesta noite", level=2)
    add_para(doc, "Cada aluno sai sabendo trocar nome e senha do Wi-Fi e escolher a segurança, mudar o endereço do painel e reencontrar o roteador depois, ajustar a faixa de endereços e reservar um endereço fixo, conferir qual agenda (DNS) o roteador entrega, criar rede de convidados, salvar e restaurar a configuração e preencher a ficha de entrega.")

    add_heading(doc, "Esta é a oitava de quinze noites", level=2)
    add_para(doc, "Fecha o que a Noite 6 abriu (o painel) e o que a Noite 7 prometeu (onde se troca a agenda). A Noite 9 usa estes mesmos roteadores, já configurados, pra medir cobertura de Wi-Fi. A Noite 10 volta a eles pra auditar a segurança.")
    add_para(doc, "Não entre em máscara de sub-rede, NAT, redirecionamento de portas nem controle dos pais. Senha forte e WPS ficam pra Noite 10: hoje a senha só precisa ser longa e não ser a da etiqueta.", italic=True, color=COR_ALERTA)

    add_hr(doc)

    add_heading(doc, "ANTES DA AULA")
    add_heading(doc, "Material da mesa", level=2)
    add_bullet(doc, "Um roteador por equipe de 3 ou 4 (os da bancada), cada um na tomada, numerado com fita crepe: EQUIPE 1, EQUIPE 2…")
    add_bullet(doc, "Régua de tomadas. NÃO precisa de internet")
    add_bullet(doc, "1 ou 2 computadores com porta de rede + cabos testados da Noite 5 (pra cópia de segurança e pro ipconfig)")
    add_bullet(doc, "Fita crepe e canetão (pra etiqueta de entrega que vai no roteador)")
    add_bullet(doc, "Celulares dos alunos: são eles que configuram, pelo navegador")

    add_heading(doc, "Prepare os roteadores", level=2)
    add_bullet(doc, "Dê Reset em todos (botão Reset, ~10 segundos com o aparelho ligado). A turma começa de fábrica, como cliente que acabou de comprar")
    add_bullet(doc, "Em casa, configure UM de cada modelo seguindo este roteiro inteiro, pra conhecer as telas, e dê Reset de novo. Anote no verso da fita de cada modelo o caminho dos menus (a tabela abaixo ajuda, mas o nome muda com o firmware)")
    add_bullet(doc, "Confira se o seu modelo tem rede de convidados. O Archer C5 tem. Os TL-WR840N mais antigos podem não ter: nessas equipes o Bloco 6 vira observação no roteador de outra equipe")
    add_table(doc, [["O quê", "Onde costuma ficar"]] + [list(r) for r in MENUS], col_widths=[4.5, 12])

    add_heading(doc, "O plano de endereços das equipes", level=2)
    add_para(doc, "Cada equipe muda o endereço do painel pra um número próprio. Assim nenhum roteador fica igual ao outro e ninguém entra no da vizinha por engano:")
    add_table(doc, [["Equipe", "Novo endereço do painel", "Faixa de endereços", "Reserva"]] +
              [[str(n), plano_equipe(n)["painel"], plano_equipe(n)["faixa"], plano_equipe(n)["reserva"]] for n in range(1, 7)],
              col_widths=[2, 4.8, 6, 3.7])

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "Folha do aluno (3 páginas), uma por pessoa")
    add_bullet(doc, "Resumo do aluno, uma por pessoa (pra levar pra casa)")
    add_bullet(doc, "Uma ficha de entrega extra por equipe (é a página 3 da folha): ela vai junto com o roteador no Bloco 9")

    add_hr(doc)

    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Retomada + censo das etiquetas", "20 min", "Quantas casas ainda estão com o Wi-Fi de fábrica?"],
        ["2. O que se configura e por quê", "15 min", "Os oito itens da entrega. Onde fica cada um no painel"],
        ["3. PRÁTICA — Nome, senha e tranca", "35 min", "Primeiro acesso, senha de administrador, nome e senha do Wi-Fi, WPA2. Todo mundo reconecta"],
        ["4. PRÁTICA — Mudar a portaria de lugar", "15 min", "O painel muda de endereço. A equipe perde o roteador e reencontra pelo gateway"],
        ["5. PRÁTICA — Endereços e a agenda", "25 min", "Faixa do DHCP, reserva de endereço e o DNS que o roteador entrega"],
        ["INTERVALO", "15 min", ""],
        ["6. PRÁTICA — Rede de convidados", "15 min", "A visita navega mas não enxerga a casa. Teste: o painel abre pela rede de convidados?"],
        ["7. PRÁTICA — Cópia de segurança", "20 min", "Salvar a configuração, resetar o roteador e restaurar"],
        ["8. PRÁTICA — Ficha e etiqueta de entrega", "15 min", "A ficha pro cliente e a etiqueta no roteador"],
        ["9. PRÁTICA — O roteador trocado", "25 min", "Cada equipe recebe o roteador e a ficha de outra. A ficha funciona?"],
        ["10. Fechamento + tarefa", "10 min", "Quiz de saída e anúncio da Noite 9"],
    ], col_widths=[6.5, 2, 8])
    add_para(doc, "Blocos 3 a 9: 150 min de atividade do aluno (71% do total), mais o censo do Bloco 1. A exposição é só o Bloco 2 (15 min) e falas curtas dentro das práticas.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — RETOMADA + CENSO DAS ETIQUETAS · 20 MIN")
    add_heading(doc, "Recap relâmpago — 5 min", level=2)
    add_bullet(doc, "Qual o endereço da portaria (o painel) dos roteadores da bancada? (192.168.0.1)")
    add_bullet(doc, "Quem entrega o endereço pro celular? (o roteador, o DHCP, o porteiro)")
    add_bullet(doc, "O DNS faz o quê? (a agenda: nome vira número)")
    add_bullet(doc, "Como o intruso entrou no Wi-Fi da equipe na Noite 6? (leu a senha na etiqueta)")

    add_heading(doc, "O censo — 15 min", level=2)
    add_para(doc, "Cada aluno diz o que trouxe da tarefa e você marca no quadro, numa tabela de três colunas: marca e modelo · Wi-Fi com nome de fábrica? · alguém sabe a senha de administrador?")
    add_bullet(doc, "\"Quantas casas estão com o nome de fábrica?\" (costuma ser metade ou mais)")
    add_bullet(doc, "\"Se o nome é de fábrica, a senha provavelmente também é. Onde ela está escrita?\" (na etiqueta, embaixo do roteador: qualquer visita lê)")
    add_bullet(doc, "\"Quantos sabem a senha de administrador?\" (quase ninguém. Muitos nunca souberam que existe)")
    add_bullet(doc, "\"Quantas marcas diferentes apareceram?\" (o painel muda de marca pra marca, mas os itens são os mesmos)")
    add_para(doc, "\"Esse quadro é o mercado de vocês. Cada casa com roteador de fábrica é um serviço de meia hora: configurar, deixar seguro e entregar a ficha. Hoje vocês aprendem o serviço inteiro.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — O QUE SE CONFIGURA E POR QUÊ · 15 MIN")
    add_para(doc, "Projete a tabela e passe item por item. Pra cada um, pergunte primeiro \"o que acontece se o técnico NÃO fizer isso?\" e deixe a turma responder.")
    add_table(doc, [["Item", "Por quê"]] + [list(r) for r in ENTREGA], col_widths=[5, 11.5])
    add_heading(doc, "Uma palavra sobre a ligação com a internet", level=2)
    add_para(doc, "Na tela de WAN (ou Internet) o roteador pergunta como ele recebe a internet. Na maioria das casas, atrás da ONU do provedor, é \"IP dinâmico\" e não se mexe. Se o provedor deixou a ONU em modo ponte, é PPPoE: precisa de usuário e senha que só o provedor dá. Sem esses dados, não mexa nessa tela. Hoje os roteadores estão sem internet, então a turma só olha qual tipo está marcado.", italic=True)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — PRÁTICA · NOME, SENHA E TRANCA · 35 MIN")
    add_para(doc, "Cada equipe no seu roteador, pelo celular de um dos alunos (o \"configurador\"). Troque o configurador a cada bloco, pra todo mundo pôr a mão. Página 1 da folha.")
    add_heading(doc, "Passo a passo", level=2)
    add_bullet(doc, "1. Conectar no Wi-Fi de fábrica (nome e senha na etiqueta) e abrir 192.168.0.1 no navegador")
    add_bullet(doc, "2. Criar a senha de administrador. Desta vez a EQUIPE escolhe (no mínimo 10 caracteres) e anota na folha. Se abrir assistente de configuração, pule")
    add_bullet(doc, "3. Olhar a tela WAN / Internet e anotar o tipo marcado. Não mexer")
    add_bullet(doc, "4. Sem fio → trocar o nome do Wi-Fi para EQUIPE N - SOBRENOME DE ALGUÉM (ex.: EQUIPE 3 - SCHNEIDER). Sem acento")
    add_bullet(doc, "5. Segurança sem fio → WPA2-PSK (ou WPA2/WPA3), criptografia AES, e uma senha nova de pelo menos 12 caracteres, que a equipe inventa. Salvar")
    add_bullet(doc, "6. O celular cai. É esperado. Conectar de novo, agora no nome novo com a senha nova. Todos os celulares da equipe fazem o mesmo")
    add_para(doc, "Quando o celular cair no passo 6, pare a turma: \"Isso que aconteceu aqui acontece na casa do cliente com TODOS os aparelhos: celular de cada um, TV, impressora, câmera. Antes de trocar a senha, combine com o cliente quem vai reconectar o quê. Senão vocês vão embora e a TV fica sem internet.\"", italic=True, color=COR_PRIMARIA)
    add_para(doc, "Equipe que se trancou pra fora (esqueceu a senha, salvou errado): botão Reset e começa de novo. Diga que isso também é aula: \"agora vocês sabem quanto custa não anotar\".", italic=True, color=COR_CINZA)
    add_para(doc, "Onde dilatar: peça que cada equipe esconda o nome do Wi-Fi (\"Ocultar SSID\" / desmarcar \"Transmitir\"), veja a rede sumir da lista do celular e conecte digitando o nome na mão. Depois volte a mostrar. Discuta: esconder o nome protege? (Pouco: qualquer aplicativo de Wi-Fi acha. E dá trabalho pro cliente.)", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · MUDAR A PORTARIA DE LUGAR · 15 MIN")
    add_para(doc, "\"Na casa do cliente, a ONU do provedor muitas vezes já usa 192.168.0.1 ou 192.168.1.1. Se o roteador usar o mesmo número, dá conflito: tudo parece ligado e nada funciona. Por isso o técnico muda o endereço do painel do roteador.\"", italic=True)
    add_bullet(doc, "Rede → LAN → endereço IP: trocar pelo número da sua equipe (tabela na página 1: equipe 1 = 192.168.11.1, equipe 2 = 192.168.12.1…). Salvar")
    add_bullet(doc, "O roteador reinicia e o painel some do navegador. É esperado")
    add_bullet(doc, "Desafio: reencontrar o roteador SEM olhar a tabela. No celular: Wi-Fi → (a rede da equipe) → detalhes → \"Gateway\" ou \"Roteador\". No computador por cabo: ipconfig → \"Gateway Padrão\"")
    add_bullet(doc, "Abrir o painel no endereço novo e anotar na folha")
    add_para(doc, "\"O gateway é o endereço da portaria. Lembram da Noite 6? Quando o técnico não sabe o endereço do painel de um roteador, é no gateway que ele olha. Serve pra qualquer marca.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — PRÁTICA · ENDEREÇOS E A AGENDA · 25 MIN")
    add_heading(doc, "A faixa de endereços — 8 min", level=2)
    add_bullet(doc, "Tela do DHCP: a faixa já mudou sozinha junto com o painel (192.168.1X.100 a .199, mais ou menos). Ajustar pra .100 a .150, como na tabela")
    add_bullet(doc, "\"Cinquenta endereços dá pra casa inteira? Contem quantos aparelhos a casa de vocês tem.\" (Sobra: uma casa tem uns 8 a 15)")
    add_heading(doc, "A reserva — 10 min", level=2)
    add_para(doc, "\"A impressora pega um endereço hoje e outro amanhã, e o computador continua procurando ela no número antigo. Prometido na Noite 6: hoje a gente resolve com a reserva.\"", italic=True)
    add_bullet(doc, "Escolha um celular da equipe pra fazer o papel da impressora. Ache o MAC dele na lista de aparelhos conectados (o CPF, Noite 6)")
    add_bullet(doc, "Reserva de endereço → adicionar: o MAC desse celular → o endereço da reserva da equipe (192.168.1X.50). Salvar")
    add_bullet(doc, "No celular \"impressora\": desligar e ligar o Wi-Fi. Conferir nos detalhes da rede: ganhou o .50?")
    add_para(doc, "Atenção: celular novo usa \"MAC aleatório\" (privado) pra cada rede. Se a reserva não pegar, no celular: Wi-Fi → a rede → Privacidade → \"Usar MAC do dispositivo\". Reconecte e refaça a reserva com o MAC novo. Impressora e câmera não têm esse problema.", italic=True, color=COR_ALERTA)
    add_heading(doc, "A agenda (DNS) — 7 min", level=2)
    add_para(doc, "\"Na Noite 7 eu disse que ia mostrar onde se troca a agenda. É aqui.\"", italic=True)
    add_bullet(doc, "Na tela do DHCP, \"DNS primário\" e \"DNS secundário\" estão vazios ou 0.0.0.0: o roteador usa a agenda do provedor")
    add_bullet(doc, "Preencher 8.8.8.8 e 1.1.1.1 (as agendas públicas da Noite 7). Salvar")
    add_bullet(doc, "Conferir no computador por cabo: ipconfig /all → \"Servidores DNS\" mostra 8.8.8.8. O roteador está dizendo pro computador qual agenda usar")
    add_para(doc, "\"Quando um cliente reclamar que 'tem internet mas nenhum site abre', e os dois pings da Noite 7 apontarem pra agenda, é aqui que vocês trocam.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)
    add_heading(doc, "INTERVALO · 15 MIN", color=COR_CINZA)
    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — PRÁTICA · REDE DE CONVIDADOS · 15 MIN")
    add_bullet(doc, "Rede de convidados → ativar. Nome: EQUIPE N - VISITA. Senha própria, diferente da principal. Salvar")
    add_bullet(doc, "Deixar DESMARCADO \"permitir que convidados acessem a rede local\" (o nome varia)")
    add_bullet(doc, "Um aluno conecta o celular na rede de VISITA e tenta abrir o painel (o endereço novo da equipe). Não abre: a visita não enxerga a casa")
    add_bullet(doc, "Na lista de aparelhos do painel principal, o celular da visita aparece separado (ou nem aparece)")
    add_para(doc, "\"Pra quem serve? Visita, cliente de comércio, o amigo do filho, a TV e a câmera baratas que vocês não confiam. A senha da visita pode ir pra parede. A da casa, não.\"", italic=True, color=COR_PRIMARIA)
    add_para(doc, "Equipe com roteador sem rede de convidados: acompanha o teste na equipe vizinha e anota na folha. Diga que, na hora de recomendar roteador pro cliente, esse é um item a conferir.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — PRÁTICA · CÓPIA DE SEGURANÇA · 20 MIN")
    add_para(doc, "As equipes vêm uma de cada vez ao computador (ou fazem pelo celular, se o navegador deixar baixar e enviar arquivo). Enquanto esperam, adiantam a ficha de entrega (Bloco 8).")
    add_bullet(doc, "Ferramentas do sistema → Backup → salvar o arquivo da configuração. Renomear: equipeN-AAAA-MM-DD")
    add_bullet(doc, "Agora a prova: botão Reset no roteador. Tudo volta ao de fábrica (o Wi-Fi da etiqueta reaparece). A equipe vê o trabalho de uma hora sumir")
    add_bullet(doc, "Entrar pelo endereço de fábrica (192.168.0.1), criar uma senha de administrador qualquer, Backup e restauração → Restaurar → escolher o arquivo")
    add_bullet(doc, "O roteador reinicia e volta com tudo: nome, senha, endereço novo, reserva, convidados. Conferir conectando de novo")
    add_para(doc, "\"Raio na rede, cliente que apertou o Reset achando que ia consertar, roteador que trocou na garantia pelo mesmo modelo: com a cópia, o serviço volta em um minuto. Sem ela, vocês refazem tudo — e cobram de novo, ou de graça.\"", italic=True, color=COR_PRIMARIA)
    add_para(doc, "A cópia de segurança guarda as senhas. Ela fica com o cliente (pendrive, e-mail dele) ou com o técnico, nunca num computador de uso geral. Apague as cópias do computador da escola no fim da aula.", italic=True, color=COR_ALERTA)

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — PRÁTICA · FICHA E ETIQUETA DE ENTREGA · 15 MIN")
    add_para(doc, "Página 3 da folha: a ficha de entrega, no mesmo espírito da ficha técnica das noites anteriores. Uma ficha por equipe, preenchida a caneta, legível.")
    add_bullet(doc, "Tudo que foi configurado vai na ficha: endereço do painel, senhas, nomes, faixa, reserva, DNS, onde está a cópia")
    add_bullet(doc, "Na fita crepe, a etiqueta que fica NO roteador: \"Configurado por EQUIPE N em DD/MM · painel 192.168.1X.1\". SEM senha nenhuma")
    add_para(doc, "\"A ficha com as senhas fica com o cliente, guardada. Na etiqueta do roteador vai só quem configurou e o endereço do painel: o próximo técnico agradece, e a visita não lê senha nenhuma.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 9
    add_heading(doc, "BLOCO 9 — PRÁTICA · O ROTEADOR TROCADO · 25 MIN")
    add_para(doc, "Cada equipe entrega o roteador e a ficha de entrega pra próxima (a 1 pra 2, a 2 pra 3… a última pra 1). A equipe que recebe faz de conta que é o próximo técnico a chegar naquela casa, um ano depois, só com o papel.")
    add_heading(doc, "A missão — 15 min", level=2)
    add_bullet(doc, "Conectar todos os celulares da equipe no Wi-Fi, usando só a ficha")
    add_bullet(doc, "Entrar no painel, usando só a ficha")
    add_bullet(doc, "Conferir item por item: o nome, a segurança, a faixa, a reserva, o DNS e a rede de convidados estão como a ficha diz?")
    add_bullet(doc, "Preencher o laudo na página 3: cada item ✓ ou ✗, e o que faltou ou estava errado na ficha")
    add_heading(doc, "A volta — 10 min", level=2)
    add_bullet(doc, "Devolve o roteador e o laudo pra equipe dona. Ela corrige a ficha")
    add_bullet(doc, "No quadro: quantas fichas funcionaram de primeira? O que mais faltou? (costuma ser letra ilegível, senha com 0 e O confundidos, endereço do painel antigo)")
    add_para(doc, "\"Ficha que só quem escreveu entende não é ficha. O cliente perde o papel, o técnico muda: o papel tem que funcionar sozinho.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 10
    add_heading(doc, "BLOCO 10 — FECHAMENTO E TAREFA · 10 MIN")
    add_heading(doc, "Quiz de saída", level=2)
    add_bullet(doc, "Você mudou o endereço do painel e perdeu o roteador. Onde acha o endereço novo?")
    add_bullet(doc, "Por que trocar a senha do Wi-Fi exige combinar com o cliente antes?")
    add_bullet(doc, "A impressora muda de endereço toda semana. O que você configura?")
    add_bullet(doc, "O que vai na etiqueta colada no roteador? E o que NÃO vai?")

    add_heading(doc, "A tarefa", level=2)
    add_para(doc, "\"Na próxima noite a gente mede até onde o Wi-Fi chega. A tarefa: descubram quantos mega é o plano de internet de casa (está na fatura ou no aplicativo do provedor). E, com o celular, rodem um teste de velocidade (fast.com ou o aplicativo do provedor) duas vezes: do lado do roteador e no cômodo mais longe dele. Anotem os três números.\"", italic=True)

    add_heading(doc, "Anuncie a próxima", level=2)
    add_para(doc, "\"Na próxima noite: até onde o Wi-Fi chega. Com estes mesmos roteadores, vocês medem o sinal pela escola, mudam o roteador de lugar e de faixa e veem o que melhora.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "SE PERGUNTAREM")
    add_para(doc, "\"Posso fazer isso no roteador do provedor?\"", bold=True)
    add_para(doc, "> \"O aparelho é do provedor, emprestado. Nome e senha do Wi-Fi, quase sempre pode, e às vezes o próprio aplicativo do provedor faz. O resto, combine com o provedor antes: tem provedor que reseta remoto e apaga o que vocês fizeram.\"", italic=True)
    add_para(doc, "\"Esqueci a senha de administrador do roteador de casa. E agora?\"", bold=True)
    add_para(doc, "> \"Reset de 10 segundos e configura de novo. Por isso a cópia de segurança e a ficha. Antes de resetar, anote o nome e a senha do Wi-Fi, pra configurar igual e os aparelhos da casa reconectarem sozinhos.\"", italic=True)
    add_para(doc, "\"WPA3 é melhor que WPA2?\"", bold=True)
    add_para(doc, "> \"É, mas aparelho antigo não conecta em WPA3. Por isso muitos roteadores têm o modo misto WPA2/WPA3. O que nunca se usa é WEP nem rede aberta.\"", italic=True)
    add_para(doc, "\"Posso usar o mesmo nome e senha do roteador antigo no novo?\"", bold=True)
    add_para(doc, "> \"Pode, e é um truque bom: configurando igualzinho, todos os aparelhos da casa conectam sozinhos no novo, sem ninguém mexer na TV e na impressora.\"", italic=True)
    add_para(doc, "\"E se o cliente quiser a senha do Wi-Fi fácil de lembrar?\"", bold=True)
    add_para(doc, "> \"Fácil de lembrar não é o mesmo que curta. Uma frase com quatro palavras é longa e fácil. Na Noite 10 a gente vê isso com calma.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "Apague do computador da escola os arquivos de cópia de segurança (têm senhas)")
    add_bullet(doc, "NÃO dê Reset nos roteadores: a Noite 9 usa eles configurados. Guarde cada roteador com a ficha de entrega da equipe, num envelope ou preso com elástico")
    add_bullet(doc, "Anote quantas fichas funcionaram de primeira no Bloco 9. É o indicador da noite")

    add_hr(doc)

    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Desenvolvi atividade prática de configuração de roteador doméstico a partir das configurações de fábrica. Partimos do levantamento, feito pelos estudantes, da situação dos roteadores das próprias residências. Apresentei os itens de uma configuração de entrega e sua finalidade. Em equipes, os estudantes criaram senha de administração, configuraram nome de rede, senha e criptografia WPA2, alteraram o endereço de acesso ao equipamento e o localizaram novamente pelo gateway, ajustaram a faixa do servidor DHCP, fizeram reserva de endereço por MAC, definiram os servidores DNS entregues pelo roteador e verificaram o resultado no computador, e criaram rede de convidados isolada da rede local. Salvaram a configuração em arquivo, restauraram o equipamento após reset e registraram o serviço em ficha de entrega. Ao final, cada equipe validou a ficha e a configuração de outra equipe, registrando laudo.")

    add_hr(doc)

    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou pra usar em outra noite:", italic=True, color=COR_CINZA)
    add_para(doc, "1. Configurar sem olhar o roteiro", bold=True)
    add_para(doc, "A equipe dá Reset e configura tudo de novo, agora sem a folha, só com a ficha de entrega que ela mesma fez. Cronometre. A meta do técnico é fazer o serviço inteiro em menos de 20 minutos. 25 min.")
    add_para(doc, "2. O mesmo serviço em outra marca", bold=True)
    add_para(doc, "Se alguém tiver um roteador de outra marca em casa (ou a escola tiver), a equipe tenta achar os oito itens no painel dele, só com a tabela do Bloco 2. Mostra que o serviço é o mesmo, só muda o nome do menu. 20 min.")
    add_para(doc, "3. O guia do cliente", bold=True)
    add_para(doc, "Cada equipe escreve, em linguagem de vó, um bilhete pra deixar na casa do cliente: \"como conectar um aparelho novo no Wi-Fi\" e \"como passar o Wi-Fi de visita\". Leitura pra turma. 15 min.")
    add_para(doc, "4. Horário do Wi-Fi", bold=True)
    add_para(doc, "Nos modelos que têm (Agenda do Wi-Fi / Controle de horário), a equipe programa o Wi-Fi pra desligar de madrugada e discute quando isso é útil (economia, filhos) e quando atrapalha (câmera, alarme). 15 min.")
    add_para(doc, "5. Atualização do roteador", bold=True)
    add_para(doc, "No painel: Ferramentas do sistema → Atualização de firmware. A equipe anota a versão instalada e o modelo com a versão de hardware da etiqueta (V4, V6…). Discuta por que nunca se atualiza com o roteador ligado por Wi-Fi nem com risco de faltar luz. Não precisa atualizar de verdade. 10 min.")

    path = os.path.join(OUTDIR, "Roteiro - Redes 8 - Configurando o roteador.docx")
    doc.save(path)
    return path


# ============================================================
# SLIDES
# ============================================================
def build_slides():
    prs = new_presentation()
    P, D, K, R = PPalette.P, PPalette.D, PPalette.K, PPalette.R

    def lista(kicker, titulo, linhas, subtitulo=None, size=20, nota=None, nota_cor=None):
        s = add_slide(prs)
        add_title_block(s, kicker, titulo, subtitulo)
        add_text(s, linhas, 0.7, 2.8, 11.9, 3.8, size=size, color=K)
        if nota:
            add_text(s, nota, 0.7, 6.4, 11.9, 0.5, size=16, italic=True, color=nota_cor or D, align=PP_ALIGN.CENTER)
        add_footer(s, FOOTER)
        return s

    def frase(kicker, texto, cor=None, size=36):
        s = add_slide(prs)
        add_text(s, kicker, 0.7, 1.6, 11.9, 0.5, size=16, bold=True, color=D, align=PP_ALIGN.CENTER)
        add_text(s, texto.split("\n"), 0.9, 2.4, 11.5, 3.0, size=size, bold=True, color=cor or P, align=PP_ALIGN.CENTER)
        add_footer(s, FOOTER)
        return s

    def tabela(kicker, titulo, rows, col_widths=None, subtitulo=None, size=16, nota=None, height=3.6):
        s = add_slide(prs)
        add_title_block(s, kicker, titulo, subtitulo)
        add_table_slide(s, rows, top=2.6, height=height, col_widths=col_widths, size=size)
        if nota:
            add_text(s, nota, 0.7, 6.5, 11.9, 0.5, size=15, italic=True, color=D, align=PP_ALIGN.CENTER)
        add_footer(s, FOOTER)
        return s

    add_slide_cover(prs, "NOITE 8", "CONFIGURANDO", "O ROTEADOR",
                    "Da caixa à entrega",
                    "Redes de Computadores · Curso FIC · CEJA Itapiranga – SC")

    lista("RETOMADA", "Das Noites 6 e 7 pra hoje", [
        "•  Qual o endereço da portaria (o painel) dos roteadores da bancada?",
        "•  Quem entrega o endereço pro celular?",
        "•  O DNS faz o quê?",
        "•  Como o intruso entrou no Wi-Fi da equipe na Noite 6?",
    ])

    lista("O CENSO", "O roteador da sua casa", [
        "•  Marca e modelo",
        "•  O Wi-Fi ainda tem o nome de fábrica?",
        "•  Alguém sabe a senha de administrador?",
        "",
        "Se o nome é de fábrica, a senha também é. E está escrita embaixo do roteador.",
    ])

    frase("A FRASE DA NOITE", "Roteador de fábrica funciona.\nRoteador configurado\né serviço entregue.")

    tabela("O QUE SE CONFIGURA", "Oito itens da entrega", [["Item", "Por quê"]] + [list(r) for r in ENTREGA],
           col_widths=[3.9, 8.0], size=13, height=4.0,
           nota="Pra cada um: o que acontece se o técnico NÃO fizer?")

    lista("A LIGAÇÃO COM A INTERNET", "A tela WAN: só olhe", [
        "•  Atrás da ONU do provedor: quase sempre \"IP dinâmico\". Não mexa",
        "•  ONU em modo ponte: PPPoE, com usuário e senha que só o provedor dá",
        "•  Sem esses dados, não mexa nessa tela",
        "",
        "Hoje os roteadores estão sem internet: anote só o tipo marcado.",
    ], size=20)

    lista("PRÁTICA · 35 MIN", "Nome, senha e tranca", [
        "1.  Wi-Fi de fábrica (etiqueta) → abrir 192.168.0.1",
        "2.  Senha de administrador: a equipe escolhe (10+ caracteres) e anota",
        "3.  Tela WAN: anotar o tipo. Não mexer",
        "4.  Nome do Wi-Fi: EQUIPE N - SOBRENOME (sem acento)",
        "5.  Segurança: WPA2-PSK, AES, senha nova de 12+ caracteres",
        "6.  O celular cai. Reconectar todo mundo no nome novo",
    ], size=19, nota="Troque o configurador a cada bloco: todo mundo põe a mão.")

    frase("O CELULAR CAIU", "Na casa do cliente caem TODOS:\ncelulares, TV, impressora, câmera.\nCombine antes quem reconecta o quê.", size=30)

    lista("PRÁTICA · 15 MIN", "Mudar a portaria de lugar", [
        "•  A ONU do provedor muitas vezes já usa 192.168.0.1 ou 192.168.1.1",
        "•  Mesmo número no roteador = conflito: tudo ligado, nada funciona",
        "•  Rede → LAN → o endereço da sua equipe (página 1)",
        "•  O painel some. Agora ache de novo, SEM olhar a tabela",
    ], size=20)

    frase("ONDE ESTÁ O PAINEL?", "No gateway.\nCelular: Wi-Fi → detalhes → Gateway\nComputador: ipconfig → Gateway Padrão", size=30)

    tabela("O PLANO DE ENDEREÇOS", "Cada equipe com o seu", [["Equipe", "Painel", "Faixa", "Reserva"]] +
           [[str(n), plano_equipe(n)["painel"], plano_equipe(n)["faixa"], plano_equipe(n)["reserva"]] for n in range(1, 7)],
           col_widths=[1.6, 3.2, 4.6, 2.5], size=15, height=3.8)

    lista("PRÁTICA · 25 MIN", "Endereços e a agenda", [
        "•  DHCP: faixa .100 a .150",
        "•  Reserva: o MAC de um celular (a \"impressora\") → o endereço .50",
        "•  Desliga e liga o Wi-Fi do celular: ganhou o .50?",
        "•  DNS primário 8.8.8.8 · secundário 1.1.1.1",
        "•  No computador: ipconfig /all → Servidores DNS",
    ], size=20, nota="Celular novo usa MAC aleatório: Privacidade → \"Usar MAC do dispositivo\".", nota_cor=R)

    frase("A PROMESSA DA NOITE 7", "Onde se troca a agenda?\nNo roteador.\nEle diz pros aparelhos qual DNS usar.", size=32)

    frase("", "INTERVALO\n15 minutos", cor=K)

    lista("PRÁTICA · 15 MIN", "Rede de convidados", [
        "•  Ativar: EQUIPE N - VISITA, com senha própria",
        "•  Deixar DESMARCADO \"convidados acessam a rede local\"",
        "•  Um celular na rede de VISITA tenta abrir o painel",
        "",
        "Não abre? A visita navega, mas não enxerga a casa.",
    ], size=21)

    lista("PRÁTICA · 20 MIN", "Cópia de segurança", [
        "1.  Backup → salvar o arquivo: equipeN-data",
        "2.  Reset no roteador. Tudo some",
        "3.  192.168.0.1 → senha qualquer → Restaurar → o arquivo",
        "4.  Reiniciou: voltou tudo?",
    ], size=22, nota="O arquivo guarda as senhas: fica com o cliente, nunca em computador de uso geral.", nota_cor=R)

    frase("POR QUE A CÓPIA", "Raio, Reset sem querer,\nroteador trocado na garantia:\ncom a cópia, volta em um minuto.", size=32)

    tabela("A FICHA DE ENTREGA", "Fica com o cliente", [["O que vai na ficha", ""]] +
           [[FICHA[i], FICHA[i + 6]] for i in range(6)],
           col_widths=[5.95, 5.95], size=14, height=3.6,
           nota="Na etiqueta do roteador: quem configurou, data e painel. Nenhuma senha.")

    lista("PRÁTICA · 25 MIN", "O roteador trocado", [
        "Vocês recebem o roteador e a ficha de outra equipe.",
        "São o próximo técnico, um ano depois, só com o papel.",
        "",
        "•  Conectar todos os celulares, só com a ficha",
        "•  Entrar no painel, só com a ficha",
        "•  Conferir cada item. Laudo na página 3",
    ], size=21)

    frase("A LIÇÃO DA FICHA", "Ficha que só quem escreveu entende\nnão é ficha.\nO papel tem que funcionar sozinho.", size=30)

    lista("O QUE FICA", "Três frases pra guardar", [
        "\"Roteador de fábrica funciona. Configurado é serviço entregue.\"",
        "",
        "\"Perdeu o painel? Olhe o gateway.\"",
        "",
        "\"Senha vai na ficha do cliente. Nunca na etiqueta do roteador.\"",
    ], size=22)

    lista("PARA A PRÓXIMA NOITE", "Até onde o Wi-Fi chega", [
        "A tarefa:",
        "",
        "•  Quantos mega é o plano de internet de casa? (fatura ou aplicativo)",
        "•  Teste de velocidade no celular (fast.com) do lado do roteador",
        "•  E de novo no cômodo mais longe dele",
        "•  Anote os três números",
    ])

    path = os.path.join(OUTDIR, "Aula - Redes 8 - Configurando o roteador.pptx")
    prs.save(path)
    return path


# ============================================================
# FOLHA DO ALUNO
# ============================================================
def campo(doc, texto, size=11, space_after=4, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(texto)
    r.font.size = Pt(size)
    r.bold = bold
    return p


def linha(doc):
    p = campo(doc, "_" * 100, space_after=0)
    p.runs[0].font.color.rgb = COR_CINZA


def build_folha():
    doc = Document()
    setup_page(doc, margem=Cm(1.8))

    # PÁGINA 1 — Nome, senha, painel
    cabecalho_documento(
        doc,
        "CONFIGURANDO O ROTEADOR",
        f"Redes de Computadores · {NOITE} · Curso FIC · CEJA Itapiranga · 1 de 3",
    )
    campo(doc, "Nome: _________________________________________________     Equipe: ________", space_after=8)

    add_heading(doc, "1  A tarefa: o roteador da minha casa", level=2)
    campo(doc, "Marca e modelo: __________________________   Wi-Fi com nome de fábrica?  ☐ Sim  ☐ Não")
    campo(doc, "Alguém sabe a senha de administrador?  ☐ Sim  ☐ Não")

    add_heading(doc, "2  Nome, senha e tranca", level=2)
    add_para(doc, "Anote a lápis. No fim, passe a limpo na ficha de entrega (página 3).", italic=True, size=10, color=COR_CINZA)
    tbl = add_table(doc, [
        ["O quê", "De fábrica (etiqueta)", "Como a equipe deixou"],
        ["Endereço do painel", "192.168.0.1", ""],
        ["Senha de administrador", "(nenhuma)", ""],
        ["Nome do Wi-Fi", "", ""],
        ["Senha do Wi-Fi", "", ""],
        ["Segurança", "", "☐ WPA2-PSK  ☐ WPA2/WPA3   AES"],
        ["Tela WAN: tipo marcado", "", "(não mexer)"],
    ], col_widths=[4.6, 5.2, 7.6])
    for row in tbl.rows[1:]:
        row.height = Cm(0.85)
    campo(doc, "Quantos celulares da equipe reconectaram no nome novo? ____   Algum não conseguiu? Por quê? ____________")

    add_heading(doc, "3  Mudar a portaria de lugar", level=2)
    tbl = add_table(doc, [["Equipe", "Novo endereço do painel", "Faixa de endereços", "Reserva"]] +
                    [[str(n), plano_equipe(n)["painel"], plano_equipe(n)["faixa"], plano_equipe(n)["reserva"]] for n in range(1, 7)],
                    col_widths=[2, 4.8, 6.6, 4])
    for row in tbl.rows:
        for cell in row.cells:
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(9)
    campo(doc, "Depois de mudar, onde a equipe achou o endereço novo?  ☐ gateway no celular  ☐ ipconfig no computador")
    campo(doc, "O gateway que apareceu: ____________________   Bate com o endereço da sua equipe?  ☐ Sim  ☐ Não")
    campo(doc, "Por que o técnico muda o endereço do painel na casa do cliente?", bold=True)
    linha(doc)
    linha(doc)

    # PÁGINA 2 — Endereços, agenda, convidados, cópia
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "ENDEREÇOS, AGENDA, VISITA E CÓPIA",
        f"Configurando o roteador · {NOITE} · 2 de 3",
    )

    add_heading(doc, "1  A faixa de endereços (DHCP)", level=2)
    campo(doc, "A faixa estava: de ____________________ até ____________________")
    campo(doc, "A equipe deixou: de ____________________ até ____________________   (são ______ endereços)")

    add_heading(doc, "2  A reserva", level=2)
    campo(doc, "Aparelho que fez de \"impressora\": ______________________   MAC: ____ : ____ : ____ : ____ : ____ : ____")
    campo(doc, "Endereço reservado: ____________________   Depois de reconectar, ele ganhou esse endereço?  ☐ Sim  ☐ Não")
    campo(doc, "Teve que desligar o MAC aleatório (privado) do celular?  ☐ Sim  ☐ Não")

    add_heading(doc, "3  A agenda (DNS)", level=2)
    campo(doc, "Antes, os campos de DNS estavam: ____________________   A equipe colocou: ____________ e ____________")
    campo(doc, "No computador, ipconfig /all → Servidores DNS mostrou: ____________________")
    campo(doc, "Que reclamação de cliente faz você vir até esta tela? (lembre da Noite 7)", bold=True)
    linha(doc)

    add_heading(doc, "4  Rede de convidados", level=2)
    campo(doc, "O roteador tem rede de convidados?  ☐ Sim  ☐ Não (acompanhamos a equipe ______)")
    campo(doc, "Nome: ____________________   Senha: ____________________")
    campo(doc, "Pela rede de VISITA, o painel abriu?  ☐ Sim  ☐ Não     O que isso quer dizer? ________________________")

    add_heading(doc, "5  Cópia de segurança", level=2)
    tbl = add_table(doc, [
        ["Passo", "Feito?"],
        ["Salvamos o arquivo da configuração (nome: ______________________)", "☐"],
        ["Demos Reset: o Wi-Fi de fábrica voltou", "☐"],
        ["Entramos em 192.168.0.1 e restauramos o arquivo", "☐"],
        ["Voltou tudo: nome, senha, painel, reserva, convidados", "☐"],
    ], col_widths=[15, 2.5])
    for row in tbl.rows[1:]:
        row.height = Cm(0.75)
    campo(doc, "Onde a cópia de segurança deve ficar? E onde NÃO deve?", bold=True)
    linha(doc)

    # PÁGINA 3 — Ficha de entrega + laudo do roteador trocado + tarefa
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "FICHA DE ENTREGA E LAUDO",
        f"O serviço documentado · {NOITE} · 3 de 3",
    )

    add_heading(doc, "1  Ficha de entrega ao cliente (a caneta, letra legível)", level=2)
    tbl = add_table(doc, [["Item", "Como ficou"]] + [[c, ""] for c in FICHA], col_widths=[6.5, 11])
    for row in tbl.rows[1:]:
        row.height = Cm(0.72)
    add_para(doc, "Etiqueta no roteador (fita crepe): \"Configurado por EQUIPE __ em __/__ · painel 192.168.___.1\". Nenhuma senha.", italic=True, size=9, color=COR_ALERTA)

    add_heading(doc, "2  O roteador trocado: laudo da ficha da equipe ____", level=2)
    tbl = add_table(doc, [
        ["Conferido só com a ficha", "✓ ou ✗", "O que faltou ou estava errado"],
        ["Todos conectaram no Wi-Fi", "", ""],
        ["Entramos no painel", "", ""],
        ["Nome e segurança como na ficha", "", ""],
        ["Faixa e reserva como na ficha", "", ""],
        ["DNS como na ficha", "", ""],
        ["Rede de convidados como na ficha", "", ""],
    ], col_widths=[6.5, 2, 9])
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)
    campo(doc, "A ficha funcionou sozinha?  ☐ Sim, de primeira  ☐ Com ajuda da outra equipe  ☐ Não", size=10)

    add_heading(doc, "TAREFA DA SEMANA", level=2, color=COR_DESTAQUE)
    campo(doc, "Plano de internet de casa: ______ mega     Teste de velocidade (fast.com):", size=10)
    campo(doc, "Do lado do roteador: ______ mega     No cômodo mais longe (qual? __________): ______ mega", size=10)
    add_para(doc, "Traga esta folha na Noite 9 · CEJA Itapiranga – SC", italic=True, size=9, color=COR_CINZA, align=WD_ALIGN_PARAGRAPH.CENTER)

    path = os.path.join(OUTDIR, "Folha do Aluno - Redes 8 - Configurando o roteador.docx")
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
        "CONFIGURANDO O ROTEADOR: DA CAIXA À ENTREGA",
        f"Resumo do aluno · {NOITE} · Redes de Computadores · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "O roteador sai da caixa funcionando, mas com nome, senha e painel que qualquer um conhece. Configurar é deixar ele seguro e organizado. Entregar é deixar a ficha com o cliente, pra ele e pro próximo técnico.")

    add_heading(doc, "Os oito itens da entrega", level=2)
    add_table(doc, [["Item", "Por quê"]] + [list(r) for r in ENTREGA], col_widths=[5, 10.5])

    add_heading(doc, "Onde fica cada coisa (TP-Link)", level=2)
    add_table(doc, [["O quê", "Onde costuma ficar"]] + [list(r) for r in MENUS], col_widths=[4.5, 11])
    add_para(doc, "Outra marca? Os itens são os mesmos, só muda o nome do menu.", italic=True, size=10, color=COR_CINZA)

    add_heading(doc, "Na hora do serviço", level=2)
    add_bullet(doc, "Antes de trocar a senha do Wi-Fi, combine com o cliente: TODOS os aparelhos vão cair e precisar reconectar")
    add_bullet(doc, "Truque: no roteador novo, use o mesmo nome e senha do antigo. Os aparelhos reconectam sozinhos")
    add_bullet(doc, "Perdeu o painel depois de mudar o endereço? Olhe o gateway: no celular (detalhes da rede) ou no computador (ipconfig)")
    add_bullet(doc, "Reserva de endereço não pegou num celular? Ele está com MAC aleatório: Wi-Fi → a rede → Privacidade → \"Usar MAC do dispositivo\"")
    add_bullet(doc, "Tela WAN: atrás da ONU, quase sempre \"IP dinâmico\". PPPoE só com usuário e senha do provedor")
    add_bullet(doc, "Nunca deixe o Wi-Fi aberto nem em WEP. WPA2, ou WPA2/WPA3 se todos os aparelhos aceitarem")

    add_heading(doc, "Ficha, etiqueta e cópia", level=2)
    add_bullet(doc, "Ficha de entrega com tudo, inclusive as senhas: fica guardada com o cliente")
    add_bullet(doc, "Etiqueta no roteador: quem configurou, a data e o endereço do painel. Nenhuma senha")
    add_bullet(doc, "Cópia de segurança: salva a configuração inteira num arquivo. Depois de um Reset, restaura em um minuto. Guarda as senhas: fica com o cliente ou com o técnico")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Roteador de fábrica funciona. Configurado é serviço entregue.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Perdeu o painel? Olhe o gateway.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Senha vai na ficha do cliente. Nunca na etiqueta do roteador.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Ache o gateway da sua rede de casa no celular. É o endereço do painel do seu roteador")
    add_bullet(doc, "Se você fosse configurar o roteador da sua casa amanhã, o que mudaria primeiro? E com quem precisaria combinar?")
    add_bullet(doc, "Por que a ficha de entrega não pode depender da memória de quem configurou?")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Este material é seu. Guarde e volte quando precisar. · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, "Resumo do Aluno - Redes 8 - Configurando o roteador.docx")
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
