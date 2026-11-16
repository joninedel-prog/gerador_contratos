"""Gera os 4 arquivos do Encontro 1 do módulo de Servidores (Plano Mensal 4).

Servidores 1 — Instalando o Ubuntu Server e o primeiro site
Primeiro contato da turma com Linux. Os ALUNOS instalam o Ubuntu Server 24.04
nas máquinas antigas, em equipes (uma por máquina), ligadas aos roteadores
configurados na Noite 8. Depois entram no sistema pela primeira vez, tiram o
"raio-x" da máquina com comandos simples, instalam o Nginx e fazem uma página
da equipe que abre no celular. No fim, visitam o site das outras equipes
conectando no Wi-Fi delas com a ficha de entrega da Noite 8.
Padrão prática-pesada, sem role-play e sem duelo. Precisa de internet.
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from common import (
    COR_PRIMARIA, COR_ALERTA, COR_DESTAQUE, COR_CINZA,
    add_heading, add_para, add_bullet, add_table, add_hr,
    setup_page, cabecalho_documento,
    PPalette, new_presentation, add_slide, add_text, add_title_block,
    add_slide_cover, add_footer, add_table_slide,
)
from pptx.enum.text import PP_ALIGN

OUTDIR = "/home/nedel/curso"
ENCONTRO = "Servidores 1"
TITULO_ARQ = "Servidores 1 - Instalando o Ubuntu e o primeiro site"
FOOTER = f"Servidores com Linux · Encontro 1"
MONO = "Courier New"


# As telas do instalador (Ubuntu Server 24.04), na ordem em que aparecem
TELAS = [
    ("Menu do pendrive (GNU GRUB)", "Começar", "Try or Install Ubuntu Server → Enter"),
    ("Welcome! (idioma)", "Idioma do instalador", "English"),
    ("Installer update available", "Tem instalador mais novo", "Continue without updating"),
    ("Keyboard configuration", "Teclado", "Portuguese (Brazil) nas duas linhas"),
    ("Choose type of installation", "Tipo de instalação", "Ubuntu Server (não o \"minimized\")"),
    ("Network configuration", "Rede", "Não mexer. ANOTAR o número do DHCPv4"),
    ("Proxy configuration", "Atalho de empresa", "Deixar em branco → Done"),
    ("Ubuntu archive mirror", "De onde baixar programas", "Esperar \"passed tests\" → Done"),
    ("Guided storage configuration", "Disco", "Use an entire disk → Done"),
    ("Storage configuration", "Resumo do disco", "Done → Continue  (APAGA TUDO)"),
    ("Profile configuration", "Nome, servidor, usuário, senha", "Ver o combinado da equipe e ANOTAR"),
    ("Upgrade to Ubuntu Pro", "Versão paga", "Skip for now"),
    ("SSH configuration", "Acesso de longe", "MARCAR Install OpenSSH server (Espaço)"),
    ("Featured server snaps", "Programas extras", "Não marcar nada → Done"),
    ("Installation complete!", "Terminou", "Reboot Now → tirar o pendrive → Enter"),
]

# Raio-x da máquina: primeiros comandos
RAIOX = [
    ("whoami", "Quem está usando o servidor agora"),
    ("hostname", "O nome do servidor"),
    ("ip a", "O endereço do servidor (linha inet, fora do lo)"),
    ("ping -c 4 192.168.1N.1", "O servidor fala com o roteador da equipe?"),
    ("ping -c 4 google.com", "O servidor fala com a internet? (e a agenda, o DNS)"),
    ("free -h", "Quanta memória (RAM) tem e quanta está em uso"),
    ("df -h /", "Tamanho do disco, quanto está usado e quanto sobra"),
    ("clear", "Limpa a tela"),
]

# O primeiro site
SITE = [
    ("sudo apt update", "Atualiza a lista de programas (pede a senha)"),
    ("sudo apt install nginx", "Instala o servidor de páginas. Pergunta [Y/n]: Enter"),
    ("systemctl status nginx", "Está rodando? Procure active (running). Sai com q"),
    ("sudo nano /var/www/html/index.html", "Cria a página da equipe"),
    ("Ctrl+O, Enter · Ctrl+X", "No nano: salvar e sair"),
    ("sudo poweroff", "Desliga o servidor do jeito certo"),
]

# A página que a equipe digita (sem acento e sem ç hoje)
PAGINA = [
    "<h1>Servidor da Equipe N</h1>",
    "<p>Feito por: nomes da equipe</p>",
    "<p>Instalado em 27/11/2026 no CEJA Itapiranga</p>",
    "<p>Funcionando num computador antigo com Ubuntu Server</p>",
]

# Atividade da espera: cliente ou servidor?
CLIENTE_SERVIDOR = [
    "O celular de vocês", "O computador onde ficam os vídeos do YouTube",
    "O roteador entregando endereço (DHCP)", "A TV da sala assistindo Netflix",
    "O computador do banco que guarda o seu saldo", "A impressora compartilhada da escola",
    "O notebook que manda imprimir", "O computador do WhatsApp que guarda a mensagem até o amigo ligar o celular",
    "A câmera de segurança que grava e mostra no aplicativo", "O aplicativo da câmera no celular",
    "A agenda da internet que transforma nome em número (DNS)", "A máquina que vocês acabaram de instalar",
]

# Problemas comuns e o que fazer
TRAVOU = [
    ("Ligou e não aparece o menu do pendrive", "Tecla de boot do modelo (F12, F11, F8, Esc). Pendrive numa porta USB de trás. Na BIOS, pôr o USB primeiro"),
    ("Tela de rede sem número no DHCPv4", "Cabo frouxo ou na porta errada do roteador (tem que ser LAN, amarela/azul). Trocar o cabo. Sem rede não segue: o site precisa de internet"),
    ("Mirror não passa no teste", "Falta internet no roteador. Conferir o cabo da WAN do roteador até a rede da escola"),
    ("Erro de disco ou de memória", "Máquina não serve. A equipe vai pra máquina reserva"),
    ("Login incorrect", "Caps Lock, maiúscula, Num Lock. Senha não aparece: digite com calma e Enter"),
    ("Mensagens aparecem depois do login: prompt", "É o sistema terminando de ligar. Aperte Enter e o login: volta"),
    ("/ ou : não saem no teclado", "Teclado ficou errado: sudo dpkg-reconfigure keyboard-configuration"),
    ("Could not get lock (no apt)", "O servidor está se atualizando sozinho no primeiro boot. Esperar 5 minutos e repetir"),
    ("Celular não abre a página", "Celular no Wi-Fi da EQUIPE (não 4G), http:// e não https, número certo. systemctl status nginx"),
    ("Ainda aparece Welcome to nginx", "O arquivo não se chama index.html ou não está em /var/www/html. Recarregar a página"),
]

# Ficha técnica: os campos que começam hoje (continua nos próximos encontros)
FICHA = [
    "Equipe / máquina (marca e modelo)",
    "Nome do servidor",
    "Usuário",
    "Senha",
    "Endereço do servidor (IP)",
    "Ligado no roteador / Wi-Fi",
    "Memória (RAM) e disco",
    "Serviços instalados",
    "Como acessar o site",
    "Instalado por / data",
]


# ============================================================
# Helpers locais
# ============================================================
def shade_paragraph(p, color_hex="EEF3F3"):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    pPr.append(shd)


def add_cmd(doc, text, size=10.5):
    """Linha de comando em fonte de máquina com fundo cinza claro."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(text)
    r.font.name = MONO
    r._element.rPr.rFonts.set(qn('w:eastAsia'), MONO)
    r.font.size = Pt(size)
    shade_paragraph(p)
    return p


def mono_col(table, col, start_row=1, size=10):
    for row in table.rows[start_row:]:
        for run in row.cells[col].paragraphs[0].runs:
            run.font.name = MONO
            run.font.size = Pt(size)


def small_table(table, size=9):
    for row in table.rows:
        for cell in row.cells:
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(size)


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


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "INSTALANDO O UBUNTU E O PRIMEIRO SITE",
        f"Roteiro do professor · Servidores com Linux · Encontro 1 de 3 · sexta 27/11 · Curso FIC · CEJA Itapiranga – SC",
    )

    add_heading(doc, "A IDEIA DESTE ENCONTRO")
    add_para(doc, "É o primeiro contato da turma com Linux, e quase todos nunca viram uma tela só de texto. A noite tem que terminar com uma vitória que dá pra mostrar: cada equipe instala o Ubuntu Server numa máquina antiga com as próprias mãos e, no fim, abre no celular um site que está rodando naquela máquina.")
    add_para(doc, "O caminho é curto de propósito. Nada de partição, nada de teoria de sistema operacional. Instalar, entrar, perguntar cinco coisas pra máquina, instalar um programa e publicar uma página. O resto (pastas, usuários, permissões, acesso de longe, pasta compartilhada) vem nos encontros 2 e 3.")
    add_para(doc, "A frase da noite: \"Servidor não é um computador especial. É um computador que trabalha para os outros.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "O que se resolve neste encontro", level=2)
    add_para(doc, "Cada aluno sai sabendo diferenciar cliente de servidor com exemplos do dia a dia, dar boot por pendrive, passar pelas telas do instalador do Ubuntu Server entendendo cada escolha, entrar no sistema, descobrir o endereço, a memória e o disco da máquina, instalar um programa com o apt e publicar uma página que abre nos celulares da rede.")

    add_heading(doc, "Onde este encontro fica no módulo", level=2)
    add_para(doc, "Abre o Plano Mensal 4 (servidores com Linux). Usa os roteadores configurados na Noite 8 e a ficha de entrega deles. Retoma a Noite 1 (rede não é internet: o site só abre dentro da rede), a Noite 6 (endereço, gateway, ping) e a Noite 7 (o ping em google.com só funciona porque a agenda, o DNS, responde). O SSH já fica instalado hoje para o Encontro 2.")
    add_para(doc, "Não entre em: partições e LVM, interface gráfica, comparação entre distribuições, domínio na internet, redirecionamento de portas, HTML além das quatro linhas da página. Se alguém perguntar \"e pra abrir de casa?\", a resposta é \"dá, mas é outro serviço, com outros cuidados\" (ver Se perguntarem).", italic=True, color=COR_ALERTA)

    add_hr(doc)

    add_heading(doc, "ANTES DA AULA")
    add_heading(doc, "As máquinas (faça com antecedência)", level=2)
    add_bullet(doc, "Confira cada máquina: processador de 64 bits, pelo menos 2 GB de RAM e disco de pelo menos 25 GB. Máquina abaixo disso não aceita o 24.04 ou fica lenta demais")
    add_bullet(doc, "Confirme com a escola que NÃO há nada nos discos. A instalação apaga tudo")
    add_bullet(doc, "Descubra a tecla do menu de boot de cada modelo (F12, F11, F8, Esc) e cole numa fita crepe em cima da máquina: \"BOOT = F12\". Isso economiza 15 minutos de tentativa na aula")
    add_bullet(doc, "Faça UMA instalação completa antes, seguindo a folha do aluno, numa das máquinas. Cronometre: é o tempo que as equipes vão levar (em máquina antiga, de 20 a 40 minutos). Essa máquina fica pronta como RESERVA para a equipe cuja máquina falhar")
    add_bullet(doc, "Numere com fita crepe: SERVIDOR EQUIPE 1, SERVIDOR EQUIPE 2… Cada máquina com monitor, teclado (mouse não é usado), tomada e cabo de rede")

    add_heading(doc, "Os pendrives", level=2)
    add_bullet(doc, "Baixe a imagem ubuntu-24.04.x-live-server-amd64.iso em ubuntu.com/download/server")
    add_bullet(doc, "Grave com o Rufus (Windows) ou o balenaEtcher, um pendrive de 4 GB ou mais por equipe. Etiquete \"UBUNTU SERVER\"")
    add_bullet(doc, "Se tiver menos pendrives do que equipes: o pendrive fica preso na máquina até o fim da instalação. Comece as equipes em ondas e use a atividade de espera (Bloco 5) para quem ainda não começou")

    add_heading(doc, "A rede", level=2)
    add_bullet(doc, "Cada servidor vai num roteador de equipe da Noite 8, por cabo, numa porta LAN. A porta WAN (Internet) do roteador vai num ponto de rede da escola")
    add_bullet(doc, "Teste a internet de cada roteador com um celular antes da aula. Sem internet não instala o Nginx (há um plano B no Bloco 7)")
    add_bullet(doc, "Tenha as fichas de entrega da Noite 8 (nome e senha do Wi-Fi de cada roteador). Elas são usadas no Bloco 9")
    add_bullet(doc, "Se os roteadores ainda estão como a Noite 8 deixou, o servidor da equipe N vai ganhar um endereço 192.168.1N.1xx (equipe 3: 192.168.13.101, por exemplo)")
    add_para(doc, "Sem os roteadores da Noite 8? Todos os servidores podem ir num roteador só. Nesse caso o Bloco 9 fica mais simples (todos no mesmo Wi-Fi), mas perde a discussão de \"cada rede é uma casa\".", italic=True, color=COR_CINZA)

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "Folha do aluno (3 páginas), uma por pessoa")
    add_bullet(doc, "Resumo do aluno, uma por pessoa (pra levar pra casa)")
    add_bullet(doc, "Os slides ficam projetados durante a instalação: o slide das telas do instalador é o mapa da turma")

    add_hr(doc)

    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Abertura: o roteador já é um servidor", "15 min", "Cliente e servidor, a cozinha do restaurante, a demanda do escritório"],
        ["2. Linux, Ubuntu e o pendrive", "10 min", "Software livre, servidor sem janela, como se dá boot. Aviso: apaga o disco"],
        ["3. PRÁTICA — Boot e primeiras telas", "25 min", "Pendrive, menu de boot, idioma, teclado, tipo, rede (anotar o IP), mirror"],
        ["4. PRÁTICA — Disco, perfil e SSH", "20 min", "Disco inteiro, nome do servidor, usuário e senha, OpenSSH. Começa a cópia"],
        ["5. PRÁTICA — Enquanto copia", "20 min", "Cliente ou servidor? (12 casos) e leitura dos comandos da noite"],
        ["INTERVALO", "15 min", "Coincide com o fim da cópia nas máquinas mais lentas"],
        ["6. PRÁTICA — Primeiro login e raio-x", "25 min", "Entrar, a senha invisível, oito comandos, a ficha técnica"],
        ["7. PRÁTICA — O primeiro site", "30 min", "apt update, instalar o Nginx, abrir no celular"],
        ["8. PRÁTICA — A página da equipe", "25 min", "nano, quatro linhas, salvar, recarregar no celular"],
        ["9. PRÁTICA — Visita aos outros servidores", "15 min", "Conectar no Wi-Fi de outra equipe e abrir o site dela"],
        ["10. Desligar certo + fechamento", "10 min", "sudo poweroff, três frases, o que vem no Encontro 2"],
    ], col_widths=[6.5, 2, 8])
    add_para(doc, "Blocos 3 a 9: 160 min de atividade do aluno (76% do total). A exposição é só os Blocos 1 e 2 (25 min) e as falas curtas dentro das práticas.", italic=True, color=COR_CINZA)

    add_heading(doc, "As equipes e o rodízio", level=2)
    add_para(doc, "Uma equipe por máquina, de 3 ou 4. Três papéis que trocam a cada bloco (anotado no topo da página 1 da folha):")
    add_bullet(doc, "Digitador: o único com a mão no teclado")
    add_bullet(doc, "Leitor: lê a folha e a tela em voz alta, traduz o inglês, confere antes do Enter")
    add_bullet(doc, "Anotador: preenche a folha da equipe (IP, usuário, resultados)")
    add_para(doc, "A regra: o digitador não aperta Enter antes do leitor dizer \"confere\". Isso evita 90% dos erros de digitação e obriga todos a lerem a tela.", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — ABERTURA: O ROTEADOR JÁ É UM SERVIDOR · 15 MIN")
    add_heading(doc, "Comece pelo que eles já fizeram — 5 min", level=2)
    add_para(doc, "Segure um roteador da Noite 8 e pergunte:")
    add_bullet(doc, "\"Quem dá o endereço pro celular quando ele conecta?\" (o roteador, o DHCP)")
    add_bullet(doc, "\"Onde vocês entraram pra configurar ele?\" (numa página: 192.168.1N.1)")
    add_para(doc, "\"Então esse aparelhinho entrega endereço pra todo mundo e ainda tem uma página que abre no navegador. Ele trabalha para os outros aparelhos da casa. Isso tem nome: servidor. Vocês já configuraram um servidor sem saber. Hoje vocês vão transformar um computador velho em um servidor de verdade, com um site que abre no celular de vocês.\"", italic=True, color=COR_PRIMARIA)

    add_heading(doc, "A cozinha do restaurante — 5 min", level=2)
    add_para(doc, "Desenhe no quadro: mesas de um lado, cozinha do outro.")
    add_bullet(doc, "O cliente senta, PEDE e recebe. Celular, notebook e TV são clientes")
    add_bullet(doc, "A cozinha fica lá atrás, ninguém vê, e SERVE todas as mesas ao mesmo tempo. É o servidor")
    add_bullet(doc, "\"A foto do Instagram está no seu celular?\" (não: está num servidor, o celular só pede e mostra)")
    add_bullet(doc, "A cozinha não fecha. Servidor fica ligado o tempo todo, e quase sempre sem ninguém olhando a tela")

    add_heading(doc, "A demanda — 5 min", level=2)
    add_para(doc, "Leia a questão do plano: \"Um escritório pequeno quer guardar os arquivos num lugar só, acessar de qualquer computador e ter uma página interna para os funcionários. Como transformar um computador antigo no servidor que resolve isso?\"")
    add_para(doc, "\"Em três encontros a gente faz isso inteiro. Hoje: o sistema instalado e a página interna funcionando.\"", italic=True)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — LINUX, UBUNTU E O PENDRIVE · 10 MIN")
    add_bullet(doc, "Linux é um sistema, como o Windows. É gratuito, qualquer um pode usar e melhorar (software livre) e roda bem em máquina antiga. A maioria dos servidores da internet usa Linux. O Android também é Linux por dentro")
    add_bullet(doc, "Ubuntu é uma \"versão empacotada\" do Linux (uma distribuição). O Ubuntu Server não tem janelas nem mouse: só texto. Não precisa: ninguém fica olhando a tela de um servidor")
    add_bullet(doc, "\"Parece filme de hacker, mas é só escrever ordens. Cada ordem é uma palavra em inglês: ele faz o que vocês mandam, nem mais nem menos\"")
    add_bullet(doc, "24.04 quer dizer abril de 2024. É uma versão de suporte longo: recebe correções até 2029")
    add_heading(doc, "Como se instala", level=2)
    add_bullet(doc, "O pendrive tem o instalador. Liga a máquina e aperta a tecla de boot (está na fita crepe) para escolher o pendrive")
    add_bullet(doc, "O instalador é em inglês e não usa mouse: setas e Tab para andar, Espaço para marcar, Enter para confirmar. Esc volta uma tela")
    add_bullet(doc, "Mostre o slide das telas: são 15, a folha traduz cada uma")
    add_para(doc, "Fale sério, uma vez: \"Numa tela o instalador avisa que vai apagar o disco inteiro. Aqui as máquinas estão vazias. Na casa de um cliente, isso apaga as fotos da família. O técnico SEMPRE pergunta e faz cópia antes.\"", bold=True, color=COR_ALERTA)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — PRÁTICA · BOOT E PRIMEIRAS TELAS · 25 MIN")
    add_para(doc, "Página 1 da folha. As equipes vão no seu ritmo, você circula. Telas 1 a 8:")
    add_table(doc, [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS[:8]], col_widths=[5.2, 4.3, 7])
    add_heading(doc, "Pontos de parada", level=2)
    add_bullet(doc, "Teclado: se escolherem errado, depois / e : não saem no lugar certo e nenhum comando funciona. Peça que mostrem a tela antes do Done")
    add_bullet(doc, "Rede: pare cada equipe aqui. \"Esse número ao lado de DHCPv4 é o endereço do servidor. Quem deu esse número?\" (o roteador da equipe, o DHCP da Noite 8). Anotar sem o /24")
    add_bullet(doc, "Mirror: \"This mirror location passed tests\" quer dizer que o servidor alcançou a internet. Se não passar, o problema é a WAN do roteador")
    add_para(doc, "Onde dilatar: na tela de rede, peça que a equipe confira o número no painel do roteador (lista de clientes DHCP, Noite 6). O servidor aparece lá com o nome ubuntu-server ou sem nome. É o mesmo endereço?", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · DISCO, PERFIL E SSH · 20 MIN")
    add_table(doc, [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS[8:]], col_widths=[5.2, 4.3, 7])
    add_heading(doc, "O combinado do perfil (escreva no quadro)", level=2)
    add_table(doc, [
        ["Campo", "O que a equipe põe", "Regra"],
        ["Your name", "Equipe N", "Livre"],
        ["Your server's name", "servidor-equipeN", "Minúsculo, sem espaço, sem acento"],
        ["Pick a username", "equipeN", "Minúsculo, sem espaço, sem acento"],
        ["Choose a password", "A equipe inventa", "10 caracteres ou mais. Anotar na folha"],
    ], col_widths=[4.5, 4.5, 7.5])
    add_bullet(doc, "Storage: o padrão já vem certo (disco inteiro). O instalador mostra o resumo e pergunta \"Confirm destructive action\". Lembre a fala do Bloco 2 e deixe a equipe ler a palavra destructive em voz alta antes do Continue")
    add_bullet(doc, "SSH: no instalador, marcar é com Espaço, não Enter. Confira que aparece [X] Install OpenSSH server. É o que permite entrar no servidor pelo celular no Encontro 2")
    add_bullet(doc, "Depois dos snaps começa a cópia (Installing system). Em máquina antiga demora de 10 a 30 minutos. Se aparecer \"Cancel update and reboot\" depois de muito tempo baixando atualizações, pode usar: o servidor se atualiza depois")

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — PRÁTICA · ENQUANTO COPIA · 20 MIN")
    add_para(doc, "O anotador fica de olho na tela (chama quando aparecer Reboot Now). Os outros fazem a página 2 da folha:")
    add_bullet(doc, "Cliente ou servidor? Doze casos para marcar. Corrija em voz alta com a turma. Os casos que dão discussão: a impressora compartilhada (servidor de impressão), o WhatsApp (o servidor segura a mensagem quando o amigo está sem internet), a câmera (ela serve as imagens, o aplicativo é cliente)")
    add_bullet(doc, "Leitura dos comandos: a equipe lê a tabela do raio-x e escreve, antes de rodar, o que acha que cada comando vai mostrar")
    add_para(doc, "A cópia terminou: Reboot Now. A tela pede \"Please remove the installation medium, then press ENTER\": tira o pendrive, Enter. Se não tirar, a máquina pode começar a instalação de novo.", italic=True, color=COR_ALERTA)

    add_heading(doc, "INTERVALO — 15 MIN", level=2)
    add_para(doc, "Se alguma máquina ainda estiver copiando, ela segue sozinha durante o intervalo.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — PRÁTICA · PRIMEIRO LOGIN E RAIO-X · 25 MIN")
    add_para(doc, "A tela mostra: servidor-equipeN login:")
    add_para(doc, "Antes de qualquer um digitar, pare a turma toda: \"Quando vocês digitarem a senha, NADA vai aparecer. Nem bolinha, nem asterisco. Não travou: ele está escondendo. Digitem com calma e Enter.\" Esse é o susto número um do primeiro contato.", bold=True, color=COR_ALERTA)
    add_bullet(doc, "Se aparecerem mensagens depois do login: (cloud-init, chaves SSH), é o sistema terminando de ligar. Enter e o login: volta")
    add_bullet(doc, "Entrou: aparece equipeN@servidor-equipeN:~$ . Explique as partes: quem@qual-máquina:onde$. O $ quer dizer \"pode mandar\"")
    add_heading(doc, "O raio-x (página 2 da folha)", level=2)
    add_table(doc, [["Comando", "O que mostra"]] + [list(r) for r in RAIOX], col_widths=[5.5, 11])
    add_bullet(doc, "Em ip a, ensine a ignorar o lo (127.0.0.1, a máquina falando com ela mesma). O endereço é o inet da outra placa (enp…, eno…). Tem que bater com o da tela de rede do instalador")
    add_bullet(doc, "No primeiro ping, troque o N pelo número da equipe. -c 4 manda 4 e para. Sem o -c, o ping não para: Ctrl+C")
    add_bullet(doc, "Ping em google.com: \"Pra pingar um nome, o servidor perguntou pra quem?\" (pra agenda, o DNS da Noite 7)")
    add_bullet(doc, "free -h e df -h: a equipe anota e faz a conta de quanto sobra (Matemática). \"Esse computador que ninguém queria tem espaço pra quantas fotos de 3 MB?\"")
    add_bullet(doc, "Dois atalhos que mudam a vida: seta para cima repete o último comando; Tab completa a palavra")
    add_para(doc, "Comando digitado errado mostra \"command not found\". Não é desastre: leia a mensagem com a equipe. É o método de diagnóstico do módulo.", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — PRÁTICA · O PRIMEIRO SITE · 30 MIN")
    add_para(doc, "Página 3 da folha.")
    for c, _ in SITE[:3]:
        add_cmd(doc, c)
    add_bullet(doc, "sudo = \"faça como administrador\". Pede a senha da equipe (invisível de novo). Ele lembra por 15 minutos")
    add_bullet(doc, "apt é a \"loja de aplicativos\" do servidor, só que por texto. update atualiza a vitrine; install instala")
    add_bullet(doc, "Na pergunta Do you want to continue? [Y/n], a letra maiúscula é o padrão: só Enter já é sim")
    add_bullet(doc, "systemctl status nginx: procurar active (running) em verde. Para sair, a tecla q")
    add_heading(doc, "O momento da noite", level=2)
    add_para(doc, "Cada um conecta o celular no Wi-Fi da equipe (desliga os dados móveis) e digita no navegador http:// e o endereço do servidor. Aparece \"Welcome to nginx!\".")
    add_para(doc, "Deixe a turma comemorar. Depois: \"Essa página não veio da internet. Veio daquela máquina velha ali, pelo cabo, pelo roteador, pelo Wi-Fi até o seu celular. Vocês são o restaurante agora.\"", italic=True, color=COR_PRIMARIA)
    add_heading(doc, "Plano B: sem internet não instala o Nginx", level=2)
    add_para(doc, "O Ubuntu já vem com Python, que tem um servidor de páginas embutido. A equipe faz a página do Bloco 8 numa pasta e liga esse servidor:")
    add_cmd(doc, "mkdir site")
    add_cmd(doc, "nano site/index.html")
    add_cmd(doc, "sudo python3 -m http.server 80 --directory site")
    add_para(doc, "Abre no celular do mesmo jeito. Fica ligado enquanto a tela estiver aberta; Ctrl+C desliga. Quando a internet voltar, instale o Nginx no Encontro 2.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — PRÁTICA · A PÁGINA DA EQUIPE · 25 MIN")
    add_cmd(doc, SITE[3][0])
    add_para(doc, "Abre o nano, um editor de texto dentro do terminal. A equipe digita as quatro linhas da folha, trocando o N e os nomes. Hoje sem acento e sem ç (no terminal eles dão trabalho):")
    for l in PAGINA:
        add_cmd(doc, l, size=10)
    add_bullet(doc, "Salvar: Ctrl+O e Enter (confirma o nome). Sair: Ctrl+X. Os atalhos estão no rodapé do nano: ^ quer dizer Ctrl")
    add_bullet(doc, "Recarregar no celular: a página da equipe substitui o Welcome")
    add_bullet(doc, "Explique só isto do HTML: <h1> é título grande, <p> é parágrafo. Abre com <h1> e fecha com </h1>")
    add_bullet(doc, "Troque o digitador e peça uma linha nova: o que a equipe quiser (o time, a música, um recado). Cada um da equipe põe a mão no nano")
    add_para(doc, "Onde dilatar: o leitor muda uma letra no celular? Não: só no servidor. Peça que a equipe erre de propósito (apague um </h1>) e veja o que acontece na página. Depois conserta.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 9
    add_heading(doc, "BLOCO 9 — PRÁTICA · VISITA AOS OUTROS SERVIDORES · 15 MIN")
    add_para(doc, "Escreva no quadro o endereço do servidor de cada equipe. Primeiro, cada equipe tenta abrir o site da equipe vizinha continuando no PRÓPRIO Wi-Fi. Não abre.")
    add_bullet(doc, "\"Por quê?\" Cada roteador é uma casa. O servidor da equipe 2 está dentro da casa da equipe 2")
    add_bullet(doc, "Para visitar, conectem no Wi-Fi da outra equipe. A senha está na ficha de entrega da Noite 8 (entregue as fichas). Agora abre")
    add_bullet(doc, "Cada equipe visita pelo menos duas e anota na folha o que estava escrito na página")
    add_para(doc, "Feche com a Noite 1: \"Esse site está na REDE, não na INTERNET. Quem está em casa não abre. Para o escritório da questão, é exatamente o que se quer: uma página interna, só para quem está lá dentro.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 10
    add_heading(doc, "BLOCO 10 — DESLIGAR CERTO + FECHAMENTO · 10 MIN")
    add_cmd(doc, "sudo poweroff")
    add_bullet(doc, "\"Servidor não se desliga no botão nem na tomada. Ele pode estar escrevendo no disco. Puxar da tomada é como arrancar a panela do fogo no meio da receita\"")
    add_bullet(doc, "As três frases (slide final). Confira se a ficha técnica da página 3 está preenchida: ela continua nos próximos encontros")
    add_bullet(doc, "Encontro 2: andar pelas pastas, criar e apagar arquivos, criar usuários e entrar no servidor pelo celular, de longe (SSH)")

    add_hr(doc)

    add_heading(doc, "SE TRAVAR")
    add_table(doc, [["O que aconteceu", "O que fazer"]] + [list(r) for r in TRAVOU], col_widths=[5.5, 11])

    add_hr(doc)

    add_heading(doc, "SE PERGUNTAREM")
    add_para(doc, "\"Dá pra abrir esse site de casa?\"", bold=True)
    add_para(doc, "> \"Dá, mas aí o servidor fica exposto pra internet inteira, e precisa de outros cuidados: endereço fixo do provedor, liberar porta no roteador, segurança. É um serviço diferente. O que a gente faz aqui é o que o escritório pediu: página interna.\"", italic=True)
    add_para(doc, "\"Por que não instalar com janelas, igual Windows?\"", bold=True)
    add_para(doc, "> \"Janela gasta memória e ninguém olha a tela de servidor. Sem janela, essa máquina velha sobra força pra trabalhar. E no Encontro 2 vocês vão mexer nela pelo celular, sem nem chegar perto.\"", italic=True)
    add_para(doc, "\"Linux pega vírus?\"", bold=True)
    add_para(doc, "> \"Pode ter problema de segurança como qualquer sistema. O que protege é atualizar (o apt faz isso), senha forte e só instalar o necessário. É o mesmo cuidado da Noite 10 com o roteador.\"", italic=True)
    add_para(doc, "\"Dá pra instalar Linux no meu computador de casa?\"", bold=True)
    add_para(doc, "> \"Dá. Para usar no dia a dia, o Ubuntu Desktop (com janelas) ou o Linux Mint. Mas instalar apaga o Windows se escolher disco inteiro, como vocês viram hoje. Primeiro teste pelo pendrive, sem instalar, e faça cópia dos arquivos.\"", italic=True)
    add_para(doc, "\"Isso é profissão?\"", bold=True)
    add_para(doc, "> \"É. Pequeno escritório, clínica, escola, loja: alguém instala e cuida do servidor de arquivos e da página interna. Quem sabe Linux e documenta o serviço tem trabalho, aqui e à distância.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "NÃO reinstale nem formate: o Encontro 2 continua de onde parou, na mesma máquina, com o mesmo usuário")
    add_bullet(doc, "Guarde cada servidor junto com o roteador da equipe e a ficha de entrega")
    add_bullet(doc, "Tire foto da ficha técnica de cada equipe (página 3): se um aluno faltar ou perder a folha, a senha do servidor está salva")
    add_bullet(doc, "Anote quantas equipes chegaram ao site no celular. É o indicador do encontro")

    add_hr(doc)

    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Iniciei o estudo de servidores com Linux a partir da diferença entre cliente e servidor, usando o roteador configurado na unidade anterior como exemplo de equipamento que presta serviços à rede. Apresentei o Linux como software livre e a distribuição Ubuntu Server como opção para reaproveitar computadores antigos. Em equipes, com rodízio de funções, os estudantes instalaram o Ubuntu Server 24.04 em computadores reaproveitados, por meio de pendrive, lendo e interpretando as telas do instalador em inglês técnico e registrando cada escolha. Acessaram o sistema pelo terminal e levantaram nome, endereço de rede, memória e espaço em disco da máquina, testando a comunicação com o roteador e com a internet. Instalaram o servidor web Nginx pelo gerenciador de pacotes, publicaram uma página interna da equipe e verificaram o acesso pelo celular, inclusive nas redes das outras equipes, discutindo a diferença entre rede local e internet. Iniciaram a ficha técnica do servidor e desligaram o sistema de forma segura.")

    add_hr(doc)

    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou pra usar em outro encontro:", italic=True, color=COR_CINZA)
    add_para(doc, "1. O servidor trabalhando ao vivo", bold=True)
    add_para(doc, "sudo apt install htop e depois htop: a tela mostra processador e memória em tempo real. Enquanto a equipe recarrega a página no celular várias vezes, o digitador olha o htop. Sai com q. Discuta: por que uma máquina antiga dá conta? (servir uma página é pouco trabalho). 15 min.")
    add_para(doc, "2. Quanto custa deixar ligado", bold=True)
    add_para(doc, "Ciências da Natureza e Matemática. Cada equipe lê a potência na etiqueta da fonte da máquina (ou usa 60 W, valor comum de máquina antiga sem monitor), calcula o consumo de 24 h por 30 dias em kWh e multiplica pela tarifa da fatura de luz de casa. Comparar: e se ficar ligada só no horário do escritório? 20 min.")
    add_para(doc, "3. O cartaz do restaurante", bold=True)
    add_para(doc, "Cada equipe faz um cartaz com o caminho da página: servidor → cabo → roteador → Wi-Fi → celular, marcando em cada ponto o endereço de verdade da equipe (IP do servidor, do roteador, do celular). Apresentação de 2 minutos por equipe. 25 min.")
    add_para(doc, "4. Decifrar as mensagens", bold=True)
    add_para(doc, "Liste no quadro mensagens que apareceram na noite (command not found, Permission denied, Could not get lock, active (running), Temporary failure in name resolution). As equipes traduzem e escrevem o que o técnico faz diante de cada uma. Linguagens: inglês técnico. 15 min.")
    add_para(doc, "5. A página com mais coisas", bold=True)
    add_para(doc, "No nano, a equipe acrescenta uma lista (<ul><li>Arquivos</li><li>Recados</li></ul>) com o que a página interna do escritório deveria ter, e um link para o site de outra equipe: <a href=\"http://192.168.1N.1xx\">Equipe N</a>. 20 min.")

    path = os.path.join(OUTDIR, f"Roteiro - {TITULO_ARQ}.docx")
    doc.save(path)
    return path


# ============================================================
# SLIDES
# ============================================================
def build_slides():
    prs = new_presentation()
    P, D, K, R = PPalette.P, PPalette.D, PPalette.K, PPalette.R

    def mono(tb):
        for para in tb.text_frame.paragraphs:
            for run in para.runs:
                if run.text.startswith("$ ") or run.text.startswith("<"):
                    run.font.name = MONO

    def lista(kicker, titulo, linhas, subtitulo=None, size=20, nota=None, nota_cor=None):
        s = add_slide(prs)
        add_title_block(s, kicker, titulo, subtitulo)
        tb = add_text(s, linhas, 0.7, 2.8, 11.9, 3.8, size=size, color=K)
        mono(tb)
        if nota:
            add_text(s, nota, 0.7, 6.4, 11.9, 0.5, size=16, italic=True, color=nota_cor or D, align=PP_ALIGN.CENTER)
        add_footer(s, FOOTER)
        return s

    def frase(kicker, texto, cor=None, size=36):
        s = add_slide(prs)
        add_text(s, kicker, 0.7, 1.6, 11.9, 0.5, size=16, bold=True, color=D, align=PP_ALIGN.CENTER)
        tb = add_text(s, texto.split("\n"), 0.9, 2.4, 11.5, 3.0, size=size, bold=True, color=cor or P, align=PP_ALIGN.CENTER)
        mono(tb)
        add_footer(s, FOOTER)
        return s

    def tabela(kicker, titulo, rows, col_widths=None, subtitulo=None, size=16, nota=None, nota_cor=None, height=3.6, top=2.6):
        s = add_slide(prs)
        add_title_block(s, kicker, titulo, subtitulo)
        add_table_slide(s, rows, top=top, height=height, col_widths=col_widths, size=size)
        if nota:
            add_text(s, nota, 0.7, 6.55, 11.9, 0.45, size=15, italic=True, color=nota_cor or D, align=PP_ALIGN.CENTER)
        add_footer(s, FOOTER)
        return s

    add_slide_cover(prs, "SERVIDORES COM LINUX · ENCONTRO 1", "INSTALANDO", "O UBUNTU",
                    "e o primeiro site, aberto no celular de vocês",
                    "Curso FIC · Montador e Reparador de Computadores · CEJA Itapiranga – SC")

    lista("PRA COMEÇAR", "Vocês já configuraram um servidor", [
        "Aquele roteador da Noite 8:",
        "",
        "•  Entrega endereço pra todo aparelho que conecta (DHCP)",
        "•  Tem uma página que abre no navegador (o painel)",
        "•  Fica ligado o tempo todo, sem ninguém olhando",
        "",
        "Ele trabalha para os outros aparelhos. Isso é um servidor.",
    ], size=22)

    tabela("CLIENTE × SERVIDOR", "A cozinha do restaurante", [
        ["", "Cliente (a mesa)", "Servidor (a cozinha)"],
        ["O que faz", "Pede e mostra", "Prepara e entrega"],
        ["Quantos atende", "Só a si mesmo", "Todas as mesas ao mesmo tempo"],
        ["Onde fica", "Na mão da pessoa", "Lá atrás, ninguém vê"],
        ["Exemplos", "Celular, notebook, TV", "YouTube, banco, WhatsApp, o roteador"],
    ], col_widths=[2.9, 4.5, 4.5], size=18, height=3.4,
        nota="A foto do Instagram não está no seu celular. Está num servidor.")

    frase("A FRASE DA NOITE", "Servidor não é um computador especial.\nÉ um computador\nque trabalha para os outros.", size=34)

    lista("A DEMANDA DO CLIENTE", "O escritório pequeno", [
        "\"Queremos guardar os arquivos num lugar só,",
        "acessar de qualquer computador",
        "e ter uma página interna para os funcionários.\"",
        "",
        "Como transformar um computador antigo nesse servidor?",
        "",
        "Hoje: o sistema instalado e a página interna funcionando.",
    ], size=22)

    lista("O SISTEMA", "Linux e Ubuntu Server", [
        "•  Linux: sistema gratuito, livre, roda bem em máquina antiga",
        "•  A maioria dos servidores da internet usa Linux. O Android também",
        "•  Ubuntu: uma \"versão empacotada\" do Linux (distribuição)",
        "•  Ubuntu Server: sem janela, sem mouse. Só texto",
        "•  24.04 = abril de 2024. Recebe correções até 2029",
    ], size=21, nota="Parece filme de hacker. É só escrever ordens.")

    lista("COMO SE INSTALA", "Pendrive, tecla de boot e teclado", [
        "•  Pendrive na máquina → liga → tecla de boot (está na fita crepe)",
        "•  O instalador fala inglês. A folha traduz cada tela",
        "•  Sem mouse:",
        "        Setas e Tab = andar     Espaço = marcar",
        "        Enter = confirmar       Esc = voltar",
        "•  Digitador, leitor e anotador. Enter só depois do \"confere\"",
    ], size=21)

    frase("ANTES DE COMEÇAR", "O instalador vai apagar o disco inteiro.\nAqui as máquinas estão vazias.\nNo cliente, o técnico sempre pergunta\ne faz cópia antes.", cor=R, size=30)

    tabela("PRÁTICA · 25 MIN", "Telas 1 a 8", [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS[:8]],
           col_widths=[3.9, 3.3, 4.7], size=13, height=4.0, top=2.3,
           nota="Tela de rede: anote o número do DHCPv4. É o endereço do servidor.")

    tabela("PRÁTICA · 20 MIN", "Telas 9 a 15", [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS[8:]],
           col_widths=[3.9, 3.3, 4.7], size=13, height=3.8, top=2.3,
           nota="SSH: marca com ESPAÇO. Tem que aparecer [X].", nota_cor=R)

    tabela("O COMBINADO DO PERFIL", "Tudo minúsculo, sem espaço, sem acento", [
        ["Campo", "O que a equipe põe"],
        ["Your name", "Equipe N"],
        ["Your server's name", "servidor-equipeN"],
        ["Pick a username", "equipeN"],
        ["Choose a password", "A equipe inventa (10+ caracteres) e ANOTA"],
    ], col_widths=[4.5, 7.4], size=20, height=3.0)

    lista("ENQUANTO COPIA · 20 MIN", "Cliente ou servidor?", [
        "Página 2 da folha:",
        "",
        "•  Doze casos para marcar: cliente ou servidor?",
        "•  Leiam os comandos do raio-x e escrevam o que acham que cada um mostra",
        "•  O anotador vigia a tela: chama quando aparecer Reboot Now",
    ], size=21, nota="Reboot Now → TIRE O PENDRIVE → Enter", nota_cor=R)

    frase("", "INTERVALO\n15 minutos", cor=K)

    frase("O PRIMEIRO LOGIN", "Quando digitar a senha,\nNADA aparece.\nNão travou. Digite com calma e Enter.", cor=R, size=32)

    lista("ENTROU", "O terminal está esperando", [
        "$ equipe3@servidor-equipe3:~$",
        "",
        "•  equipe3 = quem está usando",
        "•  servidor-equipe3 = em qual máquina",
        "•  ~ = onde (a pasta da equipe)",
        "•  $ = \"pode mandar\"",
    ], size=22, nota="Mensagens depois do login:? É o sistema terminando de ligar. Enter.")

    tabela("PRÁTICA · 25 MIN", "O raio-x da máquina", [["Comando", "O que mostra"]] + [list(r) for r in RAIOX],
           col_widths=[4.2, 7.7], size=15, height=4.0, top=2.3,
           nota="Seta pra cima repete. Tab completa. Errou? Leia a mensagem.")

    lista("PRÁTICA · 30 MIN", "O primeiro site", [
        "$ sudo apt update",
        "$ sudo apt install nginx",
        "$ systemctl status nginx",
        "",
        "•  sudo = como administrador (pede a senha, invisível)",
        "•  apt = a loja de aplicativos do servidor",
        "•  [Y/n]: Enter já é sim · status: active (running) · sai com q",
    ], size=21)

    frase("NO CELULAR", "Wi-Fi da equipe, dados móveis desligados\nhttp://192.168.1N.1xx\n\nWelcome to nginx!", size=30)

    frase("VOCÊS SÃO O RESTAURANTE", "Essa página não veio da internet.\nVeio daquela máquina velha,\npelo cabo, pelo roteador, pelo Wi-Fi.", size=30)

    lista("PRÁTICA · 25 MIN", "A página da equipe", [
        "$ sudo nano /var/www/html/index.html",
        "",
    ] + PAGINA + [
        "",
        "Salvar: Ctrl+O, Enter · Sair: Ctrl+X · Recarregue o celular",
    ], size=19, nota="Hoje sem acento e sem ç. Cada um da equipe põe uma linha.")

    lista("PRÁTICA · 15 MIN", "Visita aos outros servidores", [
        "•  Tente abrir o site da equipe vizinha no SEU Wi-Fi. Abre?",
        "•  Cada roteador é uma casa",
        "•  Conecte no Wi-Fi da outra equipe (ficha de entrega da Noite 8)",
        "•  Agora abre? Anote o que estava escrito",
    ], size=22)

    frase("LEMBRA DA NOITE 1?", "Esse site está na REDE,\nnão na INTERNET.\nPágina interna: só quem está dentro vê.", size=32)

    frase("DESLIGAR CERTO", "$ sudo poweroff\n\nNunca no botão, nunca na tomada:\nele pode estar escrevendo no disco.", size=30)

    lista("O QUE FICA", "Três frases pra guardar", [
        "\"Servidor é um computador que trabalha para os outros.\"",
        "",
        "\"Senha não aparece. Não travou.\"",
        "",
        "\"Errou o comando? Leia a mensagem.\"",
    ], size=23)

    lista("NO ENCONTRO 2", "Morar no terminal", [
        "•  Andar pelas pastas, criar, copiar e apagar arquivos",
        "•  Criar usuários para os funcionários do escritório",
        "•  Entrar no servidor pelo celular, de longe (o SSH que vocês marcaram hoje)",
        "",
        "Tarefa: quatro serviços que você usou esta semana e que dependem de um servidor.",
    ], size=21)

    path = os.path.join(OUTDIR, f"Aula - {TITULO_ARQ}.pptx")
    prs.save(path)
    return path


# ============================================================
# FOLHA DO ALUNO
# ============================================================
def build_folha():
    doc = Document()
    setup_page(doc, margem=Cm(1.6))

    # PÁGINA 1 — A instalação
    cabecalho_documento(
        doc,
        "INSTALANDO O UBUNTU SERVER",
        f"Servidores com Linux · Encontro 1 · Curso FIC · CEJA Itapiranga · 1 de 3",
    )
    campo(doc, "Nome: _____________________________________________     Equipe: ______   Máquina: ____________________", space_after=6, size=10)
    campo(doc, "Rodízio · Digitador: ____________   Leitor: ____________   Anotador: ____________   (troca a cada parte)", size=10)
    add_para(doc, "Sem mouse: setas e Tab andam · Espaço marca · Enter confirma · Esc volta. Enter só depois do leitor dizer \"confere\".", italic=True, size=9, color=COR_CINZA)

    tbl = add_table(doc, [["Tela (em inglês)", "O que é", "O que escolher", "✓"]] +
                    [[t[0], t[1], t[2], "☐"] for t in TELAS], col_widths=[5, 3.8, 7.4, 1])
    small_table(tbl, size=9)

    add_heading(doc, "Anote durante a instalação", level=2)
    campo(doc, "Tecla de boot desta máquina: ______   Tela de rede · número ao lado de DHCPv4: ____ . ____ . ____ . ____", size=10)
    campo(doc, "O mirror passou no teste (\"passed tests\")?  ☐ Sim  ☐ Não      Tamanho do disco que apareceu: ________ GB", size=10)
    campo(doc, "Nome do servidor: ______________________   Usuário: ______________   Senha: ______________________", size=10)
    campo(doc, "Hora em que começou a cópia: ____ : ____     Hora do Reboot Now: ____ : ____     Levou ______ minutos", size=10)
    campo(doc, "Que palavra em inglês na tela do disco avisa que tudo vai ser apagado? O que ela quer dizer?", bold=True, size=10)
    linha(doc)

    # PÁGINA 2 — Enquanto copia + raio-x
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "CLIENTE OU SERVIDOR? E O RAIO-X",
        f"Instalando o Ubuntu · Encontro 1 · 2 de 3",
    )

    add_heading(doc, "1  Enquanto copia: cliente ou servidor?", level=2)
    add_para(doc, "Cliente pede e mostra. Servidor guarda, prepara e entrega para os outros.", italic=True, size=9, color=COR_CINZA)
    rows = [["Aparelho ou serviço", "Cliente", "Servidor"]] + [[c, "☐", "☐"] for c in CLIENTE_SERVIDOR]
    tbl = add_table(doc, rows, col_widths=[13.4, 2, 2])
    small_table(tbl, size=9)

    add_heading(doc, "2  O raio-x da máquina", level=2)
    add_para(doc, "Antes de rodar, escreva o que acha que vai aparecer. Depois rode e anote o resultado. Troque N pelo número da equipe.", italic=True, size=9, color=COR_CINZA)
    tbl = add_table(doc, [["Comando", "O que mostra", "O que deu"]] +
                    [[c, d, ""] for c, d in RAIOX[:7]], col_widths=[4.2, 6.6, 6.6])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.8)
    campo(doc, "O endereço do ip a bate com o da tela de rede do instalador?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "Disco: tem ______ GB e estão livres ______ GB. Cabem quantas fotos de 3 MB? (1 GB ≈ 1000 MB) ________", size=10)
    campo(doc, "Para pingar google.com, o servidor precisou perguntar pra quem? (Noite 7) ______________________", size=10)

    # PÁGINA 3 — O site, a visita e a ficha técnica
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "O PRIMEIRO SITE",
        f"Instalando o Ubuntu · Encontro 1 · 3 de 3",
    )

    add_heading(doc, "1  Instalar e abrir no celular", level=2)
    tbl = add_table(doc, [["Comando", "Para quê", "✓"]] + [[c, d, "☐"] for c, d in SITE[:3]],
                    col_widths=[6, 10.4, 1])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=9)
    campo(doc, "No celular (Wi-Fi da equipe): http:// ____ . ____ . ____ . ____     Apareceu Welcome to nginx?  ☐ Sim  ☐ Não", size=10)

    add_heading(doc, "2  A página da equipe", level=2)
    add_cmd(doc, SITE[3][0], size=9.5)
    add_para(doc, "Digite (sem acento e sem ç), trocando N e os nomes. Salvar: Ctrl+O e Enter · Sair: Ctrl+X.", italic=True, size=9, color=COR_CINZA)
    for l in PAGINA:
        add_cmd(doc, l, size=9.5)
    campo(doc, "A linha que a equipe inventou: _________________________________________________", size=10)

    add_heading(doc, "3  Visita aos outros servidores", level=2)
    tbl = add_table(doc, [["Equipe", "Endereço do servidor", "No meu Wi-Fi abriu?", "No Wi-Fi dela abriu?", "O que estava escrito"],
                          ["", "", "", "", ""], ["", "", "", "", ""]], col_widths=[1.6, 3.8, 2.8, 2.8, 6.4])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)
    campo(doc, "Por que alguém em casa não consegue abrir o site da equipe?", bold=True, size=10)
    linha(doc)

    add_heading(doc, "4  Ficha técnica do servidor (começa hoje, continua nos próximos encontros)", level=2)
    rows = [[FICHA[i], "", FICHA[i + 5], ""] for i in range(5)]
    tbl = add_table(doc, [["Item", "Como ficou", "Item", "Como ficou"]] + rows, col_widths=[3.6, 5.1, 3.6, 5.1])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)

    add_heading(doc, "TAREFA DA SEMANA", level=2, color=COR_DESTAQUE)
    campo(doc, "Quatro serviços que você usou esta semana e que dependem de um servidor: ______________ · ______________ · ______________ · ______________", size=10)
    add_para(doc, "Traga esta folha no Encontro 2 · Desligar sempre com sudo poweroff · CEJA Itapiranga – SC", italic=True, size=9, color=COR_CINZA, align=WD_ALIGN_PARAGRAPH.CENTER)

    path = os.path.join(OUTDIR, f"Folha do Aluno - {TITULO_ARQ}.docx")
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
        "INSTALANDO O UBUNTU E O PRIMEIRO SITE",
        f"Resumo do aluno · Servidores com Linux · Encontro 1 · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Servidor não é um computador especial. É um computador que fica ligado trabalhando para os outros: guarda, prepara e entrega. Cliente é quem pede e mostra (celular, notebook, TV). Hoje um computador antigo virou servidor, com Linux, e entregou uma página para os celulares da rede.")

    add_heading(doc, "Linux e Ubuntu Server", level=2)
    add_bullet(doc, "Linux é um sistema gratuito e livre, que roda bem em máquina antiga. A maioria dos servidores do mundo usa Linux")
    add_bullet(doc, "Ubuntu é uma distribuição (uma versão empacotada) do Linux. O Ubuntu Server não tem janela nem mouse: só texto")
    add_bullet(doc, "24.04 quer dizer abril de 2024. Recebe correções até 2029")

    add_heading(doc, "A instalação em 6 passos", level=2)
    add_bullet(doc, "1. Pendrive com o instalador na máquina, liga e aperta a tecla de boot (F12, F11, F8 ou Esc)")
    add_bullet(doc, "2. Idioma English, teclado Portuguese (Brazil), tipo Ubuntu Server")
    add_bullet(doc, "3. Rede: anotar o endereço que aparece (DHCPv4). Quem deu esse número foi o roteador")
    add_bullet(doc, "4. Disco inteiro. ATENÇÃO: apaga tudo. No cliente, sempre perguntar e fazer cópia antes")
    add_bullet(doc, "5. Perfil: nome do servidor, usuário e senha (tudo minúsculo, sem espaço). Marcar Install OpenSSH server")
    add_bullet(doc, "6. Esperar a cópia, Reboot Now, tirar o pendrive e Enter")
    add_para(doc, "Sem mouse: setas e Tab andam, Espaço marca, Enter confirma, Esc volta.", italic=True, size=10, color=COR_CINZA)

    add_heading(doc, "Os comandos da noite", level=2)
    tbl = add_table(doc, [["Comando", "O que faz"]] + [list(r) for r in RAIOX] + [list(r) for r in SITE],
                    col_widths=[6, 9.5])
    mono_col(tbl, 0, size=9.5)

    add_heading(doc, "Cuidados do primeiro contato", level=2)
    add_bullet(doc, "A senha NÃO aparece quando você digita. Não travou: digite com calma e Enter")
    add_bullet(doc, "Maiúscula e minúscula são diferentes no Linux")
    add_bullet(doc, "Errou? \"command not found\" quer dizer \"não conheço esse comando\". Leia a mensagem: ela diz o que houve")
    add_bullet(doc, "Seta para cima repete o último comando. Tab completa a palavra")
    add_bullet(doc, "No celular, o endereço começa com http:// (não https) e o celular tem que estar no Wi-Fi da mesma rede do servidor")
    add_bullet(doc, "Servidor se desliga com sudo poweroff. Nunca no botão nem na tomada")

    add_heading(doc, "Rede não é internet (de novo)", level=2)
    add_para(doc, "O site da equipe só abre para quem está conectado no mesmo roteador. Quem está em casa não vê. Para um escritório, é isso que se quer: uma página interna, só para quem está lá dentro.")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Servidor é um computador que trabalha para os outros.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Senha não aparece. Não travou.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Errou o comando? Leia a mensagem.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Na sua casa, quais aparelhos são clientes? Tem algum que trabalha como servidor?")
    add_bullet(doc, "Por que o técnico pergunta e faz cópia antes de instalar um sistema na máquina de alguém?")
    add_bullet(doc, "Se o celular não abre a página do servidor, quais são as três primeiras coisas que você confere?")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Este material é seu. Guarde e volte quando precisar. · CEJA Itapiranga – SC")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COR_CINZA

    path = os.path.join(OUTDIR, f"Resumo do Aluno - {TITULO_ARQ}.docx")
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
