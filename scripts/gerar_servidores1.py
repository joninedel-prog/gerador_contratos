"""Gera os 4 arquivos do Encontro 1 do módulo de Servidores (Plano Mensal 4).

Servidores 1 — Instalando o Ubuntu Server e o primeiro site
Primeiro contato da turma com Linux. Os ALUNOS instalam o Ubuntu Server 24.04
nas máquinas antigas, em equipes (uma por máquina), ligadas aos roteadores
configurados na Noite 8. Depois entram no sistema pela primeira vez, tiram o
"raio-x" da máquina com comandos simples, instalam o Nginx e fazem uma página
da equipe que abre no celular. No fim, visitam o site das outras equipes
conectando no Wi-Fi delas com a ficha de entrega da Noite 8.
Padrão prática-pesada, sem role-play e sem duelo. Precisa de internet.

Versão 2026-11-18 (aula dada nesse dia): abertura completa antes da prática —
o módulo, o que é servidor (tipos, data center), sistema operacional (síndico),
história do Linux e software livre, distribuições, como o computador liga —
intercalada com práticas curtas (cliente ou servidor, kernel Linux no Android).
Na espera da cópia, o dicionário do técnico (inglês das telas). Folha com 4 páginas.
As funções add_cmd, mono_col, small_table, campo e linha são usadas pelos
scripts dos Encontros 2 e 3: não mudar a assinatura delas.
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
    "<p>Instalado em 18/11/2026 no CEJA Itapiranga</p>",
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
# Conteúdo das introduções (versão 2026-11-18: aula com abertura completa)
# ============================================================
DATA_AULA = "quarta 18/11"

# O que o módulo vai construir
MODULO = [
    ("Encontro 1 (hoje)", "Instalar o Ubuntu Server, conhecer a máquina pelo terminal e publicar a primeira página"),
    ("Encontro 2", "Mandar no servidor de longe pelo PC, organizar pastas, criar os funcionários e o site do negócio"),
    ("Encontro 3", "A pasta do escritório aparece no Windows e no celular. Entrega e apresentação do servidor"),
]

# Tipos de servidor do dia a dia
TIPOS_SERVIDOR = [
    ("Servidor de páginas (web)", "Entrega sites: o do banco, o da prefeitura, a página que vocês fazem hoje"),
    ("Servidor de arquivos", "Guarda os arquivos de um escritório num lugar só (Encontro 3)"),
    ("Servidor de e-mail e de mensagens", "Segura a mensagem até a pessoa ligar o celular"),
    ("Servidor de vídeo", "YouTube, Netflix: guarda e entrega os vídeos"),
    ("Servidor de jogos", "Partida online, servidor de Minecraft: todos os jogadores se encontram nele"),
    ("DVR / NVR de câmeras", "Grava as câmeras e entrega as imagens para o aplicativo"),
    ("Servidor de endereços e de nomes", "O roteador (DHCP) e a agenda da internet (DNS)"),
]

# O sistema operacional
SO_FUNCOES = [
    "Liga a máquina e conversa com as peças: processador, memória, disco, placa de rede, teclado",
    "Divide a máquina entre os programas: quem usa a memória, quem usa o processador, em que ordem",
    "Guarda os arquivos organizados em pastas no disco",
    "Controla quem pode entrar e o que cada um pode fazer (usuários e senhas)",
]

SO_EXEMPLOS = [
    ("Windows", "Microsoft", "A maioria dos computadores de casa e de escritório. Pago"),
    ("macOS / iOS", "Apple", "Computadores Mac e iPhone. Só nas máquinas da Apple"),
    ("Android", "Google", "A maioria dos celulares. Por dentro, é Linux"),
    ("Linux", "Comunidade (milhares de pessoas e empresas)", "Servidores, roteadores, TVs, supercomputadores. Gratuito"),
]

# A história do Linux em 5 datas
LINUX_HISTORIA = [
    ("1983", "Richard Stallman começa o projeto GNU: um sistema inteiro que qualquer um pudesse usar, estudar, mudar e repartir"),
    ("1991", "Linus Torvalds, estudante de 21 anos na Finlândia, cria o núcleo (kernel) Linux no computador de casa e divulga de graça na internet"),
    ("Anos 90", "Voluntários do mundo todo melhoram o Linux juntos. GNU + Linux viram um sistema completo"),
    ("2004", "Nasce o Ubuntu, uma versão mais fácil de instalar. \"Ubuntu\" é uma palavra africana: \"sou o que sou porque somos todos\""),
    ("Hoje", "Linux está nos servidores da internet, no Android, em roteadores, TVs e carros, e em todos os 500 supercomputadores mais rápidos do mundo"),
]

DISTROS = [
    ("Ubuntu", "A mais conhecida. Tem a versão com janelas (Desktop) e a Server, que é a de hoje"),
    ("Debian", "A \"mãe\" do Ubuntu. Muito estável, muito usada em servidor"),
    ("Linux Mint", "Parece o Windows. Boa para computador antigo de casa"),
    ("Fedora", "Mantida com apoio da Red Hat, empresa que vende Linux para grandes empresas"),
]

# O que acontece quando o computador liga
BOOT = [
    ("1. Liga o botão", "A placa-mãe acorda"),
    ("2. BIOS / UEFI", "O \"porteiro\" da placa-mãe: testa as peças e procura de onde ligar"),
    ("3. Disco ou pendrive?", "A tecla de boot (F12, F11, F8, Esc) deixa você escolher. Hoje: o pendrive"),
    ("4. O sistema carrega", "Do pendrive vem o instalador. Depois de instalado, vem do disco"),
]

# Atividade da espera: o dicionário do técnico
DICIONARIO = [
    ("Install", "Instalar"), ("Language", "Idioma"), ("Keyboard", "Teclado"),
    ("Network", "Rede"), ("Storage / Disk", "Armazenamento / Disco"), ("Mirror", "Espelho: o site de onde baixar"),
    ("Profile", "Perfil"), ("Username", "Nome de usuário"), ("Password", "Senha"),
    ("Done", "Pronto"), ("Continue", "Continuar"), ("Back", "Voltar"),
    ("Skip", "Pular"), ("Confirm", "Confirmar"), ("Destructive action", "Ação que destrói (apaga)"),
    ("Reboot", "Reiniciar"), ("Remove the installation medium", "Tire o pendrive"), ("Update", "Atualização"),
]


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "INSTALANDO O UBUNTU E O PRIMEIRO SITE",
        f"Roteiro do professor · Servidores com Linux · Encontro 1 de 3 · {DATA_AULA} · Curso FIC · CEJA Itapiranga – SC",
    )

    add_heading(doc, "A IDEIA DESTE ENCONTRO")
    add_para(doc, "É o primeiro contato da turma com Linux e com servidor. Quase ninguém viu uma tela só de texto, e muitos não sabem o que é um sistema operacional, mesmo usando um todo dia. Por isso a noite abre com uma introdução de verdade: o que vamos construir no módulo, o que é um servidor, o que é um sistema operacional, de onde vem o Linux e o que acontece quando o computador liga.")
    add_para(doc, "A introdução não é um bloco só de fala. Ela vem em três partes curtas, cada uma seguida de uma atividade em que o aluno faz alguma coisa (classificar cliente e servidor, achar o Linux no próprio celular). Depois, a parte prática: cada equipe instala o Ubuntu Server com as próprias mãos e termina a noite abrindo no celular um site que está rodando naquela máquina.")
    add_para(doc, "A frase da noite: \"Servidor não é um computador especial. É um computador que trabalha para os outros.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "O que se resolve neste encontro", level=2)
    add_para(doc, "Cada aluno sai sabendo diferenciar cliente de servidor com exemplos do dia a dia, dizer o que um sistema operacional faz, contar de onde vem o Linux e por que ele é gratuito, explicar o que acontece quando o computador liga, dar boot por pendrive, passar pelas telas do instalador do Ubuntu Server entendendo cada escolha, entrar no sistema, descobrir o endereço, a memória e o disco da máquina, instalar um programa com o apt e publicar uma página que abre nos celulares da rede.")

    add_heading(doc, "Onde este encontro fica no módulo", level=2)
    add_para(doc, "Abre o Plano Mensal 4 (servidores com Linux). Usa os roteadores configurados na Noite 8 e a ficha de entrega deles. Retoma a Noite 1 (rede não é internet: o site só abre dentro da rede), a Noite 6 (endereço, gateway, ping, e a analogia do condomínio) e a Noite 7 (o ping em google.com só funciona porque a agenda, o DNS, responde). O SSH fica instalado hoje para o Encontro 2, em que a turma passa a mandar no servidor pelo PC com o PuTTY.")
    add_para(doc, "Não entre em: partições e LVM, interface gráfica, comparação técnica entre distribuições, licenças (GPL) em detalhe, domínio na internet, redirecionamento de portas, HTML além das quatro linhas da página. Se alguém perguntar \"e pra abrir de casa?\", a resposta está em Se perguntarem.", italic=True, color=COR_ALERTA)

    add_hr(doc)

    add_heading(doc, "ANTES DA AULA")
    add_heading(doc, "As máquinas", level=2)
    add_bullet(doc, "Confira cada máquina: processador de 64 bits, pelo menos 2 GB de RAM e disco de pelo menos 25 GB")
    add_bullet(doc, "Confirme com a escola que NÃO há nada nos discos. A instalação apaga tudo")
    add_bullet(doc, "Descubra a tecla do menu de boot de cada modelo (F12, F11, F8, Esc) e cole numa fita crepe em cima da máquina: \"BOOT = F12\"")
    add_bullet(doc, "Se der tempo, faça UMA instalação completa antes, seguindo a folha. Essa máquina fica como RESERVA para a equipe cuja máquina falhar. Sem tempo: tenha ao menos um pendrive testado dando boot")
    add_bullet(doc, "Numere com fita crepe: SERVIDOR EQUIPE 1, SERVIDOR EQUIPE 2… Cada máquina com monitor, teclado (mouse não é usado), tomada e cabo de rede")
    add_bullet(doc, "Separe UMA máquina aberta (ou uma placa-mãe velha) para mostrar no Bloco 4: processador, memória, disco. Opcional, mas ajuda muito")

    add_heading(doc, "Os pendrives", level=2)
    add_bullet(doc, "Imagem ubuntu-24.04.x-live-server-amd64.iso, de ubuntu.com/download/server")
    add_bullet(doc, "Gravada com o Rufus (Windows) ou o balenaEtcher, um pendrive de 4 GB ou mais por equipe")
    add_bullet(doc, "Menos pendrives do que equipes: o pendrive fica preso na máquina até o fim da cópia. Comece as equipes em ondas e use o dicionário do técnico (Bloco 8) para quem espera")

    add_heading(doc, "A rede", level=2)
    add_bullet(doc, "Cada servidor vai no roteador da sua equipe (Noite 8), por cabo, numa porta LAN. A WAN do roteador vai num ponto de rede da escola")
    add_bullet(doc, "Teste a internet de cada roteador com um celular. Sem internet não instala o Nginx (plano B no Bloco 10)")
    add_bullet(doc, "Tenha as fichas de entrega da Noite 8 (Wi-Fi de cada roteador): usadas no Bloco 12")
    add_bullet(doc, "Com os roteadores como a Noite 8 deixou, o servidor da equipe N ganha um endereço 192.168.1N.1xx. No Encontro 2 ele passa a ser fixo em 192.168.1N.10")

    add_heading(doc, "Imprima e projete", level=2)
    add_bullet(doc, "Folha do aluno (4 páginas), uma por pessoa. Resumo do aluno, uma por pessoa")
    add_bullet(doc, "Os slides seguem a ordem dos blocos. Os da introdução (slides 2 a 21) são o roteiro da sua fala")

    add_hr(doc)

    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Abertura do módulo", "10 min", "O que vamos construir em três encontros, como vai funcionar, a enquete da mão levantada"],
        ["2. O que é um servidor", "15 min", "O roteador já é um servidor, a cozinha do restaurante, os tipos de servidor, onde eles ficam"],
        ["3. PRÁTICA — Cliente ou servidor?", "15 min", "Doze casos em dupla, correção em voz alta"],
        ["4. Sistema operacional e Linux", "15 min", "O síndico do computador, os sistemas que eles já usam, a história do Linux, as distribuições"],
        ["5. PRÁTICA — O Linux no seu bolso", "10 min", "Achar a versão do kernel Linux no Android. Onde mais tem Linux na sala?"],
        ["6. Como o computador liga e como se instala", "10 min", "BIOS, boot, pendrive, o instalador sem mouse, o aviso do disco"],
        ["7. PRÁTICA — A instalação", "40 min", "As 15 telas, em equipe, com rodízio. Anotar o IP e o perfil"],
        ["8. PRÁTICA — Enquanto copia", "15 min", "O dicionário do técnico (inglês das telas) e o palpite do raio-x"],
        ["INTERVALO", "15 min", "Coincide com o fim da cópia nas máquinas mais lentas"],
        ["9. PRÁTICA — Primeiro login e raio-x", "20 min", "A senha invisível, o prompt, oito comandos, a ficha técnica"],
        ["10. PRÁTICA — O primeiro site", "20 min", "apt update, instalar o Nginx, abrir no celular"],
        ["11. PRÁTICA — A página da equipe", "15 min", "nano, quatro linhas, salvar, recarregar no celular"],
        ["12. Visita, desligar e fechamento", "10 min", "O site de outra equipe, sudo poweroff, três frases, o Encontro 2"],
    ], col_widths=[6.5, 2, 8])
    add_para(doc, "Introdução: 50 min de fala (Blocos 1, 2, 4 e 6), sempre intercalada com prática. Atividade do aluno: Blocos 3, 5 e 7 a 12, cerca de 145 min (69% do total). Nunca há dois blocos de fala seguidos.", italic=True, color=COR_CINZA)

    add_heading(doc, "As equipes e o rodízio", level=2)
    add_para(doc, "Forme as equipes no Bloco 1, uma por máquina, de 3 ou 4. Três papéis que trocam a cada bloco de prática (anotado na página 2 da folha):")
    add_bullet(doc, "Digitador: o único com a mão no teclado")
    add_bullet(doc, "Leitor: lê a folha e a tela em voz alta, traduz o inglês, confere antes do Enter")
    add_bullet(doc, "Anotador: preenche a folha da equipe (IP, usuário, resultados)")
    add_para(doc, "A regra: o digitador não aperta Enter antes do leitor dizer \"confere\".", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — ABERTURA DO MÓDULO · 10 MIN")
    add_heading(doc, "A enquete da mão levantada — 3 min", level=2)
    add_para(doc, "Antes de explicar qualquer coisa, pergunte e conte as mãos (anote no canto do quadro, vai ser retomado no fim do módulo):")
    add_bullet(doc, "\"Quem já ouviu a palavra servidor?\" \"Quem sabe dizer o que é?\"")
    add_bullet(doc, "\"Quem já ouviu falar de Linux?\" \"Quem já usou?\" (quase todos usaram sem saber: o Android)")
    add_bullet(doc, "\"Quem já viu uma tela preta só com letras, num filme ou no computador?\"")
    add_heading(doc, "O que vamos construir — 5 min", level=2)
    add_para(doc, "Leia a questão do plano: \"Um escritório pequeno quer guardar os arquivos num lugar só, acessar de qualquer computador e ter uma página interna para os funcionários. Como transformar um computador antigo no servidor que resolve isso?\"")
    add_table(doc, [["Quando", "O que a equipe faz"]] + [list(m) for m in MODULO], col_widths=[4, 12.5])
    add_para(doc, "\"Esses computadores aqui iam para o lixo. Em três encontros, cada equipe vai transformar um deles no servidor de um escritório, com site, pasta de rede e funcionários com senha. E no fim vocês entregam para outra equipe testar, como se fosse para um cliente.\"", italic=True, color=COR_PRIMARIA)
    add_heading(doc, "Como vai funcionar — 2 min", level=2)
    add_bullet(doc, "Equipes de 3 ou 4, uma máquina por equipe, a mesma até o fim do módulo")
    add_bullet(doc, "Rodízio: digitador, leitor, anotador. Todo mundo põe a mão no teclado")
    add_bullet(doc, "A ficha técnica do servidor começa hoje e é entregue no último encontro. É a avaliação")
    add_bullet(doc, "Errar faz parte: no Linux, toda mensagem de erro diz o que aconteceu. A gente vai aprender a ler")

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — O QUE É UM SERVIDOR · 15 MIN")
    add_heading(doc, "Comece pelo que eles já fizeram — 3 min", level=2)
    add_para(doc, "Segure um roteador da Noite 8 e pergunte:")
    add_bullet(doc, "\"Quem dá o endereço pro celular quando ele conecta?\" (o roteador, o DHCP)")
    add_bullet(doc, "\"Onde vocês entraram pra configurar ele?\" (numa página: 192.168.1N.1)")
    add_para(doc, "\"Esse aparelhinho entrega endereço pra todo mundo e ainda tem uma página que abre no navegador. Ele trabalha para os outros aparelhos da casa, sem ninguém olhando. Isso tem nome: servidor. Vocês já configuraram um servidor sem saber.\"", italic=True, color=COR_PRIMARIA)
    add_heading(doc, "A cozinha do restaurante — 4 min", level=2)
    add_para(doc, "Desenhe no quadro: mesas de um lado, cozinha do outro.")
    add_bullet(doc, "O cliente senta, PEDE e recebe. Celular, notebook e TV são clientes")
    add_bullet(doc, "A cozinha fica lá atrás, ninguém vê, e SERVE todas as mesas ao mesmo tempo. É o servidor")
    add_bullet(doc, "A cozinha não fecha. Servidor fica ligado o tempo todo, quase sempre sem tela e sem ninguém do lado")
    add_bullet(doc, "\"A foto do Instagram está no seu celular?\" (não: está num servidor, o celular só pede e mostra)")
    add_heading(doc, "Os tipos de servidor — 5 min", level=2)
    add_para(doc, "Projete a tabela e pergunte, para cada linha, \"quem aqui usou isso hoje?\":")
    add_table(doc, [["Servidor", "O que faz"]] + [list(t) for t in TIPOS_SERVIDOR], col_widths=[5.5, 11])
    add_para(doc, "Puxe o DVR de câmeras: é o servidor que o técnico da região mais instala sem chamar de servidor. Quem já viu um DVR na casa de alguém?", italic=True, color=COR_DESTAQUE)
    add_heading(doc, "Onde os servidores ficam — 3 min", level=2)
    add_bullet(doc, "Os grandes ficam em data centers: galpões enormes, cheios de servidores empilhados em armários, com ar-condicionado e energia de reserva. O YouTube, o banco e o WhatsApp moram lá")
    add_bullet(doc, "Os pequenos ficam numa sala do escritório, da escola, da clínica, da cooperativa. É esse o servidor que vocês vão montar")
    add_bullet(doc, "\"Por que um computador velho serve?\" Um escritório de 5 pessoas pede pouco: uns arquivos, uma página. E sem janelas, o sistema gasta pouca memória. O que importa é ficar ligado e não perder os dados")

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — PRÁTICA · CLIENTE OU SERVIDOR? · 15 MIN")
    add_para(doc, "Página 1 da folha, em dupla. Doze casos para marcar cliente ou servidor (8 min). Depois corrija em voz alta (7 min), pedindo que a dupla justifique.")
    add_bullet(doc, "Os que dão discussão: a impressora compartilhada (servidor de impressão), o WhatsApp (o servidor segura a mensagem quando o amigo está sem internet), a câmera (ela serve as imagens, o aplicativo é cliente), a máquina que vão instalar (hoje ainda nada; no fim da noite, servidor)")
    add_bullet(doc, "Pergunta de fechamento: \"um mesmo aparelho pode ser cliente e servidor?\" (pode: o computador do escritório é cliente quando abre o site do banco e pode ser servidor quando compartilha a impressora)")

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — SISTEMA OPERACIONAL E LINUX · 15 MIN")
    add_heading(doc, "O síndico do computador — 5 min", level=2)
    add_para(doc, "Retome o condomínio da Noite 6. Se tiver a máquina aberta, mostre as peças: processador, memória, disco.")
    add_para(doc, "\"O computador é um prédio: as peças são a estrutura, os programas são os moradores. Quem organiza tudo é o síndico: o sistema operacional. Sem ele, o Word não sabe usar a memória, o WhatsApp não sabe usar a placa de rede, e ninguém sabe onde está guardado o quê.\"", italic=True, color=COR_PRIMARIA)
    add_para(doc, "O que o sistema operacional faz:")
    for f in SO_FUNCOES:
        add_bullet(doc, f)
    add_table(doc, [["Sistema", "De quem", "Onde está"]] + [list(s) for s in SO_EXEMPLOS], col_widths=[3, 5, 8.5])
    add_heading(doc, "De onde vem o Linux — 6 min", level=2)
    add_table(doc, [["Quando", "O que aconteceu"]] + [list(h) for h in LINUX_HISTORIA], col_widths=[2.3, 14.2])
    add_bullet(doc, "Software livre: qualquer um pode usar, estudar como funciona, mudar e repartir. Não é só \"de graça\": é aberto. Por isso empresas do mundo todo melhoram o Linux juntas e todos ganham")
    add_bullet(doc, "\"Como alguém ganha dinheiro com algo gratuito?\" Com o serviço: instalar, configurar, dar suporte, manter. Exatamente o que vocês vão aprender a vender")
    add_heading(doc, "As distribuições — 4 min", level=2)
    add_para(doc, "O Linux é o motor. Uma distribuição é o carro pronto: o motor mais os programas, o instalador e as atualizações, montados por alguém.")
    add_table(doc, [["Distribuição", "Como é"]] + [list(d) for d in DISTROS], col_widths=[3.5, 13])
    add_bullet(doc, "Hoje: Ubuntu Server 24.04. 24.04 = abril de 2024. É LTS (suporte longo): recebe correções de segurança até 2029")
    add_bullet(doc, "Desktop tem janelas e mouse, para usar no dia a dia. Server não tem: só texto. \"Ninguém fica olhando a tela de um servidor. Sem janelas, a máquina velha sobra força para trabalhar\"")

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — PRÁTICA · O LINUX NO SEU BOLSO · 10 MIN")
    add_para(doc, "Página 1 da folha. Quem tem Android procura a versão do kernel (o núcleo Linux) no próprio celular:")
    add_bullet(doc, "Samsung: Configurações → Sobre o telefone → Informações do software → Versão do kernel")
    add_bullet(doc, "Motorola, Xiaomi e outros: Configurações → Sobre o telefone → (Versão do Android ou Todas as especificações) → Versão do kernel")
    add_bullet(doc, "Aparece algo como 5.10.198 ou 6.1.75. É a versão do Linux que está rodando no celular. Quem tem iPhone acompanha o colega (o iPhone não é Linux, mas é \"parente\": vem da família Unix)")
    add_bullet(doc, "No quadro, anote as versões que apareceram. Qual é a mais nova? E a mais antiga? (celular mais antigo, Linux mais antigo: por isso celular velho para de receber atualização)")
    add_para(doc, "Feche com a caça: \"Onde mais tem Linux nesta sala?\" (os roteadores da bancada, a TV smart, o projetor talvez, a máquina de cartão, o painel do carro). \"Vocês usam Linux o dia inteiro. Hoje vão instalar um.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — COMO O COMPUTADOR LIGA E COMO SE INSTALA · 10 MIN")
    add_table(doc, [["Etapa", "O que acontece"]] + [list(b) for b in BOOT], col_widths=[4.5, 12])
    add_bullet(doc, "Instalar é copiar o sistema do pendrive para o disco e deixar ele pronto para ligar sozinho. O instalador pergunta umas coisas antes: idioma, teclado, rede, disco, nome e senha")
    add_bullet(doc, "O instalador é em inglês e não usa mouse: setas e Tab para andar, Espaço para marcar, Enter para confirmar, Esc volta. A folha traduz cada tela")
    add_bullet(doc, "Mostre o slide das 15 telas: é o mapa da instalação")
    add_para(doc, "Fale sério, uma vez: \"Numa tela o instalador avisa que vai apagar o disco inteiro. Aqui as máquinas estão vazias. Na casa de um cliente, isso apaga as fotos da família. O técnico SEMPRE pergunta e faz cópia antes.\"", bold=True, color=COR_ALERTA)

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — PRÁTICA · A INSTALAÇÃO · 40 MIN")
    add_para(doc, "Página 2 da folha. As equipes vão no seu ritmo, você circula. Troque o digitador na tela 9 (o disco).")
    add_table(doc, [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS], col_widths=[5.2, 4.3, 7])
    add_heading(doc, "Pontos de parada", level=2)
    add_bullet(doc, "Teclado: se escolherem errado, / e : não saem no lugar certo e nenhum comando funciona. Peça que mostrem a tela antes do Done")
    add_bullet(doc, "Rede: pare cada equipe. \"Esse número ao lado de DHCPv4 é o endereço do servidor. Quem deu esse número?\" (o roteador da equipe, o DHCP). Anotar sem o /24")
    add_bullet(doc, "Mirror: \"This mirror location passed tests\" = o servidor alcançou a internet. Se não passar, o problema é a WAN do roteador")
    add_bullet(doc, "Disco: o instalador pergunta \"Confirm destructive action\". O leitor lê a palavra destructive em voz alta antes do Continue")
    add_table(doc, [
        ["Perfil: campo", "O que a equipe põe", "Regra"],
        ["Your name", "Equipe N", "Livre"],
        ["Your server's name", "servidor-equipeN", "Minúsculo, sem espaço, sem acento"],
        ["Pick a username", "equipeN", "Minúsculo, sem espaço, sem acento"],
        ["Choose a password", "A equipe inventa", "10 caracteres ou mais. Anotar na folha"],
    ], col_widths=[4.5, 4.5, 7.5])
    add_bullet(doc, "SSH: marcar é com Espaço, não Enter. Tem que aparecer [X] Install OpenSSH server. É o que permite mandar no servidor pelo PC no Encontro 2")
    add_bullet(doc, "Depois dos snaps começa a cópia (Installing system): de 10 a 30 minutos em máquina antiga. \"Cancel update and reboot\" pode ser usado se as atualizações demorarem muito")
    add_para(doc, "Onde dilatar: na tela de rede, a equipe confere o número no painel do roteador (lista de clientes DHCP, Noite 6). É o mesmo endereço?", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — PRÁTICA · ENQUANTO COPIA · 15 MIN")
    add_para(doc, "O anotador vigia a tela e chama quando aparecer Reboot Now. Os outros fazem a página 3 da folha:")
    add_bullet(doc, "O dicionário do técnico: 18 palavras e expressões em inglês que apareceram nas telas. A equipe traduz sem celular primeiro, depois confere com a tabela abaixo. Linguagens: inglês técnico")
    add_table(doc, [["Na tela", "Quer dizer"]] + [list(d) for d in DICIONARIO], col_widths=[6, 10.5])
    add_bullet(doc, "O palpite do raio-x: a equipe lê os comandos que vão rodar depois do intervalo e escreve o que acha que cada um mostra")
    add_para(doc, "A cópia terminou: Reboot Now. A tela pede \"Please remove the installation medium, then press ENTER\": tira o pendrive, Enter. Se não tirar, a máquina pode começar a instalação de novo.", italic=True, color=COR_ALERTA)

    add_heading(doc, "INTERVALO — 15 MIN", level=2)

    add_hr(doc)

    # BLOCO 9
    add_heading(doc, "BLOCO 9 — PRÁTICA · PRIMEIRO LOGIN E RAIO-X · 20 MIN")
    add_para(doc, "A tela mostra: servidor-equipeN login:")
    add_para(doc, "Antes de qualquer um digitar, pare a turma toda: \"Quando vocês digitarem a senha, NADA vai aparecer. Nem bolinha, nem asterisco. Não travou: ele está escondendo. Digitem com calma e Enter.\"", bold=True, color=COR_ALERTA)
    add_bullet(doc, "Mensagens depois do login: (cloud-init, chaves SSH) = o sistema terminando de ligar. Enter e o login: volta")
    add_bullet(doc, "Entrou: equipeN@servidor-equipeN:~$ . As partes: quem@qual-máquina:onde$. O $ quer dizer \"pode mandar\"")
    add_table(doc, [["Comando", "O que mostra"]] + [list(r) for r in RAIOX], col_widths=[5.5, 11])
    add_bullet(doc, "Em ip a, ignorar o lo (127.0.0.1, a máquina falando com ela mesma). O endereço é o inet da outra placa (enp…, eno…)")
    add_bullet(doc, "No primeiro ping, trocar o N pelo número da equipe. -c 4 manda 4 e para. Sem o -c, Ctrl+C")
    add_bullet(doc, "Ping em google.com: \"Pra pingar um nome, o servidor perguntou pra quem?\" (pra agenda, o DNS da Noite 7)")
    add_bullet(doc, "free -h e df -h: a equipe anota e faz a conta de quanto sobra. \"Cabem quantas fotos de 3 MB?\"")
    add_bullet(doc, "Seta para cima repete o último comando; Tab completa a palavra. Errou? \"command not found\": leia a mensagem")

    add_hr(doc)

    # BLOCO 10
    add_heading(doc, "BLOCO 10 — PRÁTICA · O PRIMEIRO SITE · 20 MIN")
    for c, _ in SITE[:3]:
        add_cmd(doc, c)
    add_bullet(doc, "sudo = \"faça como administrador\". Pede a senha da equipe (invisível de novo)")
    add_bullet(doc, "apt é a \"loja de aplicativos\" do servidor, por texto. update atualiza a vitrine; install instala. Liga com o Bloco 4: quem mantém essa loja é a distribuição")
    add_bullet(doc, "Do you want to continue? [Y/n]: a maiúscula é o padrão, só Enter já é sim. status: active (running) em verde, sai com q")
    add_para(doc, "Cada um conecta o celular no Wi-Fi da equipe (dados móveis desligados) e abre http:// e o endereço do servidor. Aparece \"Welcome to nginx!\". Deixe comemorar. \"Essa página não veio da internet. Veio daquela máquina velha, pelo cabo, pelo roteador, pelo Wi-Fi. Vocês são a cozinha agora.\"", italic=True, color=COR_PRIMARIA)
    add_heading(doc, "Plano B: sem internet não instala o Nginx", level=2)
    add_cmd(doc, "mkdir site")
    add_cmd(doc, "nano site/index.html")
    add_cmd(doc, "sudo python3 -m http.server 80 --directory site")
    add_para(doc, "O Ubuntu já vem com Python, que tem um servidor de páginas embutido. Fica ligado enquanto a tela estiver aberta; Ctrl+C desliga. Instale o Nginx no começo do Encontro 2.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 11
    add_heading(doc, "BLOCO 11 — PRÁTICA · A PÁGINA DA EQUIPE · 15 MIN")
    add_cmd(doc, SITE[3][0])
    add_para(doc, "O nano é um editor de texto dentro do terminal. A equipe digita as quatro linhas da folha, trocando o N e os nomes. Hoje sem acento e sem ç:")
    for l in PAGINA:
        add_cmd(doc, l, size=10)
    add_bullet(doc, "Salvar: Ctrl+O e Enter. Sair: Ctrl+X. No rodapé do nano, ^ quer dizer Ctrl")
    add_bullet(doc, "<h1> é título grande, <p> é parágrafo. Abre com <h1> e fecha com </h1>")
    add_bullet(doc, "Troca o digitador: cada um da equipe acrescenta uma linha (o time, a música, um recado)")

    add_hr(doc)

    # BLOCO 12
    add_heading(doc, "BLOCO 12 — VISITA, DESLIGAR E FECHAMENTO · 10 MIN")
    add_bullet(doc, "Visita (5 min): cada um conecta o celular no Wi-Fi de outra equipe (ficha de entrega da Noite 8) e abre o site dela. No próprio Wi-Fi não abre: \"cada roteador é uma casa. Esse site está na REDE, não na INTERNET (Noite 1). Para o escritório, é isso que se quer: página interna\"")
    add_cmd(doc, "sudo poweroff")
    add_bullet(doc, "\"Servidor não se desliga no botão nem na tomada: ele pode estar escrevendo no disco\"")
    add_bullet(doc, "Volte à enquete do Bloco 1: \"quem sabe dizer agora o que é um servidor? E o Linux?\". As três frases (slide final)")
    add_bullet(doc, "Encontro 2: o PC com Windows vai para a bancada, e a turma manda no servidor de longe, pelo PuTTY. Organiza as pastas, cria os funcionários e faz o site de um negócio")

    add_hr(doc)

    add_heading(doc, "SE TRAVAR")
    add_table(doc, [["O que aconteceu", "O que fazer"]] + [list(r) for r in TRAVOU], col_widths=[5.5, 11])

    add_hr(doc)

    add_heading(doc, "SE PERGUNTAREM")
    add_para(doc, "\"Se o Linux é tão bom, por que todo mundo usa Windows?\"", bold=True)
    add_para(doc, "> \"Porque o Windows já vem instalado quando se compra o computador, e os programas mais conhecidos de escritório e jogos foram feitos para ele. No servidor é o contrário: quase tudo roda Linux. Cada um é melhor num lugar.\"", italic=True)
    add_para(doc, "\"Dá pra abrir esse site de casa?\"", bold=True)
    add_para(doc, "> \"Dá, mas aí o servidor fica exposto pra internet inteira e precisa de outros cuidados: endereço do provedor, liberar porta no roteador, segurança. É outro serviço. O que a gente faz é o que o escritório pediu: página interna.\"", italic=True)
    add_para(doc, "\"Linux pega vírus?\"", bold=True)
    add_para(doc, "> \"Pode ter problema de segurança como qualquer sistema. O que protege é atualizar (o apt faz isso), senha forte e só instalar o necessário.\"", italic=True)
    add_para(doc, "\"Dá pra instalar Linux no meu computador de casa?\"", bold=True)
    add_para(doc, "> \"Dá. Para o dia a dia, Ubuntu Desktop ou Linux Mint. Mas instalar no disco inteiro apaga o Windows, como vocês viram hoje. Primeiro teste pelo pendrive, sem instalar, e faça cópia dos arquivos.\"", italic=True)
    add_para(doc, "\"O Linus Torvalds ficou rico?\"", bold=True)
    add_para(doc, "> \"Não como o dono da Microsoft. Ele trabalha até hoje coordenando o Linux, pago por uma fundação que as empresas mantêm. Quem ganha dinheiro com Linux é quem presta serviço com ele.\"", italic=True)
    add_para(doc, "\"Isso é profissão?\"", bold=True)
    add_para(doc, "> \"É. Escritório, clínica, escola, loja: alguém instala e cuida do servidor de arquivos, das câmeras, da página interna. Quem sabe Linux e documenta o serviço tem trabalho, aqui e a distância.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "NÃO reinstale nem formate: o Encontro 2 continua de onde parou, na mesma máquina, com o mesmo usuário")
    add_bullet(doc, "Guarde cada servidor junto com o roteador da equipe e a ficha de entrega")
    add_bullet(doc, "Tire foto da ficha técnica de cada equipe (página 4): se um aluno faltar ou perder a folha, a senha do servidor está salva")
    add_bullet(doc, "Anote quantas equipes chegaram ao site no celular. É o indicador do encontro")

    add_hr(doc)

    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Iniciei o estudo de servidores com Linux apresentando a situação-problema do módulo e o produto a ser construído. Trabalhei a diferença entre cliente e servidor a partir do roteador configurado na unidade anterior e de exemplos do cotidiano, e os estudantes classificaram casos em dupla. Apresentei a função do sistema operacional, a história do Linux e do software livre e o conceito de distribuição, e os estudantes identificaram o núcleo Linux em seus próprios celulares. Expliquei a inicialização do computador e o processo de instalação. Em equipes, com rodízio de funções, os estudantes instalaram o Ubuntu Server 24.04 em computadores reaproveitados, por meio de pendrive, interpretando as telas do instalador em inglês técnico e registrando cada escolha, e elaboraram um glossário dos termos do instalador. Acessaram o sistema pelo terminal e levantaram nome, endereço de rede, memória e espaço em disco da máquina, testando a comunicação com o roteador e com a internet. Instalaram o servidor web Nginx pelo gerenciador de pacotes, publicaram uma página interna da equipe e verificaram o acesso pelo celular, discutindo a diferença entre rede local e internet. Iniciaram a ficha técnica do servidor e desligaram o sistema de forma segura.")

    add_hr(doc)

    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou para usar no começo do Encontro 2:", italic=True, color=COR_CINZA)
    add_para(doc, "1. O servidor trabalhando ao vivo", bold=True)
    add_para(doc, "sudo apt install htop e depois htop: processador e memória em tempo real. Enquanto a equipe recarrega a página no celular várias vezes, o digitador olha o htop. Sai com q. Por que uma máquina antiga dá conta? 15 min.")
    add_para(doc, "2. Quanto custa deixar ligado", bold=True)
    add_para(doc, "Potência da fonte (ou 60 W) × 24 h × 30 dias em kWh, vezes a tarifa da fatura de casa. E se ficar ligado só no horário do escritório? Matemática e Ciências da Natureza. 20 min.")
    add_para(doc, "3. O cartaz da cozinha", bold=True)
    add_para(doc, "Cada equipe desenha o caminho da página: servidor → cabo → roteador → Wi-Fi → celular, com os endereços de verdade da equipe em cada ponto. Apresentação de 2 minutos. 25 min.")
    add_para(doc, "4. A linha do tempo do Linux", bold=True)
    add_para(doc, "Em dupla, com o celular, cada dupla pesquisa um fato (o primeiro Android, o pinguim Tux, por que o Ubuntu tem nome africano, o que é a Red Hat) e escreve num papel com o ano. A turma monta a linha do tempo na parede. 20 min.")
    add_para(doc, "5. Decifrar as mensagens", bold=True)
    add_para(doc, "No quadro, mensagens que apareceram na noite (command not found, Could not get lock, active (running), Temporary failure in name resolution). As equipes traduzem e escrevem o que o técnico faz diante de cada uma. 15 min.")

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

    # Capa + abertura do módulo
    add_slide_cover(prs, "SERVIDORES COM LINUX · ENCONTRO 1", "INSTALANDO", "O UBUNTU",
                    "e o primeiro site, aberto no celular de vocês",
                    "Curso FIC · Montador e Reparador de Computadores · CEJA Itapiranga – SC")

    lista("ANTES DE COMEÇAR", "Levante a mão", [
        "•  Quem já ouviu a palavra servidor? Quem sabe dizer o que é?",
        "•  Quem já ouviu falar de Linux? Quem já usou?",
        "•  Quem já viu uma tela preta só com letras?",
    ], size=24, nota="No fim do módulo a gente faz as mesmas perguntas.")

    lista("A DEMANDA DO CLIENTE", "O escritório pequeno", [
        "\"Queremos guardar os arquivos num lugar só,",
        "acessar de qualquer computador",
        "e ter uma página interna para os funcionários.\"",
        "",
        "Como transformar um computador antigo nesse servidor?",
    ], size=24)

    tabela("O MÓDULO", "Três encontros, um servidor por equipe", [["Quando", "O que a equipe faz"]] + [list(m) for m in MODULO],
           col_widths=[3.0, 8.9], size=18, height=2.8,
           nota="No fim, outra equipe testa o seu servidor só com a ficha técnica.")

    lista("COMO VAI FUNCIONAR", "Equipe, rodízio e ficha", [
        "•  Equipes de 3 ou 4. A mesma máquina até o fim",
        "•  Digitador, leitor e anotador. Troca a cada parte",
        "•  Enter só depois do leitor dizer \"confere\"",
        "•  A ficha técnica começa hoje e é entregue no último encontro",
        "•  Errar faz parte: a mensagem de erro diz o que aconteceu",
    ], size=22)

    # O que é um servidor
    lista("O QUE É UM SERVIDOR", "Vocês já configuraram um", [
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
        ["Quando funciona", "Quando a pessoa usa", "O tempo todo"],
        ["Exemplos", "Celular, notebook, TV", "YouTube, banco, WhatsApp, o roteador"],
    ], col_widths=[2.9, 4.5, 4.5], size=17, height=3.6,
        nota="A foto do Instagram não está no seu celular. Está num servidor.")

    tabela("OS TIPOS", "Quem aqui usou isso hoje?", [["Servidor", "O que faz"]] + [list(t) for t in TIPOS_SERVIDOR],
           col_widths=[4.3, 7.6], size=14, height=4.0, top=2.3)

    lista("ONDE ELES FICAM", "Do galpão gigante à salinha do escritório", [
        "•  Data center: galpões cheios de servidores empilhados,",
        "   com ar-condicionado e energia de reserva. YouTube, banco, WhatsApp",
        "•  Servidor pequeno: numa sala do escritório, da clínica, da escola",
        "",
        "Por que um computador velho serve?",
        "Um escritório de 5 pessoas pede pouco. O que importa é",
        "ficar ligado e não perder os dados.",
    ], size=20)

    frase("A FRASE DA NOITE", "Servidor não é um computador especial.\nÉ um computador\nque trabalha para os outros.", size=34)

    lista("PRÁTICA · 15 MIN", "Cliente ou servidor?", [
        "Página 1 da folha, em dupla:",
        "",
        "•  Doze casos para marcar",
        "•  Depois a gente corrige junto: justifique a resposta",
        "",
        "Um mesmo aparelho pode ser cliente e servidor?",
    ], size=22)

    # Sistema operacional e Linux
    lista("O SISTEMA OPERACIONAL", "O síndico do computador", [
        "O computador é um prédio. As peças são a estrutura.",
        "Os programas são os moradores. O síndico é o sistema operacional:",
        "",
    ] + ["•  " + f for f in SO_FUNCOES], size=19)

    tabela("VOCÊS JÁ USAM", "Os sistemas do dia a dia", [["Sistema", "De quem", "Onde está"]] + [list(s) for s in SO_EXEMPLOS],
           col_widths=[2.4, 4.0, 5.5], size=15, height=3.4,
           nota="O Android é Linux por dentro.")

    tabela("DE ONDE VEM O LINUX", "Um estudante, um computador de casa e o mundo todo ajudando",
           [["Quando", "O que aconteceu"]] + [list(h) for h in LINUX_HISTORIA],
           col_widths=[1.6, 10.3], size=13, height=4.2, top=2.4)

    lista("SOFTWARE LIVRE", "Não é só de graça: é aberto", [
        "•  Qualquer um pode usar, estudar, mudar e repartir",
        "•  Empresas do mundo todo melhoram o Linux juntas",
        "",
        "Como alguém ganha dinheiro com algo gratuito?",
        "Com o serviço: instalar, configurar, dar suporte, manter.",
        "",
        "É isso que vocês vão aprender a vender.",
    ], size=22)

    tabela("AS DISTRIBUIÇÕES", "O Linux é o motor. A distribuição é o carro pronto",
           [["Distribuição", "Como é"]] + [list(d) for d in DISTROS],
           col_widths=[2.8, 9.1], size=16, height=3.2,
           nota="Hoje: Ubuntu Server 24.04 (abril de 2024), com correções até 2029.")

    lista("DESKTOP × SERVER", "Por que sem janelas?", [
        "•  Ubuntu Desktop: janelas e mouse, para usar no dia a dia",
        "•  Ubuntu Server: só texto",
        "",
        "Ninguém fica olhando a tela de um servidor.",
        "Sem janelas, a máquina velha sobra força para trabalhar.",
        "",
        "Parece filme de hacker. É só escrever ordens, em inglês.",
    ], size=22)

    lista("PRÁTICA · 10 MIN", "O Linux no seu bolso", [
        "Android → Configurações → Sobre o telefone →",
        "   (Informações do software / Todas as especificações) →",
        "   Versão do kernel",
        "",
        "Apareceu um número como 5.10 ou 6.1? É o Linux do seu celular.",
        "",
        "Onde mais tem Linux nesta sala?",
    ], size=21, nota="iPhone não é Linux, mas é parente: família Unix.")

    tabela("COMO O COMPUTADOR LIGA", "Do botão ao sistema", [["Etapa", "O que acontece"]] + [list(b) for b in BOOT],
           col_widths=[3.6, 8.3], size=17, height=3.0,
           nota="Instalar = copiar o sistema do pendrive para o disco.")

    lista("O INSTALADOR", "Inglês e sem mouse", [
        "•  Pendrive na máquina → liga → tecla de boot (está na fita crepe)",
        "•  O instalador fala inglês. A folha traduz cada tela",
        "        Setas e Tab = andar     Espaço = marcar",
        "        Enter = confirmar       Esc = voltar",
        "•  Digitador, leitor e anotador. Enter só depois do \"confere\"",
    ], size=21)

    frase("ANTES DE COMEÇAR", "O instalador vai apagar o disco inteiro.\nAqui as máquinas estão vazias.\nNo cliente, o técnico sempre pergunta\ne faz cópia antes.", cor=R, size=30)

    # Prática
    tabela("PRÁTICA · 40 MIN", "Telas 1 a 8", [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS[:8]],
           col_widths=[3.9, 3.3, 4.7], size=13, height=4.0, top=2.3,
           nota="Tela de rede: anote o número do DHCPv4. É o endereço do servidor.")

    tabela("PRÁTICA · 40 MIN", "Telas 9 a 15", [["Tela", "O que é", "O que escolher"]] + [list(t) for t in TELAS[8:]],
           col_widths=[3.9, 3.3, 4.7], size=13, height=3.8, top=2.3,
           nota="SSH: marca com ESPAÇO. Tem que aparecer [X].", nota_cor=R)

    tabela("O COMBINADO DO PERFIL", "Tudo minúsculo, sem espaço, sem acento", [
        ["Campo", "O que a equipe põe"],
        ["Your name", "Equipe N"],
        ["Your server's name", "servidor-equipeN"],
        ["Pick a username", "equipeN"],
        ["Choose a password", "A equipe inventa (10+ caracteres) e ANOTA"],
    ], col_widths=[4.5, 7.4], size=20, height=3.0)

    lista("ENQUANTO COPIA · 15 MIN", "O dicionário do técnico", [
        "Página 3 da folha:",
        "",
        "•  18 palavras do instalador: traduzam sem celular, depois confiram",
        "•  O palpite do raio-x: o que cada comando vai mostrar?",
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

    tabela("PRÁTICA · 20 MIN", "O raio-x da máquina", [["Comando", "O que mostra"]] + [list(r) for r in RAIOX],
           col_widths=[4.2, 7.7], size=15, height=4.0, top=2.3,
           nota="Seta pra cima repete. Tab completa. Errou? Leia a mensagem.")

    lista("PRÁTICA · 20 MIN", "O primeiro site", [
        "$ sudo apt update",
        "$ sudo apt install nginx",
        "$ systemctl status nginx",
        "",
        "•  sudo = como administrador (pede a senha, invisível)",
        "•  apt = a loja de aplicativos do servidor",
        "•  [Y/n]: Enter já é sim · status: active (running) · sai com q",
    ], size=21)

    frase("NO CELULAR", "Wi-Fi da equipe, dados móveis desligados\nhttp://192.168.1N.1xx\n\nWelcome to nginx!", size=30)

    frase("VOCÊS SÃO A COZINHA", "Essa página não veio da internet.\nVeio daquela máquina velha,\npelo cabo, pelo roteador, pelo Wi-Fi.", size=30)

    lista("PRÁTICA · 15 MIN", "A página da equipe", [
        "$ sudo nano /var/www/html/index.html",
        "",
    ] + PAGINA + [
        "",
        "Salvar: Ctrl+O, Enter · Sair: Ctrl+X · Recarregue o celular",
    ], size=19, nota="Hoje sem acento e sem ç. Cada um da equipe põe uma linha.")

    frase("VISITA", "No seu Wi-Fi, o site da vizinha não abre.\nCada roteador é uma casa.\nEsse site está na REDE, não na INTERNET.", size=28)

    frase("DESLIGAR CERTO", "$ sudo poweroff\n\nNunca no botão, nunca na tomada:\nele pode estar escrevendo no disco.", size=30)

    lista("O QUE FICA", "Três frases pra guardar", [
        "\"Servidor é um computador que trabalha para os outros.\"",
        "",
        "\"Linux é livre: quem ganha é quem presta o serviço.\"",
        "",
        "\"Senha não aparece. Não travou.\"",
    ], size=23)

    lista("NO ENCONTRO 2", "A bancada do técnico", [
        "•  Um PC com Windows ao lado do servidor, no mesmo roteador",
        "•  Mandar no servidor de longe, pelo PuTTY",
        "•  Pastas, funcionários com senha e o site de um negócio",
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

    # PÁGINA 1 — Antes de instalar
    cabecalho_documento(
        doc,
        "SERVIDOR, SISTEMA E LINUX",
        "Servidores com Linux · Encontro 1 · Curso FIC · CEJA Itapiranga · 1 de 4",
    )
    campo(doc, "Nome: _____________________________________________     Equipe: ______   Máquina: ____________________", space_after=6, size=10)

    add_heading(doc, "1  Cliente ou servidor? (em dupla)", level=2)
    add_para(doc, "Cliente pede e mostra. Servidor guarda, prepara e entrega para os outros.", italic=True, size=9, color=COR_CINZA)
    rows = [["Aparelho ou serviço", "Cliente", "Servidor"]] + [[c, "☐", "☐"] for c in CLIENTE_SERVIDOR]
    tbl = add_table(doc, rows, col_widths=[13.4, 2, 2])
    small_table(tbl, size=9)
    campo(doc, "Um mesmo aparelho pode ser cliente e servidor? Dê um exemplo: ________________________________________", size=10)

    add_heading(doc, "2  O síndico do computador", level=2)
    campo(doc, "Cite duas coisas que o sistema operacional faz: _________________________________________________", size=10)
    campo(doc, "Que sistema tem no seu celular? ______________   No computador de casa (se tiver)? ______________", size=10)

    add_heading(doc, "3  O Linux no seu bolso", level=2)
    campo(doc, "Marca do celular: ______________   Versão do kernel (Linux): ______________   ☐ iPhone: olhei no do colega", size=10)
    campo(doc, "A versão mais nova da turma: ________   A mais antiga: ________", size=10)
    campo(doc, "Onde mais tem Linux nesta sala ou na sua casa? ____________________________________________", size=10)
    campo(doc, "O Linux é software livre. O que isso quer dizer, além de ser gratuito?", bold=True, size=10)
    linha(doc)

    # PÁGINA 2 — A instalação
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "INSTALANDO O UBUNTU SERVER",
        "Servidores com Linux · Encontro 1 · 2 de 4",
    )
    campo(doc, "Rodízio · Digitador: ____________   Leitor: ____________   Anotador: ____________   (troca na tela 9)", size=10)
    add_para(doc, "Sem mouse: setas e Tab andam · Espaço marca · Enter confirma · Esc volta. Enter só depois do leitor dizer \"confere\".", italic=True, size=9, color=COR_CINZA)

    tbl = add_table(doc, [["Tela (em inglês)", "O que é", "O que escolher", "✓"]] +
                    [[t[0], t[1], t[2], "☐"] for t in TELAS], col_widths=[5, 3.8, 7.4, 1])
    small_table(tbl, size=9)

    add_heading(doc, "Anote durante a instalação", level=2)
    campo(doc, "Tecla de boot desta máquina: ______   Tela de rede · número ao lado de DHCPv4: ____ . ____ . ____ . ____", size=10)
    campo(doc, "O mirror passou no teste (\"passed tests\")?  ☐ Sim  ☐ Não      Tamanho do disco que apareceu: ________ GB", size=10)
    campo(doc, "Nome do servidor: ______________________   Usuário: ______________   Senha: ______________________", size=10)
    campo(doc, "Hora em que começou a cópia: ____ : ____     Hora do Reboot Now: ____ : ____     Levou ______ minutos", size=10)

    # PÁGINA 3 — Enquanto copia + raio-x
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "O DICIONÁRIO DO TÉCNICO E O RAIO-X",
        "Servidores com Linux · Encontro 1 · 3 de 4",
    )

    add_heading(doc, "1  Enquanto copia: o dicionário do técnico", level=2)
    add_para(doc, "Traduza sem o celular. Depois confira com a turma.", italic=True, size=9, color=COR_CINZA)
    metade = (len(DICIONARIO) + 1) // 2
    rows = [["Na tela", "Quer dizer", "Na tela", "Quer dizer"]]
    for i in range(metade):
        a = DICIONARIO[i][0]
        b = DICIONARIO[i + metade][0] if i + metade < len(DICIONARIO) else ""
        rows.append([a, "", b, ""])
    tbl = add_table(doc, rows, col_widths=[4.2, 4.5, 4.2, 4.5])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)
    campo(doc, "Que palavra na tela do disco avisa que tudo vai ser apagado? ______________________", size=10)

    add_heading(doc, "2  O raio-x da máquina (depois do intervalo)", level=2)
    add_para(doc, "Antes de rodar, escreva o palpite. Depois rode e anote o resultado. Troque N pelo número da equipe.", italic=True, size=9, color=COR_CINZA)
    tbl = add_table(doc, [["Comando", "O que mostra", "Palpite", "O que deu"]] +
                    [[c, d, "", ""] for c, d in RAIOX[:7]], col_widths=[3.8, 5.4, 3.8, 4.4])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.8)
    campo(doc, "O endereço do ip a bate com o da tela de rede do instalador?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "Disco: tem ______ GB e estão livres ______ GB. Cabem quantas fotos de 3 MB? (1 GB ≈ 1000 MB) ________", size=10)
    campo(doc, "Para pingar google.com, o servidor precisou perguntar pra quem? (Noite 7) ______________________", size=10)

    # PÁGINA 4 — O site, a visita e a ficha técnica
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "O PRIMEIRO SITE",
        "Servidores com Linux · Encontro 1 · 4 de 4",
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

    add_heading(doc, "3  Visita", level=2)
    campo(doc, "Site da equipe ____ · No meu Wi-Fi abriu?  ☐ Sim  ☐ Não · No Wi-Fi dela abriu?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "Por que alguém em casa não consegue abrir o site da equipe? ________________________________________", size=10)

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
        "Resumo do aluno · Servidores com Linux · Encontro 1 · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Servidor não é um computador especial. É um computador que fica ligado trabalhando para os outros: guarda, prepara e entrega. Cliente é quem pede e mostra (celular, notebook, TV). Hoje um computador antigo virou servidor, com Linux, e entregou uma página para os celulares da rede.")

    add_heading(doc, "Tipos de servidor", level=2)
    add_table(doc, [["Servidor", "O que faz"]] + [list(t) for t in TIPOS_SERVIDOR], col_widths=[5.5, 10])

    add_heading(doc, "Sistema operacional e Linux", level=2)
    add_bullet(doc, "O sistema operacional é o síndico do computador: conversa com as peças, divide a máquina entre os programas, guarda os arquivos e controla quem entra")
    add_bullet(doc, "Windows, macOS, Android e Linux são sistemas operacionais. O Android é Linux por dentro")
    add_bullet(doc, "1991: Linus Torvalds, estudante na Finlândia, cria o Linux e divulga de graça. Junto com o projeto GNU (1983), vira um sistema completo")
    add_bullet(doc, "Software livre: qualquer um pode usar, estudar, mudar e repartir. Quem ganha dinheiro é quem presta o serviço")
    add_bullet(doc, "Distribuição é o Linux pronto para usar: Ubuntu, Debian, Linux Mint, Fedora. Hoje: Ubuntu Server 24.04 (abril de 2024, correções até 2029), sem janelas")

    add_heading(doc, "Como o computador liga", level=2)
    add_para(doc, "Botão → BIOS/UEFI testa as peças e procura de onde ligar → a tecla de boot escolhe o pendrive → o sistema carrega. Instalar é copiar o sistema do pendrive para o disco.", size=10.5)

    add_heading(doc, "A instalação em 6 passos", level=2)
    add_bullet(doc, "1. Pendrive na máquina, liga e aperta a tecla de boot (F12, F11, F8 ou Esc)")
    add_bullet(doc, "2. Idioma English, teclado Portuguese (Brazil), tipo Ubuntu Server")
    add_bullet(doc, "3. Rede: anotar o endereço que aparece (DHCPv4). Quem deu esse número foi o roteador")
    add_bullet(doc, "4. Disco inteiro. ATENÇÃO: apaga tudo. No cliente, sempre perguntar e fazer cópia antes")
    add_bullet(doc, "5. Perfil: nome do servidor, usuário e senha (tudo minúsculo, sem espaço). Marcar Install OpenSSH server")
    add_bullet(doc, "6. Esperar a cópia, Reboot Now, tirar o pendrive e Enter")

    add_heading(doc, "Os comandos da noite", level=2)
    tbl = add_table(doc, [["Comando", "O que faz"]] + [list(r) for r in RAIOX] + [list(r) for r in SITE],
                    col_widths=[6, 9.5])
    mono_col(tbl, 0, size=9.5)

    add_heading(doc, "Cuidados do primeiro contato", level=2)
    add_bullet(doc, "A senha NÃO aparece quando você digita. Não travou")
    add_bullet(doc, "Maiúscula e minúscula são diferentes no Linux")
    add_bullet(doc, "\"command not found\" = \"não conheço esse comando\". Leia a mensagem")
    add_bullet(doc, "No celular: http:// (não https), no Wi-Fi da mesma rede do servidor")
    add_bullet(doc, "Servidor se desliga com sudo poweroff. Nunca no botão nem na tomada")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Servidor é um computador que trabalha para os outros.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Linux é livre: quem ganha é quem presta o serviço.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Senha não aparece. Não travou.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Na sua casa, quais aparelhos são clientes? Tem algum que trabalha como servidor (roteador, DVR de câmera)?")
    add_bullet(doc, "Explique para alguém da família o que é um sistema operacional, usando o síndico do prédio")
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
