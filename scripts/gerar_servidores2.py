"""Gera os 4 arquivos do Encontro 2 do módulo de Servidores (Plano Mensal 4).

Servidores 2 — Pastas, usuários e acesso de longe
Continua das máquinas instaladas pelos alunos no Encontro 1 (usuário equipeN,
servidor servidor-equipeN, Nginx com a página da equipe, OpenSSH instalado).
Hoje: andar pelas pastas, montar uma pasta de treino do escritório (mkdir, nano,
cp, mv, rm, tree), fixar o endereço do servidor por reserva no roteador da
Noite 8, entrar pelo celular e pelo Windows (SSH), criar os usuários e grupos
dos funcionários e travar as pastas do escritório com permissões. A pasta
/srv/escritorio criada hoje é a que vira pasta compartilhada (Samba) no
Encontro 3. Padrão prática-pesada, sem role-play e sem duelo.
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
from gerar_servidores1 import add_cmd, mono_col, small_table, campo, linha, MONO
from pptx.enum.text import PP_ALIGN

OUTDIR = "/home/nedel/curso"
TITULO_ARQ = "Servidores 2 - Pastas, usuarios e acesso de longe"
FOOTER = "Servidores com Linux · Encontro 2"


# O passeio pelas pastas
PASSEIO = [
    ("pwd", "Onde eu estou?"),
    ("ls -a", "O que tem aqui? (-a mostra os escondidos, com ponto na frente)"),
    ("cd /", "Vai para a raiz, a porta de entrada de tudo"),
    ("ls", "As pastas principais do sistema"),
    ("cd /var/www/html", "Vai direto para a pasta do site"),
    ("cat index.html", "Mostra o conteúdo do arquivo (a página da equipe!)"),
    ("cd ..", "Sobe uma pasta"),
    ("cat /etc/hostname", "O arquivo onde está guardado o nome do servidor"),
    ("cd ~", "Volta para casa (a pasta da equipe)"),
]

# As pastas principais da raiz
RAIZ = [
    ("/home", "As casas dos usuários. Cada um tem a sua"),
    ("/etc", "As configurações do sistema e dos programas"),
    ("/var", "O que muda o tempo todo: o site, os registros (logs)"),
    ("/srv", "Os dados dos serviços. O escritório vai morar aqui"),
    ("/usr e /bin", "Os programas"),
    ("/root", "A casa do administrador principal"),
]

# O pedido do escritório: pasta de treino
PEDIDO = [
    ("mkdir treino", "Criar a pasta de treino"),
    ("cd treino", "Entrar nela"),
    ("mkdir financeiro recepcao publico", "Criar três pastas de uma vez"),
    ("nano recepcao/recados.txt", "Escrever dois recados (salvar: Ctrl+O, Enter, Ctrl+X)"),
    ("cat recepcao/recados.txt", "Conferir o que foi escrito"),
    ("cp recepcao/recados.txt publico/", "Copiar os recados para o público"),
    ("nano financeiro/contas.txt", "Escrever três contas a pagar"),
    ("mv financeiro/contas.txt financeiro/contas-dezembro.txt", "Renomear (mv move ou muda o nome)"),
    ("nano errado.txt", "Criar um arquivo qualquer"),
    ("rm errado.txt", "Apagar. Não tem lixeira!"),
    ("sudo apt install tree", "Instalar o programa que desenha a árvore"),
    ("tree ~/treino", "Conferir com o desenho da folha"),
]

ARVORE = [
    "treino",
    "├── financeiro",
    "│   └── contas-dezembro.txt",
    "├── publico",
    "│   └── recados.txt",
    "└── recepcao",
    "    └── recados.txt",
]

# Usuários e grupos do escritório
FUNCIONARIOS = [
    ("ana", "Ana", "financeiro"),
    ("bruno", "Bruno", "financeiro"),
    ("carla", "Carla", "recepcao"),
]

USUARIOS_CMD = [
    ("sudo adduser ana", "Cria a usuária ana (pede senha 2 vezes, o resto Enter, Y no fim)"),
    ("sudo groupadd financeiro", "Cria o grupo financeiro"),
    ("sudo usermod -aG financeiro ana", "Põe a ana no grupo financeiro (-aG: acrescenta)"),
    ("groups ana", "Mostra em quais grupos a ana está"),
]

PASTAS_CMD = [
    "sudo mkdir -p /srv/escritorio/financeiro /srv/escritorio/recepcao /srv/escritorio/publico",
    "sudo chgrp financeiro /srv/escritorio/financeiro",
    "sudo chgrp recepcao /srv/escritorio/recepcao",
    "sudo chmod 770 /srv/escritorio/financeiro /srv/escritorio/recepcao",
    "sudo chmod 755 /srv/escritorio/publico",
    "ls -l /srv/escritorio",
]

# Os números das permissões
NUMEROS = [
    ("7", "rwx", "4+2+1", "Lê, altera e entra/executa: tudo"),
    ("5", "r-x", "4+0+1", "Lê e entra, mas não altera"),
    ("4", "r--", "4+0+0", "Só lê"),
    ("0", "---", "0", "Nada"),
]

# O teste das permissões
TESTE = [
    ("ana", "entra em financeiro e cria um arquivo", "Consegue"),
    ("ana", "entra em recepcao", "Permission denied"),
    ("ana", "entra em publico e lê", "Consegue"),
    ("ana", "cria arquivo em publico", "Permission denied"),
    ("carla", "entra em recepcao e cria um arquivo", "Consegue"),
    ("carla", "entra em financeiro", "Permission denied"),
]

# SSH
SSH_CMD = [
    ("ip a", "Na linha link/ether: o MAC do servidor"),
    ("sudo reboot", "Reinicia para pegar o endereço reservado"),
    ("ssh equipeN@192.168.1N.10", "Entrar de longe (Windows: no Prompt de Comando)"),
    ("yes", "Na primeira vez: \"confia nesta máquina?\""),
    ("wall Oi da equipe N", "Manda um recado pra tela do servidor"),
    ("exit", "Sai do servidor (a sessão de longe acaba)"),
]

TRAVOU = [
    ("Connection timed out", "O celular ou o computador não está na rede da equipe, ou o endereço está errado. Wi-Fi da equipe, dados móveis desligados"),
    ("Connection refused", "O SSH não ficou instalado. No servidor: sudo apt install openssh-server"),
    ("Permission denied, please try again", "Usuário ou senha errados. Tudo minúsculo. A senha não aparece"),
    ("REMOTE HOST IDENTIFICATION HAS CHANGED", "Aquele endereço já foi de outra máquina. No Windows: ssh-keygen -R 192.168.1N.10"),
    ("O endereço não mudou depois da reserva", "Conferir o MAC digitado no roteador, salvar e sudo reboot de novo. Se insistir, ver nota do professor no Bloco 5"),
    ("No such file or directory", "Caminho errado. Use pwd para ver onde está e Tab para completar"),
    ("tree: command not found", "Falta instalar: sudo apt install tree"),
    ("ana is not in the sudoers file", "Esperado: funcionário não é administrador. É o teste dando certo"),
    ("Permission denied numa pasta que devia abrir", "Usuário entrou antes de ser posto no grupo. exit e entrar de novo"),
    ("Apagou com rm o que não devia", "Não tem lixeira. Refazer. Por isso o treino é numa pasta só de treino"),
]

FICHA_NOVA = [
    "Endereço fixo do servidor (reserva)",
    "MAC do servidor",
    "Como entrar de longe (SSH)",
    "Usuários criados e grupos",
    "Pastas do escritório e quem acessa",
    "Testado por / data",
]


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "PASTAS, USUÁRIOS E ACESSO DE LONGE",
        "Roteiro do professor · Servidores com Linux · Encontro 2 de 3 · sexta 04/12 · Curso FIC · CEJA Itapiranga – SC",
    )

    add_heading(doc, "A IDEIA DESTE ENCONTRO")
    add_para(doc, "No Encontro 1 a turma instalou o servidor e publicou a página. Hoje ela aprende a morar no terminal e prepara a máquina para o escritório: as pastas, os funcionários e quem pode mexer em quê. E descobre que não precisa mais ficar na frente do servidor: entra nele pelo celular, do outro lado da sala.")
    add_para(doc, "O fio da noite é a demanda do escritório. Tudo que a turma cria hoje (os usuários ana, bruno e carla, os grupos e a pasta /srv/escritorio) é usado no Encontro 3, quando essa pasta vira pasta compartilhada no Windows e no celular.")
    add_para(doc, "A frase da noite: \"Cada um entra com o seu nome e só abre a gaveta que é sua.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "O que se resolve neste encontro", level=2)
    add_para(doc, "Cada aluno sai sabendo onde está e como andar entre as pastas (pwd, ls, cd), criar, ler, copiar, renomear e apagar arquivos (mkdir, nano, cat, cp, mv, rm), pedir ajuda a um comando, fixar o endereço do servidor com reserva no roteador, entrar no servidor de longe pelo celular e pelo Windows (SSH), criar usuários e grupos e ler e aplicar permissões em pastas.")

    add_heading(doc, "Onde este encontro fica no módulo", level=2)
    add_para(doc, "Usa as máquinas do Encontro 1 sem reinstalar. Retoma a reserva por MAC da Noite 8 (agora para o servidor, não para a \"impressora\") e o gateway e o ping da Noite 6. A senha forte é a da Noite 10. A pasta /srv/escritorio e os usuários de hoje são a base do Encontro 3 (Samba e entrega ao cliente).")
    add_para(doc, "Não entre em: chave SSH no lugar de senha, editar o arquivo sudoers, permissões especiais (setgid, sticky), ACL, netplan e IP fixo configurado no próprio servidor, redirecionamento de portas. O endereço fixo de hoje é pela reserva no roteador, que eles já sabem fazer.", italic=True, color=COR_ALERTA)

    add_hr(doc)

    add_heading(doc, "ANTES DA AULA")
    add_heading(doc, "As máquinas e a rede", level=2)
    add_bullet(doc, "Cada equipe volta para a MESMA máquina e o MESMO roteador do Encontro 1. Ligue tudo e confira se o login com equipeN funciona (as senhas estão nas fotos das fichas técnicas que você tirou)")
    add_bullet(doc, "Internet nos roteadores: hoje se instala o tree pelo apt")
    add_bullet(doc, "Separe as fichas de entrega da Noite 8: a reserva do Bloco 5 precisa da senha de administrador de cada roteador")
    add_bullet(doc, "Plano de endereço fixo: o servidor da equipe N fica em 192.168.1N.10 (fora da faixa .100 a .150 da Noite 8, como foi a reserva da \"impressora\" em .50)")

    add_heading(doc, "O SSH do lado do cliente", level=2)
    add_bullet(doc, "Celular Android: ConnectBot (gratuito, sem cadastro). iPhone: Termius. Peça no grupo da turma, durante a semana, que instalem. Na aula, quem não instalou usa o Wi-Fi da equipe para baixar")
    add_bullet(doc, "Windows: o Prompt de Comando do Windows 10 e 11 já tem o comando ssh. O computador precisa estar na rede da equipe: um computador do laboratório (ou o seu notebook) ligado por cabo numa porta LAN do roteador da equipe, com os cabos testados da Noite 5")
    add_bullet(doc, "Teste antes num computador do laboratório: abra o Prompt e digite ssh. Se aparecer \"usage: ssh…\", está pronto")

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "Folha do aluno (3 páginas), uma por pessoa")
    add_bullet(doc, "Resumo do aluno, uma por pessoa")
    add_bullet(doc, "As folhas do Encontro 1 voltam: a ficha técnica (página 3) é completada hoje")

    add_hr(doc)

    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Retomada: ligar, entrar, a página ainda abre?", "15 min", "Login, a tarefa dos serviços, o que vem hoje"],
        ["2. O armário de pastas", "10 min", "A raiz, as pastas principais, caminho completo e caminho curto"],
        ["3. PRÁTICA — O passeio pelas pastas", "25 min", "pwd, ls, cd, cat, a pasta do site, ajuda dos comandos"],
        ["4. PRÁTICA — A pasta de treino do escritório", "30 min", "mkdir, nano, cp, mv, rm, tree. Conferir com o desenho"],
        ["INTERVALO", "15 min", ""],
        ["5. PRÁTICA — Endereço fixo do servidor", "15 min", "O MAC do servidor, reserva no roteador, reboot, conferir"],
        ["6. PRÁTICA — Entrar de longe (SSH)", "35 min", "Pelo celular e pelo Windows. Recado na tela, mudar o site de longe"],
        ["7. PRÁTICA — Os funcionários", "30 min", "adduser, groupadd, usermod. A ana tenta usar sudo"],
        ["8. PRÁTICA — As gavetas do escritório", "25 min", "/srv/escritorio, chgrp, chmod, ls -l. O teste de quem abre o quê"],
        ["9. Fechamento + ficha técnica", "10 min", "Completar a ficha, três frases, o Encontro 3"],
    ], col_widths=[6.8, 2, 7.7])
    add_para(doc, "Blocos 3 a 8: 160 min de atividade do aluno (76% do total). Exposição: os Blocos 1 e 2 (25 min) e falas curtas dentro das práticas. Mantenha o rodízio do Encontro 1: digitador, leitor, anotador.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — RETOMADA · 15 MIN")
    add_bullet(doc, "Cada equipe liga o servidor e entra com equipeN. \"O que acontece quando digita a senha?\" (nada aparece)")
    add_bullet(doc, "Um celular por equipe no Wi-Fi da equipe: a página do Encontro 1 ainda abre? (sim: o Nginx liga sozinho junto com o servidor)")
    add_bullet(doc, "A tarefa: cada um diz um serviço que usou na semana e que depende de servidor. Anote no quadro. Pergunte de dois deles: \"o que aconteceria se esse servidor desligasse?\"")
    add_para(doc, "\"Hoje o servidor vai ganhar funcionários. Cada um vai entrar com o próprio nome e só vai abrir a gaveta que é dele. E vocês vão mandar no servidor pelo celular, sem encostar no teclado.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — O ARMÁRIO DE PASTAS · 10 MIN")
    add_para(doc, "Desenhe no quadro uma árvore de cabeça para baixo: a raiz / em cima, as pastas principais embaixo.")
    add_table(doc, [["Pasta", "O que guarda"]] + [list(r) for r in RAIZ], col_widths=[3.5, 13])
    add_bullet(doc, "No Linux não tem C: nem D:. Tudo começa na raiz /. O pendrive e o segundo disco aparecem como pastas dentro dela")
    add_bullet(doc, "Caminho completo é o endereço inteiro: /var/www/html (\"Rua X, 123, Itapiranga\"). Funciona de qualquer lugar")
    add_bullet(doc, "Caminho curto parte de onde você está: recepcao/recados.txt (\"a casa do lado\"). Só funciona se você estiver no lugar certo")
    add_bullet(doc, ".. é a pasta de cima. ~ é a sua casa")

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — PRÁTICA · O PASSEIO PELAS PASTAS · 25 MIN")
    add_para(doc, "Página 1 da folha. O anotador escreve o que apareceu em cada passo.")
    add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in PASSEIO], col_widths=[5, 11.5])
    add_bullet(doc, "Em ls -a, as linhas com ponto na frente (.bashrc, .profile) são arquivos escondidos de configuração. Não mexer hoje")
    add_bullet(doc, "Em cat index.html, a equipe vê a página dela como texto: \"o navegador lê exatamente isso e desenha\"")
    add_bullet(doc, "Depois do passeio, ajuda: ls --help (a lista de opções) e man ls (o manual, sai com q). Peça que achem na ajuda o que faz ls -l e ls -h")
    add_bullet(doc, "Erro de propósito: cd /pasta-que-nao-existe. Lê a mensagem: No such file or directory = não tem esse arquivo ou pasta")
    add_para(doc, "Onde dilatar: o leitor fala só o destino (\"vai pra pasta do site\", \"volta pra casa\", \"sobe uma\") e o digitador tem que achar o comando sozinho. Depois trocam.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · A PASTA DE TREINO DO ESCRITÓRIO · 30 MIN")
    add_para(doc, "Antes de mexer no escritório de verdade, a equipe treina numa pasta que pode estragar. Página 1, a lista de 12 passos. O digitador troca a cada 4 passos.")
    add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in PEDIDO], col_widths=[7.5, 9])
    add_para(doc, "No fim, tree ~/treino tem que desenhar exatamente isto (também está na folha):")
    for l in ARVORE:
        add_cmd(doc, l, size=10)
    add_bullet(doc, "Pare a turma no rm: \"No Linux não tem lixeira. O rm apaga de vez. Na casa do cliente, antes de apagar, o técnico confere duas vezes com ls e pwd onde está.\"")
    add_bullet(doc, "rm -r apaga pasta com tudo dentro. Mostre, mas não deixe usar fora da pasta treino")
    add_bullet(doc, "Desenho diferente? A equipe descobre sozinha o que faltou e conserta. É o laudo da atividade")
    add_para(doc, "Onde dilatar: a equipe desenha no papel uma estrutura nova (o escritório com uma pasta de clientes e uma de modelos de contrato, por exemplo), troca o desenho com a equipe do lado, e cada uma constrói o desenho da outra no próprio servidor. Confere com tree.", italic=True, color=COR_DESTAQUE)

    add_heading(doc, "INTERVALO — 15 MIN", level=2)

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — PRÁTICA · ENDEREÇO FIXO DO SERVIDOR · 15 MIN")
    add_para(doc, "\"Hoje vocês vão entrar no servidor pelo endereço dele. Se o roteador der outro número amanhã, ninguém acha o servidor: nem o SSH, nem o site, nem a pasta do Encontro 3. Servidor tem que ter endereço fixo. Vocês já sabem fazer isso: a reserva da Noite 8.\"", italic=True, color=COR_PRIMARIA)
    add_bullet(doc, "No servidor: ip a. Na placa que não é lo, a linha link/ether mostra o MAC (ex.: 00:1a:2b:3c:4d:5e). Anotar com cuidado")
    add_bullet(doc, "No celular, painel do roteador da equipe (192.168.1N.1, senha da ficha de entrega da Noite 8) → DHCP → Reserva de endereço → MAC do servidor → 192.168.1N.10. Salvar")
    add_bullet(doc, "No servidor: sudo reboot. Depois de ligar, ip a: o endereço tem que ser .10")
    add_bullet(doc, "O site mudou de endereço: abrir http://192.168.1N.10 no celular. Atualizar a ficha técnica")
    add_para(doc, "Nota do professor: se a reserva não pegar mesmo com o MAC certo, o servidor está se identificando pro roteador por outro código (client-id). Conserto, feito por você: em /etc/netplan/50-cloud-init.yaml, abaixo de dhcp4: true, acrescente a linha dhcp-identifier: mac (mesmo recuo), depois sudo netplan apply. Não mostre para a turma.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — PRÁTICA · ENTRAR DE LONGE (SSH) · 35 MIN")
    add_para(doc, "Página 2 da folha. SSH é uma conexão segura (com tudo embaralhado no caminho) até o terminal do servidor. É o mesmo terminal, só que dentro do celular.")
    add_heading(doc, "Pelo celular — 15 min", level=2)
    add_bullet(doc, "Celular no Wi-Fi da equipe. No ConnectBot (ou Termius): novo host, equipeN@192.168.1N.10, porta 22")
    add_bullet(doc, "Na primeira vez ele pergunta se confia nessa máquina e mostra um código (a \"impressão digital\" do servidor). Aceitar. Ele guarda e, se um dia outra máquina se passar por essa, ele avisa")
    add_bullet(doc, "Senha da equipe. Entrou: o mesmo equipeN@servidor-equipeN:~$")
    add_bullet(doc, "A prova: pelo celular, wall Oi da equipe N. O recado aparece na tela do servidor, do outro lado da mesa")
    add_bullet(doc, "Agora cada um da equipe entra pelo próprio celular. O servidor aceita várias pessoas ao mesmo tempo (a cozinha atende todas as mesas)")
    add_heading(doc, "Pelo Windows — 10 min", level=2)
    add_bullet(doc, "No computador ligado por cabo no roteador da equipe: Iniciar → cmd → ssh equipeN@192.168.1N.10")
    add_bullet(doc, "Pergunta Are you sure you want to continue connecting (yes/no)? Escrever yes por extenso")
    add_bullet(doc, "Mesmo terminal. Para sair: exit")
    add_heading(doc, "Mudar o site de longe — 10 min", level=2)
    add_bullet(doc, "Pelo SSH (celular ou Windows): sudo nano /var/www/html/index.html, acrescentar uma linha: <p>Alterado de longe pela equipe N</p>")
    add_bullet(doc, "Outro celular recarrega a página. \"O técnico arrumou o site do escritório sem sair da cadeira.\"")
    add_para(doc, "Fale da segurança: \"Agora qualquer um que esteja no Wi-Fi da equipe pode TENTAR entrar no servidor. O que segura é a senha. Por isso a senha do servidor é longa e não fica colada no monitor.\" Liga com a Noite 10.", bold=True, color=COR_ALERTA)
    add_para(doc, "Onde dilatar: last mostra quem entrou e de onde. sudo grep Failed /var/log/auth.log mostra as tentativas com senha errada. Peça que alguém erre a senha de propósito e a equipe ache o registro.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — PRÁTICA · OS FUNCIONÁRIOS · 30 MIN")
    add_para(doc, "Página 3 da folha. Pode ser feito pelo SSH ou na tela do servidor. O escritório tem três funcionários:")
    add_table(doc, [["Usuário", "Nome", "Grupo (setor)"]] + [list(f) for f in FUNCIONARIOS], col_widths=[4, 5, 7.5])
    add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in USUARIOS_CMD], col_widths=[6.5, 10])
    add_bullet(doc, "Repetir para bruno e carla. O grupo recepcao também precisa ser criado (sudo groupadd recepcao)")
    add_bullet(doc, "adduser pergunta senha (duas vezes), nome completo e telefones. Senha: a equipe escolhe uma para cada funcionário (10+ caracteres) e anota. Telefones: só Enter. No fim, Y")
    add_bullet(doc, "No usermod é -aG, com o a. Sem o a, a pessoa sai de todos os outros grupos. Erro clássico")
    add_bullet(doc, "Teste: entrar como ana, pelo celular: ssh ana@192.168.1N.10. whoami responde ana. pwd responde /home/ana: cada um tem a sua casa")
    add_bullet(doc, "Ainda como ana: sudo apt update. Resposta: ana is not in the sudoers file. Leia com a turma: \"a ana não está na lista de administradores\". É o certo: funcionário usa, não administra")
    add_para(doc, "\"Por que não deixar todo mundo entrar com equipeN, que é mais fácil?\" Deixe a turma responder: senha compartilhada vaza, não dá pra saber quem fez o quê, e se a ana sair do escritório, troca a senha de todo mundo. Com um usuário por pessoa, é só apagar o dela.", italic=True, color=COR_PRIMARIA)
    add_para(doc, "Onde dilatar: cada aluno cria um usuário com o próprio nome (minúsculo, sem acento), entra pelo próprio celular e deixa um arquivo na casa dele com nano. Depois tenta ler a casa do colega: ls /home/colega → Permission denied.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — PRÁTICA · AS GAVETAS DO ESCRITÓRIO · 25 MIN")
    add_para(doc, "Agora o escritório de verdade, em /srv/escritorio. De volta como equipeN (exit para sair da ana).")
    for c in PASTAS_CMD:
        add_cmd(doc, c, size=9.5)
    add_heading(doc, "Lendo o ls -l", level=2)
    add_cmd(doc, "drwxrwx--- 2 root financeiro 4096 dez  4 21:10 financeiro", size=9.5)
    add_bullet(doc, "d = é pasta. Depois, três trios: o que o DONO pode, o que o GRUPO pode, o que os OUTROS podem")
    add_bullet(doc, "r = ler, w = alterar (escrever), x = entrar na pasta (ou executar). - = não pode")
    add_bullet(doc, "root é o dono, financeiro é o grupo. rwx rwx --- = o dono e o grupo tudo, os outros nada")
    add_heading(doc, "Os números (Matemática)", level=2)
    add_table(doc, [["Número", "Letras", "Conta", "Quer dizer"]] + [list(n) for n in NUMEROS], col_widths=[2, 2.2, 2.5, 9.8])
    add_bullet(doc, "r vale 4, w vale 2, x vale 1. Soma em cada trio. 770 = dono 7, grupo 7, outros 0")
    add_bullet(doc, "755 no publico: todo mundo lê (é o mural), só o dono altera")
    add_para(doc, "Nunca use 777 para \"resolver\" um Permission denied: é deixar a gaveta aberta para qualquer um. O técnico descobre quem precisa de acesso e põe no grupo certo.", bold=True, color=COR_ALERTA)
    add_heading(doc, "O teste (página 3 da folha)", level=2)
    add_para(doc, "Entrar como ana (su - ana, ou SSH) e depois como carla. A equipe marca o que aconteceu. exit volta:")
    add_table(doc, [["Quem", "Tenta", "O certo é"]] + [list(t) for t in TESTE], col_widths=[2.5, 8.5, 5.5])

    add_hr(doc)

    # BLOCO 9
    add_heading(doc, "BLOCO 9 — FECHAMENTO + FICHA TÉCNICA · 10 MIN")
    add_bullet(doc, "A equipe completa a ficha técnica (página 3 da folha de hoje): endereço fixo, MAC, como entrar por SSH, usuários e grupos, pastas e quem acessa")
    add_bullet(doc, "As três frases (slide)")
    add_bullet(doc, "Encontro 3: a pasta /srv/escritorio aparece no Windows como uma pasta de rede, e a ana abre a gaveta do financeiro do computador dela. E a entrega do servidor: outra equipe vai testar tudo só com a ficha técnica")
    add_bullet(doc, "sudo poweroff. Pelo SSH também funciona: e aí a sessão cai, claro")

    add_hr(doc)

    add_heading(doc, "SE TRAVAR")
    add_table(doc, [["O que aparece", "O que fazer"]] + [list(r) for r in TRAVOU], col_widths=[5.5, 11])

    add_hr(doc)

    add_heading(doc, "SE PERGUNTAREM")
    add_para(doc, "\"Dá pra entrar por SSH de casa?\"", bold=True)
    add_para(doc, "> \"Tecnicamente dá, mas aí o servidor fica de frente pra internet inteira, e tem gente testando senhas o dia todo, automaticamente. Para isso se usa chave em vez de senha e VPN. É assunto para quando vocês forem administrar servidor de verdade.\"", italic=True)
    add_para(doc, "\"Qual a diferença entre o equipeN e o root?\"", bold=True)
    add_para(doc, "> \"O root é o administrador total. No Ubuntu ninguém entra como root: o equipeN pede o poder emprestado com sudo, um comando por vez, e fica registrado. É mais seguro.\"", italic=True)
    add_para(doc, "\"Se eu apagar um arquivo do sistema, estraga?\"", bold=True)
    add_para(doc, "> \"Pode estragar. Por isso o sistema não deixa usuário comum apagar nada fora da casa dele, e o sudo pede senha. Com sudo, o sistema obedece: confira antes de dar Enter.\"", italic=True)
    add_para(doc, "\"Por que o comando se chama mv se ele renomeia?\"", bold=True)
    add_para(doc, "> \"Porque renomear é mover para outro nome no mesmo lugar. Para o Linux é a mesma coisa.\"", italic=True)
    add_para(doc, "\"No Windows também tem usuário e permissão?\"", bold=True)
    add_para(doc, "> \"Tem: botão direito na pasta → Propriedades → Segurança. A ideia é a mesma: quem pode ler, quem pode alterar. No Encontro 3 os dois se encontram.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "NÃO apague usuários, grupos nem /srv/escritorio: o Encontro 3 usa tudo")
    add_bullet(doc, "NÃO mexa nas reservas dos roteadores: o endereço .10 é o que vai na ficha técnica")
    add_bullet(doc, "Tire foto das fichas técnicas atualizadas (agora com as senhas dos funcionários)")
    add_bullet(doc, "Anote quantas equipes passaram no teste das gavetas (Bloco 8) sem ajuda. É o indicador do encontro")

    add_hr(doc)

    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Dei continuidade ao estudo de servidores com Linux nas máquinas instaladas pelos estudantes no encontro anterior. Apresentei a estrutura de pastas do sistema e a diferença entre caminho completo e relativo. Em equipes, os estudantes navegaram pelo sistema e consultaram a ajuda dos comandos, criaram, leram, copiaram, renomearam e apagaram arquivos e pastas em uma estrutura de treino, conferindo o resultado com a representação em árvore. Fixaram o endereço do servidor por reserva no roteador a partir do endereço físico (MAC) e acessaram o servidor a distância por conexão segura (SSH) a partir do celular e do computador com Windows, discutindo a importância da senha forte. Criaram usuários e grupos para os funcionários de um escritório fictício, verificaram que usuários comuns não têm poder de administração e aplicaram permissões às pastas de cada setor, interpretando a notação em letras e em números. Testaram o acesso de cada usuário e atualizaram a ficha técnica do servidor.")

    add_hr(doc)

    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou pra usar em outro encontro:", italic=True, color=COR_CINZA)
    add_para(doc, "1. O arquivo escondido", bold=True)
    add_para(doc, "Pelo SSH, você entra no servidor de cada equipe e esconde um arquivo tesouro.txt em alguma pasta funda (sudo cp para /var/lib/… ou /usr/share/…), com uma frase dentro. A equipe acha com find / -name tesouro.txt 2>/dev/null e lê com cat. Ganha quem disser a frase e o caminho completo. 15 min.")
    add_para(doc, "2. Tradutor de permissões", bold=True)
    add_para(doc, "No quadro, 8 permissões em letras (rwxr-x---, rw-r--r--…) e 8 em números (640, 750…). As equipes traduzem de um lado para o outro e dizem quem pode o quê em cada uma. Matemática. 15 min.")
    add_para(doc, "3. O que digitamos hoje", bold=True)
    add_para(doc, "history mostra todos os comandos da equipe na noite, numerados. A equipe escolhe os 10 mais importantes e escreve, em linguagem de vó, o que cada um faz: vira o cartão de comandos da equipe para o Encontro 3. Linguagens. 20 min.")
    add_para(doc, "4. O funcionário que saiu", bold=True)
    add_para(doc, "O bruno saiu do escritório. A equipe tira ele do grupo (sudo gpasswd -d bruno financeiro), testa que ele não abre mais a gaveta e depois decide: apagar o usuário (sudo deluser bruno) ou só bloquear (sudo passwd -l bruno)? Discuta o que acontece com os arquivos dele. 15 min.")
    add_para(doc, "5. Guia de acesso para o cliente", bold=True)
    add_para(doc, "Cada equipe escreve uma folha para a ana: como entrar no servidor pelo celular, passo a passo, com o nome do aplicativo, o endereço e onde fica a gaveta dela. Outra equipe tenta seguir o guia num celular que nunca entrou no servidor. 25 min.")

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
                if run.text.startswith("$ ") or run.text[:1] in "<│├└" or run.text.startswith("treino") or run.text.startswith("drwx"):
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

    add_slide_cover(prs, "SERVIDORES COM LINUX · ENCONTRO 2", "PASTAS, USUÁRIOS", "E ACESSO DE LONGE",
                    "O servidor ganha funcionários. E vocês mandam nele pelo celular",
                    "Curso FIC · Montador e Reparador de Computadores · CEJA Itapiranga – SC")

    lista("RETOMADA", "Do Encontro 1 pra hoje", [
        "•  Liguem o servidor e entrem com equipeN. O que acontece com a senha?",
        "•  A página da equipe ainda abre no celular?",
        "•  A tarefa: um serviço que você usou e que depende de servidor",
        "•  E se esse servidor desligasse?",
    ], size=22)

    frase("A FRASE DA NOITE", "Cada um entra com o seu nome\ne só abre a gaveta\nque é sua.", size=36)

    tabela("O ARMÁRIO DE PASTAS", "Tudo começa na raiz  /", [["Pasta", "O que guarda"]] + [list(r) for r in RAIZ],
           col_widths=[2.8, 9.1], size=17, height=3.6,
           nota="No Linux não tem C: nem D:. Tudo é pasta dentro da raiz.")

    lista("O CAMINHO", "Endereço completo e endereço curto", [
        "$ /var/www/html",
        "   Completo: \"Rua X, 123, Itapiranga\". Funciona de qualquer lugar",
        "",
        "$ recepcao/recados.txt",
        "   Curto: \"a casa do lado\". Só funciona se você estiver no lugar certo",
        "",
        "•   ..  = a pasta de cima          ~  = a sua casa",
    ], size=21)

    tabela("PRÁTICA · 25 MIN", "O passeio pelas pastas", [["Comando", "Para quê"]] + [list(r) for r in PASSEIO],
           col_widths=[3.6, 8.3], size=15, height=4.0, top=2.3,
           nota="Ajuda: ls --help  ·  man ls (sai com q)")

    tabela("PRÁTICA · 30 MIN", "A pasta de treino do escritório", [["Comando", "Para quê"]] + [list(r) for r in PEDIDO],
           col_widths=[6.3, 5.6], size=12, height=4.2, top=2.2,
           nota="O digitador troca a cada 4 passos.")

    lista("O DESENHO", "tree ~/treino tem que mostrar isto", ARVORE, size=24,
          nota="Diferente? A equipe descobre o que faltou e conserta.")

    frase("CUIDADO", "No Linux não tem lixeira.\nO rm apaga de vez.\nConfira com pwd e ls antes.", cor=R, size=34)

    frase("", "INTERVALO\n15 minutos", cor=K)

    frase("POR QUE ENDEREÇO FIXO", "Se o roteador der outro número amanhã,\nninguém acha o servidor:\nnem o SSH, nem o site, nem a pasta.", size=30)

    lista("PRÁTICA · 15 MIN", "A reserva da Noite 8, agora pro servidor", [
        "1.  No servidor: ip a → linha link/ether = o MAC",
        "2.  Painel do roteador → DHCP → Reserva de endereço",
        "3.  MAC do servidor → 192.168.1N.10 → Salvar",
        "4.  No servidor: sudo reboot",
        "5.  ip a de novo: ficou .10? O site abre no .10?",
    ], size=22, nota="A senha do painel está na ficha de entrega da Noite 8.")

    lista("PRÁTICA · 35 MIN", "Entrar de longe: SSH", [
        "Celular (Wi-Fi da equipe):  ConnectBot (Android) · Termius (iPhone)",
        "$ equipeN@192.168.1N.10   porta 22",
        "",
        "Windows (cabo no roteador da equipe):  Iniciar → cmd",
        "$ ssh equipeN@192.168.1N.10",
        "",
        "Primeira vez: \"confia nesta máquina?\" → yes",
    ], size=20)

    frase("A PROVA", "$ wall Oi da equipe N\n\nO recado aparece na tela do servidor,\ndo outro lado da mesa.", size=30)

    lista("DE LONGE", "Mudar o site sem levantar da cadeira", [
        "$ sudo nano /var/www/html/index.html",
        "",
        "<p>Alterado de longe pela equipe N</p>",
        "",
        "•  Outro celular recarrega a página",
        "•  Cada um da equipe entra pelo próprio celular, ao mesmo tempo",
    ], size=21)

    frase("SEGURANÇA", "Agora qualquer um no Wi-Fi\npode TENTAR entrar.\nO que segura é a senha.", cor=R, size=34)

    tabela("PRÁTICA · 30 MIN", "Os funcionários do escritório", [["Usuário", "Nome", "Grupo (setor)"]] + [list(f) for f in FUNCIONARIOS],
           col_widths=[3.5, 3.5, 4.9], size=20, height=2.0, top=2.4)
    s = prs.slides[-1]
    tb = add_text(s, ["$ " + c for c, _ in USUARIOS_CMD] + ["", "adduser: senha 2 vezes, o resto Enter, Y no fim.  -aG: com o a!"],
                  0.7, 4.6, 11.9, 2.2, size=17, color=K)
    for para in tb.text_frame.paragraphs:
        for run in para.runs:
            if run.text.startswith("$ "):
                run.font.name = MONO

    lista("O TESTE", "Entrar como ana", [
        "$ ssh ana@192.168.1N.10",
        "$ whoami        → ana",
        "$ pwd           → /home/ana",
        "$ sudo apt update",
        "",
        "ana is not in the sudoers file",
        "Certo: funcionário usa, não administra.",
    ], size=21)

    lista("PRÁTICA · 25 MIN", "As gavetas do escritório", ["$ " + c for c in PASTAS_CMD], size=15,
          nota="O escritório mora em /srv/escritorio. No Encontro 3 ela vira pasta de rede.")

    lista("LENDO O ls -l", "Quem pode o quê", [
        "drwxrwx--- 2 root financeiro  financeiro",
        "",
        "•  d = pasta.  Depois três trios:  DONO · GRUPO · OUTROS",
        "•  r = ler     w = alterar     x = entrar     - = não pode",
        "•  rwx rwx --- = dono e grupo tudo, os outros nada",
    ], size=21)

    tabela("OS NÚMEROS", "r = 4   w = 2   x = 1", [["Número", "Letras", "Conta", "Quer dizer"]] + [list(n) for n in NUMEROS],
           col_widths=[1.8, 1.8, 2.0, 6.3], size=18, height=2.8,
           nota="770 = dono 7, grupo 7, outros 0.  Nunca 777.", nota_cor=R)

    tabela("O TESTE DAS GAVETAS", "su - ana  ·  su - carla  ·  exit volta", [["Quem", "Tenta", "O certo é"]] + [list(t) for t in TESTE],
           col_widths=[1.8, 6.4, 3.7], size=16, height=3.6)

    lista("O QUE FICA", "Três frases pra guardar", [
        "\"Cada um entra com o seu nome e só abre a gaveta que é sua.\"",
        "",
        "\"Servidor tem endereço fixo.\"",
        "",
        "\"rm não tem lixeira. 777 não é solução.\"",
    ], size=23)

    lista("NO ENCONTRO 3", "A entrega do servidor", [
        "•  A pasta /srv/escritorio aparece no Windows como pasta de rede",
        "•  A ana abre a gaveta do financeiro do computador dela",
        "•  Outra equipe testa o seu servidor só com a ficha técnica",
        "",
        "Tarefa: pergunte a alguém que trabalha em escritório ou comércio:",
        "onde ficam os arquivos? Todo mundo vê tudo?",
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

    # PÁGINA 1 — Pastas e treino
    cabecalho_documento(
        doc,
        "ANDANDO PELAS PASTAS",
        "Servidores com Linux · Encontro 2 · Curso FIC · CEJA Itapiranga · 1 de 3",
    )
    campo(doc, "Nome: _____________________________________________     Equipe: ______   Servidor: servidor-equipe____", space_after=6, size=10)

    add_heading(doc, "1  O passeio", level=2)
    tbl = add_table(doc, [["Comando", "Para quê", "O que apareceu"]] + [[c, d, ""] for c, d in PASSEIO],
                    col_widths=[3.8, 7.2, 6.4])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=9)
    campo(doc, "Na ajuda (ls --help), o que faz ls -l? _____________________   E ls -h? _____________________", size=10)
    campo(doc, "cd /pasta-que-nao-existe mostrou: _____________________________   Quer dizer: ____________________", size=10)

    add_heading(doc, "2  A pasta de treino do escritório", level=2)
    tbl = add_table(doc, [["Comando", "Para quê", "✓"]] + [[c, d, "☐"] for c, d in PEDIDO],
                    col_widths=[7.6, 8.8, 1])
    small_table(tbl, size=8.5)
    mono_col(tbl, 0, size=8.5)
    add_para(doc, "O tree tem que desenhar:", size=10, bold=True)
    for l in ARVORE:
        add_cmd(doc, l, size=9)
    campo(doc, "Saiu igual?  ☐ Sim, de primeira  ☐ Não: o que faltou e como a equipe consertou? ______________________", size=10)

    # PÁGINA 2 — Endereço fixo e SSH
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "ENDEREÇO FIXO E ACESSO DE LONGE",
        "Pastas, usuários e acesso de longe · Encontro 2 · 2 de 3",
    )

    add_heading(doc, "1  Endereço fixo do servidor (reserva no roteador)", level=2)
    campo(doc, "MAC do servidor (ip a → link/ether): ____ : ____ : ____ : ____ : ____ : ____", size=10)
    campo(doc, "Endereço antes da reserva: 192.168.____.____      Endereço reservado: 192.168.____.10", size=10)
    campo(doc, "Depois do sudo reboot, o ip a mostrou: 192.168.____.____     Ficou o reservado?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "O site abriu no endereço novo (http://192.168.____.10)?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "Por que servidor precisa de endereço fixo e celular não precisa?", bold=True, size=10)
    linha(doc)
    linha(doc)

    add_heading(doc, "2  Entrar de longe (SSH)", level=2)
    tbl = add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in SSH_CMD], col_widths=[6.4, 11])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=9)
    tbl = add_table(doc, [
        ["Teste", "Funcionou?", "Se deu erro, o que dizia"],
        ["Entrar pelo celular (aplicativo: ____________)", "☐ Sim  ☐ Não", ""],
        ["O recado do wall apareceu na tela do servidor", "☐ Sim  ☐ Não", ""],
        ["Quantos da equipe entraram ao mesmo tempo: ____", "", ""],
        ["Entrar pelo Windows (Prompt de Comando)", "☐ Sim  ☐ Não", ""],
        ["Mudar o site de longe e ver no outro celular", "☐ Sim  ☐ Não", ""],
    ], col_widths=[7.4, 3, 7])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.75)
    campo(doc, "Na primeira vez, o SSH perguntou se você confia na máquina. Por que ele pergunta isso?", bold=True, size=10)
    linha(doc)
    campo(doc, "Agora qualquer pessoa no Wi-Fi da equipe pode tentar entrar no servidor. O que protege ele?", bold=True, size=10)
    linha(doc)

    # PÁGINA 3 — Usuários, gavetas, ficha
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "OS FUNCIONÁRIOS E AS GAVETAS",
        "Pastas, usuários e acesso de longe · Encontro 2 · 3 de 3",
    )

    add_heading(doc, "1  Os funcionários", level=2)
    tbl = add_table(doc, [["Usuário", "Grupo", "Senha (a equipe escolhe, 10+)", "Criado", "No grupo (groups)"]] +
                    [[u, g, "", "☐", "☐"] for u, _, g in FUNCIONARIOS], col_widths=[2.4, 3, 6.4, 2, 3.6])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)
    campo(doc, "Como ana, sudo apt update respondeu: _______________________________   Isso está certo?  ☐ Sim  ☐ Não", size=10)

    add_heading(doc, "2  As gavetas: ls -l /srv/escritorio", level=2)
    tbl = add_table(doc, [["Pasta", "Permissão (letras)", "Número", "Grupo", "Quem pode entrar"],
                          ["financeiro", "", "", "", ""], ["recepcao", "", "", "", ""], ["publico", "", "", "", ""]],
                    col_widths=[2.8, 4, 2, 3, 5.6])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)
    add_para(doc, "r = 4 · w = 2 · x = 1 · soma em cada trio: dono, grupo, outros", italic=True, size=9, color=COR_CINZA)

    add_heading(doc, "3  O teste das gavetas", level=2)
    tbl = add_table(doc, [["Quem", "Tenta", "O que aconteceu", "Certo?"]] + [[q, t, "", "☐"] for q, t, _ in TESTE],
                    col_widths=[1.8, 7, 6.6, 2])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.65)
    campo(doc, "Por que não usar chmod 777 quando dá Permission denied?", bold=True, size=10)
    linha(doc)

    add_heading(doc, "4  Ficha técnica: o que entra hoje (passe para a folha do Encontro 1)", level=2)
    rows = [[FICHA_NOVA[i], "", FICHA_NOVA[i + 3], ""] for i in range(3)]
    tbl = add_table(doc, [["Item", "Como ficou", "Item", "Como ficou"]] + rows, col_widths=[3.8, 4.9, 3.8, 4.9])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.8)

    add_heading(doc, "TAREFA DA SEMANA", level=2, color=COR_DESTAQUE)
    campo(doc, "Pergunte a alguém que trabalha em escritório ou comércio: onde ficam os arquivos? ____________________", size=10)
    campo(doc, "Todo mundo vê tudo, ou cada um tem a sua parte? ________________________________________________", size=10)
    add_para(doc, "Traga esta folha e a do Encontro 1 no Encontro 3 · CEJA Itapiranga – SC", italic=True, size=9, color=COR_CINZA, align=WD_ALIGN_PARAGRAPH.CENTER)

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
        "PASTAS, USUÁRIOS E ACESSO DE LONGE",
        "Resumo do aluno · Servidores com Linux · Encontro 2 · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Um servidor de escritório tem pastas organizadas, um usuário para cada funcionário e regras de quem abre o quê. Ele tem endereço fixo e é administrado de longe, pelo celular ou pelo computador, com SSH.")

    add_heading(doc, "As pastas", level=2)
    add_table(doc, [["Pasta", "O que guarda"]] + [list(r) for r in RAIZ], col_widths=[3.5, 12])
    add_bullet(doc, "Caminho completo começa com / e funciona de qualquer lugar. Caminho curto parte de onde você está")
    add_bullet(doc, ".. é a pasta de cima. ~ é a sua casa")

    add_heading(doc, "Os comandos de arquivo", level=2)
    tbl = add_table(doc, [["Comando", "O que faz"],
                          ["pwd", "Mostra onde você está"],
                          ["ls  ·  ls -l  ·  ls -a", "Lista · com detalhes e permissões · com os escondidos"],
                          ["cd pasta  ·  cd ..  ·  cd ~", "Entra · sobe uma · volta pra casa"],
                          ["mkdir pasta", "Cria pasta"],
                          ["nano arquivo", "Cria ou edita (Ctrl+O, Enter salva · Ctrl+X sai)"],
                          ["cat arquivo", "Mostra o conteúdo"],
                          ["cp origem destino", "Copia"],
                          ["mv origem destino", "Move ou renomeia"],
                          ["rm arquivo  ·  rm -r pasta", "Apaga. NÃO TEM LIXEIRA"],
                          ["comando --help  ·  man comando", "Ajuda (no man, sai com q)"],
                          ], col_widths=[6, 9.5])
    mono_col(tbl, 0, size=9.5)

    add_heading(doc, "Endereço fixo e SSH", level=2)
    add_bullet(doc, "Servidor precisa de endereço fixo: reserva no roteador com o MAC do servidor (ip a → link/ether), depois sudo reboot")
    add_bullet(doc, "Entrar de longe: ssh usuario@endereco. No celular, ConnectBot (Android) ou Termius (iPhone). No Windows, no Prompt de Comando")
    add_bullet(doc, "Na primeira vez, ele pergunta se você confia na máquina: yes. Para sair: exit")
    add_bullet(doc, "Qualquer um na rede pode tentar entrar. O que protege é a senha longa")

    add_heading(doc, "Usuários, grupos e permissões", level=2)
    tbl = add_table(doc, [["Comando", "O que faz"]] + [list(r) for r in USUARIOS_CMD] +
                    [["sudo chgrp grupo pasta", "Muda o grupo dono da pasta"],
                     ["sudo chmod 770 pasta", "Muda as permissões"]], col_widths=[6.5, 9])
    mono_col(tbl, 0, size=9.5)
    add_para(doc, "Em ls -l: drwxrwx--- → d (pasta), depois três trios: DONO, GRUPO, OUTROS. r = ler (4), w = alterar (2), x = entrar (1). 770 = dono e grupo tudo, outros nada. 755 = todos leem, só o dono altera. Nunca 777.", size=10)
    add_bullet(doc, "Funcionário não é administrador: \"is not in the sudoers file\" é o certo")
    add_bullet(doc, "No usermod é -aG, com o a. Sem ele, a pessoa sai dos outros grupos")
    add_bullet(doc, "Pôs alguém num grupo e não funcionou? A pessoa precisa sair (exit) e entrar de novo")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Cada um entra com o seu nome e só abre a gaveta que é sua.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Servidor tem endereço fixo.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"rm não tem lixeira. 777 não é solução.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Qual a diferença entre cd /var e cd var?")
    add_bullet(doc, "Uma pasta está com rwxr-x---. Quanto dá em número? Quem pode alterar e quem só lê?")
    add_bullet(doc, "Por que é melhor cada funcionário ter o próprio usuário, em vez de todos usarem a mesma senha?")

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
