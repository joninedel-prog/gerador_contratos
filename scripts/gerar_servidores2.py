"""Gera os 4 arquivos do Encontro 2 do módulo de Servidores (Plano Mensal 4).

Servidores 2 — Pastas, usuários e acesso de longe
Continua das máquinas instaladas pelos alunos no Encontro 1 (usuário equipeN,
servidor servidor-equipeN, Nginx com a página da equipe, OpenSSH instalado).

Versão "bancada" (2026-11-17): cada equipe monta servidor + PC com Windows no
roteador da equipe (o roteador traz a internet da escola pela WAN). O PC tem o
PuTTY (instalado pelo professor) e abre a página de comandos publicada
(LINK_COMANDOS). A noite inteira é feita pelo PuTTY:
- endereço fixo do servidor (reserva 192.168.1N.10) antes do primeiro acesso;
- pastas e pasta de treino DIGITADAS (é ali que se aprende o comando);
- grupos e gavetas COLADOS da página, com pergunta na folha sobre cada linha;
- site da equipe gerado com o ChatGPT no PC e colado no nano pelo PuTTY.
Os usuários ana, bruno, carla e /srv/escritorio continuam: base do Samba no
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
LINK_COMANDOS = "https://claude.ai/artifact/7kZVEZKi9rJSuGs35weUj5"
ARQ_COMANDOS = f"Comandos - {TITULO_ARQ}.html"


# A bancada da equipe
BANCADA = [
    "Internet da escola ──cabo──► WAN do roteador da equipe",
    "                                   │ LAN",
    "                    ┌──────────────┴──────────────┐",
    "             Servidor (Ubuntu)              PC da equipe (Windows)",
    "             192.168.1N.10                  PuTTY + navegador",
]

# PuTTY
PUTTY = [
    ("Host Name", "192.168.1N.10   ·   Port 22   ·   SSH marcado"),
    ("Saved Sessions", "servidor-equipeN → Save → Open"),
    ("PuTTY Security Alert", "Primeira vez: a impressão digital do servidor → Accept"),
    ("login as:", "equipeN, depois a senha (não aparece)"),
    ("Entrou", "equipeN@servidor-equipeN:~$"),
]

TECLAS = [
    ("Copiar no PuTTY", "Só selecionar com o mouse"),
    ("Colar no PuTTY", "Botão direito do mouse"),
    ("Ctrl+V", "Não cola"),
    ("Ctrl+C", "Não copia: CANCELA o comando que está rodando"),
]

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

# Digitados: o adduser é interativo (pede senha), não vai em bloco colado
ADDUSER_CMD = [
    ("sudo adduser ana", "Senha 2 vezes (a equipe escolhe, 10+), o resto Enter, Y no fim"),
    ("sudo adduser bruno", "Igual"),
    ("sudo adduser carla", "Igual"),
]

# Colados da página
SETORES_CMD = [
    "sudo groupadd financeiro",
    "sudo groupadd recepcao",
    "sudo usermod -aG financeiro ana",
    "sudo usermod -aG financeiro bruno",
    "sudo usermod -aG recepcao carla",
    "groups ana bruno carla",
]

PASTAS_CMD = [
    "sudo mkdir -p /srv/escritorio/financeiro /srv/escritorio/recepcao /srv/escritorio/publico",
    "sudo chgrp financeiro /srv/escritorio/financeiro",
    "sudo chgrp recepcao /srv/escritorio/recepcao",
    "sudo chmod 770 /srv/escritorio/financeiro /srv/escritorio/recepcao",
    "sudo chmod 755 /srv/escritorio/publico",
    "ls -l /srv/escritorio",
]

# Na folha: "o que essa linha fez?"
LINHAS_EXPLICAR = [
    "sudo groupadd financeiro",
    "sudo usermod -aG financeiro ana",
    "groups ana bruno carla",
    "sudo mkdir -p /srv/escritorio/financeiro …",
    "sudo chgrp financeiro /srv/escritorio/financeiro",
    "sudo chmod 770 … financeiro … recepcao",
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

# O site com o ChatGPT
PEDIDO_IA = [
    "Crie o código de uma página de site em um único arquivo HTML.",
    "Regras: comece com <!DOCTYPE html> e coloque <meta charset=\"UTF-8\"> no <head>. Todo o CSS dentro de <style>. Sem JavaScript. Sem imagens de outros sites (use cores, emojis e formas com CSS). Fica boa no celular. Textos em português do Brasil.",
    "O site é de: [o negócio], em Itapiranga – SC. Tem: nome e slogan, serviços com preços inventados, horário, contato (telefone e endereço inventados).",
    "No rodapé: \"Site hospedado no servidor-equipeN · Curso FIC CEJA Itapiranga\". Responda só com o código.",
]

PUBLICAR_CMD = [
    "sudo cp /var/www/html/index.html ~/site-encontro1.html",
    "sudo rm /var/www/html/index.html",
    "sudo nano /var/www/html/index.html",
]

TRAVOU = [
    ("Network error: Connection timed out", "O PC não está no roteador da equipe (cabo na WAN em vez da LAN?) ou o Host Name está errado. ipconfig no cmd do PC: tem que ser 192.168.1N.x"),
    ("Network error: Connection refused", "O SSH não ficou instalado. Na tela do servidor: sudo apt install openssh-server"),
    ("Access denied", "Usuário ou senha errados. Tudo minúsculo. A senha não aparece"),
    ("POTENTIAL SECURITY BREACH", "Aquele endereço já foi de outra máquina (a reserva mudou). Se você sabe por quê, Accept. Se não sabe, pare e investigue: é para isso que o aviso existe"),
    ("O endereço não mudou depois da reserva", "Conferir o MAC digitado no roteador, salvar e sudo reboot de novo. Se insistir, ver nota do professor no Bloco 2"),
    ("Ctrl+V não fez nada / Ctrl+C parou o comando", "No PuTTY: selecionar copia, botão direito cola. Ctrl+C cancela"),
    ("Pediu senha no meio de um bloco colado", "Ctrl+C, digitar sudo -v e a senha, colar o bloco de novo"),
    ("No such file or directory", "Caminho errado. Use pwd para ver onde está e Tab para completar"),
    ("ana is not in the sudoers file", "Esperado: funcionário não é administrador. É o teste dando certo"),
    ("Permission denied numa pasta que devia abrir", "Usuário entrou antes de ser posto no grupo. exit e su - ana de novo"),
    ("Acentos estranhos no site (Ã©, Ã§)", "Faltou <meta charset=\"UTF-8\">. Pedir ao ChatGPT e publicar de novo"),
    ("403 Forbidden no navegador", "Não existe index.html: o nano fechou sem salvar. Rodar de novo o bloco Publicar a versão nova"),
    ("O site não mudou", "Ctrl+F5 no navegador. Conferir se salvou no nano (Ctrl+O, Enter)"),
    ("chatgpt.com não abre", "O PC está sem internet: o roteador da equipe está com a WAN ligada? Se a rede da escola bloquear, usar copilot.microsoft.com com o mesmo pedido"),
]

FICHA_NOVA = [
    "Endereço fixo do servidor (reserva)",
    "MAC do servidor",
    "Como entrar de longe (PuTTY)",
    "Usuários criados e grupos",
    "Pastas do escritório e quem acessa",
    "O site: de quê, testado por / data",
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
    add_para(doc, "No Encontro 1 a turma instalou o servidor e publicou a página. Hoje cada equipe monta a bancada do técnico: o servidor e um PC com Windows no mesmo roteador, e o roteador trazendo a internet da escola. A partir daí, a turma não usa mais o teclado do servidor: faz tudo do PC, pelo PuTTY (SSH), como o técnico faz com o servidor do cliente.")
    add_para(doc, "O PC tem internet, então a noite tem duas ferramentas novas: a página de comandos (para copiar e colar os blocos longos) e o ChatGPT, que escreve o site da equipe. A regra que segura a aprendizagem: o que é comando do dia a dia (pastas, arquivos, adduser) se DIGITA; o que é bloco longo de configuração se COLA, mas a equipe explica na folha o que cada linha fez.")
    add_para(doc, "O fio da noite continua sendo a demanda do escritório. Os usuários ana, bruno e carla, os grupos e a pasta /srv/escritorio são usados no Encontro 3, quando essa pasta vira pasta compartilhada no Windows. E o PC da bancada já é o computador da ana.")
    add_para(doc, "A frase da noite: \"Cada um entra com o seu nome e só abre a gaveta que é sua.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "O que se resolve neste encontro", level=2)
    add_para(doc, "Cada aluno sai sabendo montar a rede de bancada (servidor e PC no mesmo roteador), fixar o endereço do servidor com reserva, entrar no servidor de longe pelo PuTTY, copiar e colar com segurança no terminal, andar entre as pastas e mexer em arquivos (pwd, ls, cd, mkdir, nano, cat, cp, mv, rm), criar usuários e grupos, ler e aplicar permissões, e publicar um site feito com IA, conferindo o que a IA entregou.")

    add_heading(doc, "Onde este encontro fica no módulo", level=2)
    add_para(doc, "Usa as máquinas do Encontro 1 sem reinstalar. Retoma a WAN e a LAN do roteador (Noites 2 e 8), a reserva por MAC (Noite 8), o gateway e o ping (Noite 6), os cabos testados (Noite 5) e a senha forte (Noite 10). A pasta /srv/escritorio e os usuários são a base do Encontro 3 (Samba e entrega ao cliente).")
    add_para(doc, "Não entre em: chave SSH no lugar de senha, editar o sudoers, permissões especiais (setgid, sticky), ACL, netplan e IP fixo configurado no próprio servidor, redirecionamento de portas, HTML e CSS linha por linha (o site é da IA; a turma confere e ajusta pedindo). O endereço fixo de hoje é pela reserva no roteador.", italic=True, color=COR_ALERTA)

    add_hr(doc)

    add_heading(doc, "ANTES DA AULA")
    add_heading(doc, "A bancada de cada equipe", level=2)
    for l in BANCADA:
        add_cmd(doc, l, size=9)
    add_bullet(doc, "Cada equipe volta para a MESMA máquina e o MESMO roteador do Encontro 1. Ao lado, um PC do laboratório com Windows. Dois cabos testados (Noite 5): servidor → LAN 1, PC → LAN 2")
    add_bullet(doc, "A porta WAN (azul) de cada roteador ligada na rede da escola, como na Noite 8. Confira num PC: ligado só no roteador da equipe, ele abre chatgpt.com? Então o roteador está passando a internet")
    add_bullet(doc, "Os PCs largam a rede da escola e passam a receber endereço do roteador da equipe (192.168.1N.1xx). Se algum PC tiver IP fixo configurado pela escola, troque para automático (DHCP) ou use outro PC")
    add_bullet(doc, "Plano de endereço fixo: o servidor da equipe N fica em 192.168.1N.10 (fora da faixa .100 a .150 da Noite 8, como foi a reserva da \"impressora\" em .50)")
    add_bullet(doc, "Separe as fichas de entrega da Noite 8: a reserva do Bloco 2 precisa da senha de administrador de cada roteador. Separe também as fotos das fichas técnicas do Encontro 1 (senha do equipeN)")

    add_heading(doc, "Em cada PC", level=2)
    add_bullet(doc, "Instale o PuTTY (putty.org → Download PuTTY → instalador 64-bit x86 .msi). No Windows, Iniciar → PuTTY")
    add_bullet(doc, "Abra a página de comandos e deixe nos favoritos: " + LINK_COMANDOS + ". A página é PRIVADA até você mudar: no menu Compartilhar, libere para quem tiver o link. Plano B: o arquivo \"" + ARQ_COMANDOS + "\" (na pasta do curso) num pendrive, aberto direto no navegador")
    add_bullet(doc, "Abra chatgpt.com e confira se abre sem login (o ChatGPT gratuito funciona sem conta). Se a rede bloquear, o plano B é copilot.microsoft.com com o mesmo pedido")
    add_bullet(doc, "Ensaio de 10 minutos numa bancada: entre pelo PuTTY, cole o bloco \"Os setores\" da página e veja como ele se comporta. No Ubuntu 24.04 o bloco colado costuma aparecer inteiro na linha e esperar o Enter. Saber isso antes evita susto na turma")
    add_bullet(doc, "No ensaio, peça ao ChatGPT o site, cole no nano pelo PuTTY e veja se os acentos aparecem certos no navegador. Se vierem estranhos: no PuTTY, Window → Translation → UTF-8")

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "Folha do aluno (4 páginas), uma por pessoa")
    add_bullet(doc, "Resumo do aluno, uma por pessoa")
    add_bullet(doc, "As folhas do Encontro 1 voltam: a ficha técnica é completada hoje")

    add_hr(doc)

    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Retomada + montar a bancada", "20 min", "Servidor e PC no roteador da equipe, a internet chega no PC, a página ainda abre?"],
        ["2. PRÁTICA — Endereço fixo + entrar pelo PuTTY", "25 min", "MAC, reserva .10, reboot, PuTTY, wall. Copiar e colar no PuTTY"],
        ["3. PRÁTICA — As pastas (digitando)", "40 min", "O armário (5 min de fala), o passeio, a pasta de treino, tree"],
        ["INTERVALO", "15 min", ""],
        ["4. PRÁTICA — Funcionários e gavetas (colando)", "40 min", "adduser digitado, setores e gavetas colados, o que cada linha fez, o teste de quem abre o quê"],
        ["5. PRÁTICA — O site da equipe com o ChatGPT", "45 min", "O pedido, publicar no nano pelo PuTTY, conferir o que a IA inventou, um ajuste"],
        ["6. Visita aos sites + ficha técnica", "25 min", "Celular no Wi-Fi das outras equipes, ficha técnica, três frases, Encontro 3"],
    ], col_widths=[6.8, 2, 7.7])
    add_para(doc, "Blocos 2 a 5 e a visita: cerca de 160 min de atividade do aluno (76% do total). Exposição: a retomada, os 5 min do armário de pastas e o fechamento. Rodízio com duas telas: um no PC (digitador), um lendo a folha e conferindo antes do Enter (leitor), um na folha (anotador). Troca a cada bloco.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — RETOMADA + MONTAR A BANCADA · 20 MIN")
    add_bullet(doc, "A tarefa: cada um diz um serviço que usou na semana e que depende de servidor. Anote no quadro. \"O que aconteceria se esse servidor desligasse?\"")
    add_bullet(doc, "Desenhe a bancada no quadro (está na página 1 da folha). \"Onde entra a internet no roteador? E onde ligam o servidor e o PC?\" (WAN e LAN, Noite 2)")
    add_bullet(doc, "As equipes montam: servidor e PC por cabo nas portas LAN. Liga o servidor, entra com equipeN na tela dele")
    add_bullet(doc, "No PC: Iniciar → cmd → ipconfig. O endereço tem que ser 192.168.1N.algo: o PC está na rede da equipe. Abre chatgpt.com: a internet chega pelo roteador")
    add_bullet(doc, "No navegador do PC: http:// e o endereço antigo do servidor (o da ficha técnica). A página do Encontro 1 ainda abre?")
    add_para(doc, "\"Hoje o teclado do servidor vai ficar parado. Vocês vão mandar nele do PC, como o técnico faz com o servidor do cliente, que fica trancado numa sala. E no fim o servidor vai ter um site de verdade, feito com ajuda de uma IA.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — PRÁTICA · ENDEREÇO FIXO + ENTRAR PELO PUTTY · 25 MIN")
    add_para(doc, "\"Para entrar de longe, o PuTTY precisa do endereço do servidor. Se o roteador der outro número amanhã, ninguém acha o servidor: nem o PuTTY, nem o site, nem a pasta do Encontro 3. Servidor tem endereço fixo. Vocês já sabem fazer: a reserva da Noite 8.\"", italic=True, color=COR_PRIMARIA)
    add_heading(doc, "Endereço fixo — 10 min", level=2)
    add_bullet(doc, "Na tela do servidor: ip a. Na placa que não é lo, a linha link/ether mostra o MAC (ex.: 00:1a:2b:3c:4d:5e). Anotar com cuidado")
    add_bullet(doc, "No navegador do PC: painel do roteador (192.168.1N.1, senha da ficha de entrega) → DHCP → Reserva de endereço → MAC do servidor → 192.168.1N.10. Salvar")
    add_bullet(doc, "No servidor: sudo reboot. Depois de ligar, ip a: tem que ser .10. No navegador do PC: http://192.168.1N.10 abre a página")
    add_para(doc, "Nota do professor: se a reserva não pegar mesmo com o MAC certo, o servidor está se identificando pro roteador por outro código (client-id). Conserto, feito por você: em /etc/netplan/50-cloud-init.yaml, abaixo de dhcp4: true, acrescente a linha dhcp-identifier: mac (mesmo recuo), depois sudo netplan apply. Não mostre para a turma.", italic=True, color=COR_CINZA)
    add_heading(doc, "O PuTTY — 15 min", level=2)
    add_para(doc, "SSH é uma conexão segura (tudo embaralhado no caminho) até o terminal do servidor. O PuTTY é o programa que faz isso no Windows. Página 1 da folha e passo 1 da página de comandos.")
    add_table(doc, [["Onde", "O que fazer"]] + [list(r) for r in PUTTY], col_widths=[4.5, 12])
    add_bullet(doc, "O Security Alert só aparece na primeira vez. Explique: \"o PuTTY guardou a impressão digital do servidor. Se um dia outra máquina se passar por ele, o aviso volta, e aí o técnico desconfia\"")
    add_bullet(doc, "A prova: pelo PuTTY, digitar wall Oi da equipe N. O recado aparece na tela do servidor, do lado do PC")
    add_bullet(doc, "Desligue o monitor do servidor de propósito (ou vire para a parede). \"De agora em diante, a tela do servidor é o PuTTY.\"")
    add_para(doc, "Pare a turma nas teclas, antes de qualquer cópia:", bold=True)
    add_table(doc, [["Tecla", "No PuTTY"]] + [list(r) for r in TECLAS], col_widths=[4.5, 12])
    add_para(doc, "Fale da segurança: \"Agora qualquer um ligado no roteador da equipe pode TENTAR entrar no servidor. O que segura é a senha. Por isso ela é longa e não fica colada no monitor.\" Liga com a Noite 10.", bold=True, color=COR_ALERTA)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — PRÁTICA · AS PASTAS (DIGITANDO) · 40 MIN")
    add_para(doc, "Diga antes: \"Isto aqui não se cola. É o dia a dia do técnico, tem que estar na mão.\" Na página de comandos esses passos estão marcados como DIGITE e não têm botão de copiar.", bold=True, color=COR_PRIMARIA)
    add_heading(doc, "O armário de pastas — 5 min de fala", level=2)
    add_para(doc, "Desenhe no quadro uma árvore de cabeça para baixo: a raiz / em cima, as pastas principais embaixo.")
    add_table(doc, [["Pasta", "O que guarda"]] + [list(r) for r in RAIZ], col_widths=[3.5, 13])
    add_bullet(doc, "No Linux não tem C: nem D:. Tudo começa na raiz /")
    add_bullet(doc, "Caminho completo: /var/www/html (\"Rua X, 123, Itapiranga\"), funciona de qualquer lugar. Caminho curto: recepcao/recados.txt (\"a casa do lado\"), só funciona se você estiver no lugar certo. .. é a pasta de cima, ~ é a sua casa")
    add_heading(doc, "O passeio — 15 min", level=2)
    add_para(doc, "Página 2 da folha. O anotador escreve o que apareceu em cada passo.")
    add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in PASSEIO], col_widths=[5, 11.5])
    add_bullet(doc, "Em cat index.html, a equipe vê a página dela como texto: \"o navegador lê exatamente isso e desenha\". Daqui a pouco esse arquivo vai ser trocado pelo site da IA")
    add_bullet(doc, "Ajuda: ls --help e man ls (sai com q). Peça que achem o que faz ls -l e ls -h")
    add_bullet(doc, "A tecla Tab completa nomes de pastas: cd /va + Tab. Mostre. Evita metade dos erros")
    add_heading(doc, "A pasta de treino — 20 min", level=2)
    add_para(doc, "Antes de mexer no escritório de verdade, a equipe treina numa pasta que pode estragar. O digitador troca a cada 4 passos.")
    add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in PEDIDO], col_widths=[7.5, 9])
    add_para(doc, "No fim, tree ~/treino tem que desenhar exatamente isto (também está na folha):")
    for l in ARVORE:
        add_cmd(doc, l, size=10)
    add_bullet(doc, "Pare a turma no rm: \"No Linux não tem lixeira. O rm apaga de vez. Antes de apagar, o técnico confere com ls e pwd onde está.\"")
    add_bullet(doc, "Desenho diferente? A equipe descobre sozinha o que faltou e conserta. É o laudo da atividade")
    add_para(doc, "Onde dilatar: o leitor fala só o destino (\"vai pra pasta do site\", \"volta pra casa\") e o digitador acha o comando sozinho. Ou a equipe desenha uma estrutura nova no papel, troca com a do lado, e cada uma constrói o desenho da outra. Confere com tree.", italic=True, color=COR_DESTAQUE)

    add_heading(doc, "INTERVALO — 15 MIN", level=2)

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · FUNCIONÁRIOS E GAVETAS (COLANDO) · 40 MIN")
    add_para(doc, "Página 3 da folha. O escritório tem três funcionários:")
    add_table(doc, [["Usuário", "Nome", "Grupo (setor)"]] + [list(f) for f in FUNCIONARIOS], col_widths=[4, 5, 7.5])
    add_heading(doc, "Os usuários — digitados (10 min)", level=2)
    add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in ADDUSER_CMD], col_widths=[5, 11.5])
    add_bullet(doc, "Por que digitado: o adduser faz perguntas (senha, nome, telefone). Comando que pergunta não vai em bloco colado")
    add_bullet(doc, "Senha: a equipe escolhe uma para cada funcionário (10+ caracteres) e anota na folha. Nome e telefones: só Enter. No fim, Y")
    add_heading(doc, "Os setores e as gavetas — colados (15 min)", level=2)
    add_para(doc, "Na página de comandos, passos 3 e 4: botão Copiar → no PuTTY, botão direito. O bloco aparece na tela: o leitor lê linha por linha em voz alta ANTES do Enter.")
    add_para(doc, "Os setores:", bold=True)
    for c in SETORES_CMD:
        add_cmd(doc, c, size=9.5)
    add_para(doc, "As gavetas:", bold=True)
    for c in PASTAS_CMD:
        add_cmd(doc, c, size=9.5)
    add_para(doc, "\"Colar é rápido, mas o servidor é seu. Na casa do cliente, o técnico não cola nada que não saiba explicar. Comando de internet colado sem ler já apagou muito servidor.\" A folha tem uma linha para a equipe explicar cada comando colado: sem a explicação, o bloco não vale.", bold=True, color=COR_ALERTA)
    add_heading(doc, "Lendo o ls -l", level=2)
    add_cmd(doc, "drwxrwx--- 2 root financeiro 4096 dez  4 21:10 financeiro", size=9.5)
    add_bullet(doc, "d = é pasta. Depois três trios: o que o DONO pode, o que o GRUPO pode, o que os OUTROS podem")
    add_bullet(doc, "r = ler, w = alterar, x = entrar na pasta. - = não pode. root é o dono, financeiro é o grupo")
    add_table(doc, [["Número", "Letras", "Conta", "Quer dizer"]] + [list(n) for n in NUMEROS], col_widths=[2, 2.2, 2.5, 9.8])
    add_bullet(doc, "r vale 4, w vale 2, x vale 1. 770 = dono 7, grupo 7, outros 0. 755 no publico: todo mundo lê (é o mural), só o dono altera. Matemática")
    add_para(doc, "Nunca use 777 para \"resolver\" um Permission denied: é deixar a gaveta aberta para qualquer um.", bold=True, color=COR_ALERTA)
    add_heading(doc, "O teste — digitado (15 min)", level=2)
    add_para(doc, "No PuTTY: su - ana (senha da ana). whoami → ana. pwd → /home/ana. Depois as gavetas. exit volta para equipeN. Repetir com su - carla.")
    add_table(doc, [["Quem", "Tenta", "O certo é"]] + [list(t) for t in TESTE], col_widths=[2.5, 8.5, 5.5])
    add_bullet(doc, "Ainda como ana: sudo apt update → ana is not in the sudoers file. \"A ana não está na lista de administradores.\" É o certo: funcionário usa, não administra")
    add_para(doc, "\"Por que não deixar todo mundo entrar com equipeN, que é mais fácil?\" Deixe a turma responder: senha compartilhada vaza, não dá pra saber quem fez o quê, e se a ana sair do escritório, troca a senha de todo mundo.", italic=True, color=COR_PRIMARIA)
    add_para(doc, "Onde dilatar: abrir uma SEGUNDA janela do PuTTY e entrar direto como ana (login as: ana). Duas pessoas no servidor ao mesmo tempo: who mostra as duas.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — PRÁTICA · O SITE DA EQUIPE COM O CHATGPT · 45 MIN")
    add_para(doc, "Página 4 da folha e passo 5 da página de comandos. A equipe escolhe um negócio fictício (a página já vem com \"assistência técnica de celulares e computadores\", mas pode ser a oficina, a padaria, o salão, a loja de peças). Na página, o pedido se monta sozinho com o negócio e o número da equipe.")
    add_heading(doc, "O pedido — 10 min", level=2)
    for l in PEDIDO_IA:
        add_cmd(doc, l, size=9)
    add_bullet(doc, "Leia o pedido com a turma antes de mandar: \"cada regra está ali por um motivo\". Um arquivo só: é o que cabe no nano. charset UTF-8: os acentos. Sem imagem de fora: se o outro site sair do ar, a imagem some. Telefone inventado: o site é de treino")
    add_para(doc, "\"Nunca cole senha, nome completo ou telefone de verdade no ChatGPT. O que vai para a IA sai da sala.\" Vale para o técnico na casa do cliente também: não se cola configuração do cliente com senha dentro.", bold=True, color=COR_ALERTA)
    add_heading(doc, "Publicar — 15 min", level=2)
    for c in PUBLICAR_CMD:
        add_cmd(doc, c, size=9.5)
    add_bullet(doc, "O bloco guarda a página do Encontro 1 (cp, que eles acabaram de aprender), apaga a antiga e abre o nano vazio")
    add_bullet(doc, "No ChatGPT, Copiar código. No nano, botão direito: o código entra todo. Ctrl+O, Enter, Ctrl+X")
    add_bullet(doc, "No navegador do PC: http://192.168.1N.10. Um celular no Wi-Fi da equipe também abre")
    add_heading(doc, "Conferir o que a IA fez — 10 min", level=2)
    add_bullet(doc, "Cada equipe acha pelo menos UMA coisa que a IA inventou, errou ou que não serve (preço absurdo, rua que não existe em Itapiranga, texto em inglês, serviço que o negócio não faz). Anota na folha")
    add_para(doc, "\"A IA escreve rápido, mas não assina. Quem entrega o site ao cliente e responde por ele é o técnico.\"", italic=True, color=COR_PRIMARIA)
    add_heading(doc, "Um ajuste — 10 min", level=2)
    add_bullet(doc, "A equipe pede uma mudança na mesma conversa (cores, promoção da semana, perguntas frequentes) e termina com \"me mande o código completo de novo\"")
    add_bullet(doc, "Publica pelo bloco \"Publicar a versão nova\" (rm + nano) e recarrega com Ctrl+F5")
    add_bullet(doc, "Pergunta para a folha: \"o site é feito pela IA. Quem está servindo ele para o navegador?\" (o Nginx, no servidor da equipe. A IA não está no caminho)")
    add_para(doc, "Onde dilatar: ls -l /var/www/html (quem é o dono do index.html?). cat ~/site-encontro1.html (a página antiga está guardada). Pedir à IA uma segunda página, contato.html, publicar com sudo nano /var/www/html/contato.html e ligar as duas com um link.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — VISITA AOS SITES + FICHA TÉCNICA · 25 MIN")
    add_bullet(doc, "Visita (15 min): como no Encontro 1, cada um conecta o celular no Wi-Fi de outra equipe (dados móveis desligados) e abre http://192.168.1M.10. Anota na folha de quê é o site e uma coisa boa")
    add_bullet(doc, "\"Por que o PC da equipe 1 não abre o site da equipe 2?\" (cada roteador é uma rede separada. Do lado de fora, só se vê o roteador. Liga com a Noite 6)")
    add_bullet(doc, "Ficha técnica (5 min): endereço fixo, MAC, como entrar pelo PuTTY, usuários e grupos, pastas e quem acessa, o site")
    add_bullet(doc, "As três frases e o Encontro 3: \"a pasta /srv/escritorio vai aparecer neste PC como pasta de rede, e a ana abre a gaveta do financeiro dele\". Desligar: sudo poweroff pelo PuTTY (a janela avisa que a conexão caiu)")

    add_hr(doc)

    add_heading(doc, "SE TRAVAR")
    add_table(doc, [["O que aparece", "O que fazer"]] + [list(r) for r in TRAVOU], col_widths=[5.5, 11])

    add_hr(doc)

    add_heading(doc, "SE PERGUNTAREM")
    add_para(doc, "\"Dá pra entrar por SSH de casa?\"", bold=True)
    add_para(doc, "> \"Tecnicamente dá, mas aí o servidor fica de frente pra internet inteira, e tem gente testando senhas o dia todo, automaticamente. Para isso se usa chave em vez de senha e VPN. É assunto para quando vocês forem administrar servidor de verdade.\"", italic=True)
    add_para(doc, "\"Dá pra pôr esse site na internet de verdade?\"", bold=True)
    add_para(doc, "> \"Dá: precisa de um domínio (o nome, tipo .com.br, que se paga por ano) e de um servidor que fique ligado com internet o tempo todo, que normalmente se aluga. O arquivo é o mesmo que vocês fizeram hoje.\"", italic=True)
    add_para(doc, "\"Se a IA faz o site, pra que aprender?\"", bold=True)
    add_para(doc, "> \"A IA escreveu a página. Quem montou a rede, fixou o endereço, instalou o Nginx, publicou e achou o erro da IA foram vocês. É isso que o cliente paga.\"", italic=True)
    add_para(doc, "\"Qual a diferença entre o equipeN e o root?\"", bold=True)
    add_para(doc, "> \"O root é o administrador total. No Ubuntu ninguém entra como root: o equipeN pede o poder emprestado com sudo, um comando por vez, e fica registrado. É mais seguro.\"", italic=True)
    add_para(doc, "\"Por que o comando se chama mv se ele renomeia?\"", bold=True)
    add_para(doc, "> \"Porque renomear é mover para outro nome no mesmo lugar. Para o Linux é a mesma coisa.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "NÃO apague usuários, grupos nem /srv/escritorio: o Encontro 3 usa tudo")
    add_bullet(doc, "NÃO mexa nas reservas dos roteadores nem desmonte as bancadas se der: no Encontro 3 o PC vira o computador da ana")
    add_bullet(doc, "Se os PCs precisarem voltar para a rede da escola, anote em qual porta de qual roteador cada um estava")
    add_bullet(doc, "Tire foto das fichas técnicas atualizadas (agora com as senhas dos funcionários)")
    add_bullet(doc, "Anote quantas equipes passaram no teste das gavetas sem ajuda e quantas explicaram todas as linhas coladas. São os indicadores do encontro")

    add_hr(doc)

    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Dei continuidade ao estudo de servidores com Linux nas máquinas instaladas pelos estudantes no encontro anterior. Em equipes, os estudantes montaram uma rede de bancada com o servidor e um computador com Windows ligados ao roteador da equipe, fixaram o endereço do servidor por reserva no roteador a partir do endereço físico (MAC) e passaram a administrá-lo a distância por conexão segura (SSH) com o programa PuTTY, discutindo a importância da senha forte. Apresentei a estrutura de pastas do sistema e a diferença entre caminho completo e relativo; os estudantes navegaram pelo sistema, consultaram a ajuda dos comandos e criaram, copiaram, renomearam e apagaram arquivos e pastas, conferindo o resultado com a representação em árvore. Criaram usuários e grupos para os funcionários de um escritório fictício e aplicaram permissões às pastas de cada setor a partir de blocos de comandos copiados, explicando por escrito a função de cada linha e interpretando a notação em letras e em números. Utilizaram uma ferramenta de inteligência artificial para gerar o site de um negócio fictício, publicaram-no no servidor, identificaram erros e informações inventadas pela ferramenta e discutiram o cuidado com dados pessoais. Visitaram os sites das outras equipes e atualizaram a ficha técnica do servidor.")

    add_hr(doc)

    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou pra usar em outro encontro:", italic=True, color=COR_CINZA)
    add_para(doc, "1. O arquivo escondido", bold=True)
    add_para(doc, "Pelo PuTTY, você entra no servidor de cada equipe e esconde um arquivo tesouro.txt numa pasta funda (/var/lib/… ou /usr/share/…), com uma frase dentro. A equipe acha com find / -name tesouro.txt 2>/dev/null e lê com cat. Ganha quem disser a frase e o caminho completo. 15 min.")
    add_para(doc, "2. Tradutor de permissões", bold=True)
    add_para(doc, "No quadro, 8 permissões em letras (rwxr-x---, rw-r--r--…) e 8 em números (640, 750…). As equipes traduzem de um lado para o outro e dizem quem pode o quê. Matemática. 15 min.")
    add_para(doc, "3. O que digitamos hoje", bold=True)
    add_para(doc, "history mostra todos os comandos da noite, numerados. A equipe escolhe os 10 mais importantes e escreve, em linguagem de vó, o que cada um faz: vira o cartão de comandos da equipe para o Encontro 3. Linguagens. 20 min.")
    add_para(doc, "4. O funcionário que saiu", bold=True)
    add_para(doc, "O bruno saiu do escritório. A equipe tira ele do grupo (sudo gpasswd -d bruno financeiro), testa que ele não abre mais a gaveta e decide: apagar (sudo deluser bruno) ou só bloquear (sudo passwd -l bruno)? E os arquivos dele? 15 min.")
    add_para(doc, "5. Pergunte à IA, confira no servidor", bold=True)
    add_para(doc, "A equipe pergunta ao ChatGPT \"como vejo quanto espaço livre tem no disco do Ubuntu?\" e \"como vejo quem está logado?\". Antes de rodar a resposta no PuTTY, explica o que o comando vai fazer. Depois roda e confere se a IA acertou. Mostra o hábito certo: ler antes de colar. 15 min.")

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
                t = run.text
                if t.startswith("$ ") or t[:1] in "<│├└┌ " or t.startswith("treino") or t.startswith("drwx") or t.startswith("Internet da escola"):
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
                    "A bancada do técnico: o servidor, o PC, o PuTTY. E um site feito com IA",
                    "Curso FIC · Montador e Reparador de Computadores · CEJA Itapiranga – SC")

    lista("RETOMADA", "Do Encontro 1 pra hoje", [
        "•  Um serviço que você usou na semana e que depende de servidor",
        "•  E se esse servidor desligasse?",
        "•  A página da equipe ainda abre?",
    ], size=24)

    frase("A FRASE DA NOITE", "Cada um entra com o seu nome\ne só abre a gaveta\nque é sua.", size=36)

    lista("PRÁTICA · MONTAR", "A bancada da equipe", BANCADA + [
        "",
        "•  PC: Iniciar → cmd → ipconfig → tem que ser 192.168.1N.algo",
        "•  PC: abre chatgpt.com? A internet chega pelo roteador",
    ], size=17)

    frase("POR QUE ENDEREÇO FIXO", "O PuTTY precisa achar o servidor.\nSe o roteador der outro número amanhã,\nninguém acha: nem o PuTTY, nem o site.", size=30)

    lista("PRÁTICA · 10 MIN", "A reserva da Noite 8, agora pro servidor", [
        "1.  No servidor: ip a → linha link/ether = o MAC",
        "2.  Navegador do PC: 192.168.1N.1 → DHCP → Reserva de endereço",
        "3.  MAC do servidor → 192.168.1N.10 → Salvar",
        "4.  No servidor: sudo reboot",
        "5.  ip a de novo: ficou .10? O site abre no .10?",
    ], size=22, nota="A senha do painel está na ficha de entrega da Noite 8.")

    tabela("PRÁTICA · 15 MIN", "Entrar de longe: PuTTY", [["Onde", "O que fazer"]] + [list(r) for r in PUTTY],
           col_widths=[3.6, 8.3], size=17, height=3.4,
           nota="SSH: uma conexão segura até o terminal do servidor.")

    tabela("ANTES DE COPIAR", "As teclas mudam no PuTTY", [["Tecla", "No PuTTY"]] + [list(r) for r in TECLAS],
           col_widths=[3.6, 8.3], size=20, height=2.8,
           nota="Ctrl+C no PuTTY = parar. Não é copiar.", nota_cor=R)

    frase("A PROVA", "$ wall Oi da equipe N\n\nO recado aparece na tela do servidor.\nDe agora em diante, a tela é o PuTTY.", size=30)

    frase("SEGURANÇA", "Agora qualquer um no roteador\npode TENTAR entrar.\nO que segura é a senha.", cor=R, size=34)

    frase("A REGRA DA NOITE", "DIGITE o que é do dia a dia.\nCOLE o que é bloco longo,\nmas só depois de saber explicar.", size=32)

    tabela("O ARMÁRIO DE PASTAS", "Tudo começa na raiz  /", [["Pasta", "O que guarda"]] + [list(r) for r in RAIZ],
           col_widths=[2.8, 9.1], size=17, height=3.6,
           nota="Caminho completo: /var/www/html · curto: recepcao/recados.txt · .. sobe · ~ casa")

    tabela("PRÁTICA · DIGITE", "O passeio pelas pastas", [["Comando", "Para quê"]] + [list(r) for r in PASSEIO],
           col_widths=[3.6, 8.3], size=15, height=4.0, top=2.3,
           nota="Ajuda: ls --help · man ls (sai com q) · Tab completa")

    tabela("PRÁTICA · DIGITE", "A pasta de treino do escritório", [["Comando", "Para quê"]] + [list(r) for r in PEDIDO],
           col_widths=[6.3, 5.6], size=12, height=4.2, top=2.2,
           nota="O digitador troca a cada 4 passos.")

    lista("O DESENHO", "tree ~/treino tem que mostrar isto", ARVORE, size=24,
          nota="Diferente? A equipe descobre o que faltou e conserta.")

    frase("CUIDADO", "No Linux não tem lixeira.\nO rm apaga de vez.\nConfira com pwd e ls antes.", cor=R, size=34)

    frase("", "INTERVALO\n15 minutos", cor=K)

    tabela("PRÁTICA · DIGITE", "Os funcionários do escritório", [["Usuário", "Nome", "Grupo (setor)"]] + [list(f) for f in FUNCIONARIOS],
           col_widths=[3.5, 3.5, 4.9], size=20, height=2.0, top=2.4)
    s = prs.slides[-1]
    tb = add_text(s, ["$ " + c for c, _ in ADDUSER_CMD] + ["", "Senha 2 vezes, o resto Enter, Y no fim. Comando que pergunta não se cola."],
                  0.7, 4.6, 11.9, 2.2, size=17, color=K)
    mono(tb)

    lista("PRÁTICA · COLE", "Os setores", ["$ " + c for c in SETORES_CMD], size=19,
          nota="O leitor lê cada linha em voz alta antes do Enter.")

    lista("PRÁTICA · COLE", "As gavetas do escritório", ["$ " + c for c in PASTAS_CMD], size=15,
          nota="Na folha: o que cada linha fez. Sem explicação, o bloco não vale.")

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

    tabela("O TESTE · DIGITE", "su - ana  ·  su - carla  ·  exit volta", [["Quem", "Tenta", "O certo é"]] + [list(t) for t in TESTE],
           col_widths=[1.8, 6.4, 3.7], size=16, height=3.6,
           nota="sudo como ana: \"not in the sudoers file\". Certo: funcionário usa, não administra.")

    lista("PRÁTICA · 45 MIN", "O site da equipe com o ChatGPT", [
        "1.  chatgpt.com no PC",
        "2.  Página de comandos, passo 5: escolha o negócio, Copiar o pedido",
        "3.  Cole no ChatGPT. Ele responde com o código",
        "4.  Leia as regras do pedido: um arquivo só, UTF-8, sem imagem de fora",
    ], size=21, nota="Nunca cole senha, nome ou telefone de verdade na IA.", nota_cor=R)

    lista("PRÁTICA · COLE", "Publicar no servidor", ["$ " + c for c in PUBLICAR_CMD] + [
        "",
        "•  ChatGPT: Copiar código → nano: botão direito",
        "•  Ctrl+O, Enter, Ctrl+X → navegador: http://192.168.1N.10",
    ], size=18)

    frase("QUEM ASSINA", "A IA escreve rápido,\nmas não assina.\nQuem entrega o site é o técnico.", size=34)

    lista("CONFERIR E AJUSTAR", "O que a IA inventou?", [
        "•  Achem UMA coisa errada ou inventada: preço, rua, serviço, inglês",
        "•  Peçam um ajuste: cores, promoção da semana, perguntas frequentes",
        "•  Terminem com: \"me mande o código completo de novo\"",
        "$ sudo rm /var/www/html/index.html",
        "$ sudo nano /var/www/html/index.html",
        "•  Ctrl+F5 no navegador",
    ], size=19)

    lista("A VISITA", "Os sites das outras equipes", [
        "•  Celular no Wi-Fi de outra equipe, dados móveis desligados",
        "•  http://192.168.1M.10",
        "•  De quê é o site? Uma coisa boa nele",
        "",
        "Por que o PC da equipe 1 não abre o site da equipe 2?",
    ], size=22)

    lista("O QUE FICA", "Três frases pra guardar", [
        "\"Cada um entra com o seu nome e só abre a gaveta que é sua.\"",
        "",
        "\"Servidor tem endereço fixo.\"",
        "",
        "\"Só cole o que você sabe explicar.\"",
    ], size=23)

    lista("NO ENCONTRO 3", "A entrega do servidor", [
        "•  A pasta /srv/escritorio aparece neste PC como pasta de rede",
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

    # PÁGINA 1 — Bancada, endereço fixo, PuTTY
    cabecalho_documento(
        doc,
        "A BANCADA DO TÉCNICO",
        "Servidores com Linux · Encontro 2 · Curso FIC · CEJA Itapiranga · 1 de 4",
    )
    campo(doc, "Nome: _____________________________________________     Equipe: ______   Servidor: servidor-equipe____", space_after=6, size=10)

    add_heading(doc, "1  A bancada", level=2)
    for l in BANCADA:
        add_cmd(doc, l, size=8.5)
    campo(doc, "O servidor está na porta LAN ____ e o PC na porta LAN ____.   A internet entra na porta ________", size=10)
    campo(doc, "ipconfig no PC mostrou: 192.168.____.____      O PC abriu chatgpt.com?  ☐ Sim  ☐ Não", size=10)

    add_heading(doc, "2  Endereço fixo do servidor", level=2)
    campo(doc, "MAC do servidor (ip a → link/ether): ____ : ____ : ____ : ____ : ____ : ____", size=10)
    campo(doc, "Endereço antes da reserva: 192.168.____.____      Depois do sudo reboot: 192.168.____.____", size=10)
    campo(doc, "Por que servidor precisa de endereço fixo e o PC não precisa?", bold=True, size=10)
    linha(doc)

    add_heading(doc, "3  Entrar pelo PuTTY", level=2)
    tbl = add_table(doc, [["Onde", "O que fazer"]] + [list(r) for r in PUTTY], col_widths=[4.2, 13.2])
    small_table(tbl, size=9)
    tbl = add_table(doc, [
        ["Teste", "Funcionou?", "Se deu erro, o que dizia"],
        ["Entrar pelo PuTTY como equipeN", "☐ Sim  ☐ Não", ""],
        ["O recado do wall apareceu na tela do servidor", "☐ Sim  ☐ Não", ""],
        ["Abrir http://192.168.1N.10 no navegador do PC", "☐ Sim  ☐ Não", ""],
    ], col_widths=[7.4, 3, 7])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.75)
    campo(doc, "No PuTTY:  copiar = ______________________   colar = ______________________   Ctrl+C = ______________________", size=10)
    campo(doc, "Na primeira vez o PuTTY mostrou um Security Alert. Por que ele avisa?", bold=True, size=10)
    linha(doc)
    campo(doc, "Agora qualquer um ligado no roteador da equipe pode tentar entrar no servidor. O que protege ele?", bold=True, size=10)
    linha(doc)

    # PÁGINA 2 — Pastas e treino
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "ANDANDO PELAS PASTAS",
        "Pastas, usuários e acesso de longe · Encontro 2 · 2 de 4 · tudo DIGITADO",
    )

    add_heading(doc, "1  O passeio", level=2)
    tbl = add_table(doc, [["Comando", "Para quê", "O que apareceu"]] + [[c, d, ""] for c, d in PASSEIO],
                    col_widths=[3.8, 7.2, 6.4])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=9)
    campo(doc, "Na ajuda (ls --help), o que faz ls -l? _____________________   E ls -h? _____________________", size=10)
    campo(doc, "Diferença entre cd /var e cd var: ____________________________________________________________", size=10)

    add_heading(doc, "2  A pasta de treino do escritório", level=2)
    tbl = add_table(doc, [["Comando", "Para quê", "✓"]] + [[c, d, "☐"] for c, d in PEDIDO],
                    col_widths=[7.6, 8.8, 1])
    small_table(tbl, size=8.5)
    mono_col(tbl, 0, size=8.5)
    add_para(doc, "O tree tem que desenhar:", size=10, bold=True)
    for l in ARVORE:
        add_cmd(doc, l, size=9)
    campo(doc, "Saiu igual?  ☐ Sim, de primeira  ☐ Não: o que faltou e como a equipe consertou? ______________________", size=10)

    # PÁGINA 3 — Usuários, gavetas
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "OS FUNCIONÁRIOS E AS GAVETAS",
        "Pastas, usuários e acesso de longe · Encontro 2 · 3 de 4",
    )

    add_heading(doc, "1  Os funcionários (DIGITADO)", level=2)
    tbl = add_table(doc, [["Usuário", "Grupo", "Senha (a equipe escolhe, 10+)", "Criado", "No grupo (groups)"]] +
                    [[u, g, "", "☐", "☐"] for u, _, g in FUNCIONARIOS], col_widths=[2.4, 3, 6.4, 2, 3.6])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.65)
    campo(doc, "Por que o adduser é digitado e não colado? ____________________________________________________", size=10)

    add_heading(doc, "2  O que cada linha colada fez", level=2)
    tbl = add_table(doc, [["Linha colada", "O que ela fez (com as suas palavras)"]] + [[c, ""] for c in LINHAS_EXPLICAR],
                    col_widths=[7, 10.4])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=8.5)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)

    add_heading(doc, "3  ls -l /srv/escritorio", level=2)
    tbl = add_table(doc, [["Pasta", "Permissão (letras)", "Número", "Grupo", "Quem pode entrar"],
                          ["financeiro", "", "", "", ""], ["recepcao", "", "", "", ""], ["publico", "", "", "", ""]],
                    col_widths=[2.8, 4, 2, 3, 5.6])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.6)
    add_para(doc, "r = 4 · w = 2 · x = 1 · soma em cada trio: dono, grupo, outros", italic=True, size=9, color=COR_CINZA)

    add_heading(doc, "4  O teste das gavetas (DIGITADO: su - ana, su - carla)", level=2)
    tbl = add_table(doc, [["Quem", "Tenta", "O que aconteceu", "Certo?"]] + [[q, t, "", "☐"] for q, t, _ in TESTE],
                    col_widths=[1.8, 7, 6.6, 2])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.6)
    campo(doc, "Como ana, sudo apt update respondeu: _______________________________   Isso está certo?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "Por que não usar chmod 777 quando dá Permission denied? _______________________________________", size=10)

    # PÁGINA 4 — O site, visita, ficha
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "O SITE DA EQUIPE",
        "Pastas, usuários e acesso de longe · Encontro 2 · 4 de 4",
    )

    add_heading(doc, "1  O pedido para o ChatGPT", level=2)
    campo(doc, "O site é de: _____________________________________________   Nome do negócio: ______________________", size=10)
    campo(doc, "Cada regra do pedido tem um motivo. Explique:", bold=True, size=10)
    tbl = add_table(doc, [["Regra", "Por quê?"],
                          ["Um arquivo só", ""],
                          ["<meta charset=\"UTF-8\">", ""],
                          ["Sem imagens de outros sites", ""],
                          ["Telefone e endereço inventados", ""]], col_widths=[5.6, 11.8])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.65)

    add_heading(doc, "2  Publicar e conferir", level=2)
    tbl = add_table(doc, [["Passo", "✓"],
                          ["Colei o bloco Primeira publicação (guardou a página do Encontro 1 em ~/site-encontro1.html)", "☐"],
                          ["Copiei o código do ChatGPT e colei no nano com o botão direito", "☐"],
                          ["Salvei: Ctrl+O, Enter, Ctrl+X", "☐"],
                          ["Abriu em http://192.168.1____.10 no PC", "☐"],
                          ["Abriu no celular (Wi-Fi da equipe)", "☐"]], col_widths=[15.8, 1.6])
    small_table(tbl, size=9)
    campo(doc, "Uma coisa que a IA inventou, errou ou que não serve: _______________________________________________", size=10)
    campo(doc, "O ajuste que a equipe pediu: _____________________________________   Apareceu depois do Ctrl+F5?  ☐ Sim", size=10)
    campo(doc, "A IA fez a página. Quem entrega a página para o navegador? ___________________________________", size=10)

    add_heading(doc, "3  A visita", level=2)
    tbl = add_table(doc, [["Equipe visitada", "Endereço", "O site é de quê", "Uma coisa boa"],
                          ["", "192.168.1___.10", "", ""], ["", "192.168.1___.10", "", ""]],
                    col_widths=[3, 3.6, 5.2, 5.6])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.65)
    campo(doc, "Por que o PC da equipe não abre o site das outras equipes? ____________________________________", size=10)

    add_heading(doc, "4  Ficha técnica: o que entra hoje (passe para a folha do Encontro 1)", level=2)
    rows = [[FICHA_NOVA[i], "", FICHA_NOVA[i + 3], ""] for i in range(3)]
    tbl = add_table(doc, [["Item", "Como ficou", "Item", "Como ficou"]] + rows, col_widths=[3.8, 4.9, 3.8, 4.9])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.75)

    add_heading(doc, "TAREFA DA SEMANA", level=2, color=COR_DESTAQUE)
    campo(doc, "Pergunte a alguém que trabalha em escritório ou comércio: onde ficam os arquivos? Todo mundo vê tudo? ______", size=10)
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
    add_para(doc, "O técnico não fica na frente do servidor: ele monta o servidor e o computador na mesma rede e administra de longe, pelo SSH. O servidor tem endereço fixo, pastas organizadas, um usuário para cada funcionário e regras de quem abre o quê. E o que a IA entrega, o técnico confere antes de publicar.")

    add_heading(doc, "A bancada e o PuTTY", level=2)
    add_bullet(doc, "Internet na porta WAN do roteador. Servidor e PC nas portas LAN: ficam na mesma rede")
    add_bullet(doc, "Endereço fixo: reserva no roteador com o MAC do servidor (ip a → link/ether), depois sudo reboot")
    add_bullet(doc, "PuTTY: Host Name = endereço do servidor, porta 22, SSH. Na primeira vez, Accept. Para sair: exit")
    add_bullet(doc, "No PuTTY, selecionar copia e o botão direito cola. Ctrl+C cancela o comando")
    add_bullet(doc, "Qualquer um na rede pode tentar entrar. O que protege é a senha longa")

    add_heading(doc, "As pastas", level=2)
    add_table(doc, [["Pasta", "O que guarda"]] + [list(r) for r in RAIZ], col_widths=[3.5, 12])
    add_bullet(doc, "Caminho completo começa com / e funciona de qualquer lugar. Caminho curto parte de onde você está. .. sobe, ~ é a sua casa")

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

    add_heading(doc, "Usuários, grupos e permissões", level=2)
    tbl = add_table(doc, [["Comando", "O que faz"],
                          ["sudo adduser ana", "Cria a usuária (pergunta senha: digitar, não colar)"],
                          ["sudo groupadd financeiro", "Cria o grupo"],
                          ["sudo usermod -aG financeiro ana", "Põe a ana no grupo (-aG, com o a!)"],
                          ["groups ana", "Mostra os grupos da ana"],
                          ["sudo chgrp grupo pasta", "Muda o grupo dono da pasta"],
                          ["sudo chmod 770 pasta", "Muda as permissões"],
                          ["su - ana  ·  exit", "Vira a ana · volta"]], col_widths=[6.5, 9])
    mono_col(tbl, 0, size=9.5)
    add_para(doc, "Em ls -l: drwxrwx--- → d (pasta), depois três trios: DONO, GRUPO, OUTROS. r = ler (4), w = alterar (2), x = entrar (1). 770 = dono e grupo tudo, outros nada. 755 = todos leem, só o dono altera. Nunca 777.", size=10)
    add_bullet(doc, "Funcionário não é administrador: \"is not in the sudoers file\" é o certo")

    add_heading(doc, "O site com IA", level=2)
    add_bullet(doc, "O pedido certo: um arquivo HTML só, com <meta charset=\"UTF-8\">, sem imagens de outros sites, dados inventados")
    add_bullet(doc, "Publicar: sudo nano /var/www/html/index.html, colar com o botão direito, Ctrl+O, Enter, Ctrl+X. No navegador, Ctrl+F5")
    add_bullet(doc, "A IA inventa e erra. Quem confere e entrega é o técnico")
    add_bullet(doc, "Nunca cole senha, nome ou telefone de verdade na IA")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Cada um entra com o seu nome e só abre a gaveta que é sua.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Servidor tem endereço fixo.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Só cole o que você sabe explicar.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Por que o PC e o servidor precisam estar no mesmo roteador para o PuTTY funcionar?")
    add_bullet(doc, "Uma pasta está com rwxr-x---. Quanto dá em número? Quem pode alterar e quem só lê?")
    add_bullet(doc, "Alguém te manda um comando pelo WhatsApp dizendo \"cola isso no servidor que resolve\". O que você faz antes?")

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
