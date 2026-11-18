"""Gera os 4 arquivos do Encontro 3 do módulo de Servidores (Plano Mensal 4).

Servidores 3 — Pasta na rede e entrega do servidor
Último encontro do módulo (sexta 11/12), com a avaliação prática integradora.
Continua da bancada do Encontro 2 (servidor + PC com Windows no roteador da
equipe, tudo pelo PuTTY, página de comandos publicada em LINK_COMANDOS).
Hoje: instalar o Samba e compartilhar /srv/escritorio (financeiro, recepcao,
publico) com os usuários ana, bruno e carla; abrir as gavetas no Windows
(\\\\IP, unidade de rede Z:) e no celular; conferir pelo servidor (smbstatus).
Depois a entrega: ficha técnica final, conferência cruzada (uma equipe testa o
servidor da outra só com a ficha), conserto do que o laudo apontou e
apresentação da equipe. Padrão prática-pesada, sem role-play e sem duelo.
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
TITULO_ARQ = "Servidores 3 - Pasta na rede e entrega do servidor"
FOOTER = "Servidores com Linux · Encontro 3"
LINK_COMANDOS = "https://claude.ai/artifact/AwdfaVUJQgniKQLR6JFv5Z"
ARQ_COMANDOS = f"Comandos - {TITULO_ARQ}.html"


# Samba: digitados
SAMBA_INSTALAR = [
    ("sudo apt update", "Atualiza a lista de programas"),
    ("sudo apt install samba", "Instala o Samba. [Y/n]: Enter"),
    ("sudo -v", "Pede a senha antes, para o bloco colado não parar no meio"),
]

# Samba: o bloco colado (como está na página)
BLOCO_SAMBA = [
    "sudo cp /etc/samba/smb.conf /etc/samba/smb.conf.original",
    "sudo tee -a /etc/samba/smb.conf > /dev/null <<'FIM'",
    "",
    "[financeiro]",
    "   path = /srv/escritorio/financeiro",
    "   valid users = @financeiro",
    "   read only = no",
    "   force group = financeiro",
    "   create mask = 0660",
    "   directory mask = 0770",
    "",
    "[recepcao]",
    "   (igual, trocando financeiro por recepcao)",
    "",
    "[publico]",
    "   path = /srv/escritorio/publico",
    "   valid users = @financeiro @recepcao",
    "   read only = yes",
    "FIM",
]

# Na folha: o que cada linha da gaveta quer dizer
LINHAS_SAMBA = [
    ("[financeiro]", "O nome que aparece no Windows"),
    ("path = /srv/escritorio/financeiro", "Qual pasta do servidor vai para a rede"),
    ("valid users = @financeiro", ""),
    ("read only = no", ""),
    ("force group = financeiro", ""),
    ("create mask = 0660", ""),
    ("read only = yes (no publico)", ""),
]

LINHAS_SAMBA_RESP = [
    ("valid users = @financeiro", "Só quem é do grupo financeiro entra (o @ quer dizer grupo)"),
    ("read only = no", "Pode criar e alterar arquivos"),
    ("force group = financeiro", "Todo arquivo que entra fica do grupo financeiro: o bruno abre o que a ana salvou"),
    ("create mask = 0660", "Arquivo novo: dono e grupo leem e alteram (6 = 4+2), os outros nada"),
    ("directory mask = 0770", "Pasta nova: dono e grupo tudo, os outros nada"),
    ("read only = yes", "No publico: o mural. Todos leem, ninguém altera pela rede"),
]

SAMBA_LIGAR = [
    ("tail -25 /etc/samba/smb.conf", "As três gavetas têm que aparecer no fim do arquivo"),
    ("testparm", "Confere: Loaded services file OK. Depois Enter"),
    ("sudo systemctl restart smbd", "Reinicia o Samba com a configuração nova"),
    ("systemctl status smbd", "active (running). Sai com q"),
    ("sudo smbpasswd -a ana", "A senha de rede da ana (a mesma do Linux), duas vezes"),
    ("sudo smbpasswd -a bruno", "Igual"),
    ("sudo smbpasswd -a carla", "Igual"),
    ("sudo cp ~/treino/recepcao/recados.txt /srv/escritorio/publico/", "Um arquivo no mural, para ter o que ver"),
]

# Os testes das gavetas pela rede
TESTES = [
    ("ana", "Windows", "abrir financeiro e criar um arquivo", "Consegue"),
    ("ana", "Windows", "abrir recepcao", "Negado"),
    ("ana", "Windows", "abrir publico e ler recados.txt", "Consegue"),
    ("ana", "Windows", "alterar e salvar recados.txt", "Negado"),
    ("bruno", "Windows", "abrir o arquivo que a ana criou e alterar", "Consegue"),
    ("carla", "Celular", "abrir recepcao e ver os arquivos", "Consegue"),
    ("carla", "Celular", "abrir financeiro", "Negado"),
]

TRAVOU = [
    ("O Windows não encontra \\\\192.168.1N.10", "O PC está no roteador da equipe? Barras invertidas (\\\\). No servidor: systemctl status smbd"),
    ("Nome de usuário ou senha incorretos", "Faltou o smbpasswd -a daquele funcionário, ou a senha é outra. Tentar servidor-equipeN\\ana"),
    ("Erro 1219: várias conexões com mais de um nome de usuário", "O Windows já entrou com outro funcionário. Fechar o Explorador e no cmd: net use * /delete /y"),
    ("O Windows nega sem pedir senha", "Ficou lembrado o usuário errado. net use * /delete /y"),
    ("testparm: Unknown parameter", "Uma linha entrou cortada. sudo nano /etc/samba/smb.conf, ir ao fim e conferir com a página"),
    ("As gavetas aparecem duas vezes", "O bloco foi colado duas vezes. sudo cp /etc/samba/smb.conf.original /etc/samba/smb.conf e colar uma vez só"),
    ("Abre a gaveta mas não salva", "O funcionário não está no grupo (groups ana), ou é o publico, que é só leitura"),
    ("O bruno não consegue alterar o arquivo da ana", "Faltou force group ou create mask no bloco. Conferir com tail e reiniciar o smbd"),
    ("O celular não conecta", "Wi-Fi da equipe, dados móveis desligados. No Android, o tipo é SMB (Windows). Usuário sem o servidor-equipeN\\ na frente"),
    ("Pediu senha no meio do bloco colado", "Ctrl+C. Voltar ao original (linha acima), sudo -v, colar de novo"),
]

# A ficha técnica final (vai com o servidor)
FICHA = [
    ("Cliente / negócio do site", ""),
    ("Máquina (marca e modelo), RAM e disco", ""),
    ("Nome do servidor", "servidor-equipe__"),
    ("Endereço fixo (IP) e MAC", "192.168.___.10  ·  MAC:"),
    ("Roteador: painel e Wi-Fi", "Painel 192.168.___.1 · Wi-Fi:"),
    ("Administrador: usuário e senha", "equipe__  ·"),
    ("Como entrar de longe", "PuTTY · Host 192.168.___.10 · porta 22"),
    ("Site", "http://192.168.___.10"),
    ("Pasta na rede", "\\\\192.168.___.10"),
    ("Gaveta financeiro: quem entra", ""),
    ("Gaveta recepcao: quem entra", ""),
    ("Gaveta publico", "Todos leem, ninguém altera"),
    ("Funcionários: usuário e senha", "ana:            bruno:            carla:"),
    ("Como desligar", "sudo poweroff (nunca no botão nem na tomada)"),
    ("Instalado por / data", ""),
]

# O laudo da conferência cruzada
LAUDO = [
    "Conectamos no Wi-Fi da equipe só com a ficha",
    "O site abriu no navegador e no celular",
    "Entramos pelo PuTTY como administrador",
    "Como ana, a pasta na rede abriu e salvamos um arquivo no financeiro",
    "Como ana, a recepcao foi negada",
    "Como carla (celular), a recepcao abriu e o financeiro foi negado",
    "O publico abriu, mas não deixou alterar",
    "No PuTTY, sudo smbstatus mostrou a nossa conexão",
    "A ficha estava completa e legível",
]

# Critérios da avaliação integradora
RUBRICA = [
    ("Servidor funcionando", "Site, PuTTY e pasta na rede funcionam", "Um dos três falha", "Dois ou mais falham"),
    ("Gavetas certas", "Cada funcionário só na sua gaveta; publico só leitura", "Uma permissão errada", "Todos entram em tudo, ou ninguém entra"),
    ("Ficha técnica", "A outra equipe fez tudo só com a ficha", "Precisou de ajuda em 1 ou 2 itens", "Não deu para usar sem ajuda"),
    ("Laudo da outra equipe", "Completo, com o que faltou escrito", "Incompleto", "Não fez"),
    ("Apresentação", "Explica o que o servidor faz e como se usa, em linguagem de cliente", "Explica só a parte técnica", "Não apresentou"),
    ("Rodízio e participação", "Todos passaram por digitador, leitor e anotador", "Um ou dois concentraram", "Só um trabalhou"),
]


# ============================================================
# ROTEIRO
# ============================================================
def build_roteiro():
    doc = Document()
    setup_page(doc)

    cabecalho_documento(
        doc,
        "PASTA NA REDE E ENTREGA DO SERVIDOR",
        "Roteiro do professor · Servidores com Linux · Encontro 3 de 3 · sexta 11/12 · Curso FIC · CEJA Itapiranga – SC",
    )

    add_heading(doc, "A IDEIA DESTE ENCONTRO")
    add_para(doc, "É o fechamento do módulo e a avaliação prática integradora. Na primeira metade, o servidor cumpre o último pedido do escritório: guardar os arquivos num lugar só e abrir de qualquer computador. As gavetas criadas no Encontro 2 (/srv/escritorio) viram pastas de rede com o Samba, e a ana abre o financeiro como se fosse o disco Z: do computador dela.")
    add_para(doc, "Na segunda metade, a entrega. Cada equipe fecha a ficha técnica e entrega o servidor para outra equipe testar só com o papel, como o próximo técnico que chega um ano depois. O laudo volta, a equipe conserta o que faltou e apresenta o servidor para a turma.")
    add_para(doc, "A frase da noite: \"Servidor entregue é servidor que outra pessoa consegue usar só com a ficha.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "O que se resolve neste encontro", level=2)
    add_para(doc, "Cada aluno sai sabendo instalar e configurar uma pasta compartilhada com o Samba, explicar cada linha da configuração, criar a senha de rede dos usuários, abrir a pasta no Windows (pelo endereço e como unidade de rede) e no celular, conferir pelo servidor quem está conectado, documentar o servidor numa ficha técnica que funciona sozinha e testar o servidor de outra equipe com laudo.")

    add_heading(doc, "Onde este encontro fica no módulo", level=2)
    add_para(doc, "Fecha a questão motriz do Plano Mensal 4. Usa tudo que veio antes: a instalação (Encontro 1), a bancada com PuTTY, o endereço fixo, os usuários, os grupos e as permissões e o site da equipe (Encontro 2), o roteador e a ficha de entrega (Noite 8). A conferência cruzada repete a ideia do \"roteador trocado\" da Noite 8, agora com o servidor inteiro.")
    add_para(doc, "Não entre em: Samba como controlador de domínio, acesso de convidado sem senha (o Windows 11 bloqueia), lixeira de rede, cotas de disco, backup automático. Se perguntarem de backup, ver a Caixa de Ferramentas.", italic=True, color=COR_ALERTA)

    add_hr(doc)

    add_heading(doc, "ANTES DA AULA")
    add_heading(doc, "As bancadas", level=2)
    add_bullet(doc, "Cada equipe na MESMA bancada do Encontro 2: servidor e PC no roteador da equipe, PuTTY com a sessão servidor-equipeN salva")
    add_bullet(doc, "Internet nos roteadores: hoje se instala o Samba pelo apt")
    add_bullet(doc, "Confira em cada servidor, pelo PuTTY: groups ana bruno carla e ls -l /srv/escritorio. Se uma equipe não terminou as gavetas no Encontro 2, ela começa colando os blocos 3 e 4 da página do Encontro 2 (10 min)")
    add_bullet(doc, "Tenha as fotos das fichas técnicas: as senhas da ana, do bruno e da carla estão lá")

    add_heading(doc, "A página de comandos", level=2)
    add_bullet(doc, "Abra e deixe nos favoritos dos PCs: " + LINK_COMANDOS + ". A página é PRIVADA até você mudar: no menu Compartilhar, libere para quem tiver o link. Plano B: o arquivo \"" + ARQ_COMANDOS + "\" (na pasta do curso) num pendrive, aberto direto no navegador")
    add_bullet(doc, "Ensaio de 15 minutos numa bancada: instale o Samba, cole o bloco do passo 3, abra \\\\192.168.1N.10 no Windows como ana, depois rode net use * /delete /y e entre como bruno. Saber como o Windows se comporta com a troca de usuário evita a maior parte dos travamentos")
    add_bullet(doc, "Celulares: CX Explorador de Arquivos (Android, gratuito) ou o app Arquivos (iPhone, já vem instalado). Peça no grupo da turma que instalem o CX durante a semana")

    add_heading(doc, "Imprima", level=2)
    add_bullet(doc, "Folha do aluno (4 páginas), uma por pessoa. A página 3 (ficha técnica) também uma EXTRA por equipe, a limpo: é ela que vai com o servidor para a outra equipe")
    add_bullet(doc, "Resumo do aluno: é o resumo do módulo inteiro, para guardar")
    add_bullet(doc, "Uma cópia dos critérios de avaliação (abaixo) para você marcar durante a noite")

    add_hr(doc)

    add_heading(doc, "QUADRO DE TEMPO — 210 MINUTOS")
    add_table(doc, [
        ["Bloco", "Tempo", "O que acontece"],
        ["1. Retomada + bancada", "15 min", "A questão do escritório: o que já foi resolvido e o que falta"],
        ["2. PRÁTICA — O Samba no servidor", "30 min", "Instalar, colar as gavetas, explicar cada linha, testparm, senhas da rede"],
        ["3. PRÁTICA — As gavetas no Windows", "30 min", "\\\\192.168.1N.10 como ana, os testes, unidade Z:, trocar para o bruno"],
        ["4. PRÁTICA — No celular + conferir pelo servidor", "15 min", "A carla pelo celular, sudo smbstatus, o grupo do arquivo da ana"],
        ["INTERVALO", "15 min", ""],
        ["5. PRÁTICA — A ficha técnica final", "20 min", "Passar a limpo e conferir: cada item testado antes de entregar"],
        ["6. PRÁTICA — Conferência cruzada", "40 min", "30 min no servidor da outra equipe só com a ficha + 10 min consertando o que o laudo apontou"],
        ["7. Apresentação das equipes", "30 min", "4 minutos por equipe, na bancada, a turma em volta"],
        ["8. Fechamento do módulo", "15 min", "A questão motriz respondida, a profissão, a autoavaliação"],
    ], col_widths=[6.8, 2, 7.7])
    add_para(doc, "Blocos 2 a 7: 165 min de atividade do aluno (79% do total). Exposição: a retomada, o fechamento e falas curtas dentro das práticas. Rodízio de sempre: digitador, leitor, anotador, trocando a cada bloco.", italic=True, color=COR_CINZA)

    add_heading(doc, "Critérios da avaliação integradora", level=2)
    add_para(doc, "Marque durante os Blocos 6 e 7. O laudo da outra equipe é a sua principal evidência do critério Ficha técnica.", size=10, italic=True, color=COR_CINZA)
    tbl = add_table(doc, [["Critério", "Atingiu", "Parcial", "Não atingiu"]] + [list(r) for r in RUBRICA],
                    col_widths=[3.4, 4.8, 4, 4.3])
    small_table(tbl, size=9)

    add_hr(doc)

    # BLOCO 1
    add_heading(doc, "BLOCO 1 — RETOMADA + BANCADA · 15 MIN")
    add_bullet(doc, "As equipes ligam a bancada, entram pelo PuTTY e abrem o site no navegador do PC")
    add_bullet(doc, "Escreva a questão do escritório no quadro, em três partes, e pergunte o que já está feito:")
    add_table(doc, [
        ["O escritório pediu", "Feito?"],
        ["Ter uma página interna para os funcionários", "Sim: o site da equipe (Nginx), Encontros 1 e 2"],
        ["Guardar os arquivos num lugar só", "Metade: as gavetas existem no servidor, mas só pelo PuTTY"],
        ["Acessar de qualquer computador", "Ainda não: a ana não usa PuTTY. Ela quer abrir a pasta no Windows"],
    ], col_widths=[7, 9.5])
    add_bullet(doc, "A tarefa do Encontro 2: \"onde ficam os arquivos no trabalho de quem vocês perguntaram? todo mundo vê tudo?\". Duas ou três respostas. Normalmente aparece \"cada um no seu computador\" ou \"num pendrive\": pergunte o que acontece quando aquele computador estraga")
    add_para(doc, "\"Hoje a gaveta do financeiro vai aparecer no computador da ana como se fosse um disco dela. E no fim vocês entregam o servidor: outra equipe vai testar tudo só com o papel que vocês escreverem.\"", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 2
    add_heading(doc, "BLOCO 2 — PRÁTICA · O SAMBA NO SERVIDOR · 30 MIN")
    add_para(doc, "Página 1 da folha, passos 2 a 5 da página de comandos.")
    add_bullet(doc, "\"O Windows fala uma língua para pastas de rede (SMB). O Samba é o programa que ensina o Linux a falar essa língua.\"")
    add_heading(doc, "Instalar (digitando)", level=2)
    for c, _ in SAMBA_INSTALAR:
        add_cmd(doc, c)
    add_heading(doc, "As gavetas na rede (colando)", level=2)
    add_para(doc, "O bloco do passo 3. O leitor lê em voz alta antes do Enter. A primeira linha é o hábito do técnico: guardar o arquivo original antes de mexer. O tee -a acrescenta o texto até a palavra FIM no fim do arquivo de configuração.")
    for l in BLOCO_SAMBA:
        add_cmd(doc, l if l else " ", size=9.5)
    add_para(doc, "Pare a turma e faça a tabela da página 1 da folha junto. Cada linha da gaveta liga com algo que eles já fizeram:", size=11)
    add_table(doc, [["Linha", "Quer dizer"]] + [list(r) for r in LINHAS_SAMBA_RESP], col_widths=[5.5, 11])
    add_bullet(doc, "O 0660 é o mesmo jogo de números do chmod (Encontro 2): 6 = 4 + 2 = ler e alterar. O zero da frente pode ignorar")
    add_bullet(doc, "Por que o force group: sem ele, o arquivo que a ana salva fica do grupo \"ana\", e o bruno, do mesmo setor, não consegue alterar. É o problema mais comum de pasta compartilhada mal configurada")
    add_heading(doc, "Conferir, ligar e as senhas (digitando)", level=2)
    tbl = add_table(doc, [["Comando", "Para quê"]] + [list(r) for r in SAMBA_LIGAR], col_widths=[7.5, 9])
    mono_col(tbl, 0, size=9)
    add_bullet(doc, "testparm é o \"corretor\" da configuração. Se ele reclamar, a equipe lê a mensagem: ela diz a linha")
    add_bullet(doc, "\"Por que o Samba pede outra senha se a ana já tem senha no Linux?\" O Samba guarda as senhas de rede num lugar próprio, no formato que o Windows entende. Usem a mesma, para o funcionário não ter duas")
    add_para(doc, "Onde dilatar: cat /etc/samba/smb.conf.original | grep -v '^#' | grep -v '^;' | grep -v '^$' mostra só as linhas que valem no arquivo original. A equipe acha a gaveta [printers] e a [homes] e discute o que elas fariam.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 3
    add_heading(doc, "BLOCO 3 — PRÁTICA · AS GAVETAS NO WINDOWS · 30 MIN")
    add_para(doc, "Página 2 da folha, passo 6 da página. O PC da bancada agora é o computador da ana.")
    add_bullet(doc, "Explorador de Arquivos (Windows + E) → barra de endereço → \\\\192.168.1N.10 → Enter. Aparecem as três gavetas. Explique as barras: \\\\ quer dizer \"outro computador na rede\"")
    add_bullet(doc, "Abrir financeiro: usuário ana e a senha. Desmarcar \"Lembrar minhas credenciais\" (hoje vai haver troca de usuário)")
    add_bullet(doc, "Os testes da folha como ana: criar arquivo no financeiro, tentar a recepcao, ler e tentar alterar o mural")
    add_bullet(doc, "Unidade de rede: botão direito em Este Computador → Mapear unidade de rede → Z: → \\\\192.168.1N.10\\financeiro → Reconectar ao entrar. \"É assim que o funcionário usa: o disco Z: é a gaveta do setor, e ele nem precisa saber que existe um servidor\"")
    add_heading(doc, "Trocar para o bruno", level=2)
    add_bullet(doc, "Fechar todas as janelas do Explorador. cmd → net use * /delete /y. Abrir \\\\192.168.1N.10 de novo e entrar como bruno")
    add_bullet(doc, "O teste que importa: o bruno abre o arquivo que a ana criou, altera e salva. Funcionou por causa do force group")
    add_para(doc, "O erro 1219 vai aparecer em alguma equipe (\"não são permitidas várias conexões... com mais de um nome de usuário\"). Não resolva: aponte para a tabela Se travar da página e deixe a equipe ler a mensagem. É o método do módulo.", italic=True, color=COR_PRIMARIA)
    add_para(doc, "Onde dilatar: botão direito no arquivo da ana → Propriedades → Segurança. O Windows mostra o dono e as permissões do arquivo que está no Linux. Compare com o ls -l no PuTTY.", italic=True, color=COR_DESTAQUE)

    add_hr(doc)

    # BLOCO 4
    add_heading(doc, "BLOCO 4 — PRÁTICA · NO CELULAR + CONFERIR PELO SERVIDOR · 15 MIN")
    add_bullet(doc, "Celular no Wi-Fi da equipe. iPhone: Arquivos → os três pontinhos → Conectar ao Servidor → smb://192.168.1N.10. Android: CX Explorador → Rede → Novo local → Remoto → SMB")
    add_bullet(doc, "Entrar como carla: a recepcao abre, o financeiro não")
    add_bullet(doc, "Com a carla conectada no celular e a ana (ou o bruno) no PC: no PuTTY, sudo smbstatus. Aparecem as duas pessoas, os endereços e as gavetas. \"O técnico vê quem está usando o servidor agora, sem perguntar para ninguém\"")
    add_bullet(doc, "ls -l /srv/escritorio/financeiro: o arquivo criado pelo Windows tem dono ana e grupo financeiro")

    add_heading(doc, "INTERVALO — 15 MIN", level=2)

    add_hr(doc)

    # BLOCO 5
    add_heading(doc, "BLOCO 5 — PRÁTICA · A FICHA TÉCNICA FINAL · 20 MIN")
    add_para(doc, "Página 3 da folha, mais a cópia extra por equipe, a limpo e a caneta. Ela junta as fichas dos Encontros 1 e 2 e a de hoje.")
    add_bullet(doc, "Regra: nenhum item vai para a ficha sem ser testado de novo agora. O anotador lê o item, o digitador testa, o leitor confere com a tela")
    add_bullet(doc, "Senhas vão na ficha: ela é do cliente, guardada com ele. Na etiqueta da máquina vai só o nome do servidor, o endereço e quem instalou (como na Noite 8)")
    add_bullet(doc, "\"Escrevam para alguém que nunca viu este servidor. Se precisar perguntar para vocês, a ficha falhou.\"")
    add_bullet(doc, "Antes de entregar: a equipe fecha o PuTTY, roda net use * /delete /y no PC e tira a sessão salva do PuTTY (Delete). A outra equipe tem que começar do zero, só com o papel")

    add_hr(doc)

    # BLOCO 6
    add_heading(doc, "BLOCO 6 — PRÁTICA · CONFERÊNCIA CRUZADA · 40 MIN")
    add_para(doc, "A equipe 1 vai para a bancada da equipe 2, a 2 para a 3, e a última para a 1. Cada uma leva só a ficha da equipe dona da bancada, que fica sentada longe e não pode ajudar (só se a equipe visitante chamar o professor e o professor autorizar).")
    add_heading(doc, "No servidor da outra equipe — 30 min", level=2)
    add_bullet(doc, "Página 4 da folha: o laudo com 9 itens. Cada item ✓ ou ✗, e quando ✗, o que aconteceu, com a mensagem de erro")
    add_bullet(doc, "Primeiro o Wi-Fi e o site no celular, depois o PuTTY (sessão nova, criada com os dados da ficha), depois a pasta como ana no PC e como carla no celular")
    add_bullet(doc, "No fim, a equipe visitante escreve uma frase para a equipe dona: o que estava melhor na ficha e o que faltou")
    add_heading(doc, "De volta ao próprio servidor — 10 min", level=2)
    add_bullet(doc, "As equipes voltam, leem o laudo e consertam o que deu ✗ (no servidor ou na ficha). Anotam na folha o que consertaram")
    add_para(doc, "O laudo não é nota da outra equipe: é a prova de que o serviço funciona sem o técnico do lado. Diga isso antes de começar, para ninguém esconder problema nem exagerar.", italic=True, color=COR_PRIMARIA)

    add_hr(doc)

    # BLOCO 7
    add_heading(doc, "BLOCO 7 — APRESENTAÇÃO DAS EQUIPES · 30 MIN")
    add_para(doc, "A turma vai de bancada em bancada. Cada equipe tem 4 minutos e responde, com o servidor ligado, às quatro perguntas da página 4 da folha (cada um da equipe fala pelo menos uma):")
    add_bullet(doc, "1. O que este servidor faz pelo escritório? (em linguagem de cliente, sem comando)")
    add_bullet(doc, "2. Como a ana usa no dia a dia? (mostrar o Z: abrindo e o site)")
    add_bullet(doc, "3. O que o laudo da outra equipe encontrou e como vocês consertaram?")
    add_bullet(doc, "4. O que o cliente precisa saber para não estragar? (desligar certo, não trocar a reserva do roteador, guardar a ficha)")
    add_para(doc, "Controle o tempo com o celular. Com 6 equipes, são 24 minutos de fala e 6 de deslocamento. Com menos equipes, peça a cada uma que mostre também o site do negócio e uma coisa que a IA errou no Encontro 2.", italic=True, color=COR_CINZA)

    add_hr(doc)

    # BLOCO 8
    add_heading(doc, "BLOCO 8 — FECHAMENTO DO MÓDULO · 15 MIN")
    add_bullet(doc, "Volte à tabela do Bloco 1 e marque as três linhas como feitas. Acrescente uma quarta que o escritório nem pediu: administrar de longe (PuTTY)")
    add_bullet(doc, "\"Em três sextas, uma máquina que ia para o lixo virou o servidor de um escritório: site, pasta de rede, funcionários com senha, tudo documentado. Esse é um serviço que se vende.\"")
    add_bullet(doc, "Onde isso aparece no trabalho: escritório de contabilidade, clínica, escola, loja de peças, cooperativa. Quem sabe instalar, documentar e dar manutenção tem cliente fixo (a manutenção é mensal)")
    add_bullet(doc, "Autoavaliação (página 4 da folha, 3 min). Todos desligam pelo PuTTY: sudo poweroff")

    add_hr(doc)

    add_heading(doc, "SE TRAVAR")
    add_table(doc, [["O que aparece", "O que fazer"]] + [list(r) for r in TRAVOU], col_widths=[5.5, 11])

    add_hr(doc)

    add_heading(doc, "SE PERGUNTAREM")
    add_para(doc, "\"Não é mais fácil usar o Google Drive?\"", bold=True)
    add_para(doc, "> \"Para muita gente é. O Drive cobra por mês, depende da internet e os arquivos ficam com o Google. O servidor no escritório funciona sem internet, é rápido dentro da rede e os dados ficam lá. Mas alguém precisa cuidar dele: é aí que entra o técnico.\"", italic=True)
    add_para(doc, "\"E se o disco do servidor queimar?\"", bold=True)
    add_para(doc, "> \"Perde tudo. Por isso servidor de escritório precisa de cópia de segurança: um HD externo, outra máquina ou a nuvem, de preferência automática. O primeiro passo é a Caixa de Ferramentas de hoje.\"", italic=True)
    add_para(doc, "\"Funciona no Mac e no Linux também?\"", bold=True)
    add_para(doc, "> \"Funciona. O SMB é a língua de pasta de rede que quase todo sistema fala. No Mac: Finder → Ir → Conectar ao Servidor → smb://endereço.\"", italic=True)
    add_para(doc, "\"Dá para abrir a pasta de casa?\"", bold=True)
    add_para(doc, "> \"Não deixando a pasta aberta para a internet: é a porta que os vírus de sequestro de arquivos procuram. O jeito certo é uma VPN, que coloca o computador de casa dentro da rede do escritório. Fica para um próximo curso.\"", italic=True)
    add_para(doc, "\"Quanto se cobra por um serviço desses?\"", bold=True)
    add_para(doc, "> \"Depende da cidade e do tamanho. Faça a conta: horas de instalação, deslocamento, a máquina (se for sua) e a manutenção mensal. A Caixa de Ferramentas tem uma atividade para isso.\"", italic=True)

    add_hr(doc)

    add_heading(doc, "DEPOIS DA AULA")
    add_bullet(doc, "Recolha as fichas técnicas a limpo: têm todas as senhas. Guarde ou destrua; não deixe no laboratório")
    add_bullet(doc, "Decida o destino das máquinas: se servem para outra turma, um Reset no roteador e reinstalação. Se ficam, troque as senhas de equipeN")
    add_bullet(doc, "Registre os critérios de cada equipe no diário. Indicador do módulo: quantas equipes passaram na conferência cruzada sem ajuda")

    add_hr(doc)

    add_heading(doc, "DESCRITIVO PARA O DIÁRIO")
    add_para(doc, "Realizei a avaliação prática integradora do estudo de servidores com Linux. Os estudantes retomaram a situação-problema do escritório e identificaram o que já estava resolvido. Em equipes, instalaram o serviço de compartilhamento de arquivos Samba, configuraram as pastas de cada setor com acesso restrito por grupo e permissões de criação, explicando por escrito cada linha da configuração, e criaram as senhas de rede dos usuários. Acessaram as pastas a partir de computador com Windows, pelo endereço de rede e como unidade de rede, e a partir do celular, testando o acesso de cada usuário e verificando pelo servidor as conexões ativas. Consolidaram a ficha técnica do servidor e realizaram conferência cruzada, em que cada equipe testou o servidor de outra equipe apenas com a ficha recebida e registrou laudo; as equipes corrigiram os problemas apontados e apresentaram o servidor à turma, explicando seu funcionamento e uso. Encerramos com a discussão da administração de pequenos servidores como frente de prestação de serviço e autoavaliação.")

    add_hr(doc)

    add_heading(doc, "CAIXA DE FERRAMENTAS DIDÁTICAS", color=COR_DESTAQUE)
    add_para(doc, "Se sobrar tempo, ou para uma aula de reforço:", italic=True, color=COR_CINZA)
    add_para(doc, "1. A primeira cópia de segurança", bold=True)
    add_para(doc, "Pelo PuTTY: sudo tar czf ~/backup-escritorio.tar.gz /srv/escritorio e ls -lh ~ para ver o tamanho. Depois a equipe apaga um arquivo da gaveta pelo Windows e restaura do backup (sudo tar xzf ~/backup-escritorio.tar.gz -C / srv/escritorio/financeiro/arquivo.txt). Discuta: o backup está no mesmo disco. Se o disco queimar, adianta? 20 min.")
    add_para(doc, "2. O setor novo, sem página", bold=True)
    add_para(doc, "O escritório contratou a dora, do setor compras. A equipe cria o usuário, o grupo, a pasta, a gaveta no Samba e a senha de rede sozinha, olhando só a ficha, o resumo e o smb.conf. Testa no Windows e atualiza a ficha. É o teste de autonomia do módulo. 25 min.")
    add_para(doc, "3. Servidor próprio ou nuvem?", bold=True)
    add_para(doc, "Matemática. As equipes pesquisam no celular o preço mensal de um plano de armazenamento em nuvem para 5 pessoas e comparam com o servidor: consumo de energia de 24 h por 30 dias (60 W, tarifa da fatura de casa) mais uma manutenção mensal que eles mesmos estimam. Em quantos meses um empata com o outro? 20 min.")
    add_para(doc, "4. O orçamento do serviço", bold=True)
    add_para(doc, "Cada equipe monta o orçamento para o escritório na planilha da Noite 3: horas de instalação (usem o tempo real dos três encontros), deslocamento, materiais, e um valor de manutenção mensal. Compara com as outras equipes e discute por que deu diferente. 25 min.")
    add_para(doc, "5. O manual do funcionário", bold=True)
    add_para(doc, "Uma folha para colar do lado do computador da ana, em linguagem de vó: onde está o disco Z:, o que é o mural, o que fazer se pedir senha, a quem ligar. Outra equipe lê e diz se entenderia sem ajuda. Linguagens. 20 min.")

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
                if run.text.startswith("$ ") or run.text.startswith("\\\\") or run.text.startswith("[") or run.text.startswith("   "):
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

    add_slide_cover(prs, "SERVIDORES COM LINUX · ENCONTRO 3", "PASTA NA REDE", "E ENTREGA",
                    "O último pedido do escritório. E o servidor entregue",
                    "Curso FIC · Montador e Reparador de Computadores · CEJA Itapiranga – SC")

    tabela("RETOMADA", "O que o escritório pediu", [
        ["O escritório pediu", "Feito?"],
        ["Uma página interna para os funcionários", "Sim: o site da equipe"],
        ["Guardar os arquivos num lugar só", "Metade: as gavetas existem, mas só pelo PuTTY"],
        ["Acessar de qualquer computador", "Ainda não: a ana quer abrir no Windows"],
    ], col_widths=[5.6, 6.3], size=19, height=2.8,
        nota="Hoje: a gaveta do financeiro vira o disco Z: da ana.")

    frase("A FRASE DA NOITE", "Servidor entregue é servidor\nque outra pessoa consegue usar\nsó com a ficha.", size=34)

    lista("O SAMBA", "A língua das pastas de rede", [
        "•  O Windows fala uma língua para pastas de rede: SMB",
        "•  O Samba ensina o Linux a falar essa língua",
        "•  Funciona com Windows, Mac, Linux e celular",
        "",
        "$ sudo apt update",
        "$ sudo apt install samba",
        "$ sudo -v",
    ], size=21)

    lista("PRÁTICA · 30 MIN", "As gavetas na rede (passo 3, colar)", BLOCO_SAMBA[3:10] + [
        "",
        "O leitor lê em voz alta antes do Enter. A 1ª linha guarda o original.",
    ], size=17)

    tabela("O QUE CADA LINHA QUER DIZER", "Igual ao chmod do Encontro 2", [["Linha", "Quer dizer"]] + [list(r) for r in LINHAS_SAMBA_RESP],
           col_widths=[4.3, 7.6], size=14, height=3.9, top=2.3,
           nota="force group: sem ele, o bruno não altera o que a ana salvou.")

    tabela("CONFERIR, LIGAR E AS SENHAS", "Passos 4 e 5, digitando", [["Comando", "Para quê"]] + [list(r) for r in SAMBA_LIGAR],
           col_widths=[6.2, 5.7], size=12, height=4.0, top=2.3,
           nota="A senha de rede é a mesma do Linux: o funcionário não fica com duas.")

    lista("PRÁTICA · 30 MIN", "No Windows, como ana", [
        "1.  Explorador de Arquivos (Windows + E)",
        "\\\\192.168.1N.10   (\\\\ = outro computador na rede)",
        "2.  financeiro → usuário ana e a senha. DESMARCAR \"Lembrar\"",
        "3.  Os testes da folha: criar, tentar a recepcao, o mural",
        "4.  Este Computador → Mapear unidade de rede → Z:",
        "\\\\192.168.1N.10\\financeiro",
    ], size=20, nota="O disco Z: é a gaveta do setor. O funcionário nem sabe que existe um servidor.")

    frase("TROCAR DE FUNCIONÁRIO", "Feche o Explorador. No cmd:\n$ net use * /delete /y\n\nEntre como bruno: ele altera\no arquivo que a ana criou?", size=28)

    frase("O ERRO 1219", "\"Não são permitidas várias conexões\ncom mais de um nome de usuário\"\n\nLeia a mensagem. A resposta está na página.", cor=R, size=26)

    lista("PRÁTICA · 15 MIN", "No celular, como carla", [
        "iPhone:  Arquivos → ··· → Conectar ao Servidor",
        "         smb://192.168.1N.10",
        "",
        "Android:  CX Explorador → Rede → Novo local → Remoto → SMB",
        "",
        "•  A recepcao abre? E o financeiro?",
    ], size=21)

    lista("CONFERIR PELO SERVIDOR", "Quem está usando agora?", [
        "$ sudo smbstatus",
        "   Quem está conectado, de onde e em qual gaveta",
        "",
        "$ ls -l /srv/escritorio/financeiro",
        "   O arquivo do Windows: dono ana, grupo financeiro",
    ], size=22)

    frase("", "INTERVALO\n15 minutos", cor=K)

    lista("PRÁTICA · 20 MIN", "A ficha técnica final", [
        "•  Junta os Encontros 1, 2 e 3. A limpo, a caneta",
        "•  Nenhum item entra sem ser testado de novo agora",
        "•  Senhas na ficha (fica com o cliente). Na etiqueta, nunca",
        "•  Escrevam para quem nunca viu este servidor",
        "",
        "Antes de entregar: feche o PuTTY, apague a sessão salva,",
        "net use * /delete /y",
    ], size=20)

    frase("A LIÇÃO DA NOITE 8", "Ficha que só quem escreveu entende\nnão é ficha.", size=34)

    lista("PRÁTICA · 40 MIN", "Conferência cruzada", [
        "•  Equipe 1 → bancada da 2, a 2 → 3 … a última → 1",
        "•  Só com a ficha. A equipe dona fica longe e não ajuda",
        "•  Laudo com 9 itens: ✓ ou ✗, e a mensagem de erro",
        "•  30 min no servidor da outra equipe",
        "•  10 min de volta: consertar o que o laudo apontou",
    ], size=21, nota="O laudo não é nota: é a prova de que o serviço funciona sem você do lado.")

    lista("APRESENTAÇÃO · 4 MIN POR EQUIPE", "Na bancada, com o servidor ligado", [
        "1.  O que este servidor faz pelo escritório? (sem comando)",
        "2.  Como a ana usa no dia a dia? (o Z: e o site)",
        "3.  O que o laudo encontrou e como vocês consertaram?",
        "4.  O que o cliente precisa saber para não estragar?",
    ], size=22, nota="Cada um da equipe fala pelo menos uma.")

    tabela("O MÓDULO", "O pedido do escritório, resolvido", [
        ["O escritório pediu", "Como"],
        ["Uma página interna", "Nginx, com o site do negócio"],
        ["Guardar os arquivos num lugar só", "As gavetas em /srv/escritorio"],
        ["Acessar de qualquer computador", "Samba: \\\\endereço e o disco Z:"],
        ["(bônus) Administrar de longe", "PuTTY / SSH"],
    ], col_widths=[5.6, 6.3], size=19, height=3.2,
        nota="Uma máquina que ia para o lixo virou o servidor de um escritório.")

    lista("O QUE FICA", "Três frases pra guardar", [
        "\"Servidor entregue é o que outro usa só com a ficha.\"",
        "",
        "\"Antes de mexer na configuração, guarde o original.\"",
        "",
        "\"Cada setor na sua gaveta. O mural, todos leem.\"",
    ], size=23)

    lista("E AGORA", "Onde isso vira trabalho", [
        "•  Escritório de contabilidade, clínica, escola, loja, cooperativa",
        "•  Instalar é um serviço. Manter é um cliente fixo (todo mês)",
        "•  Quem documenta bem é chamado de novo",
        "",
        "Autoavaliação na página 4. Desligar: sudo poweroff",
    ], size=22)

    path = os.path.join(OUTDIR, f"Aula - {TITULO_ARQ}.pptx")
    prs.save(path)
    return path


# ============================================================
# FOLHA DO ALUNO
# ============================================================
def build_folha():
    doc = Document()
    setup_page(doc, margem=Cm(1.6))

    # PÁGINA 1 — O Samba no servidor
    cabecalho_documento(
        doc,
        "A PASTA NA REDE: O SAMBA",
        "Servidores com Linux · Encontro 3 · Curso FIC · CEJA Itapiranga · 1 de 4",
    )
    campo(doc, "Nome: _____________________________________________     Equipe: ______   Servidor: 192.168.____.10", space_after=6, size=10)
    campo(doc, "Rodízio · Digitador: ____________   Leitor: ____________   Anotador: ____________   (troca a cada bloco)", size=10)

    add_heading(doc, "1  Instalar e colar as gavetas (passos 2 e 3 da página)", level=2)
    tbl = add_table(doc, [["Passo", "✓"],
                          ["sudo apt install samba terminou sem erro", "☐"],
                          ["sudo -v antes de colar", "☐"],
                          ["Bloco \"Compartilhar /srv/escritorio\" colado UMA vez, lido em voz alta antes do Enter", "☐"],
                          ], col_widths=[15.4, 2])
    small_table(tbl, size=9.5)

    add_heading(doc, "2  O que cada linha da gaveta quer dizer", level=2)
    add_para(doc, "As duas primeiras estão respondidas. Complete as outras com a equipe.", italic=True, size=9, color=COR_CINZA)
    tbl = add_table(doc, [["Linha", "Quer dizer"]] + [list(r) for r in LINHAS_SAMBA], col_widths=[6.2, 11.2])
    small_table(tbl, size=9.5)
    mono_col(tbl, 0, size=9)
    for row in tbl.rows[3:]:
        row.height = Cm(0.85)
    campo(doc, "No create mask = 0660, o 6 quer dizer ____ + ____ = ler e ______________. O último 0 quer dizer: os outros ______________", size=10)
    campo(doc, "Por que a primeira linha do bloco guarda uma cópia (smb.conf.original)?", bold=True, size=10)
    linha(doc)

    add_heading(doc, "3  Conferir, ligar e as senhas (passos 4 e 5)", level=2)
    tbl = add_table(doc, [["Comando", "O que deu", "✓"]] + [[c, "", "☐"] for c, _ in SAMBA_LIGAR],
                    col_widths=[7.8, 8.6, 1])
    small_table(tbl, size=9)
    mono_col(tbl, 0, size=8.5)
    for row in tbl.rows[1:]:
        row.height = Cm(0.62)
    campo(doc, "Por que o Samba pede uma senha separada, se a ana já tem senha no Linux? ___________________________________", size=10)

    # PÁGINA 2 — Testes no Windows e no celular
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "AS GAVETAS NO WINDOWS E NO CELULAR",
        "Pasta na rede e entrega do servidor · Encontro 3 · 2 de 4",
    )

    add_heading(doc, "1  Abrir pela rede", level=2)
    campo(doc, "No Explorador do Windows, a equipe digitou: ______________________________   Apareceram quantas gavetas? ____", size=10)
    campo(doc, "O que quer dizer o \\\\ no começo do endereço? __________________________________________", size=10)

    add_heading(doc, "2  Os testes", level=2)
    tbl = add_table(doc, [["Quem", "Onde", "Tenta", "O certo é", "O que aconteceu", "✓"]] +
                    [[q, o, t, c, "", "☐"] for q, o, t, c in TESTES],
                    col_widths=[1.5, 1.8, 5.4, 2, 5.6, 1.1])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.8)
    campo(doc, "Celular usado para a carla: ☐ iPhone (Arquivos)  ☐ Android (CX Explorador)  ☐ outro: __________", size=10)

    add_heading(doc, "3  O disco Z: da ana", level=2)
    campo(doc, "Mapeamos \\\\192.168.____.10\\______________ como unidade ____:     Apareceu em Este Computador?  ☐ Sim  ☐ Não", size=10)
    campo(doc, "Para trocar da ana para o bruno no mesmo PC, a equipe usou: _______________________________", size=10)
    campo(doc, "Apareceu o erro 1219?  ☐ Sim  ☐ Não     Se apareceu, o que ele dizia, com as suas palavras? ____________________", size=10)

    add_heading(doc, "4  Conferir pelo servidor", level=2)
    campo(doc, "sudo smbstatus mostrou ______ conexões. Usuários: ______________   Endereços: ______________________", size=10)
    campo(doc, "ls -l do arquivo da ana: dono ______________   grupo ______________", size=10)
    campo(doc, "Qual linha da gaveta fez o arquivo ficar com esse grupo? E por que isso é bom para o bruno?", bold=True, size=10)
    linha(doc)
    linha(doc)

    # PÁGINA 3 — Ficha técnica
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "FICHA TÉCNICA DO SERVIDOR",
        "Vai junto com o servidor · fica guardada com o cliente · Encontro 3 · 3 de 4",
    )
    add_para(doc, "A limpo e a caneta. Cada item testado de novo antes de escrever. Escreva para alguém que nunca viu este servidor.", italic=True, size=9.5, color=COR_ALERTA)
    tbl = add_table(doc, [["Item", "Como ficou"]] + [list(r) for r in FICHA], col_widths=[6, 11.4])
    small_table(tbl, size=9.5)
    for row in tbl.rows[1:]:
        row.height = Cm(0.95)
    add_heading(doc, "Como o funcionário abre a gaveta dele", level=2)
    campo(doc, "No Windows: ________________________________________________________________________", size=10)
    campo(doc, "No celular: _________________________________________________________________________", size=10)
    add_para(doc, "Na etiqueta da máquina (fita crepe): nome do servidor, endereço e quem instalou. Nenhuma senha.", italic=True, size=9, color=COR_CINZA)

    # PÁGINA 4 — Laudo, conserto, apresentação, autoavaliação
    doc.add_page_break()
    cabecalho_documento(
        doc,
        "CONFERÊNCIA CRUZADA E ENTREGA",
        "Pasta na rede e entrega do servidor · Encontro 3 · 4 de 4",
    )

    add_heading(doc, "1  Laudo do servidor da equipe ____ (só com a ficha deles)", level=2)
    tbl = add_table(doc, [["Conferido", "✓ ou ✗", "Se ✗: o que aconteceu / a mensagem"]] + [[l, "", ""] for l in LAUDO],
                    col_widths=[8, 1.6, 7.8])
    small_table(tbl, size=9)
    for row in tbl.rows[1:]:
        row.height = Cm(0.7)
    campo(doc, "Recado para a equipe dona: o melhor da ficha foi _______________________ e faltou _______________________", size=10)

    add_heading(doc, "2  O laudo que recebemos: o que consertamos", level=2)
    linha(doc)
    linha(doc)

    add_heading(doc, "3  Apresentação (4 minutos, cada um fala pelo menos uma)", level=2)
    tbl = add_table(doc, [["Pergunta", "Quem fala"],
                          ["O que este servidor faz pelo escritório? (sem comando)", ""],
                          ["Como a ana usa no dia a dia? (o Z: e o site)", ""],
                          ["O que o laudo encontrou e como consertamos?", ""],
                          ["O que o cliente precisa saber para não estragar?", ""]],
                    col_widths=[12.4, 5])
    small_table(tbl, size=9.5)

    add_heading(doc, "4  Autoavaliação", level=2)
    tbl = add_table(doc, [["No módulo de servidores, eu…", "Sozinho", "Com ajuda", "Ainda não"],
                          ["instalo o Ubuntu Server numa máquina", "☐", "☐", "☐"],
                          ["entro de longe e ando pelas pastas", "☐", "☐", "☐"],
                          ["crio usuários, grupos e permissões", "☐", "☐", "☐"],
                          ["publico um site e uma pasta na rede", "☐", "☐", "☐"],
                          ["escrevo uma ficha que outro técnico usa", "☐", "☐", "☐"]],
                    col_widths=[10.4, 2.3, 2.4, 2.3])
    small_table(tbl, size=9.5)
    campo(doc, "O que foi mais difícil, e o que eu faria para um cliente de verdade? ______________________________________", size=10)
    add_para(doc, "Desligar sempre com sudo poweroff · CEJA Itapiranga – SC", italic=True, size=9, color=COR_CINZA, align=WD_ALIGN_PARAGRAPH.CENTER)

    path = os.path.join(OUTDIR, f"Folha do Aluno - {TITULO_ARQ}.docx")
    doc.save(path)
    return path


# ============================================================
# RESUMO DO ALUNO — o módulo inteiro
# ============================================================
def build_resumo():
    doc = Document()
    setup_page(doc, margem=Cm(1.8))

    cabecalho_documento(
        doc,
        "O SERVIDOR DO ESCRITÓRIO, DO COMEÇO AO FIM",
        "Resumo do aluno · Servidores com Linux · Encontro 3 e o módulo inteiro · Curso FIC · CEJA Itapiranga",
    )

    add_heading(doc, "A ideia principal", level=2)
    add_para(doc, "Em três encontros, um computador antigo virou o servidor de um escritório: tem um site interno, guarda os arquivos de cada setor numa pasta de rede, cada funcionário entra com o próprio nome e só abre a sua gaveta, e o técnico administra tudo de longe. Servidor entregue é o que outra pessoa consegue usar só com a ficha técnica.")

    add_heading(doc, "O caminho do módulo", level=2)
    add_table(doc, [
        ["Encontro", "O que foi feito", "Comandos e ferramentas"],
        ["1", "Instalar o Ubuntu Server, raio-x da máquina, primeiro site", "pendrive, ip a, ping, free -h, df -h, apt, Nginx, nano"],
        ["2", "Bancada, endereço fixo, acesso de longe, pastas, usuários, permissões, site com IA", "PuTTY, reserva no roteador, cd, ls, mkdir, cp, mv, rm, adduser, groupadd, usermod, chgrp, chmod"],
        ["3", "Pasta na rede para Windows e celular, ficha técnica, entrega", "Samba, smbpasswd, testparm, \\\\endereço, unidade Z:, smbstatus"],
    ], col_widths=[2, 7, 7.5])

    add_heading(doc, "A pasta na rede (Samba)", level=2)
    add_bullet(doc, "O Samba faz uma pasta do Linux aparecer no Windows, no Mac e no celular (a língua SMB)")
    add_bullet(doc, "Antes de mexer num arquivo de configuração, guarde uma cópia do original")
    add_bullet(doc, "Cada gaveta tem: o nome que aparece na rede, a pasta do servidor (path), quem entra (valid users, @ = grupo), se pode alterar (read only) e de quem ficam os arquivos novos (force group, create mask)")
    add_bullet(doc, "testparm confere a configuração; sudo systemctl restart smbd liga com a nova")
    add_bullet(doc, "Cada funcionário precisa de senha de rede: sudo smbpasswd -a nome (use a mesma do Linux)")
    add_bullet(doc, "sudo smbstatus mostra quem está conectado agora")

    add_heading(doc, "Do lado do funcionário", level=2)
    add_bullet(doc, "Windows: Explorador → \\\\endereço do servidor → a gaveta → usuário e senha")
    add_bullet(doc, "Unidade Z: botão direito em Este Computador → Mapear unidade de rede → \\\\endereço\\gaveta → Reconectar ao entrar")
    add_bullet(doc, "Trocar de usuário no mesmo PC: fechar o Explorador e, no cmd, net use * /delete /y")
    add_bullet(doc, "iPhone: Arquivos → Conectar ao Servidor → smb://endereço · Android: CX Explorador → Rede → SMB")

    add_heading(doc, "A ficha técnica tem que ter", level=2)
    add_para(doc, "Máquina, nome e endereço fixo do servidor, MAC, roteador e Wi-Fi, administrador e senha, como entrar de longe, o site, a pasta na rede, as gavetas e quem entra em cada uma, os funcionários e senhas, como desligar, quem instalou e quando. Fica guardada com o cliente. Na etiqueta da máquina, nenhuma senha.", size=10.5)

    add_heading(doc, "Cuidados de quem cuida de servidor", level=2)
    add_bullet(doc, "Servidor tem endereço fixo e se desliga com sudo poweroff, nunca na tomada")
    add_bullet(doc, "Senha longa, um usuário por pessoa, funcionário não é administrador")
    add_bullet(doc, "Nunca chmod 777. Descubra quem precisa e ponha no grupo certo")
    add_bullet(doc, "Pasta de rede não se abre para a internet. De fora, só com VPN")
    add_bullet(doc, "Servidor sem cópia de segurança perde tudo quando o disco queima")

    add_heading(doc, "Três frases pra guardar", level=2)
    add_para(doc, "\"Servidor entregue é o que outro usa só com a ficha.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Antes de mexer na configuração, guarde o original.\"", bold=True, color=COR_PRIMARIA)
    add_para(doc, "\"Cada setor na sua gaveta. O mural, todos leem.\"", bold=True, color=COR_PRIMARIA)

    add_heading(doc, "Pra se testar em casa", level=2)
    add_bullet(doc, "Um escritório de 4 pessoas guarda tudo num pendrive que passa de mão em mão. Que problemas isso causa, e o que você ofereceria?")
    add_bullet(doc, "Na gaveta do financeiro, o bruno não consegue alterar o que a ana salvou. Qual linha da configuração você conferiria primeiro?")
    add_bullet(doc, "Quanto você cobraria para instalar esse servidor? E pela manutenção de cada mês?")

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
