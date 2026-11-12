"""Apoio do professor: passo a passo para configurar TP-Link Archer C5 e Archer C50.

Não é slide nem folha do aluno: é o papel que o professor segura e vai passando
para a turma, tela por tela. Segue a mesma ordem de serviço da Noite 8.

Gera: curso/Apoio - Redes - Passo a passo TP-Link Archer C5 e C50.docx
"""
import os

from docx import Document
from docx.shared import Pt, Cm

from common import (
    COR_PRIMARIA, COR_ALERTA, COR_CINZA,
    add_heading, add_para, add_bullet, add_table, add_hr,
    setup_page, cabecalho_documento, set_cell_shading,
)

SAIDA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "Apoio - Redes - Passo a passo TP-Link Archer C5 e C50.docx")


def caixa(doc, titulo, linhas, fundo="F4F7F7", cor_titulo=None):
    """Caixa de destaque (tabela de 1 célula com fundo)."""
    t = doc.add_table(rows=1, cols=1)
    cell = t.rows[0].cells[0]
    set_cell_shading(cell, fundo)
    p = cell.paragraphs[0]
    r = p.add_run(titulo)
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = cor_titulo or COR_PRIMARIA
    for ln in linhas:
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(ln)
        r.font.size = Pt(10)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def passo(doc, n, titulo, caminho_azul=None, caminho_verde=None):
    add_heading(doc, f"Passo {n} — {titulo}", level=2)
    if caminho_azul or caminho_verde:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        if caminho_azul:
            r = p.add_run("Tela azul: ")
            r.bold = True
            r.font.size = Pt(10)
            r = p.add_run(caminho_azul)
            r.font.size = Pt(10)
            r.font.color.rgb = COR_PRIMARIA
        if caminho_azul and caminho_verde:
            p.add_run("     ")
        if caminho_verde:
            r = p.add_run("Tela verde: ")
            r.bold = True
            r.font.size = Pt(10)
            r = p.add_run(caminho_verde)
            r.font.size = Pt(10)
            r.font.color.rgb = COR_CINZA


def num(doc, itens):
    for i, txt in enumerate(itens, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{i}.  ")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = COR_PRIMARIA
        r = p.add_run(txt)
        r.font.size = Pt(11)


def build():
    doc = Document()
    setup_page(doc, margem=Cm(1.8))
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)

    cabecalho_documento(
        doc,
        "Passo a passo — Configurando o TP-Link Archer C5 e Archer C50",
        "Apoio do professor · Redes de Computadores · FIC Montador e Reparador de Computadores · CEJA Itapiranga",
    )

    add_para(doc, "Roteiro pra ir passando pra turma, tela por tela. A ordem é a mesma do serviço na casa do "
                  "cliente: entrar, trancar, dar nome, organizar os endereços e guardar a configuração. Os dois "
                  "modelos são Dual Band (2,4 GHz e 5 GHz) e, nas versões novas, a tela é praticamente a mesma. "
                  "O que muda de um pro outro está marcado.")

    # ------------------------------------------------------------------
    add_heading(doc, "0. Antes de ligar: qual tela vai aparecer?")
    add_para(doc, "Vire o roteador e leia a etiqueta: tem o modelo (Archer C5 / Archer C50) e a versão "
                  "(Ver: 1.0, 2.0, 4.0, 6.0…). A versão decide qual painel vai abrir.")
    add_table(doc, [
        ["", "Tela AZUL (firmware novo)", "Tela VERDE (firmware antigo)"],
        ["Quais aparelhos", "Archer C50 (v3 em diante) · Archer C5 v4 em diante", "Archer C5 v1.2 e v2"],
        ["Endereço", "192.168.0.1 ou tplinkwifi.net", "192.168.0.1 ou tplinkwifi.net"],
        ["Primeiro acesso", "Pede pra CRIAR a senha de administrador", "Usuário admin / senha admin"],
        ["Menu", "Abas no alto: Básico · Avançado. Menu à esquerda dentro de Avançado",
         "Menu comprido à esquerda, sem abas"],
        ["Wi-Fi 2,4 e 5 GHz", "Na mesma tela, uma aba pra cada faixa", "Dois menus separados: Sem fio 2,4 GHz e Sem fio 5 GHz"],
    ], col_widths=[3.2, 7.2, 7.2])
    add_para(doc, "Os roteadores da bancada pedem pra criar a senha no primeiro acesso: são tela azul. Os caminhos "
                  "da tela verde vão junto em cada passo, pra quando aparecer um C5 antigo na casa do cliente.",
             size=10, italic=True, color=COR_CINZA)
    add_para(doc, "Os nomes dos menus podem vir em português ou em inglês, conforme o firmware. Aqui vai o nome em "
                  "português com o inglês entre parênteses.", size=10, italic=True, color=COR_CINZA)

    caixa(doc, "Ligações físicas (as duas versões)", [
        "•  Fonte: 12 V. Não use a fonte de 9 V do TL-WR840N.",
        "•  Porta WAN / Internet (cor diferente, geralmente azul): cabo vindo da ONU ou do modem do provedor.",
        "•  Portas LAN (amarelas, numeradas 1 a 4): computador, TV, impressora.",
        "•  Na bancada, sem internet, a WAN pode ficar vazia. Dá pra configurar tudo do mesmo jeito.",
    ])

    # ------------------------------------------------------------------
    add_heading(doc, "1. Entrar no roteador")

    passo(doc, 1, "Começar do zero (se o roteador veio usado)")
    num(doc, [
        "Roteador ligado. Aguarde 1 minuto pra ele terminar de iniciar.",
        "Com a ponta de uma caneta, segure o botão RESET por uns 10 segundos, até as luzes piscarem juntas. "
        "No C50 e no C5 v4 o botão costuma ser um só, escrito WPS/RESET: apertar rápido liga o WPS, segurar reseta.",
        "Solte e espere 1 a 2 minutos. Ele volta com o nome e a senha de Wi-Fi da etiqueta.",
    ])

    passo(doc, 2, "Conectar e abrir o painel")
    num(doc, [
        "Conecte no roteador: cabo na porta LAN amarela ou Wi-Fi com o nome e a senha da etiqueta "
        "(SSID e Wireless Password/PIN).",
        "No celular: desligue os dados móveis. Se o celular avisar \"rede sem internet\", escolha manter conectado.",
        "No navegador, digite na barra de endereço (não na busca do Google): 192.168.0.1 ou tplinkwifi.net",
    ])
    caixa(doc, "Não abriu?", [
        "•  Confira se está mesmo no Wi-Fi do roteador e não no da escola.",
        "•  No computador: cmd → ipconfig → olhe o Gateway Padrão. O painel está nesse número.",
        "•  Se alguém já mudou o endereço (Noite 8: 192.168.1X.1), use o gateway. Se ninguém sabe a senha, reset (Passo 1).",
    ], fundo="FFF2EF", cor_titulo=COR_ALERTA)

    passo(doc, 3, "Senha de administrador", "System Tools → Administração (Administration)",
          "Ferramentas do sistema → Senha (System Tools → Password)")
    num(doc, [
        "Tela azul: no primeiro acesso aparece \"Criar uma senha\" (Create a password). Digite duas vezes e clique "
        "em Vamos começar (Let's Get Started).",
        "Tela verde: entre com admin / admin e depois troque em Ferramentas do sistema → Senha.",
        "A equipe escolhe a senha (no mínimo 10 caracteres) e anota na ficha. Na casa do cliente ela vai na ficha "
        "de entrega, nunca na etiqueta do roteador.",
    ])
    caixa(doc, "Fale pra turma", [
        "São duas senhas diferentes: a do Wi-Fi (pra usar a internet) e a de administrador (pra mexer no roteador). "
        "Quem tem a de administrador manda no aparelho.",
    ])

    passo(doc, 4, "Assistente de configuração rápida (Quick Setup)")
    add_para(doc, "Logo depois da senha, a tela azul abre o assistente. Na bancada dá pra pular (botão Pular / Skip "
                  "ou clicar na aba Avançado) e fazer item por item, que ensina mais. Na casa do cliente ele "
                  "resolve o básico em 2 minutos:")
    num(doc, [
        "Fuso horário: (GMT-03:00) Brasília.",
        "Tipo de conexão com a internet: veja o quadro abaixo. Se não souber, clique em Detectar (Auto Detect).",
        "Clonagem de MAC: deixe Não clonar (Do NOT clone), a não ser que o provedor peça.",
        "Sem fio: nome e senha da rede de 2,4 GHz e da de 5 GHz (detalhes no Passo 5).",
        "Confira o resumo e clique em Salvar. Ele reinicia o Wi-Fi: reconecte com o nome e a senha novos.",
    ])
    add_table(doc, [
        ["Situação na casa", "Tipo de conexão", "O que precisa"],
        ["ONU/modem do provedor já é roteador (tem Wi-Fi, entrega IP)", "IP Dinâmico (Dynamic IP)", "Nada. É o caso mais comum"],
        ["ONU em modo bridge (só converte a fibra)", "PPPoE", "Usuário e senha do provedor (pedir pro provedor ou ver no contrato)"],
        ["Provedor deu um IP fixo", "IP Estático (Static IP)", "IP, máscara, gateway e DNS que o provedor passou"],
    ], col_widths=[6.6, 4.4, 6.6])
    add_para(doc, "Tela verde: Rede → WAN (Network → WAN), mesmo raciocínio.", size=10, italic=True, color=COR_CINZA)

    # ------------------------------------------------------------------
    add_heading(doc, "2. Wi-Fi: nome, senha e tranca")

    passo(doc, 5, "Nome e senha das duas redes",
          "Avançado → Sem fio → Configurações sem fio (Advanced → Wireless → Wireless Settings)",
          "Sem fio 2,4 GHz → Configurações e Segurança; repetir em Sem fio 5 GHz")
    add_para(doc, "Na tela azul tem uma aba 2,4 GHz e outra 5 GHz. Faça as duas.")
    num(doc, [
        "Habilitar rádio sem fio (Enable Wireless Radio): marcado.",
        "Nome da rede (SSID): nome que não identifica a família nem o endereço. Sugestão pra turma: "
        "EQUIPE1 na 2,4 GHz e EQUIPE1_5G na 5 GHz (nomes diferentes deixam escolher a faixa, e isso vai ser usado na próxima aula).",
        "Ocultar SSID (Hide SSID): desmarcado. Esconder o nome não protege nada e só dá trabalho pra conectar.",
        "Segurança: WPA/WPA2-Pessoal (WPA/WPA2-Personal) → Versão: WPA2-PSK → Criptografia: AES.",
        "Senha: no mínimo 10 caracteres, misturando letras e números. Pode ser frase: casaverde2026.",
        "Canal e largura de canal: deixe Automático (Auto) por enquanto.",
        "Salvar. O Wi-Fi cai. Todo mundo reconecta com o nome e a senha novos e abre o painel de novo.",
    ])
    caixa(doc, "Tela verde (C5 antigo)", [
        "Primeiro veja Seleção de banda dupla (Dual Band Selection): deixe as duas faixas ligadas.",
        "Nome em Sem fio 2,4 GHz → Configurações sem fio. Senha em Sem fio 2,4 GHz → Segurança sem fio "
        "(WPA/WPA2 - Personal, WPA2-PSK, AES, Senha PSK). Depois repita tudo no Sem fio 5 GHz. "
        "Alguns salvamentos pedem pra reiniciar: clique no link de reinício que aparece em vermelho.",
    ])

    passo(doc, 6, "Desligar o WPS", "Avançado → Sem fio → WPS",
          "WPS (no menu da esquerda)")
    num(doc, [
        "Desative o PIN do roteador (Router's PIN) e o WPS.",
        "Por quê: o PIN do WPS tem só 8 dígitos e dá pra descobrir por tentativa. Com ele ligado, a senha forte do Wi-Fi não adianta.",
    ])

    # ------------------------------------------------------------------
    add_heading(doc, "3. Endereços: painel, DHCP e DNS")

    passo(doc, 7, "Mudar o endereço do painel (LAN)", "Avançado → Rede → LAN (Advanced → Network → LAN)",
          "Rede → LAN (Network → LAN)")
    num(doc, [
        "Endereço IP: troque 192.168.0.1 pelo da equipe (equipe 1 = 192.168.11.1, equipe 2 = 192.168.12.1…). "
        "Máscara: 255.255.255.0.",
        "Salvar. O roteador reinicia (1 minuto). Reconecte no Wi-Fi.",
        "O endereço antigo não abre mais. Descubra o novo pelo gateway (ipconfig ou detalhes do Wi-Fi no celular) e entre por ele.",
    ])
    caixa(doc, "Por que mudar", [
        "A ONU do provedor muitas vezes já usa 192.168.0.1 ou 192.168.1.1. Se o roteador usar o mesmo número, dá "
        "conflito: tudo parece ligado e nada funciona.",
    ])

    passo(doc, 8, "DHCP: faixa de endereços e DNS",
          "Avançado → Rede → Servidor DHCP (Advanced → Network → DHCP Server)",
          "DHCP → Configurações DHCP (DHCP → DHCP Settings)")
    num(doc, [
        "Servidor DHCP: Habilitado (Enable).",
        "Faixa de endereços (IP Address Pool): na tela azul ela muda sozinha junto com a LAN. Ajuste pra "
        "192.168.1X.100 até 192.168.1X.150.",
        "Tempo de concessão (Lease Time): pode deixar o padrão (120 minutos).",
        "Gateway padrão: o endereço do painel (192.168.1X.1). Geralmente já vem preenchido.",
        "DNS primário: 8.8.8.8 · DNS secundário: 1.1.1.1.",
        "Salvar. Pra conferir, no computador: desconectar e conectar de novo, cmd → ipconfig /all → "
        "o IP tem que estar na faixa e o Servidor DNS tem que ser 8.8.8.8.",
    ])

    passo(doc, 9, "Reserva de endereço por MAC",
          "Avançado → Rede → Servidor DHCP → Reserva de endereço → Adicionar",
          "DHCP → Reserva de endereço (Address Reservation) → Adicionar novo")
    num(doc, [
        "Na mesma tela do DHCP, veja a Lista de clientes (DHCP Client List) e ache o aparelho pelo nome.",
        "Na tela azul dá pra clicar em Ver clientes conectados (View Existing Devices) e escolher da lista: o MAC vem preenchido.",
        "Endereço IP reservado: 192.168.1X.50 (fora da faixa do DHCP). Marque Habilitar e salve.",
        "Desconecte e conecte o aparelho. Ele tem que voltar com o .50.",
    ])
    caixa(doc, "Pegadinha do celular", [
        "Android e iPhone usam MAC aleatório por rede (\"Endereço MAC privado\" ou \"MAC aleatório\"). Se a reserva "
        "não pegar, entre nas configurações daquela rede Wi-Fi no celular, troque para MAC do aparelho e reconecte. "
        "Pra impressora, câmera e TV não tem esse problema: é nelas que a reserva mais se usa.",
    ], fundo="FFF2EF", cor_titulo=COR_ALERTA)

    # ------------------------------------------------------------------
    add_heading(doc, "4. Rede de convidados")

    passo(doc, 10, "Criar a rede de convidados", "Avançado → Rede de convidados (Advanced → Guest Network)",
          "Rede de convidados (Guest Network), no menu da esquerda")
    num(doc, [
        "Habilite na 2,4 GHz (e na 5 GHz, se quiser).",
        "Nome: EQUIPE1_VISITA. Segurança WPA2-PSK, outra senha, diferente da rede principal.",
        "Desmarque \"Permitir que convidados acessem minha rede local\" (Allow guests to access my local network). "
        "Desmarque também \"permitir que os convidados vejam uns aos outros\" (see each other).",
        "Salvar. Teste: conecte um celular na rede de visita e tente abrir o painel 192.168.1X.1. Não pode abrir.",
    ])
    add_para(doc, "Serve pra visita, cliente do comércio, funcionário: usa a internet e não enxerga as câmeras, a "
                  "impressora nem o painel.", size=10, italic=True, color=COR_CINZA)

    # ------------------------------------------------------------------
    add_heading(doc, "5. Manutenção: horário, firmware, cópia e restauração")

    passo(doc, 11, "Horário e firmware (rápido)",
          "Avançado → Ferramentas do sistema → Configurações de hora / Atualização de firmware",
          "Ferramentas do sistema → Configurações de hora / Atualização de firmware")
    num(doc, [
        "Hora: fuso (GMT-03:00) Brasília e Obter do computador/da internet. Hora errada bagunça os registros (logs) e o controle dos pais.",
        "Firmware: veja a versão instalada. Atualização só com o arquivo baixado do site da TP-Link pra mesma "
        "versão de hardware da etiqueta (Ver: X.0). Arquivo da versão errada pode estragar o roteador. "
        "Nunca desligar no meio. Na bancada: só olhar, não atualizar.",
    ])

    passo(doc, 12, "Cópia de segurança (backup)",
          "Avançado → Ferramentas do sistema → Backup e restauração → Backup",
          "Ferramentas do sistema → Backup e restauração (Backup & Restore) → Backup")
    num(doc, [
        "Clique em Backup. Baixa um arquivo (config.bin).",
        "Renomeie: EQUIPE1_ArcherC5_DDMM.bin e guarde no computador ou no pen drive da turma.",
        "Na casa do cliente: guardar com a ficha. Se um dia o roteador resetar, você volta tudo em 1 minuto.",
    ])

    passo(doc, 13, "Restaurar (o teste de verdade)",
          "Backup e restauração → Restaurar (Restore)",
          "Backup e restauração → Restaurar")
    num(doc, [
        "Reset no roteador (Passo 1). Ele volta para 192.168.0.1 e Wi-Fi da etiqueta.",
        "Conecte, abra 192.168.0.1, crie uma senha de administrador qualquer (a tela azul exige) e pule o assistente.",
        "Restaurar → Procurar/Escolher arquivo → o .bin da equipe → Restaurar. Não desligue: ele reinicia sozinho.",
        "Reconecte no Wi-Fi da EQUIPE e entre em 192.168.1X.1 com a senha de administrador da equipe, a original e não a provisória.",
    ])
    caixa(doc, "Cuidado", [
        "O backup só volta em roteador do mesmo modelo e versão. Arquivo do C5 não entra no C50, e arquivo do C5 v2 não entra no C5 v4.",
    ], fundo="FFF2EF", cor_titulo=COR_ALERTA)

    # ------------------------------------------------------------------
    add_heading(doc, "6. Conferência antes de entregar")
    add_table(doc, [
        ["Conferir", "Como", "OK"],
        ["Senha de administrador trocada", "Sair e entrar de novo no painel", "☐"],
        ["Wi-Fi 2,4 e 5 GHz com nome e senha novos", "Conectar um celular em cada", "☐"],
        ["WPA2-PSK / AES nas duas faixas", "Olhar a tela de segurança", "☐"],
        ["WPS desligado", "Tela do WPS", "☐"],
        ["Painel no endereço novo", "Abrir 192.168.1X.1", "☐"],
        ["DHCP na faixa .100–.150 e DNS 8.8.8.8", "ipconfig /all no computador", "☐"],
        ["Reserva funcionando", "Aparelho reconecta e volta com o .50", "☐"],
        ["Rede de convidados isolada", "Pela visita, o painel não abre", "☐"],
        ["Backup salvo e testado", "Restaurou depois do reset", "☐"],
        ["Ficha de entrega preenchida", "Outra pessoa consegue usar só com o papel", "☐"],
    ], col_widths=[7.0, 8.6, 1.6])

    # ------------------------------------------------------------------
    add_heading(doc, "7. Deu problema? (os mais comuns na bancada)")
    add_table(doc, [
        ["O que aparece", "Causa provável", "O que fazer"],
        ["Painel não abre depois de salvar a LAN", "O endereço mudou", "Ver o gateway e entrar pelo número novo"],
        ["Celular não volta pro Wi-Fi depois de trocar a senha", "Celular guardou a senha antiga", "Esquecer a rede e conectar de novo"],
        ["Abre a busca do Google em vez do painel", "Digitou na caixa de busca", "Digitar o número na barra de endereço, com http:// na frente se precisar"],
        ["Conectado, mas \"sem internet\"", "Na bancada é normal: WAN sem cabo", "Seguir configurando. Na casa do cliente: ver o tipo de conexão (Passo 4)"],
        ["Ninguém lembra a senha de administrador", "Não anotaram", "Reset (Passo 1) e restaurar o backup (Passo 13)"],
        ["Reserva não pega no celular", "MAC aleatório", "Trocar pra MAC do aparelho naquela rede"],
        ["Não aparece a rede de 5 GHz", "Aparelho antigo (só 2,4 GHz) ou rádio 5 GHz desligado", "Testar com outro celular e conferir o Passo 5"],
        ["Luzes piscando, nada responde", "Está reiniciando ou atualizando", "Esperar 2 minutos e não desligar da tomada"],
    ], col_widths=[5.6, 5.4, 6.6])

    add_hr(doc)
    add_para(doc, "Os nomes dos menus variam um pouco entre versões de firmware. Se um item não estiver onde este "
                  "papel diz, procure pelo assunto (Sem fio, Rede, DHCP, Ferramentas do sistema): a lógica é sempre a mesma.",
             size=9, italic=True, color=COR_CINZA)

    doc.save(SAIDA)
    print("OK:", SAIDA)


if __name__ == "__main__":
    build()
