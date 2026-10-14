"""Gera o Plano de Aula Mensal do curso de Redes — 15/10/2026 a 14/11/2026.

Cobre as Noites 7 a 11 (uma aula por semana, 210 min cada), no padrão
prática-pesada: 50%+ do tempo em atividade do aluno, sem role-play e sem duelo.
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.shared import Pt, Cm

from common import (
    COR_PRIMARIA, COR_CINZA, COR_DESTAQUE,
    add_heading, add_para, add_bullet, add_table, add_hr,
    setup_page, cabecalho_documento,
)

OUTDIR = "/home/nedel/curso"
PERIODO = "15/10/2026 a 14/11/2026"


IDENTIFICACAO = [
    ["Campo", "Informação"],
    ["Instituição", "CEJA — Centro de Educação de Jovens e Adultos · Itapiranga – SC"],
    ["Curso", "FIC — Montador e Reparador de Computadores"],
    ["Componente / Módulo", "Redes de Computadores"],
    ["Professor", "Joni Nedel"],
    ["Turno", "Noturno · 1 encontro por semana"],
    ["Duração do encontro", "210 min (3h30)"],
    ["Período deste plano", PERIODO],
    ["Encontros previstos", "5 (Noites 7, 8, 9, 10 e 11 do módulo)"],
    ["Carga horária do período", "17h30 (5 × 3h30)"],
    ["Público", "Jovens e adultos da EJA"],
]

CRONOGRAMA = [
    ["Semana", "Encontro", "Tema", "Atividade central do aluno"],
    ["15/10 a 21/10", "Noite 7", "Nomes na rede: DNS, URL e o que o HTTPS resolve",
     "Agenda viva (DNS encenado), clínica de 16 links, estações, cartilha anti-golpe"],
    ["22/10 a 28/10", "Noite 8", "Rede local em casa I — planejando a instalação",
     "Levantamento da planta da casa, lista de equipamentos e orçamento"],
    ["29/10 a 04/11", "Noite 9", "Rede local em casa II — configurando o roteador",
     "Configuração completa de um roteador de fábrica em bancada"],
    ["05/11 a 11/11", "Noite 10", "Cobertura de Wi-Fi: posição, faixas e vizinhos",
     "Mapa de sinal da escola com o celular e reposicionamento do roteador"],
    ["12/11 a 14/11*", "Noite 11", "Segurança para o técnico: senhas, HTTPS e golpes com Wi-Fi",
     "Auditoria de segurança do roteador e da cartilha da Noite 7"],
]

# Cada aula: título, objetivos, conteúdos, metodologia, recursos, avaliação
AULAS = [
    {
        "titulo": "NOITE 7 — NOMES NA REDE: DNS, URL E O QUE O HTTPS RESOLVE",
        "objetivos": [
            "Explicar com as próprias palavras o que o DNS faz (traduz nome em número de IP)",
            "Ler um endereço de site e identificar quem é o dono do link (domínio principal)",
            "Diferenciar o que o cadeado (HTTPS) garante e o que ele não garante",
            "Reconhecer links falsos e os truques de mensagens de golpe",
        ],
        "conteudos": [
            "DNS: a agenda de contatos da internet",
            "Partes de uma URL: protocolo, subdomínio, domínio, extensão, caminho",
            "HTTP × HTTPS: conversa lacrada, certificado e seus limites",
            "Golpes por link: nomes parecidos, encurtadores, pressa, medo e prêmio",
        ],
        "metodologia": [
            "Retomada com censo dos endereços trazidos de casa (tarefa da Noite 6)",
            "Exposição curta em três blocos de 10–15 min, cada um seguido de prática",
            "Agenda viva: encenação do DNS com placas e cartões de pedido",
            "Clínica de links em equipe: 16 cartões classificados em verdadeiro × golpe",
            "Estações rotativas: ping por nome × por número, DNS sabotado, cadeado no celular",
            "Fábrica de golpes e cartilha anti-golpe para a família (produção em equipe)",
        ],
        "recursos": "Quadro, projetor, slides, folha do aluno, anexos impressos (agenda, placas, cartões), "
                     "computador com hotspot do celular, celulares dos alunos, cartolina e canetões.",
        "avaliacao": "Correção coletiva da clínica de links, laudo das estações, qualidade da cartilha "
                     "e quiz de saída. Tarefa: desenhar a planta da própria casa para a Noite 8.",
    },
    {
        "titulo": "NOITE 8 — REDE LOCAL EM CASA I: PLANEJANDO A INSTALAÇÃO",
        "objetivos": [
            "Levantar as necessidades de uma residência antes de instalar a rede",
            "Escolher a posição do roteador a partir da planta da casa",
            "Definir onde usar cabo e onde usar Wi-Fi",
            "Montar uma lista de materiais com orçamento",
        ],
        "conteudos": [
            "Visita técnica: o que perguntar e o que observar no cliente",
            "Planta baixa simplificada: cômodos, paredes, ponto de entrada da internet",
            "Critérios de posicionamento: centro da casa, altura, obstáculos",
            "Cabo × Wi-Fi: quando usar cada um (retoma Noites 3 e 4)",
            "Lista de materiais e orçamento (retoma a planilha da Noite 3)",
        ],
        "metodologia": [
            "Galeria das plantas trazidas de casa: cada aluno apresenta a sua",
            "Exposição curta sobre visita técnica e critérios de posicionamento",
            "Projeto em dupla: escolher uma planta e marcar roteador, pontos de cabo e áreas sem sinal",
            "Medição de cabo na planta com escala simples e cálculo de metragem",
            "Orçamento da instalação na planilha modelo e apresentação para a turma",
        ],
        "recursos": "Plantas dos alunos, folha do aluno com malha quadriculada, réguas, lápis de cor, "
                     "planilha modelo de orçamento (.ods), slides, projetor.",
        "avaliacao": "Projeto da dupla (planta marcada + lista de materiais + orçamento) e justificativa "
                     "das escolhas na apresentação.",
    },
    {
        "titulo": "NOITE 9 — REDE LOCAL EM CASA II: CONFIGURANDO O ROTEADOR",
        "objetivos": [
            "Acessar o painel de um roteador de fábrica e fazer a configuração inicial completa",
            "Configurar nome de rede, senha forte e tipo de segurança (WPA2/WPA3)",
            "Trocar a senha de administrador e conferir DHCP e DNS",
            "Registrar a configuração em uma ficha de entrega para o cliente",
        ],
        "conteudos": [
            "Reset de fábrica e primeiro acesso (retoma a Noite 6)",
            "Modo de conexão da internet: roteador atrás do modem/ONU",
            "SSID, senha, WPA2/WPA3, rede de convidados",
            "DHCP (faixa de endereços) e DNS do roteador (retoma a Noite 7)",
            "Atualização de firmware e backup da configuração",
        ],
        "metodologia": [
            "Demonstração curta no projetor com um roteador",
            "Bancada em equipe: cada equipe reseta e configura o próprio roteador seguindo checklist",
            "Teste de conexão com os celulares da equipe",
            "Troca de bancada: uma equipe confere a configuração da outra com o checklist",
            "Preenchimento da ficha de entrega para o cliente",
        ],
        "recursos": "Roteadores de bancada (TL-WR840N, Archer C5), cabos de rede, régua de tomadas, "
                     "computador, celulares, checklist impresso, slides.",
        "avaliacao": "Roteador configurado e conferido pela equipe vizinha; ficha de entrega preenchida "
                     "corretamente.",
    },
    {
        "titulo": "NOITE 10 — COBERTURA DE WI-FI: POSICIONAMENTO, FAIXAS E VIZINHOS",
        "objetivos": [
            "Medir a intensidade do sinal Wi-Fi com o celular",
            "Diferenciar as faixas de 2,4 GHz e 5 GHz (alcance × velocidade)",
            "Identificar interferência de redes vizinhas e escolher canal",
            "Propor soluções para áreas sem sinal: reposicionar, repetidor, mesh, cabo",
        ],
        "conteudos": [
            "Intensidade de sinal (dBm) e o que é sinal bom, médio e ruim",
            "2,4 GHz × 5 GHz: alcance, velocidade e paredes",
            "Canais e redes vizinhas",
            "Repetidor, roteador mesh e ponto de acesso cabeado",
        ],
        "metodologia": [
            "Exposição curta com o resultado da planta da Noite 8",
            "Mapa de sinal: equipes medem o sinal em vários pontos da escola com app no celular",
            "Experimento: mudar a posição e a faixa do roteador e medir de novo",
            "Varredura de redes vizinhas e escolha do melhor canal",
            "Cartaz por equipe com diagnóstico e solução proposta, apresentado à turma",
        ],
        "recursos": "Roteadores de bancada, celulares com app de análise de Wi-Fi, planta da escola impressa, "
                     "cartolina, canetões, slides.",
        "avaliacao": "Mapa de sinal preenchido, conclusão do experimento e cartaz com diagnóstico e solução.",
    },
    {
        "titulo": "NOITE 11 — SEGURANÇA PARA O TÉCNICO: SENHAS, HTTPS E GOLPES COM WI-FI",
        "objetivos": [
            "Criar e avaliar senhas fortes para Wi-Fi e para o painel do roteador",
            "Reconhecer riscos de redes Wi-Fi abertas e redes falsas",
            "Fazer uma auditoria básica de segurança em um roteador",
            "Orientar o cliente sobre cuidados de segurança",
        ],
        "conteudos": [
            "Senhas: tamanho, frases-senha, reaproveitamento",
            "WPS, senha de fábrica e firmware desatualizado",
            "Wi-Fi público, rede falsa com nome parecido, HTTPS (retoma a Noite 7)",
            "Responsabilidade do técnico: o que entregar e o que explicar ao cliente",
        ],
        "metodologia": [
            "Retomada das cartilhas anti-golpe feitas na Noite 7",
            "Exposição curta sobre os riscos mais comuns",
            "Oficina de senhas: equipes criam e testam frases-senha",
            "Auditoria em bancada: cada equipe recebe um roteador com falhas plantadas, encontra e corrige",
            "Revisão da cartilha da Noite 7 com a parte de Wi-Fi",
        ],
        "recursos": "Roteadores de bancada com configurações inseguras, checklist de auditoria, cartilhas da "
                     "Noite 7, celulares, slides.",
        "avaliacao": "Laudo de auditoria (falhas encontradas e corrigidas) e cartilha revisada.",
    },
]


def gerar():
    doc = Document()
    setup_page(doc)
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    cabecalho_documento(doc, "Plano de Aula Mensal",
                        f"Redes de Computadores · FIC Montador e Reparador de Computadores · {PERIODO}")

    add_heading(doc, "1. Identificação")
    add_table(doc, IDENTIFICACAO, col_widths=[5.0, 12.0])

    add_heading(doc, "2. Apresentação do período")
    add_para(doc, "Neste período a turma fecha a parte \"invisível\" da rede (nomes e endereços, iniciada na "
                  "Noite 6) e começa o projeto da rede de uma casa: planejar a instalação, configurar o "
                  "roteador, garantir a cobertura do Wi-Fi e deixar tudo seguro. Cada encontro parte do que o "
                  "aluno traz de casa e termina com uma produção concreta (planta, orçamento, roteador "
                  "configurado, mapa de sinal, laudo), que prepara o projeto final do módulo (Noite 15).")

    add_heading(doc, "3. Objetivo geral")
    add_para(doc, "Capacitar o aluno a planejar, configurar e proteger uma rede doméstica com acesso à "
                  "internet, entendendo como os aparelhos encontram sites e uns aos outros, e a orientar o "
                  "cliente sobre uso seguro.")

    add_heading(doc, "4. Competências e habilidades trabalhadas")
    for t in [
        "Ler e interpretar endereços de sites, identificando links legítimos e fraudulentos",
        "Planejar uma instalação de rede a partir da planta de um imóvel",
        "Elaborar lista de materiais e orçamento de serviço",
        "Configurar um roteador doméstico de ponta a ponta",
        "Medir e diagnosticar a cobertura de Wi-Fi e propor soluções",
        "Aplicar boas práticas de segurança e orientar o cliente",
        "Trabalhar em equipe, registrar o serviço e apresentar resultados",
    ]:
        add_bullet(doc, t)

    add_heading(doc, "5. Cronograma do período")
    add_table(doc, CRONOGRAMA, col_widths=[2.8, 2.0, 6.0, 6.2])
    add_para(doc, "* Datas por semana. Feriado de Finados em 02/11 (segunda-feira): se o encontro cair "
                  "nesse dia, a Noite 9 passa para a semana seguinte e a Noite 11 fica para o plano do "
                  "próximo período. A Noite 11 só acontece neste período se o dia da aula cair entre 12/11 "
                  "e 14/11.", italic=True, size=9, color=COR_CINZA)

    add_heading(doc, "6. Detalhamento por encontro")
    for aula in AULAS:
        add_heading(doc, aula["titulo"], level=2)
        add_para(doc, "Objetivos específicos", bold=True, color=COR_PRIMARIA)
        for t in aula["objetivos"]:
            add_bullet(doc, t)
        add_para(doc, "Conteúdos", bold=True, color=COR_PRIMARIA)
        for t in aula["conteudos"]:
            add_bullet(doc, t)
        add_para(doc, "Metodologia (mín. 50% do tempo em atividade do aluno)", bold=True, color=COR_PRIMARIA)
        for t in aula["metodologia"]:
            add_bullet(doc, t)
        add_para(doc, "Recursos", bold=True, color=COR_PRIMARIA)
        add_para(doc, aula["recursos"])
        add_para(doc, "Avaliação", bold=True, color=COR_PRIMARIA)
        add_para(doc, aula["avaliacao"])
        add_hr(doc)

    add_heading(doc, "7. Metodologia geral")
    for t in [
        "Aulas no padrão prática-pesada: pelo menos metade do encontro é atividade do aluno",
        "Exposição em blocos curtos (10 a 15 min), sempre seguida de prática — nunca dois blocos expositivos seguidos",
        "Trabalho em equipe e em dupla com material real de bancada (roteadores, cabos, celulares)",
        "Produções concretas apresentadas à turma: cartaz, planta, orçamento, laudo, cartilha",
        "Tarefa curta de casa ligando um encontro ao próximo (ex.: planta da casa para a Noite 8)",
        "Linguagem respeitosa, analogias do cotidiano e foco no trabalho de campo do técnico",
    ]:
        add_bullet(doc, t)

    add_heading(doc, "8. Recursos didáticos")
    for t in [
        "Projetor, quadro e slides de cada encontro",
        "Folha do aluno (atividade guiada) e resumo para levar para casa, a cada encontro",
        "Roteadores de bancada, cabos de rede, alicate crimpador, conectores RJ45, régua de tomadas",
        "Celulares do professor e dos alunos (hotspot, apps de medição de sinal)",
        "Planilha modelo de orçamento (.ods), cartolinas, canetões, materiais impressos",
    ]:
        add_bullet(doc, t)

    add_heading(doc, "9. Avaliação")
    add_para(doc, "Avaliação contínua e processual, com foco no fazer:")
    add_table(doc, [
        ["Instrumento", "O que se observa"],
        ["Participação nas práticas", "Envolvimento na bancada, nas estações e nas equipes"],
        ["Folha do aluno", "Atividades guiadas preenchidas durante o encontro"],
        ["Produções da equipe", "Cartilha, planta + orçamento, ficha de entrega, mapa de sinal, laudo de auditoria"],
        ["Apresentação", "Clareza ao explicar as escolhas técnicas para a turma"],
        ["Quiz de saída / tarefas", "Retenção dos conceitos e retorno das tarefas de casa"],
    ], col_widths=[5.0, 12.0])

    add_heading(doc, "10. Recuperação e acompanhamento")
    for t in [
        "Retomada dos conceitos no início de cada encontro (censo, galeria, revisão de tarefa)",
        "Resumo do aluno impresso para quem faltou, com atividade de reposição",
        "Alunos com mais facilidade atuam como apoio das equipes na bancada",
        "Atividades extras na caixa de ferramentas didáticas de cada roteiro, para quem termina antes",
    ]:
        add_bullet(doc, t)

    add_heading(doc, "11. Referências")
    for t in [
        "KUROSE, J. F.; ROSS, K. W. Redes de computadores e a Internet: uma abordagem top-down. Pearson.",
        "TANENBAUM, A. S. Redes de computadores. Pearson.",
        "CERT.br. Cartilha de Segurança para Internet. Disponível em: cartilha.cert.br.",
        "Manuais dos fabricantes dos roteadores de bancada (TP-Link TL-WR840N e Archer C5).",
        "Material próprio do professor: roteiros, slides, folhas e resumos das Noites 7 a 11.",
    ]:
        add_bullet(doc, t, size=10)

    doc.add_paragraph()
    doc.add_paragraph()
    add_para(doc, "Itapiranga – SC, 14 de outubro de 2026.")
    doc.add_paragraph()
    add_table(doc, [
        ["Professor", "Coordenação"],
        ["\n\n______________________________\nJoni Nedel", "\n\n______________________________\n"],
    ], col_widths=[8.5, 8.5])

    path = os.path.join(OUTDIR, "Plano de Aula Mensal - Redes - 15-10 a 14-11-2026.docx")
    doc.save(path)
    print(path)


if __name__ == "__main__":
    gerar()
