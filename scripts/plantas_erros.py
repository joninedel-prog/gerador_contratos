"""Desenha as duas plantas com erros (atividade extra: sala sem energia).

Planta A — Casa da Dona Lurdes. Planta B — Agropecuária do Seu Ademir.
Cada uma tem 12 erros e 3 acertos (pegadinhas). A grade A–F × 1–4 serve
pro aluno dizer onde achou o erro ("C3").
"""
from PIL import Image, ImageDraw, ImageFont

W, H = 2400, 1600
X0, X1 = 120, 2340          # área da grade
Y0, Y1 = 100, 1360
COLS, ROWS = "ABCDEF", 4

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
F_TXT = ImageFont.truetype(FONT, 27)
F_DEV = ImageFont.truetype(FONT_B, 30)
F_ROOM = ImageFont.truetype(FONT_B, 34)
F_GRID = ImageFont.truetype(FONT_B, 40)
F_TIT = ImageFont.truetype(FONT_B, 44)

PAREDE = (40, 40, 40)
CABO = (25, 95, 200)          # cabo de rede
FIBRA = (20, 150, 70)
ELETRICA = (205, 45, 35)
DEV_FILL = (222, 236, 252)
NOTA_FILL = (255, 246, 190)
APAR_FILL = (225, 225, 225)
GRADE = (175, 175, 175)


def novo(titulo):
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    cw = (X1 - X0) / 6
    rh = (Y1 - Y0) / ROWS
    for i in range(7):
        x = X0 + i * cw
        for y in range(Y0, Y1, 24):
            d.line([(x, y), (x, y + 12)], fill=GRADE, width=2)
    for j in range(ROWS + 1):
        y = Y0 + j * rh
        for x in range(X0, X1, 24):
            d.line([(x, y), (x + 12, y)], fill=GRADE, width=2)
    for i, c in enumerate(COLS):
        d.text((X0 + (i + .5) * cw, Y0 - 38), c, font=F_GRID, fill=(110, 110, 110), anchor="mm")
    for j in range(ROWS):
        d.text((60, Y0 + (j + .5) * rh), str(j + 1), font=F_GRID, fill=(110, 110, 110), anchor="mm")
    img.titulo = titulo
    return img, d


def salvar(img, caminho):
    final = Image.new("RGB", (W, H + 80), "white")
    final.paste(img, (0, 80))
    ImageDraw.Draw(final).text((X0, 20), img.titulo, font=F_TIT, fill=(11, 92, 92))
    final.save(caminho)


def sala(d, x0, y0, x1, y1, nome, pos="tl", metal=False):
    d.rectangle([x0, y0, x1, y1], outline=PAREDE, width=9 if not metal else 14)
    if metal:
        for x in range(int(x0) + 10, int(x1), 34):
            d.line([(x, y0 + 4), (x + 18, y0 + 4)], fill=(150, 150, 150), width=6)
            d.line([(x, y1 - 4), (x + 18, y1 - 4)], fill=(150, 150, 150), width=6)
    if pos == "tl":
        d.text((x0 + 18, y0 + 14), nome, font=F_ROOM, fill=(90, 90, 90))
    elif pos == "bl":
        d.text((x0 + 18, y1 - 52), nome, font=F_ROOM, fill=(90, 90, 90))
    elif pos == "tr":
        d.text((x1 - 18, y0 + 14), nome, font=F_ROOM, fill=(90, 90, 90), anchor="ra")
    else:
        d.text((x1 - 18, y1 - 52), nome, font=F_ROOM, fill=(90, 90, 90), anchor="ra")


def caixa(d, cx, cy, linhas, fill=DEV_FILL, borda=(60, 60, 60), titulo_negrito=True):
    """Caixa centrada em (cx, cy). Primeira linha em negrito."""
    fonts = [F_DEV if (i == 0 and titulo_negrito) else F_TXT for i in range(len(linhas))]
    larg = max(d.textlength(t, font=f) for t, f in zip(linhas, fonts)) + 30
    alt = sum(40 if f is F_DEV else 34 for f in fonts) + 18
    x0, y0 = cx - larg / 2, cy - alt / 2
    d.rounded_rectangle([x0, y0, x0 + larg, y0 + alt], radius=12, fill=fill, outline=borda, width=3)
    y = y0 + 10
    for t, f in zip(linhas, fonts):
        d.text((x0 + 15, y), t, font=f, fill=(20, 20, 20))
        y += 40 if f is F_DEV else 34
    return (x0, y0, x0 + larg, y0 + alt)


def linha(d, pts, cor=CABO, largura=8, tracejado=False):
    if not tracejado:
        d.line(pts, fill=cor, width=largura, joint="curve")
        return
    for (xa, ya), (xb, yb) in zip(pts, pts[1:]):
        n = max(1, int(((xb - xa) ** 2 + (yb - ya) ** 2) ** .5 // 30))
        for k in range(n):
            if k % 2:
                continue
            t0, t1 = k / n, (k + 1) / n
            d.line([(xa + (xb - xa) * t0, ya + (yb - ya) * t0),
                    (xa + (xb - xa) * t1, ya + (yb - ya) * t1)], fill=cor, width=largura)


def legenda(d):
    y = 1420
    x = X0
    for cor, nome, tr in [(CABO, "cabo de rede", False), (FIBRA, "fibra do provedor", False),
                          (ELETRICA, "fiação elétrica 220 V", True)]:
        linha(d, [(x, y + 20), (x + 90, y + 20)], cor=cor, tracejado=tr)
        d.text((x + 105, y + 2), nome, font=F_TXT, fill=(30, 30, 30))
        x += 105 + d.textlength(nome, font=F_TXT) + 70
    d.rounded_rectangle([x, y, x + 60, y + 40], radius=8, fill=DEV_FILL, outline=(60, 60, 60), width=3)
    d.text((x + 75, y + 2), "equipamento", font=F_TXT, fill=(30, 30, 30))
    x += 75 + d.textlength("equipamento", font=F_TXT) + 60
    d.rounded_rectangle([x, y, x + 60, y + 40], radius=8, fill=NOTA_FILL, outline=(60, 60, 60), width=3)
    d.text((x + 75, y + 2), "papel / etiqueta / anotação", font=F_TXT, fill=(30, 30, 30))
    d.text((X0, y + 70), "Ache 12 erros e 3 coisas que estão certas. Diga onde está cada um pela grade (ex.: C3).",
           font=F_DEV, fill=(11, 92, 92))


def poste(d, x, y):
    d.ellipse([x - 22, y - 22, x + 22, y + 22], fill=(110, 110, 110))
    d.text((x + 34, y), "poste", font=F_TXT, fill=(60, 60, 60), anchor="lm")


# ============================================================
# PLANTA A — CASA DA DONA LURDES
# ============================================================
def planta_a(caminho):
    img, d = novo("PLANTA A · Casa da Dona Lurdes")
    YM = 760
    d.text((1230, 150), "R U A", font=F_ROOM, fill=(150, 150, 150), anchor="mm")
    # cômodos
    sala(d, 160, 220, 660, YM, "Garagem")
    sala(d, 660, 220, 1500, YM, "Sala")
    sala(d, 1500, 220, 2300, YM, "Quarto 1", pos="tr")
    sala(d, 160, YM, 660, 1330, "Área de serviço", pos="tl")
    sala(d, 660, YM, 1240, 1330, "Cozinha")
    sala(d, 1240, YM, 1720, 1330, "Escritório", pos="tr")
    sala(d, 1720, YM, 1960, 1330, "Banh.")
    sala(d, 1960, YM, 2300, 1330, "Quarto 2", pos="tr")

    # fibra: poste -> ONU
    poste(d, 230, 150)
    linha(d, [(230, 172), (230, 900)], cor=FIBRA, largura=7)

    # câmera do portão (cabo por fora)
    linha(d, [(500, 300), (500, 225), (120, 225), (120, 1010), (180, 1010)])
    caixa(d, 500, 335, ["Câmera do portão", "(lado de fora)"])
    caixa(d, 410, 500, ["Cabo da câmera:", "cabo de rede comum,", "de uso interno, preso", "por fora da parede,", "no sol e na chuva"],
          fill=NOTA_FILL, titulo_negrito=False)

    # ONU e roteador no canto
    caixa(d, 340, 862, ["ONU do provedor", "endereço: 192.168.1.1"])
    caixa(d, 300, 1020, ["ROTEADOR", "no chão, atrás da", "máquina de lavar", "painel: 192.168.1.1", "senha do painel:", "trocada pelo técnico"])
    maq = (455, 1140, 640, 1236)
    d.rectangle(maq, fill=APAR_FILL, outline=(90, 90, 90), width=3)
    d.text(((maq[0] + maq[2]) / 2, (maq[1] + maq[3]) / 2), "Máquina\nde lavar", font=F_TXT, fill=(40, 40, 40),
           anchor="mm", align="center")
    caixa(d, 395, 1285, ["Benjamim (T): ONU + roteador", "+ máquina de lavar"], fill=NOTA_FILL, titulo_negrito=False)

    # cabo pra TV dentro do conduíte elétrico
    linha(d, [(430, 960), (630, 960), (630, 330), (900, 330)])
    linha(d, [(645, 975), (645, 345), (900, 345)], cor=ELETRICA, tracejado=True, largura=6)
    caixa(d, 1020, 330, ["Smart TV", "no cabo"])
    caixa(d, 1030, 560, ["O cabo da TV sobe dentro do", "mesmo conduíte da fiação", "da tomada (220 V)"], fill=NOTA_FILL,
          titulo_negrito=False)

    # bilhete na geladeira
    d.rectangle([690, 820, 830, 1000], fill=APAR_FILL, outline=(90, 90, 90), width=3)
    d.text((760, 910), "Gela-\ndeira", font=F_TXT, fill=(40, 40, 40), anchor="mm", align="center")
    caixa(d, 1030, 940, ["Bilhete na geladeira:", "Wi-Fi: TP-Link_3F2A", "Senha: 12345678", "(a da etiqueta)",
                         "Segurança: WEP", "Visitas: usam esta", "mesma senha"], fill=NOTA_FILL)

    # cabo do escritório no chão, por baixo da porta
    linha(d, [(430, 1090), (680, 1090), (680, 1275), (1270, 1275), (1270, 960), (1315, 960)])
    caixa(d, 960, 1200, ["Cabo solto no chão da cozinha,", "passa por baixo da porta"], fill=NOTA_FILL,
          titulo_negrito=False)
    caixa(d, 1480, 960, ["PC do escritório", "cabo CAT5e, 15 m", "testado: 8 fios OK", "cópia de segurança", "do roteador salva aqui"])
    caixa(d, 1480, 1215, ["Impressora", "endereço automático:", "cada semana muda;", "\"os PCs perdem ela\""])

    # quartos
    caixa(d, 1900, 470, ["Notebook", "sinal: 1 de 4 barras", "\"a internet trava\""])
    caixa(d, 2130, 1000, ["REPETIDOR", "ligado no quarto", "do fundo, onde o", "sinal chega com", "1 barra"])

    legenda(d)
    salvar(img, caminho)


# ============================================================
# PLANTA B — AGROPECUÁRIA DO SEU ADEMIR
# ============================================================
def planta_b(caminho):
    img, d = novo("PLANTA B · Agropecuária do Seu Ademir")
    d.text((900, 150), "R U A", font=F_ROOM, fill=(150, 150, 150), anchor="mm")
    sala(d, 160, 220, 1560, 1000, "Loja", pos="br")
    sala(d, 1560, 220, 2300, 720, "Escritório")
    sala(d, 1560, 720, 1930, 1000, "Banh.", pos="bl")
    sala(d, 1930, 720, 2300, 1000, "Copa", pos="br")
    d.text((190, 1240), "Pátio", font=F_ROOM, fill=(150, 150, 150))
    sala(d, 1000, 1130, 2300, 1350, "Depósito (galpão de chapa metálica)", metal=True)

    # fibra -> ONU
    poste(d, 2150, 150)
    linha(d, [(2150, 172), (2150, 255)], cor=FIBRA, largura=7)
    caixa(d, 2050, 300, ["ONU + roteador do provedor", "DHCP ligado (entrega endereço)"])
    caixa(d, 2050, 450, ["Etiqueta colada no aparelho:", "painel: usuário admin", "senha admin (nunca trocada)"],
          fill=NOTA_FILL)
    caixa(d, 1730, 620, ["PC do escritório", "cabo CAT6 testado OK", "reserva: 192.168.0.50"])
    caixa(d, 2115, 620, ["Impressora", "de etiquetas", "reserva: 192.168.0.50"])

    # ONU -> switch, dividindo o eletroduto com o ar-condicionado
    linha(d, [(1810, 330), (1215, 330)])
    caixa(d, 1410, 260, ["Ar-condicionado"], fill=APAR_FILL)
    linha(d, [(1410, 285), (1410, 348), (1240, 348)], cor=ELETRICA, tracejado=True, largura=6)
    linha(d, [(1410, 348), (1560, 348)], cor=ELETRICA, tracejado=True, largura=6)
    caixa(d, 1400, 445, ["Cabo de rede no mesmo", "eletroduto da fiação", "do ar-condicionado"], fill=NOTA_FILL,
          titulo_negrito=False)
    caixa(d, 1050, 300, ["SWITCH 8 portas", "no forro, em cima da", "lâmpada, sem ventilação"])

    # roteador da loja
    linha(d, [(885, 290), (730, 290)])
    caixa(d, 520, 360, ["ROTEADOR da loja (TP-Link)", "na parede, a 2 m de altura,", "no meio da loja",
                        "ligado pela porta LAN", "DHCP também ligado", "DNS: 8.8.8.8 e 1.1.1.1"])

    # cabo atravessando o corredor até o caixa
    linha(d, [(960, 363), (960, 755), (450, 755), (450, 777)])
    caixa(d, 350, 620, ["Placa no balcão:", "WI-FI GRÁTIS — CLIENTES", "Rede: Agro_Ademir",
                        "Senha: ademir2020", "(a mesma rede do caixa", "e das câmeras)"], fill=NOTA_FILL)
    caixa(d, 380, 840, ["CAIXA: PC + maquininha", "conector com a capa fora", "do RJ45 (fios aparecendo)"])
    caixa(d, 790, 845, ["Cabo atravessa o corredor", "no chão, sem canaleta;", "o carrinho passa por cima"],
          fill=NOTA_FILL, titulo_negrito=False)
    d.rectangle([180, 915, 300, 990], fill=APAR_FILL, outline=(90, 90, 90), width=3)
    d.text((240, 952), "Freezer", font=F_TXT, fill=(40, 40, 40), anchor="mm")
    caixa(d, 640, 955, ["Mesma régua: PC, maquininha,", "freezer de sorvete. Sem nobreak"], fill=NOTA_FILL,
          titulo_negrito=False)

    # câmera do pátio com emenda
    linha(d, [(1020, 363), (1020, 1060), (730, 1060)])
    caixa(d, 520, 1080, ["Câmera do pátio", "cabo emendado com fita", "isolante, pendurado na calha"])

    # câmera do depósito: 120 m
    linha(d, [(1080, 363), (1080, 1095), (2180, 1095), (2180, 1187)])
    caixa(d, 2050, 1250, ["Câmera do depósito", "cabo direto do switch: 120 m"])
    caixa(d, 1450, 1250, ["REPETIDOR dentro do galpão", "pra pegar o Wi-Fi da loja", "(o coletor do estoque não conecta)"])

    legenda(d)
    salvar(img, caminho)


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    planta_a(f"{out}/planta_a.png")
    planta_b(f"{out}/planta_b.png")
