# -*- coding: utf-8 -*-
"""Gera Equalizacao-modelo.xlsx: ficha de EQ por canal, faixas e sintomas, guia por fonte, checklist e exemplo."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

PURPLE_T = HUE["purple"][1]
BLUE, BLUE_T = HUE["blue"]
N = 24  # canais na ficha
G1 = 14  # primeira linha de canal
GAIN_FMT = '+0.0;-0.0;0'

# ------------------------------------------------------------------ guia por fonte (alimenta a conferência do passa-altas)
# fonte, HPF mínimo, HPF máximo, pode ficar desligado, faixa de partida, onde costuma sobrar, onde pode faltar, observação
GUIA = [
    ("Voz masculina", 80, 120, "Não", "80 a 120 Hz",
     "200 a 350 Hz: embolado. 800 Hz a 1,5 kHz: anasalado. 5 a 8 kHz: sibilância.",
     "2 a 4 kHz: clareza das palavras (2 a 3 dB, banda larga).",
     "Fala: entender cada palavra do fundo da sala. Canto: equalize com a voz forte."),
    ("Voz feminina", 100, 180, "Não", "100 a 180 Hz",
     "200 a 400 Hz: embolado. 2,5 a 4 kHz: áspero nas notas fortes. 5 a 8 kHz: sibilância.",
     "3 a 5 kHz: presença. Shelving em 10 a 12 kHz: ar (2 a 3 dB).",
     "Vale também para voz infantil, lapela e headset. Lapela no peito costuma soar abafada."),
    ("Coral", 100, 180, "Não", "100 a 180 Hz",
     "200 a 400 Hz: embolado com o som da sala.",
     "3 a 5 kHz: presença, com cuidado: microfone distante apita antes.",
     "O microfone distante capta muita sala. Aproxime antes de equalizar."),
    ("Violão", 80, 150, "Não", "80 a 150 Hz",
     "100 a 250 Hz: ressonância do corpo, onde apita. 1 a 3 kHz: som duro de captador.",
     "Shelving em 8 a 10 kHz: brilho das cordas.",
     "Sozinho, passa-altas perto de 80 Hz. Com banda, mais alto: o baixo cuida do grave."),
    ("Guitarra", 80, 120, "Não", "80 a 120 Hz",
     "200 a 400 Hz: embolado. 2 a 4 kHz: se disputar com a voz.",
     "Raramente precisa.",
     "O timbre se acerta no amplificador, antes da mesa."),
    ("Caixa da bateria", 80, 120, "Não", "80 a 120 Hz",
     "400 a 800 Hz: som oco, de papelão.",
     "150 a 250 Hz: corpo. 3 a 5 kHz: estalo da esteira.",
     "O passa-altas reduz o bumbo que entra no microfone da caixa."),
    ("Tons e surdo", 40, 100, "Não", "40 a 100 Hz",
     "300 a 600 Hz: som oco, de papelão.",
     "80 a 150 Hz: corpo. 3 a 5 kHz: ataque.",
     "Quanto maior o tambor, mais baixo o filtro."),
    ("Pratos e chimbal", 200, 400, "Não", "200 a 400 Hz",
     "1 a 3 kHz: som de lata, áspero.",
     "Shelving em 10 kHz: brilho, com cuidado.",
     "O passa-altas deixa passar só os pratos, e não o resto da bateria."),
    ("Teclado", 40, 120, "Sim", "desligado, ou 40 a 120 Hz",
     "200 a 400 Hz: embolado com a banda. 2 a 4 kHz: se disputar com a voz.",
     "Raramente precisa.",
     "Filtro baixo quando o teclado faz o grave; mais alto quando há baixista. Ajuste o timbre no teclado."),
    ("Baixo", 30, 40, "Sim", "desligado, ou 30 a 40 Hz",
     "200 a 300 Hz: embolado. 40 a 60 Hz: se a sala retumbar.",
     "60 a 100 Hz: peso. 700 Hz a 1 kHz: definição. 2 a 3 kHz: ataque.",
     "É a definição, e não o peso, que faz o baixo ser ouvido em caixas pequenas."),
    ("Bumbo", 30, 40, "Sim", "desligado, ou 30 a 40 Hz",
     "300 a 500 Hz: caixa de papelão (cortes de 4 a 8 dB são comuns).",
     "50 a 80 Hz: peso. 2 a 5 kHz: ataque do batedor.",
     "Decida com o baixo quem fica mais embaixo. Não realce os dois no mesmo ponto."),
    ("Computador e mídia", 30, 40, "Sim", "desligado, ou 30 a 40 Hz",
     "Raramente: o material já vem equalizado.",
     "Raramente.",
     "Se a mídia soa mal no sistema, o defeito pode ser da sala ou das caixas."),
]
FONTES = [g[0] for g in GUIA] + ["Outra"]
GUIA_R1, GUIA_R2 = 5, 5 + len(GUIA) - 1
GUIA_RNG = f"'Guia por fonte'!$B${GUIA_R1}:$I${GUIA_R2}"

# ------------------------------------------------------------------ as sete faixas
FAIXAS = [  # nome, de, até, o que mora ali, em excesso, em falta
    ("Subgrave", 20, 60, "Peso do bumbo, notas mais graves do baixo e do teclado.", "Ronco, vibração no palco.", "Bumbo e baixo sem peso."),
    ("Grave", 60, 250, "Corpo do baixo, do bumbo e do violão; fundamental das vozes.", "Retumbante, pesado.", "Magro, sem corpo."),
    ("Médio-grave", 250, 500, "Calor das vozes e dos instrumentos. Onde tudo se acumula.", "Embolado, caixa de papelão.", "Oco, sem calor."),
    ("Médio", 500, 2000, "Miolo das vozes, do violão, do teclado e da guitarra.", "Anasalado, som de telefone.", "Distante, vazio."),
    ("Médio-agudo", 2000, 4000, "Consoantes e clareza da fala. O ouvido é mais sensível aqui.", "Áspero, estridente.", "Sem clareza nas palavras."),
    ("Agudo", 4000, 8000, "Definição, ataque, o s e o ch das vozes.", "Sibilante, cortante.", "Opaco, sem ataque."),
    ("Brilho", 8000, 20000, "Brilho dos pratos e das cordas, ar nas vozes.", "Chiado, agudo que cansa.", "Abafado, sem ar."),
]
SINTOMAS = [  # como soa, onde procurar, primeira medida
    ("Ronco, vento, pof nas letras P e B, batida de mão", "abaixo de 100 Hz", "Ligar ou subir o passa-altas."),
    ("Retumbante, pesado", "100 a 250 Hz", "Subir um pouco o passa-altas, ou cortar com o shelving de graves."),
    ("Embolado, turvo", "200 a 400 Hz", "Corte de 2 a 4 dB com uma banda bell."),
    ("Caixa de papelão, som oco de tambor", "300 a 600 Hz", "Corte com uma banda bell. No bumbo, o corte costuma ser maior."),
    ("Anasalado, corneta, telefone", "800 Hz a 1,5 kHz", "Corte de 2 a 4 dB com uma banda bell."),
    ("Áspero, estridente, cansa o ouvido", "2 a 5 kHz", "Corte suave e largo. Confira antes se há realce demais nessa região."),
    ("Sibilante: s e ch que espetam", "5 a 8 kHz", "Corte leve, ou o de-esser (estudo Dinâmica e efeitos)."),
    ("Abafado, ou sem clareza nas palavras", "2 kHz para cima", "Conferir microfone e posição. Cortar o que embola (200 a 400 Hz). Só então, realce largo de 2 a 3 dB."),
]
# defeitos do exercício 2: nome, região típica (texto), de, até
DEFEITOS = [("Retumbante", "100 a 250 Hz", 100, 250), ("Embolado", "200 a 400 Hz", 200, 400),
            ("Anasalado", "800 Hz a 1,5 kHz", 800, 1500), ("Áspero", "2 a 5 kHz", 2000, 5000),
            ("Sibilante", "5 a 8 kHz", 5000, 8000)]

CHECK = [
    ("Antes", "O ganho de cada canal foi ajustado antes do EQ, e o EQ começou plano."),
    ("Antes", "A fonte, o microfone e a posição foram conferidos antes de equalizar."),
    ("Canal", "O passa-altas está ligado em tudo o que não faz o grave da música, na faixa de partida de cada fonte."),
    ("Canal", "Cada ajuste responde a um defeito que foi ouvido e tem nome."),
    ("Canal", "Os cortes vieram antes dos realces, e o ponto de cada corte foi achado por varredura."),
    ("Canal", "Os realces são poucos, largos e de até cerca de 6 dB. Nenhum canal tem só realces."),
    ("Canal", "Cada canal foi comparado com o EQ desligado, no mesmo volume, e ficou melhor com ele ligado."),
    ("Canal", "O nível de cada canal foi conferido no medidor depois do EQ."),
    ("Mistura", "O EQ foi decidido na mistura: a voz principal está clara por cima da banda, e o bumbo e o baixo se distinguem."),
    ("Sistema", "O mesmo corte não se repete em quase todos os canais. Se repete, o sistema foi testado com música gravada."),
    ("Retorno", "O EQ de cada retorno tem só cortes, poucos e moderados, e foi comparado com o equalizador desligado."),
    ("Registro", "O EQ de cada canal e de cada saída está anotado, e qualquer mudança no EQ de sistema foi registrada."),
]

# Exemplo 1 do treinamento: o EQ como foi encontrado na passagem de som
EX_FICHA = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "mesa": "Mesa digital, passa-altas e quatro bandas por canal"},
    "lim": (6, 4),
    "canais": [  # fonte, tipo, HPF, bandas [(tipo, Hz, dB, Q)]
        ("Voz da cantora", "Voz feminina", None, [("Bell", 3000, 8, 1.4), ("Shelf", 10000, 9, None)]),
        ("Voz de apoio", "Voz feminina", 100, [("Bell", 3150, 6, 1.4), ("Shelf", 10000, 4, None)]),
        ("Voz do pastor", "Voz masculina", 100, [("Bell", 300, -3, 1.4), ("Bell", 2500, 3, 1.0)]),
        ("Violão", "Violão", 80, [("Bell", 200, -4, 1.4), ("Bell", 3000, 5, 1.4)]),
        ("Baixo", "Baixo", None, [("Bell", 250, -3, 1.4), ("Bell", 800, 3, 1.4)]),
        ("Bumbo", "Bumbo", None, [("Bell", 60, 3, 1.4), ("Bell", 400, -6, 2.9), ("Bell", 3500, 4, 1.4)]),
        ("Teclado", "Teclado", None, [("Bell", 2500, 6, 1.4), ("Shelf", 8000, 3, None)]),
        ("Computador (som dos vídeos)", "Computador e mídia", None, []),
    ],
    "ajustes": [  # canal, o que foi feito
        ("1", "Realces em 0. Passa-altas em 120 Hz. Varredura: sobra em 250 Hz, corte de 3 dB. Depois, +3 dB em 3 kHz (Q 0,7) e +2 dB no shelving."),
        ("2", "Realces em 0. Passa-altas em 120 Hz. Corte de 3 dB em 300 Hz. Nenhum realce: o apoio fica atrás da voz principal."),
        ("4", "O realce de 5 dB em 3 kHz virou corte largo de 3 dB (Q 0,7). Passa-altas em 100 Hz."),
        ("7", "Realces em 0. Passa-altas em 100 Hz, porque há baixista. Corte largo de 3 dB em 3 kHz."),
        ("3, 5, 6, 8", "Mantidos. O ataque do bumbo em 3,5 kHz é curto e não cobre a voz."),
        ("Todos", "Com a voz mais limpa, o ganho da cantora foi conferido no medidor antes de montar os retornos."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


BANDS = [("F", "G", "H", "I"), ("J", "K", "L", "M"), ("N", "O", "P", "Q"), ("R", "S", "T", "U")]
F_W = {"A": 2, "B": 5, "C": 26, "D": 18, "E": 8}
for b in BANDS:
    F_W.update({b[0]: 7, b[1]: 8, b[2]: 7, b[3]: 6})
F_W.update({"V": 16, "W": 15, "X": 9, "Y": 9, "Z": 16, "AA": 42, "AB": 26, "AC": 2})
SIT_CF = [("OK", GREEN), ("Realce alto", YELLOW), ("Sem HPF", YELLOW), ("HPF fora da faixa", YELLOW), ("Só realces", YELLOW)]
HPF_CF = [("OK", GREEN), ("Desligado", GREEN), ("Sem HPF", YELLOW), ("Abaixo da faixa", YELLOW), ("Acima da faixa", YELLOW)]
AV_DISPUTA = "Disputa: muitos realces nesta faixa. Corte nos canais que cobrem a fonte principal"
AV_SISTEMA = "Mesmo corte em muitos canais: teste o sistema com música gravada"


# ================================================================== aba Ficha de EQ
def aba_ficha(ws, tab, data=None):
    widths(ws, F_W)
    is_ex = data is not None
    title(ws, "Ficha de EQ por canal — passa-altas e até quatro bandas",
          "Exemplo preenchido, para consulta. Use a aba Ficha de EQ para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro com o EQ de cada canal. As células cinza são calculadas.", "AB")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None), (8, "Mesa", "mesa", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:I{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    lim = data["lim"] if is_ex else (6, 4)
    for r, rot, v, dica in [(4, "Realce máximo (dB)", lim[0], "Realce acima disto recebe o aviso Realce alto. Ponto de partida: 6 dB."),
                            (5, "Aviso de faixa repetida (nº)", lim[1], "Quantos realces, ou cortes, na mesma faixa disparam o aviso do mapa. Ponto de partida: 4.")]:
        label(ws, f"K{r}", rot, merge=f"K{r}:Q{r}")
        inp(ws, f"R{r}", v, h="center", bg=hb)
        hint(ws, f"S{r}", dica, merge=f"S{r}:AB{r}", h="left")
    dv_number(ws, "R4:R5")
    RMAX, NREP = "$R$4", "$R$5"
    hint(ws, "K7", "Frequências em Hz (3000, e não 3 k). Ganho em dB, com sinal: −4 para corte. Q só na banda bell.", merge="K7:AB7", h="left")
    hint(ws, "K8", "Passa-altas em branco = desligado. A faixa de partida vem da aba Guia por fonte.", merge="K8:AB8", h="left")

    band(ws, 10, "Ficha de EQ por canal", "AB")
    for col, t, bg in [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Tipo de fonte", INK), ("E", "HPF", PURPLE)]:
        head(ws, f"{col}11", t, bg=bg)
    for k, b in enumerate(BANDS, 1):
        put(ws, f"{b[0]}11", f"Banda {k}", f=font(10, True, c=WHITE), bg=PURPLE, h="center", merge=f"{b[0]}11:{b[3]}11")
    for col, t in [("V", "Partida do HPF"), ("W", "Passa-altas"), ("X", "Maior realce"), ("Y", "Maior corte"),
                   ("Z", "Situação"), ("AA", "O que fazer")]:
        head(ws, f"{col}11", t, bg=BLUE)
    head(ws, "AB11", "Observação")
    ws.row_dimensions[11].height = 30
    hints = {"B": "nº", "C": "o que está no canal", "D": "da lista", "E": "Hz", "V": "do Guia por fonte", "W": "conferência",
             "X": "dB", "Y": "dB", "Z": "", "AA": "", "AB": ""}
    for b in BANDS:
        hints.update({b[0]: "Tipo", b[1]: "Hz", b[2]: "dB", b[3]: "Q"})
    for col, t in hints.items():
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 21.75
    if is_ex:
        hint(ws, "B13", "O EQ como foi encontrado na passagem de som, antes da correção. As colunas de V a AA são calculadas.",
             merge="B13:AB13", h="left")
    else:
        exrow = {"B": "Ex.", "C": "Voz do pastor", "D": "Voz masculina", "E": 100, "F": "Bell", "G": 300, "H": -3, "I": 1.4,
                 "J": "Bell", "K": 2500, "L": 3, "M": 1, "V": "80 a 120 Hz", "W": "OK", "X": 3, "Y": -3, "Z": "OK",
                 "AA": "Manter. Decida na mistura"}
        for col in hints:
            ex(ws, f"{col}13", exrow.get(col, ""), h="left" if col in ("C", "D", "AA") else "center",
               fmt=GAIN_FMT if col in ("H", "L", "X", "Y") else None)
        note(ws, "D11", "Escolha na lista. O tipo define a faixa de partida do passa-altas.\nUse Outra quando a fonte não estiver na lista: a conferência do passa-altas fica em branco.")
        note(ws, "E11", "Frequência do passa-altas, em Hz.\nDeixe em branco se ele está desligado.")
        note(ws, "F11", "Tipo: Bell para a banda em sino, Shelf para a prateleira de graves ou de agudos.\nHz: frequência central (ou o ponto do shelf).\ndB: ganho, positivo para realce e negativo para corte.\nQ: largura (só no bell).")
        note(ws, "W11", "OK: dentro da faixa de partida. Desligado: a fonte pode ficar sem o filtro.\nSem HPF: a fonte pede o filtro. Abaixo ou acima: confira ouvindo, pode estar certo.")
        note(ws, "Z11", "Ordem dos avisos: Realce alto, Sem HPF, HPF fora da faixa, Só realces.\nSão convites para conferir: o ouvido decide.")
    ws.row_dimensions[13].height = 24

    G2 = G1 + N - 1
    bg = WHITE if is_ex else INPUT
    for k in range(N):
        r = G1 + k
        c = data["canais"][k] if is_ex and k < len(data["canais"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, bg=bg)
        inp(ws, f"E{r}", c[2] if c else None, h="center", bg=bg, fmt="0")
        for j, b in enumerate(BANDS):
            v = c[3][j] if c and j < len(c[3]) else (None, None, None, None)
            inp(ws, f"{b[0]}{r}", v[0], h="center", bg=bg)
            inp(ws, f"{b[1]}{r}", v[1], h="center", bg=bg, fmt="0")
            inp(ws, f"{b[2]}{r}", v[2], h="center", bg=bg, fmt=GAIN_FMT)
            inp(ws, f"{b[3]}{r}", v[3], h="center", bg=bg, fmt="0.0")
        gains = ",".join(f"{b[2]}{r}" for b in BANDS)
        nums = ",".join([f"E{r}"] + [f"{b[i]}{r}" for b in BANDS for i in (1, 2, 3)])
        nb = "+".join(f"({b[2]}{r}>0)" for b in BANDS)
        nc = "+".join(f"({b[2]}{r}<0)" for b in BANDS)
        D, E, V, W, X, Y, Z = (f"{col}{r}" for col in ("D", "E", "V", "W", "X", "Y", "Z"))
        calc(ws, V, f'=IF({D}="","",IFERROR(VLOOKUP({D},{GUIA_RNG},5,0),"–"))', b=False, sz=9)
        calc(ws, W, f'=IF({D}="","",IFERROR(IF({E}="",IF(VLOOKUP({D},{GUIA_RNG},4,0)="Sim","Desligado","Sem HPF"),'
                    f'IF({E}<VLOOKUP({D},{GUIA_RNG},2,0),"Abaixo da faixa",IF({E}>VLOOKUP({D},{GUIA_RNG},3,0),"Acima da faixa","OK"))),"–"))',
             b=False)
        calc(ws, X, f'=IF(COUNT({gains})=0,"",MAX(0,{gains}))', fmt=GAIN_FMT, b=False)
        calc(ws, Y, f'=IF(COUNT({gains})=0,"",MIN(0,{gains}))', fmt=GAIN_FMT, b=False)
        calc(ws, Z, f'=IF(AND(C{r}="",{D}="",COUNT({nums})=0),"",IF(AND({X}<>"",{X}>{RMAX}),"Realce alto",'
                    f'IF({W}="Sem HPF","Sem HPF",IF(OR({W}="Abaixo da faixa",{W}="Acima da faixa"),"HPF fora da faixa",'
                    f'IF(AND({nb}>0,{nc}=0),"Só realces","OK")))))')
        calc(ws, f"AA{r}", f'=IF({Z}="","",IF({Z}="Realce alto","Realce de "&{X}&" dB: procure antes o que cortar e confira o microfone",'
                           f'IF({Z}="Sem HPF","Ligue o passa-altas: comece em "&{V},'
                           f'IF({Z}="HPF fora da faixa","Partida sugerida: "&{V}&". Confira ouvindo",'
                           f'IF({Z}="Só realces","Procure o que cortar antes de realçar","Manter. Decida na mistura")))))',
             h="left", b=False, sz=9)
        inp(ws, f"AB{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{G1}:D{G2}", FONTES, "Escolha o tipo de fonte. Outra, se não estiver na lista.")
    dv_number(ws, f"E{G1}:E{G2}")
    for b in BANDS:
        dv_list(ws, f"{b[0]}{G1}:{b[0]}{G2}", ["Bell", "Shelf"], "Bell (sino) ou Shelf (prateleira).")
        dv_number(ws, f"{b[1]}{G1}:{b[3]}{G2}")
    cf_equal(ws, f"W{G1}:W{G2}", HPF_CF)
    cf_equal(ws, f"Z{G1}:Z{G2}", SIT_CF)

    # ---- mapa de cortes e realces por faixa
    s = G2 + 2
    band(ws, s, "Mapa de cortes e realces por faixa, somando todos os canais", "AB")
    r = s + 1
    for col, t, mg in [("B", "Faixa", "B{r}:C{r}"), ("D", "De (Hz)", None), ("E", "Até (Hz)", None), ("F", "Realces", "F{r}:I{r}"),
                       ("J", "Cortes", "J{r}:M{r}"), ("N", "Aviso", "N{r}:AB{r}")]:
        put(ws, f"{col}{r}", t, f=font(10, True), bg=GRAY, h="center", merge=mg.format(r=r) if mg else None)
    ws.row_dimensions[r].height = 21.75
    m1 = r + 1
    for k, (nome, lo, hi, *_rest) in enumerate(FAIXAS):
        r = m1 + k
        put(ws, f"B{r}", nome, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
        put(ws, f"D{r}", lo, bg=GRAY, h="center", fmt="0")
        put(ws, f"E{r}", hi, bg=GRAY, h="center", fmt="0")
        menor = '"<="' if k == len(FAIXAS) - 1 else '"<"'

        def conta(sinal):
            return "+".join(f'COUNTIFS({b[1]}${G1}:{b[1]}${G2},">="&$D{r},{b[1]}${G1}:{b[1]}${G2},{menor}&$E{r},'
                            f'{b[2]}${G1}:{b[2]}${G2},"{sinal}0")' for b in BANDS)
        calc(ws, f"F{r}", "=" + conta(">"), merge=f"F{r}:I{r}")
        calc(ws, f"J{r}", "=" + conta("<"), merge=f"J{r}:M{r}")
        calc(ws, f"N{r}", f'=IF(F{r}+J{r}=0,"",IF(F{r}>={NREP},"{AV_DISPUTA}",IF(J{r}>={NREP},"{AV_SISTEMA}","OK")))',
             b=False, h="left", merge=f"N{r}:AB{r}")
        ws.row_dimensions[r].height = 19.5
    m2 = m1 + len(FAIXAS) - 1
    cf_status(ws, f"N{m1}:AB{m2}", f"$N{m1}")
    hint(ws, f"B{m2 + 1}", "Conta as bandas com ganho diferente de zero pela frequência. O shelf entra pela frequência do ponto. O passa-altas não entra na conta.",
         merge=f"B{m2 + 1}:AB{m2 + 1}", h="left")

    # ---- resumo
    s = m2 + 3
    band(ws, s, "Resumo automático", "AB")
    cont = lambda t: f'COUNTIF(Z{G1}:Z{G2},"{t}")'
    linhas = [
        ("Canais registrados", f"=COUNTA(C{G1}:C{G2})"),
        ("Na situação OK", f'={cont("OK")}'),
        ("Realce alto", f'={cont("Realce alto")}'),
        ("Sem passa-altas", f'={cont("Sem HPF")}'),
        ("Passa-altas fora da faixa", f'={cont("HPF fora da faixa")}'),
        ("Só realces", f'={cont("Só realces")}'),
        ("Faixas em disputa", f'=COUNTIF(N{m1}:N{m2},"{AV_DISPUTA}")'),
        ("Faixas com o mesmo corte em muitos canais", f'=COUNTIF(N{m1}:N{m2},"{AV_SISTEMA}")'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:I{s + k}")
    a = s + len(linhas) + 1
    R = lambda k: f"F{s + k}"
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(COUNT($R$4:$R$5)<2,"Informe os dois limites no alto da aba",IF({R(1)}=0,"Liste os canais",'
                      f'IF({R(8)}>0,"Mesmo corte em muitos canais: o defeito pode ser do sistema ou da sala",'
                      f'IF({R(7)}>0,"Há faixa disputada: abra espaço com cortes nos canais que cobrem a voz",'
                      f'IF({R(3)}+{R(4)}+{R(5)}+{R(6)}>0,"Há canais a conferir: veja a coluna O que fazer","OK")))))',
         b=False, merge=f"F{a}:AB{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:AB{a}")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "AB")
        r = end + 3
        put(ws, f"B{r}", "Canal", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:AB{r}")
        ws.row_dimensions[r].height = 21.75
        for canal, feito in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", canal, f=font(10, True, c=MUTED), bg=GRAY, h="center", merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:AB{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"F{G1}"
    setup(ws, tab, f"B1:AB{end}", fit_height=True)


# ================================================================== aba Faixas e sintomas
def aba_faixas(ws):
    widths(ws, {"A": 2, "B": 30, "C": 13, "D": 16, "E": 22, "F": 30, "G": 30, "H": 34, "I": 2})
    title(ws, "Faixas e sintomas — consulta e calculadora",
          "Digite uma frequência para ver a faixa e como ela soa. Digite a frequência e o Q para ver de onde até onde a banda age.", "H")

    # tabela das sete faixas (fica embaixo, mas é a base das consultas)
    t0 = 25
    band(ws, t0, "As sete faixas", "H")
    for col, t in zip("BCDEF", ["Faixa", "De (Hz)", "Até (Hz)", "O que mora ali", "Em excesso, soa"]):
        put(ws, f"{col}{t0 + 1}", t, f=font(10, True), bg=GRAY, h="center")
    put(ws, f"G{t0 + 1}", "Em falta, soa", f=font(10, True), bg=GRAY, h="center", merge=f"G{t0 + 1}:H{t0 + 1}")
    f1 = t0 + 2
    for k, (nome, lo, hi, mora, exc, fal) in enumerate(FAIXAS):
        r = f1 + k
        put(ws, f"B{r}", nome, f=font(10, True))
        put(ws, f"C{r}", lo, h="center", fmt="0")
        put(ws, f"D{r}", hi, h="center", fmt="0")
        put(ws, f"E{r}", mora)
        put(ws, f"F{r}", exc)
        put(ws, f"G{r}", fal, merge=f"G{r}:H{r}")
        ws.row_dimensions[r].height = 31.5
    f2 = f1 + len(FAIXAS) - 1
    NOMES, LOS, EXC = f"$B${f1}:$B${f2}", f"$C${f1}:$C${f2}", f"$F${f1}:$F${f2}"

    def faixa(F):
        return f'IF({F}="","",IF(OR({F}<20,{F}>20000),"Fora de 20 Hz a 20 kHz",INDEX({NOMES},MATCH({F},{LOS},1))))'

    def excesso(F):
        return f'IF({F}="","",IF(OR({F}<20,{F}>20000),"",INDEX({EXC},MATCH({F},{LOS},1))))'

    # consulta por frequência
    band(ws, 4, "Consulta por frequência (exercício 2: o que você ouviu na varredura)", "H")
    for col, t, bg in [("B", "Defeito ouvido", INK), ("C", "Frequência", PURPLE), ("D", "Faixa", BLUE), ("E", "Em excesso, soa", BLUE),
                       ("F", "Região típica do defeito", INK), ("G", "Conferência", BLUE), ("H", "Observação", INK)]:
        head(ws, f"{col}5", t, bg=bg)
    for col, t in zip("BCDEFGH", ["", "Hz", "", "", "", "", ""]):
        hint(ws, f"{col}6", t)
    q1 = 7
    for k in range(8):
        r = q1 + k
        F = f"C{r}"
        if k < len(DEFEITOS):
            nome, txt, lo, hi = DEFEITOS[k]
            put(ws, f"B{r}", nome, f=font(10, True), bg=GRAY)
            put(ws, f"F{r}", txt, bg=GRAY, h="center")
            calc(ws, f"G{r}", f'=IF({F}="","",IF(AND({F}>={lo},{F}<={hi}),"Dentro da região típica","Fora da região típica: confira"))',
                 b=False, sz=9)
        else:
            inp(ws, f"B{r}")
            put(ws, f"F{r}", "–", bg=GRAY, h="center")
            put(ws, f"G{r}", "", bg=GRAY)
        inp(ws, F, h="center", fmt="0")
        calc(ws, f"D{r}", "=" + faixa(F), b=True)
        calc(ws, f"E{r}", "=" + excesso(F), b=False, sz=9)
        inp(ws, f"H{r}")
        ws.row_dimensions[r].height = 21.75
    q2 = q1 + 7
    dv_number(ws, f"C{q1}:C{q2}")
    cf_equal(ws, f"G{q1}:G{q1 + len(DEFEITOS) - 1}", [("Dentro da região típica", GREEN), ("Fora da região típica: confira", YELLOW)])
    hint(ws, f"B{q2 + 1}", "A frequência de um defeito muda com a voz, o microfone e a sala. Fora da região típica não é erro: é a sua voz. As três últimas linhas são para consulta livre.",
         merge=f"B{q2 + 1}:H{q2 + 1}", h="left")

    # calculadora de largura
    c0 = q2 + 3
    band(ws, c0, "Calculadora de largura de uma banda bell", "H")
    for col, t, bg in [("B", "Uso", INK), ("C", "Frequência", PURPLE), ("D", "Q", PURPLE), ("E", "Largura (oitavas)", BLUE),
                       ("F", "De (Hz)", BLUE), ("G", "Até (Hz)", BLUE), ("H", "Observação", INK)]:
        head(ws, f"{col}{c0 + 1}", t, bg=bg)
    usos = ["Banda larga (realce)", "Corte comum", "Corte estreito (microfonia)", "Livre"]
    qs = [0.7, 2, 8, None]
    for k in range(4):
        r = c0 + 2 + k
        put(ws, f"B{r}", usos[k], f=font(10, True), bg=GRAY)
        inp(ws, f"C{r}", None, h="center", fmt="0")
        inp(ws, f"D{r}", qs[k], h="center", fmt="0.0")
        x = f"(1/(2*D{r})+(1/(4*D{r}^2)+1)^0.5)"
        calc(ws, f"E{r}", f'=IF(OR(D{r}="",D{r}<=0),"",ROUND(2*LOG10({x})/LOG10(2),2))', fmt="0.00")
        calc(ws, f"F{r}", f'=IF(OR(C{r}="",E{r}=""),"",ROUND(C{r}/2^(E{r}/2),0))', fmt="0")
        calc(ws, f"G{r}", f'=IF(OR(C{r}="",E{r}=""),"",ROUND(C{r}*2^(E{r}/2),0))', fmt="0")
        inp(ws, f"H{r}")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"C{c0 + 2}:D{c0 + 5}")
    note(ws, f"D{c0 + 1}", "Q alto, banda estreita. Q 1,4 é cerca de uma oitava; Q 4,3, um terço de oitava.")
    hint(ws, f"B{c0 + 6}", "Conta usual do Q (pontos de meia intensidade em torno da frequência central). Alguns fabricantes usam outra definição: trate o resultado como ordem de grandeza.",
         merge=f"B{c0 + 6}:H{c0 + 6}", h="left")
    assert c0 + 6 < t0 - 1

    # sintomas
    s0 = f2 + 2
    band(ws, s0, "Do que se ouve para onde procurar", "H")
    put(ws, f"B{s0 + 1}", "Como soa", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"C{s0 + 1}", "Onde procurar", f=font(10, True), bg=GRAY, h="center", merge=f"C{s0 + 1}:D{s0 + 1}")
    put(ws, f"E{s0 + 1}", "Primeira medida", f=font(10, True), bg=GRAY, h="center", merge=f"E{s0 + 1}:H{s0 + 1}")
    r = s0 + 1
    for como, onde, med in SINTOMAS:
        r += 1
        put(ws, f"B{r}", como, f=font(10, True))
        put(ws, f"C{r}", onde, h="center", merge=f"C{r}:D{r}")
        put(ws, f"E{r}", med, merge=f"E{r}:H{r}")
        ws.row_dimensions[r].height = 31.5
    hint(ws, f"B{r + 1}", "Faixas típicas, para servir de ponto de partida. Os limites entre as faixas são uma convenção e variam um pouco de um autor para outro.",
         merge=f"B{r + 1}:H{r + 1}", h="left")
    setup(ws, HUE["purple"][0], f"B1:H{r + 1}", fit_height=True)


# ================================================================== aba Guia por fonte
def aba_guia(ws):
    widths(ws, {"A": 2, "B": 22, "C": 9, "D": 9, "E": 12, "F": 22, "G": 44, "H": 38, "I": 44, "J": 2})
    title(ws, "Guia por fonte — ponto de partida",
          "Onde procurar primeiro, e não receita. As colunas C a F alimentam a conferência do passa-altas na Ficha de EQ.", "I")
    r = GUIA_R1 - 1
    for col, t, bg in [("B", "Fonte", INK), ("C", "HPF mín. (Hz)", PURPLE), ("D", "HPF máx. (Hz)", PURPLE),
                       ("E", "Pode ficar desligado", PURPLE), ("F", "Passa-altas para começar", PURPLE),
                       ("G", "Onde costuma sobrar (cortar)", BLUE), ("H", "Onde pode faltar (realçar com cuidado)", BLUE),
                       ("I", "Observação", INK)]:
        head(ws, f"{col}{r}", t, bg=bg)
    ws.row_dimensions[r].height = 30
    for g in GUIA:
        r += 1
        put(ws, f"B{r}", g[0], f=font(10, True))
        put(ws, f"C{r}", g[1], h="center", fmt="0")
        put(ws, f"D{r}", g[2], h="center", fmt="0")
        put(ws, f"E{r}", g[3], h="center")
        put(ws, f"F{r}", g[4], h="center")
        for col, v in zip("GHI", g[5:]):
            put(ws, f"{col}{r}", v)
        ws.row_dimensions[r].height = 31.5
    hint(ws, f"B{r + 1}", "Valores típicos. Mudam com a voz, o instrumento, o microfone, a distância e a sala. Se mudar os números das colunas C e D, a conferência da ficha acompanha.",
         merge=f"B{r + 1}:I{r + 1}", h="left")
    ws.freeze_panes = f"C{GUIA_R1}"
    setup(ws, HUE["slate"][0], f"B1:I{r + 1}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Equalização — ficha de EQ por canal"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Equalização — como usar esta planilha",
           "Ficha de EQ de cada canal, consulta de faixas e sintomas e guia de ponto de partida por fonte.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo de fonte e tipo de banda (Bell ou Shelf) aceitam só as opções da lista."),
        ("Números", "Frequência em Hz, sem a letra k: 3000, e não 3 k. Ganho em dB, com o sinal de menos no corte: -4."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Ficha de EQ: preencha o cabeçalho e confira os dois limites dos avisos, no alto à direita."),
        ("2", "Liste a fonte e o tipo de fonte de cada canal. O tipo traz a faixa de partida do passa-altas."),
        ("3", "Na mesa: ajuste o ganho, ligue o passa-altas, procure o que sobra por varredura, corte e, se faltar, realce pouco."),
        ("4", "Anote o passa-altas e cada banda usada: tipo, frequência, ganho e Q."),
        ("5", "Veja a situação e o que fazer em cada canal. Depois, o mapa de faixas: ele mostra disputa e cortes repetidos."),
        ("6", "Aba Faixas e sintomas: consulte uma frequência, ou calcule a largura de uma banda pelo Q."),
        ("7", "Aba Checklist: confira os doze itens antes de montar os retornos."),
    ]),
    ("Abas da planilha", [
        ("Ficha de EQ", "Até 24 canais, com conferência do passa-altas, maior realce, maior corte, situação e o mapa por faixa."),
        ("Faixas e sintomas", "Consulta por frequência, calculadora de largura pelo Q, as sete faixas e a tabela de sintomas."),
        ("Guia por fonte", "Passa-altas de partida e onde cada fonte costuma sobrar ou faltar, para doze tipos de fonte."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Ficha de EQ", "Os oito canais do exemplo 1 do treinamento, como foram encontrados, e o que foi ajustado."),
    ]),
    ("Regras e premissas do modelo", [
        ("Passa-altas", "OK dentro da faixa de partida do Guia por fonte. Em branco, vale Desligado nas fontes que podem ficar sem ele e Sem HPF nas outras."),
        ("Realce alto", "Maior realce do canal acima do limite do alto da aba (6 dB, de partida)."),
        ("Só realces", "O canal tem realce e nenhum corte em bell ou shelf. Convite para procurar o que cortar."),
        ("Mapa de faixas", "Conta, em todos os canais, as bandas que realçam e as que cortam em cada faixa. Muitos realces na mesma faixa indicam disputa; muitos cortes, defeito do sistema ou da sala."),
        ("Avisos", "São convites para conferir, e não erros. O ouvido decide. Os limites são regra prática deste material, e não norma."),
        ("Exemplos", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_ficha(wb.create_sheet("Ficha de EQ"), HUE["purple"][0])
aba_faixas(wb.create_sheet("Faixas e sintomas"))
aba_guia(wb.create_sheet("Guia por fonte"))
checklist(wb.create_sheet("Checklist"), "Checklist da equalização",
          "Confira na passagem de som, depois de equalizar os canais. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["teal"][0])
aba_ficha(wb.create_sheet("Exemplo - Ficha de EQ"), MUTED, EX_FICHA)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
