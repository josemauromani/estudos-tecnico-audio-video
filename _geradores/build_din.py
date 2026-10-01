# -*- coding: utf-8 -*-
"""Gera Dinamica-e-Efeitos-modelo.xlsx: ficha de dinâmica por canal, ficha de efeitos, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

RED_H, RED_T = HUE["red"]
SLATE, SLATE_T = HUE["slate"]
PURPLE_H, PURPLE_T = HUE["purple"]
SIM_NAO = ["Sim", "Não"]
NC, NG = 16, 8      # compressores e gates na aba de dinâmica
NFX, NMUS, NENV = 4, 6, 16  # processadores, músicas e canais na aba de efeitos

CHECK = [
    ("Antes", "O ganho e o equalizador de cada canal foram ajustados antes de ligar compressor, gate e efeitos."),
    ("Compressor", "Cada compressor ligado tem um motivo: a fonte varia mais do que o fader consegue acompanhar."),
    ("Compressor", "Nos trechos fortes, o medidor de redução marca de 3 a 6 dB, e volta a zero nas pausas."),
    ("Compressor", "O make-up não passa da redução, e o canal foi comparado com o compressor desligado, no mesmo volume."),
    ("Compressor", "Os microfones distantes (coral, púlpito, lapela, ambiente) estão sem compressor ou com compressão leve e sem make-up."),
    ("Gate", "Cada gate abre com o golpe mais fraco da própria peça e fica fechado com o vazamento dos vizinhos."),
    ("Gate", "Não há gate em microfone de fala."),
    ("Limiter", "Os limiters das saídas atuam só nos picos raros, e não ficam acesos o tempo todo."),
    ("Efeitos", "Os envios de efeito saem de auxiliar pós-fader, e o processador está em 100% de efeito."),
    ("Efeitos", "O retorno de cada efeito chega ao master e não volta para o próprio envio."),
    ("Efeitos", "O tempo do reverb combina com a sala e com a música, e o delay está no tempo da música."),
    ("Efeitos", "Os efeitos ficam fechados (mute) na fala: pregação, avisos, oração e leitura."),
]

# Exemplo 1 do treinamento: compressor ligado em tudo, do mesmo jeito (leituras em dBFS)
EX_DIN = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "mesa": "Mesa digital de 24 canais, medidor em dBFS"},
    "ref": (3, 6, 10, 1, 3),
    "comp": [  # fonte, threshold, razão, attack, release, make-up, pico de entrada, redução lida
        ("Voz do pastor (headset)", -30, 8, 1, 50, 10, -8, 18),
        ("Voz da cantora", -30, 8, 1, 50, 10, -6, 20),
        ("Voz de apoio", -20, 3, 10, 150, 3, -10, 7),
        ("Violão", -18, 2.5, 20, 150, 2, -10, 5),
        ("Baixo", -20, 4, 30, 200, 3, -12, 6),
        ("Teclado", -10, 2, 30, 300, 0, -12, 0),
        ("Bumbo", -14, 4, 20, 100, 2, -6, 6),
        ("Coral (dois microfones suspensos)", -30, 4, 10, 200, 8, -22, 6),
    ],
    "gate": [  # fonte, threshold, range, attack, hold, release, vazamento, som útil mais fraco
        ("Bumbo", -30, 20, 0.5, 60, 150, -36, -14),
        ("Tom 1", -30, 15, 0.5, 100, 300, -24, -12),
        ("Surdo", -10, 15, 0.5, 100, 350, -26, -12),
    ],
    "ajustes": [  # canal, o que foi feito, leitura depois
        ("Pastor", "Threshold de −30 para −16, razão de 8:1 para 3:1, make-up de +10 para +3, attack de 15 ms, release de 200 ms.", "redução de 5 dB"),
        ("Cantora", "Threshold de −30 para −14, razão de 8:1 para 4:1, make-up de +10 para +4, attack de 10 ms, release de 150 ms.", "redução de 6 dB"),
        ("Apoio", "Threshold de −20 para −18.", "redução de 5 dB"),
        ("Teclado", "Compressor desligado: o pico nem chegava ao threshold, e o teclado não precisa.", "sem compressor"),
        ("Coral", "Compressor desligado: microfone distante com 8 dB de make-up apita nas pausas.", "sem compressor"),
        ("Tom 1", "Threshold do gate de −30 para −18: fica 6 dB acima da caixa que vaza e 6 dB abaixo do golpe fraco.", "abre só no tom"),
        ("Surdo", "Threshold do gate de −10 para −19: os golpes fracos voltaram a abrir a porta.", "nenhum golpe engolido"),
    ],
}

# Exemplo 2 do treinamento: o reverb que embolou o louvor
EX_FX = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares, sala que já reverbera bastante",
             "data": date(2026, 10, 4), "operador": "Operador de som da escala"},
    "ref": (2.5, 40, 0, 5),
    "fx": [  # tipo, tempo, pré-delay, repetição, aux de envio, envio, retorno em, volta ao envio, no retorno de palco
        ("Reverb hall", 2.8, 0, None, "Aux 5", "Pós-fader", "Retorno FX 1", "Não", "Não"),
        ("Delay", 500, None, 35, "Aux 6", "Pré-fader", "Canais 15 e 16", "Não", "Não"),
    ],
    "mus": [  # música, BPM, nota usada, tempo no aparelho
        ("Louvor lento", 72, "Colcheia", 500),
        ("Louvor animado", 120, "Colcheia pontuada", 500),
        ("Música final", 100, "Semínima", 500),
    ],
    "env": [  # fonte, tipo, envios FX1 a FX4
        ("Voz do pastor (headset)", "Fala", -10, None, None, None),
        ("Voz da cantora", "Voz cantada", -5, -12, None, None),
        ("Voz de apoio 1", "Voz cantada", -8, None, None, None),
        ("Voz de apoio 2", "Voz cantada", -8, None, None, None),
        ("Violão", "Instrumento", -10, None, None, None),
        ("Baixo", "Grave", -10, None, None, None),
        ("Teclado", "Instrumento", None, None, None, None),
        ("Bumbo", "Grave", -15, None, None, None),
        ("Caixa da bateria", "Bateria", 3, None, None, None),
        ("Computador (som dos vídeos)", "Mídia", -10, None, None, None),
    ],
    "ajustes": [
        ("FX 1", "Hall de 2,8 s trocado por plate de 1,4 s, com pré-delay de 40 ms, corte de graves em 200 Hz e de agudos em 6 kHz."),
        ("FX 2", "Envio do delay passado para auxiliar pós-fader. Repetição mantida em 35%."),
        ("Músicas", "Tempo do delay pelo tap a cada música: 417 ms, 375 ms e 600 ms."),
        ("Envios", "Fechados no pastor, no baixo, no bumbo e no computador. Caixa da bateria de +3 para −12 dB."),
        ("Operação", "Mute nos retornos de efeito a cada passagem da música para a fala."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def sub(ws, ref, text, merge=None):
    put(ws, ref, text, f=font(10, True), bg=GRAY, h="center", merge=merge)


def referencias(ws, linhas, lab1, lab2, val, hint1, hint2, bg, first=4):
    """Bloco de valores de referência, à direita do cabeçalho. Devolve as referências absolutas."""
    for k, (rot, v, dica) in enumerate(linhas):
        r = first + k
        label(ws, f"{lab1}{r}", rot, merge=f"{lab1}{r}:{lab2}{r}")
        inp(ws, f"{val}{r}", v, h="center", bg=bg)
        hint(ws, f"{hint1}{r}", dica, merge=f"{hint1}{r}:{hint2}{r}", h="left")
    dv_number(ws, f"{val}{first}:{val}{first + len(linhas) - 1}")
    return [f"${val}${first + k}" for k in range(len(linhas))]


# ================================================================== aba Dinâmica por canal
D_W = {"A": 2, "B": 7, "C": 31, "D": 11, "E": 9, "F": 9, "G": 9, "H": 10, "I": 11, "J": 11, "K": 11, "L": 11,
       "M": 11, "N": 11, "O": 14, "P": 56, "Q": 28, "R": 2}
D1 = 14
SIT_C = [("OK", GREEN), ("Leve", YELLOW), ("Forte", YELLOW), ("Sem ação", YELLOW), ("Make-up alto", YELLOW),
         ("Excesso", RED), ("Razão inválida", RED)]
SIT_G = [("OK", GREEN), ("Margem curta", YELLOW), ("Defina", YELLOW), ("Abre à toa", RED), ("Corta o som", RED),
         ("Sem janela", RED)]


def aba_din(ws, tab, data=None):
    widths(ws, D_W)
    is_ex = data is not None
    title(ws, "Dinâmica por canal — ficha de compressores e gates",
          "Exemplo preenchido, para consulta. Use a aba Dinâmica por canal para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro com os ajustes e com o que você leu nos medidores. As células cinza são calculadas.", "Q")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None), (8, "Mesa e medidor", "mesa", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    v = data["ref"] if is_ex else (3, 6, 10, 1, 3)
    GMIN, GMAX, GLIM, MTOL, GMAR = referencias(ws, [
        ("Redução mínima (dB)", v[0], "Abaixo dela, a atuação do compressor é leve."),
        ("Redução máxima (dB)", v[1], "Até aqui, a redução nos trechos fortes conta como OK."),
        ("Limite de excesso (dB)", v[2], "Acima dele, a compressão é excessiva."),
        ("Folga do make-up (dB)", v[3], "Quanto o make-up pode passar da redução antes do aviso."),
        ("Margem mínima do gate (dB)", v[4], "Distância mínima do threshold do gate ao vazamento e ao som útil."),
    ], "I", "L", "M", "N", "Q", hb)

    # ---------------------------------------------------------------- compressores
    band(ws, 10, "Compressores: ajuste, leitura e saída esperada pela razão", "Q", color=RED_H)
    heads = [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Threshold", RED_H), ("E", "Razão", RED_H),
             ("F", "Attack", RED_H), ("G", "Release", RED_H), ("H", "Make-up", RED_H), ("I", "Pico de entrada", RED_H),
             ("J", "Redução lida", RED_H), ("K", "Acima do threshold", SLATE), ("L", "Redução esperada", SLATE),
             ("M", "Pico de saída", SLATE), ("N", "Lida − esperada", SLATE), ("O", "Situação", SLATE),
             ("P", "O que fazer", SLATE), ("Q", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMNOPQ", ["nº", "o que está ligado no canal", "dB", "4 para 4:1", "ms", "ms", "dB",
                                           "dB, antes do compressor", "dB, no medidor GR", "dB", "dB", "dB, com make-up",
                                           "dB", "", "", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 27
    if is_ex:
        hint(ws, "B13", "Leituras em dBFS, com cada fonte no trecho mais forte, antes da correção. As colunas de K a P são calculadas.",
             merge="B13:Q13", h="left")
    else:
        for col, val in zip("BCDEFGHIJKLMNOP", ["Ex.", "Voz da cantora", -20, 4, 10, 150, 4, -8, 8, 12, 9, -13, -1, "Forte",
                                                "Redução forte: suba o threshold ou baixe a razão"]):
            ex(ws, f"{col}13", val, h="left" if col in "CP" else "center")
        ex(ws, "Q13", "")
        note(ws, "D11", "Nível a partir do qual o compressor atua, na mesma escala do medidor do canal.\nDigite só o número: -20, por exemplo.")
        note(ws, "E11", "Digite só o primeiro número da razão: 4 para 4:1, 2,5 para 2,5:1.\nPara um limiter (∞:1), digite 100.")
        note(ws, "I11", "O ponto mais alto que o medidor de entrada do canal alcançou no trecho mais forte, antes do compressor.")
        note(ws, "J11", "Quantos dB o medidor de redução (GR) marcou no trecho mais forte.\nDigite como número positivo: 6, e não -6.")
        note(ws, "L11", "Conta da razão: o que passa do threshold, menos esse valor dividido pela razão.\nO medidor pode marcar menos, por causa do attack e do knee.")
        note(ws, "M11", "Onde o pico deve chegar depois do compressor: entrada, menos a redução esperada, mais o make-up.")
        note(ws, "O11", "Usa a redução lida. Se ela estiver em branco, usa a esperada.")
    ws.row_dimensions[13].height = 24

    D2 = D1 + NC - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NC):
        r = D1 + k
        c = data["comp"][k] if is_ex and k < len(data["comp"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        for col, j in zip("DEFGHIJ", range(1, 8)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt='+0;-0;0' if col in "DHI" else None)
        D, E, Hc, I, J, K, L, M = (f"{x}{r}" for x in "DEHIJKLM")
        tem = f'AND({D}<>"",{E}<>"",{I}<>"")'
        calc(ws, K, f'=IF(OR({D}="",{I}=""),"",MAX(0,{I}-{D}))', b=False)
        calc(ws, L, f'=IF({tem},IF({E}<1,"",ROUND({K}*(1-1/{E}),1)),"")', fmt='0.0')
        calc(ws, M, f'=IF({L}="","",ROUND({I}-{L}+IF({Hc}="",0,{Hc}),1))', fmt='+0.0;-0.0;0.0')
        calc(ws, f"N{r}", f'=IF(OR({J}="",{L}=""),"",ROUND({J}-{L},1))', b=False, fmt='+0.0;-0.0;0.0')
        gr = f'IF({J}<>"",{J},{L})'
        mk = f'IF({Hc}="",0,{Hc})'
        calc(ws, f"O{r}", f'=IF(AND({E}<>"",{E}<1),"Razão inválida",IF(AND({J}="",{L}=""),"",'
                          f'IF({gr}>{GLIM},"Excesso",IF({mk}>{gr}+{MTOL},"Make-up alto",IF({gr}<0.5,"Sem ação",'
                          f'IF({gr}<{GMIN},"Leve",IF({gr}>{GMAX},"Forte","OK")))))))')
        O = f"O{r}"
        calc(ws, f"P{r}", f'=IF({O}="","",IF({O}="Razão inválida","A razão começa em 1. Digite 4 para 4:1",'
                          f'IF({O}="Excesso","Compressão em excesso: suba o threshold, baixe a razão e reduza o make-up",'
                          f'IF({O}="Make-up alto","Make-up acima da redução: baixe-o. Nas pausas, o canal fica mais perto de apitar",'
                          f'IF({O}="Sem ação","O sinal não chega ao threshold: baixe o threshold ou desligue o compressor",'
                          f'IF({O}="Leve","Atuação leve. Se a fonte ainda varia demais, baixe o threshold alguns dB",'
                          f'IF({O}="Forte","Redução forte: suba o threshold ou baixe a razão","Manter")))))))',
             h="left", b=False, sz=9)
        inp(ws, f"Q{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"D{D1}:J{D2}")
    cf_equal(ws, f"O{D1}:O{D2}", SIT_C)

    s = D2 + 2
    band(ws, s, "Resumo dos compressores", "Q")
    cont = lambda t: f'COUNTIF(O{D1}:O{D2},"{t}")'
    linhas = [
        ("Canais listados", f"=COUNTA(C{D1}:C{D2})"),
        ("Compressores com situação",
         "=" + "+".join(cont(t) for t in ("OK", "Leve", "Sem ação", "Forte", "Make-up alto", "Excesso", "Razão inválida"))),
        ("Na faixa (OK)", f'={cont("OK")}'),
        ("Leve ou sem ação", f'={cont("Leve")}+{cont("Sem ação")}'),
        ("Forte ou make-up alto", f'={cont("Forte")}+{cont("Make-up alto")}'),
        ("Em excesso", f'={cont("Excesso")}'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(COUNT(M4:M8)<5,"Informe os cinco valores de referência",'
                      f'IF(F{s + 2}=0,"Preencha threshold, razão e pico de entrada, ou a redução lida",'
                      f'IF(F{s + 6}+{cont("Razão inválida")}>0,"Há compressão em excesso: corrija antes do evento",'
                      f'IF(F{s + 5}>0,"Há compressor forte ou com make-up alto: confira",'
                      f'IF(F{s + 4}>0,"Há compressor que quase não atua: veja se ele é necessário","OK")))))',
         b=False, merge=f"F{a}:P{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:P{a}")

    # ---------------------------------------------------------------- gates
    g = a + 2
    band(ws, g, "Gates: a janela entre o vazamento e o som útil", "Q", color=SLATE)
    heads = [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Threshold", SLATE), ("E", "Range", SLATE),
             ("F", "Attack", SLATE), ("G", "Hold", SLATE), ("H", "Release", SLATE), ("I", "Vazamento", SLATE),
             ("J", "Som útil mais fraco", SLATE), ("K", "Janela", INK), ("L", "Margem abaixo", INK),
             ("M", "Margem acima", INK), ("N", "Threshold sugerido", INK), ("O", "Situação", INK),
             ("P", "O que fazer", INK), ("Q", "Observação", INK)]
    for col, t, bgc in heads:
        head(ws, f"{col}{g + 1}", t, bg=bgc)
    ws.row_dimensions[g + 1].height = 30
    for col, t in zip("BCDEFGHIJKLMNOPQ", ["nº", "peça ou fonte com gate", "dB", "dB", "ms", "ms", "ms",
                                           "pico, em dB", "pico, em dB", "útil − vazamento", "threshold − vazamento",
                                           "útil − threshold", "dB, no meio da janela", "", "", ""]):
        hint(ws, f"{col}{g + 2}", t)
    ws.row_dimensions[g + 2].height = 27
    if not is_ex:
        note(ws, f"I{g + 1}", "O pico mais forte do som que NÃO deve abrir o gate, lido no medidor do canal: a caixa vazando no microfone do tom, por exemplo.")
        note(ws, f"J{g + 1}", "O pico do golpe mais fraco da própria fonte, que ainda precisa abrir o gate.")
        note(ws, f"N{g + 1}", "O ponto do meio entre o vazamento e o som útil mais fraco. É um ponto de partida: confirme de ouvido.")
    G1 = g + 3
    G2 = G1 + NG - 1
    for k in range(NG):
        r = G1 + k
        c = data["gate"][k] if is_ex and k < len(data["gate"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        for col, j in zip("DEFGHIJ", range(1, 8)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt='+0;-0;0' if col in "DIJ" else None)
        D, I, J, K, O = (f"{x}{r}" for x in "DIJKO")
        lido = f'AND({I}<>"",{J}<>"")'
        calc(ws, K, f'=IF({lido},{J}-{I},"")', b=False)
        calc(ws, f"L{r}", f'=IF(AND({D}<>"",{I}<>""),{D}-{I},"")', b=False, fmt='+0;-0;0')
        calc(ws, f"M{r}", f'=IF(AND({D}<>"",{J}<>""),{J}-{D},"")', b=False, fmt='+0;-0;0')
        calc(ws, f"N{r}", f'=IF({lido},IF({K}<2*{GMAR},"",ROUND(({I}+{J})/2,0)),"")', fmt='+0;-0;0')
        calc(ws, O, f'=IF({lido},IF({K}<2*{GMAR},"Sem janela",IF({D}="","Defina",IF({D}<={I},"Abre à toa",'
                    f'IF({D}>={J},"Corta o som",IF(OR({D}-{I}<{GMAR},{J}-{D}<{GMAR}),"Margem curta","OK"))))),"")')
        calc(ws, f"P{r}", f'=IF({O}="","",IF({O}="Sem janela","O vazamento é quase tão forte quanto a fonte: aproxime o microfone ou fique sem gate",'
                          f'IF({O}="Defina","Comece pelo threshold sugerido e confirme de ouvido",'
                          f'IF({O}="Abre à toa","O vazamento abre o gate: suba o threshold para perto de "&N{r}&" dB",'
                          f'IF({O}="Corta o som","O gate engole os golpes fracos: baixe o threshold para perto de "&N{r}&" dB",'
                          f'IF({O}="Margem curta","Threshold perto demais de um dos lados: leve para perto de "&N{r}&" dB","Manter"))))))',
             h="left", b=False, sz=9)
        inp(ws, f"Q{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"D{G1}:J{G2}")
    cf_equal(ws, f"O{G1}:O{G2}", SIT_G)
    hint(ws, f"B{G2 + 1}", "As faixas de redução e a margem do gate são regra prática deste material, e não norma. Troque os valores no alto da aba, se precisar.",
         merge=f"B{G2 + 1}:Q{G2 + 1}", h="left")
    end = G2 + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "Q")
        r = end + 3
        sub(ws, f"B{r}", "Canal", merge=f"B{r}:C{r}")
        sub(ws, f"D{r}", "O que foi feito", merge=f"D{r}:O{r}")
        sub(ws, f"P{r}", "Leitura depois", merge=f"P{r}:Q{r}")
        ws.row_dimensions[r].height = 21.75
        for canal, feito, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", canal, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:O{r}")
            put(ws, f"P{r}", depois, merge=f"P{r}:Q{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{D1}"
    setup(ws, tab, f"B1:Q{end}", fit_height=True)


# ================================================================== aba Efeitos
F_W = {"A": 2, "B": 7, "C": 30, "D": 15, "E": 13, "F": 13, "G": 15, "H": 13, "I": 18, "J": 14, "K": 15, "L": 17,
       "M": 58, "N": 28, "O": 2}
TIPOS_FX = ["Reverb hall", "Reverb plate", "Reverb room", "Delay", "Outro"]
NOTAS = ["Semínima", "Colcheia", "Colcheia pontuada", "Semicolcheia"]
TIPOS_CANAL = ["Fala", "Voz cantada", "Instrumento", "Bateria", "Grave", "Mídia"]
SIT_FX = [("OK", GREEN), ("Incompleto", YELLOW), ("Tempo longo", YELLOW), ("Repetição alta", YELLOW),
          ("Envio pré", RED), ("Laço", RED)]
SIT_ENV = [("OK", GREEN), ("Sem tipo", YELLOW), ("Envio alto", YELLOW), ("Fala com efeito", RED),
           ("Grave com efeito", YELLOW), ("Mídia com efeito", YELLOW)]


def aba_fx(ws, tab, data=None):
    widths(ws, F_W)
    is_ex = data is not None
    title(ws, "Efeitos — processadores, tempo do delay e envios por canal",
          "Exemplo preenchido, para consulta. Use a aba Efeitos para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro. As células cinza calculam o tempo do delay e conferem o caminho de cada efeito.", "N")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local e sala", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:G{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    v = data["ref"] if is_ex else (2.5, 40, 0, 5)
    TMAX, RMAX, EMAX, DTOL = referencias(ws, [
        ("Tempo máximo de reverb (s)", v[0], "Acima dele, a aba avisa. Em sala que já reverbera, use um valor menor."),
        ("Repetição máxima do delay (%)", v[1], "Acima dela, o eco demora a morrer e embola."),
        ("Envio máximo (dB)", v[2], "Envio acima deste valor recebe aviso. O 0 é o ganho unitário."),
        ("Tolerância do tempo do delay (%)", v[3], "Diferença aceita entre o tempo calculado e o do aparelho."),
    ], "H", "J", "K", "L", "N", hb)

    bg = WHITE if is_ex else INPUT
    # ---------------------------------------------------------------- processadores
    band(ws, 9, "Processadores de efeito: ajuste e caminho do sinal", "N", color=PURPLE_H)
    heads = [("B", "Efeito", INK), ("C", "Tipo", PURPLE_H), ("D", "Tempo", PURPLE_H), ("E", "Pré-delay", PURPLE_H),
             ("F", "Repetição", PURPLE_H), ("G", "Auxiliar de envio", PURPLE_H), ("H", "Envio", PURPLE_H),
             ("I", "Retorno em", PURPLE_H), ("J", "Retorno volta ao envio?", PURPLE_H), ("K", "No retorno de palco?", PURPLE_H),
             ("L", "Situação", SLATE), ("M", "O que fazer", SLATE), ("N", "Observação", INK)]
    for col, t, bgc in heads:
        head(ws, f"{col}10", t, bg=bgc)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "da lista", "reverb: s; delay: ms", "ms", "%, no delay", "nome ou número",
                                        "Pré / Pós-fader", "canal ou retorno", "Sim / Não", "Sim / Não", "", "", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 27
    if not is_ex:
        note(ws, "D10", "No reverb: a duração da cauda (decay), em segundos. No delay: o intervalo entre as repetições, em milissegundos.")
        note(ws, "H10", "O auxiliar que alimenta o efeito é tirado antes ou depois do fader do canal?\nPara efeito, o certo é pós-fader.")
        note(ws, "J10", "O canal ou retorno por onde o efeito volta tem o envio para o próprio efeito aberto?\nSe tiver, o efeito realimenta a si mesmo.")
        note(ws, "K10", "O efeito também é enviado para os retornos de palco ou para os fones dos músicos?")
    X1 = 12
    X2 = X1 + NFX - 1
    for k in range(NFX):
        r = X1 + k
        c = data["fx"][k] if is_ex and k < len(data["fx"]) else None
        put(ws, f"B{r}", f"FX {k + 1}", f=font(10, True, c=MUTED), bg=GRAY, h="center")
        for col, j in zip("CDEFGHIJK", range(9)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="left" if col in "CGI" else "center", bg=bg)
        C, D, F, Hc, J, K, L = (f"{x}{r}" for x in "CDFHJKL")
        rev = f'OR({C}="Reverb hall",{C}="Reverb plate",{C}="Reverb room")'
        calc(ws, L, f'=IF({C}="","",IF({J}="Sim","Laço",IF({Hc}="Pré-fader","Envio pré",'
                    f'IF(AND({rev},{D}<>"",{D}>{TMAX}),"Tempo longo",IF(AND({C}="Delay",{F}<>"",{F}>{RMAX}),"Repetição alta",'
                    f'IF(OR({Hc}="",{J}=""),"Incompleto","OK"))))))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="Laço","Feche o envio do canal de retorno para o próprio efeito: ele realimenta a si mesmo",'
                          f'IF({L}="Envio pré","Passe o auxiliar para pós-fader: o efeito continua soando com o fader fechado",'
                          f'IF({L}="Tempo longo","Cauda longa para esta sala: reduza o tempo ou troque o tipo de reverb",'
                          f'IF({L}="Repetição alta","Reduza a repetição: o eco demora a morrer e embola",'
                          f'IF({L}="Incompleto","Informe o tipo de envio e se o retorno volta ao envio",'
                          f'IF({K}="Sim","Manter. Efeito no retorno de palco: pouco, e atenção à microfonia","Manter")))))))',
             h="left", b=False, sz=9)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"C{X1}:C{X2}", TIPOS_FX, "Reverb hall, Reverb plate, Reverb room, Delay ou Outro.")
    dv_list(ws, f"H{X1}:H{X2}", ["Pós-fader", "Pré-fader"], "O auxiliar de envio é pós-fader ou pré-fader?")
    dv_list(ws, f"J{X1}:K{X2}", SIM_NAO, "Sim ou Não.")
    dv_number(ws, f"D{X1}:F{X2}")
    cf_equal(ws, f"L{X1}:L{X2}", SIT_FX)

    # ---------------------------------------------------------------- tempo do delay
    m = X2 + 2
    band(ws, m, "Tempo do delay pela música: 60 000 ÷ BPM", "N", color=PURPLE_H)
    heads = [("B", "Ordem", INK), ("C", "Música", PURPLE_H), ("D", "Andamento", PURPLE_H), ("E", "Semínima", SLATE),
             ("F", "Colcheia", SLATE), ("G", "Colcheia pontuada", SLATE), ("H", "Semicolcheia", SLATE),
             ("I", "Figura usada", PURPLE_H), ("J", "Tempo a ajustar", SLATE), ("K", "Tempo no aparelho", PURPLE_H),
             ("L", "Situação", SLATE), ("M", "O que fazer", SLATE), ("N", "Observação", INK)]
    for col, t, bgc in heads:
        head(ws, f"{col}{m + 1}", t, bg=bgc)
    ws.row_dimensions[m + 1].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "nome ou momento", "BPM", "ms (1 tempo)", "ms (½ tempo)", "ms (¾ de tempo)",
                                        "ms (¼ de tempo)", "da lista", "ms", "ms, lido", "", "", ""]):
        hint(ws, f"{col}{m + 2}", t)
    ws.row_dimensions[m + 2].height = 27
    if not is_ex:
        note(ws, f"D{m + 1}", "Batidas por minuto. Conte as batidas em 15 segundos e multiplique por 4, ou use um aplicativo de metrônomo.")
        note(ws, f"K{m + 1}", "O tempo que o processador mostra depois do tap ou do ajuste manual. Deixe em branco se não for conferir.")
    M1 = m + 3
    M2 = M1 + NMUS - 1
    for k in range(NMUS):
        r = M1 + k
        c = data["mus"][k] if is_ex and k < len(data["mus"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, h="center", bg=bg)
        D, I, J, K, L = (f"{x}{r}" for x in "DIJKL")
        ok = f'AND({D}<>"",{D}>0)'
        for col, fator in (("E", "1"), ("F", "2"), ("G", "4*3"), ("H", "4")):
            calc(ws, f"{col}{r}", f'=IF({ok},ROUND(60000/{D}/{fator},0),"")', b=False)
        inp(ws, I, c[2] if c else None, h="center", bg=bg)
        calc(ws, J, f'=IF(OR(E{r}="",{I}=""),"",IF({I}="Semínima",E{r},IF({I}="Colcheia",F{r},IF({I}="Colcheia pontuada",G{r},H{r}))))')
        inp(ws, K, c[3] if c else None, h="center", bg=bg)
        calc(ws, L, f'=IF(OR({J}="",{K}=""),"",IF(ABS({K}-{J})<={J}*{DTOL}/100,"OK","Fora do tempo"))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="OK","Manter","Ajuste o delay para "&{J}&" ms, ou bata o tap de novo no ritmo da música"))',
             h="left", b=False, sz=9)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"I{M1}:I{M2}", NOTAS, "Semínima, Colcheia, Colcheia pontuada ou Semicolcheia.")
    dv_number(ws, f"D{M1}:D{M2}")
    dv_number(ws, f"K{M1}:K{M2}")
    cf_equal(ws, f"L{M1}:L{M2}", [("OK", GREEN), ("Fora do tempo", YELLOW)])

    # ---------------------------------------------------------------- envios por canal
    e = M2 + 2
    band(ws, e, "Envios de efeito por canal", "N", color=PURPLE_H)
    heads = [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Tipo de fonte", PURPLE_H), ("E", "Envio ao FX 1", PURPLE_H),
             ("F", "Envio ao FX 2", PURPLE_H), ("G", "Envio ao FX 3", PURPLE_H), ("H", "Envio ao FX 4", PURPLE_H),
             ("I", "Efeitos abertos", SLATE), ("J", "Maior envio", SLATE), ("K", "Situação", SLATE)]
    for col, t, bgc in heads:
        head(ws, f"{col}{e + 1}", t, bg=bgc)
    put(ws, f"L{e + 1}", "O que fazer", f=font(10, True, c=WHITE), bg=SLATE, h="center", merge=f"L{e + 1}:M{e + 1}")
    head(ws, f"N{e + 1}", "Observação")
    ws.row_dimensions[e + 1].height = 30
    for col, t in zip("BCDEFGHIJK", ["nº", "o que está ligado no canal", "da lista", "dB", "dB", "dB", "dB", "quantos", "dB", ""]):
        hint(ws, f"{col}{e + 2}", t)
    hint(ws, f"L{e + 2}", "", merge=f"L{e + 2}:M{e + 2}")
    hint(ws, f"N{e + 2}", "")
    ws.row_dimensions[e + 2].height = 27
    if not is_ex:
        note(ws, f"D{e + 1}", "Fala: pregação, avisos, leitura, palestra.\nGrave: bumbo e baixo.\nMídia: computador, vídeo, trilha.")
        note(ws, f"E{e + 1}", "Nível do envio do canal para este efeito, em dB. Deixe em branco se o envio está fechado.\nSe a mesa não mostra o número, anote a posição do botão na coluna Observação e digite um valor aproximado.")
    E1 = e + 3
    E2 = E1 + NENV - 1
    for k in range(NENV):
        r = E1 + k
        c = data["env"][k] if is_ex and k < len(data["env"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, h="center", bg=bg)
        for col, j in zip("EFGH", range(2, 6)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt='+0;-0;0')
        C, D, I, J, K = (f"{x}{r}" for x in "CDIJK")
        n = f"COUNT(E{r}:H{r})"
        calc(ws, I, f'=IF(AND({C}="",{D}="",{n}=0),"",{n})', b=False)
        calc(ws, J, f'=IF({n}=0,"",MAX(E{r}:H{r}))', b=False, fmt='+0;-0;0')
        calc(ws, K, f'=IF(AND({D}="",{n}=0),"",IF({D}="","Sem tipo",IF({n}=0,"Seco",IF({D}="Fala","Fala com efeito",'
                    f'IF({D}="Grave","Grave com efeito",IF({D}="Mídia","Mídia com efeito",IF({J}>{EMAX},"Envio alto","OK")))))))')
        calc(ws, f"L{r}", f'=IF({K}="","",IF({K}="Sem tipo","Escolha o tipo de fonte na lista",IF({K}="Seco","Canal sem efeito",'
                          f'IF({K}="Fala com efeito","Feche o envio: reverb e delay atrapalham o entendimento da fala",'
                          f'IF({K}="Grave com efeito","Feche o envio: efeito em bumbo e baixo embola os graves",'
                          f'IF({K}="Mídia com efeito","Feche o envio: o material já vem com o efeito pronto",'
                          f'IF({K}="Envio alto","Envio acima da referência: confira se o efeito não cobre a fonte",'
                          f'IF({D}="Voz cantada","Manter. Feche o efeito quando este microfone for usado para falar","Manter"))))))))',
             h="left", b=False, sz=9, merge=f"L{r}:M{r}")
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{E1}:D{E2}", TIPOS_CANAL, "Fala, Voz cantada, Instrumento, Bateria, Grave (bumbo e baixo) ou Mídia.")
    dv_number(ws, f"E{E1}:H{E2}")
    cf_equal(ws, f"K{E1}:K{E2}", SIT_ENV)

    s = E2 + 2
    band(ws, s, "Resumo automático", "N")
    for k, col in enumerate("EFGH", 1):
        summary(ws, s + k, f"Canais com envio ao FX {k}", f"=COUNT({col}{E1}:{col}{E2})", "D", "E")
    alertas_fx = "+".join(f'COUNTIF(L{X1}:L{X2},"{t}")' for t, _ in SIT_FX[1:])
    alertas_env = "+".join(f'COUNTIF(K{E1}:K{E2},"{t}")' for t, _ in SIT_ENV[1:])
    summary(ws, s + 5, "Processadores com alerta", f"={alertas_fx}", "D", "E")
    summary(ws, s + 6, "Músicas com o delay fora do tempo", f'=COUNTIF(L{M1}:L{M2},"Fora do tempo")', "D", "E")
    summary(ws, s + 7, "Canais com alerta de envio", f"={alertas_env}", "D", "E")
    a = s + 8
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(COUNT(K4:K7)<4,"Informe os quatro valores de referência",'
                      f'IF(COUNTA(C{X1}:C{X2})+COUNTA(C{E1}:C{E2})+COUNT(D{M1}:D{M2})=0,"Preencha os processadores, as músicas ou os envios",'
                      f'IF(E{s + 5}>0,"Há processador com alerta: confira o caminho do efeito",'
                      f'IF(E{s + 7}>0,"Há canal com alerta de envio",'
                      f'IF(E{s + 6}>0,"Há música com o delay fora do tempo","OK")))))',
         b=False, merge=f"E{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:M{a}")
    hint(ws, f"B{a + 1}", "Os limites de tempo, de repetição e de envio são regra prática deste material, e não norma. O ouvido, na sala, decide.",
         merge=f"B{a + 1}:N{a + 1}", h="left")
    end = a + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "N")
        r = end + 3
        sub(ws, f"B{r}", "Onde", merge=f"B{r}:C{r}")
        sub(ws, f"D{r}", "O que foi feito", merge=f"D{r}:N{r}")
        ws.row_dimensions[r].height = 21.75
        for onde, feito in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", onde, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:N{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = "D4"
    setup(ws, tab, f"B1:N{end}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Dinâmica e efeitos — modelo de ficha de dinâmica e de efeitos"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Dinâmica e efeitos — como usar esta planilha",
           "Ficha dos compressores e dos gates de cada canal e ficha dos efeitos, para usar no fim da passagem de som.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo de efeito, tipo de envio, figura do delay, tipo de fonte e as respostas Sim ou Não aceitam só as opções da lista."),
        ("Números", "Digite só o número, com o sinal de menos quando houver: -20, e não “-20 dB”. A redução lida é positiva: 6."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Na mesa: ajuste o ganho e o equalizador de cada canal antes de ligar qualquer processador."),
        ("2", "Aba Dinâmica por canal: liste só os canais que vão usar compressor e anote threshold, razão, attack, release e make-up."),
        ("3", "Com a fonte no trecho mais forte, anote o pico de entrada do canal e a redução que o medidor GR marcou."),
        ("4", "Veja a situação e o que fazer. Corrija na mesa e troque os números pela nova leitura."),
        ("5", "Gates: anote o vazamento e o som útil mais fraco de cada peça. Compare o seu threshold com o sugerido."),
        ("6", "Aba Efeitos: anote cada processador e o caminho dele, o BPM das músicas que usam delay e os envios de cada canal."),
        ("7", "Aba Checklist: confira os doze itens no fim da passagem de som."),
    ]),
    ("Abas da planilha", [
        ("Dinâmica por canal", "Ficha de até 16 compressores, com a redução e o pico de saída esperados pela razão, e de até 8 gates."),
        ("Efeitos", "Até 4 processadores, o tempo do delay de até 6 músicas e os envios de até 16 canais, com alertas."),
        ("Referência", "Pontos de partida por fonte, razões, attack e release, gate, tipos de reverb e tempos de delay."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Dinâmica", "Os compressores e os gates de um culto, antes da correção (exemplo 1 do treinamento)."),
        ("Exemplo - Efeitos", "Os efeitos, as músicas e os envios do mesmo culto, antes da correção (exemplo 2 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Conta da razão", "Saída = threshold + (entrada − threshold) ÷ razão. A redução esperada é o que passa do threshold, menos esse valor dividido pela razão."),
        ("Redução lida e esperada", "A situação usa a redução lida no medidor GR. Sem ela, usa a esperada. O medidor costuma marcar um pouco menos, por causa do attack e do knee."),
        ("Faixas de redução", "O modelo parte de 3 a 6 dB como faixa de trabalho e de 10 dB como limite de excesso. São ponto de partida: troque os valores no alto da aba."),
        ("Make-up", "Recebe aviso quando passa da redução mais a folga. Cada dB de make-up sobe os trechos fracos e reduz a margem antes da microfonia."),
        ("Gate", "O threshold sugerido é o meio da janela entre o vazamento e o som útil mais fraco. Com janela menor que duas margens, o gate não separa os dois."),
        ("Tempo do delay", "Semínima = 60 000 ÷ BPM, em ms. Colcheia: metade. Colcheia pontuada: três quartos. Semicolcheia: um quarto."),
        ("Efeitos", "Envio pré-fader, retorno que volta ao próprio envio, efeito em fala, em graves e em mídia recebem alerta. São regras práticas deste material."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_din(wb.create_sheet("Dinâmica por canal"), RED_H)
aba_fx(wb.create_sheet("Efeitos"), PURPLE_H)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 34, "C": 16, "D": 16, "E": 16, "F": 20, "G": 62, "H": 2})
title(ws, "Referência — pontos de partida", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "G")


def tabela(r, titulo, cab, linhas, alt=31.5, color=INK):
    """Tabela de seis colunas (B a G). Linhas com menos células juntam as colunas do meio."""
    band(ws, r, titulo, "G", color=color)
    r += 1
    for col, t in zip("BCDEFG", cab):
        sub(ws, f"{col}{r}", t)
    ws.row_dimensions[r].height = 21.75
    for lin in linhas:
        r += 1
        put(ws, f"B{r}", lin[0], f=font(10, True))
        if len(lin) == 3:
            put(ws, f"C{r}", lin[1], h="center", merge=f"C{r}:F{r}")
        else:
            for col, t in zip("CDEF", lin[1:5]):
                put(ws, f"{col}{r}", t, h="center")
        put(ws, f"G{r}", lin[-1])
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(4, "Compressor: ponto de partida por fonte", ["Fonte", "Razão", "Attack", "Release", "Redução nos trechos fortes", "Observação"], [
    ("Voz falada (pregação, palestra)", "2:1 a 3:1", "10 a 30 ms", "100 a 300 ms", "3 a 6 dB", "Só para segurar os gritos e as frases fortes. Com lapela ou púlpito, use menos."),
    ("Voz cantada (principal e apoio)", "3:1 a 4:1", "5 a 20 ms", "100 a 250 ms", "3 a 6 dB", "Knee suave, se houver. É onde o compressor mais ajuda."),
    ("Violão por direct box", "2:1 a 4:1", "10 a 30 ms", "100 a 250 ms", "2 a 5 dB", "Iguala o dedilhado e a batida forte."),
    ("Baixo", "3:1 a 6:1", "20 a 40 ms", "150 a 400 ms", "3 a 8 dB", "Attack mais lento, para manter o início da nota. Release rápido demais distorce os graves."),
    ("Bumbo", "4:1", "10 a 30 ms", "50 a 150 ms", "3 a 6 dB", "O attack deixa o estalo da batida passar. Muitas vezes o gate ajuda mais."),
    ("Caixa da bateria", "3:1 a 4:1", "5 a 20 ms", "50 a 150 ms", "3 a 6 dB", "Só se os golpes forem muito desiguais."),
    ("Teclado", "2:1", "20 a 40 ms", "200 a 400 ms", "0 a 3 dB", "Quase nunca precisa. Em par estéreo, ligue o link."),
    ("Coral, púlpito distante, ambiente", "Sem compressor", "Microfone longe da fonte é o primeiro a apitar."),
    ("Computador, vídeo, trilha", "Sem compressor", "O material já vem finalizado. Acerte o nível no fader."),
], color=RED_H)
r = tabela(r, "Razão: como age", ["Razão", "Como age", "16 dB acima viram", "Redução", "Nome usual", "Uso típico"], [
    ("1:1", "não comprime", "16 dB", "0 dB", "desligado", "É o mesmo que desligado."),
    ("2:1", "suave", "8 dB", "8 dB", "compressão", "Voz falada, violão, fontes que só precisam de um pouco de controle."),
    ("4:1", "moderada", "4 dB", "12 dB", "compressão", "Voz cantada, baixo, bumbo. É o ponto de partida mais comum."),
    ("8:1", "forte", "2 dB", "14 dB", "compressão forte", "Fontes muito irregulares. Ao vivo, pede cuidado com a microfonia."),
    ("10:1 ou mais", "teto", "1,6 dB ou menos", "14,4 dB ou mais", "limitação", "Proteção de saídas. O compressor já trabalha como limiter."),
], alt=21.75, color=RED_H)
r = tabela(r, "Attack e release: o que ouvir", ["Controle", "Faixa típica", "Rápido demais", "Lento demais", "Para começar", "Como corrigir"], [
    ("Attack", "1 a 50 ms", "achata o começo", "os picos passam", "10 a 20 ms", "Som sem definição no começo: alongue. Picos saltando e medidor parado: encurte."),
    ("Release", "50 a 500 ms", "bombeia, distorce", "não solta", "100 a 200 ms", "Som bombeando ou ruído respirando: alongue. Medidor que não volta a zero: encurte ou suba o threshold."),
], color=RED_H)
r = tabela(r, "Gate na bateria: ponto de partida", ["Peça", "Range", "Attack", "Hold", "Release", "Threshold"], [
    ("Bumbo", "10 a 20 dB", "abaixo de 1 ms", "30 a 80 ms", "80 a 200 ms", "Entre o vazamento mais forte (baixo, caixa) e o golpe mais fraco do bumbo."),
    ("Caixa", "6 a 15 dB", "abaixo de 1 ms", "50 a 100 ms", "100 a 250 ms", "Cuidado com as notas fracas (ghost notes): com gate fundo, elas somem."),
    ("Tons e surdo", "10 a 20 dB", "abaixo de 1 ms", "50 a 150 ms", "200 a 400 ms", "Entre a caixa que vaza e o golpe mais fraco do tom. Release longo, para a nota soar inteira."),
    ("Voz", "Sem gate", "Corta o início e o fim das palavras. Se precisar, expander com range de 6 a 10 dB."),
], color=SLATE)
r = tabela(r, "Reverb: tipos e ajustes", ["Tipo", "Tempo típico", "Pré-delay", "Corte de graves", "Corte de agudos", "Uso ao vivo"], [
    ("Room (sala)", "0,4 a 1 s", "0 a 20 ms", "150 a 300 Hz", "4 a 8 kHz", "Bateria, violão, músicas rápidas. Tira o som seco sem aparecer."),
    ("Plate (placa)", "1 a 2,5 s", "20 a 60 ms", "150 a 300 Hz", "4 a 8 kHz", "Voz cantada e caixa da bateria. É a escolha mais segura para começar."),
    ("Hall (salão)", "1,5 a 3 s", "20 a 80 ms", "150 a 300 Hz", "4 a 8 kHz", "Músicas lentas, coral, cordas. Embola com facilidade em música rápida."),
    ("Fala (pregação, avisos, leitura)", "Sem efeito", "Reverb e delay atrapalham o entendimento da fala. Mute nos retornos de efeito."),
], color=PURPLE_H)
r = tabela(r, "Delay: tempo pelo andamento (60 000 ÷ BPM)", ["Andamento", "Semínima", "Colcheia", "Colcheia pontuada", "Semicolcheia", "Onde aparece"], [
    ("60 BPM", "1000 ms", "500 ms", "750 ms", "250 ms", "Música muito lenta."),
    ("72 BPM", "833 ms", "417 ms", "625 ms", "208 ms", "Louvor lento."),
    ("90 BPM", "667 ms", "333 ms", "500 ms", "167 ms", "Andamento médio."),
    ("100 BPM", "600 ms", "300 ms", "450 ms", "150 ms", "Andamento médio."),
    ("120 BPM", "500 ms", "250 ms", "375 ms", "125 ms", "Música animada."),
    ("140 BPM", "429 ms", "214 ms", "321 ms", "107 ms", "Música rápida."),
    ("Eco curto (slapback)", "80 a 120 ms, uma repetição", "Dá corpo à voz sem deixar cauda. Não depende do andamento."),
], alt=21.75, color=PURPLE_H)
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com a fonte, o microfone, a música, a sala e o modelo do processador. Confie mais no medidor de redução e no ouvido do que na posição dos botões.",
     merge=f"B{r - 1}:G{r - 1}", h="left")
setup(ws, SLATE, f"B1:G{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplos
checklist(wb.create_sheet("Checklist"), "Checklist de dinâmica e efeitos",
          "Confira no fim da passagem de som, com a banda tocando. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["blue"][0])
aba_din(wb.create_sheet("Exemplo - Dinâmica"), MUTED, EX_DIN)
aba_fx(wb.create_sheet("Exemplo - Efeitos"), MUTED, EX_FX)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
