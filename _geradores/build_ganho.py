# -*- coding: utf-8 -*-
"""Gera Ganho-e-Fader-modelo.xlsx: registro de ganho por canal, cadeia de ganho, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIM_NAO = ["Sim", "Não"]
SIT_CF = [("OK", GREEN), ("Baixo", YELLOW), ("Alto", YELLOW), ("Fader baixo", YELLOW), ("Fader alto", YELLOW), ("Clip", RED)]
N = 24  # canais na aba de ganho

CHECK = [
    ("Preparação", "A mesa foi zerada: faders e master embaixo, ganhos no mínimo, nenhum PFL apertado."),
    ("Preparação", "Cada fonte está na entrada certa: microfone, linha ou direct box."),
    ("Preparação", "O pad está ligado só nos canais que precisam dele."),
    ("Ganho", "O ganho de cada canal foi ajustado pelo medidor, com o PFL, um canal por vez."),
    ("Ganho", "A leitura foi feita com a fonte no volume real do evento, no trecho mais forte."),
    ("Ganho", "A média de cada canal está perto do nível nominal (−18 dBFS ou 0 VU)."),
    ("Ganho", "Os picos ficam abaixo do teto (perto de −6 dBFS), e a luz de clip não acende."),
    ("Ganho", "O ganho, o pad e as leituras de cada canal estão anotados."),
    ("Cadeia", "Os faders dos canais principais trabalham perto de 0 dB."),
    ("Cadeia", "O master e os grupos estão perto de 0 dB, e o medidor principal não chega ao vermelho."),
    ("Cadeia", "O nível do amplificador ou da caixa ativa foi ajustado com a mesa no nível nominal, e a luz de limite não fica acesa."),
    ("Operação", "Os retornos foram montados depois do ganho, e o ganho não muda durante o evento."),
]

# Exemplo 1 do treinamento: passagem de som de um culto, mesa digital (leituras em dBFS)
EX_GANHO = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "mesa": "Mesa digital de 16 canais, medidor em dBFS"},
    "ref": (-18, -6, 0, 3),
    "canais": [  # fonte, entrada, pad, ganho, média, pico, luz de clip
        ("Voz do pastor (mic de mão)", "Microfone", "Não", 20, -31, -19, "Não"),
        ("Voz da cantora (mic de mão)", "Microfone", "Não", 38, -18, -7, "Não"),
        ("Voz de apoio 1", "Microfone", "Não", 40, -20, -9, "Não"),
        ("Voz de apoio 2", "Microfone", "Não", 46, -10, -1, "Não"),
        ("Violão", "DI", "Não", 25, -19, -8, "Não"),
        ("Baixo", "DI", "Não", 18, -16, -9, "Não"),
        ("Teclado E", "Linha", "Não", 5, -6, 0, "Sim"),
        ("Teclado D", "Linha", "Não", 5, -6, 0, "Sim"),
        ("Bumbo", "Microfone", "Não", 20, -24, -5, "Não"),
        ("Caixa da bateria", "Microfone", "Não", 22, -26, -2, "Não"),
        ("Computador (som dos vídeos)", "Linha", "Não", 0, -20, -10, "Não"),
        ("Microfone do púlpito", "Microfone", "Não", 30, -27, -14, "Não"),
    ],
    "ajustes": [  # canal, o que foi feito, leitura depois
        ("1", "Ganho de 20 para 33 dB. O fader voltou do topo para perto de 0 dB.", "média −18, pico −6"),
        ("4", "Ganho de 46 para 38 dB.", "média −18, pico −9"),
        ("7 e 8", "O ganho já estava no mínimo: pad ligado e volume do teclado em três quartos.", "média −19, pico −9"),
        ("10", "Ganho de 22 para 18 dB. Em fonte percussiva, quem manda é o pico.", "média −30, pico −6"),
        ("12", "Ganho de 30 para 38 dB.", "média −19, pico −6"),
        ("Todos", "Retornos montados só depois da correção dos ganhos.", ""),
    ],
}

# Exemplo 2 do treinamento: chiado em uma formatura, mesa analógica (leituras na escala VU)
EX_CADEIA = {
    "head": {"evento": "Formatura no auditório da escola", "sinal": "Orador falando forte, a um palmo do microfone",
             "data": date(2026, 12, 12), "operador": "Operador de som do evento"},
    "ref": (0, 6, 10, 3),
    "etapas": [  # posição, média, pico, luz
        (15, -20, -8, "Não"),
        (0, None, None, None),
        (None, None, None, None),
        (-5, -25, -13, "Não"),
        (None, None, None, None),
        ("máximo", None, None, "Não"),
    ],
    "ajustes": [  # etapa, antes, depois
        ("Pré do canal", "Ganho em 15 dB. PFL: média −20, pico −8.", "Ganho em 29 dB. PFL: média −6, pico +6."),
        ("Master", "Fader em −5 dB. Medidor: média −25, pico −13.", "Fader em 0 dB. Medidor: média −6, pico +6."),
        ("Caixas ativas", "Controle de nível no máximo.", "Fechadas até o volume da sala voltar ao que era."),
        ("Resultado", "Chiado contínuo nos silêncios.", "Mesmo volume na sala; o chiado deixou de ser ouvido da plateia."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def cf_status(ws, rng, first):
    """Verde para OK, amarelo para qualquer outro texto, nada para a célula vazia."""
    cf_equal(ws, rng, [("OK", GREEN)])
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(LEN({first})>0,{first}<>"OK")'],
                                                   fill=PatternFill("solid", bgColor=YELLOW, fgColor=YELLOW)))


def ajuste(media, pico, alvo, teto):
    """Quantos dB faltam (positivo) ou sobram (negativo): vale o menor entre a média e o pico."""
    return f'ROUND(IF({media}="",{teto}-{pico},IF({pico}="",{alvo}-{media},MIN({alvo}-{media},{teto}-{pico}))),0)'


def referencias(ws, lab1, lab2, val, hint1, hint2, ref, bg):
    """Bloco dos quatro valores de referência, à direita do cabeçalho. Devolve as referências absolutas."""
    linhas = [
        ("Alvo da média", ref[0], "Mesa digital: −18 dBFS. Medidor com escala VU: 0."),
        ("Teto dos picos", ref[1], "Mesa digital: −6 dBFS. Medidor com escala VU: +6."),
        ("Limite (clip)", ref[2], "Mesa digital: 0 dBFS. Escala VU: o último LED da escala."),
        ("Tolerância (dB)", ref[3], "Diferença que ainda conta como OK."),
    ]
    for k, (rot, v, dica) in enumerate(linhas):
        r = 4 + k
        label(ws, f"{lab1}{r}", rot, merge=f"{lab1}{r}:{lab2}{r}")
        inp(ws, f"{val}{r}", v, h="center", bg=bg, fmt='+0;-0;0' if k < 3 else None)
        hint(ws, f"{hint1}{r}", dica, merge=f"{hint1}{r}:{hint2}{r}", h="left")
    dv_number(ws, f"{val}4:{val}7")
    return [f"${val}${4 + k}" for k in range(4)]


# ================================================================== aba Ganho por canal
G_W = {"A": 2, "B": 7, "C": 30, "D": 13, "E": 8, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 9, "L": 9, "M": 10,
       "N": 11, "O": 44, "P": 30, "Q": 2}
G1 = 14


def aba_ganho(ws, tab, data=None):
    widths(ws, G_W)
    is_ex = data is not None
    title(ws, "Ganho por canal — registro da passagem de som",
          "Exemplo preenchido, para consulta. Use a aba Ganho por canal para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro com o que você leu no PFL. As células cinza são calculadas.", "P")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None), (8, "Mesa e medidor", "mesa", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    ALVO, TETO, LIM, TOL = referencias(ws, "I", "K", "L", "M", "P", data["ref"] if is_ex else (-18, -6, 0, 3), hb)
    hint(ws, "I8", "Leia cada canal no PFL, um por vez, com a fonte no volume real do evento.", merge="I8:P8", h="left")

    band(ws, 10, "Registro de ganho por canal", "P")
    heads = [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Entrada", INK), ("E", "Pad", INK),
             ("F", "Ganho", AMBER), ("G", "Média", AMBER), ("H", "Pico", AMBER), ("I", "Luz de clip", AMBER),
             ("J", "Crista", TEAL), ("K", "Folga", TEAL), ("L", "Ajuste", TEAL), ("M", "Ganho sugerido", TEAL),
             ("N", "Situação", TEAL), ("O", "O que fazer", TEAL), ("P", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMNOP", ["nº", "o que está ligado no canal", "tipo", "Sim / Não", "dB, no botão",
                                          "lida no PFL", "lido no PFL", "Sim / Não", "pico − média", "limite − pico",
                                          "dB", "dB", "", "", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 27
    if is_ex:
        hint(ws, "B13", "Leituras em dBFS, feitas no PFL com cada fonte no trecho mais forte. As colunas de J a O são calculadas.",
             merge="B13:P13", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNO", ["Ex.", "Voz da cantora (mic de mão)", "Microfone", "Não", 38, -24, -12, "Não",
                                             12, 12, 6, 44, "Baixo", "Suba o ganho cerca de 6 dB"]):
            ex(ws, f"{col}13", v, h="left" if col in "CO" else "center")
        ex(ws, "P13", "")
        note(ws, "F11", "Valor do botão de ganho, em dB.\nSe a mesa não mostra o número, anote a posição do botão na coluna Observação.")
        note(ws, "G11", "Onde o medidor passa a maior parte do tempo, com a fonte no volume real.\nDigite só o número: −18, por exemplo.")
        note(ws, "H11", "O ponto mais alto que o medidor alcançou no trecho mais forte.")
        note(ws, "I11", "A luz de clip (peak, OL) do canal acendeu durante a leitura?")
        note(ws, "L11", "Quantos dB subir (positivo) ou baixar (negativo) para chegar ao alvo.\nVale o menor entre o que falta para a média e o que falta para o teto dos picos.")
        note(ws, "M11", "Ganho anotado mais o ajuste.\nEm caso de clip fica em branco: o pico real é desconhecido. Baixe uns 6 dB e meça de novo.")
    ws.row_dimensions[13].height = 24

    G2 = G1 + N - 1
    for k in range(N):
        r = G1 + k
        c = data["canais"][k] if is_ex and k < len(data["canais"]) else None
        bg = WHITE if is_ex else INPUT
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        for col, j in zip("DEFGHI", range(1, 7)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt='+0;-0;0' if col in "GH" else None)
        adj = ajuste(f"G{r}", f"H{r}", ALVO, TETO)
        calc(ws, f"J{r}", f'=IF(OR(G{r}="",H{r}=""),"",H{r}-G{r})', b=False)
        calc(ws, f"K{r}", f'=IF(H{r}="","",{LIM}-H{r})', b=False)
        clip = f'OR(I{r}="Sim",AND(H{r}<>"",H{r}>={LIM}))'
        calc(ws, f"L{r}", f'=IF(OR(AND(G{r}="",H{r}=""),{clip}),"",{adj})', fmt='+0;-0;0')
        calc(ws, f"M{r}", f'=IF(OR(F{r}="",L{r}=""),"",F{r}+L{r})')
        calc(ws, f"N{r}", f'=IF(I{r}="Sim","Clip",IF(AND(G{r}="",H{r}=""),"",IF(AND(H{r}<>"",H{r}>={LIM}),"Clip",'
                          f'IF(L{r}<-{TOL},"Alto",IF(L{r}>{TOL},"Baixo","OK")))))')
        calc(ws, f"O{r}", f'=IF(N{r}="","",IF(N{r}="Clip",IF(E{r}="Sim","Baixe o ganho uns 6 dB e meça de novo",'
                          f'"Baixe o ganho uns 6 dB e meça de novo. No mínimo: ligue o pad"),'
                          f'IF(N{r}="Alto","Baixe o ganho cerca de "&ABS(L{r})&" dB",'
                          f'IF(N{r}="Baixo",IF(E{r}="Sim","Veja se o pad é necessário. Sem ele, meça de novo",'
                          f'"Suba o ganho cerca de "&L{r}&" dB"),"Manter"))))', h="left", b=False, sz=9)
        inp(ws, f"P{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{G1}:D{G2}", ["Microfone", "Linha", "DI", "Sem fio", "Digital"], "Microfone, Linha, DI, Sem fio ou Digital.")
    dv_list(ws, f"E{G1}:E{G2}", SIM_NAO, "O pad do canal está ligado?")
    dv_list(ws, f"I{G1}:I{G2}", SIM_NAO, "A luz de clip acendeu durante a leitura?")
    dv_number(ws, f"F{G1}:H{G2}")
    cf_equal(ws, f"N{G1}:N{G2}", SIT_CF)

    s = G2 + 2
    band(ws, s, "Resumo automático", "P")
    cont = lambda t: f'COUNTIF(N{G1}:N{G2},"{t}")'
    linhas = [
        ("Canais registrados", f"=COUNTA(C{G1}:C{G2})"),
        ("Canais com leitura", f'={cont("OK")}+{cont("Baixo")}+{cont("Alto")}+{cont("Clip")}'),
        ("Na faixa (OK)", f'={cont("OK")}'),
        ("Ganho baixo", f'={cont("Baixo")}'),
        ("Ganho alto", f'={cont("Alto")}'),
        ("Em clip", f'={cont("Clip")}'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(COUNT({ALVO[1:2]}4:{ALVO[1:2]}7)<4,"Informe os quatro valores de referência",'
                      f'IF(F{s + 1}=0,"Liste os canais",IF(F{s + 6}>0,"Há canal em clip: corrija antes de mixar",'
                      f'IF(F{s + 2}<F{s + 1},"Há canal sem leitura",IF(F{s + 4}+F{s + 5}>0,"Há ganho a ajustar","OK")))))',
         b=False, merge=f"F{a}:O{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:O{a}")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "P")
        r = end + 3
        put(ws, f"B{r}", "Canal", f=font(10, True), bg=GRAY, h="center")
        put(ws, f"C{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:M{r}")
        put(ws, f"N{r}", "Leitura depois (dBFS)", f=font(10, True), bg=GRAY, h="center", merge=f"N{r}:P{r}")
        ws.row_dimensions[r].height = 21.75
        for canal, feito, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", canal, f=font(10, True, c=MUTED), bg=GRAY, h="center")
            put(ws, f"C{r}", feito, merge=f"C{r}:M{r}")
            put(ws, f"N{r}", depois, merge=f"N{r}:P{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{G1}"
    setup(ws, tab, f"B1:P{end}", fit_height=True)


# ================================================================== aba Cadeia de ganho
C_W = {"A": 2, "B": 5, "C": 28, "D": 30, "E": 13, "F": 10, "G": 10, "H": 14, "I": 14, "J": 52, "K": 30, "L": 2}
ETAPAS = [  # etapa, controle, tipo de conferência
    ("Pré do canal", "Ganho (leitura no PFL)", "medidor"),
    ("Canal", "Fader do canal", "fader"),
    ("Grupo (se houver)", "Fader do grupo", "fader"),
    ("Master", "Fader master (medidor principal)", "master"),
    ("Processador (se houver)", "Ganho de entrada e de saída", "luz"),
    ("Amplificador ou caixa ativa", "Controle de nível", "luz"),
]
C1 = 12


def aba_cadeia(ws, tab, data=None):
    widths(ws, C_W)
    is_ex = data is not None
    title(ws, "Cadeia de ganho — do pré à caixa",
          "Exemplo preenchido, para consulta. Use a aba Cadeia de ganho para o seu sistema." if is_ex
          else "Com um sinal de teste no volume real, anote a posição de cada controle e o que os medidores mostram.", "K")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Sistema ou evento", "evento", None), (5, "Sinal de teste", "sinal", None),
                           (6, "Data", "data", DATE), (7, "Operador", "operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:E{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    ALVO, TETO, LIM, TOL = referencias(ws, "F", "H", "I", "J", "K", data["ref"] if is_ex else (-18, -6, 0, 3), hb)

    band(ws, 9, "Cadeia de ganho, da entrada para a saída", "K")
    for col, t, bg in [("B", "#", INK), ("C", "Etapa", INK), ("D", "Controle", INK), ("E", "Posição", AMBER),
                       ("F", "Média", AMBER), ("G", "Pico", AMBER), ("H", "Luz de clip ou de limite", AMBER),
                       ("I", "Situação", TEAL), ("J", "O que fazer", TEAL), ("K", "Observação", INK)]:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJK", ["", "", "", "dB", "no medidor", "no medidor", "Sim / Não", "", "", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 21.75
    if not is_ex:
        note(ws, "E10", "Valor marcado no controle, em dB. No fader, 0 é o ganho unitário.\nNo amplificador ou na caixa ativa, pode ser um texto: máximo, meio, 12 h.")
        note(ws, "H10", "Acendeu a luz de clip (peak, OL) da etapa, ou a luz de limite (limit) da caixa ou do amplificador?")
        note(ws, "I10", "Fader baixo: abaixo de −10 dB. Fader alto: acima de +5 dB. No master, o aviso de nível baixo só aparece quando faltam mais de 10 dB.")

    bg = WHITE if is_ex else INPUT
    for k, (etapa, controle, tipo) in enumerate(ETAPAS):
        r = C1 + k
        v = data["etapas"][k] if is_ex else (None, None, None, None)
        num(ws, f"B{r}", k + 1)
        put(ws, f"C{r}", etapa, f=font(10, True), bg=GRAY)
        put(ws, f"D{r}", controle, bg=GRAY)
        inp(ws, f"E{r}", v[0], h="center", bg=bg, fmt='+0;-0;0')
        tem_medidor, tem_luz = tipo in ("medidor", "master"), tipo != "fader"
        for col, val, usa in (("F", v[1], tem_medidor), ("G", v[2], tem_medidor), ("H", v[3], tem_luz)):
            if usa:
                inp(ws, f"{col}{r}", val, h="center", bg=bg, fmt='+0;-0;0' if col in "FG" else None)
            else:
                hint(ws, f"{col}{r}", "–")
        E, F, G, Hc, I = (f"{c}{r}" for c in "EFGHI")
        adj = ajuste(F, G, ALVO, TETO)
        if tipo == "medidor":
            sit = (f'=IF({Hc}="Sim","Clip",IF(AND({F}="",{G}=""),"",IF(AND({G}<>"",{G}>={LIM}),"Clip",'
                   f'IF({adj}<-{TOL},"Alto",IF({adj}>{TOL},"Baixo","OK")))))')
            acao = (f'=IF({I}="","",IF({I}="Clip","Baixe o ganho uns 6 dB e meça de novo. No mínimo: ligue o pad",'
                    f'IF({I}="Alto","Baixe o ganho cerca de "&ABS({adj})&" dB",'
                    f'IF({I}="Baixo","Suba o ganho cerca de "&{adj}&" dB","Manter"))))')
        elif tipo == "fader":
            sit = f'=IF({E}="","",IF({E}<-10,"Fader baixo",IF({E}>5,"Fader alto","OK")))'
            acao = (f'=IF({I}="","",IF({I}="Fader baixo","Sobra nível antes deste fader: confira o ganho dos canais",'
                    f'IF({I}="Fader alto","Falta nível antes deste fader: confira o ganho dos canais","Manter")))')
        elif tipo == "master":
            lido = f'OR({F}<>"",{G}<>"")'
            sit = (f'=IF({Hc}="Sim","Clip",IF(AND({G}<>"",{G}>={LIM}),"Clip",IF(AND({lido},{adj}<-{TOL}),"Alto",'
                   f'IF(AND({E}<>"",{E}<-10),"Fader baixo",IF(AND({E}<>"",{E}>5),"Fader alto",'
                   f'IF(AND({lido},{adj}>10),"Baixo",IF(OR({E}<>"",{lido}),"OK","")))))))')
            acao = (f'=IF({I}="","",IF({I}="Clip","Baixe os faders dos canais ou o master. Os ganhos ficam como estão",'
                    f'IF({I}="Alto","Mistura acima do nível nominal: baixe os faders dos canais ou o master",'
                    f'IF({I}="Fader baixo","Feche o amplificador ou a caixa e traga o master para perto de 0 dB",'
                    f'IF({I}="Fader alto","Confira os ganhos e os faders. Se estão certos, abra a caixa; no máximo, falta sistema",'
                    f'IF({I}="Baixo","Saída abaixo do nível nominal: confira os ganhos, suba a mistura e feche a caixa na mesma medida",'
                    f'"Manter"))))))')
        else:
            sit = f'=IF({Hc}="Sim","Clip",IF({Hc}="Não","OK",""))'
            acao = (f'=IF({I}="","",IF({I}="Clip","'
                    + ("Baixe o nível que chega ao processador" if k == 4 else "O sistema chegou ao limite: baixe o master")
                    + '","Manter"))')
        calc(ws, I, sit)
        calc(ws, f"J{r}", acao, h="left", b=False, sz=9)
        inp(ws, f"K{r}", bg=bg)
        ws.row_dimensions[r].height = 30
    C2 = C1 + len(ETAPAS) - 1
    dv_number(ws, f"E{C1}:G{C2 - 1}")
    dv_list(ws, f"H{C1}", SIM_NAO, "A luz de clip acendeu?")
    dv_list(ws, f"H{C1 + 3}:H{C2}", SIM_NAO, "A luz de clip ou de limite acendeu?")
    cf_equal(ws, f"I{C1}:I{C2}", SIT_CF)

    s = C2 + 2
    band(ws, s, "Leitura da cadeia", "K")
    fora = "+".join(f'COUNTIF(I{C1}:I{C2},"{t}")' for t in ("Baixo", "Alto", "Fader baixo", "Fader alto"))
    summary(ws, s + 1, "Soma dos faders: canal + grupo + master (dB)", f'=IF(COUNT(E{C1 + 1}:E{C1 + 3})=0,"",SUM(E{C1 + 1}:E{C1 + 3}))',
            "D", "E", fmt='+0;-0;0')
    label(ws, f"B{s + 2}", "Leitura da soma", merge=f"B{s + 2}:D{s + 2}", h="right")
    calc(ws, f"E{s + 2}", f'=IF(E{s + 1}="","",IF(E{s + 1}<-20,"Faders muito baixos: sobra ganho no início, ou a caixa está aberta demais",'
                          f'IF(E{s + 1}>10,"Faders muito altos: falta ganho no início, ou falta sistema","OK")))',
         b=False, merge=f"E{s + 2}:J{s + 2}", h="left")
    summary(ws, s + 3, "Etapas em clip", f'=COUNTIF(I{C1}:I{C2},"Clip")', "D", "E")
    summary(ws, s + 4, "Etapas fora da faixa", f"={fora}", "D", "E")
    label(ws, f"B{s + 5}", "Aviso", merge=f"B{s + 5}:D{s + 5}", h="right")
    calc(ws, f"E{s + 5}", f'=IF(COUNTIF(I{C1}:I{C2},"OK")+E{s + 3}+E{s + 4}=0,"Preencha as etapas",'
                          f'IF(E{s + 3}>0,"Há etapa em clip: corrija da entrada para a saída",'
                          f'IF(E{s + 4}>0,"Há etapa fora da faixa: comece pela primeira da lista","OK")))',
         b=False, merge=f"E{s + 5}:J{s + 5}", h="left")
    for r in (s + 2, s + 5):
        ws.row_dimensions[r].height = 21.75
    cf_status(ws, f"E{s + 2}:J{s + 2}", f"$E${s + 2}")
    cf_ok(ws, f"E{s + 5}:J{s + 5}")
    hint(ws, f"B{s + 6}", "Os cortes de −10 e +5 dB para um fader, e de −20 e +10 dB para a soma, são regra prática deste material, e não norma.",
         merge=f"B{s + 6}:K{s + 6}", h="left")
    end = s + 6

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "K")
        r = end + 3
        put(ws, f"B{r}", "Etapa", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "Antes", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:H{r}")
        put(ws, f"I{r}", "Depois", f=font(10, True), bg=GRAY, h="center", merge=f"I{r}:K{r}")
        ws.row_dimensions[r].height = 21.75
        for etapa, antes, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", etapa, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", antes, merge=f"D{r}:H{r}")
            put(ws, f"I{r}", depois, merge=f"I{r}:K{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    setup(ws, tab, f"B1:K{end}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Ganho e fader — modelo de registro de ganho"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Ganho e fader — como usar esta planilha",
           "Registro do ganho de cada canal e conferência da cadeia de ganho, para usar na passagem de som.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Entrada, pad e luz de clip aceitam só as opções da lista."),
        ("Números", "Digite as leituras só com o número, com o sinal de menos quando houver: -18, e não “-18 dBFS”."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Ganho por canal: preencha o cabeçalho e confira os quatro valores de referência do seu medidor."),
        ("2", "Aba Ganho por canal: liste a fonte e o tipo de entrada de cada canal."),
        ("3", "Na mesa: aperte o PFL do canal, peça o volume real e leve o ganho até a faixa de trabalho."),
        ("4", "Aba Ganho por canal: anote o ganho, o pad, a média, o pico e se a luz de clip acendeu."),
        ("5", "Veja a situação e o que fazer. Corrija na mesa, meça de novo e troque os números pela nova leitura."),
        ("6", "Aba Cadeia de ganho: com a mistura pronta, anote as posições e as leituras, do pré até a caixa."),
        ("7", "Aba Checklist: confira os doze itens antes de montar os retornos."),
    ]),
    ("Abas da planilha", [
        ("Ganho por canal", "Registro de até 24 canais, com situação calculada: OK, Baixo, Alto ou Clip."),
        ("Cadeia de ganho", "Seis etapas, do pré ao amplificador ou à caixa ativa, com a situação de cada uma e a soma dos faders."),
        ("Referência", "Faixas de leitura por tipo de medidor, ganho de partida por fonte e posição dos controles."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Ganho", "A passagem de som de um culto com doze canais, em mesa digital (exemplo 1 do treinamento)."),
        ("Exemplo - Cadeia", "A cadeia de uma formatura com chiado nas caixas, em mesa analógica (exemplo 2 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Valores de referência", "O modelo parte de média em −18 dBFS, picos até −6 dBFS e clip em 0 dBFS, comuns em mesas digitais. São ponto de partida: confira o manual da sua mesa e troque os valores no alto da aba."),
        ("Medidor com escala VU", "Em mesa analógica, use alvo 0, teto +6 e, como limite, o valor do último LED da escala. O LED de pico do canal entra na coluna Luz de clip."),
        ("Ajuste sugerido", "É o menor entre o que falta para a média chegar ao alvo e o que falta para o pico chegar ao teto. Em fonte percussiva, o pico costuma mandar."),
        ("Situação", "OK quando o ajuste fica dentro da tolerância. Baixo ou Alto quando passa dela. Clip quando o pico chega ao limite ou a luz de clip acende."),
        ("Em caso de clip", "O pico real é desconhecido, porque o medidor não passa do limite. O ajuste fica em branco: baixe uns 6 dB e meça de novo."),
        ("Cadeia de ganho", "Fader abaixo de −10 dB ou acima de +5 dB recebe aviso. São cortes de regra prática, para chamar a atenção, e não norma."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_ganho(wb.create_sheet("Ganho por canal"), AMBER)
aba_cadeia(wb.create_sheet("Cadeia de ganho"), TEAL)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 36, "C": 26, "D": 26, "E": 60, "F": 2})
title(ws, "Referência — faixas típicas", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")


def tabela(r, titulo, cab, linhas, alt=31.5):
    band(ws, r, titulo, "E")
    r += 1
    for col, t in zip("BCDE", cab):
        put(ws, f"{col}{r}", t, f=font(10, True), bg=GRAY, h="center")
    ws.row_dimensions[r].height = 21.75
    for lin in linhas:
        r += 1
        put(ws, f"B{r}", lin[0], f=font(10, True))
        for col, t in zip("CDE", lin[1:]):
            put(ws, f"{col}{r}", t, h="left" if col == "E" else "center")
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(4, "Faixas de leitura por tipo de medidor", ["Medidor", "Média", "Picos mais fortes", "Sinal de que passou"], [
    ("Digital (dBFS)", "perto de −18 dBFS", "até perto de −6 dBFS", "Luz de clip ou leitura em 0 dBFS. Em algumas mesas o nível nominal é −20 dBFS."),
    ("Analógico (escala VU)", "em torno de 0", "até perto de +6", "LED de pico (peak, clip, OL) aceso. Ele costuma acender alguns dB antes do clip."),
    ("Só dois LEDs por canal", "LED de sinal aceso", "LED de pico apagado", "LED de pico aceso no trecho mais forte."),
])
r = tabela(r, "Ganho de partida por fonte", ["Fonte", "Nível típico na entrada", "Ganho para começar", "Observação"], [
    ("Fala em microfone dinâmico", "−55 a −40 dBu", "45 a 60 dB", "É a fonte que mais pede ganho. Microfone perto da boca reduz o ganho necessário."),
    ("Canto perto do microfone dinâmico", "−40 a −25 dBu", "30 a 45 dB", "Ajuste com o trecho mais forte da música."),
    ("Microfone condensador, mesma fonte", "10 a 20 dB acima do dinâmico", "10 a 20 dB a menos", "Depende muito da distância: um coral captado de longe volta a pedir ganho alto."),
    ("Bumbo, caixa, amplificador de guitarra", "−20 a 0 dBu", "5 a 25 dB", "Se passar da faixa com o ganho no mínimo, ligue o pad."),
    ("Violão ou baixo por direct box", "−40 a −20 dBu", "20 a 40 dB", "A direct box entrega nível de microfone. Captação ativa chega mais forte."),
    ("Teclado, computador, tocador (linha)", "−10 a +4 dBu", "0 a 15 dB", "Use a entrada de linha. Na entrada de microfone, use direct box ou pad."),
])
r = tabela(r, "Média e pico por tipo de fonte", ["Tipo de fonte", "Pico acima da média", "O que guia o ajuste", "Exemplos"], [
    ("Sustentada", "poucos dB", "a média", "Teclado com som de cordas ou de órgão, baixo, trilha de vídeo já finalizada."),
    ("Voz", "10 a 15 dB", "a média e o pico", "Fala e canto. Peça o trecho mais forte."),
    ("Percussiva", "20 dB ou mais", "o pico", "Bumbo, caixa, tons, percussão, violão tocado com força."),
])
r = tabela(r, "Posição dos controles na cadeia", ["Etapa", "Onde deixar", "Faixa prática", "O que conferir"], [
    ("Pré do canal (ganho)", "pelo medidor, no PFL", "–", "Média perto do nominal, picos abaixo do teto, luz de clip apagada."),
    ("Pad", "desligado", "em geral −20 dB", "Ligue só se o canal passar da faixa com o ganho no mínimo."),
    ("Fader do canal", "perto de 0 dB", "−10 a +5 dB", "Fader no topo ou lá embaixo indica ganho errado."),
    ("Grupo", "0 dB", "−10 a +5 dB", "Se o medidor do grupo passar do teto, baixe os faders dos canais."),
    ("Master", "0 dB", "−10 a +5 dB", "Medidor principal perto do nominal nos trechos fortes, sem chegar ao vermelho."),
    ("Amplificador ou caixa ativa", "no volume da sala", "–", "Ajuste com a mesa no nível nominal. Luz de limite nunca acesa direto."),
])
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com o microfone, a distância, a fonte e o modelo da mesa. As faixas dos faders são regra prática deste material.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplos
checklist(wb.create_sheet("Checklist"), "Checklist do ajuste de ganho",
          "Confira na passagem de som, antes de montar os retornos. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["blue"][0])
aba_ganho(wb.create_sheet("Exemplo - Ganho"), MUTED, EX_GANHO)
aba_cadeia(wb.create_sheet("Exemplo - Cadeia"), MUTED, EX_CADEIA)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
