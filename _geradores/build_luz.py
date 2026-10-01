# -*- coding: utf-8 -*-
"""Gera Iluminacao-modelo.xlsx: plano de luz com ângulos e contraste, carga elétrica, DMX, referência e checklist."""
import math
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

AMBER, AMBER_T = HUE["amber"]
BLUE, BLUE_T = HUE["blue"]
TEAL, TEAL_T = HUE["teal"]
FUNCOES = ["Principal", "Preenchimento", "Contraluz", "Fundo", "Plateia", "Efeito"]
NP = 20   # refletores no plano de luz
NM = 6    # alvos na medição no rosto
NC = 8    # circuitos
NA = 24   # aparelhos na carga elétrica
ND = 24   # aparelhos DMX
INT_FMT = "0"
DEC1 = "0.0"

CHECK = [
    ("Plano", "Cada alvo (púlpito, banda, leitura) tem luz principal, preenchimento e contraluz anotados no plano de luz."),
    ("Ângulo", "A luz de frente chega ao rosto entre 30° e 45° acima da linha dos olhos, e os olhos aparecem na câmera."),
    ("Cor", "Todas as luzes brancas de rosto têm a mesma temperatura de cor, e o balanço de branco foi feito com elas acesas."),
    ("Cor", "Os refletores usados no rosto têm IRC de 90 ou mais."),
    ("Intensidade", "A luz no rosto foi medida, ou conferida na câmera, e não cai em nenhum ponto da área em que a pessoa anda."),
    ("Contraste", "A diferença entre o lado claro e o lado escuro do rosto está perto de 2:1, sem passar de 3:1."),
    ("Contraste", "O fundo está mais escuro que o rosto, e a contraluz separa a pessoa do fundo sem estourar."),
    ("Telão", "Nenhuma luz de frente bate no telão, e o brilho do telão foi ajustado olhando a imagem da câmera."),
    ("Câmera", "A imagem de cada câmera foi conferida sem faixas nem cintilação, com os refletores na intensidade do evento."),
    ("Energia", "A carga de cada circuito está em até 80% do disjuntor, e o som está em circuito separado do da luz."),
    ("DMX", "Cada aparelho DMX tem modo e endereço anotados, sem sobreposição, e a linha termina com terminador."),
    ("Segurança", "Todo refletor suspenso tem cabo de segurança, os tripés estão travados e com lastro, e nada quente encosta em tecido ou cabo."),
]

# Exemplo 1 do treinamento: o pastor com os olhos escuros na transmissão (plano antes da correção)
EX_PLANO = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares, sala fechada", "data": date(2026, 10, 4),
             "operador": "Operador de luz da escala", "cameras": "Duas câmeras: central e lateral"},
    "ref": (5600, 300, 90, 500, 3),
    "refletores": [  # refletor, posição, alvo, função, distância, altura, intensidade, temperatura, IRC, observação
        ("PAR LED branco frio 1", "Vara sobre a frente do palco", "Púlpito", "Principal", 1.5, 3.2, 100, 6500, 72, ""),
        ("PAR LED branco frio 2", "Vara sobre a frente do palco", "Púlpito", "Principal", 1.5, 3.2, 100, 6500, 72, ""),
        ("Halógeno de 500 W", "Parede lateral, à direita", "Púlpito", "Preenchimento", 4.0, 1.0, 100, 3000, 100, ""),
        ("Barra de LED colorida", "Piso, ao pé da parede do fundo", "Parede do fundo", "Fundo", None, None, 60, None, None, "Azul, para dar clima"),
    ],
    "medicao": [  # alvo, lado claro, lado escuro, fundo
        ("Púlpito, antes", 520, 65, 300),
        ("Púlpito, depois", 800, 400, 200),
    ],
    "ajustes": [  # o quê, o que foi feito
        ("Principal", "Fresnel de LED, IRC 95, 5.600 K, em vara nova 4,5 m à frente do púlpito e 3,2 m acima do rosto: 35°."),
        ("Preenchimento", "Painel de LED, IRC 95, 5.600 K, do outro lado, na mesma vara, com difusor. O halógeno da parede foi desligado."),
        ("Contraluz", "Os dois PAR antigos passaram para trás do púlpito, a 60°: ali o IRC baixo e a cor fria não atrapalham."),
        ("Câmeras", "Balanço de branco refeito com a luz final, nas duas câmeras."),
        ("Resultado", "Olhos visíveis, pele natural, razão de 2:1 e o fundo 2 stops abaixo do rosto."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def cf_fills(ws, rng, pairs):
    cf_equal(ws, rng, pairs)


def referencias(ws, lab1, lab2, val, hint1, hint2, linhas, valores, bg, r0=4):
    """Bloco de valores de referência à direita do cabeçalho. Devolve as referências absolutas."""
    refs = []
    for k, ((rot, dica, fmt), v) in enumerate(zip(linhas, valores)):
        r = r0 + k
        label(ws, f"{lab1}{r}", rot, merge=f"{lab1}{r}:{lab2}{r}")
        inp(ws, f"{val}{r}", v, h="center", bg=bg, fmt=fmt)
        hint(ws, f"{hint1}{r}", dica, merge=f"{hint1}{r}:{hint2}{r}", h="left")
        refs.append(f"${val}${r}")
    dv_number(ws, f"{val}{r0}:{val}{r0 + len(linhas) - 1}")
    return refs


def cabecalho(ws, campos, H, hb, lab_to, inp_from, inp_to):
    for r, rot, k, fmt in campos:
        label(ws, f"B{r}", rot, merge=f"B{r}:{lab_to}{r}")
        inp(ws, f"{inp_from}{r}", H.get(k), merge=f"{inp_from}{r}:{inp_to}{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75


# ================================================================== aba Plano de luz
P_W = {"A": 2, "B": 5, "C": 24, "D": 22, "E": 15, "F": 14, "G": 10, "H": 10, "I": 11, "J": 11, "K": 8,
       "L": 9, "M": 11, "N": 9, "O": 9, "P": 13, "Q": 42, "R": 24, "S": 2}
P1 = 14


def angulo(dist, alt):
    """Ângulo vertical, em graus inteiros, pela tabela de tangentes da coluna oculta U (só funções da lista)."""
    return (f'IF(OR({dist}="",{alt}=""),"",IF({dist}<=0,90,IFERROR(MATCH({alt}/{dist},$U$5:$U$93,1),0)))')


def aba_plano(ws, tab, data=None):
    widths(ws, P_W)
    is_ex = data is not None
    title(ws, "Plano de luz — refletores, ângulos e contraste",
          "Exemplo preenchido, para consulta. Use a aba Plano de luz para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro. As células cinza são calculadas. Medidas em metros, a partir do rosto da pessoa.", "R")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    cabecalho(ws, [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                   (7, "Operador", "operador", None), (8, "Câmeras", "cameras", None)], H, hb, "C", "D", "G")
    dv_date(ws, "D6")
    REF, TOL, IRCMIN, LUXMIN, RAZMAX = referencias(ws, "H", "J", "K", "L", "R", [
        ("Temperatura de referência (K)", "A cor de todas as luzes de rosto: 3.200 K em sala com halógenas, 5.600 K com janela ou telão de LED.", "0"),
        ("Tolerância de cor (K)", "Diferença que ainda conta como a mesma cor. Ponto de partida: 200 a 300 K.", "0"),
        ("IRC mínimo no rosto", "Luz de rosto com IRC abaixo disso recebe aviso. Ponto de partida: 90.", "0"),
        ("Luz mínima no rosto (lux)", "Ponto de partida: 500 lux. Câmera de sensor grande trabalha com menos.", "0"),
        ("Razão máxima (claro ÷ escuro)", "Ponto de partida para fala e culto: até 3, com alvo perto de 2.", "0.0"),
    ], data["ref"] if is_ex else (5600, 300, 90, 500, 3), hb)

    # tabela oculta de tangentes: limite de cada grau inteiro (k - 0,5°), para o ângulo sem ATAN
    put(ws, "U4", "tan(k−0,5°)", f=font(8, c=MUTED), box=False)
    for k in range(1, 90):
        put(ws, f"U{4 + k}", round(math.tan(math.radians(k - 0.5)), 6), f=font(8, c=MUTED), box=False)
    ws.column_dimensions["U"].hidden = True

    band(ws, 10, "Refletores", "R")
    heads = [("B", "nº", INK), ("C", "Refletor", INK), ("D", "Posição", INK), ("E", "Alvo", INK), ("F", "Função", INK),
             ("G", "Distância", AMBER), ("H", "Altura", AMBER), ("I", "Intensidade", AMBER), ("J", "Cor", AMBER), ("K", "IRC", AMBER),
             ("L", "Ângulo", TEAL), ("M", "Situação do ângulo", TEAL), ("N", "Distância mín.", TEAL), ("O", "Distância máx.", TEAL),
             ("P", "Cor e IRC", TEAL), ("Q", "O que fazer", TEAL), ("R", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMNOPQR", ["", "tipo e número", "onde está preso", "quem ou o quê", "lista",
                                           "m, na horizontal", "m, acima do rosto", "% no dimmer", "K", "da ficha",
                                           "graus", "", "m, para a faixa", "m, para a faixa", "luz de rosto", "", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 27
    if is_ex:
        hint(ws, "B13", "Plano antes da correção. A altura é medida a partir do rosto da pessoa (cerca de 1,6 m do chão, em pé). As colunas de L a Q são calculadas.",
             merge="B13:R13", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNOPQ", ["Ex.", "Fresnel LED 1", "Vara de frente", "Púlpito", "Principal", 4.5, 3.2, 80,
                                               5600, 95, 35, "OK", 3.2, 5.5, "OK", "Manter"]):
            ex(ws, f"{col}13", v, h="left" if col in "CDEQ" else "center")
        ex(ws, "R13", "")
        note(ws, "G11", "Distância horizontal, no chão, entre o ponto embaixo do refletor e a pessoa, em metros.")
        note(ws, "H11", "Altura do refletor acima do rosto (e não do chão), em metros.\nRosto de pessoa em pé: cerca de 1,6 m do chão do palco.")
        note(ws, "L11", "Ângulo acima da linha dos olhos, calculado pela altura e pela distância.")
        note(ws, "M11", "Principal: OK de 30° a 45°. Preenchimento: de 25° a 45°. Contraluz: de 45° a 60°.\nPerto da faixa: Atenção. Fundo, plateia e efeito não são avaliados.")
        note(ws, "N11", "Distâncias horizontais que levam este refletor para a faixa OK da função dele, na altura anotada.")
        note(ws, "P11", "Só para principal e preenchimento: cor igual à referência (dentro da tolerância) e IRC acima do mínimo.")
    ws.row_dimensions[13].height = 24

    P2 = P1 + NP - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NP):
        r = P1 + k
        c = data["refletores"][k] if is_ex and k < len(data["refletores"]) else None
        num(ws, f"B{r}", k + 1)
        for col, j in zip("CDEF", range(4)):
            inp(ws, f"{col}{r}", c[j] if c else None, bg=bg, h="center" if col == "F" else "left")
        for col, j, fmt in zip("GHIJK", range(4, 9), (DEC1, DEC1, INT_FMT, "0", INT_FMT)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt=fmt)
        F, G, Hh, J, K, L, M, P = (f"{x}{r}" for x in "FGHJKLMP")
        calc(ws, L, "=" + angulo(G, Hh), fmt=INT_FMT)
        a = L
        sit_p = f'IF({a}<25,"Baixa",IF({a}<30,"Atenção",IF({a}<=45,"OK",IF({a}<=50,"Atenção","Íngreme"))))'
        sit_f = f'IF({a}<20,"Baixa",IF({a}<25,"Atenção",IF({a}<=45,"OK",IF({a}<=50,"Atenção","Íngreme"))))'
        sit_c = f'IF({a}<40,"Baixa",IF({a}<45,"Atenção",IF({a}<=60,"OK",IF({a}<=70,"Atenção","Íngreme"))))'
        calc(ws, M, f'=IF(OR({F}="",{L}=""),"",IF({F}="Principal",{sit_p},IF({F}="Preenchimento",{sit_f},'
                    f'IF({F}="Contraluz",{sit_c},"–"))))', b=False)
        avaliado = f'OR({F}="Principal",{F}="Preenchimento",{F}="Contraluz")'
        calc(ws, f"N{r}", f'=IF({Hh}="","",IF({avaliado},ROUND({Hh}*IF({F}="Contraluz",0.577,1),1),""))', b=False, fmt=DEC1)
        calc(ws, f"O{r}", f'=IF({Hh}="","",IF({avaliado},ROUND({Hh}*IF({F}="Principal",1.732,IF({F}="Preenchimento",2.145,1)),1),""))',
             b=False, fmt=DEC1)
        cor_off = f'IF({J}="",FALSE,ABS({J}-{REF})>{TOL})'
        irc_low = f'IF({K}="",FALSE,{K}<{IRCMIN})'
        calc(ws, P, f'=IF({F}="","",IF(OR({F}="Principal",{F}="Preenchimento"),IF(AND({J}="",{K}=""),"",'
                    f'IF({cor_off},IF({irc_low},"Cor e IRC","Cor diferente"),IF({irc_low},"IRC baixo","OK"))),"–"))', b=False)
        ang_t = (f'IF({M}="Íngreme","Íngreme: afaste o refletor do alvo ou baixe-o",IF({M}="Baixa",IF({F}="Contraluz",'
                 f'"Baixa: suba a contraluz, para não entrar na lente","Baixa: suba o refletor ou aproxime-o do alvo"),'
                 f'IF({M}="Atenção","No limite da faixa: confira na câmera","")))')
        cor_t = (f'IF({P}="Cor diferente","Iguale a cor à referência",IF({P}="IRC baixo","IRC baixo para rosto: use em contraluz ou fundo",'
                 f'IF({P}="Cor e IRC","Cor e IRC fora: passe para contraluz ou fundo","")))')
        calc(ws, f"Q{r}", f'=IF(AND({M}="",{P}=""),"",IF(AND(OR({M}="–",{M}=""),OR({P}="–",{P}="")),"Não avaliado: não é luz de rosto",'
                          f'IF({ang_t}&{cor_t}="","Manter",{ang_t}&IF(AND({ang_t}<>"",{cor_t}<>""),". ","")&{cor_t})))',
             h="left", b=False, sz=9)
        inp(ws, f"R{r}", c[9] if c else None, bg=bg)
        ws.row_dimensions[r].height = 27
    dv_list(ws, f"F{P1}:F{P2}", FUNCOES, "Principal, Preenchimento, Contraluz, Fundo, Plateia ou Efeito.")
    dv_number(ws, f"G{P1}:K{P2}")
    cf_equal(ws, f"M{P1}:M{P2}", [("OK", GREEN), ("Atenção", YELLOW), ("Baixa", RED), ("Íngreme", RED)])
    cf_equal(ws, f"P{P1}:P{P2}", [("OK", GREEN), ("Cor diferente", YELLOW), ("IRC baixo", YELLOW), ("Cor e IRC", RED)])

    # ---------------------------------------------------------- medição no rosto
    m0 = P2 + 2
    band(ws, m0, "Medição no rosto, por alvo (sensor no lugar do rosto, voltado para a câmera)", "R")
    for col, t, bg2 in [("B", "nº", INK), ("C", "Alvo", INK), ("D", "Lado claro", AMBER), ("E", "Lado escuro", BLUE),
                        ("F", "Fundo", INK), ("G", "Razão", TEAL), ("H", "Diferença", TEAL), ("I", "Fundo abaixo", TEAL),
                        ("J", "Situação", TEAL), ("K", "O que fazer", TEAL), ("R", "Observação", INK)]:
        head(ws, f"{col}{m0 + 1}", t, bg=bg2)
    put(ws, f"K{m0 + 1}", "O que fazer", f=font(10, True, c=WHITE), bg=TEAL, h="center", merge=f"K{m0 + 1}:Q{m0 + 1}")
    ws.row_dimensions[m0 + 1].height = 30
    for col, t in zip("BCDEFGHIJ", ["", "", "lux, tudo aceso", "lux, sem a principal", "lux", "claro ÷ escuro", "stops", "stops", ""]):
        hint(ws, f"{col}{m0 + 2}", t)
    hint(ws, f"K{m0 + 2}", "", merge=f"K{m0 + 2}:Q{m0 + 2}")
    hint(ws, f"R{m0 + 2}", "")
    ws.row_dimensions[m0 + 2].height = 27
    if not is_ex:
        note(ws, f"D{m0 + 1}", "Leitura com todas as luzes do alvo acesas: principal mais preenchimento.")
        note(ws, f"E{m0 + 1}", "Leitura com a principal apagada ou tampada: só o preenchimento.")
        note(ws, f"F{m0 + 1}", "Leitura no fundo, atrás da pessoa, com o sensor voltado para a câmera.")
        note(ws, f"I{m0 + 1}", "Quantos stops o fundo está abaixo do rosto. Ponto de partida: de 1 a 2.")
    M1 = m0 + 3
    M2 = M1 + NM - 1
    for k in range(NM):
        r = M1 + k
        c = data["medicao"][k] if is_ex and k < len(data["medicao"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        for col, j in zip("DEF", (1, 2, 3)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt=INT_FMT)
        D, E, Fm, Gm, Hm, Im, Jm = (f"{x}{r}" for x in "DEFGHIJ")
        calc(ws, Gm, f'=IF(OR({D}="",{E}=""),"",IF(OR({D}<=0,{E}<=0),"",ROUND({D}/{E},1)))', fmt=DEC1)
        calc(ws, Hm, f'=IF({Gm}="","",ROUND(LOG10({D}/{E})/LOG10(2),1))', b=False, fmt=DEC1)
        calc(ws, Im, f'=IF(OR({D}="",{Fm}=""),"",IF(OR({D}<=0,{Fm}<=0),"",ROUND(LOG10({D}/{Fm})/LOG10(2),1)))', b=False, fmt=DEC1)
        calc(ws, Jm, f'=IF({D}="","",IF({D}<{LUXMIN},"Luz baixa",IF({Gm}="",IF({Im}="","OK",IF({Im}<1,"Fundo claro","OK")),'
                     f'IF({Gm}>{RAZMAX},"Contraste alto",IF({Gm}<1.5,"Chapado",IF({Im}="","OK",IF({Im}<1,"Fundo claro","OK")))))))', b=False)
        calc(ws, f"K{r}", f'=IF({Jm}="","",IF({Jm}="Luz baixa","Suba a principal ou aproxime-a até a luz mínima de referência",'
                          f'IF({Jm}="Contraste alto","Suba o preenchimento até a razão chegar perto de 2:1",'
                          f'IF({Jm}="Chapado","Baixe o preenchimento até a razão chegar perto de 2:1",'
                          f'IF({Jm}="Fundo claro","Baixe a luz do fundo ou o brilho do telão: de 1 a 2 stops abaixo do rosto","Manter")))))',
             h="left", b=False, sz=9, merge=f"K{r}:Q{r}")
        inp(ws, f"R{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"D{M1}:F{M2}")
    cf_equal(ws, f"J{M1}:J{M2}", [("OK", GREEN), ("Luz baixa", YELLOW), ("Contraste alto", YELLOW), ("Chapado", YELLOW), ("Fundo claro", YELLOW)])

    # ---------------------------------------------------------- resumo
    s = M2 + 2
    band(ws, s, "Resumo automático", "R")
    cnt = lambda rng, t: f'COUNTIF({rng},"{t}")'
    Mr, Pr, Fr, Jr = f"M{P1}:M{P2}", f"P{P1}:P{P2}", f"F{P1}:F{P2}", f"J{M1}:J{M2}"
    linhas = [
        ("Refletores registrados", f"=COUNTA(C{P1}:C{P2})"),
        ("Luzes de rosto (principal e preenchimento)", f'={cnt(Fr, "Principal")}+{cnt(Fr, "Preenchimento")}'),
        ("Contraluzes", f'={cnt(Fr, "Contraluz")}'),
        ("Ângulos na faixa (OK)", f'={cnt(Mr, "OK")}'),
        ("Ângulos a conferir ou corrigir", f'={cnt(Mr, "Atenção")}+{cnt(Mr, "Baixa")}+{cnt(Mr, "Íngreme")}'),
        ("Luzes de rosto com cor ou IRC fora", f'={cnt(Pr, "Cor diferente")}+{cnt(Pr, "IRC baixo")}+{cnt(Pr, "Cor e IRC")}'),
        ("Alvos medidos", f"=COUNT(D{M1}:D{M2})"),
        ("Alvos fora da referência", f'=COUNT(D{M1}:D{M2})-{cnt(Jr, "OK")}'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "F", "G", val_merge=f"G{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    v = lambda i: f"G{s + i}"
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:F{a}", h="right")
    calc(ws, f"G{a}", f'=IF(COUNT({REF},{TOL},{IRCMIN},{LUXMIN},{RAZMAX})<5,"Informe os cinco valores de referência",'
                      f'IF({v(1)}=0,"Liste os refletores",IF({cnt(Fr, "Principal")}=0,"Não há luz principal no plano",'
                      f'IF({v(5)}+{v(6)}>0,"Há refletor de rosto a corrigir: veja a coluna O que fazer",'
                      f'IF({v(7)}=0,"Meça a luz no rosto de cada alvo",IF({v(8)}>0,"Há alvo fora da referência na medição",'
                      f'IF({v(3)}=0,"OK, mas não há contraluz no plano","OK")))))))',
         b=False, merge=f"G{a}:Q{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"G{a}:Q{a}")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "R")
        r = end + 3
        put(ws, f"B{r}", "Luz", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:R{r}")
        ws.row_dimensions[r].height = 21.75
        for oq, feito in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", oq, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:R{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{P1}"
    setup(ws, tab, f"B1:R{end}", fit_height=True)


# ================================================================== aba Carga elétrica
C_W = {"A": 2, "B": 5, "C": 30, "D": 24, "E": 11, "F": 12, "G": 12, "H": 12, "I": 10, "J": 13, "K": 12, "L": 14, "M": 44, "N": 2}


def aba_carga(ws, tab):
    widths(ws, C_W)
    title(ws, "Carga elétrica — quanto cabe em cada circuito",
          "Cadastre os circuitos (tensão e disjuntor) e depois os aparelhos. Corrente = potência ÷ tensão.", "M")
    cabecalho(ws, [(4, "Evento ou local", "", None), (5, "Data", "", DATE), (6, "Responsável", "", None)], {}, INPUT, "C", "D", "G")
    dv_date(ws, "D5")
    (LIM,) = referencias(ws, "H", "J", "K", "L", "M", [
        ("Limite de uso contínuo", "Fração do disjuntor para carga ligada por horas. Prática comum: 80%.", "0%"),
    ], (0.8,), INPUT)
    hint(ws, "H5", "Quem define circuitos e mexe no quadro é um eletricista habilitado. Na dúvida sobre uma tomada, pergunte ao responsável pelo prédio.",
         merge="H5:M6", h="left")

    # ---------------------------------------------------------- circuitos
    band(ws, 8, "Circuitos", "M")
    for col, t, bg in [("B", "#", INK), ("C", "Circuito", INK), ("D", "Onde fica", INK), ("E", "Tensão", AMBER), ("F", "Disjuntor", AMBER),
                       ("G", "Carga", TEAL), ("H", "Corrente", TEAL), ("I", "Uso", TEAL), ("J", "Cabe até", TEAL), ("K", "Sobra", TEAL),
                       ("L", "Situação", TEAL), ("M", "O que fazer", TEAL)]:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKLM", ["", "nome curto (A, B, palco 1)", "tomada, quadro, disjuntor nº", "V (127 ou 220)", "A",
                                       "W, soma dos aparelhos", "A", "% do disjuntor", "W, no limite de uso", "W", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 27
    for col, v in zip("BCDEFGHIJKLM", ["Ex.", "A", "Tomada do palco, lado esquerdo", 127, 20, 2000, 15.7, "79%", 2032, 32, "OK", "Manter"]):
        ex(ws, f"{col}11", v, h="left" if col in "CDM" else "center")
    ws.row_dimensions[11].height = 21.75
    note(ws, "E9", "Tensão medida ou informada da tomada: 127 V ou 220 V.\nNa mesma potência, a corrente em 127 V é quase o dobro da corrente em 220 V.")
    note(ws, "F9", "Valor em ampères escrito no disjuntor do quadro que protege esta tomada.")
    note(ws, "J9", "Tensão × disjuntor × limite de uso. É a potência que o circuito aguenta por horas, com folga.")
    C1, C2 = 12, 12 + NC - 1
    A1, A2 = C2 + 5 + 1, C2 + 5 + NA
    for k in range(NC):
        r = C1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center")
        inp(ws, f"D{r}")
        inp(ws, f"E{r}", h="center", fmt=INT_FMT)
        inp(ws, f"F{r}", h="center", fmt=INT_FMT)
        C, E, F, G, Hc, I, J, K, L = (f"{x}{r}" for x in "CEFGHIJKL")
        calc(ws, G, f'=IF({C}="","",SUMIF($D${A1}:$D${A2},{C},$G${A1}:$G${A2}))', fmt=INT_FMT)
        calc(ws, Hc, f'=IF(OR({G}="",{E}=""),"",IF({E}<=0,"",ROUND({G}/{E},1)))', fmt=DEC1)
        calc(ws, I, f'=IF(OR({Hc}="",{F}=""),"",IF({F}<=0,"",{Hc}/{F}))', fmt="0%")
        calc(ws, J, f'=IF(OR({E}="",{F}=""),"",ROUND({E}*{F}*{LIM},0))', b=False, fmt=INT_FMT)
        calc(ws, K, f'=IF(OR({J}="",{G}=""),"",{J}-{G})', b=False, fmt=INT_FMT)
        calc(ws, L, f'=IF({C}="","",IF(OR({E}="",{F}=""),"Falta dado",IF({G}=0,"Sem carga",IF({I}>1,"Sobrecarga",IF({I}>{LIM},"Atenção","OK")))))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="Sobrecarga","Passou do disjuntor: tire aparelhos deste circuito",'
                          f'IF({L}="Atenção","Acima do limite de uso contínuo: passe aparelhos para outro circuito",'
                          f'IF({L}="Falta dado","Informe a tensão e o disjuntor",IF({L}="Sem carga","Livre","Manter")))))',
             h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"E{C1}:F{C2}")
    cf_equal(ws, f"L{C1}:L{C2}", [("OK", GREEN), ("Atenção", YELLOW), ("Falta dado", YELLOW), ("Sobrecarga", RED)])

    # ---------------------------------------------------------- aparelhos
    a0 = C2 + 2
    band(ws, a0, "Aparelhos", "M")
    for col, t, bg in [("B", "#", INK), ("C", "Aparelho", INK), ("D", "Circuito", AMBER), ("E", "Quantidade", AMBER),
                       ("F", "Potência", AMBER), ("G", "Potência total", TEAL), ("H", "Corrente", TEAL), ("I", "Situação", TEAL)]:
        head(ws, f"{col}{a0 + 1}", t, bg=bg)
    put(ws, f"J{a0 + 1}", "Observação", f=font(10, True, c=WHITE), bg=INK, h="center", merge=f"J{a0 + 1}:M{a0 + 1}")
    ws.row_dimensions[a0 + 1].height = 30
    for col, t in zip("BCDEFGHI", ["", "tipo e modelo", "nome da lista", "unidades", "W, por aparelho", "W", "A", ""]):
        hint(ws, f"{col}{a0 + 2}", t)
    hint(ws, f"J{a0 + 2}", "Use a potência ou a corrente máxima da etiqueta, e não a “equivalência” da propaganda.", merge=f"J{a0 + 2}:M{a0 + 2}", h="left")
    ws.row_dimensions[a0 + 2].height = 27
    for col, v in zip("BCDEFGHI", ["Ex.", "Fresnel halógeno de 1.000 W", "A", 2, 1000, 2000, 15.7, "OK"]):
        ex(ws, f"{col}{a0 + 3}", v, h="left" if col == "C" else "center")
    ex(ws, f"J{a0 + 3}", "")
    ws.merge_cells(f"J{a0 + 3}:M{a0 + 3}")
    ws.row_dimensions[a0 + 3].height = 21.75
    assert a0 + 4 == A1
    for k in range(NA):
        r = A1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}")
        inp(ws, f"D{r}", h="center")
        inp(ws, f"E{r}", h="center", fmt=INT_FMT)
        inp(ws, f"F{r}", h="center", fmt=INT_FMT)
        C, D, E, F, G = (f"{x}{r}" for x in "CDEFG")
        calc(ws, G, f'=IF(OR({E}="",{F}=""),"",{E}*{F})', fmt=INT_FMT)
        calc(ws, f"H{r}", f'=IF(OR({G}="",{D}=""),"",IFERROR(ROUND({G}/VLOOKUP({D},$C${C1}:$E${C2},3,FALSE),1),""))', b=False, fmt=DEC1)
        calc(ws, f"I{r}", f'=IF(AND({C}="",{D}=""),"",IF({D}="","Sem circuito",IF(COUNTIF($C${C1}:$C${C2},{D})=0,"Circuito não cadastrado",'
                          f'IF(OR({E}="",{F}=""),"Falta potência","OK"))))', b=False, sz=9)
        inp(ws, f"J{r}", merge=f"J{r}:M{r}")
        ws.row_dimensions[r].height = 19.5
    dv = DataValidation(type="list", formula1=f"$C${C1}:$C${C2}", allow_blank=True)
    dv.promptTitle, dv.prompt = "Circuito", "Escolha um dos circuitos cadastrados no alto da aba."
    dv.errorTitle, dv.error = "Circuito inválido", "Cadastre o circuito no alto da aba antes de usá-lo aqui."
    dv.showErrorMessage = dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(f"D{A1}:D{A2}")
    dv_number(ws, f"E{A1}:F{A2}")
    cf_status(ws, f"I{A1}:I{A2}", f"I{A1}")

    s = A2 + 2
    band(ws, s, "Resumo automático", "M")
    Lr = f"L{C1}:L{C2}"
    linhas = [
        ("Carga total ligada (W)", f"=SUM(G{A1}:G{A2})", INT_FMT),
        ("Corrente total em 127 V, se fosse um circuito só (A)", f"=ROUND(SUM(G{A1}:G{A2})/127,1)", DEC1),
        ("Circuitos em sobrecarga", f'=COUNTIF({Lr},"Sobrecarga")', None),
        ("Circuitos acima do limite de uso", f'=COUNTIF({Lr},"Atenção")', None),
        ("Aparelhos sem circuito válido", f'=COUNTIF(I{A1}:I{A2},"Sem circuito")+COUNTIF(I{A1}:I{A2},"Circuito não cadastrado")', None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "F", "G", fmt=fmt, val_merge=f"G{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:F{a}", h="right")
    calc(ws, f"G{a}", f'=IF(COUNTA(C{C1}:C{C2})=0,"Cadastre os circuitos",IF(COUNT(G{A1}:G{A2})=0,"Liste os aparelhos",'
                      f'IF(G{s + 3}>0,"Há circuito em sobrecarga: redistribua antes de ligar",IF(G{s + 5}>0,"Há aparelho sem circuito válido",'
                      f'IF(G{s + 4}>0,"Há circuito acima do limite de uso contínuo","OK")))))',
         b=False, merge=f"G{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"G{a}:M{a}")
    hint(ws, f"B{a + 1}", "O limite de 80% é prática comum para carga contínua, e não norma. Tomada, extensão, régua e canal de dimmer têm cada um o seu limite, que pode ser menor que o do disjuntor.",
         merge=f"B{a + 1}:M{a + 1}", h="left")
    ws.row_dimensions[a + 1].height = 27
    ws.freeze_panes = "C12"
    setup(ws, tab, f"B1:M{a + 1}", fit_height=True)


# ================================================================== aba DMX
D_W = {"A": 2, "B": 5, "C": 28, "D": 22, "E": 10, "F": 18, "G": 10, "H": 11, "I": 11, "J": 11, "K": 16, "L": 36, "M": 2}


def aba_dmx(ws, tab):
    widths(ws, D_W)
    title(ws, "DMX — endereçamento dos aparelhos",
          "Para cada aparelho: universo, modo, número de canais e endereço inicial. A planilha confere a sobreposição.", "L")
    cabecalho(ws, [(4, "Evento ou local", "", None), (5, "Mesa ou interface", "", None), (6, "Data", "", DATE)], {}, INPUT, "C", "D", "F")
    dv_date(ws, "D6")
    hint(ws, "G4", "Um universo tem 512 canais. Cada aparelho ocupa do endereço inicial até endereço + canais − 1. "
                   "O número de canais de cada modo está no manual do aparelho.", merge="G4:L6", h="left")
    band(ws, 8, "Aparelhos na linha DMX", "L")
    for col, t, bg in [("B", "#", INK), ("C", "Aparelho", INK), ("D", "Posição", INK), ("E", "Universo", AMBER), ("F", "Modo", AMBER),
                       ("G", "Canais", AMBER), ("H", "Endereço inicial", AMBER), ("I", "Último canal", TEAL), ("J", "Próximo livre", TEAL),
                       ("K", "Situação", TEAL), ("L", "Observação", INK)]:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKL", ["", "tipo e número", "onde está", "1, 2…", "como no menu", "nº do modo", "1 a 512", "", "", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    for col, v in zip("BCDEFGHIJK", ["Ex.", "PAR LED 1", "Vara de contraluz", 1, "7 canais (7CH)", 7, 1, 7, 8, "OK"]):
        ex(ws, f"{col}11", v, h="left" if col in "CDF" else "center")
    ex(ws, "L11", "")
    ws.row_dimensions[11].height = 21.75
    note(ws, "H9", "Primeiro canal que o aparelho escuta, ajustado no menu ou nas chaves DIP.")
    note(ws, "K9", "Conflito: os canais deste aparelho se sobrepõem aos de outro no mesmo universo.\n"
                   "Se for de propósito (dois aparelhos iguais, no mesmo modo, respondendo juntos), anote na observação.")
    D1, D2 = 12, 12 + ND - 1
    for k in range(ND):
        r = D1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}")
        inp(ws, f"D{r}")
        inp(ws, f"E{r}", h="center", fmt=INT_FMT)
        inp(ws, f"F{r}", h="center")
        inp(ws, f"G{r}", h="center", fmt=INT_FMT)
        inp(ws, f"H{r}", h="center", fmt=INT_FMT)
        C, E, G, Hd, I = (f"{x}{r}" for x in "CEGHI")
        calc(ws, I, f'=IF(OR({G}="",{Hd}=""),"",{Hd}+{G}-1)', fmt=INT_FMT)
        calc(ws, f"J{r}", f'=IF({I}="","",IF({I}>=512,"–",{I}+1))', b=False, fmt=INT_FMT)
        calc(ws, f"K{r}", f'=IF(AND({C}="",{Hd}=""),"",IF(OR({E}="",{G}="",{Hd}=""),"Falta dado",IF(OR({Hd}<1,{G}<1,{I}>512),"Fora do universo",'
                          f'IF(COUNTIFS($E${D1}:$E${D2},{E},$H${D1}:$H${D2},"<="&{I},$I${D1}:$I${D2},">="&{Hd})>1,"Conflito","OK"))))')
        inp(ws, f"L{r}")
        ws.row_dimensions[r].height = 19.5
    dv_number(ws, f"E{D1}:E{D2}")
    dv_number(ws, f"G{D1}:H{D2}")
    cf_equal(ws, f"K{D1}:K{D2}", [("OK", GREEN), ("Falta dado", YELLOW), ("Conflito", RED), ("Fora do universo", RED)])

    s = D2 + 2
    band(ws, s, "Resumo automático", "L")
    Kr = f"K{D1}:K{D2}"
    linhas = [
        ("Aparelhos registrados", f"=COUNTA(C{D1}:C{D2})"),
        ("Aparelhos no universo 1", f"=COUNTIF(E{D1}:E{D2},1)"),
        ("Canais ocupados no universo 1", f"=SUMIF(E{D1}:E{D2},1,G{D1}:G{D2})"),
        ("Maior canal ocupado", f'=IF(COUNT(I{D1}:I{D2})=0,"",MAX(I{D1}:I{D2}))'),
        ("Aparelhos em conflito", f'=COUNTIF({Kr},"Conflito")'),
        ("Aparelhos fora do universo", f'=COUNTIF({Kr},"Fora do universo")'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "F", "G", val_merge=f"G{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:F{a}", h="right")
    calc(ws, f"G{a}", f'=IF(G{s + 1}=0,"Liste os aparelhos",IF(G{s + 5}+G{s + 6}>0,"Há endereço em conflito ou fora do universo",'
                      f'IF(COUNTIF({Kr},"Falta dado")>0,"Há aparelho sem universo, canais ou endereço",'
                      f'IF(G{s + 2}>32,"Mais de 32 aparelhos no universo 1: use divisor (splitter) e confira cada linha","OK"))))',
         b=False, merge=f"G{a}:L{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"G{a}:L{a}")

    # ---------------------------------------------------------- chaves DIP
    d0 = a + 2
    band(ws, d0, "Endereço em chaves DIP (para aparelhos sem visor)", "L")
    label(ws, f"B{d0 + 1}", "Endereço desejado", merge=f"B{d0 + 1}:C{d0 + 1}")
    inp(ws, f"D{d0 + 1}", h="center", fmt=INT_FMT)
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="511", allow_blank=True)
    dv.errorTitle, dv.error = "Endereço inválido", "Digite um número inteiro de 1 a 511."
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(f"D{d0 + 1}")
    hint(ws, f"E{d0 + 1}", "A conta vai da chave de maior valor para a menor: liga a chave quando o resto ainda comporta o valor dela.",
         merge=f"E{d0 + 1}:L{d0 + 1}", h="left")
    ws.row_dimensions[d0 + 1].height = 21.75
    for col, t in zip("CDEF", ["Chave", "Valor", "Posição", "Resto"]):
        put(ws, f"{col}{d0 + 2}", t, f=font(10, True), bg=GRAY, h="center")
    ADDR = f"$D${d0 + 1}"
    prev = ADDR
    first = d0 + 3
    for i, chave in enumerate(range(9, 0, -1)):
        r = first + i
        val = 2 ** (chave - 1)
        num(ws, f"C{r}", chave)
        put(ws, f"D{r}", val, bg=GRAY, h="center")
        calc(ws, f"E{r}", f'=IF({ADDR}="","",IF({prev}>={val},"Ligada","Desligada"))')
        calc(ws, f"F{r}", f'=IF({ADDR}="","",{prev}-IF(E{r}="Ligada",{val},0))', b=False, fmt=INT_FMT)
        prev = f"F{r}"
        ws.row_dimensions[r].height = 19.5
    last = first + 8
    cf_equal(ws, f"E{first}:E{last}", [("Ligada", GREEN)])
    summary(ws, last + 1, "Conferência: soma das chaves ligadas", f'=IF({ADDR}="","",SUMIF(E{first}:E{last},"Ligada",D{first}:D{last}))', "D", "E")
    hint(ws, f"G{first}", "Exemplo: endereço 37 = chaves 6 (32), 3 (4) e 1 (1) ligadas. Em alguns aparelhos a chave 10 escolhe outra função: confira o manual.",
         merge=f"G{first}:L{first + 2}", h="left")
    ws.freeze_panes = "C12"
    setup(ws, tab, f"B1:L{last + 1}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Iluminação — modelo de plano de luz"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Iluminação — como usar esta planilha",
           "Plano de luz com ângulos e contraste, carga elétrica por circuito e endereçamento DMX, para montar a luz do evento.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Função do refletor e circuito do aparelho aceitam só as opções da lista."),
        ("Números", "Digite só o número: 3,2 para metros, 5600 para kelvins, 800 para lux. Sem unidade."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Carga elétrica: cadastre os circuitos do local e liste os aparelhos. Corrija a distribuição antes de ligar."),
        ("2", "Aba Plano de luz: confira os cinco valores de referência no alto (cor, tolerância, IRC, luz mínima e razão máxima)."),
        ("3", "Aba Plano de luz: liste os refletores, com a função, a distância horizontal e a altura acima do rosto."),
        ("4", "Veja o ângulo e a situação de cada um. Mova os refletores que estão fora da faixa e troque os números."),
        ("5", "Com a luz apontada, meça o lado claro, o lado escuro e o fundo de cada alvo e anote na medição no rosto."),
        ("6", "Aba DMX: anote o modo, os canais e o endereço de cada aparelho e confira que não há conflito."),
        ("7", "Aba Checklist: confira os doze itens antes do ajuste final das câmeras."),
    ]),
    ("Abas da planilha", [
        ("Plano de luz", "Até 20 refletores, com ângulo vertical, situação, faixa de distância, cor e IRC; medição no rosto de até 6 alvos."),
        ("Carga elétrica", "Até 8 circuitos e 24 aparelhos: carga, corrente, uso do disjuntor e alerta de sobrecarga."),
        ("DMX", "Até 24 aparelhos: último canal, próximo endereço livre, conflito; conversão de endereço em chaves DIP."),
        ("Referência", "Três pontos, temperatura de cor, IRC, luz no rosto, razão de contraste, ângulos, facho, corrente e DMX."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Plano", "O plano de luz do exemplo 1 do treinamento, antes da correção, com a medição antes e depois."),
    ]),
    ("Regras e premissas do modelo", [
        ("Ângulo", "Calculado pela altura acima do rosto e pela distância horizontal até a pessoa, arredondado para o grau inteiro."),
        ("Faixas de ângulo", "Principal: OK de 30° a 45°. Preenchimento: de 25° a 45°. Contraluz: de 45° a 60°. Perto da faixa: Atenção. São pontos de partida: o que decide é a imagem no monitor."),
        ("Cor e IRC", "Avaliados só nas luzes de rosto (principal e preenchimento). Contraluz, fundo, plateia e efeito podem ter outra cor."),
        ("Razão de contraste", "Lado claro (tudo aceso) dividido pelo lado escuro (sem a principal). Abaixo de 1,5: chapado. Acima do máximo: contraste alto."),
        ("Fundo", "Fundo menos de 1 stop abaixo do rosto recebe aviso de fundo claro."),
        ("Carga elétrica", "Corrente = potência ÷ tensão. Atenção acima do limite de uso (80%, como prática); sobrecarga acima de 100% do disjuntor."),
        ("Exemplos", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_plano(wb.create_sheet("Plano de luz"), AMBER)
aba_carga(wb.create_sheet("Carga elétrica"), HUE["red"][0])
aba_dmx(wb.create_sheet("DMX"), PURPLE)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 32, "C": 26, "D": 26, "E": 58, "F": 2})
title(ws, "Referência — valores típicos", "Valores de referência, para servir de ponto de partida. Confira o manual e a ficha técnica do seu equipamento.", "E")


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


r = tabela(4, "Os três pontos de luz", ["Luz", "De onde vem", "Ângulo vertical", "Força, como ponto de partida"], [
    ("Principal (key)", "de frente, 30° a 45° para um lado", "30° a 45° acima dos olhos", "A referência. Define a exposição e dá volume ao rosto."),
    ("Preenchimento (fill)", "de frente, do outro lado", "igual ou um pouco abaixo", "Suave. Ajustado pela razão de contraste, perto de 2:1."),
    ("Contraluz (back light)", "de trás, oposta à principal", "45° a 60°", "Da metade ao mesmo valor da principal. Cabelo e roupa escuros pedem mais."),
])
r = tabela(r, "Temperatura de cor", ["Fonte", "Temperatura típica", "Aparência", "Observação"], [
    ("Lâmpada doméstica “branco quente”", "2.700 K", "alaranjada", "Comum em luminárias de sala."),
    ("Halógena de palco (“tungstênio”)", "3.200 K", "alaranjada", "Referência de sala fechada. Fica mais alaranjada com o dimmer baixo."),
    ("Branco neutro", "4.000 K", "neutra", "Fluorescentes e LEDs de iluminação geral."),
    ("Luz do dia, LED “daylight”", "5.600 K", "branca a azulada", "Referência com janela ou telão de LED."),
    ("Céu nublado, monitores e telões", "6.500 K", "azulada", "Painéis de LED costumam vir daí para cima."),
])
r = tabela(r, "IRC (CRI) do refletor", ["IRC", "Pele na câmera", "Onde usar", "Observação"], [
    ("90 ou mais", "natural", "luz de rosto", "Se a ficha trouxer R9 e TLCI, prefira os mais altos."),
    ("80 a 89", "um pouco sem vida", "rosto, se não houver outro", "Pode pedir correção na câmera."),
    ("abaixo de 80", "acinzentada ou esverdeada", "contraluz, fundo, efeito", "Branco feito só de LED RGB costuma ficar aqui."),
])
r = tabela(r, "Luz no rosto (lux)", ["Situação", "Luz típica", "Câmera", "Observação"], [
    ("Sala de estar à noite", "50 a 150 lux", "ganho alto", "Imagem com ruído."),
    ("Sala de aula, nave com luz geral", "200 a 500 lux", "íris aberta, algum ganho", "Foco crítico."),
    ("Rosto no palco, para vídeo", "500 a 1.000 lux", "imagem limpa", "Ponto de partida, e não norma."),
    ("Estúdio de televisão", "1.000 a 2.000 lux", "íris mais fechada", "Foco confortável em zoom longo."),
])
r = tabela(r, "Razão de contraste", ["Razão (claro ÷ escuro)", "Diferença", "Aparência", "Uso"], [
    ("1:1", "0 stop", "chapado, sem volume", "Evite no rosto principal."),
    ("2:1", "1 stop", "volume natural", "Ponto de partida para fala, culto e transmissão."),
    ("3:1", "1,6 stop", "mais volume", "Limite para várias câmeras."),
    ("4:1", "2 stops", "dramático", "Limite prático ao vivo."),
    ("8:1", "3 stops", "lado escuro sem detalhe", "Evite em evento ao vivo."),
])
ang = []
for g in (20, 25, 30, 35, 40, 45, 50, 55, 60):
    t = math.tan(math.radians(g))
    uso = ("baixa para frente" if g < 25 else "preenchimento" if g < 30 else "principal e preenchimento" if g <= 45
           else "íngreme para frente; contraluz baixa" if g < 50 else "íngreme para frente; contraluz" if g < 55 else "contraluz")
    ang.append((f"{g}°", f"{t:.2f}".replace(".", ","), f"{3 / t:.1f} m".replace(".", ","), uso))
r = tabela(r, "Ângulo pela razão altura ÷ distância", ["Ângulo", "Altura ÷ distância", "Distância para 3 m de altura", "Uso típico"], ang, alt=19.5)
facho = []
for g in (10, 15, 19, 26, 36, 45, 50, 60):
    m = 2 * math.tan(math.radians(g / 2))
    facho.append((f"{g}°", f"{m:.2f}".replace(".", ","), f"{6 * m:.1f} m".replace(".", ","), f"{10 * m:.1f} m".replace(".", ",")))
r = tabela(r, "Tamanho do facho: diâmetro ≈ distância × multiplicador", ["Ângulo do facho", "Multiplicador", "Diâmetro a 6 m", "Diâmetro a 10 m"],
           facho, alt=19.5)
r = tabela(r, "Corrente por potência (corrente = potência ÷ tensão)", ["Potência", "Em 127 V", "Em 220 V", "Observação"], [
    ("500 W", "3,9 A", "2,3 A", "PAR halógeno de 500 W."),
    ("1.000 W", "7,9 A", "4,5 A", "Fresnel halógeno de 1.000 W."),
    ("2.000 W", "15,7 A", "9,1 A", "Perto do limite de uso de um circuito de 20 A em 127 V."),
    ("3.000 W", "23,6 A", "13,6 A", "Passa de um disjuntor de 20 A em 127 V."),
], alt=19.5)
r = tabela(r, "Limite de uso contínuo (80% do disjuntor)", ["Disjuntor", "Em 127 V", "Em 220 V", "Observação"], [
    ("10 A", "1.016 W", "1.760 W", "O máximo absoluto é tensão × disjuntor; o limite de uso fica em 80% dele."),
    ("16 A", "1.626 W", "2.816 W", "Tomada, extensão e régua podem ter limite menor que o disjuntor."),
    ("20 A", "2.032 W", "3.520 W", "Exemplo 2 do treinamento: 5.000 W em 127 V pedem três circuitos."),
    ("25 A", "2.540 W", "4.400 W", "Confira com o responsável pelo prédio."),
], alt=19.5)
r = tabela(r, "DMX: regras da linha", ["Regra", "Valor de referência", "Por quê", "Observação"], [
    ("Universo", "512 canais", "valores de 0 a 255", "Um universo por saída da mesa ou da interface."),
    ("Ligação", "em cascata", "da saída de um para a entrada do outro", "Sem cabo em Y. Para dividir, use divisor (splitter)."),
    ("Terminador", "120 Ω no último", "evita reflexão do sinal", "Linha sem terminador pode falhar só no evento."),
    ("Aparelhos por linha", "até 32", "limite do circuito de cada linha", "Acima disso, use divisor."),
    ("Cabo", "cabo próprio de DMX", "impedância para dados", "Cabo de microfone falha em linha longa."),
    ("Perda de sinal", "manter o último valor", "luz de rosto não apaga", "Configure no menu de cada aparelho."),
])
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com o modelo do refletor, a câmera e o local. Confira a ficha técnica e o manual.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=False, landscape=False)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist da iluminação",
          "Confira depois de apontar a luz e antes do ajuste final das câmeras. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["teal"][0])
aba_plano(wb.create_sheet("Exemplo - Plano"), MUTED, EX_PLANO)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
