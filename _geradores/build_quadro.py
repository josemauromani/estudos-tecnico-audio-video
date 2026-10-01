# -*- coding: utf-8 -*-
"""Gera Enquadramento-e-Movimento-modelo.xlsx: posições de câmera, lista de planos, autoavaliação do quadro,
referência e checklist."""
import math
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
PLANOS = ["PG", "PC", "PA", "PM", "PP", "Close", "Detalhe"]
MOVS = ["Fixo", "Pan", "Tilt", "Zoom", "Travelling", "Acompanha"]
SUPORTES = ["Tripé fluido", "Tripé foto", "Monopé", "Mão", "Estabilizador", "PTZ"]
LADOS = ["Fundo da sala", "Lateral esquerda", "Lateral direita", "Frente da plateia", "Mezanino", "Fundo do palco"]
DEG = '+0"°";-0"°";0"°"'

CHECK = [
    ("Preparação", "Li o roteiro e a lista de planos, e sei o plano de base da minha câmera em cada momento."),
    ("Posição", "A posição tem visão livre do palco com o público em pé, e está do lado certo do eixo."),
    ("Posição", "O tripé está firme, fora da passagem, e os cabos estão presos ao chão."),
    ("Tripé", "A cabeça está nivelada: a bolha está no centro, e o horizonte fica reto na panorâmica de ponta a ponta."),
    ("Tripé", "A câmera está equilibrada, com contrapeso e fricção ajustados: fica parada onde eu a deixo."),
    ("Imagem", "A grade de terços está ligada no visor, e exposição, foco e cor foram conferidos."),
    ("Quadro", "Em cada plano de pessoa, os olhos ficam no terço superior e o teto está na medida."),
    ("Quadro", "Há espaço livre na direção do olhar ou do movimento."),
    ("Quadro", "As bordas e o fundo estão limpos, e o corte do corpo fica fora das articulações."),
    ("Movimento", "Ensaiei os movimentos: sei o quadro inicial e o final, e seguro dois ou três segundos em cada um."),
    ("Movimento", "Testei e anotei o plano mais fechado que mantenho firme desta posição."),
    ("Equipe", "O intercomunicador e a luz de tally foram testados, e os comandos foram combinados com o diretor."),
]

# Exemplo 1 do treinamento: a primeira lista de planos de um culto com três câmeras
EX_PLANOS = {
    "head": {"evento": "Culto de domingo, 19 h, com transmissão", "data": date(2026, 10, 4), "diretor": "Diretor de imagem da escala"},
    "cams": [("Central, ao fundo, sobre praticável", "Tripé fluido", "Operador 1"),
             ("Lateral direita, a uns 35° do centro", "Tripé fluido", "Operador 2"),
             ("Móvel, perto do palco", "Mão", "Operador 3")],
    "momentos": ["Abertura", "Louvor", "Leitura", "Pregação", "Encerramento"],
    "planos": [  # momento, câmera, assunto, plano, movimento
        ("Abertura", 1, "Igreja e palco", "PG", "Fixo"),
        ("Abertura", 2, "Apresentador", "PM", "Fixo"),
        ("Abertura", 3, "Congregação chegando", "PC", "Travelling"),
        ("Louvor", 1, "Palco e congregação", "PG", "Fixo"),
        ("Louvor", 2, "Cantora", "PP", "Fixo"),
        ("Louvor", 3, "Banda", "PC", "Acompanha"),
        ("Leitura", 1, "Leitor", "PM", "Fixo"),
        ("Leitura", 2, "Leitor", "PM", "Fixo"),
        ("Leitura", 3, "Bíblia nas mãos", "Detalhe", "Fixo"),
        ("Pregação", 1, "Pastor", "PM", "Fixo"),
        ("Pregação", 2, "Pastor", "PM", "Acompanha"),
        ("Pregação", 3, "Pastor", "PP", "Acompanha"),
        ("Encerramento", 1, "Igreja e palco", "PG", "Fixo"),
        ("Encerramento", 2, "Pastor", "PM", "Fixo"),
    ],
    "ajustes": [  # momento e câmera, antes, depois
        ("Leitura · câm. 1", "Leitor, PM (repetido com a câmera 2)", "Palco e leitor, PC. O momento ganhou o plano aberto."),
        ("Leitura · câm. 3", "Bíblia nas mãos, detalhe, na mão", "Congregação acompanhando a leitura, PC."),
        ("Pregação · câm. 1", "Pastor, PM (repetido com a câmera 2)", "Pastor e púlpito, PC, com espaço para ele andar."),
        ("Pregação · câm. 3", "Pastor, PP, na mão", "Congregação ouvindo, PC."),
        ("Resultado", "6 planos a rever e 2 momentos sem plano aberto", "Todos os 14 planos em OK depois das quatro trocas."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def dv_range(ws, rng, source, prompt):
    dv = DataValidation(type="list", formula1=source, allow_blank=True)
    dv.promptTitle, dv.prompt = "Opções", prompt
    dv.errorTitle, dv.error = "Valor inválido", "Escolha um dos momentos listados no alto da aba."
    dv.showErrorMessage = dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def cabecalho(ws, linhas, data, cols=("B", "C", "D", "G")):
    """Linhas de cabeçalho: (linha, rótulo, chave, formato). Devolve a cor das entradas."""
    l1, l2, v1, v2 = cols
    hb = WHITE if data else INPUT
    H = data["head"] if data else {}
    for r, rot, k, fmt in linhas:
        label(ws, f"{l1}{r}", rot, merge=f"{l1}{r}:{l2}{r}")
        inp(ws, f"{v1}{r}", H.get(k), merge=f"{v1}{r}:{v2}{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
        if fmt == DATE:
            dv_date(ws, f"{v1}{r}")
    return hb


# ================================================================== aba Posições de câmera
P_W = {"A": 2, "B": 6, "C": 30, "D": 18, "E": 15, "F": 15, "G": 11, "H": 11, "I": 28, "J": 11, "K": 11, "L": 19,
       "M": 12, "N": 52, "O": 2}
P1, PN = 15, 6


def angulo(desn, dist):
    """Ângulo vertical em graus, com sinal, por uma aproximação do arco tangente (erro abaixo de 0,3°, antes do arredondamento)."""
    x, u = f"(ABS({desn})/{dist})", f"({dist}/ABS({desn}))"
    return (f'ROUND(IF({desn}<0,-1,1)*57.2958*IF(ABS({desn})<={dist},{x}*0.7854+0.273*{x}*(1-{x}),'
            f'1.5708-{u}*0.7854-0.273*{u}*(1-{u})),0)')


def aba_posicoes(ws, tab):
    widths(ws, P_W)
    title(ws, "Posições de câmera — plano de posições e ângulo vertical",
          "Preencha as células em amarelo-claro. As células cinza são calculadas.", "N")
    cabecalho(ws, [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                   (7, "Diretor de imagem", "diretor", None)], None, ("B", "C", "D", "F"))
    refs = [(4, "Altura do palco (m)", 0.4, "Do piso da plateia ao piso do palco. Sem palco: 0."),
            (5, "Olhos de quem fala, em pé (m)", 1.6, "Acima do piso em que a pessoa está. Sentada: cerca de 1,2."),
            (6, "Altura dos olhos, do piso da plateia (m)", None, "Calculada: palco mais olhos."),
            (7, "Ângulo normal até (°)", 10, "Regra prática deste material."),
            (8, "Ângulo evidente acima de (°)", 25, "Regra prática deste material.")]
    for r, rot, v, dica in refs:
        label(ws, f"H{r}", rot, merge=f"H{r}:J{r}")
        if r == 6:
            calc(ws, f"K{r}", '=IF(OR(K4="",K5=""),"",K4+K5)', fmt="0.00")
        else:
            inp(ws, f"K{r}", v, h="center", fmt="0.00" if r < 6 else "0")
        hint(ws, f"L{r}", dica, merge=f"L{r}:N{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, "K4:K5")
    dv_number(ws, "K7:K8")

    band(ws, 10, "Câmeras e posições", "N")
    heads = [("B", "Câm.", INK), ("C", "Posição", INK), ("D", "Lado da sala", INK), ("E", "Suporte", INK),
             ("F", "Operador", INK), ("G", "Distância", BLUE), ("H", "Altura da lente", BLUE), ("I", "Planos que faz", INK),
             ("J", "Desnível", TEAL), ("K", "Ângulo", TEAL), ("L", "Leitura", TEAL), ("M", "Situação", TEAL),
             ("N", "O que fazer", TEAL)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "onde fica", "lista", "lista", "nome", "m, até quem fala", "m, do piso",
                                         "ex.: PG, PM", "lente − olhos (m)", "+ alta, − baixa", "", "", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 27
    for col, v in zip("BCDEFGHIJKLMN", ["Ex.", "Central, ao fundo, no chão", "Fundo da sala", "Tripé fluido", "Operador 1",
                                         12, 1.5, "PG e PM", -0.5, -2, "Normal", "OK", "Manter"]):
        ex(ws, f"{col}13", v, h="left" if col in "CIN" else "center", fmt="0.00" if col in "GHJ" else (DEG if col == "K" else None))
    ws.row_dimensions[13].height = 21.75
    note(ws, "G11", "Distância horizontal da lente até quem aparece, em metros.\nMeça com trena ou conte os passos (um passo ≈ 0,75 m).")
    note(ws, "H11", "Altura da lente, em metros, medida do piso da plateia.")
    note(ws, "K11", "Positivo: câmera acima dos olhos (plongée).\nNegativo: câmera abaixo dos olhos (contra-plongée).")
    note(ws, "D11", "Fundo do palco fica do outro lado do eixo para quem está no palco: serve só para mostrar o público.")

    P2 = P1 + PN - 1
    for k in range(PN):
        r = P1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDEFI":
            inp(ws, f"{col}{r}", h="center" if col in "DE" else "left")
        for col in "GH":
            inp(ws, f"{col}{r}", h="center", fmt="0.00")
        G, H, J, K, L, M = (f"{c}{r}" for c in "GHJKLM")
        calc(ws, J, f'=IF(OR({H}="",$K$6=""),"",ROUND({H}-$K$6,2))', fmt='+0.00;-0.00;0.00', b=False)
        calc(ws, K, f'=IF(OR({J}="",{G}=""),"",IF({G}<=0,"",{angulo(J, G)}))', fmt=DEG)
        calc(ws, L, f'=IF({K}="","",IF(ABS({K})<=$K$7,"Normal",IF({K}>0,IF(ABS({K})<=$K$8,"Plongée leve","Plongée"),'
                    f'IF(ABS({K})<=$K$8,"Contra-plongée leve","Contra-plongée"))))', b=False)
        calc(ws, M, f'=IF(AND(C{r}="",{G}=""),"",IF(D{r}="Fundo do palco","Outro lado",IF({L}="","Medir",'
                    f'IF({L}="Normal","OK",IF(OR({L}="Plongée leve",{L}="Contra-plongée leve"),"Atenção","Revisar")))))')
        calc(ws, f"N{r}", f'=IF({M}="","",IF({M}="Outro lado","Do outro lado do eixo do palco: use só para mostrar o público",'
                          f'IF({M}="Medir","Informe a distância e a altura da lente",IF({M}="OK","Manter",'
                          f'IF({K}>0,IF({M}="Atenção","Um pouco alta: serve ao plano geral; para o rosto, baixe a lente ou afaste a câmera",'
                          f'"Alta demais para o rosto: baixe a lente, afaste a câmera ou use só no plano geral"),'
                          f'IF({M}="Atenção","Um pouco baixa: suba a lente ou afaste a câmera",'
                          f'"Baixa demais: suba a lente (praticável) ou afaste a câmera"))))))', h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 30
    dv_list(ws, f"D{P1}:D{P2}", LADOS, "Onde fica a câmera em relação ao palco.")
    dv_list(ws, f"E{P1}:E{P2}", SUPORTES, "Tipo de suporte da câmera.")
    dv_number(ws, f"G{P1}:H{P2}")
    cf_equal(ws, f"M{P1}:M{P2}", [("OK", GREEN), ("Atenção", YELLOW), ("Medir", YELLOW), ("Outro lado", YELLOW), ("Revisar", RED)])

    s = P2 + 2
    band(ws, s, "Resumo automático", "N")
    rng = f"M{P1}:M{P2}"
    linhas = [("Câmeras listadas", f"=COUNTA(C{P1}:C{P2})"),
              ("Ângulo normal (OK)", f'=COUNTIF({rng},"OK")'),
              ("Ângulo a conferir (Atenção)", f'=COUNTIF({rng},"Atenção")'),
              ("Ângulo a revisar", f'=COUNTIF({rng},"Revisar")'),
              ("Do outro lado do eixo", f'=COUNTIF({rng},"Outro lado")')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "F", "G")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:F{a}", h="right")
    calc(ws, f"G{a}", f'=IF(G{s + 1}=0,"Liste as câmeras",IF(COUNTIF({rng},"Medir")>0,"Há câmera sem medida",'
                      f'IF(G{s + 4}>0,"Há câmera com ângulo a revisar",IF(G{s + 5}>0,"Há câmera do outro lado do eixo: só para o público",'
                      f'IF(G{s + 3}>0,"Há câmera com ângulo a conferir","OK")))))', b=False, merge=f"G{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"G{a}:N{a}")
    hint(ws, f"B{a + 1}", "O ângulo usa uma aproximação com erro menor que um grau. Os limites de 10° e 25° são regra prática deste material, e não norma.",
         merge=f"B{a + 1}:N{a + 1}", h="left")
    ws.freeze_panes = f"C{P1}"
    setup(ws, tab, f"B1:N{a + 1}", fit_height=True)


# ================================================================== aba Lista de planos
L_W = {"A": 2, "B": 6, "C": 18, "D": 9, "E": 28, "F": 10, "G": 13, "H": 15, "I": 9, "J": 19, "K": 50, "L": 26, "M": 2}
CAM1, NCAM = 10, 4
MOM1, NMOM = 17, 8
LP1, NLP = 30, 36


def aba_planos(ws, tab, data=None):
    widths(ws, L_W)
    is_ex = data is not None
    title(ws, "Lista de planos — o que cada câmera mostra em cada momento",
          "Exemplo preenchido, para consulta. Use a aba Lista de planos para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro: câmeras, momentos e planos. As células cinza são calculadas.", "L")
    hb = cabecalho(ws, [(4, "Evento", "evento", None), (5, "Data", "data", DATE), (6, "Diretor de imagem", "diretor", None)],
                   data, ("B", "C", "D", "G"))
    hint(ws, "H4", "Regras conferidas: sempre um plano aberto (PG ou PC) por momento; duas câmeras não repetem assunto e plano; "
                   "plano fechado só em tripé.", merge="H4:L6", h="left")

    # câmeras
    band(ws, CAM1 - 2, "Câmeras", "L")
    for col, t, m in [("B", "Câm.", None), ("C", "Posição", "C9:E9"), ("F", "Suporte", "F9:G9"), ("H", "Operador", "H9:I9"),
                      ("J", "Planos na lista", None), ("K", "Planos fechados (PP, close, detalhe)", "K9:L9")]:
        put(ws, f"{col}9", t, f=font(10, True, c=WHITE), bg=INK, h="center", merge=m)
    ws.row_dimensions[9].height = 21.75
    D1, D2 = f"$D${LP1}", f"$D${LP1 + NLP - 1}"
    F1, F2 = f"$F${LP1}", f"$F${LP1 + NLP - 1}"
    for k in range(NCAM):
        r = CAM1 + k
        c = data["cams"][k] if is_ex and k < len(data["cams"]) else (None, None, None)
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0], merge=f"C{r}:E{r}", bg=hb)
        inp(ws, f"F{r}", c[1], merge=f"F{r}:G{r}", bg=hb, h="center")
        inp(ws, f"H{r}", c[2], merge=f"H{r}:I{r}", bg=hb)
        calc(ws, f"J{r}", f'=IF(AND(C{r}="",F{r}=""),"",COUNTIF({D1}:{D2},B{r}))')
        fech = "+".join(f'COUNTIFS({D1}:{D2},B{r},{F1}:{F2},"{p}")' for p in ("PP", "Close", "Detalhe"))
        calc(ws, f"K{r}", f'=IF(AND(C{r}="",F{r}=""),"",{fech})', merge=f"K{r}:L{r}")
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"F{CAM1}:F{CAM1 + NCAM - 1}", SUPORTES, "Tipo de suporte da câmera.")

    # momentos
    band(ws, MOM1 - 2, "Momentos do evento, na ordem", "L")
    for col, t, m in [("B", "nº", None), ("C", "Momento", "C16:E16"), ("F", "Planos", None), ("G", "Abertos", None),
                      ("H", "Repetidos", None), ("I", "Situação", "I16:J16"), ("K", "O que fazer", "K16:L16")]:
        put(ws, f"{col}16", t, f=font(10, True, c=WHITE), bg=INK if col in "BC" else TEAL, h="center", merge=m)
    ws.row_dimensions[16].height = 21.75
    C1_, C2_ = f"$C${LP1}", f"$C${LP1 + NLP - 1}"
    I1_, I2_ = f"$I${LP1}", f"$I${LP1 + NLP - 1}"
    J1_, J2_ = f"$J${LP1}", f"$J${LP1 + NLP - 1}"
    for k in range(NMOM):
        r = MOM1 + k
        m = data["momentos"][k] if is_ex and k < len(data["momentos"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", m, merge=f"C{r}:E{r}", bg=hb)
        calc(ws, f"F{r}", f'=IF(C{r}="","",COUNTIF({C1_}:{C2_},C{r}))')
        calc(ws, f"G{r}", f'=IF(C{r}="","",COUNTIFS({C1_}:{C2_},C{r},{I1_}:{I2_},"Sim"))')
        calc(ws, f"H{r}", f'=IF(C{r}="","",COUNTIFS({C1_}:{C2_},C{r},{J1_}:{J2_},"Plano repetido"))')
        calc(ws, f"I{r}", f'=IF(C{r}="","",IF(F{r}=0,"Sem planos",IF(G{r}=0,"Sem plano aberto",IF(H{r}>0,"Plano repetido","OK"))))',
             merge=f"I{r}:J{r}")
        calc(ws, f"K{r}", f'=IF(I{r}="","",IF(I{r}="Sem planos","Inclua os planos deste momento na lista",'
                          f'IF(I{r}="Sem plano aberto","Deixe uma câmera em PG ou PC: é o plano de segurança",'
                          f'IF(I{r}="Plano repetido","Duas câmeras com o mesmo assunto e plano: mude uma delas","Manter"))))',
             merge=f"K{r}:L{r}", h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 21.75
    cf_equal(ws, f"I{MOM1}:I{MOM1 + NMOM - 1}", [("OK", GREEN), ("Sem planos", YELLOW), ("Plano repetido", RED), ("Sem plano aberto", RED)])

    # lista
    band(ws, LP1 - 4, "Lista de planos (shot list)", "L")
    for col, t, bg in [("B", "#", INK), ("C", "Momento", INK), ("D", "Câmera", INK), ("E", "Assunto", INK), ("F", "Plano", BLUE),
                       ("G", "Movimento", BLUE), ("H", "Suporte", TEAL), ("I", "Aberto", TEAL), ("J", "Situação", TEAL),
                       ("K", "O que fazer", TEAL), ("L", "Observação", INK)]:
        head(ws, f"{col}{LP1 - 3}", t, bg=bg)
    ws.row_dimensions[LP1 - 3].height = 30
    for col, t in zip("BCDEFGHIJKL", ["", "lista", "1 a 4", "quem ou o quê", "lista", "lista", "da câmera", "PG ou PC", "", "", ""]):
        hint(ws, f"{col}{LP1 - 2}", t)
    ws.row_dimensions[LP1 - 2].height = 21.75
    if is_ex:
        hint(ws, f"B{LP1 - 1}", "A primeira lista do exemplo 1, antes das correções. As colunas de H a K são calculadas.",
             merge=f"B{LP1 - 1}:L{LP1 - 1}", h="left")
    else:
        for col, v in zip("BCDEFGHIJK", ["Ex.", "Pregação", 2, "Pastor", "PM", "Acompanha", "Tripé fluido", "Não", "OK", "Manter"]):
            ex(ws, f"{col}{LP1 - 1}", v, h="left" if col in "EK" else "center")
        ex(ws, f"L{LP1 - 1}", "")
        note(ws, f"C{LP1 - 3}", "Escolha um dos momentos listados no bloco Momentos do evento, acima.")
        note(ws, f"F{LP1 - 3}", "PG geral, PC conjunto, PA americano, PM médio, PP primeiro plano, Close, Detalhe.")
        note(ws, f"J{LP1 - 3}", "Plano repetido: outra câmera tem o mesmo assunto e plano no mesmo momento.\n"
                                "Fechado na mão: PP, close ou detalhe em mão, monopé ou estabilizador.\n"
                                "Sem plano aberto: nenhuma câmera em PG ou PC neste momento.")
    ws.row_dimensions[LP1 - 1].height = 21.75

    LP2 = LP1 + NLP - 1
    E1_, E2_ = f"$E${LP1}", f"$E${LP2}"
    bg = WHITE if is_ex else INPUT
    for k in range(NLP):
        r = LP1 + k
        p = data["planos"][k] if is_ex and k < len(data["planos"]) else (None,) * 5
        num(ws, f"B{r}", k + 1)
        for col, v in zip("CDEFG", p):
            inp(ws, f"{col}{r}", v, h="left" if col in "CE" else "center", bg=bg)
        C, D, E, F, H, I, J = (f"{c}{r}" for c in "CDEFHIJ")
        look = f"VLOOKUP({D},$B${CAM1}:$F${CAM1 + NCAM - 1},5,FALSE)"
        calc(ws, H, f'=IF({D}="","",IFERROR(IF({look}&""="","Sem suporte",{look}&""),"Câmera não listada"))', b=False)
        calc(ws, I, f'=IF({F}="","",IF(OR({F}="PG",{F}="PC"),"Sim","Não"))', b=False)
        rep = f'COUNTIFS({C1_}:{C2_},{C},{E1_}:{E2_},{E},{F1}:{F2},{F})>1'
        mao = f'AND(OR({H}="Mão",{H}="Monopé",{H}="Estabilizador"),OR({F}="PP",{F}="Close",{F}="Detalhe"))'
        sem = f'COUNTIFS({C1_}:{C2_},{C},{I1_}:{I2_},"Sim")=0'
        calc(ws, J, f'=IF(AND({C}="",{D}="",{E}="",{F}=""),"",IF(OR({C}="",{D}="",{E}="",{F}=""),"Incompleto",'
                    f'IF({rep},"Plano repetido",IF({mao},"Fechado na mão",IF({sem},"Sem plano aberto","OK")))))')
        calc(ws, f"K{r}", f'=IF({J}="","",IF({J}="Incompleto","Preencha momento, câmera, assunto e plano",'
                          f'IF({J}="Plano repetido","Outra câmera tem o mesmo assunto e plano neste momento: mude o plano ou o assunto",'
                          f'IF({J}="Fechado na mão","Este suporte não sustenta plano fechado: abra o plano ou use tripé",'
                          f'IF({J}="Sem plano aberto","Nenhuma câmera em PG ou PC neste momento: abra uma delas","Manter")))))',
             h="left", b=False, sz=9)
        inp(ws, f"L{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_range(ws, f"C{LP1}:C{LP2}", f"$C${MOM1}:$C${MOM1 + NMOM - 1}", "Um dos momentos listados no alto da aba.")
    dv_list(ws, f"D{LP1}:D{LP2}", ["1", "2", "3", "4"], "Número da câmera, de 1 a 4.")
    dv_list(ws, f"F{LP1}:F{LP2}", PLANOS, "PG, PC, PA, PM, PP, Close ou Detalhe.")
    dv_list(ws, f"G{LP1}:G{LP2}", MOVS, "Fixo, Pan, Tilt, Zoom, Travelling ou Acompanha.")
    cf_equal(ws, f"J{LP1}:J{LP2}", [("OK", GREEN), ("Incompleto", YELLOW), ("Sem plano aberto", YELLOW),
                                     ("Fechado na mão", YELLOW), ("Plano repetido", RED)])

    s = LP2 + 2
    band(ws, s, "Resumo automático", "L")
    cnt = lambda t: f'COUNTIF({J1_}:{J2_},"{t}")'
    linhas = [("Planos na lista", f"=COUNTA(F{LP1}:F{LP2})"),
              ("Planos em OK", f"={cnt('OK')}"),
              ("Planos repetidos", f"={cnt('Plano repetido')}"),
              ("Planos fechados na mão", f"={cnt('Fechado na mão')}"),
              ("Planos a rever (todos os avisos)", f"=COUNTA(F{LP1}:F{LP2})-{cnt('OK')}"),
              ("Momentos sem plano aberto", f'=COUNTIF(I{MOM1}:I{MOM1 + NMOM - 1},"Sem plano aberto")')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Preencha a lista de planos",IF(COUNTIF({J1_}:{J2_},"Incompleto")>0,"Há linha incompleta",'
                      f'IF(F{s + 6}>0,"Há momento sem plano aberto: o diretor fica sem plano de segurança",'
                      f'IF(F{s + 3}>0,"Há planos repetidos: não dão corte",IF(F{s + 4}>0,"Há plano fechado em suporte de mão","OK")))))',
         b=False, merge=f"F{a}:L{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:L{a}")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "L")
        r = end + 3
        put(ws, f"B{r}", "Momento e câmera", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:D{r}")
        put(ws, f"E{r}", "Na primeira lista", f=font(10, True), bg=GRAY, h="center", merge=f"E{r}:H{r}")
        put(ws, f"I{r}", "Como ficou", f=font(10, True), bg=GRAY, h="center", merge=f"I{r}:L{r}")
        ws.row_dimensions[r].height = 21.75
        for quem, antes, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", quem, f=font(10, True), bg=GRAY, merge=f"B{r}:D{r}")
            put(ws, f"E{r}", antes, merge=f"E{r}:H{r}")
            put(ws, f"I{r}", depois, merge=f"I{r}:L{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"C{LP1}" if not is_ex else None
    setup(ws, tab, f"B1:L{end}")


# ================================================================== aba Autoavaliação
A_W = {"A": 2, "B": 6, "C": 10, "D": 10, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 10, "M": 12,
       "N": 34, "O": 2}
CRIT = ["Plano pedido", "Olhos no terço superior", "Corte fora das articulações", "Espaço do olhar",
        "Horizonte nivelado", "Bordas e fundo limpos", "Quadro firme"]
AV1, NAV = 12, 12


def aba_autoavaliacao(ws, tab):
    widths(ws, A_W)
    title(ws, "Autoavaliação do quadro — revisão da gravação",
          "Pause a gravação em até doze pontos e marque cada critério. As células cinza são calculadas.", "N")
    cabecalho(ws, [(4, "Gravação", "g", None), (5, "Data", "d", DATE), (6, "Operador", "o", None), (7, "Câmera e posição", "c", None)],
              None, ("B", "C", "D", "H"))
    for r, rot, v, dica in [(4, "Nota para OK", 0.85, "Igual ou acima: o quadro está pronto."),
                            (5, "Nota mínima para treinar", 0.6, "Entre esta e a de OK: treinar. Abaixo: refazer.")]:
        label(ws, f"J{r}", rot, merge=f"J{r}:K{r}")
        inp(ws, f"L{r}", v, h="center", fmt="0%")
        hint(ws, f"M{r}", dica, merge=f"M{r}:N{r}", h="left")
    dv_number(ws, "L4:L5")
    hint(ws, "J6", "Marque Não se aplica quando o critério não vale para o quadro (espaço do olhar em quem olha para a câmera, por exemplo).",
         merge="J6:N7", h="left")

    band(ws, 9, "Quadros avaliados", "N")
    for col, t in [("B", "#"), ("C", "Tempo"), ("D", "Plano")]:
        head(ws, f"{col}10", t)
    for k, t in enumerate(CRIT):
        head(ws, f"{'EFGHIJK'[k]}10", t, bg=BLUE)
    for col, t in [("L", "Nota"), ("M", "Situação")]:
        head(ws, f"{col}10", t, bg=TEAL)
    head(ws, "N10", "Observação")
    ws.row_dimensions[10].height = 45
    for col, t in zip("BCDEFGHIJKLMN", ["", "mm:ss", "lista"] + ["Sim / Não"] * 7 + ["% de Sim", "", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 21.75
    AV2 = AV1 + NAV - 1
    for k in range(NAV):
        r = AV1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center")
        inp(ws, f"D{r}", h="center")
        for col in "EFGHIJK":
            inp(ws, f"{col}{r}", h="center")
        sim, nao = f'COUNTIF(E{r}:K{r},"Sim")', f'COUNTIF(E{r}:K{r},"Não")'
        calc(ws, f"L{r}", f'=IF({sim}+{nao}=0,"",{sim}/({sim}+{nao}))', fmt="0%")
        calc(ws, f"M{r}", f'=IF(L{r}="","",IF(L{r}>=$L$4,"OK",IF(L{r}>=$L$5,"Treinar","Refazer")))')
        inp(ws, f"N{r}")
        ws.row_dimensions[r].height = 19.5
    dv_list(ws, f"D{AV1}:D{AV2}", PLANOS, "PG, PC, PA, PM, PP, Close ou Detalhe.")
    dv_list(ws, f"E{AV1}:K{AV2}", ["Sim", "Não", "Não se aplica"], "O quadro atende ao critério?")
    cf_equal(ws, f"E{AV1}:K{AV2}", [("Sim", GREEN), ("Não", RED)])
    cf_equal(ws, f"M{AV1}:M{AV2}", [("OK", GREEN), ("Treinar", YELLOW), ("Refazer", RED)])

    r = AV2 + 1
    label(ws, f"B{r}", "Atendimento por critério", merge=f"B{r}:D{r}", h="right")
    for col in "EFGHIJK":
        sim, nao = f'COUNTIF({col}{AV1}:{col}{AV2},"Sim")', f'COUNTIF({col}{AV1}:{col}{AV2},"Não")'
        calc(ws, f"{col}{r}", f'=IF({sim}+{nao}=0,"",{sim}/({sim}+{nao}))', fmt="0%")
    ws.row_dimensions[r].height = 21.75
    crit_row = r

    s = r + 2
    band(ws, s, "Resumo automático", "N")
    sim, nao = f'COUNTIF(E{AV1}:K{AV2},"Sim")', f'COUNTIF(E{AV1}:K{AV2},"Não")'
    pr = f"E{crit_row}:K{crit_row}"
    summary(ws, s + 1, "Quadros avaliados", f"=COUNT(L{AV1}:L{AV2})", "F", "G")
    summary(ws, s + 2, "Nota geral", f'=IF({sim}+{nao}=0,"",{sim}/({sim}+{nao}))', "F", "G", fmt="0%")
    summary(ws, s + 3, "Quadros a refazer", f'=COUNTIF(M{AV1}:M{AV2},"Refazer")', "F", "G")
    label(ws, f"B{s + 4}", "Critério a treinar primeiro", merge=f"B{s + 4}:F{s + 4}", h="right")
    calc(ws, f"G{s + 4}", f'=IF(COUNT({pr})=0,"",IF(MIN({pr})>=$L$4,"Nenhum: todos os critérios estão na faixa de OK",'
                          f'INDEX(E10:K10,MATCH(MIN({pr}),{pr},0))))', merge=f"G{s + 4}:N{s + 4}", h="left")
    label(ws, f"B{s + 5}", "Aviso", merge=f"B{s + 5}:F{s + 5}", h="right")
    calc(ws, f"G{s + 5}", f'=IF(G{s + 1}=0,"Avalie pelo menos um quadro",IF(G{s + 2}<$L$5,"Nota geral baixa: refaça o exercício de bancada",'
                          f'IF(G{s + 2}<$L$4,"Treine o critério indicado e grave de novo","OK")))',
         b=False, merge=f"G{s + 5}:N{s + 5}", h="left")
    for rr in (s + 4, s + 5):
        ws.row_dimensions[rr].height = 21.75
    cf_ok(ws, f"G{s + 5}:N{s + 5}")
    hint(ws, f"B{s + 6}", "Os cortes de 85% e 60% são regra prática deste material. Avalie a gravação, e não a lembrança do que foi feito.",
         merge=f"B{s + 6}:N{s + 6}", h="left")
    ws.freeze_panes = f"E{AV1}"
    setup(ws, tab, f"B1:N{s + 6}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Enquadramento e movimento — modelo de plano de câmeras"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Enquadramento e movimento — como usar esta planilha",
           "Plano de posições de câmera, lista de planos por momento e autoavaliação do quadro.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Lado da sala, suporte, momento, câmera, plano, movimento e critérios aceitam só as opções da lista."),
        ("Números", "Digite distâncias e alturas em metros, só com o número: 12 ou 1,5."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Posições de câmera: informe a altura do palco e a dos olhos de quem fala, e liste as câmeras."),
        ("2", "Meça a distância e a altura da lente de cada câmera. Veja o ângulo, a leitura e o que fazer."),
        ("3", "Aba Lista de planos: preencha as câmeras (com o suporte) e os momentos do evento, na ordem."),
        ("4", "Liste os planos de cada câmera em cada momento: assunto, plano e movimento."),
        ("5", "Corrija os avisos: plano repetido, plano fechado na mão e momento sem plano aberto."),
        ("6", "Aba Checklist: confira os doze itens em cada câmera antes do evento."),
        ("7", "Depois do evento, aba Autoavaliação: avalie quadros da gravação e veja o critério a treinar."),
    ]),
    ("Abas da planilha", [
        ("Posições de câmera", "Até seis câmeras, com desnível, ângulo vertical aproximado, leitura e situação."),
        ("Lista de planos", "Até 4 câmeras, 8 momentos e 36 planos, com a conferência das regras de cobertura."),
        ("Autoavaliação", "Até doze quadros de uma gravação, em sete critérios, com nota e critério a treinar."),
        ("Referência", "Planos, composição, ângulo por distância, movimentos, suportes, cobertura e comandos."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Planos", "A primeira lista de planos de um culto com três câmeras (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Ângulo vertical", "Calculado pelo desnível entre a lente e os olhos de quem fala e pela distância horizontal, com uma aproximação de erro menor que um grau."),
        ("Limites do ângulo", "Até 10° conta como normal; acima de 25°, o efeito de plongée ou de contra-plongée é evidente. São regra prática deste material, e podem ser trocados no alto da aba."),
        ("Plano aberto", "PG e PC contam como plano aberto. Todo momento precisa de pelo menos um, como plano de segurança."),
        ("Plano repetido", "Mesmo momento, mesmo assunto e mesmo plano em duas câmeras. Escreva o assunto sempre do mesmo jeito, para a conferência funcionar."),
        ("Fechado na mão", "PP, close ou detalhe em câmera na mão, em monopé ou em estabilizador. Teste antes: o limite depende da câmera, da lente e de quem opera."),
        ("Autoavaliação", "Nota é a fração de Sim entre os critérios que se aplicam. 85% ou mais: OK; de 60% a 85%: treinar; abaixo: refazer."),
        ("Exemplos", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_posicoes(wb.create_sheet("Posições de câmera"), BLUE)
aba_planos(wb.create_sheet("Lista de planos"), AMBER)
aba_autoavaliacao(wb.create_sheet("Autoavaliação"), TEAL)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 30, "D": 34, "E": 52, "F": 2})
title(ws, "Referência — planos, composição e cobertura", "Valores e regras de referência, para servir de ponto de partida.", "E")


def tabela(r, titulo, cab, linhas, alt=31.5, center=False):
    band(ws, r, titulo, "E")
    r += 1
    for col, t in zip("BCDE", cab):
        put(ws, f"{col}{r}", t, f=font(10, True), bg=GRAY, h="center")
    ws.row_dimensions[r].height = 21.75
    for lin in linhas:
        r += 1
        put(ws, f"B{r}", lin[0], f=font(10, True))
        for col, t in zip("CDE", lin[1:]):
            put(ws, f"{col}{r}", t, h="center" if center else "left")
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(4, "Planos", ["Plano", "Onde termina o quadro", "Para que serve", "Em um culto ou em uma palestra"], [
    ("PG · Plano geral", "O ambiente inteiro", "Situar: o lugar, o público, o clima", "Abertura, congregação cantando, trocas de momento. Plano de segurança."),
    ("PC · Plano conjunto", "Um grupo de corpo inteiro", "Quem está junto e o que cada um faz", "A banda, o coral, a mesa de convidados."),
    ("PA · Plano americano", "No meio da coxa", "Espaço para quem anda e gesticula", "O pregador que caminha pelo palco, o regente."),
    ("PM · Plano médio", "Abaixo da cintura", "Quem fala, com os gestos", "Pregação, palestra, avisos: o plano de trabalho."),
    ("PP · Primeiro plano", "No peito", "O rosto manda no quadro", "Leitura, oração, apelo, o trecho forte de uma canção."),
    ("Close", "No alto dos ombros", "A emoção", "Poucos segundos, com a pessoa parada."),
    ("Detalhe", "Uma parte do corpo ou um objeto", "O que os outros planos não mostram", "As mãos no teclado, a Bíblia aberta. Cobre troca de câmera."),
], alt=21.75)
r = tabela(r, "Composição", ["Regra", "Como aplicar", "Erro comum", "Correção"], [
    ("Regra dos terços", "Assunto sobre as linhas da grade", "Tudo no centro exato", "Ligue a grade no visor e desloque o quadro."),
    ("Olhos no terço superior", "Do plano americano ao close", "Rosto no meio do quadro", "Tilt até os olhos tocarem a linha de cima."),
    ("Teto (headroom)", "Diminui quando o plano fecha", "Zoom sem corrigir o tilt", "Ao fechar, a câmera sobe; ao abrir, desce."),
    ("Espaço do olhar", "Espaço maior à frente do rosto ou do movimento", "Nariz na borda", "Pessoa no terço oposto ao olhar."),
    ("Corte do corpo", "Entre as articulações", "Corte no pescoço, na cintura, no joelho", "Abra ou feche até o peito, abaixo da cintura, o meio da coxa."),
    ("Horizonte", "Linhas da sala paralelas à borda", "Tripé desnivelado", "Bolha no centro e pan de teste de ponta a ponta."),
    ("Bordas e fundo", "Nada que distraia", "Pedestal saindo da cabeça, meio rosto na borda", "Um passo para o lado ou um pouco de zoom."),
], alt=21.75)

dists = [3, 5, 8, 12, 15, 20]
lin = []
for d in dists:
    lin.append((f"{d} m", *[f"{math.degrees(math.atan(h / d)):.0f}°" for h in (0.5, 1, 2)]))
r = tabela(r, "Ângulo vertical pela distância e pelo desnível entre a lente e os olhos",
           ["Distância horizontal", "Desnível de 0,5 m", "Desnível de 1 m", "Desnível de 2 m"], lin, alt=19.5, center=True)
hint(ws, f"B{r - 1}", "Até uns 10°: parece normal. De 10° a 25°: perceptível. Acima de 25°: evidente. Regra prática deste material.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
r += 1
r = tabela(r, "Movimentos", ["Movimento", "O que acontece", "Quando usar", "Cuidado"], [
    ("Pan (panorâmica)", "Giro para os lados", "Acompanhar quem anda, ir de um assunto a outro", "Tripé nivelado. Um objeto leva de 5 a 7 s para atravessar o quadro, como ponto de partida."),
    ("Tilt", "Inclinação para cima ou para baixo", "Corrigir o teto, ir das mãos ao rosto", "Contrapeso ajustado, ou a câmera cai ao soltar."),
    ("Zoom", "A lente fecha ou abre o ângulo", "Trocar de plano fora do ar", "Corrigir o tilt junto e conferir o foco."),
    ("Travelling", "A câmera inteira se desloca", "Dar profundidade, chegar perto", "Estabilizador, trilho ou rodas, e caminho livre."),
], alt=31.5)
r = tabela(r, "Suportes", ["Suporte", "Planos que sustenta bem", "Movimento", "Observação"], [
    ("Tripé com cabeça fluida", "Todos, do geral ao close", "Pan e tilt suaves", "O suporte padrão. O plano fechado pede piso firme."),
    ("Tripé com cabeça de fotografia", "Todos, com a câmera parada", "Nenhum no ar", "Câmera fixa sem operador."),
    ("Monopé", "Do geral ao médio", "Pan curta", "Rápido de mudar de lugar."),
    ("Mão ou ombro", "Do geral ao médio, lente aberta", "Travelling curto", "Cansa: trechos curtos."),
    ("Estabilizador", "Do geral ao médio, em movimento", "Travelling longo", "Equilíbrio antes de ligar e bateria carregada."),
], alt=21.75)
r = tabela(r, "Cobertura com uma, duas e três câmeras", ["Câmeras", "Câmera 1", "Câmera 2", "Câmera 3"], [
    ("Uma", "Central: PA ou PM folgado de quem fala; PG no louvor", "–", "–"),
    ("Duas", "Central, aberta (PG ou PC): plano de segurança", "Fechada em quem fala ou canta (PM, PP)", "–"),
    ("Três", "Central, aberta: plano de segurança", "Fechada em quem fala (PM, PP)", "Móvel: banda, detalhes, congregação"),
], alt=31.5)
r = tabela(r, "Comandos do diretor (combine a lista com a equipe)", ["Comando", "O que o operador faz", "", ""], [
    ("“Abre” / “fecha”", "Zoom para um plano mais aberto ou mais fechado", "", ""),
    ("“Sobe” / “desce”", "Tilt para cima ou para baixo", "", ""),
    ("“Esquerda” / “direita”", "Pan, em geral pelo lado do quadro", "", ""),
    ("“Corrige”", "Acerta teto, nível, espaço do olhar ou foco", "", ""),
    ("“Segura”", "Não mexe em nada", "", ""),
    ("“Acompanha”", "Segue a pessoa mantendo o plano", "", ""),
    ("“Livre”", "Saiu do ar: pode buscar outro quadro", "", ""),
    ("“Volta para a base”", "Retorna ao plano de base do momento", "", ""),
], alt=19.5)
hint(ws, f"B{r - 1}", "Valores e regras de referência deste material, e não norma. Os nomes dos planos variam entre equipes: combine a lista antes do evento.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", landscape=False)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist da câmera",
          "Confira em cada câmera, antes de o evento começar. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_planos(wb.create_sheet("Exemplo - Planos"), MUTED, EX_PLANOS)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
