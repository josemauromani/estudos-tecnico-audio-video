# -*- coding: utf-8 -*-
"""Gera Lentes-Foco-e-Cor-modelo.xlsx: calculadora de enquadramento, ficha de balanço de branco,
referência, checklist e os exemplos 1 e 2 do treinamento."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIM_NAO = ["Sim", "Não"]
REF = "'Referência'"

# ------------------------------------------------------------------ sensores (aba Referência)
# nome, largura (mm), altura (mm), onde aparece. A linha "Outro" recebe a largura digitada pelo usuário.
SENSORES = [
    ("Full frame", 36.0, 24.0, "Câmeras fotográficas e de cinema de sensor grande"),
    ("Super 35", 24.0, 13.5, "Câmeras de cinema digital (varia de 22 a 25 mm)"),
    ("APS-C (fator 1,5)", 23.5, 15.6, "Câmeras fotográficas intermediárias"),
    ("APS-C (fator 1,6)", 22.3, 14.9, "Câmeras fotográficas intermediárias de alguns fabricantes"),
    ("Micro 4/3", 17.3, 13.0, "Câmeras compactas de lente intercambiável"),
    ("1 polegada", 13.2, 8.8, "Filmadoras de mão e PTZ de linha mais alta"),
    ("2/3 polegada", 9.6, 5.4, "Câmeras de estúdio e de ombro de emissoras"),
    ("1/2,3 polegada", 6.2, 4.6, "Filmadoras compactas e algumas PTZ"),
    ("1/2,8 polegada", 5.6, 3.2, "A maioria das câmeras PTZ"),
    ("1/3 polegada", 4.8, 3.6, "PTZ simples e webcams"),
    ("Outro", None, None, "Digite a largura do sensor da sua câmera (ficha técnica, em mm)"),
]
S1 = 6                      # primeira linha da tabela de sensores na aba Referência
S2 = S1 + len(SENSORES) - 1
SENS_NOMES = f"{REF}!$B${S1}:$B${S2}"
SENS_TAB = f"{REF}!$B${S1}:$C${S2}"

CHECK = [
    ("Lente", "A distância de cada câmera até o assunto foi medida, e a distância focal necessária cabe na lente."),
    ("Lente", "O zoom digital está desligado em todas as câmeras, ou limitado à faixa sem perda."),
    ("Lente", "As lentes estão limpas e com para-sol, e a estabilização está de acordo com o suporte (mão ou tripé)."),
    ("Lente", "A abertura escolhida vale para toda a faixa de zoom: a imagem não escurece ao fechar o quadro."),
    ("Foco", "O modo de foco de cada câmera foi escolhido: manual para quadro fixo, automático com detecção de rosto para quem anda."),
    ("Foco", "O foco manual foi feito com o zoom fechado no rosto, com peaking ou ampliação."),
    ("Foco", "A faixa em foco cobre o espaço em que a pessoa se move."),
    ("Foco", "O foco automático foi testado com a luz do evento e com o fundo real (telão ligado)."),
    ("Cor", "Nenhuma câmera está com o balanço de branco em automático."),
    ("Cor", "O balanço foi feito com cartão branco no lugar do rosto, sob a luz do evento, com as cortinas fechadas."),
    ("Cor", "Todas as câmeras usam o mesmo perfil de imagem, e nenhuma está em perfil log sem correção."),
    ("Cor", "As câmeras foram comparadas lado a lado no mesmo monitor, e os valores de cada uma estão anotados."),
]

# Exemplo 1 do treinamento: quatro câmeras de uma igreja, antes da mudança de posição
EX_ENQ = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Equipe de vídeo da escala"},
    "cams": [  # posição, sensor, distância, cena, lente mín, lente máx, abertura
        ("Fundo, centro: pastor da cintura para cima", "Micro 4/3", 18, 1.6, 14, 140, 5.6),
        ("PTZ no fundo: palco inteiro", "1/2,8 polegada", 18, 10, 4.3, 129, 2.8),
        ("Lateral: close da cantora", "APS-C (fator 1,5)", 6, 1.0, 18, 55, 5.6),
        ("Frente do palco: banda inteira", "Full frame", 3, 8, 24, 70, 4),
    ],
    "ajustes": [  # câmera, o que foi feito, resultado
        ("1", "Levada para um praticável a 12 m do púlpito.", "130 mm, dentro da lente. Em f/5,6, faixa em foco de cerca de 1,4 m."),
        ("2", "Nada a mudar.", "10 mm, dentro da lente."),
        ("3", "O close (1,0 m) só caberia a 2,3 m da cantora. O plano pedido virou plano médio, de 2,6 m.", "55 mm, no fim da lente."),
        ("4", "Recuada para 5,5 m do palco.", "24,8 mm, dentro da lente."),
        ("Todas", "Zoom digital desligado no menu.", ""),
    ],
}

# Exemplo 2 do treinamento: a cor que muda a cada corte, como as câmeras foram encontradas
EX_WB = {
    "head": {"evento": "Formatura no auditório da escola", "local": "Auditório com janelas laterais",
             "data": date(2026, 12, 12), "operador": "Operador de vídeo do evento"},
    "ref": (3200, 1, 500, 200),
    "cams": [  # posição, tipo, modo, kelvin, tint, perfil, conferido
        ("Centro, em tripé (referência)", "Filmadora", "Cartão branco", 3300, 0, "Padrão", "Sim"),
        ("Lateral esquerda", "Câmera de lente intercambiável", "Kelvin", 3200, 0, "Vívido", "Não"),
        ("Móvel, perto do palco", "Filmadora compacta", "Preset", 5600, 0, "Padrão", "Não"),
        ("Plano geral, no fundo", "PTZ", "Automático", None, None, "Padrão", "Não"),
    ],
    "ajustes": [  # câmera, antes, depois
        ("Todas", "Cortinas das janelas abertas.", "Cortinas fechadas antes do balanço."),
        ("1", "Cartão branco: 3.300 K.", "Refeito depois de fechar as cortinas: continuou em 3.300 K."),
        ("2", "Kelvin 3.200 K, perfil vívido.", "Perfil padrão. Balanço com cartão: 3.250 K."),
        ("3", "Preset de luz do dia (5.600 K): imagem laranja.", "Balanço com cartão: 3.350 K."),
        ("4", "Automático: a cor variava com a janela no quadro.", "Manual. Balanço com cartão: 3.300 K, com um passo de tint para o magenta."),
        ("Resultado", "Quatro cores diferentes no corte.", "Branco igual nas quatro câmeras, comparado no multiview."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def dv_range(ws, rng, source, prompt):
    """Lista suspensa a partir de um intervalo de células (aceita opções com vírgula)."""
    dv = DataValidation(type="list", formula1=source, allow_blank=True)
    dv.promptTitle, dv.prompt = "Opções", prompt
    dv.errorTitle, dv.error = "Valor inválido", "Escolha uma das opções da lista."
    dv.showErrorMessage = dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def cabecalho(ws, H, hb, ate):
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None),
                           (6, "Data", "data", DATE), (7, "Operador", "operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:{ate}{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")


def ajustes(ws, end, data, cols, titulos, last):
    """Tabela "O que foi ajustado" das abas de exemplo. cols = (col1, (ini2, fim2), (ini3, fim3))."""
    band(ws, end + 2, "O que foi ajustado", last)
    r = end + 3
    c1, (a2, b2), (a3, b3) = cols
    put(ws, f"{c1}{r}", titulos[0], f=font(10, True), bg=GRAY, h="center")
    put(ws, f"{a2}{r}", titulos[1], f=font(10, True), bg=GRAY, h="center", merge=f"{a2}{r}:{b2}{r}")
    put(ws, f"{a3}{r}", titulos[2], f=font(10, True), bg=GRAY, h="center", merge=f"{a3}{r}:{b3}{r}")
    ws.row_dimensions[r].height = 21.75
    for c, x, y in data:
        r += 1
        put(ws, f"{c1}{r}", c, f=font(10, True, c=MUTED), bg=GRAY, h="center")
        put(ws, f"{a2}{r}", x, merge=f"{a2}{r}:{b2}{r}")
        put(ws, f"{a3}{r}", y, merge=f"{a3}{r}:{b3}{r}")
        ws.row_dimensions[r].height = 30
    return r


# ================================================================== aba Enquadramento
E_W = {"A": 2, "B": 5, "C": 30, "D": 17, "E": 9, "F": 9, "G": 9, "H": 8, "I": 8, "J": 8, "K": 10, "L": 10,
       "M": 13, "N": 9, "O": 9, "P": 9, "Q": 10, "R": 8, "S": 9, "T": 44, "U": 26, "V": 2}
E1 = 13
NE = 12
ENQ_CF = [("OK", GREEN), ("Falta tele", YELLOW), ("Falta angular", YELLOW)]


def aba_enq(ws, tab, data=None):
    widths(ws, E_W)
    is_ex = data is not None
    title(ws, "Enquadramento — que lente usar de onde a câmera está",
          "Exemplo preenchido, para consulta. Use a aba Enquadramento para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro. A planilha calcula a distância focal necessária, a cena obtida e a faixa em foco.", "U")
    hb = WHITE if is_ex else INPUT
    cabecalho(ws, data["head"] if is_ex else {}, hb, "H")
    hint(ws, "J4", "largura da cena = distância × largura do sensor ÷ distância focal", merge="J4:U4", h="left")
    hint(ws, "J5", "distância focal necessária = distância × largura do sensor ÷ largura desejada", merge="J5:U5", h="left")
    hint(ws, "J6", "A faixa em foco é uma estimativa usual (círculo de confusão = largura do sensor ÷ 1200), para assunto distante.", merge="J6:U6", h="left")
    hint(ws, "J7", "Sensor que não está na lista: escolha Outro e digite a largura na aba Referência.", merge="J7:U7", h="left")

    band(ws, 9, "Calculadora de enquadramento, uma linha por câmera", "U")
    heads = [("B", "#", INK), ("C", "Câmera e quadro pedido", INK), ("D", "Sensor", INK), ("E", "Largura do sensor", INK),
             ("F", "Distância", BLUE), ("G", "Cena desejada", BLUE), ("H", "Lente mín.", BLUE), ("I", "Lente máx.", BLUE),
             ("J", "Abertura", AMBER),
             ("K", "Focal necessária", BLUE), ("L", "Equivalente full frame", BLUE), ("M", "Situação", BLUE),
             ("N", "Focal a usar", BLUE), ("O", "Cena obtida", BLUE), ("P", "Altura (16:9)", BLUE),
             ("Q", "Distância que fecha a cena", BLUE), ("R", "Foco de", AMBER), ("S", "Foco até", AMBER),
             ("T", "O que fazer", INK), ("U", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 42
    for col, t in zip("BCDEFGHIJKLMNOPQRSTU", ["nº", "posição e o que mostra", "lista", "mm", "m", "largura, m", "mm", "mm",
                                              "f/", "mm", "mm", "", "mm", "largura, m", "m", "m", "m", "m", "", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 21.75
    if is_ex:
        hint(ws, "B12", "As quatro câmeras do exemplo 1, antes da mudança de posição. As colunas cinza são calculadas.", merge="B12:U12", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNOPQRS", ["Ex.", "Fundo: pastor da cintura para cima", "Micro 4/3", 17.3, 18, 1.6, 14, 140, 5.6,
                                                 194.6, 405, "Falta tele", 140, 2.22, 1.25, 12.9, 16.77, 19.43]):
            ex(ws, f"{col}12", v, h="left" if col in "CD" else "center")
        ex(ws, "T12", "Aproxime a câmera até a distância indicada, use lente mais longa ou aceite a cena obtida")
        ex(ws, "U12", "")
        note(ws, "F10", "Distância da câmera até a pessoa ou o centro da cena, em metros.\nMeça com trena ou em passos largos (cerca de 1 m cada).")
        note(ws, "G10", "Largura da cena que o quadro deve mostrar, em metros.\nUma pessoa da cintura para cima: cerca de 1,6 m. Corpo inteiro: cerca de 4 m. Rosto e ombros: cerca de 0,8 m.")
        note(ws, "H10", "Menor e maior distância focal da lente, em mm, como estão escritas na lente ou na ficha técnica.\nPara lente fixa, repita o mesmo valor.")
        note(ws, "J10", "Abertura que você vai usar, só o número: 2.8, 4, 5.6.\nUsada só para estimar a faixa em foco. Pode ficar em branco.")
        note(ws, "Q10", "Falta tele: até onde aproximar a câmera para fechar a cena com a lente no máximo.\nFalta angular: até onde afastar a câmera para abrir a cena com a lente no mínimo.")
        note(ws, "R10", "Faixa em foco estimada, com o foco na distância da coluna F e a lente na focal a usar.\nÉ uma estimativa: em tela grande ou em 4K, a faixa que parece nítida é menor.")
    ws.row_dimensions[12].height = 24

    E2 = E1 + NE - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NE):
        r = E1 + k
        c = data["cams"][k] if is_ex and k < len(data["cams"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, bg=bg)
        for col, j, fmt in (("F", 2, "0.0"), ("G", 3, "0.0#"), ("H", 4, "0.#"), ("I", 5, "0.#"), ("J", 6, "0.0")):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg, fmt=fmt)
        D, E, F, G, Hh, I, J, K, M, N = (f"{x}{r}" for x in "DEFGHIJKMN")
        calc(ws, E, f'=IF({D}="","",IFERROR(VLOOKUP({D},{SENS_TAB},2,FALSE),""))', b=False, fmt="0.0")
        base_ok = f'OR({E}="",{F}="",{G}="",{G}=0,{E}=0)'
        calc(ws, K, f'=IF({base_ok},"",ROUND({F}*{E}/{G},1))', fmt="0.0")
        calc(ws, f"L{r}", f'=IF({K}="","",ROUND({K}*36/{E},0))', b=False, fmt="0")
        calc(ws, M, f'=IF(OR({K}="",{Hh}="",{I}=""),"",IF({K}>{I},"Falta tele",IF({K}<{Hh},"Falta angular","OK")))')
        calc(ws, N, f'=IF({K}="","",IF(OR({Hh}="",{I}=""),{K},MIN(MAX({K},{Hh}),{I})))', b=False, fmt="0.0")
        calc(ws, f"O{r}", f'=IF({N}="","",ROUND({F}*{E}/{N},2))', fmt="0.00")
        calc(ws, f"P{r}", f'=IF(O{r}="","",ROUND(O{r}*9/16,2))', b=False, fmt="0.00")
        calc(ws, f"Q{r}", f'=IF({M}="Falta tele",ROUND({I}*{G}/{E},1),IF({M}="Falta angular",ROUND({Hh}*{G}/{E},1),""))', fmt="0.0")
        # profundidade de campo: hiperfocal H = f²/(N·c) + f, com c = largura do sensor ÷ 1200 (mm)
        hip = f'({N}^2/({J}*{E}/1200)+{N})'
        s = f'({F}*1000)'
        sem = f'OR({N}="",{J}="",{J}=0)'
        calc(ws, f"R{r}", f'=IF({sem},"",ROUND({s}*({hip}-{N})/({hip}+{s}-2*{N})/1000,2))', b=False, fmt="0.00")
        calc(ws, f"S{r}", f'=IF({sem},"",IF({s}>={hip},"infinito",ROUND({s}*({hip}-{N})/({hip}-{s})/1000,2)))', b=False, fmt="0.00")
        calc(ws, f"T{r}", f'=IF({M}="","",IF({M}="Falta tele","Aproxime a câmera até a distância indicada, use lente mais longa ou aceite a cena obtida",'
                          f'IF({M}="Falta angular","Afaste a câmera até a distância indicada, use lente mais aberta ou aceite a cena obtida",'
                          f'IF(AND(ISNUMBER(S{r}),ISNUMBER(R{r})),IF(S{r}-R{r}<0.5,"A lente alcança. Faixa em foco estreita: feche a íris se a pessoa anda",'
                          f'"A lente alcança. Deixe o zoom digital desligado"),"A lente alcança. Deixe o zoom digital desligado"))))',
             h="left", b=False, sz=9)
        inp(ws, f"U{r}", bg=bg)
        ws.row_dimensions[r].height = 30
    dv_range(ws, f"D{E1}:D{E2}", SENS_NOMES, "Escolha o formato do sensor. Outro: digite a largura na aba Referência.")
    dv_number(ws, f"F{E1}:J{E2}")
    cf_equal(ws, f"M{E1}:M{E2}", ENQ_CF)

    s = E2 + 2
    band(ws, s, "Resumo automático", "U")
    cont = lambda t: f'COUNTIF(M{E1}:M{E2},"{t}")'
    linhas = [("Câmeras registradas", f"=COUNTA(C{E1}:C{E2})"),
              ("Lente alcança (OK)", f'={cont("OK")}'),
              ("Falta tele", f'={cont("Falta tele")}'),
              ("Falta angular", f'={cont("Falta angular")}')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:G{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste as câmeras",IF({cont("OK")}+{cont("Falta tele")}+{cont("Falta angular")}<F{s + 1},'
                      f'"Há câmera sem sensor, distância, cena ou lente",IF(F{s + 3}+F{s + 4}>0,'
                      f'"Há quadro que a lente não alcança: decida antes de montar","OK")))',
         b=False, merge=f"F{a}:P{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:P{a}")
    end = a
    if is_ex:
        end = ajustes(ws, end, data["ajustes"], ("B", ("C", "L"), ("M", "U")), ("Câm.", "O que foi feito", "Resultado"), "U")
    ws.freeze_panes = f"D{E1}"
    setup(ws, tab, f"B1:U{end}", fit_height=True)


# ================================================================== aba Balanço de branco
W_W = {"A": 2, "B": 5, "C": 30, "D": 24, "E": 15, "F": 10, "G": 8, "H": 11, "I": 12, "J": 11, "K": 11,
       "L": 18, "M": 50, "N": 26, "O": 2}
W1 = 14
NW = 8
MODOS = ["Cartão branco", "Kelvin", "Preset", "Automático"]
PERFIS = ["Padrão", "Neutro", "Vívido", "Cinema", "Log", "Outro"]
WB_CF = [("OK", GREEN), ("Conferir no monitor", YELLOW), ("Diferente da referência", YELLOW), ("Perfil diferente", YELLOW),
         ("Sem kelvin", YELLOW), ("Longe da luz", RED), ("Automático", RED), ("Perfil log", RED)]


def aba_wb(ws, tab, data=None):
    widths(ws, W_W)
    is_ex = data is not None
    title(ws, "Balanço de branco — ficha de cor por câmera",
          "Exemplo preenchido, para consulta. Use a aba Balanço de branco para o seu evento." if is_ex
          else "Depois do balanço com cartão, anote o que cada câmera mostra. As células cinza conferem e apontam o que fazer.", "N")
    hb = WHITE if is_ex else INPUT
    cabecalho(ws, data["head"] if is_ex else {}, hb, "E")
    ref = data["ref"] if is_ex else (3200, 1, 500, 200)
    linhas = [("Luz principal (K)", ref[0], "Temperatura de cor da luz que ilumina o rosto. Veja a aba Referência."),
              ("Câmera de referência (nº)", ref[1], "A câmera principal. As outras são comparadas com ela."),
              ("Tolerância para a luz (K)", ref[2], "Diferença aceita entre o kelvin da câmera e a luz principal."),
              ("Tolerância entre câmeras (K)", ref[3], "Diferença aceita entre cada câmera e a de referência.")]
    for k, (rot, v, dica) in enumerate(linhas):
        r = 4 + k
        label(ws, f"G{r}", rot, merge=f"G{r}:I{r}")
        inp(ws, f"J{r}", v, h="center", bg=hb, fmt="0")
        hint(ws, f"K{r}", dica, merge=f"K{r}:N{r}", h="left")
    dv_number(ws, "J4:J7")
    LUZ, CREF, TL, TR = "$J$4", "$J$5", "$J$6", "$J$7"
    W2 = W1 + NW - 1
    hint(ws, "G8", "O kelvin que a câmera mostra depois do balanço é uma leitura dela. O que decide é a comparação no mesmo monitor.",
         merge="G8:N8", h="left")

    label(ws, "B9", "Kelvin da referência", merge="B9:D9", h="right")
    calc(ws, "E9", f'=IF(OR({CREF}="",{CREF}<1,{CREF}>{NW}),"",IF(INDEX(F{W1}:F{W2},{CREF})="","",INDEX(F{W1}:F{W2},{CREF})))', fmt="0")
    label(ws, "B10", "Perfil da referência", merge="B10:D10", h="right")
    calc(ws, "E10", f'=IF(OR({CREF}="",{CREF}<1,{CREF}>{NW}),"",IF(INDEX(H{W1}:H{W2},{CREF})="","",INDEX(H{W1}:H{W2},{CREF})))')
    KREF, PREF = "$E$9", "$E$10"

    band(ws, 11, "Ficha de cor, uma linha por câmera", "N")
    heads = [("B", "#", INK), ("C", "Câmera e posição", INK), ("D", "Tipo de câmera", INK), ("E", "Modo", TEAL),
             ("F", "Kelvin", TEAL), ("G", "Tint", TEAL), ("H", "Perfil", TEAL), ("I", "Comparada no monitor", TEAL),
             ("J", "Diferença para a luz", INK), ("K", "Diferença para a ref.", INK), ("L", "Situação", INK),
             ("M", "O que fazer", INK), ("N", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}12", t, bg=bg)
    ws.row_dimensions[12].height = 36
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "", "", "lista", "K", "passos", "lista", "Sim / Não", "K", "K", "", "", ""]):
        hint(ws, f"{col}13", t)
    ws.row_dimensions[13].height = 21.75
    if not is_ex:
        note(ws, "E12", "Cartão branco: balanço medido pela câmera.\nKelvin: número digitado.\nPreset: tungstênio, luz do dia etc.\nAutomático: AWB ou ATW.")
        note(ws, "F12", "O kelvin que a câmera mostra depois do balanço, ou o número digitado. Só o número: 3200.")
        note(ws, "G12", "Correção verde–magenta, no número que a câmera mostra. Deixe em branco se a câmera não tem esse ajuste.")
        note(ws, "I12", "A imagem desta câmera foi comparada com a de referência, lado a lado, no mesmo monitor (multiview)?")

    bg = WHITE if is_ex else INPUT
    for k in range(NW):
        r = W1 + k
        c = data["cams"][k] if is_ex and k < len(data["cams"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, bg=bg)
        inp(ws, f"E{r}", c[2] if c else None, h="center", bg=bg)
        inp(ws, f"F{r}", c[3] if c else None, h="center", bg=bg, fmt="0")
        inp(ws, f"G{r}", c[4] if c else None, h="center", bg=bg, fmt='+0;-0;0')
        inp(ws, f"H{r}", c[5] if c else None, h="center", bg=bg)
        inp(ws, f"I{r}", c[6] if c else None, h="center", bg=bg)
        C, E, F, Hc, I, J, K, L = (f"{x}{r}" for x in "CEFHIJKL")
        calc(ws, J, f'=IF(OR({F}="",{LUZ}=""),"",{F}-{LUZ})', b=False, fmt='+0;-0;0')
        calc(ws, K, f'=IF(OR({F}="",{KREF}="",{CREF}={k + 1}),"",{F}-{KREF})', b=False, fmt='+0;-0;0')
        calc(ws, L, f'=IF(AND({C}="",{E}=""),"",IF({E}="Automático","Automático",IF({F}="","Sem kelvin",'
                    f'IF({Hc}="Log","Perfil log",IF(AND({PREF}<>"",{Hc}<>"",{Hc}<>{PREF}),"Perfil diferente",'
                    f'IF(IF({J}="",FALSE,ABS({J})>{TL}),"Longe da luz",IF(IF({K}="",FALSE,ABS({K})>{TR}),"Diferente da referência",'
                    f'IF({I}<>"Sim","Conferir no monitor","OK"))))))))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="Automático","Passe para manual e faça o balanço com o cartão branco",'
                          f'IF({L}="Sem kelvin","Anote o kelvin que a câmera mostra",'
                          f'IF({L}="Perfil log","Ponha o perfil padrão, ou aplique a correção na mesa de corte",'
                          f'IF({L}="Perfil diferente","Ponha o mesmo perfil da câmera de referência e refaça o balanço",'
                          f'IF({L}="Longe da luz","Confira o modo e refaça o balanço com o cartão sob a luz do evento",'
                          f'IF({L}="Diferente da referência","Compare no monitor e ajuste alguns passos no kelvin ou no tint",'
                          f'IF({L}="Conferir no monitor","Compare lado a lado com a câmera de referência","Manter"))))))))',
             h="left", b=False, sz=9)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 24
    dv_list(ws, f"E{W1}:E{W2}", MODOS, "Cartão branco, Kelvin, Preset ou Automático.")
    dv_list(ws, f"H{W1}:H{W2}", PERFIS, "O perfil de imagem da câmera.")
    dv_list(ws, f"I{W1}:I{W2}", SIM_NAO, "Comparada lado a lado no mesmo monitor?")
    dv_number(ws, f"F{W1}:G{W2}")
    cf_equal(ws, f"L{W1}:L{W2}", WB_CF)

    s = W2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(L{W1}:L{W2},"{t}")'
    linhas = [("Câmeras registradas", f"=COUNTA(C{W1}:C{W2})"),
              ("Na situação OK", f'={cont("OK")}'),
              ("Em automático ou em log", f'={cont("Automático")}+{cont("Perfil log")}'),
              ("Maior diferença entre câmeras (K)", f'=IF(COUNT(F{W1}:F{W2})<2,"",MAX(F{W1}:F{W2})-MIN(F{W1}:F{W2}))')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:G{s + k}", fmt="0" if k == 4 else None)
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste as câmeras",IF({LUZ}="","Informe a luz principal",IF(F{s + 3}>0,'
                      f'"Há câmera em automático ou em log: corrija antes do evento",IF(F{s + 2}<F{s + 1},"Há câmera a ajustar ou a conferir","OK"))))',
         b=False, merge=f"F{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:M{a}")
    end = a
    if is_ex:
        end = ajustes(ws, end, data["ajustes"], ("B", ("C", "H"), ("I", "N")), ("Câm.", "Como estava", "O que foi feito"), "N")
    ws.freeze_panes = f"D{W1}"
    setup(ws, tab, f"B1:N{end}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Lentes, foco e cor — calculadora de enquadramento e ficha de cor"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Lentes, foco e cor — como usar esta planilha",
           "Calculadora de enquadramento, faixa em foco e ficha de balanço de branco, para planejar as câmeras e ajustá-las no evento.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Sensor, modo de balanço, perfil e as colunas Sim / Não aceitam só as opções da lista."),
        ("Números", "Digite só o número, sem unidade: 18 (metros), 140 (mm), 5.6 (abertura), 3200 (kelvin)."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Na visita ao local: meça a distância de cada posição de câmera até o lugar da pessoa."),
        ("2", "Aba Enquadramento: para cada câmera, o sensor, a distância, a largura de cena desejada e a faixa da lente."),
        ("3", "Veja a situação. Em Falta tele ou Falta angular, mude a posição (coluna Q), troque a lente ou aceite a cena obtida."),
        ("4", "Informe a abertura e veja a faixa em foco estimada. Se for estreita para quem anda, feche a íris."),
        ("5", "No evento, com a luz ligada: faça o balanço com cartão branco em cada câmera."),
        ("6", "Aba Balanço de branco: anote modo, kelvin, tint e perfil, compare no monitor e siga a coluna O que fazer."),
        ("7", "Aba Checklist: confira os doze itens antes de o evento começar."),
    ]),
    ("Abas da planilha", [
        ("Enquadramento", "Calculadora para até 12 câmeras: focal necessária, situação, cena obtida, distância que fecha a cena e faixa em foco."),
        ("Balanço de branco", "Ficha de cor para até 8 câmeras, com diferença para a luz e para a câmera de referência."),
        ("Referência", "Largura dos sensores (usada pela calculadora), temperaturas de cor, distâncias focais e modos de foco e de balanço."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Enquadramento", "As quatro câmeras de uma igreja, antes da mudança de posição (exemplo 1 do treinamento)."),
        ("Exemplo - Balanço", "As quatro câmeras de uma formatura com cores diferentes no corte (exemplo 2 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Conta da lente", "Largura da cena = distância × largura do sensor ÷ distância focal. Vale para assunto distante, como em evento. Diferenças de 5 a 10% para o quadro real são normais."),
        ("Equivalente full frame", "Focal necessária × 36 ÷ largura do sensor. Serve para comparar câmeras de sensores diferentes."),
        ("Situação", "OK quando a focal necessária está dentro da faixa da lente. Falta tele quando passa do máximo; Falta angular quando fica abaixo do mínimo. Zoom digital não conta."),
        ("Faixa em foco", "Estimativa usual, com círculo de confusão igual à largura do sensor ÷ 1200. Em tela grande ou em 4K, a faixa nítida é menor. Confira no local."),
        ("Balanço de branco", "Os cortes de 500 K para a luz e 200 K entre câmeras são regra prática deste material, e podem ser trocados no alto da aba."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_enq(wb.create_sheet("Enquadramento"), BLUE)
aba_wb(wb.create_sheet("Balanço de branco"), TEAL)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 16, "D": 16, "E": 16, "F": 56, "G": 2})
title(ws, "Referência — valores típicos",
      "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "F")

band(ws, 4, "Sensores (a calculadora de enquadramento usa esta tabela)", "F")
for col, t in zip("BCDEF", ["Sensor", "Largura (mm)", "Altura (mm)", "Fator de corte (pela largura)", "Onde costuma aparecer"]):
    put(ws, f"{col}5", t, f=font(10, True), bg=GRAY, h="center")
ws.row_dimensions[5].height = 21.75
for k, (nome, w, h, onde) in enumerate(SENSORES):
    r = S1 + k
    put(ws, f"B{r}", nome, f=font(10, True))
    if w is None:
        inp(ws, f"C{r}", None, h="center", fmt="0.0")
        inp(ws, f"D{r}", None, h="center", fmt="0.0")
        note(ws, f"C{r}", "Largura do sensor da sua câmera, em mm, na ficha técnica.\nNa aba Enquadramento, escolha Outro.")
    else:
        put(ws, f"C{r}", w, h="center", fmt="0.0")
        put(ws, f"D{r}", h, h="center", fmt="0.0")
    calc(ws, f"E{r}", f'=IF(C{r}="","",ROUND(36/C{r},1))', b=False, fmt="0.0")
    put(ws, f"F{r}", onde)
    ws.row_dimensions[r].height = 19.5
dv_number(ws, f"C{S2}:D{S2}")
hint(ws, f"B{S2 + 1}", "Medidas nominais. O fator publicado costuma ser calculado pela diagonal e pode diferir alguns por cento deste, pela largura, que é o que vale para o enquadramento. Algumas câmeras usam só parte do sensor em certos modos (4K, câmera lenta, estabilização eletrônica): o manual informa o recorte.",
     merge=f"B{S2 + 1}:F{S2 + 1}", h="left")
ws.row_dimensions[S2 + 1].height = 27


def tabela(r, titulo, cab, linhas, alt=31.5):
    band(ws, r, titulo, "F")
    r += 1
    put(ws, f"B{r}", cab[0], f=font(10, True), bg=GRAY, h="center")
    put(ws, f"C{r}", cab[1], f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:D{r}")
    put(ws, f"E{r}", cab[2], f=font(10, True), bg=GRAY, h="center")
    put(ws, f"F{r}", cab[3], f=font(10, True), bg=GRAY, h="center")
    ws.row_dimensions[r].height = 21.75
    for lin in linhas:
        r += 1
        put(ws, f"B{r}", lin[0], f=font(10, True))
        put(ws, f"C{r}", lin[1], h="center", merge=f"C{r}:D{r}")
        put(ws, f"E{r}", lin[2], h="center")
        put(ws, f"F{r}", lin[3])
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(S2 + 3, "Temperatura de cor das fontes", ["Fonte", "Temperatura de cor", "Preset mais próximo", "Observação"], [
    ("Vela", "cerca de 1.900 K", "–", "Muito alaranjada. Nenhum preset chega lá."),
    ("Lâmpada incandescente, LED branco quente", "2.700 a 3.000 K", "Tungstênio", "O valor do LED vem na embalagem."),
    ("Halógena, refletor de tungstênio", "3.000 a 3.200 K", "Tungstênio", "Fica mais alaranjada com o dimmer abaixo do máximo."),
    ("LED branco neutro, fluorescente comum", "4.000 a 4.500 K", "–", "Fluorescente e LED barato costumam puxar para o verde: use cartão branco ou tint."),
    ("Sol direto, flash", "5.000 a 5.600 K", "Luz do dia", "O sol fica mais quente no começo e no fim do dia."),
    ("LED luz do dia, branco frio", "5.600 a 6.500 K", "Luz do dia", "Refletores de LED de palco costumam ter a temperatura ajustável."),
    ("Céu nublado", "6.000 a 7.500 K", "Nublado", "Varia com a nuvem e a hora."),
    ("Telão de LED, monitor, projetor", "6.500 a 9.000 K", "–", "Com a câmera em 3.200 K, sai azulado. Corrija no telão, se possível."),
    ("Sombra sob céu azul", "7.500 a 10.000 K", "Sombra", "Muito azulada."),
])
r = tabela(r, "Distância focal e ângulo de visão (equivalente a full frame)", ["Focal equivalente", "Ângulo horizontal", "Cena a 10 m", "Nome usual e uso"], [
    ("16 mm", "97°", "22,5 m", "Grande-angular extrema. Câmera muito perto do palco; distorce as bordas."),
    ("24 mm", "74°", "15 m", "Grande-angular. Plano geral do palco e da plateia."),
    ("35 mm", "54°", "10,3 m", "Grande-angular moderada. Palco pequeno, banda inteira."),
    ("50 mm", "40°", "7,2 m", "Normal. Grupo de pessoas."),
    ("85 mm", "24°", "4,2 m", "Teleobjetiva curta. Pessoa de corpo inteiro a 10 m."),
    ("135 mm", "15°", "2,7 m", "Teleobjetiva. Pessoa dos joelhos para cima a 10 m."),
    ("200 mm", "10°", "1,8 m", "Teleobjetiva. Pessoa da cintura para cima a 10 m."),
    ("400 mm", "5°", "0,9 m", "Teleobjetiva longa. Cintura para cima a 20 m; pede tripé firme."),
], alt=19.5)
r = tabela(r, "Modo de foco por situação", ["Situação", "Modo indicado", "Abertura", "Por quê"], [
    ("Pessoa parada no púlpito, tripé", "Manual", "a da exposição", "A distância não muda. Foque com o zoom fechado no rosto."),
    ("Plano geral do palco", "Manual", "a da exposição", "A faixa em foco cobre o palco inteiro."),
    ("Palestrante que anda pelo palco", "AF contínuo, detecção de rosto, área limitada", "um ou dois pontos mais fechada", "A distância muda o tempo todo; a íris dá margem."),
    ("Câmera na mão, perto das pessoas", "AF contínuo, detecção de rosto", "a da exposição", "O operador já cuida do quadro e do equilíbrio."),
    ("Palco escuro, contraluz, fumaça", "Manual", "a da exposição", "O AF caça. Foque antes, com a luz acesa."),
    ("Câmera PTZ com posições gravadas", "Manual, gravado na posição", "a da exposição", "Em muitas PTZ, a posição gravada guarda o foco. Confira no manual."),
])
r = tabela(r, "Modos de balanço de branco", ["Modo", "Como funciona", "Em evento", "Risco"], [
    ("Automático (AWB, ATW)", "A câmera corrige o tempo todo", "Evite", "A cor muda com o telão, a roupa, a luz de cor, e cada câmera decide uma coisa."),
    ("Preset", "Valores fixos, perto de 3.200 e 5.600 K", "Só se a luz for a do preset", "Não corrige o verde de fluorescentes e de certos LEDs."),
    ("Kelvin", "Número digitado, de 100 em 100 K", "Bom, se a luz é conhecida", "Também não corrige o verde."),
    ("Manual com cartão branco", "A câmera mede um cartão sob a luz da cena", "O indicado", "Vale para a luz medida. Refaça se a luz mudar."),
])
hint(ws, f"B{r - 1}", "Faixas típicas. Os ângulos valem para assunto distante. Confira o manual do seu equipamento e teste no local.",
     merge=f"B{r - 1}:F{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:F{r - 1}", landscape=False, fit_height=False)

# ------------------------------------------------------------------ Checklist e exemplos
checklist(wb.create_sheet("Checklist"), "Checklist de lente, foco e cor",
          "Confira antes de o evento começar, com a luz do evento ligada. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=AMBER)
aba_enq(wb.create_sheet("Exemplo - Enquadramento"), MUTED, EX_ENQ)
aba_wb(wb.create_sheet("Exemplo - Balanço"), MUTED, EX_WB)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
