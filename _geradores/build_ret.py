# -*- coding: utf-8 -*-
"""Gera Monitores-e-Retornos-modelo.xlsx: matriz de retornos, pedidos por músico, plano de palco, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

TEAL, TEAL_T = HUE["teal"]
BLUE, BLUE_T = HUE["blue"]
SIM_NAO = ["Sim", "Não"]
TIPOS_RET = ["Caixa de chão", "Side fill", "Drum fill", "In-ear com fio", "In-ear sem fio", "Pessoal"]
TIPOS_FONTE = ["Voz", "Instrumento", "Bateria", "Linha ou DI", "Mídia", "Ambiente"]
MIX_COLS = "FGHIJKLM"   # oito mixes
NM = len(MIX_COLS)
MIXES = [f"Mix {k}" for k in range(1, NM + 1)]
N = 16                  # canais na matriz
NP = 24                 # linhas de pedidos

CHECK = [
    ("Planejamento", "A matriz de retornos está preenchida: quem ouve o quê, em qual mix e em qual saída."),
    ("Planejamento", "Os auxiliares de retorno estão em pré-fader."),
    ("Palco", "Cada saída foi testada: o mix sai na caixa ou no fone certo."),
    ("Palco", "Cada caixa de chão está no ponto surdo do microfone e apontada para os ouvidos do músico."),
    ("Palco", "O passa-altas e o ring-out foram feitos na saída de cada caixa de chão."),
    ("Mixagem", "Os ganhos dos canais foram ajustados antes de montar os retornos."),
    ("Mixagem", "Cada mix começou pelo próprio músico e tem só o que ele pediu."),
    ("Mixagem", "Com o PA fechado, o som do palco na plateia fica bem abaixo do som completo."),
    ("In-ear", "In-ear: limitador ligado, pilhas conferidas e os dois fones no ouvido."),
    ("In-ear", "In-ear: o mix tem microfone de ambiente, e o volume do receptor começou baixo."),
    ("Operação", "Os gestos de comunicação foram combinados, e o talkback foi testado."),
    ("Operação", "Os envios estão anotados ou a cena está salva, e não há pedido pendente."),
]

# Exemplo 1 do treinamento: os quatro mixes do culto, antes do ajuste
EX_MATRIZ = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala"},
    "lim": (6, 4),
    "mixes": [  # para quem, tipo, saída
        ("Cantora", "Caixa de chão", "Aux 1"),
        ("Vozes de apoio", "Caixa de chão", "Aux 2"),
        ("Tecladista e violonista", "Caixa de chão", "Aux 3"),
        ("Baterista", "In-ear com fio", "Aux 4"),
    ],
    "canais": [  # fonte, tipo, microfone aberto, níveis nos mixes 1 a 4
        ("Voz que conduz (cantora)", "Voz", "Sim", (3, 3, 3, 4)),
        ("Voz de apoio 1", "Voz", "Sim", (3, 5, None, None)),
        ("Voz de apoio 2", "Voz", "Sim", (3, 5, None, None)),
        ("Violão", "Linha ou DI", "Não", (3, None, 4, 3)),
        ("Teclado", "Linha ou DI", "Não", (3, 3, 5, 3)),
        ("Baixo", "Linha ou DI", "Não", (3, None, 2, 5)),
        ("Bumbo", "Bateria", "Sim", (3, None, None, None)),
        ("Caixa da bateria", "Bateria", "Sim", (3, None, None, None)),
        ("Pratos (microfone aéreo)", "Bateria", "Sim", (2, None, None, None)),
        ("Púlpito (pastor)", "Voz", "Sim", (None, None, None, None)),
        ("Computador (som dos vídeos)", "Mídia", "Não", (3, None, 2, None)),
    ],
    "ajustes": [  # mix, o que foi feito, resultado
        ("Mix 1", "Saíram bumbo, caixa, pratos, baixo e computador. Voz da cantora em 5, teclado em 3, violão e vozes de apoio em 2.",
         "Cinco fontes. Master do Aux 1 baixou 4 dB."),
        ("Mix 2", "Mantido.", "Quatro fontes."),
        ("Mix 3", "Mantido.", "Cinco fontes."),
        ("Mix 4", "Entraram o bumbo, em 3, e um par de microfones de ambiente, em 2, em um canal novo com o fader fechado para o PA.",
         "Seis fontes. Os dois fones no ouvido."),
        ("Púlpito", "Continua fora dos mixes: o pastor se ouve pela sala.", "Um microfone a menos captando caixa de chão."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def dv_int(ws, rng, lo, hi, msg):
    dv = DataValidation(type="whole", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True)
    dv.errorTitle, dv.error = "Valor inválido", msg
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def cf_fill(ws, rng, operator, formula, color):
    ws.conditional_formatting.add(rng, CellIsRule(operator=operator, formula=formula,
                                                  fill=PatternFill("solid", bgColor=color, fgColor=color)))


def fone(c):
    """Teste: o tipo de retorno da célula c é de fone."""
    return f'OR({c}="In-ear com fio",{c}="In-ear sem fio",{c}="Pessoal")'


# ================================================================== aba Matriz de retornos
M_W = {"A": 2, "B": 6, "C": 30, "D": 14, "E": 11, **{c: 13 for c in MIX_COLS}, "N": 11, "O": 30, "P": 2}
M1 = 16  # primeira linha de canal
SIT_MIX = [("OK", GREEN), ("Cheio", RED), ("Incompleto", YELLOW), ("Sem ambiente", YELLOW), ("Sem principal", YELLOW),
           ("Informe o tipo", YELLOW)]


def aba_matriz(ws, tab, data=None):
    widths(ws, M_W)
    is_ex = data is not None
    title(ws, "Matriz de retornos — quem ouve o quê",
          "Exemplo preenchido, para consulta. Use a aba Matriz de retornos para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro: os mixes nas colunas, os canais nas linhas e o nível pedido de 1 a 5.", "O")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:G{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    lim = data["lim"] if is_ex else (6, 4)
    label(ws, "I4", "Máximo de fontes em mix de caixa", merge="I4:K4")
    inp(ws, "L4", lim[0], h="center", bg=hb)
    hint(ws, "M4", "Acima disso, o mix fica Cheio. Regra prática.", merge="M4:O4", h="left")
    label(ws, "I5", "Mínimo de fontes em mix de fone", merge="I5:K5")
    inp(ws, "L5", lim[1], h="center", bg=hb)
    hint(ws, "M5", "Abaixo disso, o mix fica Incompleto. Regra prática.", merge="M5:O5", h="left")
    dv_int(ws, "L4:L5", 1, 32, "Digite um número inteiro de 1 a 32.")
    hint(ws, "I6", "Nível pedido: 5 = a fonte principal do mix; 3 = médio; 1 = só um pouco; vazio = não entra no mix.",
         merge="I6:O6", h="left")
    hint(ws, "I7", "A escala serve para planejar. Na mesa, o envio é ajustado de ouvido, com o músico.", merge="I7:O7", h="left")
    MAXC, MINF = "$L$4", "$L$5"

    band(ws, 9, "Mixes (colunas) e canais (linhas)", "O")
    for col, t in zip("BCDE", ["Canal", "Fonte", "Tipo de fonte", "Microfone aberto"]):
        head(ws, f"{col}10", t)
    for col, t in zip(MIX_COLS, MIXES):
        head(ws, f"{col}10", t, bg=TEAL)
    head(ws, "N10", "Em quantos mixes", bg=BLUE)
    head(ws, "O10", "Observação")
    ws.row_dimensions[10].height = 30
    mx = data["mixes"] if is_ex else []
    for r, rot, j in [(11, "Para quem (músico ou grupo)", 0), (12, "Tipo de retorno", 1), (13, "Saída da mesa", 2)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:E{r}", h="right")
        for k, col in enumerate(MIX_COLS):
            inp(ws, f"{col}{r}", mx[k][j] if k < len(mx) else None, h="center", bg=hb)
        put(ws, f"N{r}", None, bg=GRAY)
        put(ws, f"O{r}", None, bg=GRAY)
        ws.row_dimensions[r].height = 33
    dv_list(ws, f"{MIX_COLS[0]}12:{MIX_COLS[-1]}12", TIPOS_RET, "Caixa de chão, Side fill, Drum fill, In-ear com fio, In-ear sem fio ou Pessoal.")
    for col, t in zip("BCDE" + MIX_COLS + "NO", ["nº", "o que está ligado no canal", "lista", "Sim / Não"] + ["1 a 5"] * NM + ["conta", ""]):
        hint(ws, f"{col}14", t)
    ws.row_dimensions[14].height = 19.5
    if is_ex:
        hint(ws, "B15", "A matriz como estava antes do ajuste. As linhas de leitura, abaixo, são calculadas.", merge="B15:O15", h="left")
    else:
        vals = ["Ex.", "Voz que conduz (cantora)", "Voz", "Sim", 5, 3, 3, 4, "", "", "", "", 4, ""]
        for col, v in zip("BCDE" + MIX_COLS + "NO", vals):
            ex(ws, f"{col}15", v, h="left" if col in "CO" else "center")
        note(ws, "D10", "Voz, Instrumento (com microfone), Bateria, Linha ou DI, Mídia ou Ambiente.\nO tipo Ambiente é o que a planilha procura nos mixes de fone.")
        note(ws, "E10", "O canal é um microfone aberto no palco?\nMicrofone em caixa de chão é o que realimenta. Linha, DI e mídia não contam.")
        note(ws, "F10", "Em cada cruzamento, o nível pedido: de 1 (só um pouco) a 5 (a fonte principal do mix).\nDeixe vazio o que não entra no mix.")
        note(ws, "N10", "Em quantos mixes este canal entra.\nZero quer dizer que o canal não vai para nenhum retorno.")
    ws.row_dimensions[15].height = 21.75

    M2 = M1 + N - 1
    bg = WHITE if is_ex else INPUT
    for k in range(N):
        r = M1 + k
        c = data["canais"][k] if is_ex and k < len(data["canais"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, h="center", bg=bg)
        inp(ws, f"E{r}", c[2] if c else None, h="center", bg=bg)
        for j, col in enumerate(MIX_COLS):
            v = c[3][j] if c and j < len(c[3]) else None
            inp(ws, f"{col}{r}", v, h="center", bg=bg)
        calc(ws, f"N{r}", f'=IF(C{r}="","",COUNT({MIX_COLS[0]}{r}:{MIX_COLS[-1]}{r}))')
        inp(ws, f"O{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    rng = f"{MIX_COLS[0]}{M1}:{MIX_COLS[-1]}{M2}"
    dv_list(ws, f"D{M1}:D{M2}", TIPOS_FONTE, "Voz, Instrumento, Bateria, Linha ou DI, Mídia ou Ambiente.")
    dv_list(ws, f"E{M1}:E{M2}", SIM_NAO, "O canal é um microfone aberto no palco?")
    dv_int(ws, rng, 1, 5, "Digite um nível de 1 a 5, ou deixe vazio.")
    cf_fill(ws, rng, "greaterThanOrEqual", ["4"], "9FD3CC")
    cf_fill(ws, rng, "between", ["1", "3"], TEAL_T)

    s = M2 + 2
    band(ws, s, "Leitura de cada mix", "O")
    linhas = ["Fontes no mix", "Microfones abertos no mix", "Fontes no nível 4 ou 5", "Ambiente no mix", "Situação"]
    for k, t in enumerate(linhas, 1):
        label(ws, f"B{s + k}", t, merge=f"B{s + k}:E{s + k}", h="right")
        put(ws, f"N{s + k}", None, bg=GRAY)
        put(ws, f"O{s + k}", None, bg=GRAY)
        ws.row_dimensions[s + k].height = 21.75
    ws.row_dimensions[s + 5].height = 30
    rF, rM, rP, rA, rS = (s + k for k in range(1, 6))
    for col in MIX_COLS:
        cr = f"{col}{M1}:{col}{M2}"
        calc(ws, f"{col}{rF}", f'=IF(COUNT({cr})=0,"",COUNT({cr}))')
        calc(ws, f"{col}{rM}", f'=IF({col}{rF}="","",COUNTIFS($E${M1}:$E${M2},"Sim",{cr},">0"))', b=False)
        calc(ws, f"{col}{rP}", f'=IF({col}{rF}="","",COUNTIF({cr},">=4"))', b=False)
        calc(ws, f"{col}{rA}", f'=IF({col}{rF}="","",IF(COUNTIFS($D${M1}:$D${M2},"Ambiente",{cr},">0")>0,"Sim","Não"))', b=False)
        t = f"{col}12"
        calc(ws, f"{col}{rS}", f'=IF({col}{rF}="","",IF({t}="","Informe o tipo",IF({fone(t)},'
                               f'IF({col}{rF}<{MINF},"Incompleto",IF({col}{rA}="Não","Sem ambiente",IF({col}{rP}=0,"Sem principal","OK"))),'
                               f'IF({col}{rF}>{MAXC},"Cheio",IF({col}{rP}=0,"Sem principal","OK")))))')
    cf_equal(ws, f"{MIX_COLS[0]}{rS}:{MIX_COLS[-1]}{rS}", SIT_MIX)

    a = rS + 2
    band(ws, a, "O que fazer em cada mix", "O")
    put(ws, f"B{a + 1}", "Mix", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"C{a + 1}", "Para quem", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"D{a + 1}", "Situação", f=font(10, True), bg=GRAY, h="center", merge=f"D{a + 1}:E{a + 1}")
    put(ws, f"F{a + 1}", "O que fazer", f=font(10, True), bg=GRAY, h="center", merge=f"F{a + 1}:O{a + 1}")
    ws.row_dimensions[a + 1].height = 21.75
    for k, col in enumerate(MIX_COLS):
        r = a + 2 + k
        num(ws, f"B{r}", k + 1)
        calc(ws, f"C{r}", f'=IF({col}11="","",{col}11)', h="left", b=False)
        calc(ws, f"D{r}", f'=IF({col}{rS}="","",{col}{rS})', merge=f"D{r}:E{r}")
        calc(ws, f"F{r}", f'=IF(D{r}="","",IF(D{r}="Cheio","Fontes demais para uma caixa: tire o que o músico não pediu e o que ele já ouve no palco",'
                          f'IF(D{r}="Incompleto","O fone isola o músico: complete com harmonia, baixo, tempo e a voz que conduz",'
                          f'IF(D{r}="Sem ambiente","Inclua um canal do tipo Ambiente neste mix, em nível baixo",'
                          f'IF(D{r}="Sem principal","Nenhuma fonte no nível 4 ou 5: defina o que o músico mais precisa ouvir, em geral ele mesmo",'
                          f'IF(D{r}="Informe o tipo","Escolha o tipo de retorno na linha 12","Manter"))))))',
             h="left", b=False, merge=f"F{r}:O{r}")
        ws.row_dimensions[r].height = 21.75
    d1, d2 = a + 2, a + 1 + NM
    cf_equal(ws, f"D{d1}:E{d2}", SIT_MIX)

    t = d2 + 2
    band(ws, t, "Resumo automático", "O")
    summary(ws, t + 1, "Canais registrados", f"=COUNTA(C{M1}:C{M2})", "E", "F")
    summary(ws, t + 2, "Canais fora de todos os mixes", f'=COUNTIF(N{M1}:N{M2},0)', "E", "F")
    summary(ws, t + 3, "Mixes em uso", f"=COUNT({MIX_COLS[0]}{rF}:{MIX_COLS[-1]}{rF})", "E", "F")
    summary(ws, t + 4, "Mixes com situação OK", f'=COUNTIF({MIX_COLS[0]}{rS}:{MIX_COLS[-1]}{rS},"OK")', "E", "F")
    summary(ws, t + 5, "Mixes a ajustar", f"=F{t + 3}-F{t + 4}", "E", "F")
    label(ws, f"B{t + 6}", "Aviso", merge=f"B{t + 6}:E{t + 6}", h="right")
    calc(ws, f"F{t + 6}", f'=IF(F{t + 1}=0,"Liste os canais",IF(F{t + 3}=0,"Preencha os níveis de pelo menos um mix",'
                          f'IF(F{t + 5}>0,"Há mix a ajustar: veja o que fazer em cada um","OK")))',
         b=False, merge=f"F{t + 6}:O{t + 6}", h="left")
    ws.row_dimensions[t + 6].height = 21.75
    cf_ok(ws, f"F{t + 6}:O{t + 6}")
    end = t + 6

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "O")
        r = end + 3
        put(ws, f"B{r}", "Mix", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:K{r}")
        put(ws, f"L{r}", "Resultado", f=font(10, True), bg=GRAY, h="center", merge=f"L{r}:O{r}")
        ws.row_dimensions[r].height = 21.75
        for mix, feito, res in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", mix, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:K{r}")
            put(ws, f"L{r}", res, merge=f"L{r}:O{r}")
            ws.row_dimensions[r].height = 31.5
        end = r
    ws.freeze_panes = f"F{M1}"
    setup(ws, tab, f"B1:O{end}", fit_height=True)


# ================================================================== aba Pedidos por músico
def aba_pedidos(ws, tab):
    widths(ws, {"A": 2, "B": 6, "C": 13, "D": 24, "E": 10, "F": 11, "G": 26, "H": 13, "I": 62, "J": 12, "K": 12, "L": 30, "M": 2})
    title(ws, "Pedidos por músico — passagem e evento",
          "Anote cada pedido na hora. A planilha conta os pedidos de mais em cada mix e lembra o que tentar primeiro.", "L")
    for r, rot, fmt in [(4, "Evento", None), (5, "Data", DATE), (6, "Operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:G{r}", fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    hint(ws, "H4", "Gestos combinados: apontar para a fonte, depois polegar para cima (mais) ou para baixo (menos).", merge="H4:L4", h="left")
    hint(ws, "H5", "Um pedido por vez, um músico por vez. Ajuste de 2 a 3 dB e espere a confirmação.", merge="H5:L5", h="left")
    hint(ws, "H6", "Três pedidos de mais no mesmo mix são sinal de guerra de volume: baixe o que cobre. Regra prática.", merge="H6:L6", h="left")

    band(ws, 8, "Registro de pedidos", "L")
    heads = [("B", "#", INK), ("C", "Momento", INK), ("D", "Músico", INK), ("E", "Mix", TEAL), ("F", "Pedido", TEAL),
             ("G", "De quê (fonte)", TEAL), ("H", "Pedidos de mais no mix", BLUE), ("I", "O que tentar primeiro", BLUE),
             ("J", "Atendido", TEAL), ("K", "Mudança no envio (dB)", TEAL), ("L", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKL", ["nº", "lista", "quem pediu", "lista", "lista", "voz, teclado, bumbo…", "conta até esta linha",
                                      "", "Sim / Parcial / Não", "opcional", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 27
    for col, v in zip("BCDEFGHIJKL", ["Ex.", "Passagem", "Cantora", "Mix 1", "Mais", "A própria voz", 1,
                                      "Antes de subir: o que dá para baixar neste mix? A caixa aponta para os ouvidos?",
                                      "Sim", 0, "Baixou o teclado 3 dB; a voz não subiu"]):
        ex(ws, f"{col}11", v, h="left" if col in "DGIL" else "center", fmt='+0;-0;0' if col == "K" else None)
    ws.row_dimensions[11].height = 27
    note(ws, "F9", "Mais ou Menos: de uma fonte que já está no mix.\nIncluir ou Tirar: uma fonte que entra ou sai.\nTimbre: o som da caixa ou do fone (abafado, áspero).")
    note(ws, "H9", "Quantos pedidos de Mais este mix já recebeu, contando até esta linha.\nA partir do terceiro, a célula fica vermelha.")
    note(ws, "K9", "Quanto o envio mudou, em dB: +3, -2. Pode ficar em branco.\nZero quer dizer que o pedido foi resolvido sem subir nada.")

    P1, P2 = 12, 12 + NP - 1
    for k in range(NP):
        r = P1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center")
        inp(ws, f"D{r}")
        inp(ws, f"E{r}", h="center")
        inp(ws, f"F{r}", h="center")
        inp(ws, f"G{r}")
        calc(ws, f"H{r}", f'=IF(OR(E{r}="",F{r}<>"Mais"),"",COUNTIFS(E${P1}:E{r},E{r},F${P1}:F{r},"Mais"))')
        calc(ws, f"I{r}", f'=IF(F{r}="","",IF(F{r}="Mais",IF(H{r}="","Informe o mix",IF(H{r}>=3,'
                          f'"Terceiro pedido de mais neste mix, ou além: baixe as outras fontes e confira o volume de palco",'
                          f'"Antes de subir: o que dá para baixar neste mix? A caixa aponta para os ouvidos?")),'
                          f'IF(F{r}="Menos","Baixe o envio desta fonte só neste mix",'
                          f'IF(F{r}="Incluir","Abra o envio devagar, a partir do zero, até o músico confirmar",'
                          f'IF(F{r}="Tirar","Feche o envio desta fonte neste mix",'
                          f'IF(F{r}="Timbre","Ajuste o EQ da saída deste retorno; o EQ do canal muda também a sala",'
                          f'"Descreva o pedido na observação"))))))', h="left", b=False, sz=9)
        inp(ws, f"J{r}", h="center")
        inp(ws, f"K{r}", h="center", fmt='+0;-0;0')
        inp(ws, f"L{r}")
        ws.row_dimensions[r].height = 24
    dv_list(ws, f"C{P1}:C{P2}", ["Passagem", "Evento"], "Passagem ou Evento.")
    dv_list(ws, f"E{P1}:E{P2}", MIXES, "O mix do músico, como na aba Matriz de retornos.")
    dv_list(ws, f"F{P1}:F{P2}", ["Mais", "Menos", "Incluir", "Tirar", "Timbre", "Outro"], "Mais, Menos, Incluir, Tirar, Timbre ou Outro.")
    dv_list(ws, f"J{P1}:J{P2}", ["Sim", "Parcial", "Não"], "Sim, Parcial ou Não.")
    dv_number(ws, f"K{P1}:K{P2}")
    cf_equal(ws, f"J{P1}:J{P2}", [("Sim", GREEN), ("Parcial", YELLOW), ("Não", RED)])
    cf_fill(ws, f"H{P1}:H{P2}", "greaterThanOrEqual", ["3"], RED)

    s = P2 + 2
    band(ws, s, "Resumo automático", "L")
    summary(ws, s + 1, "Pedidos registrados", f"=COUNTA(F{P1}:F{P2})", "E", "F")
    summary(ws, s + 2, "Pedidos de mais", f'=COUNTIF(F{P1}:F{P2},"Mais")', "E", "F")
    summary(ws, s + 3, "Pedidos de menos ou de tirar", f'=COUNTIF(F{P1}:F{P2},"Menos")+COUNTIF(F{P1}:F{P2},"Tirar")', "E", "F")
    summary(ws, s + 4, "Pedidos atendidos", f'=COUNTIF(J{P1}:J{P2},"Sim")', "E", "F")
    summary(ws, s + 5, "Pedidos pendentes", f"=F{s + 1}-F{s + 4}", "E", "F")
    summary(ws, s + 6, "Maior número de pedidos de mais em um mix", f"=MAX(H{P1}:H{P2})", "E", "F")
    label(ws, f"B{s + 7}", "Aviso", merge=f"B{s + 7}:E{s + 7}", h="right")
    calc(ws, f"F{s + 7}", f'=IF(F{s + 1}=0,"Registre os pedidos",IF(F{s + 5}>0,"Há pedido pendente: resolva antes do evento",'
                          f'IF(F{s + 6}>=3,"Há mix com três ou mais pedidos de mais: reveja as fontes do mix e o volume de palco","OK")))',
         b=False, merge=f"F{s + 7}:L{s + 7}", h="left")
    ws.row_dimensions[s + 7].height = 21.75
    cf_ok(ws, f"F{s + 7}:L{s + 7}")
    ws.freeze_panes = f"C{P1}"
    setup(ws, tab, f"B1:L{s + 7}", fit_height=True)


# ================================================================== aba Plano de palco
SIT_PALCO = [("OK", GREEN), ("Posição", YELLOW), ("Ring-out", YELLOW), ("Falta dado", YELLOW), ("Omni", RED), ("Limitador", RED)]


def aba_palco(ws, tab):
    widths(ws, {"A": 2, "B": 6, "C": 26, "D": 12, "E": 16, "F": 17, "G": 12, "H": 13, "I": 12, "J": 12, "K": 11, "L": 11,
                "M": 13, "N": 54, "O": 28, "P": 2})
    title(ws, "Plano de palco dos retornos",
          "Uma linha por mix: onde fica cada caixa ou fone, e se a posição combina com o microfone do músico.", "O")
    for r, rot, fmt in [(4, "Evento", None), (5, "Data", DATE), (6, "Operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:F{r}", fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    hint(ws, "G4", "Ângulo da caixa: contado a partir da frente do microfone. 0 = para onde ele aponta; 180 = bem atrás dele, na direção do cabo.",
         merge="G4:O4", h="left")
    hint(ws, "G5", "Faixas aceitas: cardioide, de 150 a 180; supercardioide, de 105 a 145; hipercardioide, de 95 a 125. Regra prática em torno do ponto surdo teórico.",
         merge="G5:O5", h="left")
    hint(ws, "G6", "O ângulo só é conferido em caixa de chão. O limitador só é conferido em fone.", merge="G6:O6", h="left")

    band(ws, 8, "Retornos no palco", "O")
    heads = [("B", "Mix", INK), ("C", "Para quem", INK), ("D", "Saída da mesa", INK), ("E", "Tipo de retorno", TEAL),
             ("F", "Padrão polar do microfone", TEAL), ("G", "Ângulo da caixa", TEAL), ("H", "Ângulo de referência", BLUE),
             ("I", "Distância ao ouvido", TEAL), ("J", "Passa-altas da saída", TEAL), ("K", "Ring-out feito", TEAL),
             ("L", "Limitador ligado", TEAL), ("M", "Situação", BLUE), ("N", "O que fazer", BLUE), ("O", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKLMNO", ["nº", "músico ou grupo", "Aux 1, Aux 2…", "lista", "lista", "graus", "graus", "metros", "Hz",
                                         "Sim / Não", "Sim / Não", "", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    for col, v in zip("BCDEFGHIJKLMNO", ["Ex.", "Cantora", "Aux 1", "Caixa de chão", "Supercardioide", 180, "105 a 145", 1.5, 120, "Sim", "",
                                         "Posição", "Leve a caixa para a diagonal traseira do microfone, perto de 125 graus", ""]):
        ex(ws, f"{col}11", v, h="left" if col in "CNO" else "center")
    ws.row_dimensions[11].height = 27
    note(ws, "F9", "O padrão do microfone de voz do músico, como está no manual ou no corpo do microfone.\nSe o músico não usa microfone, escolha Sem microfone.")
    note(ws, "G9", "Ângulo entre a frente do microfone e a caixa, visto de cima.\n180 = a caixa está bem atrás do microfone.")
    note(ws, "J9", "Frequência do filtro passa-altas na saída deste retorno. Em retorno de voz, de 100 a 150 Hz, como ponto de partida.\nÉ só registro: não entra na situação.")
    note(ws, "M9", "Posição: a caixa está fora do ponto surdo do microfone.\nOmni: microfone omnidirecional com caixa de chão.\nRing-out: falta fazer nesta saída.\nLimitador: fone sem limitador ligado.")

    Q1, Q2 = 12, 12 + NM - 1
    for k in range(NM):
        r = Q1 + k
        E, F, G, K, L, M = (f"{c}{r}" for c in "EFGKLM")
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}")
        for col in "DEFGIJKL":
            inp(ws, f"{col}{r}", h="center", fmt="0.0" if col == "I" else None)
        cx = f'{E}="Caixa de chão"'
        calc(ws, f"H{r}", f'=IF(OR({E}<>"Caixa de chão",{F}=""),"",IF({F}="Cardioide","150 a 180",IF({F}="Supercardioide","105 a 145",'
                          f'IF({F}="Hipercardioide","95 a 125",IF({F}="Omni","não há","")))))', b=False)
        fora = (f'OR(AND({F}="Cardioide",{G}<150),AND({F}="Supercardioide",OR({G}<105,{G}>145)),'
                f'AND({F}="Hipercardioide",OR({G}<95,{G}>125)))')
        direc = f'OR({F}="Cardioide",{F}="Supercardioide",{F}="Hipercardioide")'
        calc(ws, M, f'=IF({E}="","",IF({fone(E)},IF({L}="Sim","OK",IF({L}="","Falta dado","Limitador")),'
                    f'IF(AND({cx},{F}="Omni"),"Omni",IF(AND({cx},{F}=""),"Falta dado",IF(AND({cx},{direc},{G}=""),"Falta dado",'
                    f'IF(AND({cx},{direc},{fora}),"Posição",IF({K}="Sim","OK",IF({K}="","Falta dado","Ring-out"))))))))')
        calc(ws, f"N{r}", f'=IF({M}="","",IF({M}="OK","Manter",IF({M}="Limitador","Ligue o limitador do receptor ou do amplificador de fones antes de o músico pôr os fones",'
                          f'IF({M}="Omni","Microfone omnidirecional capta a caixa de todos os lados: troque o microfone ou use fone",'
                          f'IF({M}="Posição",IF({F}="Cardioide","Leve a caixa para trás do microfone, perto de 180 graus",'
                          f'IF({F}="Supercardioide","Leve a caixa para a diagonal traseira do microfone, perto de 125 graus",'
                          f'"Leve a caixa para a diagonal do microfone, perto de 110 graus")),'
                          f'IF({M}="Ring-out","Faça o ring-out desta saída antes de montar a mistura",'
                          f'"Preencha o padrão polar, o ângulo e as colunas Sim / Não desta linha"))))))', h="left", b=False, sz=9)
        inp(ws, f"O{r}")
        ws.row_dimensions[r].height = 27
    dv_list(ws, f"E{Q1}:E{Q2}", TIPOS_RET, "Caixa de chão, Side fill, Drum fill, In-ear com fio, In-ear sem fio ou Pessoal.")
    dv_list(ws, f"F{Q1}:F{Q2}", ["Cardioide", "Supercardioide", "Hipercardioide", "Omni", "Sem microfone"],
            "Cardioide, Supercardioide, Hipercardioide, Omni ou Sem microfone.")
    dv_list(ws, f"K{Q1}:L{Q2}", SIM_NAO, "Sim ou Não.")
    dv_number(ws, f"G{Q1}:G{Q2}")
    dv_number(ws, f"I{Q1}:J{Q2}")
    cf_equal(ws, f"M{Q1}:M{Q2}", SIT_PALCO)

    s = Q2 + 2
    band(ws, s, "Resumo automático", "O")
    summary(ws, s + 1, "Retornos registrados", f"=COUNTA(E{Q1}:E{Q2})", "E", "F")
    summary(ws, s + 2, "Caixas no palco (chão, side fill e drum fill)",
            f'=COUNTIF(E{Q1}:E{Q2},"Caixa de chão")+COUNTIF(E{Q1}:E{Q2},"Side fill")+COUNTIF(E{Q1}:E{Q2},"Drum fill")', "E", "F")
    summary(ws, s + 3, "Retornos por fone", f"=F{s + 1}-F{s + 2}", "E", "F")
    summary(ws, s + 4, "Situação OK", f'=COUNTIF(M{Q1}:M{Q2},"OK")', "E", "F")
    summary(ws, s + 5, "A resolver", f"=F{s + 1}-F{s + 4}", "E", "F")
    label(ws, f"B{s + 6}", "Aviso", merge=f"B{s + 6}:E{s + 6}", h="right")
    calc(ws, f"F{s + 6}", f'=IF(F{s + 1}=0,"Registre os retornos",IF(F{s + 5}>0,"Há retorno a resolver: veja a coluna O que fazer","OK"))',
         b=False, merge=f"F{s + 6}:N{s + 6}", h="left")
    ws.row_dimensions[s + 6].height = 21.75
    cf_ok(ws, f"F{s + 6}:N{s + 6}")
    hint(ws, f"B{s + 7}", "Os ângulos de 180, 125 e 110 graus são os pontos surdos teóricos de cada padrão. O microfone real varia com a frequência: confirme no ring-out.",
         merge=f"B{s + 7}:O{s + 7}", h="left")
    setup(ws, tab, f"B1:O{s + 7}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Monitores e retornos — modelo de matriz de retornos"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Monitores e retornos — como usar esta planilha",
           "Matriz de retornos, registro de pedidos dos músicos e plano de palco, para usar na passagem de som e no evento.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo de fonte, tipo de retorno, padrão polar, pedido e as colunas Sim / Não aceitam só as opções da lista."),
        ("Nível pedido", "Na matriz, digite um número de 1 a 5. A célula fica verde-clara, e mais forte nos níveis 4 e 5."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Matriz de retornos: preencha o cabeçalho e, nas colunas, para quem é cada mix, o tipo de retorno e a saída da mesa."),
        ("2", "Aba Matriz de retornos: liste os canais, o tipo de fonte e se cada um é um microfone aberto no palco."),
        ("3", "Aba Matriz de retornos: em cada cruzamento, o nível pedido de 1 a 5. Veja a situação de cada mix e o que fazer."),
        ("4", "Aba Plano de palco: para cada mix, o padrão polar do microfone e o ângulo da caixa. Corrija a posição antes do ring-out."),
        ("5", "Na mesa: confira cada saída, faça o passa-altas e o ring-out, e monte um mix por vez, começando pelo próprio músico."),
        ("6", "Aba Pedidos por músico: anote cada pedido da passagem e do evento, e marque quando for atendido."),
        ("7", "Aba Checklist: confira os doze itens antes de abrir as portas."),
    ]),
    ("Abas da planilha", [
        ("Matriz de retornos", "Até 16 canais e 8 mixes, com a leitura de cada mix: OK, Cheio, Incompleto, Sem ambiente ou Sem principal."),
        ("Pedidos por músico", "Registro de até 24 pedidos, com a contagem de pedidos de mais por mix e o que tentar primeiro."),
        ("Plano de palco", "Uma linha por mix, com a conferência do ângulo da caixa pelo padrão polar, do ring-out e do limitador."),
        ("Referência", "Tipos de retorno, o que cada músico costuma pedir, posição da caixa, EQ de partida e gestos."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Matriz", "Os quatro mixes de um culto, antes do ajuste, e o que foi mudado (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Escala de 1 a 5", "É uma escala de planejamento deste material: 5 é a fonte principal do mix, 3 é médio, 1 é só um pouco. Não é uma medida em dB."),
        ("Cheio", "Mix de caixa (caixa de chão, side fill, drum fill) com mais fontes do que o máximo informado no alto da aba. O valor inicial é 6."),
        ("Incompleto", "Mix de fone (in-ear ou monitoração pessoal) com menos fontes do que o mínimo informado. O valor inicial é 4."),
        ("Sem ambiente", "Mix de fone sem nenhum canal do tipo Ambiente. O fone isola o músico da sala."),
        ("Sem principal", "Mix sem nenhuma fonte no nível 4 ou 5. Todo mix precisa de uma fonte à frente das outras."),
        ("Plano de palco", "O ângulo é conferido só em caixa de chão, com faixas em torno do ponto surdo teórico de cada padrão polar. São regra prática, e não norma."),
        ("Valores de referência", "Os limites e as faixas são ponto de partida. Confira o manual do seu equipamento e troque os valores do alto da aba se precisar."),
        ("Exemplos", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_matriz(wb.create_sheet("Matriz de retornos"), TEAL)
aba_pedidos(wb.create_sheet("Pedidos por músico"), BLUE)
aba_palco(wb.create_sheet("Plano de palco"), HUE["amber"][0])

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 36, "D": 36, "E": 56, "F": 2})
title(ws, "Referência — retornos", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")


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
            put(ws, f"{col}{r}", t)
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(4, "Tipos de retorno", ["Tipo", "Como é", "Ponto forte", "Cuidado"], [
    ("Caixa de chão (wedge)", "No chão, inclinada para o rosto do músico.", "Simples e natural.", "Soma volume no palco e é a origem mais comum de microfonia."),
    ("Side fill", "Caixas nas laterais do palco, apontadas para dentro.", "Cobre uma área grande: coral, dança.", "Muito volume de palco. Evite em sala pequena."),
    ("Drum fill", "Caixa ao lado do baterista, com mais graves.", "Vence o som da bateria acústica.", "É o retorno mais alto do palco. Fone costuma ser melhor."),
    ("In-ear", "Fones com vedação, com fio ou sem fio.", "Palco silencioso e mais margem antes da microfonia.", "Pede mix mais completo, ambiente, limitador e os dois fones."),
    ("Monitoração pessoal", "O músico faz a própria mistura em um misturador pequeno.", "Libera o operador para o som da frente.", "O músico precisa aprender a usar."),
])
r = tabela(r, "O que cada músico costuma pedir", ["Quem", "O principal do mix", "O que costuma completar", "Observação"], [
    ("Cantor que conduz", "A própria voz, bem à frente.", "Teclado ou violão. Um pouco das vozes de apoio.", "Em caixa de chão, sem reverb ou quase sem."),
    ("Vozes de apoio", "As vozes do grupo.", "A voz que conduz e a harmonia.", "Caixa dividida: o mix é um acordo."),
    ("Violonista ou guitarrista", "O próprio instrumento.", "A voz que conduz.", "Violão em caixa de chão alta realimenta nos graves."),
    ("Tecladista", "O teclado.", "A voz que conduz. Um pouco de baixo.", "Sem retorno, ele não ouve o que toca."),
    ("Baixista", "O baixo, se não houver amplificador.", "Bumbo e a voz que conduz.", "Caixa pequena não reproduz bem o grave."),
    ("Baterista", "Baixo e harmonia.", "A voz que conduz. O clique.", "Ele já ouve a bateria. Prefira fone."),
    ("Pregador ou palestrante", "Nada da própria voz, ou muito pouco.", "O áudio dos vídeos e as perguntas.", "Lapela em caixa de chão apita cedo."),
    ("Coral", "O teclado ou a base gravada.", "O solista.", "Não mande os microfones do coral para o retorno do coral."),
])
r = tabela(r, "Posição da caixa de chão pelo padrão polar", ["Padrão do microfone", "Onde capta menos", "Onde pôr a caixa", "Observação"], [
    ("Cardioide", "Bem atrás, a 180 graus.", "Uma caixa, atrás do microfone.", "O mais comum em microfone de mão."),
    ("Supercardioide", "Na diagonal traseira, a cerca de 125 graus.", "Duas caixas, ou uma deslocada para a diagonal.", "Bem atrás, a caixa cai no lóbulo traseiro."),
    ("Hipercardioide", "Na diagonal, a cerca de 110 graus.", "Caixas mais para os lados.", "O lóbulo traseiro é maior."),
    ("Omnidirecional", "Em nenhuma direção.", "Evite caixa de chão. Prefira fone.", "Comum em lapela e em alguns microfones de cabeça."),
])
r = tabela(r, "EQ da saída do retorno: ponto de partida", ["Ajuste", "Faixa", "Quando", "Observação"], [
    ("Passa-altas", "100 a 150 Hz em retorno de voz.", "Sempre, em caixa de chão de voz.", "Mais baixo ou desligado em drum fill e em retorno de baixo ou teclado."),
    ("Corte no embolado", "200 a 400 Hz, corte leve e largo.", "Quando a caixa soa embolada.", "Corte antes de realçar."),
    ("Corte no áspero", "2 a 4 kHz, corte leve.", "Quando a caixa soa áspera ou cansa.", "Realçar agudos aproxima o apito."),
    ("Cortes do ring-out", "As frequências que soaram primeiro.", "Na passagem, com os microfones abertos.", "Filtros estreitos, três ou quatro no máximo. Veja o estudo Microfonia."),
])
r = tabela(r, "Gestos combinados", ["Gesto", "Significa", "O operador", "Observação"], [
    ("Levantar a mão, olhando a mesa", "Tenho um pedido.", "Acena, para mostrar que viu.", "Combine os gestos antes da passagem."),
    ("Apontar para a fonte", "É sobre esta fonte.", "Localiza o canal.", "Para si mesmo, para o colega ou para o instrumento."),
    ("Polegar para cima ou para baixo", "Mais ou menos, no meu retorno.", "Ajusta de 2 a 3 dB e espera.", "Repetir o gesto: muito mais ou muito menos."),
    ("Mão aberta, parada", "Está bom.", "Para e anota.", ""),
    ("Apontar para o ouvido e negar", "Não ouço nada, ou algo está errado.", "Confere mute, cabo e pilha.", ""),
], alt=24)
hint(ws, f"B{r - 1}", "Faixas típicas e combinados comuns, que variam com o equipamento, a sala e a equipe. Não são norma.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist dos retornos",
          "Confira na passagem de som, antes de abrir as portas. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_matriz(wb.create_sheet("Exemplo - Matriz"), MUTED, EX_MATRIZ)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
