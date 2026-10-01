# -*- coding: utf-8 -*-
"""Gera Microfonia-modelo.xlsx: registro de ring-out, diagnóstico de ocorrências, margem por microfones abertos,
referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

RED_H, RED_T = HUE["red"]
SLATE_H, SLATE_T = HUE["slate"]
TEAL_H = HUE["teal"][0]
SIM_NAO = ["Sim", "Não"]
NUM1 = '+0.0;-0.0;0.0'
SGN = '+0;-0;0'

CHECK = [
    ("Posição", "As caixas da frente estão à frente da linha dos microfones e apontam para o público."),
    ("Posição", "As caixas apontam para as pessoas, e não para paredes, vidro ou teto."),
    ("Posição", "Cada retorno está na direção em que o microfone capta menos: atrás do cardioide, na diagonal do supercardioide."),
    ("Microfone", "Cada microfone está o mais perto possível da fonte, e quem usa foi orientado: perto da boca, grade livre."),
    ("Ganho", "O ganho de cada canal foi ajustado pelo medidor, e o filtro passa-altas está ligado nos microfones de voz."),
    ("Ganho", "O volume dos retornos é o menor que atende os músicos."),
    ("Ganho", "Só ficam abertos os microfones em uso. Os demais estão em mute, e há grupos de mute para os conjuntos."),
    ("Ring-out", "O ring-out foi feito em cada retorno e nas caixas da frente, com os microfones que ficam abertos juntos."),
    ("Ring-out", "Os cortes são estreitos e poucos: até uns cinco por saída, nenhum além de uns 9 dB."),
    ("Ring-out", "Há pelo menos 6 dB de margem entre o nível de operação e o ponto do primeiro apito, em cada saída."),
    ("Teste", "Os microfones que andam foram testados no caminho real, inclusive perto das caixas."),
    ("Registro", "As frequências, os cortes e os níveis de cada saída estão anotados."),
]

CAUSAS = [  # causa, medida definitiva
    ("Caixa mal posicionada", "Levar as caixas para a frente da linha dos microfones, apontadas para o público."),
    ("Microfone longe da fonte", "Aproximar o microfone da boca; trocar lapela por headset ou microfone de mão."),
    ("Mão na cápsula", "Orientar a segurar pelo corpo do microfone, com a grade livre."),
    ("Microfone virado para a caixa", "Mute nos intervalos; combinar que o microfone parado aponta para cima ou para o lado."),
    ("Microfones abertos sem uso", "Mute em tudo o que não está em uso; criar grupos de mute por momento do evento."),
    ("Retorno alto", "Baixar o retorno e combinar um volume de palco menor; rever a posição pelo padrão polar."),
    ("Volume da sala alto", "Operar com pelo menos 6 dB de margem; trazer o som para perto do público."),
    ("Realce de EQ ou compressor", "Trocar o realce por corte; reduzir a compensação (make-up) do compressor."),
    ("Pessoa na frente da caixa", "Combinar o caminho de quem anda com o microfone por fora da cobertura das caixas."),
    ("Sala reflexiva", "Apontar as caixas para as pessoas; aproximar o microfone; tratar a superfície mais próxima."),
    ("Não identificada", "Refazer o ring-out da saída com os microfones abertos no evento."),
]
assert len(",".join(c for c, _ in CAUSAS)) <= 255

# Exemplo 1 do treinamento: ring-out de um culto, três saídas, mesa digital (níveis = posição do master da saída)
EX_RING = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "mesa": "Mesa digital, EQ paramétrico nas saídas"},
    "par": (6, -9, 4, 5),
    "saidas": [  # saída, tipo, microfones abertos, apito antes, apito depois, operação
        ("Retorno 1 - voz principal", "Retorno", 1, -8, -2, -9),
        ("Retorno 2 - vozes de apoio", "Retorno", 3, -10, -5, -9),
        ("Frente (PA)", "Frente", 4, 3, 7, 0),
    ],
    "cortes": [  # saída, frequência, corte, Q, equalizador
        ("Retorno 1 - voz principal", 2500, -6, 8, "Paramétrico"),
        ("Retorno 1 - voz principal", 630, -4, 8, "Paramétrico"),
        ("Retorno 1 - voz principal", 4000, -5, 8, "Paramétrico"),
        ("Retorno 1 - voz principal", 250, -3, 8, "Paramétrico"),
        ("Retorno 2 - vozes de apoio", 1600, -5, 8, "Paramétrico"),
        ("Retorno 2 - vozes de apoio", 400, -10, 8, "Paramétrico"),
        ("Retorno 2 - vozes de apoio", 3150, -4, 8, "Paramétrico"),
        ("Frente (PA)", 3150, -4, 8, "Paramétrico"),
        ("Frente (PA)", 500, -3, 2, "Paramétrico"),
    ],
    "ajustes": [  # saída, o que foi feito, resultado
        ("Retorno 2", "O retorno ficava bem atrás de microfones supercardioides, no lóbulo traseiro. Foi levado para a diagonal, e o corte de 400 Hz voltou de −10 para −5 dB.",
         "Apito em −2 dB: margem de 7 dB"),
        ("Frente", "O corte de 500 Hz com Q 2 pegava mais de uma oitava. Foi refeito com Q 6.", "Corte dentro da faixa"),
        ("Todas", "Níveis de operação, cortes e microfones abertos anotados para o próximo culto.", ""),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def dv_range(ws, rng, source, prompt):
    dv = DataValidation(type="list", formula1=source, allow_blank=True)
    dv.promptTitle, dv.prompt = "Opções", prompt
    dv.errorTitle, dv.error = "Valor inválido", "Escolha uma das saídas da tabela de cima."
    dv.showErrorMessage = dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def cabecalho(ws, campos, hb, H, to="H"):
    for r, rot, k, fmt in campos:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:{to}{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75


# ================================================================== aba Ring-out
R_W = {"A": 2, "B": 6, "C": 30, "D": 13, "E": 11, "F": 11, "G": 11, "H": 13, "I": 13, "J": 11, "K": 11, "L": 15,
       "M": 50, "N": 26, "O": 2}
S1, NS = 14, 8          # saídas
C1, NC = 27, 30         # cortes


def aba_ring(ws, tab, data=None):
    widths(ws, R_W)
    is_ex = data is not None
    title(ws, "Ring-out — registro por saída e por corte",
          "Exemplo preenchido, para consulta. Use a aba Ring-out para o seu sistema." if is_ex
          else "Preencha as células em amarelo-claro durante o ring-out. As células cinza são calculadas.", "N")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    cabecalho(ws, [(4, "Evento ou sistema", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                   (7, "Operador", "operador", None), (8, "Mesa e equalizador", "mesa", None)], hb, H)
    dv_date(ws, "D6")
    par = data["par"] if is_ex else (6, -9, 4, 5)
    linhas = [("Margem mínima (dB)", "Regra prática: pelo menos 6 dB entre a operação e o primeiro apito."),
              ("Corte mais fundo aceito (dB)", "Corte além disso indica problema de posição, e não de equalização."),
              ("Q mínimo do corte", "Q perto de 4 = 1/3 de oitava. Corte de microfonia: Q de 5 a 10."),
              ("Máximo de cortes por saída", "Mais cortes que isso deixam o som oco.")]
    for k, (rot, dica) in enumerate(linhas):
        r = 4 + k
        label(ws, f"I{r}", rot, merge=f"I{r}:K{r}")
        inp(ws, f"L{r}", par[k], h="center", bg=hb, fmt=SGN if k == 1 else None)
        hint(ws, f"M{r}", dica, merge=f"M{r}:N{r}", h="left")
    dv_number(ws, "L4:L7")
    MIN_M, DEEP, QMIN, MAXC = "$L$4", "$L$5", "$L$6", "$L$7"
    hint(ws, "I8", "Níveis em dB, na posição do master de cada saída. O ponto de apito é onde a primeira nota começa a soar.",
         merge="I8:N8", h="left")

    S2, C2 = S1 + NS - 1, C1 + NC - 1
    SAI, CSAI, CSIT = f"$C${S1}:$C${S2}", f"$C${C1}:$C${C2}", f"$L${C1}:$L${C2}"

    # ---------------- saídas
    band(ws, 10, "Saídas: nível do primeiro apito e nível de operação", "N")
    for col, t, bg in [("B", "#", INK), ("C", "Saída", INK), ("D", "Tipo", INK), ("E", "Microfones abertos", INK),
                       ("F", "Apito antes dos cortes", RED_H), ("G", "Apito depois dos cortes", RED_H),
                       ("H", "Nível de operação", RED_H), ("I", "Cortes", SLATE_H), ("J", "Ganho obtido", SLATE_H),
                       ("K", "Margem", SLATE_H), ("L", "Situação", SLATE_H), ("M", "O que fazer", SLATE_H),
                       ("N", "Observação", INK)]:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["", "nome único", "", "no ring-out", "dB", "dB", "dB", "na tabela abaixo",
                                        "depois − antes", "apito − operação", "", "", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 21.75
    if is_ex:
        hint(ws, "B13", "Ring-out das três saídas do exemplo 1. As colunas de I a M são calculadas.", merge="B13:N13", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Retorno da cantora", "Retorno", 1, -8, -2, -9, 3, 6, 7, "OK", "Manter"]):
            ex(ws, f"{col}13", v, h="left" if col in "CM" else "center", fmt=SGN if col in "FGH" else None)
        ex(ws, "N13", "")
        note(ws, "C11", "Nome único de cada saída (retorno 1, frente...).\nA tabela de cortes usa esse nome para somar os cortes de cada saída.")
        note(ws, "F11", "Posição do master da saída, em dB, quando a primeira nota começou a soar, antes de qualquer corte.")
        note(ws, "G11", "O mesmo ponto, medido de novo depois dos cortes.")
        note(ws, "H11", "Posição do master em que a saída vai trabalhar no evento.")
        note(ws, "K11", "Distância entre o ponto do apito (depois dos cortes) e o nível de operação.\nRegra prática: pelo menos 6 dB.")
    ws.row_dimensions[13].height = 21.75

    for k in range(NS):
        r = S1 + k
        v = data["saidas"][k] if is_ex and k < len(data["saidas"]) else (None,) * 6
        bg = WHITE if is_ex else INPUT
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", v[0], bg=bg)
        inp(ws, f"D{r}", v[1], h="center", bg=bg)
        for col, val in zip("EFGH", v[2:]):
            inp(ws, f"{col}{r}", val, h="center", bg=bg, fmt=SGN if col != "E" else None)
        C, F, G, Hh, I, J, K, L = (f"{c}{r}" for c in "CFGHIJKL")
        calc(ws, I, f'=IF({C}="","",COUNTIF({CSAI},{C}))', b=False)
        calc(ws, J, f'=IF(OR({F}="",{G}=""),"",{G}-{F})', b=False, fmt=SGN)
        calc(ws, K, f'=IF(OR({G}="",{Hh}=""),"",{G}-{Hh})', fmt=SGN)
        rever = "+".join(f'COUNTIFS({CSAI},{C},{CSIT},"{s}")' for s in ("Corte fundo", "Corte largo", "Realce"))
        calc(ws, L, f'=IF(AND({C}="",{F}="",{G}="",{Hh}=""),"",IF(OR({G}="",{Hh}=""),"Incompleto",'
                    f'IF(AND({J}<>"",{J}<0),"Confira os níveis",IF({K}<3,"Sem margem",IF({K}<{MIN_M},"Margem curta",'
                    f'IF({I}>{MAXC},"Cortes demais",IF({rever}>0,"Rever cortes","OK")))))))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="Incompleto","Anote o apito depois dos cortes e o nível de operação",'
                          f'IF({L}="Confira os níveis","O apito depois dos cortes ficou abaixo do de antes: meça de novo",'
                          f'IF(OR({L}="Sem margem",{L}="Margem curta"),"Baixe a saída "&({MIN_M}-{K})&" dB, ou reveja a posição e a distância",'
                          f'IF({L}="Cortes demais","Cortes demais: reveja a posição das caixas e dos microfones",'
                          f'IF({L}="Rever cortes","Há corte fundo, largo ou realce nesta saída: veja a tabela de cortes","Manter"))))))',
             h="left", b=False, sz=9)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{S1}:D{S2}", ["Retorno", "Frente", "Lateral (side fill)", "Delay ou front fill", "Outra"], "Tipo de saída.")
    dv_number(ws, f"E{S1}:H{S2}")
    cf_equal(ws, f"L{S1}:L{S2}", [("OK", GREEN), ("Sem margem", RED), ("Margem curta", YELLOW), ("Cortes demais", YELLOW),
                                   ("Rever cortes", YELLOW), ("Incompleto", YELLOW), ("Confira os níveis", YELLOW)])

    # ---------------- cortes
    band(ws, C1 - 4, "Cortes de equalizador, na ordem em que foram feitos", "N")
    for col, t, bg in [("B", "#", INK), ("C", "Saída", INK), ("D", "Ordem na saída", SLATE_H), ("E", "Frequência", RED_H),
                       ("F", "Corte", RED_H), ("G", "Q", RED_H), ("H", "Equalizador", RED_H), ("I", "Faixa", SLATE_H),
                       ("J", "Largura", SLATE_H), ("K", "Largura", SLATE_H), ("L", "Situação", SLATE_H),
                       ("M", "O que fazer", SLATE_H), ("N", "Observação", INK)]:
        head(ws, f"{col}{C1 - 3}", t, bg=bg)
    ws.row_dimensions[C1 - 3].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["", "da lista", "", "Hz", "dB, negativo", "no paramétrico", "",
                                        "pelo som", "oitava", "Hz, aprox.", "", "", ""]):
        hint(ws, f"{col}{C1 - 2}", t)
    ws.row_dimensions[C1 - 2].height = 21.75
    if is_ex:
        hint(ws, f"B{C1 - 1}", "Os nove cortes do exemplo 1, antes do ajuste de posição do retorno 2. As colunas D, I, J, K, L e M são calculadas.",
             merge=f"B{C1 - 1}:N{C1 - 1}", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Retorno da cantora", 1, 2500, -6, 8, "Paramétrico", "Médio-agudo", 0.18, 313, "OK", "Manter"]):
            ex(ws, f"{col}{C1 - 1}", v, h="left" if col in "CM" else "center", fmt=SGN if col == "F" else None)
        ex(ws, f"N{C1 - 1}", "")
        note(ws, f"G{C1 - 3}", "Largura do filtro no equalizador paramétrico. Quanto maior, mais estreito.\nNo gráfico, deixe em branco: cada controle tem cerca de 1/3 de oitava.")
        note(ws, f"J{C1 - 3}", "Largura aproximada do corte, em oitavas, calculada pelo Q (no gráfico, 1/3 de oitava).")
    ws.row_dimensions[C1 - 1].height = 21.75

    for k in range(NC):
        r = C1 + k
        v = data["cortes"][k] if is_ex and k < len(data["cortes"]) else (None,) * 5
        bg = WHITE if is_ex else INPUT
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", v[0], bg=bg)
        C, D, E, F, G, Hh, I, J, K, L = (f"{c}{r}" for c in "CDEFGHIJKL")
        calc(ws, D, f'=IF(OR({C}="",{E}=""),"",COUNTIF($C${C1}:{C},{C}))', b=False)
        inp(ws, E, v[1], h="center", bg=bg)
        inp(ws, F, v[2], h="center", bg=bg, fmt=SGN)
        inp(ws, G, v[3], h="center", bg=bg)
        inp(ws, Hh, v[4], h="center", bg=bg)
        calc(ws, I, f'=IF({E}="","",IF({E}<250,"Grave",IF({E}<500,"Médio-grave",IF({E}<2000,"Médio",IF({E}<=5000,"Médio-agudo","Agudo")))))', b=False)
        sem_q = f'OR({E}="",AND({Hh}<>"Gráfico",OR({G}="",{G}<=0)))'
        calc(ws, J, f'=IF({sem_q},"",IF({Hh}="Gráfico",0.33,ROUND(2*LOG10(1/(2*{G})+(1/(4*{G}^2)+1)^0.5)/LOG10(2),2)))',
             b=False, fmt="0.00")
        calc(ws, K, f'=IF({sem_q},"",ROUND({E}*IF({Hh}="Gráfico",0.232,1/{G}),0))', b=False)
        calc(ws, L, f'=IF({E}="","",IF(COUNTIF({SAI},{C})=0,"Saída?",IF(AND({F}<>"",{F}>0),"Realce",IF(AND({F}<>"",{F}<{DEEP}),"Corte fundo",'
                    f'IF(AND({Hh}<>"Gráfico",{G}<>"",{G}<{QMIN}),"Corte largo","OK")))))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="Saída?","Escreva a saída igual à da tabela de cima",'
                          f'IF({L}="Realce","Realce aumenta o ganho do ciclo: troque por corte",'
                          f'IF({L}="Corte fundo","Corte além de "&ABS({DEEP})&" dB: reveja a posição antes de aprofundar",'
                          f'IF({L}="Corte largo","Estreite o corte: Q de pelo menos "&{QMIN},"Manter")))))', h="left", b=False, sz=9)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 19.5
    dv_range(ws, f"C{C1}:C{C2}", SAI, "Escolha uma saída já listada na tabela de cima.")
    dv_list(ws, f"H{C1}:H{C2}", ["Paramétrico", "Gráfico"], "Paramétrico ou Gráfico.")
    dv_number(ws, f"E{C1}:G{C2}")
    cf_equal(ws, f"L{C1}:L{C2}", [("OK", GREEN), ("Realce", RED), ("Corte fundo", YELLOW), ("Corte largo", YELLOW), ("Saída?", YELLOW)])

    # ---------------- resumo
    s = C2 + 2
    band(ws, s, "Resumo automático", "N")
    cs = lambda t: f'COUNTIF($L${S1}:$L${S2},"{t}")'
    cc = lambda t: f'COUNTIF({CSIT},"{t}")'
    linhas = [
        ("Saídas registradas", f"=COUNTA(C{S1}:C{S2})", None),
        ("Saídas OK", f"={cs('OK')}", None),
        ("Saídas sem margem ou com margem curta", f"={cs('Sem margem')}+{cs('Margem curta')}", None),
        ("Menor margem (dB)", f'=IF(COUNT(K{S1}:K{S2})=0,"",MIN(K{S1}:K{S2}))', SGN),
        ("Cortes registrados", f"=COUNT(E{C1}:E{C2})", None),
        ("Cortes a rever", f"={cc('Corte fundo')}+{cc('Corte largo')}+{cc('Realce')}+{cc('Saída?')}", None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", fmt=fmt, val_merge=f"F{s + k}:G{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(COUNT(L4:L7)<4,"Informe os quatro parâmetros",IF(F{s + 1}=0,"Liste as saídas",'
                      f'IF(F{s + 3}>0,"Há saída sem a margem mínima: baixe o nível ou reveja a posição",'
                      f'IF(F{s + 2}<F{s + 1},"Há saída a rever: veja a coluna O que fazer",'
                      f'IF(F{s + 6}>0,"Há corte a rever","OK")))))', b=False, merge=f"F{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:M{a}")
    hint(ws, f"B{a + 1}", "Margem mínima, corte mais fundo, Q mínimo e número de cortes são regras práticas deste material, e não norma. Troque no alto da aba.",
         merge=f"B{a + 1}:N{a + 1}", h="left")
    end = a + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado depois do ring-out", "N")
        r = end + 3
        put(ws, f"B{r}", "Saída", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:L{r}")
        put(ws, f"M{r}", "Resultado", f=font(10, True), bg=GRAY, h="center", merge=f"M{r}:N{r}")
        ws.row_dimensions[r].height = 21.75
        for sai, feito, res in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", sai, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:L{r}")
            put(ws, f"M{r}", res, merge=f"M{r}:N{r}")
            ws.row_dimensions[r].height = 31.5
        end = r
    setup(ws, tab, f"B1:N{end}", fit_height=False)


# ================================================================== aba Ocorrências
O_W = {"A": 2, "B": 5, "C": 12, "D": 22, "E": 13, "F": 20, "G": 14, "H": 11, "I": 28, "J": 32, "K": 11, "L": 14,
       "M": 50, "N": 13, "O": 24, "P": 2}
O1, NO = 9, 20


def aba_ocorrencias(ws, tab, lookup):
    widths(ws, O_W)
    title(ws, "Ocorrências — diagnóstico de cada apito durante os eventos",
          "Anote cada apito logo depois do evento: o que aconteceu, a causa e o que foi feito na hora. A medida definitiva é sugerida.", "O")
    band(ws, 4, "Registro de ocorrências", "O")
    for col, t, bg in [("B", "#", INK), ("C", "Data", INK), ("D", "Momento do evento", INK), ("E", "Saída", RED_H),
                       ("F", "Canal", RED_H), ("G", "Som do apito", RED_H), ("H", "Frequência", RED_H), ("I", "Causa", RED_H),
                       ("J", "Medida na hora", RED_H), ("K", "Resolvido?", RED_H), ("L", "Faixa pela frequência", SLATE_H),
                       ("M", "Medida definitiva sugerida", SLATE_H), ("N", "Situação", SLATE_H), ("O", "Observação", INK)]:
        head(ws, f"{col}5", t, bg=bg)
    ws.row_dimensions[5].height = 30
    for col, t in zip("BCDEFGHIJKLMNO", ["", "dd/mm/aaaa", "", "", "", "pelo ouvido", "Hz, se medida", "da lista", "",
                                         "Sim / Não", "", "", "", ""]):
        hint(ws, f"{col}6", t)
    ws.row_dimensions[6].height = 21.75
    exl = ["Ex.", "", "Pregação", "Frente", "Pastor (púlpito)", "Médio-agudo", 2500, "Microfones abertos sem uso",
           "Fader do pastor −3 dB; mute nos microfones da banda", "Sim", "Médio-agudo", CAUSAS[4][1], "Resolvida"]
    for col, v in zip("BCDEFGHIJKLMN", exl):
        ex(ws, f"{col}7", v, h="left" if col in "DFIJM" else "center")
    ex(ws, "O7", "")
    ws.row_dimensions[7].height = 31.5
    hint(ws, "B8", "A linha Ex. reproduz o exemplo 3 do treinamento.", merge="B8:O8", h="left")
    note(ws, "G5", "Grave: ronco. Médio-grave: uivo. Médio: buzina, som de cano. Médio-agudo: apito claro. Agudo: silvo fino.")
    note(ws, "I5", "A causa que você encontrou. Se não deu para saber, escolha Não identificada.")
    O2 = O1 + NO - 1
    for k in range(NO):
        r = O1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center", fmt=DATE)
        for col in "DEFGHIJK":
            inp(ws, f"{col}{r}", h="center" if col in "EGHK" else "left")
        C, D, F, H, I, K = (f"{c}{r}" for c in "CDFHIK")
        calc(ws, f"L{r}", f'=IF({H}="","",IF({H}<250,"Grave",IF({H}<500,"Médio-grave",IF({H}<2000,"Médio",IF({H}<=5000,"Médio-agudo","Agudo")))))', b=False)
        calc(ws, f"M{r}", f'=IF({I}="","",IFERROR(VLOOKUP({I},{lookup},2,FALSE),""))', h="left", b=False, sz=9)
        calc(ws, f"N{r}", f'=IF(AND({C}="",{D}="",{F}="",{I}=""),"",IF({K}="Sim","Resolvida","Aberta"))')
        inp(ws, f"O{r}")
        ws.row_dimensions[r].height = 31.5
    dv_date(ws, f"C{O1}:C{O2}")
    dv_list(ws, f"E{O1}:E{O2}", ["Frente", "Retorno", "Lateral", "Delay", "Outra"], "Em que saída apitou?")
    dv_list(ws, f"G{O1}:G{O2}", ["Grave", "Médio-grave", "Médio", "Médio-agudo", "Agudo"], "Como soava o apito.")
    dv_list(ws, f"I{O1}:I{O2}", [c for c, _ in CAUSAS], "A causa encontrada.")
    dv_list(ws, f"K{O1}:K{O2}", SIM_NAO, "A medida na hora resolveu?")
    dv_number(ws, f"H{O1}:H{O2}")
    cf_equal(ws, f"N{O1}:N{O2}", [("Resolvida", GREEN), ("Aberta", YELLOW)])

    s = O2 + 2
    band(ws, s, "Ocorrências por causa", "O")
    r = s
    for causa, _ in CAUSAS:
        r += 1
        label(ws, f"B{r}", None, merge=f"B{r}:C{r}")
        put(ws, f"D{r}", causa, f=font(10, True), bg=GRAY, merge=f"D{r}:F{r}")
        calc(ws, f"G{r}", f'=COUNTIF($I${O1}:$I${O2},D{r})')
        ws.row_dimensions[r].height = 19.5
    c1, c2 = s + 1, r
    summary(ws, r + 2, "Ocorrências registradas", f'=COUNTIF($N${O1}:$N${O2},"Resolvida")+COUNTIF($N${O1}:$N${O2},"Aberta")', "F", "G")
    summary(ws, r + 3, "Abertas", f'=COUNTIF($N${O1}:$N${O2},"Aberta")', "F", "G")
    label(ws, f"B{r + 4}", "Causa mais frequente", merge=f"B{r + 4}:F{r + 4}", h="right")
    calc(ws, f"G{r + 4}", f'=IF(MAX(G{c1}:G{c2})=0,"",INDEX(D{c1}:D{c2},MATCH(MAX(G{c1}:G{c2}),G{c1}:G{c2},0)))',
         merge=f"G{r + 4}:J{r + 4}", h="left")
    label(ws, f"B{r + 5}", "Aviso", merge=f"B{r + 5}:F{r + 5}", h="right")
    calc(ws, f"G{r + 5}", f'=IF(G{r + 2}=0,"Sem ocorrências registradas",IF(G{r + 3}>0,"Há ocorrência aberta: aplique a medida definitiva antes do próximo evento","OK"))',
         b=False, merge=f"G{r + 5}:M{r + 5}", h="left")
    ws.row_dimensions[r + 4].height = ws.row_dimensions[r + 5].height = 21.75
    cf_ok(ws, f"G{r + 5}:M{r + 5}")
    ws.freeze_panes = f"C{O1}"
    setup(ws, tab, f"B1:O{r + 5}")


# ================================================================== aba Margem
M_W = {"A": 2, "B": 5, "C": 36, "D": 14, "E": 15, "F": 15, "G": 15, "H": 58, "I": 2}
M1, NM = 14, 10


def aba_margem(ws, tab):
    widths(ws, M_W)
    title(ws, "Margem — microfones abertos e distância",
          "Estimativa da margem antes da microfonia em cada momento do evento, e calculadora de distância.", "H")
    band(ws, 4, "Ponto de partida: o ring-out", "H")
    for r, rot, dica, v in [(5, "Margem medida no ring-out (dB)", "Da aba Ring-out: apito depois dos cortes menos o nível de operação.", None),
                            (6, "Microfones abertos no ring-out", "Quantos microfones estavam abertos quando a margem foi medida.", None),
                            (7, "Margem mínima (dB)", "Regra prática: 6 dB.", 6)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", v, h="center")
        hint(ws, f"E{r}", dica, merge=f"E{r}:H{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, "D5:D7")
    label(ws, "B8", "Máximo de microfones abertos com a margem mínima", merge="B8:C8")
    calc(ws, "D8", '=IF(OR(D5="",D6="",D7="",D6<1),"",MAX(0,ROUND(D6*10^((D5-D7)/10)-0.4999,0)))')
    hint(ws, "E8", "Regra: cada vez que o número de microfones abertos dobra, a margem cai cerca de 3 dB.", merge="E8:H8", h="left")
    ws.row_dimensions[8].height = 21.75

    band(ws, 10, "Margem em cada momento do evento", "H")
    for col, t, bg in [("B", "#", INK), ("C", "Momento do evento", INK), ("D", "Microfones abertos", RED_H),
                       ("E", "Mudança da margem", SLATE_H), ("F", "Margem estimada", SLATE_H), ("G", "Situação", SLATE_H),
                       ("H", "O que fazer", SLATE_H)]:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGH", ["", "", "ao mesmo tempo", "dB", "dB", "", ""]):
        hint(ws, f"{col}12", t)
    for col, v in zip("BCDEFGH", ["Ex.", "Pregação com a banda aberta (ring-out: 7 dB com 2)", 8, -6.0, 1.0, "Sem margem",
                                  "Feche microfones sem uso: com a margem mínima, até 2 abertos"]):
        ex(ws, f"{col}13", v, h="left" if col in "CH" else "center", fmt=NUM1 if col in "EF" else None)
    ws.row_dimensions[13].height = 31.5
    M2 = M1 + NM - 1
    for k in range(NM):
        r = M1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}")
        inp(ws, f"D{r}", h="center")
        D, E, F, G = (f"{c}{r}" for c in "DEFG")
        calc(ws, E, f'=IF(OR({D}="",$D$6="",{D}<1,$D$6<1),"",ROUND(10*LOG10($D$6/{D}),1))', b=False, fmt=NUM1)
        calc(ws, F, f'=IF(OR({E}="",$D$5=""),"",$D$5+{E})', fmt=NUM1)
        calc(ws, G, f'=IF({F}="","",IF({F}<3,"Sem margem",IF({F}<$D$7,"Margem curta","OK")))')
        calc(ws, f"H{r}", f'=IF({G}="","",IF({G}="OK","Manter","Feche microfones sem uso: com a margem mínima, até "&$D$8&" abertos"))',
             h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"D{M1}:D{M2}")
    cf_equal(ws, f"G{M1}:G{M2}", [("OK", GREEN), ("Margem curta", YELLOW), ("Sem margem", RED)])

    r = M2 + 2
    band(ws, r, "Calculadora de distância", "H")
    blocos = [
        ("Fonte–microfone", "Distância atual da boca ao microfone (cm)", "Nova distância (cm)", "Mudança na voz captada (dB)",
         "D{a}/D{b}", "Positivo: a voz chega mais forte, e o sistema precisa amplificar menos. Essa diferença vira margem."),
        ("Caixa–microfone", "Distância atual da caixa ao microfone (m)", "Nova distância (m)", "Ganho de margem, no máximo (dB)",
         "D{b}/D{a}", "Valor ao ar livre. Em sala reflexiva, o som volta pelas paredes, e o ganho real é menor."),
    ]
    for nome, l1, l2, l3, razao, dica in blocos:
        r += 1
        put(ws, f"B{r}", nome, f=font(10, True, c=WHITE), bg=SLATE_H, merge=f"B{r}:D{r}")
        ws.row_dimensions[r].height = 19.5
        a, b, c = r + 1, r + 2, r + 3
        label(ws, f"B{a}", l1, merge=f"B{a}:C{a}")
        inp(ws, f"D{a}", h="center")
        label(ws, f"B{b}", l2, merge=f"B{b}:C{b}")
        inp(ws, f"D{b}", h="center")
        label(ws, f"B{c}", l3, merge=f"B{c}:C{c}")
        calc(ws, f"D{c}", f'=IF(OR(D{a}="",D{b}="",D{a}<=0,D{b}<=0),"",ROUND(20*LOG10({razao.format(a=a, b=b)}),1))', fmt=NUM1)
        hint(ws, f"E{a}", dica, merge=f"E{a}:H{c}", h="left")
        for x in (a, b, c):
            ws.row_dimensions[x].height = 19.5
        dv_number(ws, f"D{a}:D{b}")
        r = c
    hint(ws, f"B{r + 2}", "Estimativas pela lei do inverso do quadrado e pela regra dos 3 dB por dobro de microfones abertos. A medida real vem do ring-out.",
         merge=f"B{r + 2}:H{r + 2}", h="left")
    setup(ws, tab, f"B1:H{r + 2}", fit_height=True)


# ================================================================== aba Referência
def aba_referencia(ws):
    widths(ws, {"A": 2, "B": 30, "C": 22, "D": 30, "E": 60, "F": 2})
    title(ws, "Referência — microfonia", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")
    state = {"r": 4}

    def tabela(titulo, cab, linhas, alt=31.5):
        r = state["r"]
        band(ws, r, titulo, "E")
        r += 1
        for col, t in zip("BCDE", cab):
            put(ws, f"{col}{r}", t, f=font(10, True), bg=GRAY, h="center")
        ws.row_dimensions[r].height = 21.75
        first = r + 1
        for lin in linhas:
            r += 1
            put(ws, f"B{r}", lin[0], f=font(10, True))
            for col, t in zip("CDE", lin[1:]):
                put(ws, f"{col}{r}", t, h="left" if col == "E" else "center")
            ws.row_dimensions[r].height = alt
        state["r"] = r + 2
        return first, r

    tabela("Faixas do apito", ["Faixa", "Frequência", "Como soa", "Onde costuma aparecer"], [
        ("Grave", "até 250 Hz", "ronco", "Lapela, microfone de violão ou de bumbo, microfone perto do subwoofer. Ligue o filtro passa-altas."),
        ("Médio-grave", "250 a 500 Hz", "uivo, embolado", "Púlpito e coral captados de longe, sala reverberante, caixa atrás de quem fala."),
        ("Médio", "500 Hz a 2 kHz", "buzina, som de cano", "Retorno de chão alto, mão em volta da cápsula."),
        ("Médio-agudo", "2 a 5 kHz", "apito claro", "Microfone de voz com retorno de chão, microfone apontado para a caixa."),
        ("Agudo", "acima de 5 kHz", "silvo fino", "Condensador, lapela ou headset que passa perto de uma caixa."),
    ])
    tabela("Q e largura do corte (equalizador paramétrico)", ["Q", "Largura", "Uso", "Observação"], [
        ("1,4", "1 oitava", "largo demais", "Tira o apito e leva junto um pedaço grande do som."),
        ("2,9", "1/2 oitava", "largo", "Ainda largo para microfonia."),
        ("4,3", "1/3 de oitava", "como o EQ gráfico", "Largura de cada controle do equalizador gráfico de 31 bandas."),
        ("8,7", "1/6 de oitava", "bom ponto de partida", "Faixa usual para cortes de microfonia: Q de 5 a 10."),
        ("14,4", "1/10 de oitava", "muito estreito", "Preciso, mas pede a frequência exata."),
    ], alt=21.75)
    tabela("Microfones abertos e margem", ["Microfones abertos", "Mudança da margem", "Regra", "Observação"], [
        ("1", "0 dB", "referência", "Supõe microfones parecidos, no mesmo nível, captando o sistema de forma parecida."),
        ("2", "−3 dB", "dobrou", "Feche os microfones que não estão em uso."),
        ("4", "−6 dB", "dobrou duas vezes", "Use grupos de mute ou DCA por momento do evento."),
        ("8", "−9 dB", "três vezes", "Mixagem automática (automixer) ajuda em debates com muitos microfones."),
        ("16", "−12 dB", "quatro vezes", "Faça o ring-out com os microfones que ficam abertos juntos."),
    ], alt=21.75)
    tabela("Medidas e efeito na margem", ["Medida", "Efeito", "Ordem", "Do que depende"], [
        ("Posição das caixas", "vários dB", "1", "Caixas à frente da linha dos microfones; retorno na direção de menor captação."),
        ("Microfone na metade da distância", "+6 dB", "2", "Vale a cada metade. De 20 cm para 5 cm, 12 dB."),
        ("Caixa ao dobro da distância", "até +6 dB", "1", "O valor cheio só ao ar livre. Em sala reflexiva, menos."),
        ("Metade dos microfones abertos", "+3 dB", "3", "De os microfones captarem o sistema de forma parecida."),
        ("Ring-out", "3 a 6 dB", "4", "Faixa típica. Se os primeiros cortes rendem pouco, a posição está ruim."),
    ])
    tabela("Posição do retorno pelo padrão polar", ["Padrão polar", "Menor captação", "Onde fica o retorno", "Observação"], [
        ("Cardioide", "atrás, a 180°", "bem à frente do músico", "Chega pela parte de trás do microfone."),
        ("Supercardioide", "cerca de 125°", "na diagonal, ou dois retornos", "Bem atrás ainda há um pequeno lóbulo de captação."),
        ("Hipercardioide", "cerca de 110°", "mais para o lado", "Lóbulo traseiro maior que o do supercardioide."),
        ("Onidirecional", "nenhuma", "não há posição boa", "A margem vem só da distância e do volume."),
    ])
    first, last = tabela("Causa e medida definitiva (usada na aba Ocorrências)", ["Causa", "Medida definitiva", "", ""],
                         [(c, m, "", "") for c, m in CAUSAS], alt=21.75)
    for r in range(first - 1, last + 1):
        ws.merge_cells(f"C{r}:E{r}")
        ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    r = state["r"]
    hint(ws, f"B{r - 1}", "Faixas típicas, que variam com o microfone, a caixa, a sala e o equipamento. Os ângulos são valores de projeto: veja o diagrama polar no manual do microfone.",
         merge=f"B{r - 1}:E{r - 1}", h="left")
    setup(ws, SLATE_H, f"B1:E{r - 1}", landscape=False)
    return f"'Referência'!$B${first}:$C${last}"


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Microfonia — modelo de ring-out e de ocorrências"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Microfonia — como usar esta planilha",
           "Registro do ring-out, diagnóstico dos apitos e estimativa da margem, para usar na montagem e depois do evento.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo de saída, equalizador, saída do corte, som do apito, causa e Sim / Não aceitam só as opções da lista."),
        ("Números", "Digite só o número, com o sinal de menos quando houver: -6, e não “-6 dB”. Frequência em Hz: 2500, e não 2,5k."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Antes do ring-out: posição das caixas, microfones perto da fonte e ganho pelo medidor (módulo 4 do treinamento)."),
        ("2", "Aba Ring-out: preencha o cabeçalho, confira os quatro parâmetros e liste as saídas, com nomes únicos."),
        ("3", "Para cada saída: anote o nível do primeiro apito antes dos cortes. Registre cada corte na tabela de baixo."),
        ("4", "Anote o apito depois dos cortes e o nível de operação. Veja a margem, a situação e o que fazer."),
        ("5", "Aba Margem: com a margem do ring-out, estime a margem em cada momento do evento pelos microfones abertos."),
        ("6", "Aba Checklist: confira os doze itens antes da passagem de som."),
        ("7", "Depois do evento, aba Ocorrências: anote cada apito, a causa e o que foi feito. Aplique a medida definitiva."),
    ]),
    ("Abas da planilha", [
        ("Ring-out", "Até 8 saídas e 30 cortes, com ganho obtido, margem, situação de cada saída e de cada corte."),
        ("Ocorrências", "Até 20 apitos, com a medida definitiva sugerida para cada causa e a contagem por causa."),
        ("Margem", "Margem estimada por microfones abertos em cada momento do evento, e calculadora de distância."),
        ("Referência", "Faixas do apito, Q e largura, microfones abertos, medidas, posição do retorno e causas."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Ring-out", "O ring-out de um culto com três saídas e nove cortes (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Níveis", "Os níveis da aba Ring-out são a posição do master de cada saída, em dB. Use sempre o mesmo controle para medir e para operar."),
        ("Margem", "Apito depois dos cortes menos o nível de operação. Sem margem: abaixo de 3 dB. Margem curta: abaixo da margem mínima (6 dB)."),
        ("Cortes", "Corte fundo: além do corte mais fundo aceito (−9 dB). Corte largo: Q abaixo do mínimo (4). Realce: valor positivo."),
        ("Largura", "Largura em oitavas calculada pelo Q do paramétrico. No gráfico de 31 bandas, cerca de 1/3 de oitava."),
        ("Microfones abertos", "Cada vez que o número dobra, a margem cai cerca de 3 dB (10 × log da razão). É uma estimativa, e não uma medida."),
        ("Distância", "Metade da distância: +6 dB (20 × log da razão). Na caixa, o valor cheio só aparece ao ar livre."),
        ("Exemplos", "Os dados do exemplo são ilustrativos, criados para este material. As regras são práticas, e não norma."),
    ]),
])

ws_ring = wb.create_sheet("Ring-out")
ws_oc = wb.create_sheet("Ocorrências")
ws_mg = wb.create_sheet("Margem")
ws_ref = wb.create_sheet("Referência")
lookup = aba_referencia(ws_ref)
aba_ring(ws_ring, RED_H)
aba_ocorrencias(ws_oc, SLATE_H, lookup)
aba_margem(ws_mg, TEAL_H)
checklist(wb.create_sheet("Checklist"), "Checklist contra a microfonia",
          "Confira na montagem, antes da passagem de som. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["blue"][0])
aba_ring(wb.create_sheet("Exemplo - Ring-out"), MUTED, EX_RING)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
