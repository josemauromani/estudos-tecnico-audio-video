# -*- coding: utf-8 -*-
"""Gera Passagem-de-Som-modelo.xlsx: input list com mapa de palco, cronograma da passagem,
registro do evento (volume e ocorrências), referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *
from datetime import time

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIM_NAO = ["Sim", "Não"]
HORA = "hh:mm"
N = 24  # canais na input list

CHECK = [
    ("Preparação", "A input list está completa e ao lado da mesa: canal, fonte, microfone ou DI, entrada no palco e retorno."),
    ("Preparação", "O mapa de palco mostra cada posição, os retornos e as tomadas."),
    ("Preparação", "O cronograma foi montado de trás para a frente, e a passagem termina antes de abrir as portas."),
    ("Montagem", "Os cabos estão presos e fora dos caminhos, e o sistema foi ligado na ordem, com a mesa zerada."),
    ("Montagem", "O line check foi feito em todos os canais, retornos e caixas, e as pendências foram resolvidas."),
    ("Passagem", "O ganho foi ajustado canal por canal, com cada fonte no volume real do evento."),
    ("Passagem", "Cada músico confirmou o retorno com a banda tocando junta e com o PA aberto."),
    ("Passagem", "A mistura foi montada com uma música inteira: base firme, voz por cima, letra entendida."),
    ("Passagem", "O volume foi medido e está na faixa combinada, e a sala foi ouvida na frente, no meio e no fundo."),
    ("Passagem", "As transições foram passadas (fala, vídeo com som, música ambiente), e a cena foi salva ou a mesa fotografada."),
    ("Operação", "Durante o evento, os microfones sem uso ficam fechados, e o operador tem o roteiro e olha o palco."),
    ("Operação", "As ocorrências foram anotadas com a hora, o sistema foi desligado na ordem, e as pendências têm responsável."),
]

# Etapas e durações do cronograma da Figura 3 do treinamento (ponto de partida)
ETAPAS = [
    ("Montagem do palco e do sistema", 60),
    ("Ligar o sistema na ordem e zerar a mesa", 10),
    ("Line check: canais, retornos e caixas", 20),
    ("Rodada 1: ganho, canal por canal", 30),
    ("Rodada 2: retornos", 25),
    ("Rodada 3: PA e mistura", 25),
    ("Volume medido e caminhada pela sala", 10),
    ("Transições: fala, vídeo com som, música ambiente", 10),
    ("Folga", 10),
]
N_ETAPAS = 12

TIPOS_OC = ["Microfonia", "Ruído", "Sem sinal", "Volume", "Mute ou deixa", "Sem fio", "Mídia", "Outro"]

# Exemplo 1 do treinamento: o line check de um culto com dezesseis canais
EX_INPUT = {
    "head": {"evento": "Culto de domingo, 19 h (portas às 18h30)", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "mesa": "Mesa digital de 16 canais, caixa de palco com multicabo"},
    "canais": [  # fonte, tipo, microfone ou DI, phantom, pedestal, entrada, posição, retorno, line check, observação
        ("Bumbo", "Microfone", "Dinâmico de bumbo", "Não", "Baixo", 1, "A3", "Aux 4", "OK", "Dentro do bumbo, perto da pele."),
        ("Caixa da bateria", "Microfone", "Dinâmico", "Não", "Baixo", 2, "A3", "Aux 4", "OK", None),
        ("Over da bateria, lado E", "Microfone", "Condensador", "Sim", "Girafa", 3, "A3", None, "OK", "Não vai para os retornos."),
        ("Over da bateria, lado D", "Microfone", "Condensador", "Sim", "Girafa", 4, "A3", None, "OK", "Não vai para os retornos."),
        ("Baixo", "DI", "DI ativa", "Sim", "Não", 5, "A2", "Aux 4", "OK", "Saída paralela da DI para o amplificador."),
        ("Violão", "DI", "DI passiva", "Não", "Não", 6, "B2", "Aux 3", "Problema", "Sem sinal na mesa."),
        ("Teclado, lado E", "DI", "DI passiva", "Não", "Não", 7, "A4", "Aux 2", "OK", None),
        ("Teclado, lado D", "DI", "DI passiva", "Não", "Não", 7, "A4", "Aux 2", "Problema", "Não aparece na mesa."),
        ("Voz principal", "Sem fio", "Dinâmico de mão", "Não", "Reto", 9, "C3", "Aux 1", "OK", "Pilhas novas. Receptor ao lado do palco."),
        ("Voz de apoio 1", "Microfone", "Dinâmico de mão", "Não", "Girafa", 10, "C4", "Aux 2", "OK", None),
        ("Voz de apoio 2", "Microfone", "Dinâmico de mão", "Não", "Girafa", 11, "C5", "Aux 2", "OK", None),
        ("Pastor", "Sem fio", "Headset", "Não", "Não", 12, "C3", "Aux 1", "OK", "Pilhas novas."),
        ("Púlpito", "Microfone", "Condensador de púlpito", "Sim", "Não", 13, "C2", None, "OK", None),
        ("Computador, lado E", "Linha", None, "Não", "Não", "M14", None, None, "OK", "Vídeos e música ambiente. Direto na mesa."),
        ("Computador, lado D", "Linha", None, "Não", "Não", "M15", None, None, "OK", "Direto na mesa."),
        ("Microfone reserva", "Microfone", "Dinâmico de mão", "Não", "Reto", 16, "C2", None, None, "Com fio, ao lado do púlpito."),
    ],
    "zonas": {"A1": "Caixa de palco", "A2": "1 tomada (amplificador do baixo)", "A3": "Retorno R4 (bateria e baixo)",
              "A4": "1 tomada (teclado)", "B2": "Retorno R3 · 1 tomada", "C3": "Retorno R1", "C4": "Retorno R2"},
    "ajustes": [  # canal, o que foi feito, situação depois
        ("6", "Sem sinal na mesa. Com outro cabo entre a DI e a caixa de palco, o sinal chegou: cabo trocado e separado para conserto.", "OK"),
        ("7 e 8", "A lista apontava a entrada 7 para os dois lados do teclado. Teclado D ligado na entrada 8, e a lista corrigida.", "OK"),
        ("16", "Microfone reserva ligado ao lado do púlpito e testado.", "OK"),
        ("Todos", "As dezesseis linhas estavam certas quando a banda subiu ao palco. Line check feito em 20 minutos, por duas pessoas.", ""),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def sub(ws, ref, text, merge=None):
    put(ws, ref, text, f=font(10, True), bg=GRAY, h="center", merge=merge)


# ================================================================== aba Input list e mapa
I_W = {"A": 2, "B": 9, "C": 30, "D": 12, "E": 20, "F": 9, "G": 10, "H": 9, "I": 9, "J": 10, "K": 11, "L": 22, "M": 34, "N": 2}
G1 = 14
ZONA_COLS = [("C", "C"), ("D", "E"), ("F", "I"), ("J", "L"), ("M", "M")]  # as cinco posições, da esquerda para a direita
ZONA_CAB = ["1 · esquerda", "2", "3 · centro", "4", "5 · direita"]
FUNDOS = [("A", "A\nfundo"), ("B", "B\nmeio"), ("C", "C\nfrente")]
ZONAS = [f"{l}{n}" for l, _ in FUNDOS for n in range(1, 6)]


def aba_input(ws, tab, data=None):
    widths(ws, I_W)
    is_ex = data is not None
    title(ws, "Input list e mapa de palco",
          "Exemplo preenchido, para consulta. Use a aba Input list e mapa para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro. As células cinza são calculadas. O mapa de palco fica no pé da aba.", "M")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None), (8, "Mesa e caixa de palco", "mesa", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    dicas = ["Uma linha por canal da mesa, na mesma ordem toda semana.",
             "Entrada: número da entrada na caixa de palco. Fonte ligada direto na mesa: M e o número (M14).",
             "Posição: letra é a profundidade (A fundo, B meio, C frente); número vai de 1, esquerda, a 5, direita.",
             "Esquerda e direita são sempre vistas da plateia, que é a visão de quem está na mesa.",
             "Line check: marque OK ou Problema em cada linha, antes de a banda subir ao palco."]
    for k, t in enumerate(dicas):
        hint(ws, f"I{4 + k}", t, merge=f"I{4 + k}:M{4 + k}", h="left")

    band(ws, 10, "Input list — uma linha por canal", "M")
    heads = [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Tipo", BLUE), ("E", "Microfone ou DI", BLUE),
             ("F", "Phantom", BLUE), ("G", "Pedestal", BLUE), ("H", "Entrada no palco", BLUE), ("I", "Posição no mapa", BLUE),
             ("J", "Retorno", BLUE), ("K", "Line check", AMBER), ("L", "Situação", TEAL), ("M", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLM", ["nº na mesa", "o que entra no canal", "lista", "modelo ou tipo", "Sim / Não", "lista",
                                       "nº ou M+nº", "A1 a C5", "mix de quem toca", "OK / Problema", "", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 27
    if is_ex:
        hint(ws, "B13", "Estado em que o line check encontrou as linhas, antes das correções. A coluna Situação e o resumo são calculados.",
             merge="B13:M13", h="left")
    else:
        for col, v in zip("BCDEFGHIJKL", ["Ex.", "Voz principal", "Sem fio", "Dinâmico de mão", "Não", "Reto", 9, "C3", "Aux 1", "OK", "OK"]):
            ex(ws, f"{col}13", v, h="left" if col in "CE" else "center")
        ex(ws, "M13", "Pilhas novas.")
        note(ws, "D11", "Microfone: com fio.\nSem fio: o receptor entra na mesa ou na caixa de palco.\nDI: instrumento por direct box.\nLinha: fonte de linha ligada direto.")
        note(ws, "F11", "O canal precisa da alimentação phantom (+48 V)?\nCondensadores e direct boxes ativas costumam precisar.")
        note(ws, "H11", "Em que entrada da caixa de palco ou do multicabo o cabo é ligado.\nA mesma entrada em dois canais gera o aviso Entrada repetida.")
        note(ws, "I11", "Posição do músico ou da fonte no palco.\nO mapa no pé da aba é montado por esta coluna.")
        note(ws, "J11", "Auxiliar ou mix de retorno em que quem toca este canal se ouve.")
        note(ws, "L11", "OK só aparece quando a linha tem tipo, microfone ou DI, entrada sem repetição e line check OK.")
    ws.row_dimensions[13].height = 24

    G2 = G1 + N - 1
    bg = WHITE if is_ex else INPUT
    for k in range(N):
        r = G1 + k
        c = data["canais"][k] if is_ex and k < len(data["canais"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, h="center", bg=bg)
        inp(ws, f"E{r}", c[2] if c else None, bg=bg)
        for col, j in zip("FGHIJK", range(3, 9)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg)
        calc(ws, f"L{r}", f'=IF(C{r}="","",IF(D{r}="","Falta o tipo",IF(AND(D{r}<>"Linha",E{r}=""),"Falta microfone ou DI",'
                          f'IF(H{r}="","Falta a entrada",IF(COUNTIF($H${G1}:$H${G2},H{r})>1,"Entrada repetida",'
                          f'IF(K{r}="Problema","Linha com problema",IF(K{r}="","Sem line check","OK")))))))')
        inp(ws, f"M{r}", c[9] if c else None, bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{G1}:D{G2}", ["Microfone", "Sem fio", "DI", "Linha"], "Microfone, Sem fio, DI ou Linha.")
    dv_list(ws, f"F{G1}:F{G2}", SIM_NAO, "O canal precisa de phantom (+48 V)?")
    dv_list(ws, f"G{G1}:G{G2}", ["Girafa", "Reto", "Baixo", "Clipe", "Não"], "Girafa, Reto, Baixo, Clipe ou Não.")
    dv_list(ws, f"I{G1}:I{G2}", ZONAS, "De A1 a C5. A é o fundo do palco, C é a frente; 1 é a esquerda, vista da plateia.")
    dv_list(ws, f"K{G1}:K{G2}", ["OK", "Problema"], "Resultado do line check desta linha.")
    cf_status(ws, f"L{G1}:L{G2}", f"L{G1}")

    # ---------------- resumo e lista de material
    s = G2 + 2
    band(ws, s, "Resumo e lista de material", "M")
    cnt = lambda col, t: f'COUNTIF({col}{G1}:{col}{G2},"{t}")'
    linhas = [
        ("Canais em uso", f"=COUNTA(C{G1}:C{G2})"),
        ("Microfones com fio", f'={cnt("D", "Microfone")}'),
        ("Microfones sem fio", f'={cnt("D", "Sem fio")}'),
        ("Direct boxes (DI)", f'={cnt("D", "DI")}'),
        ("Fontes de linha", f'={cnt("D", "Linha")}'),
        ("Canais com phantom", f'={cnt("F", "Sim")}'),
        ("Pedestais (girafa, reto e baixo)", f'={cnt("G", "Girafa")}+{cnt("G", "Reto")}+{cnt("G", "Baixo")}'),
        ("Cabos de microfone no palco (com fio e DI)", f'={cnt("D", "Microfone")}+{cnt("D", "DI")}'),
        ("Linhas com line check OK", f'={cnt("K", "OK")}'),
        ("Linhas a resolver", f'=F{s + 1}-{cnt("L", "OK")}'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste os canais",IF(F{s + 10}>0,"Há linhas a resolver antes de a banda subir ao palco","OK"))',
         b=False, merge=f"F{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:M{a}")
    hint(ws, f"I{s + 1}", "A conta de cabos é o mínimo: some os cabos dos receptores sem fio, dos retornos e os de reserva.",
         merge=f"I{s + 1}:M{s + 2}", h="left")

    # ---------------- mapa de palco
    m = a + 2
    band(ws, m, "Mapa de palco — montado pela coluna Posição", "M")
    hint(ws, f"B{m + 1}", "FUNDO DO PALCO", merge=f"B{m + 1}:M{m + 1}")
    put(ws, f"B{m + 2}", "", bg=GRAY)
    for (c1, c2), t in zip(ZONA_COLS, ZONA_CAB):
        sub(ws, f"{c1}{m + 2}", t, merge=None if c1 == c2 else f"{c1}{m + 2}:{c2}{m + 2}")
    ws.row_dimensions[m + 2].height = 21.75
    Z = data.get("zonas", {}) if is_ex else {}
    for i, (letra, rot) in enumerate(FUNDOS):
        r1, r2 = m + 3 + 2 * i, m + 4 + 2 * i
        put(ws, f"B{r1}", rot, f=font(10, True), bg=GRAY, h="center", merge=f"B{r1}:B{r2}")
        for n, (c1, c2) in enumerate(ZONA_COLS, 1):
            cod = f"{letra}{n}"
            f = "=" + "&".join(f'IF($I${r}="{cod}","• "&$B${r}&" "&$C${r}&"   ","")' for r in range(G1, G2 + 1))
            mg = lambda r: None if c1 == c2 else f"{c1}{r}:{c2}{r}"
            calc(ws, f"{c1}{r1}", f, b=True, merge=mg(r1), sz=9)
            inp(ws, f"{c1}{r2}", Z.get(cod), h="center", merge=mg(r2), bg=bg)
            ws.conditional_formatting.add(f"{c1}{r1}:{c2}{r1}", FormulaRule(
                formula=[f'LEN(${c1}${r1})>0'], fill=PatternFill("solid", bgColor=TEAL_T, fgColor=TEAL_T)))
        ws.row_dimensions[r1].height = 54
        ws.row_dimensions[r2].height = 24
    p = m + 9
    hint(ws, f"B{p}", "PLATEIA  ·  esquerda e direita vistas da plateia (da mesa de som)", merge=f"B{p}:M{p}")
    hint(ws, f"B{p + 1}", "Em cada posição: em cima, os canais que a input list põe ali (calculado); embaixo, anote os retornos e as tomadas.",
         merge=f"B{p + 1}:M{p + 1}", h="left")
    end = p + 1

    if is_ex:
        band(ws, end + 2, "O que foi corrigido depois do line check", "M")
        r = end + 3
        sub(ws, f"B{r}", "Canal")
        sub(ws, f"C{r}", "O que foi feito", merge=f"C{r}:K{r}")
        sub(ws, f"L{r}", "Situação depois", merge=f"L{r}:M{r}")
        ws.row_dimensions[r].height = 21.75
        for canal, feito, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", canal, f=font(10, True, c=MUTED), bg=GRAY, h="center")
            put(ws, f"C{r}", feito, merge=f"C{r}:K{r}")
            put(ws, f"L{r}", depois, merge=f"L{r}:M{r}", h="center")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{G1}"
    setup(ws, tab, f"B1:M{end}", fit_height=True)


# ================================================================== aba Cronograma
def aba_cronograma(ws, tab):
    widths(ws, {"A": 2, "B": 5, "C": 44, "D": 22, "E": 11, "F": 10, "G": 10, "H": 12, "I": 10, "J": 14, "K": 36, "L": 2})
    title(ws, "Cronograma da passagem — de trás para a frente",
          "Informe a hora de abrir as portas e as durações. A planilha calcula a hora de cada etapa e a hora em que a equipe precisa começar.", "K")
    for r, rot in [(4, "Evento"), (5, "Data"), (6, "Responsável")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:F{r}", fmt=DATE if r == 5 else None)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    for r, rot, dica in [(4, "Início do evento", "Só para registro. Digite as horas no formato 19:00."),
                         (5, "Abertura das portas", "A passagem precisa terminar antes desta hora."),
                         (6, "Chegada da equipe", "Hora em que a equipe pode começar a montar.")]:
        label(ws, f"G{r}", rot, merge=f"G{r}:I{r}")
        inp(ws, f"J{r}", h="center", fmt=HORA)
        hint(ws, f"K{r}", dica, h="left")
    PORTAS, CHEG = "$J$5", "$J$6"

    band(ws, 8, "Etapas, na ordem em que acontecem", "K")
    for col, t, bg in [("B", "#", INK), ("C", "Etapa", INK), ("D", "Responsável", INK), ("E", "Duração", BLUE),
                       ("F", "Início", TEAL), ("G", "Fim", TEAL), ("H", "Terminou às", AMBER), ("I", "Atraso", TEAL),
                       ("J", "Situação", TEAL), ("K", "Observação", INK)]:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJK", ["", "sem pular linhas", "", "minutos", "calculado", "calculado", "hora real", "minutos", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    note(ws, "E9", "Duração prevista, em minutos.\nOs valores iniciais são um ponto de partida: troque pelos tempos da sua equipe.")
    note(ws, "F9", "Contado para trás: a última etapa termina na abertura das portas, e cada etapa termina quando a seguinte começa.")
    note(ws, "H9", "Hora em que a etapa terminou de fato, no formato 17:10.\nPreencha durante a passagem.")
    note(ws, "I9", "Diferença entre a hora real e o fim previsto.\nPositivo: atrasou. Negativo: adiantou.")

    r1, r2 = 11, 11 + N_ETAPAS - 1
    for k in range(N_ETAPAS):
        r = r1 + k
        e = ETAPAS[k] if k < len(ETAPAS) else (None, None)
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", e[0])
        inp(ws, f"D{r}")
        inp(ws, f"E{r}", e[1], h="center")
        prox = PORTAS if r == r2 else f'IF(F{r + 1}="",{PORTAS},F{r + 1})'
        calc(ws, f"G{r}", f'=IF(OR(E{r}="",{PORTAS}=""),"",{prox})', b=False, fmt=HORA)
        calc(ws, f"F{r}", f'=IF(G{r}="","",G{r}-E{r}/1440)', fmt=HORA)
        inp(ws, f"H{r}", h="center", fmt=HORA)
        calc(ws, f"I{r}", f'=IF(OR(H{r}="",G{r}=""),"",ROUND((H{r}-G{r})*1440,0))', b=False, fmt='+0;-0;0')
        calc(ws, f"J{r}", f'=IF(I{r}="","",IF(I{r}>0,"Atrasou","OK"))')
        inp(ws, f"K{r}")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"E{r1}:E{r2}")
    cf_status(ws, f"J{r1}:J{r2}", f"J{r1}")

    s = r2 + 2
    band(ws, s, "Leitura do cronograma", "K")
    summary(ws, s + 1, "Duração total, em minutos", f'=IF(SUM(E{r1}:E{r2})=0,"",SUM(E{r1}:E{r2}))', "D", "E")
    summary(ws, s + 2, "A equipe precisa começar às", f'=IF(F{r1}="","",F{r1})', "D", "E", fmt=HORA)
    summary(ws, s + 3, "Folga em relação à chegada da equipe, em minutos",
            f'=IF(OR(F{r1}="",{CHEG}=""),"",ROUND((F{r1}-{CHEG})*1440,0))', "D", "E", fmt='+0;-0;0')
    summary(ws, s + 4, "Maior atraso de uma etapa, em minutos", f'=IF(COUNT(I{r1}:I{r2})=0,"",MAX(I{r1}:I{r2}))', "D", "E", fmt='+0;-0;0')
    label(ws, f"B{s + 5}", "Aviso", merge=f"B{s + 5}:D{s + 5}", h="right")
    calc(ws, f"E{s + 5}", f'=IF({PORTAS}="","Informe a hora de abrir as portas",IF(E{s + 1}="","Informe as durações",'
                          f'IF({CHEG}="","Informe a hora de chegada da equipe",'
                          f'IF(E{s + 3}<0,"Faltam "&ABS(E{s + 3})&" min: chegue mais cedo ou encurte etapas","OK"))))',
         b=False, merge=f"E{s + 5}:K{s + 5}", h="left")
    ws.row_dimensions[s + 5].height = 21.75
    cf_ok(ws, f"E{s + 5}:K{s + 5}")
    hint(ws, f"B{s + 6}", "Avise aos músicos a hora de início da rodada 1, e não a da montagem. As durações iniciais são típicas de uma banda de seis a oito pessoas.",
         merge=f"B{s + 6}:K{s + 6}", h="left")
    ws.freeze_panes = f"D{r1}"
    setup(ws, tab, f"B1:K{s + 6}", fit_height=True)


# ================================================================== aba Registro do evento
def aba_registro(ws, tab):
    widths(ws, {"A": 2, "B": 5, "C": 9, "D": 24, "E": 24, "F": 15, "G": 40, "H": 40, "I": 13, "J": 13, "K": 13, "L": 2})
    title(ws, "Registro do evento — volume e ocorrências",
          "Anote na hora, em poucas palavras. O que fica pendente aparece no resumo, para resolver antes do próximo evento.", "K")
    for r, rot in [(4, "Evento"), (5, "Data"), (6, "Operador")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:E{r}", fmt=DATE if r == 5 else None)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    label(ws, "F4", "Faixa combinada: mínimo, em dB(A)", merge="F4:G4")
    inp(ws, "H4", h="center")
    label(ws, "F5", "Faixa combinada: máximo, em dB(A)", merge="F5:G5")
    inp(ws, "H5", h="center")
    label(ws, "F6", "Como e onde a faixa é medida", merge="F6:G6")
    inp(ws, "H6", merge="H6:K6")
    hint(ws, "I4", "A faixa é a que a equipe e a liderança combinarem.", merge="I4:K4", h="left")
    hint(ws, "I5", "Exemplo de medição: dB(A), resposta lenta, meio da sala.", merge="I5:K5", h="left")
    dv_number(ws, "H4:H5")
    MIN, MAX = "$H$4", "$H$5"

    # ---------------- volume
    band(ws, 8, "Volume da sala — leituras ao longo do evento", "K")
    for col, t, bg in [("B", "#", INK), ("C", "Hora", INK), ("D", "Momento do evento", INK), ("E", "Ponto de medição", INK),
                       ("F", "Leitura", AMBER), ("G", "Situação", TEAL)]:
        head(ws, f"{col}9", t, bg=bg)
    put(ws, "H9", "Observação", f=font(10, True, c=WHITE), bg=INK, h="center", merge="H9:K9")
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFG", ["", "19:20", "louvor, pregação…", "meio da sala, mesa…", "dB(A), só o número", ""]):
        hint(ws, f"{col}10", t)
    hint(ws, "H10", "", merge="H10:K10")
    ws.row_dimensions[10].height = 21.75
    v1, v2 = 11, 18
    for k, r in enumerate(range(v1, v2 + 1), 1):
        num(ws, f"B{r}", k)
        inp(ws, f"C{r}", h="center", fmt=HORA)
        inp(ws, f"D{r}")
        inp(ws, f"E{r}")
        inp(ws, f"F{r}", h="center")
        calc(ws, f"G{r}", f'=IF(F{r}="","",IF(OR({MIN}="",{MAX}=""),"Informe a faixa combinada",'
                          f'IF(F{r}>{MAX},"Acima da faixa em "&ROUND(F{r}-{MAX},0)&" dB",'
                          f'IF(F{r}<{MIN},"Abaixo da faixa em "&ROUND({MIN}-F{r},0)&" dB","OK"))))', h="left")
        inp(ws, f"H{r}", merge=f"H{r}:K{r}")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"F{v1}:F{v2}")
    cf_status(ws, f"G{v1}:G{v2}", f"G{v1}")
    note(ws, "F9", "Leitura do medidor de nível sonoro, só o número.\nUse sempre a mesma ponderação, a mesma resposta e o mesmo aparelho.")
    summary(ws, v2 + 1, "Leituras registradas", f"=COUNT(F{v1}:F{v2})", "E", "F")
    summary(ws, v2 + 2, "Maior leitura, em dB(A)", f'=IF(COUNT(F{v1}:F{v2})=0,"",MAX(F{v1}:F{v2}))', "E", "F")
    summary(ws, v2 + 3, "Leituras acima da faixa", f'=IF({MAX}="","",COUNTIF(F{v1}:F{v2},">"&{MAX}))', "E", "F")
    hint(ws, f"G{v2 + 1}", "Leituras de pontos diferentes não se comparam direto: a primeira fileira costuma ler mais do que o meio da sala.",
         merge=f"G{v2 + 1}:K{v2 + 1}", h="left")

    # ---------------- ocorrências
    b = v2 + 5
    band(ws, b, "Ocorrências — o que aconteceu e o que fica para resolver", "K")
    for col, t, bg in [("B", "#", INK), ("C", "Hora", INK), ("D", "Momento do evento", INK), ("E", "Canal ou equipamento", INK),
                       ("F", "Tipo", AMBER), ("G", "O que aconteceu", AMBER), ("H", "O que foi feito na hora", AMBER),
                       ("I", "Resolver antes do próximo?", AMBER), ("J", "Andamento", AMBER), ("K", "Pendência", TEAL)]:
        head(ws, f"{col}{b + 1}", t, bg=bg)
    ws.row_dimensions[b + 1].height = 30
    for col, t in zip("BCDEFGHIJK", ["", "19:42", "", "", "lista", "em poucas palavras", "", "Sim / Não", "Aberto / Resolvido", ""]):
        hint(ws, f"{col}{b + 2}", t)
    ws.row_dimensions[b + 2].height = 21.75
    for col, v in zip("BCDEFGHIJK", ["Ex.", time(19, 42), "Louvor, 3ª música", "Canal 6, violão", "Ruído", "Estalos ao mexer no cabo do violão.",
                                     "Cabo trocado entre as músicas.", "Sim", "Aberto", "Pendente"]):
        ex(ws, f"{col}{b + 3}", v, h="left" if col in "DEGH" else "center", fmt=HORA if col == "C" else None)
    ws.row_dimensions[b + 3].height = 24
    o1, o2 = b + 4, b + 23
    for k, r in enumerate(range(o1, o2 + 1), 1):
        num(ws, f"B{r}", k)
        inp(ws, f"C{r}", h="center", fmt=HORA)
        for col in "DEGH":
            inp(ws, f"{col}{r}")
        for col in "FIJ":
            inp(ws, f"{col}{r}", h="center")
        calc(ws, f"K{r}", f'=IF(G{r}="","",IF(AND(I{r}="Sim",J{r}<>"Resolvido"),"Pendente","OK"))')
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"F{o1}:F{o2}", TIPOS_OC, "Escolha o tipo que mais se aproxima.")
    dv_list(ws, f"I{o1}:I{o2}", SIM_NAO, "Precisa ser resolvido antes do próximo evento?")
    dv_list(ws, f"J{o1}:J{o2}", ["Aberto", "Resolvido"], "Aberto ou Resolvido.")
    cf_status(ws, f"K{o1}:K{o2}", f"K{o1}")
    note(ws, f"K{b + 1}", "Pendente: marcada para resolver antes do próximo evento e ainda não resolvida.")

    s = o2 + 2
    band(ws, s, "Resumo das ocorrências", "K")
    summary(ws, s + 1, "Ocorrências registradas", f"=COUNTA(G{o1}:G{o2})", "E", "F")
    summary(ws, s + 2, "Pendentes para o próximo evento", f'=COUNTIF(K{o1}:K{o2},"Pendente")', "E", "F")
    for k, t in enumerate(TIPOS_OC):
        summary(ws, s + 3 + k, f"Tipo: {t}", f'=COUNTIF(F{o1}:F{o2},"{t}")', "E", "F")
    a = s + 3 + len(TIPOS_OC)
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 2}>0,"Há pendências: defina quem resolve, antes do próximo evento","OK")',
         b=False, merge=f"F{a}:K{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:K{a}")
    hint(ws, f"G{s + 1}", "O tipo que mais se repete, semana após semana, mostra onde a equipe deve investir: cabos, pilhas, roteiro, treinamento.",
         merge=f"G{s + 1}:K{s + 2}", h="left")
    ws.freeze_panes = "D11"
    setup(ws, tab, f"B1:K{a}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Passagem de som e mixagem ao vivo — modelo de input list, cronograma e registro"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Passagem de som — como usar esta planilha",
           "Input list com mapa de palco, cronograma da passagem e registro do evento, para usar do planejamento à desmontagem.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo, phantom, pedestal, posição, line check, tipo de ocorrência e andamento aceitam só as opções da lista."),
        ("Horas", "Digite as horas no formato 18:30. As durações são em minutos, só o número."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Dias antes, aba Input list e mapa: liste os canais, com tipo, microfone ou DI, phantom, pedestal, entrada e retorno."),
        ("2", "Dê a cada canal uma posição (de A1 a C5) e confira o mapa de palco no pé da aba. Anote os retornos e as tomadas."),
        ("3", "Separe o material pelo resumo da aba: microfones, direct boxes, pedestais e cabos."),
        ("4", "Aba Cronograma: informe a abertura das portas e a chegada da equipe, e ajuste as durações. Avise aos músicos a hora da rodada 1."),
        ("5", "No dia, faça o line check e marque OK ou Problema em cada linha da input list. Resolva o que a coluna Situação apontar."),
        ("6", "Durante a passagem, anote na aba Cronograma a hora em que cada etapa terminou."),
        ("7", "Aba Registro do evento: informe a faixa de volume combinada, anote as leituras e registre as ocorrências na hora."),
        ("8", "Aba Checklist: confira os doze itens. Depois do evento, veja as pendências e defina quem resolve."),
    ]),
    ("Abas da planilha", [
        ("Input list e mapa", "Lista de até 24 canais com a situação de cada linha, lista de material e mapa de palco em tabela."),
        ("Cronograma", "Até 12 etapas, com início e fim contados para trás a partir da abertura das portas, folga e atrasos."),
        ("Registro do evento", "Oito leituras de volume, comparadas com a faixa combinada, e até 20 ocorrências, com as pendências."),
        ("Referência", "Line check e soundcheck, as três rodadas, prioridades de mixagem, faixas de volume e durações típicas."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Input list", "Os dezesseis canais de um culto, como o line check os encontrou (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Situação da linha", "OK exige tipo, microfone ou DI (menos em fonte de linha), entrada no palco sem repetição e line check OK."),
        ("Entrada repetida", "O aviso aparece nas duas linhas que usam a mesma entrada. Fonte ligada direto na mesa: escreva M e o número do canal."),
        ("Mapa de palco", "Três profundidades (A fundo, B meio, C frente) e cinco posições, de 1, esquerda, a 5, direita, vistas da plateia."),
        ("Cronograma", "A última etapa preenchida termina na abertura das portas. Preencha as etapas sem pular linhas. Eventos que cruzam a meia-noite não são tratados."),
        ("Durações iniciais", "São um ponto de partida, típico de uma banda de seis a oito pessoas. Meça os tempos da sua equipe e troque."),
        ("Faixa de volume", "A planilha não sugere a faixa: ela é combinada entre a equipe e a liderança, respeitando os limites de exposição ao som."),
        ("Exemplo", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_input(wb.create_sheet("Input list e mapa"), BLUE)
aba_cronograma(wb.create_sheet("Cronograma"), AMBER)
aba_registro(wb.create_sheet("Registro do evento"), TEAL)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 34, "C": 34, "D": 34, "E": 60, "F": 2})
title(ws, "Referência — a passagem em tabelas", "Valores e durações típicos, para servir de ponto de partida. Ajuste à sua sala e à sua equipe.", "E")


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


r = tabela(4, "Line check e soundcheck", ["Conferência", "Quem faz", "Quando", "O que se confere"], [
    ("Line check", "duas pessoas da equipe", "antes de a banda subir", "Cada linha chega ao canal certo, sem ruído. Phantom onde precisa. Cada retorno e cada caixa tocando."),
    ("Soundcheck (passagem)", "o operador e os músicos", "depois do line check", "Nível e timbre de cada canal, retorno de cada músico, mistura da sala e volume."),
])
r = tabela(r, "As três rodadas da passagem", ["Rodada", "Quem toca", "Caixas da frente", "O que o operador faz"], [
    ("1 · Ganho", "um músico por vez", "fechadas", "Ganho pelo medidor, com o PFL, no volume real. Filtro de graves e equalização básica. Anota o ganho."),
    ("2 · Retornos", "a banda, um trecho", "fechadas ou baixas", "Um retorno por vez: primeiro o próprio músico, depois o que serve de guia. Abre o PA no fim e confirma."),
    ("3 · PA", "a banda, uma música inteira", "no volume do evento", "Base primeiro, voz por cima. Mede o volume, caminha pela sala, passa as transições e salva a cena."),
])
r = tabela(r, "Prioridades da mistura por tipo de evento", ["Evento", "Prioridade", "O que fica aberto", "Como a mistura se organiza"], [
    ("Louvor congregacional", "a letra e a melodia", "banda, vozes e reverb", "Voz principal clara, apoio logo abaixo, base firme. O público precisa conseguir se ouvir cantando."),
    ("Pregação, palestra, leitura", "cada palavra", "só o microfone de quem fala", "Sem reverb. Música de fundo, se houver, de 10 a 20 dB abaixo da voz, acertada pelo ouvido."),
    ("Apresentação musical, coral", "o conjunto", "os microfones do grupo", "O equilíbrio vem do palco. O PA reforça por igual e destaca solos."),
    ("Formatura, cerimônia", "os nomes e os discursos", "um microfone por vez", "Vídeos no nível da fala. Música de entrada mais forte, baixada antes de alguém falar."),
])
r = tabela(r, "Volume na plateia — faixas típicas, em dB(A), resposta lenta", ["Situação", "Faixa típica", "Onde medir", "Observação"], [
    ("Conversa a um metro", "55 a 65", "–", "Só para comparar."),
    ("Palestra, pregação, leitura", "65 a 75", "meio da sala", "A voz precisa ficar bem acima do ruído da sala."),
    ("Culto ou evento com banda", "80 a 95", "meio da sala", "Depende do estilo, da sala e do público. A faixa do seu evento é a que for combinada."),
    ("Show com banda forte", "95 a 105", "meio da sala", "Exige cuidado com o tempo de exposição da equipe e do público."),
    ("Diferença entre a frente e o fundo", "até uns 6 dB", "primeira fileira e fundo", "Acima disso, o problema é de cobertura das caixas ou de som de palco, e não de fader."),
])
r = tabela(r, "Durações típicas das etapas", ["Etapa", "Duração", "Quem participa", "Observação"], [
    ("Montagem do palco e do sistema", "45 a 90 min", "equipe técnica", "Depende de quanto fica montado de uma semana para a outra."),
    ("Line check", "15 a 30 min", "equipe técnica", "Cerca de um minuto por canal, mais os retornos e as caixas."),
    ("Rodada 1: ganho", "20 a 40 min", "equipe e músicos", "Um a dois minutos por canal."),
    ("Rodada 2: retornos", "15 a 30 min", "equipe e músicos", "Cerca de cinco minutos por retorno."),
    ("Rodada 3: PA e mistura", "15 a 30 min", "equipe e músicos", "Uma ou duas músicas inteiras."),
    ("Volume, caminhada e transições", "15 a 20 min", "equipe e quem fala", "Inclui tocar cada vídeo e cada música de fundo por alguns segundos."),
    ("Folga", "10 a 15 min", "–", "Absorve atrasos. Sem ela, qualquer imprevisto empurra a passagem para depois da abertura das portas."),
])
r = tabela(r, "Durante o evento: resposta rápida", ["Imprevisto", "Primeira ação", "Depois", "Registro"], [
    ("Apito (microfonia)", "baixar alguns dB o fader ou o envio suspeito", "investigar a causa", "Canal, momento e o que estava aberto."),
    ("Microfone falhou", "trocar pelo microfone reserva", "testar pilha, cabo e canal", "Qual microfone, e o que se achou."),
    ("Ruído forte em um canal", "fechar o canal", "trocar o cabo no intervalo", "Marcar o cabo para conserto."),
    ("Primeira palavra perdida", "abrir o microfone", "marcar a deixa no roteiro", "Em que transição aconteceu."),
])
hint(ws, f"B{r - 1}", "Valores típicos, que variam com a sala, o sistema, o estilo e a equipe. Não são norma: servem de ponto de partida para a conversa e para o cronograma.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist da passagem de som",
          "Confira ao longo do dia do evento. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_input(wb.create_sheet("Exemplo - Input list"), MUTED, EX_INPUT)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
