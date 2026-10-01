# -*- coding: utf-8 -*-
"""Gera Operacao-do-Evento-modelo.xlsx: roteiro técnico, escala e cronograma, inventário, ocorrências e revisão."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *
from datetime import time

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
HORA = "hh:mm"
SIM_NAO = ["Sim", "Não"]
MIN = '+0;-0;0'

CHECK = [
    ("Antes", "O roteiro técnico tem duração, deixa e instrução de som, vídeo, luz e mídia em cada momento."),
    ("Antes", "A soma das durações cabe no horário, e o que sai em caso de atraso está combinado."),
    ("Antes", "Todos os postos têm a mesma versão do roteiro, e ele foi lido em voz alta com a equipe."),
    ("Antes", "Cada função tem um responsável confirmado, e as funções críticas têm reserva."),
    ("Antes", "O cronograma foi contado para trás, com folga, e a montagem terminou antes de abrir as portas."),
    ("Antes", "A tensão, o terra e a carga de cada circuito foram conferidos, e nenhuma régua está ligada em outra."),
    ("Antes", "Os cabos nas passagens estão presos, os tripés estão travados e as saídas estão livres."),
    ("Antes", "O kit de emergência está à mão, e o plano B dos itens críticos tem responsável."),
    ("Durante", "A comunicação foi testada com todos os postos, e os comandos seguem o padrão combinado."),
    ("Durante", "A hora real de cada momento e as ocorrências foram anotadas durante o evento."),
    ("Depois", "O sistema foi desligado na ordem, os itens foram contados e o que falhou recebeu etiqueta."),
    ("Depois", "A reunião de revisão foi feita, e cada ação tem responsável e prazo."),
]

# Exemplo 1 do treinamento: o culto que sempre terminava atrasado
EX_ROT = {
    "head": {"evento": "Culto de domingo, com transmissão", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "diretor": "Diretor técnico da escala", "versao": "Versão 3, fechada no sábado às 18:00"},
    "par": (time(19, 0), time(20, 30), 5, 3),
    "linhas": [  # momento, duração, quem conduz, som, vídeo, luz, mídia, deixa, hora real
        ("Abertura e boas-vindas", 3, "Dirigente", "Mic do dirigente", "Câmera 2 de frente", "Palco aceso",
         "Tela de abertura", "Comando do diretor técnico", time(19, 2)),
        ("Louvor, três músicas", 15, "Dirigente e banda", "Banda aberta, voz principal por cima; púlpito fechado",
         "Câmera 1 no geral; 2 e 3 nos músicos e na voz", "Cena de louvor; plateia em meia-luz",
         "Letras, música por música", "O dirigente diz “vamos cantar”", time(19, 5)),
        ("Avisos, em vídeo", 4, "Vídeo", "Som do vídeo no canal 11; banda fechada", "Vídeo no programa",
         "Palco em meia-luz", "avisos-04-out.mp4, com áudio", "Último acorde da terceira música", time(19, 21)),
        ("Leitura bíblica", 3, "Leitor", "Mic do púlpito", "Câmera 2 no leitor", "Foco no púlpito",
         "Texto bíblico", "O vídeo termina", time(19, 25)),
        ("Oferta, com música", 5, "Banda", "Teclado, violão e vozes de apoio", "Câmera 1 no geral e em detalhes",
         "Cena de louvor", "Letra da música", "O leitor diz “amém”", time(19, 28)),
        ("Coral, uma música", 5, "Coral", "Mics do coral abertos; banda fechada", "Câmera 1 no coral",
         "Palco aceso", "Letra do coral", "O coral está posicionado", time(19, 33)),
        ("Pregação", 40, "Pastor", "Lapela do pastor; coral fechado", "Câmera 2 no pastor; câmera 1 no geral",
         "Foco no púlpito; plateia em meia-luz", "Versículos citados", "O coral desce; o pastor sobe", time(19, 38)),
        ("Oração, com teclado de fundo", 7, "Pastor", "Lapela e teclado baixo", "Câmera 2 no pastor",
         "Palco em meia-luz", "Tela de oração", "O pastor diz “vamos orar”", time(20, 21)),
        ("Música final", 5, "Banda", "Banda aberta", "Câmeras 1, 2 e 3", "Cena de louvor", "Letra da música",
         "O pastor diz “amém”", time(20, 28)),
        ("Bênção e encerramento", 3, "Pastor", "Lapela do pastor; banda em fundo", "Câmera 2; depois, tela final",
         "Plateia acesa no fim", "Tela de encerramento", "Último acorde da música", time(20, 32)),
    ],
    "revisao": [  # o que, quem, até quando
        ("Começar na hora: contagem regressiva de 5 min no telão, disparada às 18:55.", "Operador de mídia", "Próximo domingo"),
        ("Relógio de palco visível do púlpito, com aviso aos 35 min de pregação.", "Diretor técnico", "Próximo domingo"),
        ("Roteiro modelo com as durações medidas: louvor 17 min, pregação 40 a 45 min.", "Diretor técnico", "Quinta-feira"),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def dv_time(ws, rng):
    dv = DataValidation(type="time", operator="between", formula1="0", formula2="0.99999", allow_blank=True)
    dv.errorTitle, dv.error = "Hora inválida", "Digite a hora no formato 19:00."
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def heads(ws, row, spec, hints_row=None, hints=None):
    for col, t, bg in spec:
        head(ws, f"{col}{row}", t, bg=bg)
    ws.row_dimensions[row].height = 30
    if hints_row:
        for (col, _, _), t in zip(spec, hints):
            hint(ws, f"{col}{hints_row}", t)
        ws.row_dimensions[hints_row].height = 24


def ex_row(ws, row, cols, vals, left):
    for col, v in zip(cols, vals):
        ex(ws, f"{col}{row}", v, h="left" if col in left else "center")
    ws.row_dimensions[row].height = 30


def aviso(ws, row, lab_to, val, formula, last):
    label(ws, f"B{row}", "Aviso", merge=f"B{row}:{lab_to}{row}", h="right")
    calc(ws, f"{val}{row}", formula, b=False, merge=f"{val}{row}:{last}{row}", h="left")
    ws.row_dimensions[row].height = 21.75
    cf_ok(ws, f"{val}{row}:{last}{row}")


def data_ref(ws, r_in, lab1, lab2, val, hint1, hint2, bg):
    """Data de referência (em branco = hoje) e a data usada nas contas. Devolve a referência absoluta da data usada."""
    label(ws, f"{lab1}{r_in}", "Data de referência", merge=f"{lab1}{r_in}:{lab2}{r_in}")
    inp(ws, f"{val}{r_in}", None, h="center", bg=bg, fmt=DATE)
    dv_date(ws, f"{val}{r_in}")
    hint(ws, f"{hint1}{r_in}", "Em branco, a planilha usa a data de hoje.", merge=f"{hint1}{r_in}:{hint2}{r_in}", h="left")
    label(ws, f"{lab1}{r_in + 1}", "Data usada nas contas", merge=f"{lab1}{r_in + 1}:{lab2}{r_in + 1}")
    calc(ws, f"{val}{r_in + 1}", f'=IF({val}{r_in}="",TODAY(),{val}{r_in})', fmt=DATE)
    return f"${val}${r_in + 1}"


# ================================================================== aba Roteiro técnico
R_W = {"A": 2, "B": 5, "C": 26, "D": 9, "E": 9, "F": 9, "G": 17, "H": 28, "I": 25, "J": 20, "K": 24, "L": 26,
       "M": 9, "N": 9, "O": 15, "P": 2}
R1, RN = 14, 30


def aba_roteiro(ws, tab, data=None):
    widths(ws, R_W)
    is_ex = data is not None
    title(ws, "Roteiro técnico — o evento minuto a minuto",
          "Exemplo preenchido, para consulta. Use a aba Roteiro técnico para o seu evento." if is_ex
          else "Liste os momentos na ordem, com a duração. O início e o fim de cada momento são calculados.", "O")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Diretor técnico", "diretor", None), (8, "Versão do roteiro", "versao", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:G{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    par = data["par"] if is_ex else (time(19, 0), None, 5, 3)
    for k, (rot, dica, fmt) in enumerate([
            ("Hora de início", "A hora em que o evento começa.", HORA),
            ("Hora limite", "A hora em que o evento precisa terminar.", HORA),
            ("Folga mínima (min)", "Quanto deve sobrar até a hora limite.", None),
            ("Tolerância (min)", "Atraso que ainda conta como no horário.", None)]):
        r = 4 + k
        label(ws, f"I{r}", rot, merge=f"I{r}:J{r}")
        inp(ws, f"K{r}", par[k], h="center", bg=hb, fmt=fmt)
        hint(ws, f"L{r}", dica, merge=f"L{r}:O{r}", h="left")
    dv_time(ws, "K4:K5")
    dv_number(ws, "K6:K7")
    hint(ws, "I8", "Horas no formato 19:00. Durações e desvios em minutos.", merge="I8:O8", h="left")
    INI, LIM, FMIN, TOL = "$K$4", "$K$5", "$K$6", "$K$7"

    band(ws, 10, "Momentos do evento, na ordem", "O")
    heads(ws, 11, [("B", "#", INK), ("C", "Momento", INK), ("D", "Duração", BLUE), ("E", "Início", TEAL),
                   ("F", "Fim", TEAL), ("G", "Quem conduz", BLUE), ("H", "Som", BLUE), ("I", "Vídeo e câmeras", BLUE),
                   ("J", "Luz", BLUE), ("K", "Mídia e telão", BLUE), ("L", "Deixa de entrada", BLUE),
                   ("M", "Hora real", AMBER), ("N", "Desvio", TEAL), ("O", "Situação", TEAL)],
          12, ["", "nome usado pela equipe", "min", "calculado", "calculado", "no palco", "canais, cuidados",
               "programa", "cena", "nome exato do arquivo", "o que dispara", "anotada no evento", "min", ""])
    if is_ex:
        hint(ws, "B13", "Exemplo 1 do treinamento. As colunas Início, Fim, Desvio e Situação são calculadas.",
             merge="B13:O13", h="left")
        ws.row_dimensions[13].height = 21.75
    else:
        ex_row(ws, 13, "BCDEFGHIJKLMNO",
               ["Ex.", "Avisos, em vídeo", 4, "19:18", "19:22", "Vídeo", "Som do vídeo no canal 11", "Vídeo no programa",
                "Palco em meia-luz", "avisos.mp4, com áudio", "Último acorde da música", "19:21", "+3", "No horário"], "CGHIJKL")
        note(ws, "D11", "Duração prevista, em minutos.\nUse durações medidas em eventos anteriores, e conte as trocas de palco.")
        note(ws, "E11", "Hora de início mais a soma das durações das linhas de cima.")
        note(ws, "L11", "O sinal combinado que dispara a entrada do momento: uma frase, um gesto, o fim de uma música ou um comando.")
        note(ws, "M11", "Hora em que o momento começou de fato. Anote durante o evento, no formato 19:21.")
        note(ws, "O11", "Passa do limite: o momento termina depois da hora limite.\nAtrasado ou Adiantado: desvio maior que a tolerância.")

    R2 = R1 + RN - 1
    bg = WHITE if is_ex else INPUT
    for k in range(RN):
        r = R1 + k
        v = data["linhas"][k] if is_ex and k < len(data["linhas"]) else (None,) * 9
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", v[0], bg=bg)
        inp(ws, f"D{r}", v[1], h="center", bg=bg)
        for col, val in zip("GHIJKL", v[2:8]):
            inp(ws, f"{col}{r}", val, bg=bg)
        inp(ws, f"M{r}", v[8], h="center", bg=bg, fmt=HORA)
        calc(ws, f"E{r}", f'=IF(OR(D{r}="",{INI}=""),"",{INI}+(SUM(D${R1}:D{r})-D{r})/1440)', fmt=HORA, b=False)
        calc(ws, f"F{r}", f'=IF(E{r}="","",E{r}+D{r}/1440)', fmt=HORA, b=False)
        calc(ws, f"N{r}", f'=IF(OR(M{r}="",E{r}=""),"",ROUND((M{r}-E{r})*1440,0))', fmt=MIN)
        calc(ws, f"O{r}", f'=IF(E{r}="","",IF(AND({LIM}<>"",ROUND((F{r}-{LIM})*1440,0)>0),"Passa do limite",'
                          f'IF(N{r}="","Previsto",IF(N{r}>{TOL},"Atrasado",IF(N{r}<-{TOL},"Adiantado","No horário")))))')
        ws.row_dimensions[r].height = 30
    dv_number(ws, f"D{R1}:D{R2}")
    dv_time(ws, f"M{R1}:M{R2}")
    cf_equal(ws, f"O{R1}:O{R2}", [("No horário", GREEN), ("Atrasado", YELLOW), ("Adiantado", YELLOW),
                                  ("Passa do limite", RED)])

    s = R2 + 2
    band(ws, s, "Resumo automático", "O")
    linhas = [
        ("Momentos no roteiro", f"=COUNTA(C{R1}:C{R2})", None),
        ("Duração total (min)", f"=SUM(D{R1}:D{R2})", None),
        ("Fim previsto", f'=IF(OR({INI}="",F{s + 2}=0),"",{INI}+F{s + 2}/1440)', HORA),
        ("Folga até a hora limite (min)", f'=IF(OR(F{s + 3}="",{LIM}=""),"",ROUND(({LIM}-F{s + 3})*1440,0))', MIN),
        ("Momentos atrasados", f'=COUNTIF(O{R1}:O{R2},"Atrasado")', None),
        ("Maior desvio (min)", f'=IF(COUNT(N{R1}:N{R2})=0,"",MAX(N{R1}:N{R2}))', MIN),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", fmt=fmt, val_merge=f"F{s + k}:G{s + k}")
    a = s + len(linhas) + 1
    T, FO, AT = f"F{s + 2}", f"F{s + 4}", f"F{s + 5}"
    aviso(ws, a, "E", "F", f'=IF(OR({INI}="",{LIM}=""),"Informe a hora de início e a hora limite",'
                           f'IF({T}=0,"Liste os momentos com a duração",'
                           f'IF({FO}<0,"O roteiro passa "&-{FO}&" min da hora limite: combine com quem conduz o que sai",'
                           f'IF({FO}<{FMIN},"Folga de "&{FO}&" min, abaixo da mínima de "&{FMIN}&" min",'
                           f'IF({AT}>0,"Houve momentos atrasados: veja na coluna Desvio onde o tempo se perdeu","OK")))))', "O")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi decidido na reunião de revisão", "O")
        r = end + 3
        put(ws, f"B{r}", "#", f=font(10, True), bg=GRAY, h="center")
        put(ws, f"C{r}", "Ação", f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:J{r}")
        put(ws, f"K{r}", "Responsável", f=font(10, True), bg=GRAY, h="center", merge=f"K{r}:L{r}")
        put(ws, f"M{r}", "Prazo", f=font(10, True), bg=GRAY, h="center", merge=f"M{r}:O{r}")
        ws.row_dimensions[r].height = 21.75
        for n, (o, quem, prazo) in enumerate(data["revisao"], 1):
            r += 1
            num(ws, f"B{r}", n)
            put(ws, f"C{r}", o, merge=f"C{r}:J{r}")
            put(ws, f"K{r}", quem, merge=f"K{r}:L{r}")
            put(ws, f"M{r}", prazo, merge=f"M{r}:O{r}", h="center")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{R1}"
    setup(ws, tab, f"B1:O{end}", fit_height=True)


# ================================================================== aba Escala e cronograma
E_W = {"A": 2, "B": 5, "C": 30, "D": 20, "E": 14, "F": 10, "G": 10, "H": 20, "I": 12, "J": 16, "K": 34, "L": 2}
E1, EN = 13, 10       # escala
C1, CN = 27, 12       # cronograma
ETAPAS = [("Chegada e descarga", "Montagem", 30), ("Montagem e ligações", "Montagem", 90),
          ("Conferência de linhas", "Teste", 30), ("Passagem de som e de vídeo", "Teste", 60),
          ("Folga para imprevistos", "Folga", 15), ("Palco limpo, teste final e reunião", "Teste", 15),
          ("Portas abertas", "Portas", 30)]


def aba_escala(ws, tab):
    widths(ws, E_W)
    title(ws, "Escala e cronograma — quem faz o quê, e a que horas",
          "O cronograma é contado para trás, a partir da hora do evento. As células cinza são calculadas.", "K")
    for r, rot, fmt in [(4, "Evento", None), (5, "Local", None), (6, "Data", DATE), (7, "Diretor técnico", None)]:
        label(ws, f"B{r}", rot)
        inp(ws, f"C{r}", None, merge=f"C{r}:E{r}", fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "C6")
    for k, (rot, v, fmt, dica) in enumerate([
            ("Hora do evento", time(19, 0), HORA, "A hora em que o evento começa."),
            ("Liberação do local", None, HORA, "A partir de que horas o espaço está disponível."),
            ("Tolerância (min)", 5, None, "Atraso de etapa que ainda conta como OK.")]):
        r = 4 + k
        label(ws, f"G{r}", rot, merge=f"G{r}:H{r}")
        inp(ws, f"I{r}", v, h="center", fmt=fmt)
        hint(ws, f"J{r}", dica, merge=f"J{r}:K{r}", h="left")
    dv_time(ws, "I4:I5")
    dv_number(ws, "I6")
    EV, LIB, TOL = "$I$4", "$I$5", "$I$6"

    # ---- escala
    band(ws, 9, "Escala da equipe", "K")
    heads(ws, 10, [("B", "#", INK), ("C", "Função", INK), ("D", "Nome", BLUE), ("E", "Contato", BLUE),
                   ("F", "Chegada", BLUE), ("G", "Posto", BLUE), ("H", "Reserva", BLUE), ("I", "Confirmado", BLUE),
                   ("J", "Situação", TEAL), ("K", "Observação", INK)],
          11, ["", "", "", "telefone", "hora", "onde fica", "quem cobre", "Sim / Não", "", ""])
    ex_row(ws, 12, "BCDEFGHIJK", ["Ex.", "Operador de som", "Nome da pessoa", "(00) 00000-0000", "14:30", "Mesa de som",
                                  "Nome do reserva", "Sim", "OK", ""], "CDGHK")
    note(ws, "I10", "A pessoa confirmou presença na véspera?")
    note(ws, "J10", "Sem responsável: a função está sem nome.\nA confirmar: o nome existe, mas a presença não foi confirmada.")
    E2 = E1 + EN - 1
    for k in range(EN):
        r = E1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDEGHK":
            inp(ws, f"{col}{r}")
        inp(ws, f"F{r}", h="center", fmt=HORA)
        inp(ws, f"I{r}", h="center")
        calc(ws, f"J{r}", f'=IF(C{r}="","",IF(D{r}="","Sem responsável",IF(I{r}<>"Sim","A confirmar","OK")))')
        ws.row_dimensions[r].height = 21.75
    dv_time(ws, f"F{E1}:F{E2}")
    dv_list(ws, f"I{E1}:I{E2}", SIM_NAO, "A presença foi confirmada?")
    cf_equal(ws, f"J{E1}:J{E2}", [("OK", GREEN), ("A confirmar", YELLOW), ("Sem responsável", YELLOW)])

    # ---- cronograma
    band(ws, 23, "Cronograma de montagem, contado para trás", "K")
    heads(ws, 24, [("B", "#", INK), ("C", "Etapa", INK), ("D", "Tipo", BLUE), ("E", "Duração", BLUE),
                   ("F", "Início", TEAL), ("G", "Fim", TEAL), ("H", "Responsável", BLUE), ("I", "Concluída às", AMBER),
                   ("J", "Situação", TEAL), ("K", "Observação", INK)],
          25, ["na ordem", "lista", "min", "calculado", "calculado", "", "hora real", "", ""])
    hint(ws, "B26", "As etapas e as durações abaixo são as do exemplo da figura 4 do treinamento. Troque pelas da sua equipe. "
                    "Etapas feitas ao mesmo tempo por pessoas diferentes entram uma vez só, com a duração da mais longa.",
         merge="B26:K26", h="left")
    ws.row_dimensions[26].height = 30
    note(ws, "D24", "Folga: reserve sempre uma etapa de folga, com hora marcada.")
    note(ws, "F24", "Hora do evento menos a soma das durações desta etapa e das de baixo.")
    C2 = C1 + CN - 1
    for k in range(CN):
        r = C1 + k
        v = ETAPAS[k] if k < len(ETAPAS) else (None, None, None)
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", v[0])
        inp(ws, f"D{r}", v[1], h="center")
        inp(ws, f"E{r}", v[2], h="center")
        inp(ws, f"H{r}")
        inp(ws, f"I{r}", h="center", fmt=HORA)
        inp(ws, f"K{r}")
        calc(ws, f"F{r}", f'=IF(OR(E{r}="",{EV}=""),"",{EV}-SUM(E{r}:E${C2})/1440)', fmt=HORA, b=False)
        calc(ws, f"G{r}", f'=IF(F{r}="","",F{r}+E{r}/1440)', fmt=HORA, b=False)
        calc(ws, f"J{r}", f'=IF(OR(F{r}="",I{r}=""),"",IF(ROUND((I{r}-G{r})*1440,0)>{TOL},"Atrasada","OK"))')
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{C1}:D{C2}", ["Montagem", "Teste", "Folga", "Portas"], "Montagem, Teste, Folga ou Portas.")
    dv_number(ws, f"E{C1}:E{C2}")
    dv_time(ws, f"I{C1}:I{C2}")
    cf_equal(ws, f"J{C1}:J{C2}", [("OK", GREEN), ("Atrasada", YELLOW)])

    s = C2 + 2
    band(ws, s, "Resumo automático", "K")
    linhas = [
        ("Tempo total, da chegada ao início (h:mm)", f'=IF(SUM(E{C1}:E{C2})=0,"",SUM(E{C1}:E{C2})/1440)', "[h]:mm"),
        ("Hora de chegada da equipe", f'=IF(OR({EV}="",SUM(E{C1}:E{C2})=0),"",{EV}-SUM(E{C1}:E{C2})/1440)', HORA),
        ("Folga reservada (min)", f'=SUMIF(D{C1}:D{C2},"Folga",E{C1}:E{C2})', None),
        ("Etapas atrasadas", f'=COUNTIF(J{C1}:J{C2},"Atrasada")', None),
        ("Funções sem responsável ou a confirmar", f'=COUNTIF(J{E1}:J{E2},"Sem responsável")+COUNTIF(J{E1}:J{E2},"A confirmar")', None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "D", "E", fmt=fmt)
    a = s + len(linhas) + 1
    CH, FG, ATR, FN = f"E{s + 2}", f"E{s + 3}", f"E{s + 4}", f"E{s + 5}"
    aviso(ws, a, "D", "E", f'=IF({EV}="","Informe a hora do evento",IF({CH}="","Liste as etapas com a duração",'
                           f'IF(AND({LIB}<>"",ROUND(({LIB}-{CH})*1440,0)>0),"O local só libera "&ROUND(({LIB}-{CH})*1440,0)&'
                           f'" min depois da chegada calculada: encurte etapas ou combine outra liberação",'
                           f'IF({FG}=0,"Reserve uma etapa de folga",IF(COUNTA(C{E1}:C{E2})=0,"Preencha a escala",'
                           f'IF({FN}>0,"Há função sem responsável ou sem confirmação",'
                           f'IF({ATR}>0,"Houve etapa atrasada: anote a causa para o próximo cronograma","OK")))))))', "K")
    ws.freeze_panes = "C4"
    setup(ws, tab, f"B1:K{a}", fit_height=True)


# ================================================================== aba Inventário
I_W = {"A": 2, "B": 5, "C": 10, "D": 30, "E": 12, "F": 7, "G": 20, "H": 13, "I": 13, "J": 11, "K": 13, "L": 13, "M": 32, "N": 2}
I1, IN = 13, 40
SIT_INV = [("Em dia", GREEN), ("Vence logo", YELLOW), ("Sem data", YELLOW), ("Vencida", RED), ("Fora de uso", RED)]


def aba_inventario(ws, tab):
    widths(ws, I_W)
    title(ws, "Inventário — equipamentos, estado e manutenção",
          "Um item por linha. A próxima manutenção e a situação são calculadas a partir da última data e do intervalo.", "M")
    for r, rot in [(4, "Equipe ou local"), (5, "Responsável pelo inventário")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:D{r}")
        inp(ws, f"E{r}", None, merge=f"E{r}:H{r}")
        ws.row_dimensions[r].height = 21.75
    REF = data_ref(ws, 4, "J", "K", "L", "M", "M", INPUT)
    hint(ws, "J6", "Vence logo: a próxima manutenção cai nos próximos 30 dias.", merge="J6:M6", h="left")

    band(ws, 9, "Equipamentos", "M")
    heads(ws, 10, [("B", "#", INK), ("C", "Código", BLUE), ("D", "Item", BLUE), ("E", "Categoria", BLUE),
                   ("F", "Qtde", BLUE), ("G", "Local de guarda", BLUE), ("H", "Estado", BLUE),
                   ("I", "Última manutenção", AMBER), ("J", "Intervalo", AMBER), ("K", "Próxima", TEAL),
                   ("L", "Situação", TEAL), ("M", "Observação", INK)],
          11, ["", "etiqueta", "descrição", "lista", "", "caixa, armário", "lista", "data", "meses", "calculada", "", ""])
    ex_row(ws, 12, "BCDEFGHIJKLM", ["Ex.", "C-12", "Cabo XLR de 10 m", "Cabos", 1, "Caixa de cabos 2", "Bom",
                                    "01/09/2026", 1, "01/10/2026", "Vence logo", "Testado com o testador de cabos"], "DGM")
    note(ws, "C10", "Código curto, igual ao da etiqueta no equipamento: C-12, MIC-03.\nUse o mesmo código nas ocorrências.")
    note(ws, "J10", "De quantos em quantos meses o item é revisado ou testado.\nCabos: 1. Tripés e suportes: 6. Quadro e tomadas: 12.")
    note(ws, "L10", "Fora de uso: item com defeito ou em conserto.\nSem data: falta a última manutenção ou o intervalo.")
    I2 = I1 + IN - 1
    for k in range(IN):
        r = I1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDGM":
            inp(ws, f"{col}{r}")
        for col in "EFHJ":
            inp(ws, f"{col}{r}", h="center")
        inp(ws, f"I{r}", h="center", fmt=DATE)
        calc(ws, f"K{r}", f'=IF(OR(I{r}="",J{r}=""),"",DATE(YEAR(I{r}),MONTH(I{r})+J{r},DAY(I{r})))', fmt=DATE, b=False)
        calc(ws, f"L{r}", f'=IF(D{r}="","",IF(OR(H{r}="Com defeito",H{r}="Em conserto"),"Fora de uso",IF(K{r}="","Sem data",'
                          f'IF(K{r}<{REF},"Vencida",IF(K{r}-{REF}<=30,"Vence logo","Em dia")))))')
        ws.row_dimensions[r].height = 19.5
    dv_list(ws, f"E{I1}:E{I2}", ["Som", "Vídeo", "Luz", "Mídia", "Energia", "Cabos", "Estrutura", "Outros"], "Escolha a categoria.")
    dv_list(ws, f"H{I1}:H{I2}", ["Bom", "Regular", "Com defeito", "Em conserto"], "Bom, Regular, Com defeito ou Em conserto.")
    dv_date(ws, f"I{I1}:I{I2}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="120", allow_blank=True)
    dv.errorTitle, dv.error = "Valor inválido", "Digite o intervalo em meses, de 1 a 120."
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(f"J{I1}:J{I2}")
    dv_number(ws, f"F{I1}:F{I2}")
    cf_equal(ws, f"L{I1}:L{I2}", SIT_INV)

    s = I2 + 2
    band(ws, s, "Resumo automático", "M")
    cont = lambda t: f'COUNTIF(L{I1}:L{I2},"{t}")'
    linhas = [("Itens cadastrados", f"=COUNTA(D{I1}:D{I2})"), ("Quantidade total", f"=SUM(F{I1}:F{I2})"),
              ("Em dia", f"={cont('Em dia')}"), ("Vence nos próximos 30 dias", f"={cont('Vence logo')}"),
              ("Manutenção vencida", f"={cont('Vencida')}"), ("Fora de uso", f"={cont('Fora de uso')}"),
              ("Sem data", f"={cont('Sem data')}")]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    aviso(ws, a, "E", "F", f'=IF(F{s + 1}=0,"Cadastre os itens",IF(F{s + 5}>0,"Há manutenção vencida: faça antes do próximo evento",'
                           f'IF(F{s + 6}>0,"Há itens fora de uso: não conte com eles no evento",'
                           f'IF(F{s + 7}>0,"Há itens sem data de manutenção","OK"))))', "M")
    ws.freeze_panes = f"E{I1}"
    setup(ws, tab, f"B1:M{a}")


# ================================================================== aba Ocorrências e revisão
O_W = {"A": 2, "B": 5, "C": 8, "D": 18, "E": 12, "F": 15, "G": 30, "H": 26, "I": 26, "J": 32, "K": 17, "L": 12,
       "M": 11, "N": 16, "O": 2}
O1, ON = 13, 15
ALERTAS = [("OK", GREEN), ("Defina a ação", YELLOW), ("Sem responsável", YELLOW), ("Sem prazo", YELLOW), ("Atrasada", RED)]


def aba_ocorrencias(ws, tab):
    widths(ws, O_W)
    title(ws, "Ocorrências e revisão — do que aconteceu ao que muda",
          "Anote no evento, com a hora. Na reunião de revisão, complete a causa, a ação, o responsável e o prazo.", "N")
    for r, rot, fmt in [(4, "Evento", None), (5, "Data do evento", DATE), (6, "Diretor técnico", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:D{r}")
        inp(ws, f"E{r}", None, merge=f"E{r}:G{r}", fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "E5")
    REF = data_ref(ws, 4, "I", "I", "J", "K", "N", INPUT)

    band(ws, 9, "Registro de ocorrências", "N")
    heads(ws, 10, [("B", "#", INK), ("C", "Hora", AMBER), ("D", "Momento", AMBER), ("E", "Área", AMBER),
                   ("F", "Gravidade", AMBER), ("G", "O que aconteceu", AMBER), ("H", "Causa provável", BLUE),
                   ("I", "O que foi feito na hora", AMBER), ("J", "Ação para o próximo evento", BLUE),
                   ("K", "Responsável", BLUE), ("L", "Prazo", BLUE), ("M", "Ação", BLUE), ("N", "Alerta", TEAL)],
          11, ["", "hh:mm", "do roteiro", "lista", "lista", "o sintoma, sem culpado", "no processo", "plano B usado",
               "o que muda", "um nome", "data", "Aberta / Feita", ""])
    ex_row(ws, 12, "BCDEFGHIJKLMN", ["Ex.", "19:21", "Avisos", "Mídia", "Público percebeu", "Vídeo de avisos entrou sem som",
                                     "O roteiro não dizia que o vídeo tinha áudio", "Canal 11 aberto 20 s depois",
                                     "Coluna “tem áudio?” na lista de mídias e teste na passagem", "Operador de mídia",
                                     "10/10/2026", "Aberta", "No prazo"], "DGHIJK")
    note(ws, "F10", "Público percebeu: a falha chegou à sala ou à transmissão.\nQuase falha: foi notada a tempo.\nInterna: só a equipe percebeu.")
    note(ws, "H10", "Procure a causa no processo (roteiro, checklist, equipamento), e não um culpado.")
    note(ws, "N10", "Defina a ação, Sem responsável, Sem prazo: a ocorrência ainda não virou mudança.\nAtrasada: o prazo passou e a ação não foi feita.")
    O2 = O1 + ON - 1
    for k in range(ON):
        r = O1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center", fmt=HORA)
        for col in "DGHIJK":
            inp(ws, f"{col}{r}")
        for col in "EFM":
            inp(ws, f"{col}{r}", h="center")
        inp(ws, f"L{r}", h="center", fmt=DATE)
        calc(ws, f"N{r}", f'=IF(AND(G{r}="",J{r}=""),"",IF(J{r}="","Defina a ação",IF(K{r}="","Sem responsável",'
                          f'IF(L{r}="","Sem prazo",IF(M{r}="Feita","OK",IF(L{r}<{REF},"Atrasada","No prazo"))))))')
        ws.row_dimensions[r].height = 30
    dv_time(ws, f"C{O1}:C{O2}")
    dv_list(ws, f"E{O1}:E{O2}", ["Som", "Vídeo", "Luz", "Mídia", "Transmissão", "Energia", "Equipe", "Outros"], "Escolha a área.")
    dv_list(ws, f"F{O1}:F{O2}", ["Público percebeu", "Quase falha", "Interna"], "Público percebeu, Quase falha ou Interna.")
    dv_date(ws, f"L{O1}:L{O2}")
    dv_list(ws, f"M{O1}:M{O2}", ["Aberta", "Feita"], "A ação já foi feita?")
    cf_equal(ws, f"N{O1}:N{O2}", ALERTAS)

    s = O2 + 2
    band(ws, s, "Resumo automático", "N")
    linhas = [("Ocorrências registradas", f"=COUNTA(G{O1}:G{O2})"),
              ("Percebidas pelo público", f'=COUNTIF(F{O1}:F{O2},"Público percebeu")'),
              ("Ações abertas", f'=COUNTIF(N{O1}:N{O2},"No prazo")+COUNTIF(N{O1}:N{O2},"Atrasada")'),
              ("Ações atrasadas", f'=COUNTIF(N{O1}:N{O2},"Atrasada")'),
              ("Ocorrências sem ação, responsável ou prazo",
               f'=COUNTIF(N{O1}:N{O2},"Defina a ação")+COUNTIF(N{O1}:N{O2},"Sem responsável")+COUNTIF(N{O1}:N{O2},"Sem prazo")')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    aviso(ws, a, "E", "F", f'=IF(F{s + 1}=0,"Nenhuma ocorrência registrada",IF(F{s + 5}>0,"Há ocorrência que ainda não virou ação com responsável e prazo",'
                           f'IF(F{s + 4}>0,"Há ação atrasada: cobre na próxima reunião","OK")))', "N")

    r = a + 2
    band(ws, r, "Reunião de revisão: as três perguntas", "N")
    for t in ["O que funcionou e deve ser mantido?", "O que falhou, ou quase falhou?", "O que muda no próximo evento?"]:
        r += 1
        label(ws, f"B{r}", t, merge=f"B{r}:E{r}")
        inp(ws, f"F{r}", None, merge=f"F{r}:N{r}")
        ws.row_dimensions[r].height = 48
    ws.freeze_panes = f"D{O1}"
    setup(ws, tab, f"B1:N{r}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Operação do evento — roteiro, equipe e ocorrências"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Operação do evento — como usar esta planilha",
           "Roteiro técnico, escala e cronograma de montagem, inventário e registro de ocorrências, do antes ao depois do evento.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, horas calculadas, situações, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Horas e datas", "Horas no formato 19:00; datas no formato dia/mês/ano. Durações em minutos, só o número."),
        ("Listas suspensas", "Tipo de etapa, estado, categoria, área, gravidade e situação aceitam só as opções da lista."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Dias antes. Aba Roteiro técnico: hora de início, hora limite e os momentos, na ordem, com a duração e a deixa."),
        ("2", "Veja a folga até a hora limite. Se o roteiro não couber, combine com quem conduz o que sai."),
        ("3", "Aba Escala e cronograma: a hora do evento, a escala e as etapas de montagem com as durações da sua equipe."),
        ("4", "Aba Inventário: confira se o que vai ao evento está em dia e se nada está fora de uso."),
        ("5", "No evento. Anote no Roteiro técnico a hora real de cada momento, e nas Ocorrências tudo o que falhar."),
        ("6", "Depois. Na reunião de revisão, complete em Ocorrências a causa, a ação, o responsável e o prazo."),
        ("7", "Atualize o Inventário (estado e manutenção) e marque o Checklist."),
    ]),
    ("Abas da planilha", [
        ("Roteiro técnico", "Até 30 momentos, com início e fim calculados, folga até a hora limite e desvio da hora real."),
        ("Escala e cronograma", "Escala de até 10 funções e até 12 etapas de montagem, contadas para trás a partir do evento."),
        ("Inventário", "Até 40 itens, com a próxima manutenção e a situação de cada um."),
        ("Ocorrências e revisão", "Até 15 ocorrências, com alerta das ações sem dono, sem prazo ou atrasadas, e as três perguntas da revisão."),
        ("Referência", "Sintomas e causas prováveis, capacidade dos circuitos, durações de montagem, kit e ordem de ligar."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Roteiro", "O roteiro do culto do exemplo 1 do treinamento, com as horas reais e o que a revisão decidiu."),
    ]),
    ("Regras e premissas do modelo", [
        ("Início de cada momento", "Hora de início do evento mais a soma das durações das linhas de cima. Linha sem duração fica sem hora."),
        ("Situação no roteiro", "Passa do limite: termina depois da hora limite. Atrasado ou Adiantado: desvio da hora real maior que a tolerância."),
        ("Cronograma para trás", "Cada etapa começa na hora do evento menos a soma da sua duração e das de baixo. Liste na ordem em que acontecem."),
        ("Inventário", "Próxima manutenção = última + intervalo em meses. Vence logo: nos próximos 30 dias. Item com defeito ou em conserto fica Fora de uso."),
        ("Data de referência", "Em branco, as abas Inventário e Ocorrências usam a data de hoje para dizer o que venceu."),
        ("Valores de referência", "Durações típicas, consumo e limite de 80% do circuito são ponto de partida, e não norma. Confira a etiqueta dos equipamentos e o quadro do local."),
        ("Exemplos", "Os dados das linhas de exemplo e da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_roteiro(wb.create_sheet("Roteiro técnico"), BLUE)
aba_escala(wb.create_sheet("Escala e cronograma"), AMBER)
aba_inventario(wb.create_sheet("Inventário"), TEAL)
aba_ocorrencias(wb.create_sheet("Ocorrências e revisão"), HUE["red"][0])

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 44, "D": 30, "E": 30, "F": 2})
title(ws, "Referência — consulta rápida", "Valores de referência, para servir de ponto de partida. Confira o manual e a etiqueta do seu equipamento.", "E")


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
            put(ws, f"{col}{r}", t, h="left")
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(4, "Sintomas e causas prováveis", ["Sintoma", "Causas mais prováveis, na ordem", "Primeiro teste", "Estudo para rever"], [
    ("Um canal sem som", "Mute; fader ou roteamento; cabo; phantom desligado em condensador ou DI ativa; pilha do sem fio.", "Medidor do canal no PFL.", "Mesa de som; Cabos"),
    ("Tudo sem som", "Master ou saída em mute; amplificador ou caixas desligados; cabo de saída; disjuntor.", "Medidor do master.", "Caixas e amplificação"),
    ("Zumbido grave contínuo", "Laço de terra entre aparelhos em tomadas diferentes; cabo desbalanceado longo; sinal junto de energia ou dimmer.", "Feche canal por canal até sumir.", "Cabos e conexões"),
    ("Chiado constante", "Ganho baixo compensado adiante; caixa ativa aberta demais; sem fio com sinal fraco.", "Feche o canal: o chiado some?", "Ganho e fader"),
    ("Estalos que vão e voltam", "Mau contato em cabo ou conector; sem fio com interferência ou pilha fraca.", "Mexa de leve no cabo, com o canal baixo.", "Cabos; Microfones sem fio"),
    ("Distorção", "Ganho alto; fonte de linha em entrada de microfone; ganho alto no transmissor; sistema no limite.", "Luz de clip do canal.", "Ganho e fader"),
    ("Apito", "Microfone de frente para a caixa; retorno alto; muitos microfones abertos.", "Baixe o canal ou o retorno suspeito.", "Microfonia"),
    ("Câmera “sem sinal” no switcher", "Cabo solto ou com defeito; formato diferente do sistema; conversor sem energia.", "Câmera direto em um monitor.", "Sinais de vídeo"),
    ("Imagem pisca ou cai", "Cabo longo demais ou ruim; conector frouxo; adaptador.", "Troque por um cabo curto e bom.", "Sinais de vídeo"),
    ("Telão sem imagem", "Entrada errada no projetor; segunda tela desativada; cabo; resolução não aceita.", "O projetor diz “sem sinal”?", "Projeção e mídia"),
    ("Vídeo toca sem som", "Saída de áudio do computador errada; volume do computador; canal da mesa fechado.", "Medidor do canal do computador.", "Projeção e mídia"),
    ("Faixas ou cintilação na imagem", "Obturador em desacordo com as lâmpadas ou o telão de LED.", "Mude o obturador e observe.", "Exposição"),
    ("Transmissão travando", "Upload insuficiente ou instável; bitrate alto; Wi-Fi; computador sobrecarregado.", "Quadros perdidos no codificador.", "Transmissão ao vivo"),
    ("Transmissão sem áudio", "Fonte de áudio errada no codificador; saída da mesa em mute.", "Medidor de áudio do codificador.", "Transmissão ao vivo"),
])
r = tabela(r, "Capacidade dos circuitos (corrente × tensão)", ["Disjuntor ou elo", "Capacidade em 127 V e em 220 V", "Até 80%, em 127 V e em 220 V", "Onde aparece"], [
    ("10 A", "1.270 W · 2.200 W", "1.016 W · 1.760 W", "Tomada comum, a maioria das réguas e extensões."),
    ("16 A", "2.032 W · 3.520 W", "1.626 W · 2.816 W", "Disjuntor comum de circuitos de tomadas."),
    ("20 A", "2.540 W · 4.400 W", "2.032 W · 3.520 W", "Tomada de 20 A, circuitos de maior carga."),
    ("25 A", "3.175 W · 5.500 W", "2.540 W · 4.400 W", "Circuitos dedicados, conforme o projeto."),
], alt=21.75)
r = tabela(r, "Durações típicas de montagem", ["Etapa", "Do que depende", "Duração típica", "Observação"], [
    ("Chegada, descarga e posicionamento", "Distância do carro ao palco, escadas, volume de carga.", "20 a 40 min", ""),
    ("Montagem de som, de vídeo ou de luz", "Número de linhas e de câmeras, refletores a pendurar.", "45 a 120 min", "Feitas ao mesmo tempo, entra a mais longa."),
    ("Conferência de linhas", "Número de canais.", "15 a 30 min", "Sem os músicos."),
    ("Passagem de som e de vídeo", "Banda ou só fala.", "60 a 90 min", "Só com fala: 20 a 30 min."),
    ("Folga para imprevistos", "Local, equipe ou equipamento novos.", "15 a 30 min", "É uma etapa, com hora marcada."),
    ("Palco limpo, teste final e reunião", "", "15 min", "Termina quando as portas abrem."),
    ("Portas abertas até o início", "Combinado com a organização.", "20 a 30 min", ""),
], alt=21.75)
r = tabela(r, "Ordem de ligar e de desligar", ["Ligar, nesta ordem", "Desligar, nesta ordem", "Cuidado", ""], [
    ("1. Energia: disjuntores, no-break, réguas", "1. Amplificadores e caixas ativas", "O que amplifica liga por último e desliga primeiro.", ""),
    ("2. Fontes e periféricos", "2. Mesa e switcher (salve a cena)", "Espere cada etapa terminar de iniciar.", ""),
    ("3. Mesa e switcher, master fechado", "3. Fontes e periféricos (computador pelo sistema)", "Projetor de lâmpada: espere a ventoinha parar.", ""),
    ("4. Amplificadores e caixas ativas", "4. Energia", "Se a energia cair, desligue as caixas antes que ela volte.", ""),
], alt=21.75)
r = tabela(r, "Kit de emergência", ["Item", "Para quê", "Quantidade sugerida", "Conferir"], [
    ("Microfone com fio e cabo XLR longo", "Plano B do microfone de quem fala.", "1", "Pronto perto do palco."),
    ("Cabos XLR, P10, HDMI (e SDI, se usado)", "Trocar cabo com defeito.", "2 ou 3 de cada", "Testados no mês."),
    ("Direct box e cabo P2", "Ligar celular ou computador de última hora.", "1", ""),
    ("Pilhas ou baterias", "Sem fio.", "Um jogo novo por sistema, mais um", "Validade e carga."),
    ("Adaptadores, extensão, régua, fontes", "Energia e conexões.", "1 de cada", "Capacidade da régua e da extensão."),
    ("Mídias em pen drive e segundo computador", "Plano B da mídia.", "1", "Arquivos da versão final."),
    ("Testador de cabos e de tomadas, lanterna", "Diagnóstico.", "1 de cada", "Pilhas do testador."),
    ("Fita de palco, fita crepe, caneta, abraçadeiras", "Prender e identificar.", "1 de cada", ""),
    ("Etiquetas “com defeito”", "Tirar de uso o que falhou.", "10", ""),
    ("Roteiro impresso e lista de contatos", "Se a tela falhar.", "1 por posto", "Versão final."),
], alt=21.75)
hint(ws, f"B{r - 1}", "Faixas típicas de eventos pequenos e médios. O limite de 80% do circuito é regra prática deste material, e não norma.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}")

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist da operação do evento",
          "Os oito primeiros antes de abrir as portas; os outros, durante e depois. Escolha a situação na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_roteiro(wb.create_sheet("Exemplo - Roteiro"), MUTED, EX_ROT)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
