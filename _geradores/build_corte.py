# -*- coding: utf-8 -*-
"""Gera Corte-ao-Vivo-modelo.xlsx: roteiro de corte com câmeras e gravação, equipe, registro de erros,
referência, checklist e o exemplo 1 do treinamento."""
import os, sys
from datetime import time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

TEAL, TEAL_T = HUE["teal"]
RED_H, RED_T = HUE["red"]
SLATE = HUE["slate"][0]
SIM_NAO = ["Sim", "Não"]
HORA = "hh:mm"

CHECK = [
    ("Preparação", "O roteiro de corte está pronto, com fonte principal e alternativa em cada momento."),
    ("Equipe", "Cada função da escala tem um nome, e todos conhecem as palavras de comando."),
    ("Equipe", "Cada câmera tem um plano de base, e a câmera geral é o plano de segurança."),
    ("Sistema", "Todas as fontes aparecem no multiview, no formato do sistema e com o nome certo."),
    ("Sistema", "As câmeras foram igualadas em exposição e em cor, com a luz do evento."),
    ("Sistema", "O tally e o intercom foram testados com cada operador de câmera."),
    ("Gráficos e mídia", "As legendas, as letras e os vídeos foram testados no preview, com a grafia dos nomes conferida."),
    ("Gráficos e mídia", "O som de cada vídeo foi combinado com o operador de áudio."),
    ("Switcher", "O tempo da fusão está ajustado, e o FTB e o key final começam como combinado."),
    ("Switcher", "O telão da sala e a transmissão recebem, cada um, a saída certa."),
    ("Gravação", "As gravações foram iniciadas, com espaço conferido, e os contadores estão andando."),
    ("Revisão", "Os erros do evento foram anotados com a hora, para a reunião de revisão."),
]

TIPOS = ["Abertura", "Louvor rápido", "Louvor lento", "Oração", "Leitura", "Pregação", "Avisos", "Mídia",
         "Oferta ou ceia", "Entrevista", "Encerramento"]
FONTES = ["Câmera 1", "Câmera 2", "Câmera 3", "Câmera 4", "Câmera 5", "Câmera 6", "Mídia", "Slides", "Letras", "Vinheta"]
TRANS = ["Corte seco", "Fusão", "Do preto", "Para o preto"]
PEND = ["Informe a duração", "Defina a fonte principal", "Defina o áudio da mídia", "Alternativa repetida",
        "Sem alternativa", "Câmera sem operador", "Alternativa sem operador"]
CAMS = ["Câmera 1", "Câmera 2", "Câmera 3", "Câmera 4", "Câmera 5", "Câmera 6"]
GRAV = ["Programa (PGM)", "Programa limpo"]

# Exemplo 1 do treinamento: o roteiro de um culto com três câmeras, como estava na quinta-feira
EX_ROTEIRO = {
    "head": {"evento": "Culto de domingo", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "hora": time(19, 0), "diretor": "Diretor de imagem da escala"},
    "par": (5, 30, 0.6),
    "momentos": [  # momento, tipo, duração, principal, alt 1, alt 2, gráfico, transição, áudio da mídia
        ("Abertura e acolhida", "Abertura", 3, "Câmera 1", "Câmera 3", None, "Logo da igreja", "Do preto", None),
        ("Louvor", "Louvor rápido", 21, "Câmera 2", "Câmera 1", "Câmera 3", "Letra", "Fusão", None),
        ("Oração", "Oração", 3, "Câmera 1", "Câmera 3", None, None, "Fusão", None),
        ("Avisos em vídeo", "Mídia", 4, "Mídia", "Câmera 1", None, None, "Corte seco", None),
        ("Leitura bíblica", "Leitura", 4, "Câmera 2", "Câmera 2", None, "Texto do versículo", "Corte seco", None),
        ("Pregação", "Pregação", 40, "Câmera 2", None, None, "Nome do pastor e versículos", "Corte seco", None),
        ("Oferta", "Oferta ou ceia", 2, "Câmera 3", "Câmera 1", None, None, "Fusão", None),
        ("Louvor final", "Louvor lento", 5, "Câmera 4", "Câmera 1", None, "Letra", "Fusão", None),
        ("Bênção e encerramento", "Encerramento", 2, "Câmera 1", None, None, "Tela de encerramento", "Para o preto", None),
    ],
    "cams": [  # posição, plano de base, operador, grava, taxa, livre
        ("Fundo, centro do corredor", "Geral do palco", "Voluntário A", "Sim", 50, 64),
        ("Fundo, à esquerda", "Médio do púlpito", "Voluntário B", "Sim", 50, 32),
        ("Frente, à esquerda do palco", "Banda e congregação", "Voluntário C", "Não", None, None),
        ("Frente, à direita do palco", "Close de quem canta", None, "Não", None, None),
        (None, None, None, None, None, None),
        (None, None, None, None, None, None),
    ],
    "grav": [("Gravador do switcher", "Sim", 20, 256), (None, "Não", None, None)],
    "ajustes": [  # o quê, o que a planilha apontou, o que foi feito
        ("Avisos em vídeo", "Defina o áudio da mídia", "O vídeo tem som: marcado “Sim”, e o operador de áudio foi avisado."),
        ("Leitura bíblica", "Alternativa repetida", "A alternativa também era a câmera 2. Trocada pela câmera 1, em plano geral."),
        ("Pregação", "Sem alternativa", "Alternativas definidas: câmera 1, geral, e câmera 3, congregação."),
        ("Louvor final", "Câmera sem operador", "A câmera 4 não tinha operador neste domingo. O momento passou para a câmera 3."),
        ("Câmera 2", "Falta espaço (42,8 GB para 32 GB livres)", "Cartão trocado por um maior antes do culto."),
        ("Câmera 2", "Concentrada: 77% do tempo como principal", "Combinado que a câmera 1 faz um plano médio do púlpito, como plano B."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


# ================================================================== aba Roteiro de corte
R_W = {"A": 2, "B": 5, "C": 26, "D": 16, "E": 9, "F": 9, "G": 13, "H": 13, "I": 13, "J": 20, "K": 13, "L": 10,
       "M": 26, "N": 24, "O": 30, "P": 2}
N = 20
R1 = 14
R2 = R1 + N - 1
CAM1 = R2 + 5          # primeira linha da tabela de câmeras
CAM_END = CAM1 + len(CAMS) - 1
GRAV1 = CAM_END + 1
GRAV_END = GRAV1 + len(GRAV) - 1
LOOK = "'Referência'!$B$6:$D$16"   # tabela de ritmo da aba Referência


def aba_roteiro(ws, tab, data=None):
    widths(ws, R_W)
    is_ex = data is not None
    title(ws, "Roteiro de corte — momento a momento",
          "Exemplo preenchido (exemplo 1 do treinamento), para consulta. Use a aba Roteiro de corte para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro. As células cinza são calculadas e apontam o que falta resolver.", "O")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Hora de início", "hora", HORA), (8, "Diretor de imagem", "diretor", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt, h="center" if k == "hora" else "left")
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    par = data["par"] if is_ex else (5, 30, 0.6)
    pars = [("Momento sem alternativa: alerta acima de (min)", par[0], None, "Regra prática deste material. Momento curto pode ficar só na câmera geral."),
            ("Margem de gravação (min)", par[1], None, "Tempo a mais que cada gravador deve suportar, além do roteiro."),
            ("Concentração numa câmera: alerta acima de", par[2], "0%", "Parcela do tempo como principal. Acima dela, combine um plano B.")]
    for k, (rot, v, fmt, dica) in enumerate(pars):
        r = 4 + k
        label(ws, f"I{r}", rot, merge=f"I{r}:K{r}")
        inp(ws, f"L{r}", v, h="center", bg=hb, fmt=fmt)
        hint(ws, f"M{r}", dica, merge=f"M{r}:O{r}", h="left")
        ws.row_dimensions[r].height = 30 if k == 0 else 21.75
    dv_number(ws, "L4:L6")
    LIM, MARG, CONC = "$L$4", "$L$5", "$L$6"
    HORA0 = "$D$7"
    hint(ws, "I7", "A hora de cada momento é calculada a partir da hora de início e das durações.", merge="I7:O7", h="left")
    hint(ws, "I8", "Fontes: Câmera 1 a 6, Mídia, Slides, Letras ou Vinheta. A pendência some quando a linha está resolvida.", merge="I8:O8", h="left")

    band(ws, 10, "Roteiro de corte, na ordem do evento", "O")
    heads = [("B", "#", INK), ("C", "Momento", INK), ("D", "Tipo", INK), ("E", "Duração", INK), ("F", "Início", TEAL),
             ("G", "Fonte principal", RED_H), ("H", "Alternativa 1", RED_H), ("I", "Alternativa 2", RED_H),
             ("J", "Gráfico", INK), ("K", "Transição de entrada", INK), ("L", "Áudio da mídia", INK),
             ("M", "Ritmo sugerido", TEAL), ("N", "Pendência", TEAL), ("O", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMNO", ["", "o que acontece no palco", "da lista", "min", "calculado", "da lista",
                                         "da lista", "da lista", "legenda, letra, texto", "da lista", "Sim / Não",
                                         "da aba Referência", "calculada", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 21.75
    if is_ex:
        hint(ws, "B13", "Roteiro como estava na quinta-feira, antes das correções. As colunas F, M e N são calculadas.",
             merge="B13:O13", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNO", ["Ex.", "Pregação", "Pregação", 40, "19:40", "Câmera 2", "Câmera 1", "Câmera 3",
                                             "Nome do pastor", "Corte seco", "", "20 a 60 s ou mais, Corte seco", "OK", ""]):
            ex(ws, f"{col}13", v, h="left" if col in "CJMO" else "center")
        note(ws, "E11", "Duração prevista do momento, em minutos.\nDigite só o número.")
        note(ws, "G11", "A fonte que fica no ar a maior parte do momento.\nAs alternativas são as fontes para onde sair.")
        note(ws, "L11", "Só para fonte principal Mídia ou Vinheta: o vídeo tem som?\nCombine com o operador de áudio.")
        note(ws, "N11", "Pendências, na ordem em que aparecem:\nduração, fonte principal, áudio da mídia, alternativa repetida, "
                        "momento longo sem alternativa, câmera principal sem operador, alternativa sem operador.")
    ws.row_dimensions[13].height = 21.75

    camnames, opcol = f"$C${CAM1}:$C${CAM_END}", f"$G${CAM1}:$G${CAM_END}"

    def semop(c):
        return f'IFERROR(INDEX({opcol},MATCH({c},{camnames},0))="",FALSE)'

    for k in range(N):
        r = R1 + k
        m = data["momentos"][k] if is_ex and k < len(data["momentos"]) else None
        bg = WHITE if is_ex else INPUT
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", m[0] if m else None, bg=bg)
        inp(ws, f"D{r}", m[1] if m else None, bg=bg, h="center")
        inp(ws, f"E{r}", m[2] if m else None, bg=bg, h="center")
        if k == 0:
            calc(ws, f"F{r}", f'=IF(OR(E{r}="",{HORA0}=""),"",{HORA0})', fmt=HORA)
        else:
            calc(ws, f"F{r}", f'=IF(OR(E{r}="",{HORA0}=""),"",{HORA0}+SUM($E${R1}:E{r - 1})/1440)', fmt=HORA)
        for col, j in zip("GHIJKL", range(3, 9)):
            inp(ws, f"{col}{r}", m[j] if m else None, bg=bg, h="left" if col == "J" else "center")
        calc(ws, f"M{r}", f'=IF(D{r}="","",IFERROR(VLOOKUP(D{r},{LOOK},2,FALSE)&", "&VLOOKUP(D{r},{LOOK},3,FALSE),""))',
             h="left", b=False, sz=9)
        G, Hh, I, E, L = (f"{c}{r}" for c in "GHIEL")
        calc(ws, f"N{r}",
             f'=IF(AND(C{r}="",{G}=""),"",IF({E}="","{PEND[0]}",IF({G}="","{PEND[1]}",'
             f'IF(AND(OR({G}="Mídia",{G}="Vinheta"),{L}=""),"{PEND[2]}",'
             f'IF(OR(AND({Hh}<>"",{Hh}={G}),AND({I}<>"",{I}={G}),AND({Hh}<>"",{Hh}={I})),"{PEND[3]}",'
             f'IF(AND({Hh}="",{I}="",{E}>{LIM}),"{PEND[4]}",'
             f'IF({semop(G)},"{PEND[5]}",IF(OR({semop(Hh)},{semop(I)}),"{PEND[6]}","OK"))))))))', h="left", b=False)
        inp(ws, f"O{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{R1}:D{R2}", TIPOS, "Tipo de momento: define o ritmo sugerido.")
    dv_number(ws, f"E{R1}:E{R2}")
    dv_list(ws, f"G{R1}:I{R2}", FONTES, "Câmera 1 a 6, Mídia, Slides, Letras ou Vinheta.")
    dv_list(ws, f"K{R1}:K{R2}", TRANS, "Corte seco, Fusão, Do preto ou Para o preto.")
    dv_list(ws, f"L{R1}:L{R2}", SIM_NAO, "O vídeo tem som?")
    cf_status(ws, f"N{R1}:N{R2}", f"N{R1}")

    # ---------------------------------------------------------------- câmeras e gravação
    TOTAL = f"SUM($E${R1}:$E${R2})"
    band(ws, CAM1 - 3, "Câmeras e gravação", "O")
    r = CAM1 - 2
    for col, t, bg in [("B", "#", INK), ("C", "Câmera ou gravador", INK), ("D", "Posição", INK), ("E", "Plano de base", INK),
                       ("G", "Operador", INK), ("H", "Grava?", AMBER_H), ("I", "Taxa", AMBER_H), ("J", "Espaço livre", AMBER_H),
                       ("K", "Minutos como principal", TEAL), ("L", "% do tempo", TEAL), ("M", "Espaço necessário", TEAL),
                       ("N", "Gravação", TEAL), ("O", "Uso no roteiro", TEAL)]:
        head(ws, f"{col}{r}", t, bg=bg)
    ws.merge_cells(f"E{r}:F{r}")
    ws.row_dimensions[r].height = 30
    for col, t in zip("BCDEGHIJKLMNO", ["", "", "onde fica", "quadro quando livre", "nome", "Sim / Não", "Mbps",
                                        "GB", "min", "", "GB, com a margem", "", ""]):
        hint(ws, f"{col}{r + 1}", t, merge=f"E{r + 1}:F{r + 1}" if col == "E" else None)
    ws.row_dimensions[r + 1].height = 21.75
    if not is_ex:
        note(ws, f"I{r}", "Taxa de dados da gravação, em megabits por segundo, como está no menu do gravador ou da câmera.\nVeja a tabela de taxas na aba Referência.")
        note(ws, f"M{r}", "Taxa × (duração do roteiro + margem) × 60 ÷ 8 ÷ 1000.")
    bg = WHITE if is_ex else INPUT
    for k, nome in enumerate(CAMS + GRAV):
        r = CAM1 + k
        is_cam = k < len(CAMS)
        if is_ex:
            v = data["cams"][k] if is_cam else (data["grav"][k - len(CAMS)][0], None, None) + data["grav"][k - len(CAMS)][1:]
        else:
            v = (None,) * 6
        num(ws, f"B{r}", k + 1)
        put(ws, f"C{r}", nome, f=font(10, True), bg=GRAY)
        inp(ws, f"D{r}", v[0], bg=bg)
        if is_cam:
            inp(ws, f"E{r}", v[1], bg=bg, merge=f"E{r}:F{r}")
            inp(ws, f"G{r}", v[2], bg=bg)
        else:
            hint(ws, f"E{r}", "–", merge=f"E{r}:F{r}")
            hint(ws, f"G{r}", "–")
        inp(ws, f"H{r}", v[3], bg=bg, h="center")
        inp(ws, f"I{r}", v[4], bg=bg, h="center")
        inp(ws, f"J{r}", v[5], bg=bg, h="center")
        if is_cam:
            calc(ws, f"K{r}", f'=IF({TOTAL}=0,"",SUMIF($G${R1}:$G${R2},C{r},$E${R1}:$E${R2}))')
            calc(ws, f"L{r}", f'=IF(K{r}="","",K{r}/{TOTAL})', fmt="0%")
        else:
            hint(ws, f"K{r}", "–")
            hint(ws, f"L{r}", "–")
        calc(ws, f"M{r}", f'=IF(OR(H{r}<>"Sim",I{r}="",{TOTAL}=0),"",ROUND(I{r}*({TOTAL}+{MARG})*60/8/1000,1))', fmt="0.0")
        calc(ws, f"N{r}", f'=IF(H{r}<>"Sim","",IF(OR(I{r}="",J{r}=""),"Informe taxa e espaço",IF(M{r}="","",IF(M{r}>J{r},"Falta espaço","OK"))))')
        if is_cam:
            calc(ws, f"O{r}", f'=IF(COUNTIF($G${R1}:$I${R2},C{r})=0,"",IF(G{r}="","Em uso sem operador",'
                              f'IF(AND(L{r}<>"",L{r}>{CONC}),"Concentrada: combine plano B","OK")))', h="left", b=False)
        else:
            hint(ws, f"O{r}", "–")
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"H{CAM1}:H{GRAV_END}", SIM_NAO, "Esta câmera ou este gravador vai gravar?")
    dv_number(ws, f"I{CAM1}:J{GRAV_END}")
    cf_status(ws, f"N{CAM1}:N{GRAV_END}", f"N{CAM1}")
    cf_status(ws, f"O{CAM1}:O{CAM_END}", f"O{CAM1}")

    # ---------------------------------------------------------------- resumo
    s = GRAV_END + 2
    band(ws, s, "Resumo automático", "O")
    Ncol = f"N{R1}:N{R2}"
    pend = "+".join(f'COUNTIF({Ncol},"{p}")' for p in PEND)
    Kc, Lc = f"K{CAM1}:K{CAM_END}", f"L{CAM1}:L{CAM_END}"
    linhas = [
        ("Momentos registrados", f"=COUNTA(C{R1}:C{R2})", None),
        ("Duração total (min)", f"={TOTAL}", None),
        ("Hora de término", f'=IF(OR({HORA0}="",{TOTAL}=0),"",{HORA0}+{TOTAL}/1440)', HORA),
        ("Momentos sem pendência (OK)", f'=COUNTIF({Ncol},"OK")', None),
        ("Momentos com pendência", f"={pend}", None),
        ("Câmera mais usada como principal", f'=IF(COUNT({Kc})=0,"",IF(MAX({Kc})=0,"",INDEX(C{CAM1}:C{CAM_END},MATCH(MAX({Kc}),{Kc},0))))', None),
        ("Parcela do tempo dessa câmera", f'=IF(COUNT({Lc})=0,"",MAX({Lc}))', "0%"),
        ("Gravações sem espaço suficiente", f'=COUNTIF(N{CAM1}:N{GRAV_END},"Falta espaço")', None),
        ("Câmeras em uso sem operador", f'=COUNTIF(O{CAM1}:O{CAM_END},"Em uso sem operador")', None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "F", "G", fmt=fmt, val_merge=f"G{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    v = lambda k: f"G{s + k}"
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:F{a}", h="right")
    calc(ws, f"G{a}", f'=IF({v(1)}=0,"Liste os momentos do evento",IF({HORA0}="","Informe a hora de início",'
                      f'IF({v(5)}>0,"Há momentos com pendência: veja a coluna Pendência",'
                      f'IF({v(8)}>0,"Há gravação sem espaço suficiente: troque o cartão ou baixe a taxa",'
                      f'IF({v(9)}>0,"Há câmera no roteiro sem operador",'
                      f'IF(AND({v(7)}<>"",{v(7)}>{CONC}),"OK, mas uma câmera concentra o programa: combine um plano B","OK"))))))',
         b=False, merge=f"G{a}:O{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    ws.conditional_formatting.add(f"G{a}:O{a}", CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", bgColor=GREEN, fgColor=GREEN)))
    ws.conditional_formatting.add(f"G{a}:O{a}", CellIsRule(operator="notEqual", formula=['"OK"'], fill=PatternFill("solid", bgColor=YELLOW, fgColor=YELLOW)))
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi feito antes do culto", "O")
        r = end + 3
        put(ws, f"B{r}", "Onde", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que a planilha apontou", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:H{r}")
        put(ws, f"I{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"I{r}:O{r}")
        ws.row_dimensions[r].height = 21.75
        for onde, apontou, feito in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", onde, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", apontou, merge=f"D{r}:H{r}")
            put(ws, f"I{r}", feito, merge=f"I{r}:O{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{R1}"
    setup(ws, tab, f"B1:O{end}", fit_height=True)


AMBER_H = HUE["amber"][0]

# ================================================================== aba Equipe
FUNCOES = ["Diretor de imagem", "Operador de corte", "Câmera 1", "Câmera 2", "Câmera 3", "Câmera 4", "Câmera 5",
           "Câmera 6", "Gráficos e mídia", "Áudio da transmissão", "Telão", "Apoio"]
COMANDOS = [  # ação, preparo, execução, observação
    ("Corte seco", "Prepara 2.", "Corta.", "Sempre o número da câmera, nunca o nome da pessoa."),
    ("Fusão", "Prepara fusão para a 1.", "Funde.", "Avise que é fusão já no preparo."),
    ("Legenda ou letra", "Prepara legenda do pastor.", "Entra legenda.", "Todo “entra” tem um “sai legenda”."),
    ("Vídeo", "Prepara vídeo de avisos.", "Roda vídeo.", "A mídia responde “vídeo pronto” e avisa “dez segundos”."),
    ("Pedido de plano", "3, congregação, plano aberto.", "–", "A câmera só se move se estiver livre."),
    ("Sem tally", "–", "2 no ar. / 2, liberada.", "O diretor avisa quem entrou e quem saiu do ar."),
    ("Abertura e fim", "Prepara abertura.", "Sai do preto. / Vai a preto.", "Combine se o FTB começa ativo."),
]
E1, NE = 13, 14


def aba_equipe(ws, tab):
    widths(ws, {"A": 2, "B": 5, "C": 22, "D": 24, "E": 24, "F": 11, "G": 22, "H": 10, "I": 18, "J": 34, "K": 2})
    title(ws, "Equipe — escala e linguagem de comando",
          "Preencha as células em amarelo-claro. A situação de cada linha e o aviso são calculados.", "J")
    for r, rot, fmt in [(4, "Evento", None), (5, "Data", DATE), (6, "Diretor de imagem", None), (7, "Hora de chegada da equipe", HORA)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", None, merge=f"D{r}:E{r}", fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    hint(ws, "G4", "Em equipe pequena, uma pessoa pode ter duas funções: registre as duas linhas com o mesmo nome.", merge="G4:J5", h="left")
    hint(ws, "G6", "Diretor, corte e câmeras precisam de intercom. Sem intercom, combine sinais e planos de base fixos.", merge="G6:J7", h="left")

    band(ws, 9, "Escala da equipe", "J")
    for col, t, bg in [("B", "#", INK), ("C", "Função", INK), ("D", "Nome", INK), ("E", "Posto", INK),
                       ("F", "Intercom", INK), ("G", "Substituto", INK), ("H", "Chegada", INK),
                       ("I", "Situação", TEAL), ("J", "Observação", INK)]:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 21.75
    for col, t in zip("BCDEFGHIJ", ["", "da lista", "", "onde fica", "Sim / Não", "quem cobre", "hh:mm", "calculada", ""]):
        hint(ws, f"{col}11", t)
    for col, v in zip("BCDEFGHIJ", ["Ex.", "Câmera 2", "Voluntário B", "Fundo, à esquerda", "Sim", "Voluntário D", "17:30", "OK", "Plano de base: médio do púlpito"]):
        ex(ws, f"{col}12", v, h="left" if col in "DEGJ" else "center")
    ws.row_dimensions[12].height = 21.75
    precisa = lambda c: "OR(" + ",".join(f'{c}="{f}"' for f in FUNCOES[:8]) + ")"
    for k in range(NE):
        r = E1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDEFGJ":
            inp(ws, f"{col}{r}", h="center" if col in "CF" else "left")
        inp(ws, f"H{r}", h="center", fmt=HORA)
        calc(ws, f"I{r}", f'=IF(C{r}="","",IF(D{r}="","Sem nome",IF(AND({precisa(f"C{r}")},F{r}<>"Sim"),"Sem intercom",'
                          f'IF(G{r}="","Sem substituto","OK"))))', b=False)
        ws.row_dimensions[r].height = 21.75
    E2 = E1 + NE - 1
    dv_list(ws, f"C{E1}:C{E2}", FUNCOES, "Escolha a função.")
    dv_list(ws, f"F{E1}:F{E2}", SIM_NAO, "Esta pessoa tem intercom?")
    cf_status(ws, f"I{E1}:I{E2}", f"I{E1}")

    s = E2 + 2
    band(ws, s, "Resumo automático", "J")
    Cc, Ic = f"C{E1}:C{E2}", f"I{E1}:I{E2}"
    linhas = [
        ("Funções na escala", f"=COUNTA({Cc})"),
        ("Diretor de imagem", f'=COUNTIF({Cc},"Diretor de imagem")'),
        ("Operador de corte", f'=COUNTIF({Cc},"Operador de corte")'),
        ("Câmeras com operador", "=" + "+".join(f'COUNTIF({Cc},"{c}")' for c in CAMS)),
        ("Linhas sem nome ou sem intercom", f'=COUNTIF({Ic},"Sem nome")+COUNTIF({Ic},"Sem intercom")'),
        ("Linhas sem substituto", f'=COUNTIF({Ic},"Sem substituto")'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Preencha a escala",IF(F{s + 2}=0,"Falta o diretor de imagem",'
                      f'IF(F{s + 3}=0,"Falta o operador de corte: se o diretor corta, registre as duas funções",'
                      f'IF(F{s + 4}=0,"Nenhuma câmera com operador",IF(F{s + 5}>0,"Há linha sem nome ou sem intercom",'
                      f'IF(F{s + 6}>0,"OK, mas há funções sem substituto","OK"))))))', b=False, merge=f"F{a}:J{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:J{a}")

    c = a + 2
    band(ws, c, "Palavras de comando combinadas", "J")
    for col, t, mg in [("B", "Ação", f"B{c + 1}:C{c + 1}"), ("D", "Preparo", None), ("E", "Execução", f"E{c + 1}:F{c + 1}"),
                       ("G", "Observação", f"G{c + 1}:J{c + 1}")]:
        put(ws, f"{col}{c + 1}", t, f=font(10, True), bg=GRAY, h="center", merge=mg)
    ws.row_dimensions[c + 1].height = 21.75
    r = c + 1
    for acao, prep, exe, obs in COMANDOS:
        r += 1
        put(ws, f"B{r}", acao, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", prep)
        inp(ws, f"E{r}", exe, merge=f"E{r}:F{r}")
        inp(ws, f"G{r}", obs, merge=f"G{r}:J{r}")
        ws.row_dimensions[r].height = 21.75
    hint(ws, f"B{r + 1}", "As palavras são sugestões. Troque pelas que a sua equipe usa, e use sempre as mesmas.", merge=f"B{r + 1}:J{r + 1}", h="left")
    ws.freeze_panes = f"C{E1}"
    setup(ws, tab, f"B1:J{r + 1}", fit_height=True)


# ================================================================== aba Registro de erros
TIPOS_ERRO = ["Corte atrasado", "Corte sem motivo", "Câmera mexendo no ar", "Salto", "Eixo cruzado", "Gráfico esquecido",
              "Gráfico errado", "Mídia sem som", "Tela preta ou buraco", "Fonte com defeito", "Comunicação"]
FONTES_ERRO = ["Câmera 1", "Câmera 2", "Câmera 3", "Câmera 4", "Câmera 5", "Câmera 6", "Mídia", "Gráficos", "Switcher", "Intercom"]
CAUSAS = ["Comando", "Roteiro", "Operação", "Equipamento", "Treinamento", "Comunicação"]
X1, NX = 13, 30


def aba_erros(ws, tab):
    widths(ws, {"A": 2, "B": 5, "C": 9, "D": 18, "E": 20, "F": 12, "G": 14, "H": 11, "I": 34, "J": 34, "K": 16, "L": 9, "M": 15, "N": 2})
    title(ws, "Registro de erros — para a reunião de revisão",
          "Anote cada erro com a hora em que apareceu na gravação. A situação de cada linha e o resumo são calculados.", "M")
    for r, rot, fmt in [(4, "Evento", None), (5, "Data do evento", DATE), (6, "Diretor de imagem", None), (7, "Revisão feita em", DATE)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:D{r}")
        inp(ws, f"E{r}", None, merge=f"E{r}:G{r}", fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "E5")
    dv_date(ws, "E7")
    hint(ws, "I4", "Assista à gravação do programa com a equipe. Para cada erro, uma linha.", merge="I4:M4", h="left")
    hint(ws, "I5", "A causa diz onde agir: Comando e Roteiro são da direção; Operação e Treinamento, de quem opera; Equipamento, da manutenção.", merge="I5:M6", h="left")
    hint(ws, "I7", "O objetivo é corrigir o processo, e não achar culpados.", merge="I7:M7", h="left")

    band(ws, 9, "Erros vistos na revisão", "M")
    for col, t, bg in [("B", "#", INK), ("C", "Hora", INK), ("D", "Momento", INK), ("E", "Tipo", RED_H), ("F", "Fonte", RED_H),
                       ("G", "Causa", RED_H), ("H", "Gravidade", RED_H), ("I", "O que aconteceu", INK),
                       ("J", "Ação para o próximo evento", INK), ("K", "Responsável", INK), ("L", "Ação feita?", INK),
                       ("M", "Situação", TEAL)]:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLM", ["", "hh:mm", "do roteiro", "da lista", "da lista", "da lista", "da lista", "", "", "", "Sim / Não", "calculada"]):
        hint(ws, f"{col}11", t)
    for col, v in zip("BCDEFGHIJKLM", ["Ex.", "19:31", "Avisos em vídeo", "Tela preta ou buraco", "Mídia", "Roteiro", "Alta",
                                       "Tela preta por 6 s no fim do vídeo", "Câmera geral no preview durante todo vídeo",
                                       "Diretor", "Sim", "OK"]):
        ex(ws, f"{col}12", v, h="left" if col in "DEIJK" else "center")
    ws.row_dimensions[12].height = 21.75
    note(ws, "G10", "Comando: faltou preparo ou o comando foi confuso.\nRoteiro: o momento não estava previsto.\n"
                    "Operação: erro de quem operava.\nEquipamento: falha de câmera, cabo, switcher.\n"
                    "Treinamento: a pessoa não sabia fazer.\nComunicação: intercom, ruído, aviso que não chegou.")
    for k in range(NX):
        r = X1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center", fmt=HORA)
        for col in "DEFGHIJKL":
            inp(ws, f"{col}{r}", h="center" if col in "FGHL" else "left")
        calc(ws, f"M{r}", f'=IF(E{r}="","",IF(J{r}="","Defina a ação",IF(L{r}="Sim","OK","Em aberto")))', b=False)
        ws.row_dimensions[r].height = 21.75
    X2 = X1 + NX - 1
    dv_list(ws, f"E{X1}:E{X2}", TIPOS_ERRO, "Escolha o tipo de erro.")
    dv_list(ws, f"F{X1}:F{X2}", FONTES_ERRO, "Onde o erro apareceu.")
    dv_list(ws, f"G{X1}:G{X2}", CAUSAS, "Comando, Roteiro, Operação, Equipamento, Treinamento ou Comunicação.")
    dv_list(ws, f"H{X1}:H{X2}", ["Baixa", "Média", "Alta"], "Alta: o público viu um defeito claro (tela preta, câmera mexendo, legenda errada).")
    dv_list(ws, f"L{X1}:L{X2}", SIM_NAO, "A ação já foi feita?")
    cf_equal(ws, f"H{X1}:H{X2}", [("Alta", RED), ("Média", YELLOW)])
    cf_status(ws, f"M{X1}:M{X2}", f"M{X1}")

    s = X2 + 2
    band(ws, s, "Resumo por tipo e por causa", "M")
    put(ws, f"B{s + 1}", "Tipo de erro", f=font(10, True), bg=GRAY, h="center", merge=f"B{s + 1}:E{s + 1}")
    put(ws, f"F{s + 1}", "Vezes", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"H{s + 1}", "Causa", f=font(10, True), bg=GRAY, h="center", merge=f"H{s + 1}:I{s + 1}")
    put(ws, f"J{s + 1}", "Vezes", f=font(10, True), bg=GRAY, h="center")
    ws.row_dimensions[s + 1].height = 21.75
    t1 = s + 2
    for k, t in enumerate(TIPOS_ERRO):
        r = t1 + k
        put(ws, f"B{r}", t, bg=GRAY, merge=f"B{r}:E{r}")
        calc(ws, f"F{r}", f'=COUNTIF($E${X1}:$E${X2},B{r})')
        if k < len(CAUSAS):
            put(ws, f"H{r}", CAUSAS[k], bg=GRAY, merge=f"H{r}:I{r}")
            calc(ws, f"J{r}", f'=COUNTIF($G${X1}:$G${X2},H{r})')
        ws.row_dimensions[r].height = 19.5
    t2 = t1 + len(TIPOS_ERRO) - 1
    c2 = t1 + len(CAUSAS) - 1
    q = c2 + 2
    linhas = [
        ("Erros registrados", f'=COUNTA(E{X1}:E{X2})'),
        ("Tipo mais frequente", f'=IF(MAX(F{t1}:F{t2})=0,"",INDEX(B{t1}:B{t2},MATCH(MAX(F{t1}:F{t2}),F{t1}:F{t2},0)))'),
        ("Causa mais frequente", f'=IF(MAX(J{t1}:J{c2})=0,"",INDEX(H{t1}:H{c2},MATCH(MAX(J{t1}:J{c2}),J{t1}:J{c2},0)))'),
        ("Erros de gravidade alta", f'=COUNTIF(H{X1}:H{X2},"Alta")'),
        ("Ações em aberto ou sem definição", f'=COUNTIF(M{X1}:M{X2},"Em aberto")+COUNTIF(M{X1}:M{X2},"Defina a ação")'),
    ]
    for k, (t, f) in enumerate(linhas):
        label(ws, f"H{q + k}", t, merge=f"H{q + k}:I{q + k}", h="right")
        calc(ws, f"J{q + k}", f, merge=f"J{q + k}:K{q + k}")
        ws.row_dimensions[q + k].height = 19.5
    a = q + len(linhas)
    label(ws, f"H{a}", "Aviso", merge=f"H{a}:I{a}", h="right")
    calc(ws, f"J{a}", f'=IF(J{q}=0,"Nenhum erro registrado",IF(J{q + 4}>0,"Há ações a definir ou a fazer antes do próximo evento","OK"))',
         b=False, merge=f"J{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"J{a}:M{a}")
    end = max(a, t2)
    ws.freeze_panes = f"C{X1}"
    setup(ws, tab, f"B1:M{end}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Corte ao vivo — modelo de roteiro de corte"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Corte ao vivo — como usar esta planilha",
           "Roteiro de corte, câmeras e gravação, escala da equipe e registro de erros para a revisão.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo de momento, fontes, transição, função, tipo e causa de erro aceitam só as opções da lista."),
        ("Números e horas", "Duração em minutos, só o número. Horas no formato 19:00."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Roteiro de corte: preencha o cabeçalho, a hora de início e confira os três parâmetros no alto."),
        ("2", "Aba Roteiro de corte: liste os momentos do evento na ordem, com duração, fonte principal e alternativas."),
        ("3", "Aba Roteiro de corte: preencha as câmeras (posição, plano de base, operador) e os gravadores."),
        ("4", "Resolva cada pendência e cada aviso. A linha fica verde quando está OK."),
        ("5", "Aba Equipe: monte a escala e confira as palavras de comando com a equipe."),
        ("6", "Aba Checklist: confira os doze itens antes de o evento começar."),
        ("7", "Depois do evento, aba Registro de erros: anote os erros vistos na gravação e a ação de cada um."),
    ]),
    ("Abas da planilha", [
        ("Roteiro de corte", "Até 20 momentos, seis câmeras e dois gravadores, com hora de cada momento, ritmo sugerido, pendências e espaço de gravação."),
        ("Equipe", "Escala de até 14 funções, com situação de cada linha, e as palavras de comando combinadas."),
        ("Registro de erros", "Até 30 erros, com contagem por tipo e por causa e ações em aberto."),
        ("Referência", "Ritmo por tipo de momento, transições, linguagem de comando, tally e taxas de gravação."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Roteiro", "O roteiro do culto do exemplo 1 do treinamento, antes das correções."),
    ]),
    ("Regras e premissas do modelo", [
        ("Hora de cada momento", "Hora de início do evento mais a soma das durações dos momentos anteriores."),
        ("Ritmo sugerido", "Vem da tabela da aba Referência, pelo tipo de momento. É ponto de partida, e não norma."),
        ("Pendência", "Uma por linha, na ordem: duração, fonte principal, áudio da mídia, alternativa repetida, momento longo sem alternativa, câmera sem operador."),
        ("Sem alternativa", "Só recebe alerta o momento mais longo que o limite do alto da aba (5 minutos no início). Troque se quiser."),
        ("Espaço de gravação", "Taxa (Mbps) × (duração do roteiro + margem) × 60 ÷ 8 ÷ 1000, em GB. A taxa real depende do formato: confira o manual."),
        ("Concentração", "Câmera principal em mais que a parcela do alto da aba (60% no início) recebe aviso para combinar um plano B."),
        ("Exemplos", "Os dados do exemplo e das linhas “Ex.” são ilustrativos, criados para este material."),
    ]),
])

aba_roteiro(wb.create_sheet("Roteiro de corte"), RED_H)
aba_equipe(wb.create_sheet("Equipe"), TEAL)
aba_erros(wb.create_sheet("Registro de erros"), HUE["amber"][0])

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 26, "D": 26, "E": 60, "F": 2})
title(ws, "Referência — corte ao vivo", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")


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


RITMO = [  # deve ocupar B6:D16, que a aba Roteiro consulta
    ("Abertura", "plano geral, 5 a 10 s", "Do preto ou fusão", "Logo ou vinheta, se houver. Situe quem chega à transmissão."),
    ("Louvor rápido", "3 a 6 s por plano", "Corte seco", "Corte no tempo forte, nas viradas das frases. Cortes fora do tempo cansam."),
    ("Louvor lento", "6 a 12 s por plano", "Fusão de 0,5 a 1 s", "Movimentos lentos de câmera no ar, se combinados."),
    ("Oração", "plano longo", "Fusão", "Plano aberto e estável. Sem gráfico. Evite closes de quem ora."),
    ("Leitura", "10 a 30 s por plano", "Corte seco", "Texto ou referência como gráfico, sem cobrir o rosto."),
    ("Pregação", "20 a 60 s ou mais", "Corte seco", "Saídas de 3 a 6 s para o geral e a reação. Não corte no meio da frase."),
    ("Avisos", "10 a 30 s por plano", "Corte seco", "Nome e função de quem fala na legenda."),
    ("Mídia", "o vídeo inteiro", "Corte seco", "Câmera geral no preview até o fim. Aviso de dez segundos."),
    ("Oferta ou ceia", "6 a 12 s por plano", "Fusão lenta", "Detalhe das mãos e dos elementos. Combine até onde chegar perto."),
    ("Entrevista", "5 a 15 s por plano", "Corte seco", "Corte para quem vai falar, um instante antes. Câmeras do mesmo lado do eixo."),
    ("Encerramento", "plano geral", "Para o preto, 1 a 2 s", "Tela final, se houver. Pare as gravações depois do preto."),
]
assert len(RITMO) == 11
r = tabela(4, "Ritmo por tipo de momento", ["Tipo de momento", "Duração de cada plano", "Transição", "Observação"], RITMO)
r = tabela(r, "Transições", ["Transição", "O que faz", "Tempo para começar", "Quando usar"], [
    ("Corte seco (cut)", "troca na hora", "instantâneo", "O padrão: fala, música com andamento, entrevista, reação."),
    ("Fusão (mix)", "uma imagem some e a outra aparece", "0,5 a 1 s entre câmeras; até 2 s entre momentos", "Música lenta, oração, passagem de um momento para outro."),
    ("Fade para preto (FTB)", "escurece tudo, inclusive os gráficos", "1 a 2 s", "Só no começo e no fim. O botão pisca enquanto está ativo."),
    ("Cortina (wipe) e efeitos", "uma imagem empurra ou cobre a outra", "–", "Quase nunca. Só em vinheta combinada."),
])
r = tabela(r, "Linguagem de comando", ["Ação", "Preparo", "Execução", "Cuidado"],
           [(a, p, e, o) for a, p, e, o in COMANDOS])
r = tabela(r, "Tally e janelas do multiview", ["Cor", "Significa", "Operador de câmera", "Observação"], [
    ("Vermelho", "no ar (program)", "segura o quadro", "Só se move para acompanhar a pessoa, com suavidade, ou a pedido do diretor."),
    ("Verde", "próxima (preview)", "termina o ajuste e segura", "O corte pode vir a qualquer instante."),
    ("Apagado", "livre", "volta ao plano de base", "Hora de acertar foco e quadro, ou de buscar o plano pedido."),
])
r = tabela(r, "Taxas de gravação e espaço", ["Formato", "Taxa típica", "Espaço por hora", "Observação"], [
    ("1080p comprimido, qualidade padrão", "10 a 25 Mbps", "4,5 a 11 GB", "Comum em celulares e em gravadores de switcher."),
    ("1080p comprimido, qualidade alta", "25 a 50 Mbps", "11 a 23 GB", "Comum em câmeras de vídeo e fotográficas."),
    ("4K comprimido", "50 a 150 Mbps", "23 a 68 GB", "Arquivos grandes. Confira o tempo máximo de gravação da câmera."),
    ("1080p em formato de edição", "100 a 220 Mbps", "45 a 100 GB", "Edição mais leve, arquivos muito maiores."),
], alt=31.5)
hint(ws, f"B{r - 1}", "Espaço (GB) = taxa (Mbps) × minutos × 60 ÷ 8 ÷ 1000. As taxas são faixas típicas e dependem do formato e da qualidade escolhidos: confira no menu e no manual.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
ws.row_dimensions[r - 1].height = 31.5
setup(ws, SLATE, f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist do corte ao vivo",
          "Confira antes de o evento começar; o último item fica para o fim. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["blue"][0])
aba_roteiro(wb.create_sheet("Exemplo - Roteiro"), MUTED, EX_ROTEIRO)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
