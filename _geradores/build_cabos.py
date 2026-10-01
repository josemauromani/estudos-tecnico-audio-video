# -*- coding: utf-8 -*-
"""Gera Cabos-e-Conexoes-modelo.xlsx: mapa de ligações, inventário de cabos, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIM_NAO = ["Sim", "Não"]
NIVEIS = ["Microfone", "Instrumento", "Linha", "Alto-falante"]
DIS = ["Não", "Passiva", "Ativa"]
LIGACOES = ["Balanceada", "Desbalanceada", "Cabo de caixa"]
ENTRADAS = ["Microfone", "Instrumento", "Linha", "Caixa passiva"]
TIPOS = ["Microfone (XLR)", "Instrumento (P10)", "Linha balanceado", "Caixa (alto-falante)", "Adaptador ou Y", "Rede", "Multicabo"]
CONECTORES = ["XLR macho", "XLR fêmea", "P10 TS", "P10 TRS", "P2", "RCA", "Speakon", "RJ45", "Outro"]
NP = 32  # ligações no mapa
NI = 40  # cabos no inventário

CHECK = [
    ("Cabos", "Cada cabo em uso passou nos testes de continuidade e de movimento e tem etiqueta nas duas pontas."),
    ("Cabos", "Há cabos de reserva testados, pelo menos um de cada tipo."),
    ("Ligações", "Cada fonte está na entrada do seu nível: microfone, instrumento ou linha."),
    ("Ligações", "Instrumentos e teclados longe da mesa passam por direct box e seguem em cabo balanceado."),
    ("Ligações", "Nenhum cabo desbalanceado tem mais de uns 6 metros."),
    ("Ligações", "As caixas passivas estão ligadas com cabo de caixa, e nenhuma saída de amplificador vai para a mesa ou para o multicabo."),
    ("Phantom", "O phantom está ligado só nos canais que precisam dele: microfone condensador e direct box ativa."),
    ("Phantom", "O phantom foi ligado com os canais em mute, e nenhuma fonte desbalanceada recebe phantom."),
    ("Ruído", "Todo o áudio está ligado no mesmo ponto de energia, com o pino de terra em todos os plugues."),
    ("Ruído", "Os cabos de áudio passam longe dos cabos de energia e de iluminação, ou cruzam em ângulo reto."),
    ("Ruído", "Com os canais abertos e sem ninguém tocando, não há zumbido nem estalos nas caixas."),
    ("Organização", "O mapa de ligações está preenchido e confere com as vias do multicabo, e os cabos no chão estão presos."),
]

# Exemplo 1 do treinamento: o mapa de ligações de um culto, como estava antes da correção
EX_MAPA = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "resp": "Operador de som da escala"},
    "limite": 6,
    "linhas": [  # origem, nível, DI, ligação, precisa de phantom, cabo, comprimento, via, destino, entrada, phantom ligado, observação
        ("Voz do pastor (mic dinâmico de mão)", "Microfone", "Não", "Balanceada", "Não", "M01", 10, 1, "Mesa canal 1", "Microfone", "Não", ""),
        ("Voz da cantora (mic dinâmico de mão)", "Microfone", "Não", "Balanceada", "Não", "M02", 10, 2, "Mesa canal 2", "Microfone", "Não", ""),
        ("Coral (condensador suspenso)", "Microfone", "Não", "Balanceada", "Sim", "M03", 15, 3, "Mesa canal 3", "Microfone", "Não", "Coral mudo."),
        ("Violão (captador sem pré)", "Instrumento", "Não", "Desbalanceada", "Não", "P03", 9, 4, "Mesa canal 4", "Microfone", "Não", "Cabo P10–XLR. Som fraco e sem brilho."),
        ("Baixo (captador passivo)", "Instrumento", "Ativa", "Balanceada", "Sim", "M05", 5, 5, "Mesa canal 5", "Microfone", "Sim", "P10 de 2 m até a DI."),
        ("Teclado E", "Linha", "Não", "Desbalanceada", "Não", "P04", 12, 6, "Mesa canal 6", "Microfone", "Sim", "Cabo P10–XLR. Som distorcido."),
        ("Teclado D", "Linha", "Não", "Desbalanceada", "Não", "P05", 12, 7, "Mesa canal 7", "Microfone", "Sim", "Cabo P10–XLR. Som distorcido."),
        ("Bumbo (mic dinâmico)", "Microfone", "Não", "Balanceada", "Não", "M06", 5, 8, "Mesa canal 8", "Microfone", "Não", ""),
        ("Caixa da bateria (mic dinâmico)", "Microfone", "Não", "Balanceada", "Não", "M08", 5, 9, "Mesa canal 9", "Microfone", "Não", ""),
        ("Computador (som dos vídeos)", "Linha", "Não", "Desbalanceada", "Não", "Y01", 2, None, "Mesa canais 11 e 12", "Linha", "Não", "Cabo em Y: P2 para dois P10. Fica ao lado da mesa."),
        ("Mesa: saída principal E", "Linha", "Não", "Balanceada", "Não", "L01", 15, "R1", "Caixa ativa esquerda", "Linha", "Não", ""),
        ("Mesa: saída principal D", "Linha", "Não", "Balanceada", "Não", "L02", 15, "R2", "Caixa ativa direita", "Linha", "Não", ""),
        ("Mesa: auxiliar 1 (retorno)", "Linha", "Não", "Balanceada", "Não", "L03", 3, "R3", "Amplificador de retorno", "Linha", "Não", ""),
        ("Amplificador de retorno: saída", "Alto-falante", "Não", "Desbalanceada", "Não", "P06", 8, None, "Retorno de chão passivo", "Caixa passiva", "Não", "Cabo de instrumento. Som fraco, cabo quente."),
    ],
    "ajustes": [  # linha, o que foi feito, situação depois
        ("3", "Phantom ligado no canal 3, com o canal em mute.", "OK"),
        ("4", "DI ativa ao lado do músico: P10 de 3 m até a DI e XLR até a via 4. Phantom ligado no canal 4.", "OK"),
        ("6 e 7", "Phantom desligado nos canais 6 e 7. DI estéreo passiva ao lado do teclado e dois cabos XLR até as vias 6 e 7.", "OK"),
        ("14", "Cabo de instrumento trocado por cabo de caixa de 2 × 2,5 mm².", "OK"),
        ("Todas", "Ajuste de ganho feito só depois da correção das ligações.", ""),
    ],
}

# Exemplo 3 do treinamento: os dez cabos testados depois de um evento ao ar livre
EX_INV = {
    "head": {"local": "Acervo da equipe de som", "resp": "Operador de som do evento", "data": date(2026, 11, 7)},
    "cabos": [  # etiqueta, tipo, conector A, conector B, comprimento, continuidade, movimento, estado, observação
        ("M01", TIPOS[0], "XLR fêmea", "XLR macho", 10, "OK", "OK", "Bom", ""),
        ("M02", TIPOS[0], "XLR fêmea", "XLR macho", 10, "OK", "OK", "Bom", ""),
        ("M03", TIPOS[0], "XLR fêmea", "XLR macho", 15, "OK", "OK", "Regular", "Capa com marcas de pisão. Observar."),
        ("M04", TIPOS[0], "XLR fêmea", "XLR macho", 10, "OK", "Falha", "Bom", "Estalos ao dobrar junto ao XLR fêmea. Era o cabo do apresentador."),
        ("M07", TIPOS[0], "XLR fêmea", "XLR macho", 5, "Falha", None, "Bom", "Pino 1 aberto: malha partida."),
        ("P01", TIPOS[1], "P10 TS", "P10 TS", 3, "OK", "OK", "Bom", ""),
        ("P02", TIPOS[1], "P10 TS", "P10 TS", 3, "OK", "OK", "Ruim", "Capa rachada, com a malha à mostra."),
        ("L01", TIPOS[2], "XLR fêmea", "XLR macho", 15, "OK", "OK", "Bom", ""),
        ("C01", TIPOS[3], "Speakon", "Speakon", 10, "OK", "OK", "Bom", ""),
        ("C02", TIPOS[3], "Speakon", "Speakon", 10, None, None, "Bom", "Ficou montado no local. Testar na desmontagem."),
    ],
    "necessarios": {TIPOS[0]: 4, TIPOS[1]: 1, TIPOS[2]: 1, TIPOS[3]: 2},
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def nest(pares, fim):
    """Monta IF(c1,"v1",IF(c2,"v2",...,"fim")) a partir de [(condição, valor), ...]."""
    expr = f'"{fim}"'
    for cond, val in reversed(pares):
        expr = f'IF({cond},"{val}",{expr})'
    return expr


# ================================================================== aba Mapa de ligações
M_W = {"A": 2, "B": 5, "C": 34, "D": 13, "E": 10, "F": 15, "G": 10, "H": 9, "I": 9, "J": 10, "K": 26, "L": 14, "M": 10,
       "N": 13, "O": 11, "P": 50, "Q": 34, "R": 2}
M1 = 13


def aba_mapa(ws, tab, data=None):
    widths(ws, M_W)
    is_ex = data is not None
    title(ws, "Mapa de ligações — da origem ao destino",
          "Exemplo preenchido, para consulta. Use a aba Mapa de ligações para o seu evento." if is_ex
          else "Uma linha por ligação. Preencha as células em amarelo-claro. As células cinza são calculadas.", "Q")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Responsável", "resp", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    label(ws, "I4", "Limite do cabo desbalanceado (m)", merge="I4:L4")
    inp(ws, "M4", data["limite"] if is_ex else 6, h="center", bg=hb)
    dv_number(ws, "M4")
    hint(ws, "N4", "Ponto de partida: 6 m. Em local com muita interferência, use menos.", merge="N4:Q4", h="left")
    hint(ws, "I5", "Nível da origem: o que a fonte entrega. Com direct box, a planilha considera que chega nível de microfone.",
         merge="I5:Q5", h="left")
    hint(ws, "I6", "Ligação: o tipo do cabo que chega ao destino. Depois de uma direct box, é balanceada.",
         merge="I6:Q6", h="left")
    hint(ws, "I7", "A coluna O que fazer mostra um aviso por ligação, o mais grave primeiro. Corrija e veja se aparece outro.",
         merge="I7:Q7", h="left")
    LIM = "$M$4"

    band(ws, 9, "Mapa de ligações (patch)", "Q")
    heads = [("B", "#", INK), ("C", "Origem", INK), ("D", "Nível da origem", BLUE), ("E", "Direct box", BLUE),
             ("F", "Ligação", BLUE), ("G", "Precisa de phantom", BLUE), ("H", "Cabo", INK), ("I", "Compr.", INK),
             ("J", "Via do multicabo", INK), ("K", "Destino", AMBER), ("L", "Entrada do destino", AMBER),
             ("M", "Phantom ligado", AMBER), ("N", "Nível que chega", TEAL), ("O", "Situação", TEAL),
             ("P", "O que fazer", TEAL), ("Q", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLMNOPQ", ["", "fonte ou saída", "lista", "lista", "do cabo até o destino", "Sim / Não",
                                           "etiqueta", "m", "nº ou código", "canal, caixa, amplificador", "lista",
                                           "Sim / Não", "", "", "", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 27
    if is_ex:
        hint(ws, "B12", "As ligações como estavam antes da correção. As colunas Nível que chega, Situação e O que fazer são calculadas.",
             merge="B12:Q12", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNOP", ["Ex.", "Violão (captador sem pré)", "Instrumento", "Ativa", "Balanceada", "Sim",
                                              "M05", 10, 4, "Mesa canal 4", "Microfone", "Sim", "Microfone", "OK",
                                              "Ligação correta"]):
            ex(ws, f"{col}12", v, h="left" if col in "CKP" else "center")
        ex(ws, "Q12", "")
        note(ws, "D10", "O nível que a fonte entrega.\nMicrofone: microfones.\nInstrumento: captador passivo.\nLinha: teclado, computador, saídas da mesa.\nAlto-falante: saída de amplificador.")
        note(ws, "E10", "A fonte passa por direct box antes deste cabo?\nPassiva: com transformador.\nAtiva: com circuito, pede phantom ou bateria.")
        note(ws, "F10", "Balanceada: XLR ou P10 TRS, entre saída e entrada balanceadas.\nDesbalanceada: P10 TS, P2, RCA ou cabo adaptador.\nCabo de caixa: dois condutores grossos, sem malha.")
        note(ws, "G10", "Sim para microfone condensador e para direct box ativa sem bateria.")
        note(ws, "L10", "O tipo de entrada em que o cabo chega.\nCaixa passiva: entrada de alto-falante, sem amplificador próprio.")
        note(ws, "M10", "A chave de +48 V deste canal está ligada?\nEm mesa com chave geral, marque Sim em todos os canais do grupo.")
    ws.row_dimensions[12].height = 24

    M2 = M1 + NP - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NP):
        r = M1 + k
        v = data["linhas"][k] if is_ex and k < len(data["linhas"]) else None
        D, E, F, G, I, J, K, L, M, N = (f"{c}{r}" for c in "DEFGIJKLMN")
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", v[0] if v else None, bg=bg)
        for col, j in zip("DEFGHIJ", range(1, 8)):
            inp(ws, f"{col}{r}", v[j] if v else None, h="center", bg=bg)
        inp(ws, f"K{r}", v[8] if v else None, bg=bg)
        inp(ws, f"L{r}", v[9] if v else None, h="center", bg=bg)
        inp(ws, f"M{r}", v[10] if v else None, h="center", bg=bg)
        calc(ws, N, f'=IF({D}="","",IF(OR({E}="Passiva",{E}="Ativa"),"Microfone",{D}))', b=False)
        regras = [  # condição, situação, o que fazer (a primeira que valer é a que aparece)
            (f'AND({N}="Alto-falante",{L}<>"Caixa passiva")', "Erro", "Saída de amplificador só pode ir para caixa passiva"),
            (f'AND({L}="Caixa passiva",{N}<>"Alto-falante")', "Erro", "Caixa passiva precisa de um amplificador antes dela"),
            (f'AND({N}="Alto-falante",{F}<>"Cabo de caixa")', "Erro", "Use cabo de caixa: dois condutores grossos, sem malha"),
            (f'AND({F}="Cabo de caixa",{N}<>"Alto-falante")', "Erro", "Cabo de caixa não tem malha: use cabo blindado"),
            (f'AND({M}="Sim",{F}="Desbalanceada")', "Erro", "Desligue o phantom ou passe a fonte por direct box"),
            (f'AND({N}="Instrumento",{L}<>"Instrumento")', "Atenção", "Use direct box ou a entrada de instrumento (Hi-Z)"),
            (f'AND({N}="Linha",{L}="Microfone")', "Atenção", "Use a entrada de linha, ou direct box na entrada de microfone"),
            (f'AND({N}="Linha",{L}="Instrumento")', "Atenção", "Use a entrada de linha"),
            (f'AND({N}="Microfone",{L}<>"Microfone")', "Atenção", "Sinal fraco para esta entrada: use a entrada de microfone"),
            (f'AND({G}="Sim",{M}<>"Sim")', "Atenção", "Ligue o phantom deste canal, com o canal em mute"),
            (f'AND({F}="Desbalanceada",{I}<>"",{LIM}<>"",{I}>{LIM})', "Atenção", "Cabo desbalanceado longo: use direct box e cabo balanceado"),
            (f'AND({J}<>"",COUNTIF($J${M1}:$J${M2},{J})>1)', "Atenção", "Via do multicabo repetida: confira"),
            (f'AND({K}<>"",COUNTIF($K${M1}:$K${M2},{K})>1)', "Atenção", "Destino repetido: confira"),
            (f'{F}=""', "Atenção", "Informe o tipo de ligação do cabo"),
        ]
        vazio = f'OR({D}="",{L}="")'
        calc(ws, f"O{r}", f'=IF({vazio},"",{nest([(c, s) for c, s, _ in regras], "OK")})')
        calc(ws, f"P{r}", f'=IF({vazio},"",{nest([(c, m) for c, _, m in regras], "Ligação correta")})', h="left", b=False, sz=9)
        inp(ws, f"Q{r}", v[11] if v else None, bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{M1}:D{M2}", NIVEIS, "Microfone, Instrumento, Linha ou Alto-falante.")
    dv_list(ws, f"E{M1}:E{M2}", DIS, "A fonte passa por direct box? Não, Passiva ou Ativa.")
    dv_list(ws, f"F{M1}:F{M2}", LIGACOES, "Balanceada, Desbalanceada ou Cabo de caixa.")
    dv_list(ws, f"G{M1}:G{M2}", SIM_NAO, "A fonte precisa de phantom?")
    dv_list(ws, f"L{M1}:L{M2}", ENTRADAS, "Microfone, Instrumento, Linha ou Caixa passiva.")
    dv_list(ws, f"M{M1}:M{M2}", SIM_NAO, "O phantom deste canal está ligado?")
    dv_number(ws, f"I{M1}:I{M2}")
    cf_equal(ws, f"O{M1}:O{M2}", [("OK", GREEN), ("Atenção", YELLOW), ("Erro", RED)])

    s = M2 + 2
    band(ws, s, "Resumo automático", "Q")
    cont = lambda t: f'COUNTIF(O{M1}:O{M2},"{t}")'
    linhas = [
        ("Ligações listadas", f"=COUNTA(C{M1}:C{M2})"),
        ("Ligações conferidas", f'={cont("OK")}+{cont("Atenção")}+{cont("Erro")}'),
        ("Corretas (OK)", f'={cont("OK")}'),
        ("Com atenção", f'={cont("Atenção")}'),
        ("Com erro", f'={cont("Erro")}'),
        ("Canais com phantom ligado", f'=COUNTIF(M{M1}:M{M2},"Sim")'),
        ("Vias do multicabo em uso", f"=COUNTA(J{M1}:J{M2})"),
        ("Direct boxes em uso", f'=COUNTIF(E{M1}:E{M2},"Passiva")+COUNTIF(E{M1}:E{M2},"Ativa")'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste as ligações",IF(F{s + 5}>0,"Há erro de ligação: corrija antes de ligar o sistema",'
                      f'IF(F{s + 4}>0,"Há ligação a conferir",IF(F{s + 2}<F{s + 1},"Há ligação sem nível ou sem entrada informados","OK"))))',
         b=False, merge=f"F{a}:P{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:P{a}")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi corrigido", "Q")
        r = end + 3
        put(ws, f"B{r}", "Linha", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:O{r}")
        put(ws, f"P{r}", "Situação depois", f=font(10, True), bg=GRAY, h="center", merge=f"P{r}:Q{r}")
        ws.row_dimensions[r].height = 21.75
        for linha, feito, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", linha, f=font(10, True, c=MUTED), bg=GRAY, h="center", merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:O{r}")
            put(ws, f"P{r}", depois, merge=f"P{r}:Q{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{M1}"
    setup(ws, tab, f"B1:Q{end}", fit_height=True)


# ================================================================== aba Inventário de cabos
I_W = {"A": 2, "B": 5, "C": 11, "D": 22, "E": 13, "F": 13, "G": 9, "H": 14, "I": 12, "J": 12, "K": 12, "L": 12, "M": 40,
       "N": 46, "O": 2}
I1 = 12


def aba_inventario(ws, tab, data=None):
    widths(ws, I_W)
    is_ex = data is not None
    title(ws, "Inventário de cabos — teste e situação",
          "Exemplo preenchido, para consulta. Use a aba Inventário de cabos para o seu acervo." if is_ex
          else "Um cabo por linha, com a etiqueta que está nas duas pontas. Anote os dois testes e o estado físico.", "N")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Local ou acervo", "local", None), (5, "Responsável", "resp", None), (6, "Data do teste", "data", DATE)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:D{r}")
        inp(ws, f"E{r}", H.get(k), merge=f"E{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "E6")
    hint(ws, "I4", "Continuidade: no testador, cada contato chega ao seu par, sem curto e sem fio trocado.", merge="I4:N4", h="left")
    hint(ws, "I5", "Movimento: com o cabo no testador, dobre e balance junto de cada conector. Nenhuma luz pode piscar.", merge="I5:N5", h="left")
    hint(ws, "I6", "Cabo em Reparar sai de circulação: marque com fita, guarde em outro lugar e anote o defeito.", merge="I6:N6", h="left")

    band(ws, 8, "Inventário de cabos", "N")
    heads = [("B", "#", INK), ("C", "Etiqueta", INK), ("D", "Tipo", INK), ("E", "Conector A", INK), ("F", "Conector B", INK),
             ("G", "Compr.", INK), ("H", "Continuidade", AMBER), ("I", "Movimento", AMBER), ("J", "Estado físico", AMBER),
             ("K", "Data do teste", AMBER), ("L", "Situação", TEAL), ("M", "O que fazer", TEAL),
             ("N", "Defeito encontrado ou observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["", "código", "lista", "lista", "lista", "m", "OK / Falha", "OK / Falha",
                                        "Bom / Regular / Ruim", "dd/mm/aaaa", "", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 27
    if is_ex:
        hint(ws, "B11", "Os dez cabos do evento, testados na desmontagem. As colunas Situação e O que fazer são calculadas.",
             merge="B11:N11", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "M01", TIPOS[0], "XLR fêmea", "XLR macho", 10, "OK", "OK", "Bom",
                                           date(2026, 10, 4), "Liberado", "Pode usar"]):
            ex(ws, f"{col}11", v, h="left" if col in "DM" else "center", fmt=DATE if col == "K" else None)
        ex(ws, "N11", "")
        note(ws, "C9", "O código que está na etiqueta das duas pontas do cabo.\nUse o mesmo código no mapa de ligações.")
        note(ws, "J9", "Bom: capa e conectores inteiros.\nRegular: marcas de uso, a observar.\nRuim: capa rachada, malha à mostra, conector frouxo ou amassado.")
        note(ws, "L9", "Liberado: passou nos dois testes e não está em estado ruim.\nReparar: falhou em um teste, ou o estado é ruim.\nTestar: falta algum teste.")
    ws.row_dimensions[11].height = 24

    I2 = I1 + NI - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NI):
        r = I1 + k
        c = data["cabos"][k] if is_ex and k < len(data["cabos"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, h="center", bg=bg)
        inp(ws, f"D{r}", c[1] if c else None, bg=bg)
        for col, j in zip("EFGHIJ", range(2, 8)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg)
        inp(ws, f"K{r}", (data["head"]["data"] if c and (c[5] or c[6]) else None) if is_ex else None, h="center", bg=bg, fmt=DATE)
        calc(ws, f"L{r}", f'=IF(C{r}="","",IF(OR(H{r}="Falha",I{r}="Falha",J{r}="Ruim"),"Reparar",'
                          f'IF(OR(H{r}="",I{r}=""),"Testar","Liberado")))')
        calc(ws, f"M{r}", f'=IF(L{r}="","",IF(L{r}="Reparar","Tire de uso, marque e conserte ou descarte",'
                          f'IF(L{r}="Testar","Faça os dois testes antes do evento",'
                          f'IF(J{r}="Regular","Pode usar. Observe no próximo teste","Pode usar"))))', h="left", b=False, sz=9)
        inp(ws, f"N{r}", c[8] if c else None, bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{I1}:D{I2}", TIPOS, "Escolha o tipo do cabo.")
    dv_list(ws, f"E{I1}:F{I2}", CONECTORES, "Escolha o conector desta ponta.")
    dv_list(ws, f"H{I1}:I{I2}", ["OK", "Falha"], "OK ou Falha. Deixe em branco se não testou.")
    dv_list(ws, f"J{I1}:J{I2}", ["Bom", "Regular", "Ruim"], "Bom, Regular ou Ruim.")
    dv_number(ws, f"G{I1}:G{I2}")
    dv_date(ws, f"K{I1}:K{I2}")
    cf_equal(ws, f"L{I1}:L{I2}", [("Liberado", GREEN), ("Testar", YELLOW), ("Reparar", RED)])
    cf_equal(ws, f"H{I1}:I{I2}", [("Falha", RED)])
    cf_equal(ws, f"J{I1}:J{I2}", [("Ruim", RED)])

    s = I2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(L{I1}:L{I2},"{t}")'
    linhas = [
        ("Cabos no inventário", f"=COUNTA(C{I1}:C{I2})", None),
        ("Liberados", f'={cont("Liberado")}', None),
        ("Para reparar", f'={cont("Reparar")}', None),
        ("Por testar", f'={cont("Testar")}', None),
        ("Metros de cabo liberados", f'=SUMIF(L{I1}:L{I2},"Liberado",G{I1}:G{I2})', "0"),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "D", "E", fmt=fmt)
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(E{s + 1}=0,"Liste os cabos",IF(E{s + 3}>0,"Há cabo para reparar: tire de circulação",'
                      f'IF(E{s + 4}>0,"Há cabo sem teste","OK")))', b=False, merge=f"E{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:M{a}")

    t0 = a + 2
    band(ws, t0, "Estoque por tipo: o que há e o que o evento pede", "N")
    r = t0 + 1
    put(ws, f"B{r}", "Tipo", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:D{r}")
    for col, t in zip("EFGH", ["No inventário", "Liberados", "Necessários", "Reserva"]):
        put(ws, f"{col}{r}", t, f=font(10, True), bg=GRAY, h="center")
    put(ws, f"I{r}", "Situação", f=font(10, True), bg=GRAY, h="center", merge=f"I{r}:K{r}")
    hint(ws, f"L{r}", "Em Necessários, digite quantos cabos de cada tipo o evento pede. A reserva é o que sobra entre os liberados.",
         merge=f"L{r}:N{r}", h="left")
    ws.row_dimensions[r].height = 27
    nec = data["necessarios"] if is_ex else {}
    for tipo in TIPOS:
        r += 1
        label(ws, f"B{r}", tipo, merge=f"B{r}:D{r}")
        calc(ws, f"E{r}", f'=COUNTIF(D{I1}:D{I2},B{r})', b=False)
        calc(ws, f"F{r}", f'=COUNTIFS(D{I1}:D{I2},B{r},L{I1}:L{I2},"Liberado")')
        inp(ws, f"G{r}", nec.get(tipo), h="center", bg=bg)
        calc(ws, f"H{r}", f'=IF(G{r}="","",F{r}-G{r})', b=False, fmt='+0;-0;0')
        calc(ws, f"I{r}", f'=IF(G{r}="","",IF(H{r}<0,"Faltam "&ABS(H{r}),IF(H{r}=0,"Sem reserva","OK")))', merge=f"I{r}:K{r}")
        cf_status(ws, f"I{r}:K{r}", f"$I${r}")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"G{t0 + 2}:G{r}")
    ws.freeze_panes = f"D{I1}"
    setup(ws, tab, f"B1:N{r}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Cabos, conectores e níveis de sinal — modelo de mapa de ligações e inventário"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Cabos e conexões — como usar esta planilha",
           "Mapa de ligações com conferência de nível, phantom e cabo, e inventário de cabos com teste e situação.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Nível, direct box, ligação, phantom, entrada, tipo de cabo, conectores e testes aceitam só as opções da lista."),
        ("Números", "Digite o comprimento só com o número, em metros: 10, e não “10 m”."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Inventário de cabos: liste cada cabo com a etiqueta que está nas duas pontas, o tipo, os conectores e o comprimento."),
        ("2", "No testador: faça o teste de continuidade e o de movimento em cada cabo. Anote os resultados e o estado físico."),
        ("3", "Aba Inventário de cabos: veja a situação. Tire de circulação os cabos em Reparar. Informe quantos cabos de cada tipo o evento pede."),
        ("4", "Aba Mapa de ligações: uma linha por ligação, da fonte até a mesa e da mesa até as caixas."),
        ("5", "Veja a situação e o que fazer em cada linha. Corrija a ligação, troque os dados e confira se aparece outro aviso."),
        ("6", "Faça as ligações seguindo o mapa: cada cabo na via e no canal anotados."),
        ("7", "Aba Checklist: confira os doze itens antes do ajuste de ganho."),
    ]),
    ("Abas da planilha", [
        ("Mapa de ligações", "Até 32 ligações, com o nível que chega ao destino e a situação calculada: OK, Atenção ou Erro."),
        ("Inventário de cabos", "Até 40 cabos, com situação calculada (Liberado, Reparar ou Testar) e o estoque por tipo."),
        ("Referência", "Níveis de sinal, conector por tipo de sinal, cabo por trecho, phantom e tipos de ruído."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Mapa", "As catorze ligações de um culto, antes da correção (exemplo 1 do treinamento)."),
        ("Exemplo - Inventário", "Os dez cabos testados depois de um evento ao ar livre (exemplo 3 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Nível que chega", "Com direct box (passiva ou ativa), o modelo considera que chega nível de microfone. Sem ela, chega o nível da origem."),
        ("Erro", "Risco para o equipamento: saída de amplificador fora de caixa passiva, cabo de caixa trocado por cabo blindado (ou o contrário) e phantom em ligação desbalanceada."),
        ("Atenção", "Nível em entrada que não é a dele, phantom faltando, cabo desbalanceado acima do limite, via do multicabo ou destino repetidos."),
        ("Um aviso por linha", "A coluna O que fazer mostra só o primeiro problema, o mais grave. Depois de corrigir, confira de novo."),
        ("Limite do desbalanceado", "O valor inicial de 6 m é regra prática, e não norma. Troque no alto da aba, conforme o seu local."),
        ("Phantom", "O modelo não sabe se o microfone é condensador: quem informa é você, na coluna Precisa de phantom."),
        ("Situação do cabo", "Reparar: falhou em um teste, ou o estado físico é ruim. Testar: falta algum teste. Liberado: passou nos dois."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_mapa(wb.create_sheet("Mapa de ligações"), BLUE)
aba_inventario(wb.create_sheet("Inventário de cabos"), AMBER)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 34, "C": 34, "D": 34, "E": 62, "F": 2})
title(ws, "Referência — níveis, conectores e cabos", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")


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


r = tabela(4, "Os quatro níveis de sinal", ["Nível", "Tamanho típico", "Quem entrega", "Onde entra e por que cabo"], [
    ("Microfone", "−60 a −20 dBu (1 a 80 mV)", "Microfones, saída XLR da direct box", "Entrada de microfone. Cabo balanceado com XLR."),
    ("Instrumento", "−30 a −10 dBu, alta impedância", "Guitarra, baixo, violão com captador passivo", "Entrada de instrumento (Hi-Z), amplificador do músico ou direct box. Cabo P10 blindado, curto."),
    ("Linha doméstica", "−10 dBV (0,3 V)", "Computador, celular, tocador, aparelhos com P2 ou RCA", "Entrada de linha, ou direct box. Cabo desbalanceado curto."),
    ("Linha profissional", "+4 dBu (1,2 V)", "Saídas da mesa, processadores, muitos teclados e receptores", "Entrada de linha da mesa, da caixa ativa ou do amplificador. Cabo balanceado."),
    ("Alto-falante", "dezenas de volts, vários ampères", "Só a saída do amplificador de potência", "Só a caixa passiva. Cabo de caixa: dois condutores grossos, sem malha."),
])
r = tabela(r, "Conector por tipo de sinal", ["Conector", "Contatos", "Sinal típico", "Observação"], [
    ("XLR", "1 malha, 2 positivo, 3 negativo", "Microfone ou linha, balanceado", "O sinal sai pelos pinos: saída é macho, entrada é fêmea. O phantom chega por ele."),
    ("P10 TS", "ponta sinal, luva malha", "Instrumento ou linha, desbalanceado", "Trechos curtos. Em caixa passiva, só com cabo de caixa."),
    ("P10 TRS", "ponta, anel, luva", "Linha balanceada; fone estéreo; insert", "O mesmo plugue serve a três usos. Saída de fone estéreo não se liga a entrada balanceada com cabo TRS comum."),
    ("P2 (3,5 mm)", "ponta esquerdo, anel direito, luva comum", "Linha doméstica, estéreo", "Solta com facilidade: prenda o cabo. Use cabo em Y para dois canais da mesa."),
    ("RCA", "pino central sinal, anel externo malha", "Linha doméstica, em par", "Branco: esquerdo. Vermelho: direito. Trechos curtos."),
    ("Speakon", "1+ e 1−; 2+ e 2− na segunda via", "Alto-falante", "Ligue e desligue com o amplificador desligado. Confira no manual qual par a caixa usa."),
    ("Combo", "XLR e P10 no mesmo furo", "Conforme o plugue", "Em geral, XLR vai ao pré de microfone e P10 entra como linha. Confira o manual."),
])
r = tabela(r, "Cabo por trecho", ["Trecho", "Cabo", "Comprimento prático", "Observação"], [
    ("Microfone até a caixa de palco ou a mesa", "Balanceado, XLR", "dezenas de metros", "É o menor sinal do sistema: cabo inteiro e bem blindado."),
    ("Instrumento até a direct box ou o amplificador", "Blindado desbalanceado, P10 TS", "até 5 ou 6 m", "Mais longo, perde agudos e capta ruído."),
    ("Teclado ou computador até a direct box ou a mesa", "Blindado desbalanceado", "até 5 ou 6 m", "Se a mesa está longe, direct box ao lado da fonte."),
    ("Direct box até a mesa", "Balanceado, XLR", "dezenas de metros", "Entra em entrada de microfone."),
    ("Mesa até caixa ativa, amplificador ou processador", "Balanceado, XLR ou P10 TRS", "dezenas de metros", "Nível de linha. Pode ir pelos retornos do multicabo."),
    ("Amplificador até a caixa passiva", "Cabo de caixa, sem malha", "o mais curto que der", "A bitola depende do comprimento e da impedância: veja o estudo Caixas, amplificação e PA."),
    ("Caixa de palco digital até a mesa", "Cabo de rede indicado pelo fabricante, ou fibra", "em geral até cerca de 100 m", "O limite depende do sistema e do cabo. Tenha um cabo reserva testado."),
])
r = tabela(r, "Phantom (+48 V): quem precisa", ["O que está no canal", "Phantom", "Por quê", "Cuidado"], [
    ("Microfone condensador com fio", "ligado", "Alimenta o circuito interno", "Ligue com o canal em mute e espere alguns segundos."),
    ("Direct box ativa", "ligado", "Alimenta o circuito da DI", "Alguns modelos aceitam bateria."),
    ("Microfone dinâmico em cabo balanceado", "tanto faz", "Não usa a alimentação", "Não é prejudicado se o cabo estiver bom. Prefira desligado."),
    ("Direct box passiva", "tanto faz", "O transformador não usa nem repassa", "–"),
    ("Receptor sem fio, microfone de fita", "desligado", "Não precisam", "O de fita passivo pode ser danificado se houver defeito no cabo."),
    ("Fonte desbalanceada com cabo adaptador até o XLR", "desligado", "A tensão chega à saída do aparelho", "Use direct box."),
])
r = tabela(r, "Ruído: o que se ouve e onde procurar", ["O que se ouve", "Causa provável", "Onde procurar", "O que fazer"], [
    ("Zumbido grave e contínuo", "Laço de terra", "Aparelho ligado em outra tomada, por cabo desbalanceado", "Mesmo ponto de energia, ligação balanceada, direct box com ground lift. Nunca retire o pino de terra."),
    ("Zunido áspero, que muda com a luz", "Interferência", "Cabo de áudio junto de cabo de energia ou de dimmer", "Separe os caminhos. Cruze em ângulo reto. Use ligação balanceada."),
    ("Estalos, som que some e volta", "Mau contato", "Cabo ou conector do canal", "Troque o cabo. Teste, marque e tire de uso."),
    ("Chiado contínuo", "Sinal fraco em alguma etapa", "Entrada errada, pad esquecido, ganho baixo", "Confira a entrada e o nível: veja o estudo Ganho e fader."),
    ("Canal de condensador mudo", "Phantom", "Chave de +48 V desligada ou malha do cabo partida", "Ligue o phantom com o canal em mute. Se continuar, troque o cabo."),
])
hint(ws, f"B{r - 1}", "Faixas típicas e regras práticas deste material. Mudam com o equipamento, com o cabo e com a interferência do local.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplos
checklist(wb.create_sheet("Checklist"), "Checklist de cabos e ligações",
          "Confira na montagem, antes do ajuste de ganho. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=TEAL)
aba_mapa(wb.create_sheet("Exemplo - Mapa"), MUTED, EX_MAPA)
aba_inventario(wb.create_sheet("Exemplo - Inventário"), MUTED, EX_INV)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
