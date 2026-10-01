# -*- coding: utf-8 -*-
"""Gera Mesa-de-Som-modelo.xlsx: folha de canais, mapa de saídas e auxiliares, registro de cenas,
referência, checklist e exemplo."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
TEAL, TEAL_T = HUE["teal"]
AMBER, AMBER_T = HUE["amber"]
SIM_NAO = ["Sim", "Não"]

TIPOS = ["Dinâmico", "Condensador", "Fita", "DI passiva", "DI ativa", "Linha", "Sem fio", "USB ou rede"]
DESTINOS = ["LR", "Grupo 1-2", "Grupo 3-4", "Grupo 5-6", "LR + grupo", "Só auxiliares", "Nenhum"]
TIPOS_BUS = ["Master LR", "Grupo", "Auxiliar", "Matriz"]
TOMADAS = ["Pré-fader", "Pós-fader"]
USOS = ["Caixas da frente", "Retorno de palco", "Fone ou in-ear", "Efeito", "Subwoofer", "Outra sala",
        "Caixas de delay", "Gravação", "Transmissão", "Conjunto de canais"]

CHECK = [
    ("Preparação", "A folha de canais está preenchida e bate com os nomes escritos na mesa e na caixa de palco."),
    ("Preparação", "A mesa foi zerada, ou a cena de partida foi carregada, antes de ligar as caixas."),
    ("Entrada", "Cada fonte está na entrada anotada, e nenhuma entrada física aparece em dois canais sem querer."),
    ("Entrada", "O phantom (+48 V) está ligado só nos canais que precisam dele."),
    ("Entrada", "O filtro passa-altas de cada canal está como anotado na folha de canais."),
    ("Barramentos", "Cada canal tem um só caminho até o master: direto ou pelo grupo."),
    ("Barramentos", "Os grupos em uso estão endereçados ao master, e os DCA em uso estão em 0 dB."),
    ("Barramentos", "Os auxiliares de retorno estão em pré-fader, e os de efeito, em pós-fader."),
    ("Saídas", "Cada saída física está ligada ao destino anotado no mapa de saídas."),
    ("Operação", "Nenhum solo ficou apertado, e o modo de solo da mesa é PFL, e não solo in place."),
    ("Operação", "Na mesa digital, a camada à vista é a dos canais principais, e o modo de envios nos faders está desligado."),
    ("Operação", "A configuração final foi guardada: cena salva com cópia fora da mesa, ou foto do painel da mesa analógica."),
]

# Exemplo 1 do treinamento: a folha de canais de um culto, antes da correção
EX_CANAIS = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "mesa": "Mesa digital de 16 canais, caixa de palco de 16 entradas"},
    "nent": 16,
    "canais": [  # fonte, tipo, entrada, phantom, HPF, Hz, destino, DCA
        ("Voz do pastor (sem fio)", "Sem fio", 1, "Não", "Sim", 100, "LR", "Vozes"),
        ("Microfone do púlpito", "Condensador", 2, "Sim", "Sim", 100, "LR", "Vozes"),
        ("Voz da cantora", "Dinâmico", 3, "Não", "Sim", 100, "LR", "Vozes"),
        ("Voz de apoio 1", "Dinâmico", 4, "Não", "Sim", 100, "Grupo 1-2", "Vozes"),
        ("Voz de apoio 2", "Dinâmico", 5, "Não", "Sim", 100, "LR + grupo", "Vozes"),
        ("Coral (condensador suspenso)", "Condensador", 6, "Não", "Sim", 120, "Grupo 1-2", "Vozes"),
        ("Violão", "DI ativa", 7, "Sim", "Sim", 80, "LR", "Banda"),
        ("Baixo", "DI passiva", 8, "Não", "Não", None, "LR", "Banda"),
        ("Teclado E", "Linha", 9, "Não", "Não", None, "LR", "Banda"),
        ("Teclado D", "Linha", 9, "Não", "Não", None, "LR", "Banda"),
        ("Bumbo", "Dinâmico", 11, "Não", "Sim", 40, "Grupo 3-4", "Banda"),
        ("Caixa da bateria", "Dinâmico", 12, "Não", "Sim", 80, "Grupo 3-4", "Banda"),
        ("Overhead da bateria", "Condensador", 13, "Sim", "Sim", 100, "Grupo 3-4", "Banda"),
        ("Guitarra (amplificador)", "Dinâmico", 14, "Não", "Sim", 80, "LR", "Banda"),
        ("Computador E (vídeos)", "USB ou rede", None, "Não", "Não", None, "LR", "Mídia"),
        ("Computador D (vídeos)", "USB ou rede", None, "Não", "Não", None, "LR", "Mídia"),
    ],
    "ajustes": [  # canal, o que foi feito
        ("5", "Botão L-R desligado. A voz passou a ir ao master só pelo grupo 1-2 e voltou ao nível das outras vozes."),
        ("6", "Phantom ligado com o canal em mute. O sinal do coral apareceu no PFL."),
        ("9 e 10", "Patch do canal 10 corrigido para a entrada 10, que estava livre. O teclado voltou a soar dos dois lados."),
        ("Todos", "Teste de linha com a folha na mão. Cena salva depois da passagem, com cópia no pendrive."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def cf_situacao(ws, rng, first, vermelhos):
    """Vermelho para os avisos graves, verde para OK, amarelo para os outros textos, nada para a célula vazia."""
    cf_equal(ws, rng, [(t, RED) for t in vermelhos] + [("OK", GREEN)])
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(LEN({first})>0,{first}<>"OK")'],
                                                   fill=PatternFill("solid", bgColor=YELLOW, fgColor=YELLOW)))


def cabecalho(ws, H, hb, campos):
    for r, rot, k, fmt in campos:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75


def bloco_ajustes(ws, end, ajustes, last):
    band(ws, end + 2, "O que foi corrigido", last)
    r = end + 3
    put(ws, f"B{r}", "Canal", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"C{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:{last}{r}")
    ws.row_dimensions[r].height = 21.75
    for canal, feito in ajustes:
        r += 1
        put(ws, f"B{r}", canal, f=font(10, True, c=MUTED), bg=GRAY, h="center")
        put(ws, f"C{r}", feito, merge=f"C{r}:{last}{r}")
        ws.row_dimensions[r].height = 21.75
    return r


# ================================================================== aba Canais (folha de canais)
N = 32
K_W = {"A": 2, "B": 7, "C": 30, "D": 14, "E": 10, "F": 10, "G": 8, "H": 8, "I": 14, "J": 11, "K": 11, "L": 19,
       "M": 46, "N": 26, "O": 2, "P": 10, "Q": 10, "R": 11, "S": 2}
K1 = 14
VERMELHOS_K = ["Falta entrada", "Entrada inexistente", "Entrada repetida", "Falta phantom"]


def aba_canais(ws, tab, data=None):
    widths(ws, K_W)
    is_ex = data is not None
    title(ws, "Folha de canais — o que entra em cada canal, e para onde vai",
          "Exemplo preenchido, como estava antes da correção. Use a aba Canais para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro. As células cinza são calculadas.", "R")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    cabecalho(ws, H, hb, [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                          (7, "Operador", "operador", None), (8, "Mesa", "mesa", None)])
    dv_date(ws, "D6")
    label(ws, "J4", "Entradas físicas", merge="J4:K4")
    inp(ws, "L4", data["nent"] if is_ex else 16, h="center", bg=hb)
    dv_number(ws, "L4")
    hint(ws, "M4", "Quantas entradas a mesa ou a caixa de palco tem. Usado para conferir os números.", merge="M4:R4", h="left")
    NENT = "$L$4"
    hint(ws, "J6", "A entrada física é o número do conector (na mesa ou na caixa de palco) ligado ao canal. "
                   "Na mesa analógica, é o próprio número do canal. Na digital, é o que o patch diz.", merge="J6:R8", h="left")
    for r in (6, 7, 8):
        ws.row_dimensions[r].height = 21.75

    band(ws, 10, "Folha de canais (input list)", "N")
    put(ws, "P10", "Entradas físicas", f=font(10, True, c=WHITE), bg=INK, box=False, merge="P10:R10")
    heads = [("B", "Canal", INK), ("C", "Fonte", INK), ("D", "Tipo", BLUE), ("E", "Entrada física", BLUE),
             ("F", "Phantom ligado", BLUE), ("G", "HPF", BLUE), ("H", "HPF (Hz)", BLUE), ("I", "Destino", TEAL),
             ("J", "DCA", TEAL), ("K", "Precisa de phantom", MUTED), ("L", "Situação", MUTED),
             ("M", "O que fazer", MUTED), ("N", "Observação", INK),
             ("P", "Entrada", INK), ("Q", "Canais que usam", MUTED), ("R", "Situação", MUTED)]
    for col, t, bg in heads:
        head(ws, f"{col}11", t, bg=bg)
    ws.row_dimensions[11].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "o que está ligado", "lista", "nº do conector", "Sim / Não",
                                        "Sim / Não", "se ajustável", "lista", "nome", "calculado", "", "", ""]):
        hint(ws, f"{col}12", t)
    for col, t in zip("PQR", ["nº", "contagem", ""]):
        hint(ws, f"{col}12", t)
    ws.row_dimensions[12].height = 21.75
    if is_ex:
        hint(ws, "B13", "Os dados são os do exemplo 1 do treinamento, como estavam na mesa antes da correção. As colunas de K a M são calculadas.",
             merge="B13:N13", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Coral (condensador suspenso)", "Condensador", 6, "Não", "Sim", 120,
                                           "Grupo 1-2", "Vozes", "Sim", "Falta phantom", "Ligue o phantom com o canal em mute"]):
            ex(ws, f"{col}13", v, h="left" if col in "CM" else "center")
        ex(ws, "N13", "")
        note(ws, "D11", "Tipo de fonte. Condensador e DI ativa precisam de phantom.\nFita: não ligue phantom sem conferir o manual do microfone.\nUSB ou rede: sinal que chega sem conector de áudio.")
        note(ws, "E11", "Número do conector físico ligado ao canal.\nNa mesa digital, é a entrada escolhida no patch.\nDeixe em branco só para fontes por USB ou rede.")
        note(ws, "I11", "Barramento que recebe o canal depois do fader.\nLR + grupo: o canal vai ao master direto e também pelo grupo, o que dá caminho duplo.\nSó auxiliares: canal que não vai às caixas da frente (comunicação com o palco, por exemplo).")
        note(ws, "K11", "Calculado pelo tipo: Sim para condensador e DI ativa.\nHá condensadores que funcionam com pilha; nesse caso, anote na Observação.")
    ws.row_dimensions[13].height = 24

    K2 = K1 + N - 1
    E_RNG = f"$E${K1}:$E${K2}"
    bg = WHITE if is_ex else INPUT
    for k in range(N):
        r = K1 + k
        c = data["canais"][k] if is_ex and k < len(data["canais"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        for col, j in zip("DEFGHIJ", range(1, 8)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg)
        C, D, E, F, I, Kc, L = (f"{x}{r}" for x in "CDEFIKL")
        calc(ws, Kc, f'=IF({D}="","",IF(OR({D}="Condensador",{D}="DI ativa"),"Sim","Não"))', b=False)
        calc(ws, L, f'=IF(AND({C}="",{D}="",{E}=""),"",IF(AND({E}="",{D}<>"USB ou rede"),"Falta entrada",'
                    f'IF(AND({E}<>"",{NENT}<>"",{E}>{NENT}),"Entrada inexistente",'
                    f'IF(AND({E}<>"",COUNTIF({E_RNG},{E})>1),"Entrada repetida",'
                    f'IF({D}="","Falta tipo",'
                    f'IF(AND({Kc}="Sim",{F}<>"Sim"),"Falta phantom",'
                    f'IF(AND({Kc}="Não",{F}="Sim"),"Phantom sem uso",'
                    f'IF({I}="LR + grupo","Caminho duplo",'
                    f'IF(OR({I}="",{I}="Nenhum"),"Sem destino","OK")))))))))')
        calc(ws, f"M{r}", f'=IF({L}="","",IF({L}="Falta entrada","Anote o conector ligado ao canal",'
                          f'IF({L}="Entrada inexistente","A mesa não tem essa entrada: confira o número",'
                          f'IF({L}="Entrada repetida","Dois canais na mesma entrada: confira o patch e os cabos",'
                          f'IF({L}="Falta tipo","Informe o tipo da fonte",'
                          f'IF({L}="Falta phantom","Ligue o phantom com o canal em mute",'
                          f'IF({L}="Phantom sem uso","Desligue o phantom com o canal em mute",'
                          f'IF({L}="Caminho duplo","Deixe um só caminho até o master: tire o L-R ou o grupo",'
                          f'IF({L}="Sem destino","Enderece o canal ao master ou a um grupo","Manter")))))))))',
             h="left", b=False, sz=9)
        inp(ws, f"N{r}", bg=bg)
        # bloco das entradas físicas
        num(ws, f"P{r}", k + 1)
        calc(ws, f"Q{r}", f'=IF(OR({NENT}="",P{r}>{NENT}),"",COUNTIF({E_RNG},P{r}))', b=False)
        calc(ws, f"R{r}", f'=IF(Q{r}="","–",IF(Q{r}=0,"Livre",IF(Q{r}=1,"OK","Repetida")))', b=False)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{K1}:D{K2}", TIPOS, "Tipo da fonte.")
    dv_number(ws, f"E{K1}:E{K2}")
    dv_list(ws, f"F{K1}:G{K2}", SIM_NAO, "Sim ou Não.")
    dv_number(ws, f"H{K1}:H{K2}")
    dv_list(ws, f"I{K1}:I{K2}", DESTINOS, "Barramento que recebe o canal depois do fader.")
    cf_situacao(ws, f"L{K1}:L{K2}", f"L{K1}", VERMELHOS_K)
    cf_equal(ws, f"R{K1}:R{K2}", [("OK", GREEN), ("Repetida", RED)])

    s = K2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(L{K1}:L{K2},"{t}")'
    avisos = ["Falta entrada", "Entrada inexistente", "Entrada repetida", "Falta tipo", "Falta phantom",
              "Phantom sem uso", "Caminho duplo", "Sem destino"]
    linhas = [
        ("Canais registrados", f"=COUNTA(C{K1}:C{K2})"),
        ("Canais OK", f'={cont("OK")}'),
        ("Canais com aviso", "=" + "+".join(cont(t) for t in avisos)),
        ("Canais com phantom ligado", f'=COUNTIF(F{K1}:F{K2},"Sim")'),
        ("Entradas físicas repetidas", f'=COUNTIF(R{K1}:R{K2},"Repetida")'),
        ("Entradas físicas livres", f'=COUNTIF(R{K1}:R{K2},"Livre")'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", val_merge=f"F{s + k}:H{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste os canais",IF(F{s + 5}>0,"Há entrada física em dois canais: confira o patch",'
                      f'IF(F{s + 3}>0,"Há canal com aviso: corrija antes do teste de linha","OK")))',
         b=False, merge=f"F{a}:M{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:M{a}")
    end = a
    if is_ex:
        end = bloco_ajustes(ws, end, data["ajustes"], "N")
    ws.freeze_panes = f"D{K1}"
    setup(ws, tab, f"B1:R{end}", fit_height=True)


# ================================================================== aba Saídas e auxiliares
NS = 20
S_W = {"A": 2, "B": 5, "C": 16, "D": 13, "E": 12, "F": 18, "G": 12, "H": 30, "I": 11, "J": 13, "K": 20, "L": 50,
       "M": 26, "N": 2}
S1 = 13
VERMELHOS_S = ["Trocar para pré", "Trocar para pós", "Grupo fora do master", "Saída repetida"]


def aba_saidas(ws, tab):
    widths(ws, S_W)
    title(ws, "Saídas e auxiliares — o mapa dos barramentos",
          "Um barramento por linha: o que ele é, como recebe os canais, para que serve e onde termina.", "M")
    cabecalho(ws, {}, INPUT, [(4, "Evento ou local", "evento", None), (5, "Data", "data", DATE),
                              (6, "Mesa", "mesa", None)])
    dv_date(ws, "D5")
    hint(ws, "I4", "A tomada indicada vem do uso: retorno e fone em pré-fader; efeito, subwoofer, outra sala e delay em pós-fader; "
                   "gravação e transmissão dependem de quem mixa.", merge="I4:M6", h="left")

    band(ws, 8, "Mapa de saídas e auxiliares", "M")
    heads = [("B", "#", INK), ("C", "Barramento", INK), ("D", "Tipo", TEAL), ("E", "Tomada", TEAL), ("F", "Uso", TEAL),
             ("G", "Saída física", AMBER), ("H", "Destino", AMBER), ("I", "Grupo vai ao master", TEAL),
             ("J", "Tomada indicada", MUTED), ("K", "Situação", MUTED), ("L", "O que fazer", MUTED), ("M", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKLM", ["", "nome na mesa", "lista", "só auxiliar", "lista", "conector",
                                       "o que está ligado", "só grupo", "calculada", "", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    for col, v in zip("BCDEFGHIJKL", ["Ex.", "Aux 1", "Auxiliar", "Pós-fader", "Retorno de palco", "Aux 1",
                                      "Caixa de chão da cantora", "–", "Pré-fader", "Trocar para pré",
                                      "Retorno em pós muda a cada ajuste do fader: passe para pré"]):
        ex(ws, f"{col}11", v, h="left" if col in "HL" else "center")
    ex(ws, "M11", "")
    ws.row_dimensions[11].height = 24
    note(ws, "E9", "Só para auxiliar: o envio sai antes (pré) ou depois (pós) do fader do canal.\nVeja a chave PRE/POST da mesa, ou a tela da digital.")
    note(ws, "G9", "Conector por onde o barramento sai: Main L/R, Aux 1, Out 5, XLR 3...\nGrupo que só vai ao master pode ficar em branco.")
    note(ws, "I9", "Só para grupo: o botão que envia o grupo ao master está apertado?\nSem ele, os canais do grupo não saem nas caixas.")

    S2 = S1 + NS - 1
    G_RNG = f"$G${S1}:$G${S2}"
    for k in range(NS):
        r = S1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDEFGHI":
            inp(ws, f"{col}{r}", h="left" if col in "CH" else "center")
        C, D, E, F, G, I, J, K = (f"{x}{r}" for x in "CDEFGIJK")
        calc(ws, J, f'=IF({D}<>"Auxiliar","",IF({F}="","",IF(OR({F}="Retorno de palco",{F}="Fone ou in-ear"),"Pré-fader",'
                    f'IF(OR({F}="Gravação",{F}="Transmissão"),"Qualquer","Pós-fader"))))', b=False)
        calc(ws, K, f'=IF(AND({C}="",{D}=""),"",IF({D}="","Falta tipo",IF({F}="","Falta uso",'
                    f'IF(AND({D}="Auxiliar",{E}=""),"Falta tomada",'
                    f'IF(AND({D}="Auxiliar",{J}="Pré-fader",{E}="Pós-fader"),"Trocar para pré",'
                    f'IF(AND({D}="Auxiliar",{J}="Pós-fader",{E}="Pré-fader"),"Trocar para pós",'
                    f'IF(AND({D}="Grupo",{I}<>"Sim"),"Grupo fora do master",'
                    f'IF(AND({D}<>"Grupo",{G}=""),"Falta saída",'
                    f'IF(AND({G}<>"",COUNTIF({G_RNG},{G})>1),"Saída repetida","OK")))))))))')
        calc(ws, f"L{r}", f'=IF({K}="","",IF({K}="Falta tipo","Informe o tipo do barramento",'
                          f'IF({K}="Falta uso","Informe para que serve",'
                          f'IF({K}="Falta tomada","Confira a chave pré ou pós do auxiliar",'
                          f'IF({K}="Trocar para pré","Retorno em pós muda a cada ajuste do fader: passe para pré",'
                          f'IF({K}="Trocar para pós","Em pré, o canal fechado continua no efeito: passe para pós",'
                          f'IF({K}="Grupo fora do master","Enderece o grupo ao master, ou os canais dele não saem",'
                          f'IF({K}="Falta saída","Anote o conector de saída",'
                          f'IF({K}="Saída repetida","Duas linhas no mesmo conector: confira","Manter")))))))))',
             h="left", b=False, sz=9)
        inp(ws, f"M{r}")
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{S1}:D{S2}", TIPOS_BUS, "Master LR, Grupo, Auxiliar ou Matriz.")
    dv_list(ws, f"E{S1}:E{S2}", TOMADAS, "Pré-fader ou Pós-fader (só para auxiliar).")
    dv_list(ws, f"F{S1}:F{S2}", USOS, "Para que serve o barramento.")
    dv_list(ws, f"I{S1}:I{S2}", SIM_NAO, "O grupo está endereçado ao master?")
    cf_situacao(ws, f"K{S1}:K{S2}", f"K{S1}", VERMELHOS_S)

    s = S2 + 2
    band(ws, s, "Resumo automático", "M")
    cont = lambda t: f'COUNTIF(K{S1}:K{S2},"{t}")'
    avisos = ["Falta tipo", "Falta uso", "Falta tomada", "Trocar para pré", "Trocar para pós", "Grupo fora do master",
              "Falta saída", "Saída repetida"]
    linhas = [
        ("Barramentos registrados", f'=COUNTA(D{S1}:D{S2})'),
        ("Auxiliares em pré-fader", f'=COUNTIFS(D{S1}:D{S2},"Auxiliar",E{S1}:E{S2},"Pré-fader")'),
        ("Auxiliares em pós-fader", f'=COUNTIFS(D{S1}:D{S2},"Auxiliar",E{S1}:E{S2},"Pós-fader")'),
        ("Barramentos com aviso", "=" + "+".join(cont(t) for t in avisos)),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "F", "G")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:F{a}", h="right")
    calc(ws, f"G{a}", f'=IF(G{s + 1}=0,"Liste os barramentos",IF(G{s + 4}>0,"Há barramento com aviso: corrija antes de montar os retornos","OK"))',
         b=False, merge=f"G{a}:L{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"G{a}:L{a}")
    ws.freeze_panes = f"D{S1}"
    setup(ws, tab, f"B1:M{a}", fit_height=True)


# ================================================================== aba Cenas
NC = 20
C_W = {"A": 2, "B": 5, "C": 10, "D": 28, "E": 13, "F": 26, "G": 44, "H": 12, "I": 18, "J": 28, "K": 2}
C1 = 12


def aba_cenas(ws, tab):
    widths(ws, C_W)
    title(ws, "Cenas — o registro das cenas da mesa digital",
          "Uma cena por linha. Anote também o que ela muda e se há cópia fora da mesa.", "J")
    cabecalho(ws, {}, INPUT, [(4, "Local", "local", None), (5, "Mesa", "mesa", None)])
    label(ws, "B6", "Cena de partida (nº)", merge="B6:C6")
    inp(ws, "D6", h="center")
    dv_number(ws, "D6")
    calc(ws, "E6", f'=IF(D6="","Informe o número da cena de partida",IF(COUNTIF(C{C1}:C{C1 + NC - 1},D6)=0,'
                   f'"A cena de partida não está registrada na tabela","OK"))', b=False, merge="E6:H6", h="left")
    cf_ok(ws, "E6:H6")
    ws.row_dimensions[6].height = 21.75
    hint(ws, "I4", "A cena de partida guarda o patch, os nomes e as saídas do local, com o resto zerado. "
                   "Grave por cima dela só de propósito.", merge="I4:J6", h="left")

    band(ws, 8, "Registro de cenas", "J")
    heads = [("B", "#", INK), ("C", "Nº da cena", BLUE), ("D", "Nome", BLUE), ("E", "Data", BLUE),
             ("F", "Evento ou uso", BLUE), ("G", "O que ela muda ou protege", BLUE), ("H", "Cópia fora da mesa", BLUE),
             ("I", "Situação", MUTED), ("J", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJ", ["", "na mesa", "", "dd/mm/aaaa", "", "ganhos, retornos, filtro de cena", "Sim / Não", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    for col, v in zip("BCDEFGHI", ["Ex.", 12, "Culto 04-10 passagem", "04/10/2026", "Culto de domingo, 19 h",
                                   "Ganhos e retornos da banda; filtro protege o canal do púlpito", "Sim", "OK"]):
        ex(ws, f"{col}11", v, h="left" if col in "DFG" else "center")
    ex(ws, "J11", "")
    ws.row_dimensions[11].height = 24

    C2 = C1 + NC - 1
    for k in range(NC):
        r = C1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center")
        inp(ws, f"D{r}")
        inp(ws, f"E{r}", h="center", fmt=DATE)
        inp(ws, f"F{r}")
        inp(ws, f"G{r}")
        inp(ws, f"H{r}", h="center")
        C, D, E, H = (f"{x}{r}" for x in "CDEH")
        calc(ws, f"I{r}", f'=IF(AND({C}="",{D}="",{E}=""),"",IF({C}="","Falta número",'
                          f'IF(COUNTIF($C${C1}:$C${C2},{C})>1,"Número repetido",IF({E}="","Falta data",'
                          f'IF({H}<>"Sim","Sem cópia","OK")))))')
        inp(ws, f"J{r}")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"C{C1}:C{C2}")
    dv_date(ws, f"E{C1}:E{C2}")
    dv_list(ws, f"H{C1}:H{C2}", SIM_NAO, "Há cópia da cena fora da mesa (pendrive, computador)?")
    cf_situacao(ws, f"I{C1}:I{C2}", f"I{C1}", ["Número repetido"])

    s = C2 + 2
    band(ws, s, "Resumo automático", "J")
    summary(ws, s + 1, "Cenas registradas", f'=COUNTA(D{C1}:D{C2})', "D", "E")
    summary(ws, s + 2, "Cenas sem cópia fora da mesa", f'=COUNTIF(I{C1}:I{C2},"Sem cópia")', "D", "E")
    summary(ws, s + 3, "Cena mais recente (data)", f'=IF(COUNT(E{C1}:E{C2})=0,"",MAX(E{C1}:E{C2}))', "D", "E", fmt=DATE)
    a = s + 4
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(E{s + 1}=0,"Registre as cenas",IF(COUNTIF(I{C1}:I{C2},"OK")<E{s + 1},'
                      f'"Há cena com aviso: veja a coluna Situação","OK"))', b=False, merge=f"E{a}:H{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:H{a}")
    ws.freeze_panes = f"C{C1}"
    setup(ws, tab, f"B1:J{a}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "A mesa de som e o caminho do sinal — modelo de folha de canais"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Mesa de som — como usar esta planilha",
           "Folha de canais, mapa de saídas e auxiliares e registro de cenas, para preparar a mesa antes do teste de linha.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo, destino, tomada, uso e as colunas Sim / Não aceitam só as opções da lista."),
        ("Situação", "Verde: OK. Amarelo: confira. Vermelho: corrija antes do teste de linha."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Canais: preencha o cabeçalho e o número de entradas físicas da mesa ou da caixa de palco."),
        ("2", "Aba Canais: um canal por linha, com a fonte, o tipo, a entrada física, o phantom, o filtro, o destino e o DCA."),
        ("3", "Veja a situação de cada canal e o bloco Entradas físicas, à direita. Corrija na mesa e na planilha."),
        ("4", "Aba Saídas e auxiliares: um barramento por linha, com a tomada, o uso, a saída física e o destino."),
        ("5", "Na mesa: zere, carregue a cena de partida (na digital) e faça o teste de linha com a folha na mão."),
        ("6", "Aba Cenas (mesa digital): registre a cena salva ao fim da passagem e confirme a cópia fora da mesa."),
        ("7", "Aba Checklist: confira os doze itens antes do teste de linha e de novo ao fim da passagem."),
    ]),
    ("Abas da planilha", [
        ("Canais", "Folha de canais com até 32 canais, situação calculada e conferência das entradas físicas."),
        ("Saídas e auxiliares", "Mapa de até 20 barramentos, com a tomada indicada para cada uso e a situação."),
        ("Cenas", "Registro de até 20 cenas, com número repetido, falta de data e falta de cópia."),
        ("Referência", "Controles do canal e posição ao zerar, barramentos, tomada por uso e o roteiro “siga o sinal”."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Canais", "A folha de canais do exemplo 1 do treinamento, antes da correção, e o que foi corrigido."),
    ]),
    ("Regras e premissas do modelo", [
        ("Precisa de phantom", "Sim para condensador e DI ativa; Não para os outros tipos. Condensador com pilha: anote na Observação e ignore o aviso."),
        ("Phantom sem uso", "Em microfone dinâmico balanceado, o phantom costuma não causar dano, mas o modelo avisa: desligue o que não precisa. Em microfone de fita e em aparelho de linha ligado por XLR, há risco real."),
        ("Entrada repetida", "Dois canais com a mesma entrada física. Às vezes é proposital (um canal para a sala e outro para a transmissão): anote na Observação."),
        ("Caminho duplo", "Destino LR + grupo: o canal chega ao master duas vezes, até 6 dB mais alto, e o fader do grupo não o controla."),
        ("Tomada indicada", "Retorno e fone em pré-fader. Efeito, subwoofer, outra sala, caixas de delay e conjunto de canais em pós-fader. Gravação e transmissão: qualquer uma, conforme quem mixa."),
        ("Valores", "As regras são práticas comuns, e não norma. Confira o manual e o diagrama de blocos da sua mesa."),
        ("Exemplos", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_canais(wb.create_sheet("Canais"), BLUE)
aba_saidas(wb.create_sheet("Saídas e auxiliares"), TEAL)
aba_cenas(wb.create_sheet("Cenas"), AMBER)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 36, "D": 30, "E": 54, "F": 2})
title(ws, "Referência — a mesa de som", "Consulta rápida. Confira o manual e o diagrama de blocos da sua mesa.", "E")


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


r = tabela(4, "Os controles do canal, de cima para baixo", ["Controle", "O que faz", "Ao zerar a mesa", "Onde se aprofunda"], [
    ("Phantom (+48 V)", "Alimenta condensadores e DI ativas pelo cabo XLR.", "Desligado", "Cabos, conectores e níveis de sinal"),
    ("Pad", "Reduz o sinal antes do pré, para fontes fortes.", "Desligado", "Ganho e fader"),
    ("Ganho (gain, trim)", "Define com que nível o sinal entra na mesa.", "No mínimo", "Ganho e fader"),
    ("Polaridade (Ø)", "Inverte a onda, para dois microfones que se cancelam.", "Normal", "Som, frequência e decibel"),
    ("Filtro passa-altas (HPF)", "Corta os graves abaixo de uma frequência.", "Desligado; religar pela folha de canais", "Equalização"),
    ("Equalizador (EQ)", "Corta ou realça faixas de frequência.", "Plano: botões no centro", "Equalização"),
    ("Auxiliares (aux, send)", "Mandam uma cópia do canal a cada mixagem extra.", "Fechados", "Monitores e retornos; Dinâmica e efeitos"),
    ("Pan", "Posição do canal entre esquerda e direita.", "No centro", "Passagem de som e mixagem ao vivo"),
    ("Mute", "Cala o canal no master e nos envios pós-fader (e, em muitas mesas, nos pré-fader).", "Ligado", "Teste na sua mesa"),
    ("PFL ou solo", "Leva o canal ao fone e ao medidor, sem mudar as caixas.", "Solto", "Ganho e fader"),
    ("Endereçamento (L-R, 1-2, 3-4)", "Escolhe os barramentos que recebem o canal depois do fader.", "Só L-R, ou só o grupo", "Este estudo"),
    ("Fader", "Quanto do canal vai aos barramentos escolhidos.", "Embaixo", "Ganho e fader"),
    ("Master, grupos e masters dos auxiliares", "Nível geral de cada barramento.", "Embaixo", "Ganho e fader"),
])
r = tabela(r, "Os barramentos", ["Barramento", "O que soma", "Controle geral", "Uso típico"], [
    ("Master LR", "Os canais e os grupos endereçados a ele.", "Fader master", "Caixas da frente: a mixagem do público."),
    ("Grupo (subgrupo)", "Os canais endereçados a ele. Desemboca no master.", "Fader do grupo", "Controlar e processar um conjunto: vozes, bateria, coral."),
    ("Auxiliar", "Uma cópia de cada canal, na medida do botão de envio.", "Master do auxiliar", "Retornos, efeitos e mixagens independentes."),
    ("Matriz", "Outros barramentos: master, grupos e, em muitas mesas, auxiliares.", "Master da matriz", "Outra sala, caixas de delay, gravação, transmissão."),
    ("DCA ou VCA", "Nada: não carrega áudio. Comanda os faders atribuídos.", "Fader do DCA", "Subir e baixar vários canais de uma vez, com os efeitos pós-fader juntos."),
])
r = tabela(r, "A tomada de cada auxiliar", ["Uso", "Tomada", "Por quê", "Se estiver trocada"], [
    ("Retorno de chão ou fone", "Pré-fader", "A mixagem do músico não muda quando o operador ajusta a sala.", "O retorno sobe e desce com o fader."),
    ("Efeito (reverb, delay)", "Pós-fader", "O efeito acompanha o canal.", "O canal fechado continua soando no efeito."),
    ("Subwoofer por auxiliar", "Pós-fader", "Os graves sobem e descem com a mixagem.", "O grave não acompanha os faders."),
    ("Outra sala ou caixas de delay, sem matriz", "Pós-fader", "A sala anexa ouve a mesma mixagem, com os mesmos mutes.", "Canais fechados no salão continuam na outra sala."),
    ("Gravação ou transmissão", "Depende", "Pós-fader para seguir a sala; pré-fader se alguém mixa só para isso.", "Veja o estudo Transmissão ao vivo."),
])
r = tabela(r, "Siga o sinal: canal sem som", ["O que você vê", "Onde procurar", "O que conferir", "Observação"], [
    ("Sem sinal no PFL do canal", "Antes do canal ou na entrada", "Fonte, cabo, entrada, patch, phantom, ganho.", "Na digital, desconfie primeiro do patch."),
    ("Sinal no PFL, nada no master", "Parte de baixo do canal", "Mute, fader, L-R, grupo fechado ou fora do master, DCA, grupo de mute.", "Na digital: camada errada ou envios nos faders ligado."),
    ("Sinal no master, nada nas caixas", "Depois da mesa", "Patch de saída, cabo, processador, amplificador ou caixa ativa.", "Feche o master antes de mexer em cabos."),
    ("Som nas caixas, nada no retorno ou no efeito", "No auxiliar", "Envio do canal, master do auxiliar, chave pré ou pós, saída e cabo.", "Use o AFL do auxiliar para ouvir o barramento."),
])
hint(ws, f"B{r - 1}", "Práticas comuns, e não norma. A ordem dos blocos e o efeito do mute variam de uma mesa para outra: confira o diagrama de blocos.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist da mesa",
          "Confira antes do teste de linha e de novo ao fim da passagem. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_canais(wb.create_sheet("Exemplo - Canais"), MUTED, EX_CANAIS)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
