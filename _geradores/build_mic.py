# -*- coding: utf-8 -*-
"""Gera Microfones-modelo.xlsx: inventário, escolha por fonte, regra 3:1, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE = HUE["blue"][0]
AMBER = HUE["amber"][0]
TEAL = HUE["teal"][0]
SIM_NAO = ["Sim", "Não"]
FORMATOS = ["De mão", "Lapela", "Headset", "Gooseneck", "Coral suspenso", "Lápis", "Clip", "Contato", "Instrumento"]
TIPOS = ["Dinâmico", "Condensador", "Fita"]
PADROES = ["Cardioide", "Supercardioide", "Hipercardioide", "Omni", "Figura de 8"]
LIGACAO = ["Com fio", "Sem fio"]
TESTE = ["OK", "Defeito", "Não testado"]
FONTES = ["Voz cantada", "Voz falada", "Coral", "Violão", "Amplificador", "Bumbo", "Caixa e tons",
          "Chimbal e pratos", "Piano", "Ambiente"]

CHECK = [
    ("Escolha", "Cada fonte tem um microfone definido, com tipo e formato adequados a ela."),
    ("Escolha", "O padrão polar de cada microfone é conhecido, e cada retorno está na direção de maior rejeição do microfone à frente dele."),
    ("Escolha", "O phantom está ligado só nos canais que precisam dele, e foi ligado com o canal fechado."),
    ("Escolha", "Todos os microfones foram testados com a voz, um por vez, com o cabo e o canal que vão usar."),
    ("Posição", "Cada microfone está o mais perto da fonte que o formato permite, e apontado para ela."),
    ("Posição", "Os microfones que ficam abertos juntos respeitam a regra 3:1."),
    ("Posição", "Nenhum microfone fica à frente das caixas do PA nem apontado para um retorno."),
    ("Posição", "A lapela e o headset estão firmes, na posição certa, com a cápsula descoberta e o cabo preso."),
    ("Uso", "Quem usa microfone de mão foi orientado: perto da boca, apontado para ela, mão no corpo."),
    ("Uso", "Há espuma ou protetor de vento onde é preciso, e o filtro passa-altas está ligado nos canais de voz."),
    ("Uso", "Só ficam abertos os microfones em uso, e há um microfone reserva com fio, testado."),
    ("Depois", "Depois do evento, os microfones foram limpos, secos e guardados no estojo, e os defeitos foram anotados no inventário."),
]

# Ponto de partida por fonte: fonte, tipo, padrão, formato, distância, posição, tipo aceito, padrão aceito
REF_FONTES = [
    ("Voz cantada", "Dinâmico", "Cardioide ou supercardioide", "De mão", "2 a 5 cm",
     "Apontado para a boca. Retorno na direção de maior rejeição do padrão.", "Qualquer", "Direcional"),
    ("Voz falada", "Dinâmico ou condensador", "Conforme o formato", "De mão, headset, gooseneck ou lapela",
     "O mais perto que o formato permite", "Um microfone por pessoa que fala, e só o dela aberto.", "Qualquer", "Qualquer"),
    ("Coral", "Condensador", "Cardioide", "Coral suspenso ou pedestal alto", "60 a 90 cm à frente da 1ª fila",
     "Acima das cabeças, apontado para a última fila. Um para cada 15 a 20 vozes.", "Condensador", "Direcional"),
    ("Violão", "Condensador", "Cardioide", "Lápis", "15 a 30 cm",
     "Apontado para a junção do braço com o corpo, e não para a boca do violão.", "Condensador", "Direcional"),
    ("Amplificador", "Dinâmico", "Cardioide", "De mão ou instrumento", "2 a 10 cm da tela",
     "Apontado para o alto-falante: no centro, mais brilho; perto da borda, mais corpo.", "Dinâmico", "Direcional"),
    ("Bumbo", "Dinâmico", "Cardioide", "Instrumento (próprio para graves)", "Dentro do casco ou no furo da pele",
     "Perto da pele batida, mais ataque. No furo, mais corpo.", "Dinâmico", "Direcional"),
    ("Caixa e tons", "Dinâmico", "Cardioide ou supercardioide", "De mão ou clip", "3 a 5 cm acima da pele",
     "Junto ao aro, apontado para o centro da pele. Na caixa, ponto surdo para o chimbal.", "Dinâmico", "Direcional"),
    ("Chimbal e pratos", "Condensador", "Cardioide", "Lápis", "Chimbal: 10 a 15 cm. Pratos: 60 a 100 cm",
     "Chimbal de cima, do lado oposto à caixa. Pratos: um ou dois apontados para baixo.", "Condensador", "Direcional"),
    ("Piano", "Condensador", "Cardioide", "Lápis", "15 a 30 cm das cordas",
     "Tampa aberta. Com dois: um sobre as cordas agudas, outro sobre as graves.", "Condensador", "Direcional"),
    ("Ambiente", "Condensador", "Cardioide", "Lápis", "Longe das caixas do PA",
     "Voltado para o público. Só para a transmissão ou a gravação, fechado no PA.", "Condensador", "Qualquer"),
]
REF_R1 = 6  # primeira linha da tabela de fontes na aba Referência
REF_RANGE = f"Referência!$B${REF_R1}:$I${REF_R1 + len(REF_FONTES) - 1}"

# Exemplo 1 do treinamento: inventário de dez microfones e as nove fontes como foram planejadas
EX_INV = [  # nº, identificação, formato, tipo, padrão, ligação, teste, observação
    (1, "Mão 1 (fita azul)", "De mão", "Dinâmico", "Cardioide", "Com fio", "OK", ""),
    (2, "Mão 2 (fita verde)", "De mão", "Dinâmico", "Cardioide", "Com fio", "OK", ""),
    (3, "Mão 3 (fita amarela)", "De mão", "Dinâmico", "Supercardioide", "Com fio", "OK", "Retorno em diagonal"),
    (4, "Mão 4 (sem fita)", "De mão", "Dinâmico", "Cardioide", "Com fio", "Defeito", "Falha ao mexer no conector"),
    (5, "Headset do pastor", "Headset", "Condensador", "Omni", "Sem fio", "OK", "Pilhas novas a cada culto"),
    (6, "Coral 1", "Coral suspenso", "Condensador", "Cardioide", "Com fio", "OK", ""),
    (7, "Coral 2", "Coral suspenso", "Condensador", "Cardioide", "Com fio", "OK", ""),
    (8, "Lápis 1", "Lápis", "Condensador", "Cardioide", "Com fio", "OK", ""),
    (9, "Bumbo", "Instrumento", "Dinâmico", "Cardioide", "Com fio", "OK", "Próprio para graves"),
    (10, "Púlpito", "Gooseneck", "Condensador", "Cardioide", "Com fio", "OK", ""),
]
EX_ESCOLHA = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala"},
    "fontes": [  # canal, tipo de fonte, descrição, microfone nº
        (1, "Voz falada", "Pastor", 5),
        (2, "Voz cantada", "Cantora", 1),
        (3, "Voz cantada", "Voz de apoio 1", 2),
        (4, "Voz cantada", "Voz de apoio 2", 4),
        (5, "Coral", "Coral, lado esquerdo", 3),
        (6, "Coral", "Coral, lado direito", 2),
        (7, "Violão", "Violão sem captação", 8),
        (8, "Bumbo", "Bumbo", 9),
        (9, "Voz falada", "Púlpito", 10),
    ],
    "ajustes": [  # canal, o que foi feito
        ("4", "Microfone nº 4 com defeito: entrou o nº 3, supercardioide, e o retorno dessa voz foi para a diagonal."),
        ("5", "Trocado pelo condensador suspenso nº 6, a 80 cm da primeira fila. Phantom ligado."),
        ("6", "Trocado pelo condensador suspenso nº 7, a 2,7 m do nº 6. Phantom ligado."),
        ("Todos", "Depois das trocas: nove fontes com microfone certo, phantom em quatro canais (5, 6, 7 e 9), nº 4 separado para conserto."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def cabecalho(ws, H, hb, ate):
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:{ate}{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")


# ================================================================== linhas de inventário (B:J)
def inv_linhas(ws, r1, n, dados=None, completo=True):
    """Escreve n linhas de inventário a partir de r1, nas colunas B a J (e K a N se completo).
    B nº, C identificação, D formato, E tipo, F padrão, G ligação, H teste, I última limpeza, J phantom."""
    bg = WHITE if dados else INPUT
    for k in range(n):
        r = r1 + k
        d = dados[k] if dados and k < len(dados) else None
        num(ws, f"B{r}", d[0] if d else k + 1)
        inp(ws, f"C{r}", d[1] if d else None, bg=bg)
        for col, j in zip("DEFGH", range(2, 7)):
            inp(ws, f"{col}{r}", d[j] if d else None, h="center", bg=bg)
        inp(ws, f"I{r}", None, h="center", bg=bg, fmt=DATE)
        E, F, G, H = (f"{c}{r}" for c in "EFGH")
        calc(ws, f"J{r}", f'=IF({E}="","",IF({G}="Sem fio","Não",IF({E}="Condensador","Sim",'
                          f'IF({E}="Fita","Não: veja o manual","Não"))))')
        if completo:
            calc(ws, f"K{r}", f'=IF({F}="","",IF({F}="Cardioide","180°, bem atrás",IF({F}="Supercardioide","cerca de 125°",'
                              f'IF({F}="Hipercardioide","cerca de 110°",IF({F}="Omni","não rejeita","90°, nos lados")))))', b=False)
            calc(ws, f"L{r}", f'=IF({F}="","",IF({F}="Cardioide","Bem atrás do microfone",IF(OR({F}="Supercardioide",{F}="Hipercardioide"),'
                              f'"Em diagonal, nunca bem atrás",IF({F}="Omni","Longe: o microfone não rejeita","Em um dos lados"))))',
                 h="left", b=False, sz=9)
            calc(ws, f"M{r}", f'=IF(AND(C{r}="",D{r}="",{E}="",{F}=""),"",IF({H}="Defeito","Defeito",'
                              f'IF(OR({H}="",{H}="Não testado"),"Testar",IF(OR(D{r}="",{E}="",{F}=""),"Incompleto","OK"))))')
            inp(ws, f"N{r}", d[7] if d else None, bg=bg)
        ws.row_dimensions[r].height = 21.75
    r2 = r1 + n - 1
    dv_list(ws, f"D{r1}:D{r2}", FORMATOS, "Formato do microfone.")
    dv_list(ws, f"E{r1}:E{r2}", TIPOS, "Dinâmico, Condensador ou Fita.")
    dv_list(ws, f"F{r1}:F{r2}", PADROES, "Padrão polar, conforme o manual do modelo.")
    dv_list(ws, f"G{r1}:G{r2}", LIGACAO, "Com fio ou Sem fio.")
    dv_list(ws, f"H{r1}:H{r2}", TESTE, "Resultado do último teste com a voz.")
    dv_date(ws, f"I{r1}:I{r2}")
    return r2


# ================================================================== aba Inventário
I_W = {"A": 2, "B": 6, "C": 24, "D": 15, "E": 13, "F": 15, "G": 10, "H": 12, "I": 12, "J": 12, "K": 15, "L": 26,
       "M": 12, "N": 30, "O": 2}
I1 = 13
NI = 30


def aba_inventario(ws, tab):
    widths(ws, I_W)
    title(ws, "Inventário de microfones",
          "Cadastre cada microfone com o formato, o tipo e o padrão do manual. As células cinza são calculadas.", "N")
    for r, rot in [(4, "Igreja ou equipe"), (5, "Responsável"), (6, "Data da conferência")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:G{r}", fmt=DATE if r == 6 else None)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    hint(ws, "I4", "Identifique cada microfone com uma fita ou etiqueta, e use o mesmo nome aqui e na aba Escolha por fonte.",
         merge="I4:N5", h="left")

    band(ws, 8, "Microfones", "N")
    heads = [("B", "Nº", INK), ("C", "Identificação", INK), ("D", "Formato", BLUE), ("E", "Tipo", BLUE),
             ("F", "Padrão polar", BLUE), ("G", "Ligação", BLUE), ("H", "Teste", AMBER), ("I", "Última limpeza", AMBER),
             ("J", "Phantom", TEAL), ("K", "Maior rejeição", TEAL), ("L", "Onde pôr o retorno", TEAL),
             ("M", "Situação", TEAL), ("N", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["", "nome ou fita", "lista", "lista", "do manual", "lista", "com a voz", "data",
                                        "precisa?", "direção", "", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Mão 3 (fita amarela)", "De mão", "Dinâmico", "Supercardioide", "Com fio",
                                       "OK", "04/10/2026", "Não", "cerca de 125°", "Em diagonal, nunca bem atrás", "OK"]):
        ex(ws, f"{col}11", v, h="left" if col in "CL" else "center")
    ex(ws, "N11", "")
    ws.row_dimensions[11].height = 21.75
    hint(ws, "B12", "Uma linha por microfone. O número da coluna B é o que você escolhe na aba Escolha por fonte.",
         merge="B12:N12", h="left")
    note(ws, "F9", "Procure no manual ou na ficha técnica do modelo. Dois microfones parecidos por fora podem ter padrões diferentes.")
    note(ws, "J9", "Sim para condensador com fio. Microfone sem fio: quem alimenta a cápsula é o transmissor.\nMicrofone de fita: confira o manual antes de ligar o phantom.")
    note(ws, "M9", "Defeito: separe para conserto. Testar: sem teste registrado. Incompleto: falta formato, tipo ou padrão.")

    I2 = inv_linhas(ws, I1, NI)
    cf_equal(ws, f"M{I1}:M{I2}", [("OK", GREEN), ("Testar", YELLOW), ("Incompleto", YELLOW), ("Defeito", RED)])
    cf_equal(ws, f"J{I1}:J{I2}", [("Sim", HUE["amber"][1])])

    s = I2 + 2
    band(ws, s, "Resumo automático", "N")
    linhas = [
        ("Microfones cadastrados", f'=COUNTA(D{I1}:D{I2})'),
        ("Dinâmicos", f'=COUNTIF(E{I1}:E{I2},"Dinâmico")'),
        ("Condensadores", f'=COUNTIF(E{I1}:E{I2},"Condensador")'),
        ("Precisam de phantom", f'=COUNTIF(J{I1}:J{I2},"Sim")'),
        ("Sem fio", f'=COUNTIF(G{I1}:G{I2},"Sem fio")'),
        ("Com defeito", f'=COUNTIF(M{I1}:M{I2},"Defeito")'),
        ("A testar ou incompletos", f'=COUNTIF(M{I1}:M{I2},"Testar")+COUNTIF(M{I1}:M{I2},"Incompleto")'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Cadastre os microfones",IF(F{s + 6}>0,"Há microfone com defeito: separe para conserto",'
                      f'IF(F{s + 7}>0,"Há microfone sem teste ou com cadastro incompleto","OK")))',
         b=False, merge=f"F{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:N{a}")
    ws.freeze_panes = f"D{I1}"
    setup(ws, tab, f"B1:N{a}")


# ================================================================== aba Escolha por fonte
E_W = {"A": 2, "B": 7, "C": 17, "D": 22, "E": 10, "F": 14, "G": 13, "H": 15, "I": 10, "J": 22, "K": 22, "L": 15,
       "M": 32, "N": 42, "O": 2}
E1 = 13
NE = 20


def aba_escolha(ws, tab, data=None):
    widths(ws, E_W)
    is_ex = data is not None
    title(ws, "Escolha de microfone por fonte",
          "Exemplo preenchido, para consulta. Use a aba Escolha por fonte para o seu evento." if is_ex
          else "Para cada canal, escolha o tipo de fonte e o número do microfone do inventário. As células cinza são calculadas.", "N")
    hb = WHITE if is_ex else INPUT
    cabecalho(ws, data["head"] if is_ex else {}, hb, "G")
    hint(ws, "I4", "A sugestão de cada fonte vem da aba Referência e é ponto de partida. “Tipo diferente” ou “Padrão diferente” "
                   "pedem conferência, e não troca obrigatória.", merge="I4:N6", h="left")

    band(ws, 9, "Fontes do evento", "N")
    heads = [("B", "Canal", INK), ("C", "Tipo de fonte", INK), ("D", "Descrição", INK), ("E", "Microfone nº", AMBER),
             ("F", "Formato", BLUE), ("G", "Tipo", BLUE), ("H", "Padrão", BLUE), ("I", "Phantom", BLUE),
             ("J", "Sugestão: tipo e padrão", TEAL), ("K", "Distância de partida", TEAL), ("L", "Situação", TEAL),
             ("M", "O que fazer", TEAL), ("N", "Posição de partida", TEAL)]
    for col, t, bg in heads:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "lista", "quem ou o quê", "do inventário", "do inventário", "do inventário",
                                        "do inventário", "no canal", "da Referência", "da Referência", "", "", "da Referência"]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 21.75
    if is_ex:
        hint(ws, "B12", "As fontes como foram planejadas, antes da conferência. O inventário do exemplo está mais abaixo.",
             merge="B12:N12", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Coral", "Coral, lado esquerdo", 3, "De mão", "Dinâmico", "Supercardioide",
                                           "Não", "Condensador, Cardioide", "60 a 90 cm à frente da 1ª fila",
                                           "Tipo diferente", "Sugestão: Condensador. Confira se este serve"]):
            ex(ws, f"{col}12", v, h="left" if col in "DJKM" else "center")
        ex(ws, "N12", "")
        note(ws, "C10", "Escolha a categoria da fonte. Ela puxa a sugestão de microfone da aba Referência.")
        note(ws, "E10", "Número do microfone na coluna B da aba Inventário.")
        note(ws, "L10", "Ordem de conferência: Não cadastrado, Defeito, Repetido (já usado em canal anterior), "
                        "Tipo diferente, Padrão diferente, Testar. OK quando nada disso acontece.")
    ws.row_dimensions[12].height = 24

    E2 = E1 + NE - 1
    if is_ex:
        inv1 = E2 + 19
        inv2 = inv1 + len(EX_INV) - 1
        inv = f"$B${inv1}:$J${inv2}"
    else:
        inv = f"Inventário!$B${I1}:$J${I1 + NI - 1}"
    bg = WHITE if is_ex else INPUT
    for k in range(NE):
        r = E1 + k
        d = data["fontes"][k] if is_ex and k < len(data["fontes"]) else None
        num(ws, f"B{r}", d[0] if d else k + 1)
        inp(ws, f"C{r}", d[1] if d else None, h="center", bg=bg)
        inp(ws, f"D{r}", d[2] if d else None, bg=bg)
        inp(ws, f"E{r}", d[3] if d else None, h="center", bg=bg)
        C, E = f"C{r}", f"E{r}"
        lk = lambda col: f'IFERROR(""&VLOOKUP({E},{inv},{col},FALSE),"")'
        rf = lambda col: f'IFERROR(VLOOKUP({C},{REF_RANGE},{col},FALSE),"")'
        calc(ws, f"F{r}", f'=IF({E}="","",{lk(3)})', b=False)
        calc(ws, f"G{r}", f'=IF({E}="","",{lk(4)})', b=False)
        calc(ws, f"H{r}", f'=IF({E}="","",{lk(5)})', b=False)
        calc(ws, f"I{r}", f'=IF({E}="","",{lk(9)})')
        calc(ws, f"J{r}", f'=IF({C}="","",{rf(2)}&", "&{rf(3)})', h="left", b=False, sz=9)
        calc(ws, f"K{r}", f'=IF({C}="","",{rf(5)})', h="left", b=False, sz=9)
        G, H = f"G{r}", f"H{r}"
        tipo_ac, pad_ac = rf(7), rf(8)
        pad_ok = (f'OR({pad_ac}="Qualquer",AND({pad_ac}="Direcional",OR({H}="Cardioide",{H}="Supercardioide",'
                  f'{H}="Hipercardioide")),{pad_ac}={H})')
        sit = (f'=IF({E}="",IF({C}="","","Sem microfone"),IF({lk(3)}&{lk(4)}&{lk(7)}="","Não cadastrado",'
               f'IF({lk(7)}="Defeito","Defeito",IF(COUNTIF($E${E1}:{E},{E})>1,"Repetido",IF({C}="","Sem fonte",'
               f'IF(AND({tipo_ac}<>"Qualquer",{G}<>{tipo_ac}),"Tipo diferente",IF(NOT({pad_ok}),"Padrão diferente",'
               f'IF({lk(7)}<>"OK","Testar","OK"))))))))')
        calc(ws, f"L{r}", sit)
        L = f"L{r}"
        acao = (f'=IF({L}="","",IF({L}="Sem microfone","Escolha um microfone do inventário",'
                f'IF({L}="Não cadastrado","Cadastre o microfone na aba Inventário",'
                f'IF({L}="Defeito","Troque de microfone e separe este para conserto",'
                f'IF({L}="Repetido","Já está em um canal anterior: escolha outro",'
                f'IF({L}="Sem fonte","Escolha o tipo de fonte",'
                f'IF({L}="Tipo diferente","Sugestão: "&{tipo_ac}&". Confira se este serve",'
                f'IF({L}="Padrão diferente","Confira o padrão e o lugar do retorno",'
                f'IF({L}="Testar","Teste com a voz antes do evento",'
                f'IF(I{r}="Sim","Ligue o phantom com o canal fechado","Manter"))))))))))')
        calc(ws, f"M{r}", acao, h="left", b=False, sz=9)
        calc(ws, f"N{r}", f'=IF({C}="","",{rf(6)})', h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 30
    dv_list(ws, f"C{E1}:C{E2}", FONTES, "Tipo de fonte. Puxa a sugestão da aba Referência.")
    dv_number(ws, f"E{E1}:E{E2}")
    cf_equal(ws, f"L{E1}:L{E2}", [("OK", GREEN), ("Defeito", RED), ("Não cadastrado", RED), ("Repetido", RED),
                                  ("Tipo diferente", YELLOW), ("Padrão diferente", YELLOW), ("Testar", YELLOW),
                                  ("Sem microfone", YELLOW), ("Sem fonte", YELLOW)])
    cf_equal(ws, f"I{E1}:I{E2}", [("Sim", HUE["amber"][1])])

    s = E2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(L{E1}:L{E2},"{t}")'
    linhas = [
        ("Fontes listadas", f'=COUNTA(D{E1}:D{E2})'),
        ("Com microfone OK", f'={cont("OK")}'),
        ("Canais com phantom", f'=COUNTIF(I{E1}:I{E2},"Sim")'),
        ("Defeito, repetido ou não cadastrado", f'={cont("Defeito")}+{cont("Repetido")}+{cont("Não cadastrado")}'),
        ("A conferir", f'={cont("Tipo diferente")}+{cont("Padrão diferente")}+{cont("Testar")}+{cont("Sem microfone")}+{cont("Sem fonte")}'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "D", "E")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(COUNTA(C{E1}:C{E2})+COUNT(E{E1}:E{E2})=0,"Liste as fontes do evento",'
                      f'IF(E{s + 4}>0,"Há microfone com defeito, repetido ou não cadastrado: troque antes do evento",'
                      f'IF(E{s + 5}>0,"Há escolhas a conferir","OK")))', b=False, merge=f"E{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:N{a}")
    end = a

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "N")
        r = end + 3
        put(ws, f"B{r}", "Canal", f=font(10, True), bg=GRAY, h="center")
        put(ws, f"C{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:N{r}")
        ws.row_dimensions[r].height = 21.75
        for canal, feito in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", canal, f=font(10, True, c=MUTED), bg=GRAY, h="center")
            put(ws, f"C{r}", feito, merge=f"C{r}:N{r}")
            ws.row_dimensions[r].height = 21.75
        assert r + 2 < inv1 - 1, "o inventário do exemplo ficou em cima do quadro de ajustes"
        band(ws, inv1 - 2, "Inventário do exemplo (os microfones que as fórmulas acima consultam)", "N")
        for col, t in zip("BCDEFGHIJ", ["Nº", "Identificação", "Formato", "Tipo", "Padrão", "Ligação", "Teste",
                                        "Última limpeza", "Phantom"]):
            head(ws, f"{col}{inv1 - 1}", t, bg=MUTED)
        ws.row_dimensions[inv1 - 1].height = 21.75
        inv_linhas(ws, inv1, len(EX_INV), EX_INV, completo=False)
        put(ws, f"K{inv1 - 1}", "Observação", f=font(10, True, c=WHITE), bg=MUTED, h="center", merge=f"K{inv1 - 1}:N{inv1 - 1}")
        for k, d in enumerate(EX_INV):
            put(ws, f"K{inv1 + k}", d[7], merge=f"K{inv1 + k}:N{inv1 + k}")
        end = inv2
    ws.freeze_panes = f"E{E1}"
    setup(ws, tab, f"B1:N{end}")


# ================================================================== aba Regra 3 para 1
R_W = {"A": 2, "B": 6, "C": 22, "D": 22, "E": 10, "F": 11, "G": 11, "H": 12, "I": 11, "J": 11, "K": 9, "L": 11,
       "M": 15, "N": 46, "O": 2}
R1 = 13
NR = 12


def aba_regra(ws, tab):
    widths(ws, R_W)
    title(ws, "Regra 3 para 1 — distância entre microfones abertos juntos",
          "Meça as distâncias em centímetros. As células cinza são calculadas.", "N")
    for r, rot in [(4, "Evento"), (5, "Data"), (6, "Medido por")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:G{r}", fmt=DATE if r == 5 else None)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    hint(ws, "I4", "A distância entre dois microfones abertos juntos deve ser pelo menos três vezes a distância de cada um até "
                   "a sua fonte. Assim, o som de uma fonte chega ao microfone vizinho cerca de 10 dB mais fraco.",
         merge="I4:N6", h="left")

    band(ws, 8, "Pares de microfones", "N")
    heads = [("B", "Par", INK), ("C", "Microfone A", INK), ("D", "Microfone B", INK), ("E", "Abertos juntos?", AMBER),
             ("F", "A até a fonte", AMBER), ("G", "B até a fonte", AMBER), ("H", "Entre os microfones", AMBER),
             ("I", "Maior até a fonte", TEAL), ("J", "Mínimo entre eles", TEAL), ("K", "Razão", TEAL),
             ("L", "Vazamento estimado", TEAL), ("M", "Situação", TEAL), ("N", "O que fazer", TEAL)]
    for col, t, bg in heads:
        head(ws, f"{col}9", t, bg=bg)
    ws.row_dimensions[9].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["nº", "quem usa", "quem usa", "Sim / Não", "cm", "cm", "cm", "cm", "cm = 3 × maior",
                                        "entre ÷ maior", "dB abaixo", "", ""]):
        hint(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 21.75
    for col, v in zip("BCDEFGHIJKLMN", ["Ex.", "Debatedor 1", "Debatedor 2", "Sim", 50, 50, 60, 50, 150, 1.2, 3.9,
                                        "Fora da regra", "Aproxime cada microfone a até 20 cm da fonte, ou afaste os microfones para 150 cm"]):
        ex(ws, f"{col}11", v, h="left" if col in "CDN" else "center")
    ws.row_dimensions[11].height = 30
    hint(ws, "B12", "A linha Ex. é o exemplo 2 do treinamento, antes da correção: a mesa de debate com som oco.",
         merge="B12:N12", h="left")
    note(ws, "L9", "Quanto mais fraco o som de uma fonte chega ao microfone vizinho do que ao dela.\n"
                   "Estimativa em campo livre, com as fontes à frente dos microfones e os microfones lado a lado.")
    note(ws, "E9", "Se nunca ficam abertos ao mesmo tempo (um fala, depois o outro), os dois não somam e a regra não se aplica.")

    R2 = R1 + NR - 1
    for k in range(NR):
        r = R1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}")
        inp(ws, f"D{r}")
        inp(ws, f"E{r}", h="center")
        for col in "FGH":
            inp(ws, f"{col}{r}", h="center")
        F, G, H, I, J, K = (f"{c}{r}" for c in "FGHIJK")
        vazio = f'OR({F}="",{G}="",{H}="")'
        calc(ws, I, f'=IF({vazio},"",MAX({F},{G}))', b=False)
        calc(ws, J, f'=IF({I}="","",3*{I})', b=False)
        calc(ws, K, f'=IF(OR({I}="",{I}=0),"",ROUND({H}/{I},1))', fmt="0.0")
        calc(ws, f"L{r}", f'=IF(OR({I}="",{I}=0),"",ROUND(20*LOG10(SQRT({H}^2+{I}^2)/{I}),1))', b=False, fmt="0.0")
        calc(ws, f"M{r}", f'=IF({K}="","",IF(E{r}="Não","Não somam",IF({K}>=3,"OK",IF({K}>=2,"Perto do limite","Fora da regra"))))')
        calc(ws, f"N{r}", f'=IF(M{r}="","",IF(M{r}="Não somam","Mantenha só um aberto por vez",IF(M{r}="OK","Manter",'
                          f'"Aproxime cada microfone a até "&ROUND({H}/3,0)&" cm da fonte, ou afaste os microfones para "&{J}&" cm")))',
             h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 30
    dv_list(ws, f"E{R1}:E{R2}", SIM_NAO, "Os dois ficam abertos ao mesmo tempo?")
    dv_number(ws, f"F{R1}:H{R2}")
    cf_equal(ws, f"M{R1}:M{R2}", [("OK", GREEN), ("Não somam", GREEN), ("Perto do limite", YELLOW), ("Fora da regra", RED)])

    s = R2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(M{R1}:M{R2},"{t}")'
    linhas = [("Pares avaliados", f'={cont("OK")}+{cont("Não somam")}+{cont("Perto do limite")}+{cont("Fora da regra")}'),
              ("Dentro da regra", f'={cont("OK")}+{cont("Não somam")}'),
              ("Perto do limite", f'={cont("Perto do limite")}'),
              ("Fora da regra", f'={cont("Fora da regra")}')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "D", "E")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(E{s + 1}=0,"Meça os pares de microfones abertos juntos",IF(E{s + 4}>0,"Há par fora da regra: aproxime cada microfone da sua fonte",'
                      f'IF(E{s + 3}>0,"Há par perto do limite: confira o timbre com os dois abertos","OK")))',
         b=False, merge=f"E{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:N{a}")
    hint(ws, f"B{a + 1}", "O corte de 2 para “Perto do limite” é regra prática deste material. Dois microfones na mesma boca "
                          "(dois goosenecks no púlpito) nunca cumprem a regra: use um só.", merge=f"B{a + 1}:N{a + 1}", h="left")
    ws.freeze_panes = f"C{R1}"
    setup(ws, tab, f"B1:N{a + 1}")


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Microfones — inventário, escolha por fonte e regra 3:1"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Microfones — como usar esta planilha",
           "Inventário dos microfones, escolha do microfone de cada fonte e conferência da regra 3:1.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Formato, tipo, padrão, ligação, teste e tipo de fonte aceitam só as opções da lista."),
        ("Números", "Distâncias em centímetros, só o número: 30, e não “30 cm”."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Inventário: cadastre cada microfone, com o padrão tirado do manual do modelo, e o resultado do teste."),
        ("2", "Aba Escolha por fonte: liste as fontes do evento, escolha o tipo de fonte e o número do microfone."),
        ("3", "Veja a situação e o que fazer. Troque o microfone ou confira a escolha até tudo ficar OK."),
        ("4", "Na montagem: posicione cada microfone pela distância e pela posição de partida, e ligue o phantom onde indicado."),
        ("5", "Aba Regra 3 para 1: meça os microfones que ficam abertos juntos e corrija os pares fora da regra."),
        ("6", "Aba Checklist: confira os doze itens antes do ajuste de ganho."),
    ]),
    ("Abas da planilha", [
        ("Inventário", "Até 30 microfones, com phantom, maior rejeição, lugar do retorno e situação calculados."),
        ("Escolha por fonte", "Até 20 fontes, com a sugestão de microfone e a conferência do microfone escolhido."),
        ("Regra 3 para 1", "Até 12 pares de microfones, com razão, vazamento estimado e o que fazer."),
        ("Referência", "Ponto de partida por fonte, padrões polares, dinâmico e condensador, distância e nível."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Escolha", "O culto do exemplo 1 do treinamento: nove fontes como foram planejadas, com o inventário de dez microfones."),
    ]),
    ("Regras e premissas do modelo", [
        ("Phantom", "Sim para condensador com fio. Não para dinâmico e para microfone sem fio. Fita: confira o manual."),
        ("Sugestão por fonte", "Vem da tabela da aba Referência. Ela é ponto de partida, e não regra: a decisão final é pelo ouvido."),
        ("Tipo e padrão", "“Tipo diferente” aparece quando a fonte pede um tipo e o microfone é de outro. “Padrão diferente”, quando a fonte pede padrão direcional e o microfone é omni ou figura de 8."),
        ("Repetido", "O mesmo número de microfone em dois canais. A marcação aparece no segundo canal em que ele foi usado."),
        ("Regra 3 para 1", "Razão = distância entre os microfones ÷ maior distância até a fonte. OK a partir de 3; Perto do limite de 2 a 3."),
        ("Vazamento estimado", "Conta de campo livre (cerca de 6 dB a cada vez que a distância dobra). Em sala reverberante, o vazamento real é maior."),
        ("Exemplos", "Os dados das abas e da linha Ex. são ilustrativos, criados para este material."),
    ]),
])

aba_inventario(wb.create_sheet("Inventário"), BLUE)
aba_escolha(wb.create_sheet("Escolha por fonte"), AMBER)
aba_regra(wb.create_sheet("Regra 3 para 1"), TEAL)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 18, "C": 17, "D": 22, "E": 24, "F": 24, "G": 50, "H": 13, "I": 13, "J": 2})
title(ws, "Referência — microfones", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "I")
band(ws, 4, "Ponto de partida por fonte (consultada pela aba Escolha por fonte: não mude a ordem das colunas)", "I")
for col, t in zip("BCDEFGHI", ["Fonte", "Tipo", "Padrão", "Formato", "Distância", "Posição e direção", "Tipo aceito", "Padrão aceito"]):
    put(ws, f"{col}5", t, f=font(10, True), bg=GRAY, h="center")
ws.row_dimensions[5].height = 21.75
for k, lin in enumerate(REF_FONTES):
    r = REF_R1 + k
    put(ws, f"B{r}", lin[0], f=font(10, True))
    for col, t in zip("CDEFGHI", lin[1:]):
        put(ws, f"{col}{r}", t, h="left" if col in "EFG" else "center")
    ws.row_dimensions[r].height = 31.5
r = REF_R1 + len(REF_FONTES)
hint(ws, f"B{r}", "Tipo aceito e Padrão aceito definem quando a aba Escolha por fonte marca “Tipo diferente” ou “Padrão diferente”. "
                  "Direcional = cardioide, supercardioide ou hipercardioide.", merge=f"B{r}:I{r}", h="left")
ws.row_dimensions[r].height = 24


def tabela(r, titulo, cab, linhas, spans, alt=31.5):
    """spans: um intervalo de colunas por coluna da tabela, como "B", "C" ou "G:I"."""
    band(ws, r, titulo, "I")
    r += 1
    for sp, t in zip(spans, cab):
        a = sp.split(":")[0]
        put(ws, f"{a}{r}", t, f=font(10, True), bg=GRAY, h="center",
            merge=f"{a}{r}:{sp.split(':')[1]}{r}" if ":" in sp else None)
    ws.row_dimensions[r].height = 21.75
    for lin in linhas:
        r += 1
        for k, (sp, t) in enumerate(zip(spans, lin)):
            a = sp.split(":")[0]
            ultimo = k == len(spans) - 1
            put(ws, f"{a}{r}", t, f=font(10, True) if k == 0 else None, h="left" if ultimo or k == 0 else "center",
                merge=f"{a}{r}:{sp.split(':')[1]}{r}" if ":" in sp else None)
        ws.row_dimensions[r].height = alt
    return r + 2


r = tabela(r + 2, "Padrões polares (valores do padrão ideal, em relação à frente)",
           ["Padrão", "Lados (90°)", "Atrás (180°)", "Maior rejeição", "Onde pôr o retorno", "Uso típico"], spans=["B", "C", "D", "E", "F", "G:I"], linhas=[
    ("Cardioide", "−6 dB", "quase nada", "180°, bem atrás", "Bem atrás do microfone", "O mais comum: voz de mão, púlpito, coral, instrumentos."),
    ("Supercardioide", "−9 dB", "−12 dB", "cerca de 125°", "Em diagonal, nunca bem atrás", "Voz em palco com volume alto, gooseneck, caixa e tons."),
    ("Hipercardioide", "−12 dB", "−6 dB", "cerca de 110°", "Em diagonal, mais para o lado", "Quando o que atrapalha está dos lados."),
    ("Omni", "0 dB", "0 dB", "não tem", "Longe: não ajuda a rejeitar", "Lapela e headset, sempre muito perto da boca. Sem efeito de proximidade."),
    ("Figura de 8", "quase nada", "0 dB", "90°, nos dois lados", "Em um dos lados", "Raro ao vivo. Padrão natural do microfone de fita."),
])
r = tabela(r, "Dinâmico e condensador", ["", "Dinâmico", "Condensador", "Observação"], spans=["B", "C:D", "E:F", "G:I"], linhas=[
    ("Energia", "Não precisa", "Phantom, pilha ou transmissor", "", "Phantom de +48 V: ligue e desligue com o canal fechado."),
    ("Nível de saída", "1 a 3 mV/Pa", "5 a 30 mV/Pa", "", "O condensador entrega uns 10 a 20 dB a mais na mesma fonte e distância."),
    ("Robustez", "Alta", "Menor", "", "Condensador: cuidado com queda, umidade, sopro e manuseio."),
    ("Uso típico", "Voz de mão, bateria, amplificador", "Coral, púlpito, lapela, violão, pratos", "", "Faixas típicas, que variam com o modelo."),
])
r = tabela(r, "Distância até a fonte e nível no microfone (campo livre, em relação a 5 cm)",
           ["Distância", "Nível", "Formato típico", "Observação"], spans=["B", "C", "D:E", "F:I"], linhas=[
    ("2,5 cm", "+6 dB", "Headset, canto de perto", "", "Bem perto da boca, a variação real é um pouco menor."),
    ("5 cm", "0 dB", "De mão (canto e fala)", "", "Referência."),
    ("10 cm", "−6 dB", "De mão (fala)", "", "Cada vez que a distância dobra, cerca de −6 dB."),
    ("20 cm", "−12 dB", "Lapela, gooseneck de perto", "", "O ganho na mesa sobe na mesma medida, e com ele a sala e o retorno."),
    ("40 cm", "−18 dB", "Gooseneck de longe", "", "Pouca margem antes da microfonia."),
    ("80 cm", "−24 dB", "Coral suspenso", "", "A fonte com menos margem do palco."),
], alt=21.75)
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com o modelo, a fonte e a sala. Confira o manual e a ficha técnica do seu microfone.",
     merge=f"B{r - 1}:I{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:I{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist dos microfones",
          "Confira na montagem, antes do ajuste de ganho, e de novo ao guardar o equipamento.",
          CHECK, tab=HUE["purple"][0])
aba_escolha(wb.create_sheet("Exemplo - Escolha"), MUTED, EX_ESCOLHA)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
