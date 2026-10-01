# -*- coding: utf-8 -*-
"""Gera Exposicao-modelo.xlsx: calculadora de stops, ficha de ajuste por câmera, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIM_NAO = ["Sim", "Não"]
STOP = '+0.0;-0.0;0.0'
L2 = "LOG10(2)"

CHECK = [
    ("Sistema", "Todas as câmeras estão na mesma taxa de quadros, a do sistema."),
    ("Sistema", "A exposição está em manual: íris, obturador e ISO ou ganho, sem nenhum automático ligado."),
    ("Obturador", "O obturador está no dobro da taxa de quadros (1/60 em 30 fps; 1/120 ou 1/60 em 60 fps)."),
    ("Obturador", "Não há cintilação nem faixas na imagem, com a luz do evento, o telão e o projetor ligados."),
    ("Exposição", "A íris foi ajustada com a luz do evento e uma pessoa no palco, e não com a luz de serviço."),
    ("Exposição", "O ISO ou ganho é o menor que dá a exposição certa, e o ruído foi conferido em um monitor grande."),
    ("Exposição", "Ao ar livre, o filtro ND está ligado e a íris não passa de f/11. Em local fechado, o ND está desligado."),
    ("Medição", "O rosto da pessoa principal está no nível certo na forma de onda ou na zebra."),
    ("Medição", "Nada importante estoura: não há zebra de 100% no rosto, na roupa clara nem no papel."),
    ("Medição", "A exposição foi conferida com o zoom nas duas pontas."),
    ("Câmeras", "O rosto tem o mesmo nível em todas as câmeras, comparadas lado a lado no multiview."),
    ("Câmeras", "Os ajustes de cada câmera estão anotados, e a equipe sabe que a correção durante o evento é na íris."),
]

# Exemplo 1 do treinamento: imagem granulada em palco escuro
EX_CALC = {
    "ref": (30, 60),
    "A": (4, 250, 12800, None, 0),
    "B": (2.8, 60, 1600, None, 0),
}

# Exemplo 2 do treinamento: quatro câmeras em uma formatura, como foram encontradas
EX_FICHA = {
    "head": {"evento": "Formatura no auditório da escola", "local": "Auditório, luz de palco e telão de LED",
             "data": date(2026, 12, 12), "operador": "Diretor de imagem", "pessoa": "Orador no púlpito"},
    "ref": (30, 60, 45, 70, 5, 6400, 12),
    "cams": [  # câmera, modo, obturador, f, ISO, ganho, ND, rosto, zebra 100%
        ("1. Central, plano geral", "Manual", 60, 4, 1600, None, 0, 62, "Não"),
        ("2. Lateral, close do orador", "Manual", 100, 2.8, 1600, None, 0, 60, "Não"),
        ("3. PTZ no fundo", "Automático", 60, 2.8, 6400, None, 0, 96, "Sim"),
        ("4. Câmera móvel, zoom fechado", "Manual", 60, 5.6, 1600, None, 0, 48, "Não"),
    ],
    "ajustes": [  # câmera, o que foi feito, leitura depois
        ("1", "Nada. As outras câmeras foram igualadas a ela.", "rosto em 62%"),
        ("2", "Obturador de 1/100 para 1/60. Íris fechada de f/2,8 para f/3,5, para compensar.", "rosto em 61%, sem faixas"),
        ("3", "Exposição em manual. ISO de 6400 para 1600, com a íris em f/2,8.", "rosto em 63%"),
        ("4", "A lente fecha para f/5,6 no fim do zoom: ISO de 1600 para 3200.", "rosto em 61%"),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def cf_good(ws, rng, first, goods, bads=()):
    """Verde para os textos de goods, vermelho para bads, amarelo para outro texto, nada para vazio."""
    cf_equal(ws, rng, [(g, GREEN) for g in goods] + [(b, RED) for b in bads])
    cond = ",".join(f'{first}<>"{g}"' for g in list(goods) + list(bads))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(LEN({first})>0,{cond})'],
                                                   fill=PatternFill("solid", bgColor=YELLOW, fgColor=YELLOW)))


def cintila(den, hz):
    """OK quando o tempo do obturador cobre um número inteiro de pulsos da luz (2 × a rede)."""
    return f'AND({den}<=2*{hz},ABS(2*{hz}/{den}-ROUND(2*{hz}/{den},0))<0.02)'


# ================================================================== aba Calculadora de stops
K_W = {"A": 2, "B": 34, "C": 15, "D": 15, "E": 15, "F": 17, "G": 54, "H": 2}


def aba_calc(ws, tab, data=None):
    widths(ws, K_W)
    is_ex = data is not None
    title(ws, "Calculadora de stops — equivalência entre f, obturador e ISO",
          "Exemplo preenchido (exemplo 1 do treinamento). Use a aba Calculadora de stops para as suas contas." if is_ex
          else "Digite o ajuste atual (A) e o novo ajuste (B). A planilha mostra a diferença de luz e o que compensa.", "G")
    bg = WHITE if is_ex else INPUT
    ref = data["ref"] if is_ex else (30, 60)
    for k, (rot, v, dica) in enumerate([
            ("Taxa de quadros (fps)", ref[0], "A do sistema inteiro: 30 ou 60 é o mais comum. Use 30 para 29,97 e 60 para 59,94."),
            ("Rede elétrica (Hz)", ref[1], "60 no Brasil. Serve para conferir a cintilação da luz ligada na rede.")]):
        r = 4 + k
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", v, h="center", bg=bg)
        hint(ws, f"E{r}", dica, merge=f"E{r}:G{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, "D4:D5")
    FPS, HZ = "$D$4", "$D$5"

    band(ws, 7, "Comparação de dois ajustes", "G")
    for col, t, c in [("B", "Controle", INK), ("C", "Ajuste atual (A)", AMBER), ("D", "Novo ajuste (B)", AMBER),
                      ("E", "Diferença", TEAL), ("F", "Valor de B que compensa", TEAL), ("G", "O que muda de A para B", TEAL)]:
        head(ws, f"{col}8", t, bg=c)
    ws.row_dimensions[8].height = 30
    for col, t in zip("BCDEFG", ["", "", "", "stops; + = B mais clara", "mantendo os outros de B", ""]):
        hint(ws, f"{col}9", t)
    ws.row_dimensions[9].height = 24

    linhas = [
        ("Abertura (número f)", "f: digite 2.8 para f/2,8",
         lambda c, d: f'2*LOG10({c}/{d})/{L2}', lambda d: f'ROUND({d}*2^($E$15/2),1)',
         lambda c, d: f'IF({d}<{c},"Íris mais aberta: mais luz e fundo mais desfocado",IF({d}>{c},"Íris mais fechada: menos luz e mais profundidade de campo","Sem mudança"))'),
        ("Obturador (1/x de segundo)", "digite só o x: 60 para 1/60",
         lambda c, d: f'LOG10({c}/{d})/{L2}', lambda d: f'ROUND({d}*2^$E$15,0)',
         lambda c, d: f'IF({d}<{c},"Obturador mais lento: mais luz e mais borrão",IF({d}>{c},"Obturador mais rápido: menos luz e movimento mais picotado","Sem mudança"))'),
        ("ISO", "use esta linha ou a do ganho",
         lambda c, d: f'LOG10({d}/{c})/{L2}', lambda d: f'ROUND({d}/2^$E$15,0)',
         lambda c, d: f'IF({d}>{c},"ISO maior: mais claro e mais ruído",IF({d}<{c},"ISO menor: mais escuro e menos ruído","Sem mudança"))'),
        ("Ganho (dB)", "filmadoras e PTZ: 6 dB = 1 stop",
         lambda c, d: f'({d}-{c})/6', lambda d: f'ROUND({d}-6*$E$15,0)',
         lambda c, d: f'IF({d}>{c},"Ganho maior: mais claro e mais ruído",IF({d}<{c},"Ganho menor: mais escuro e menos ruído","Sem mudança"))'),
        ("Filtro ND (stops que corta)", "0 sem filtro; ND8 = 3 stops",
         lambda c, d: f'{c}-{d}', lambda d: f'ROUND({d}+$E$15,1)',
         lambda c, d: f'IF({d}>{c},"Mais ND: menos luz, sem outro efeito na imagem",IF({d}<{c},"Menos ND: mais luz, sem outro efeito na imagem","Sem mudança"))'),
    ]
    vals_a = data["A"] if is_ex else (None,) * 5
    vals_b = data["B"] if is_ex else (None,) * 5
    for k, (rot, dica, dif, comp, efeito) in enumerate(linhas):
        r = 10 + k
        put(ws, f"B{r}", rot, f=font(10, True), bg=GRAY)
        note_txt = dica
        inp(ws, f"C{r}", vals_a[k], h="center", bg=bg)
        inp(ws, f"D{r}", vals_b[k], h="center", bg=bg)
        C, D = f"C{r}", f"D{r}"
        vazio = f'OR({C}="",{D}="")' if k >= 3 else f'OR({C}="",{D}="",{C}<=0,{D}<=0)'
        calc(ws, f"E{r}", f'=IF({vazio},"",IFERROR({dif(C, D)},""))', fmt=STOP)
        calc(ws, f"F{r}", f'=IF(OR(E{r}="",$E$15=""),"",IFERROR({comp(D)},""))', b=False)
        calc(ws, f"G{r}", f'=IF(E{r}="","",{efeito(C, D)})', h="left", b=False, sz=9)
        if not is_ex:
            note(ws, f"B{r}", note_txt)
        ws.row_dimensions[r].height = 24
    dv_number(ws, "C10:D14")
    summary(ws, 15, "Diferença total (stops)", '=IF(COUNT(E10:E14)=0,"",SUM(E10:E14))', "D", "E", fmt=STOP)
    label(ws, "B16", "Leitura", merge="B16:D16", h="right")
    calc(ws, "E16", '=IF(E15="","",IF(ABS(E15)<1/3,"OK",IF(E15>0,"B mais clara que A","B mais escura que A")))', merge="E16:F16")
    hint(ws, "G16", "OK quando a diferença total fica abaixo de 1/3 de stop, que não se vê na imagem.", h="left")
    ws.row_dimensions[16].height = 21.75
    cf_status(ws, "E16:F16", "$E$16")

    band(ws, 18, "O obturador de cada ajuste", "G")
    head(ws, "B19", "Conferência")
    head(ws, "C19", "A", bg=AMBER)
    head(ws, "D19", "B", bg=AMBER)
    head(ws, "E19", "Como é lida", bg=TEAL)
    ws.merge_cells("E19:G19")
    rot19 = [("Ângulo do obturador (°)", "360 × taxa de quadros ÷ x. A regra dos 180° é o ponto de partida."),
             ("Movimento", "Regra dos 180°: de 160° a 220°. Acima, mais borrão; abaixo de 90°, movimento picotado."),
             ("Cintilação com luz da rede", "OK quando o tempo cobre um número inteiro de pulsos (2 × a rede). LED com dimmer pode piscar mesmo assim: teste.")]
    for k, (rot, dica) in enumerate(rot19):
        r = 20 + k
        put(ws, f"B{r}", rot, f=font(10, True), bg=GRAY)
        for col in "CD":
            x = f"{col}11"
            if k == 0:
                f_ = f'=IF(OR({x}="",{FPS}=""),"",IFERROR(ROUND(360*{FPS}/{x},0),""))'
            elif k == 1:
                a = f"{col}20"
                f_ = (f'=IF({a}="","",IF({a}>360,"Mais lento que um quadro: evite",IF({a}>220,"Borrão maior que o natural",'
                      f'IF({a}>=160,"Regra dos 180°: OK",IF({a}>=90,"Pouco borrão","Movimento picotado")))))')
            else:
                f_ = f'=IF(OR({x}="",{HZ}=""),"",IFERROR(IF({cintila(x, HZ)},"OK","Risco de cintilação"),""))'
            calc(ws, f"{col}{r}", f_, b=(k == 0), sz=9 if k else 10)
        hint(ws, f"E{r}", dica, merge=f"E{r}:G{r}", h="left")
        ws.row_dimensions[r].height = 30
    cf_good(ws, "C21:D22", "C21", ["OK", "Regra dos 180°: OK"], ["Risco de cintilação", "Mais lento que um quadro: evite"])

    band(ws, 24, "Conversores rápidos", "G")
    for col, t in zip("BCDE", ["Converter", "Digite", "Resultado", "Unidade do resultado"]):
        head(ws, f"{col}25", t)
    ws.merge_cells("E25:G25")
    conv = [
        ("Ganho em dB → stops", 12, '=IF(C26="","",C26/6)', "stops (6 dB = 1 stop)", STOP),
        ("Stops → ganho em dB", 2, '=IF(C27="","",C27*6)', "dB", None),
        ("Filtro ND pelo fator (ND8 → 8) → stops", 8, f'=IF(OR(C28="",C28<=0),"",ROUND(LOG10(C28)/{L2},1))', "stops que o filtro corta", None),
        ("Filtro ND pela densidade (0,9) → stops", 0.9, f'=IF(C29="","",ROUND(C29/{L2},1))', "stops que o filtro corta", None),
        ("Ângulo do obturador (°) → 1/x", 180, f'=IF(OR(C30="",C30<=0,{FPS}=""),"",ROUND(360*{FPS}/C30,0))', "x de 1/x, na taxa de quadros do alto", None),
        ("Obturador 1/x → ângulo (°)", 60, f'=IF(OR(C31="",C31<=0,{FPS}=""),"",ROUND(360*{FPS}/C31,0))', "graus, na taxa de quadros do alto", None),
    ]
    for k, (rot, v, f_, uni, fmt) in enumerate(conv):
        r = 26 + k
        put(ws, f"B{r}", rot, f=font(10, True), bg=GRAY)
        inp(ws, f"C{r}", v if is_ex else None, h="center", bg=bg)
        calc(ws, f"D{r}", f_, fmt=fmt)
        hint(ws, f"E{r}", uni, merge=f"E{r}:G{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, "C26:C31")
    end = 31
    if is_ex:
        band(ws, 33, "O que observar", "G")
        put(ws, "B34", "A íris (+1,0) e o obturador (+2,1) devolveram 3 stops de luz de verdade, e o ISO pôde descer os mesmos 3 stops, "
                       "de 12800 para 1600. A sobra de 0,1 stop não se vê. O ajuste A tinha o obturador fora da regra e com risco de cintilação.",
            merge="B34:G34")
        ws.row_dimensions[34].height = 33
        end = 34
    else:
        hint(ws, "B33", "Valores de referência. A conta de stops é exata; o ISO de base, o ruído aceitável e a abertura real da lente mudam com a câmera.",
             merge="B33:G33", h="left")
        end = 33
    ws.freeze_panes = "C10"
    setup(ws, tab, f"B1:G{end}", landscape=False, fit_height=True)


# ================================================================== aba Ficha por câmera
F_W = {"A": 2, "B": 5, "C": 28, "D": 12, "E": 10, "F": 9, "G": 9, "H": 9, "I": 8, "J": 9, "K": 11, "L": 9,
       "M": 11, "N": 11, "O": 13, "P": 46, "Q": 26, "R": 2}
N_CAM = 8
R1 = 16
SITS = ["OK", "Cintilação", "Obturador", "Estouro", "Automático", "Ganho alto", "Sem leitura", "Rosto alto", "Rosto baixo", "Igualar"]


def aba_ficha(ws, tab, data=None):
    widths(ws, F_W)
    is_ex = data is not None
    title(ws, "Ficha de ajuste por câmera",
          "Exemplo preenchido: as câmeras do exemplo 2 do treinamento, como foram encontradas." if is_ex
          else "Com a luz do evento e alguém no palco, anote o ajuste e o nível do rosto de cada câmera. As células cinza são calculadas.", "Q")
    bg = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local e luz", "local", None), (6, "Data", "data", DATE),
                           (7, "Responsável", "operador", None), (8, "Pessoa de referência", "pessoa", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=bg, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    ref = data["ref"] if is_ex else (30, 60, 45, 70, 5, 6400, 12)
    refs = [("Taxa de quadros (fps)", "A do sistema. Use 30 para 29,97 e 60 para 59,94."),
            ("Rede elétrica (Hz)", "60 no Brasil."),
            ("Rosto: nível mínimo (%)", "Parte mais iluminada da pele. Pele negra: perto de 30. Morena: 45."),
            ("Rosto: nível máximo (%)", "Pele clara: perto de 70. Ajuste a faixa à pessoa de referência."),
            ("Tolerância entre câmeras", "Diferença de nível do rosto, em pontos de %, que ainda conta como igual."),
            ("ISO máximo aceitável", "O que a sua câmera aguenta sem ruído incômodo. Teste."),
            ("Ganho máximo aceitável (dB)", "Para câmeras que mostram ganho em dB.")]
    for k, ((rot, dica), v) in enumerate(zip(refs, ref)):
        r = 4 + k
        label(ws, f"I{r}", rot, merge=f"I{r}:K{r}")
        inp(ws, f"L{r}", v, h="center", bg=bg)
        hint(ws, f"M{r}", dica, merge=f"M{r}:Q{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, "L4:L10")
    FPS, HZ, MIN_, MAX_, TOL, ISOMAX, GMAX = (f"$L${4 + k}" for k in range(7))

    band(ws, 12, "Ajuste e medição de cada câmera", "Q")
    heads = [("B", "nº", INK), ("C", "Câmera e posição", INK), ("D", "Modo", AMBER), ("E", "Obturador", AMBER),
             ("F", "Abertura", AMBER), ("G", "ISO", AMBER), ("H", "Ganho", AMBER), ("I", "ND", AMBER),
             ("J", "Rosto", AMBER), ("K", "Zebra 100%", AMBER), ("L", "Ângulo", TEAL), ("M", "Cintilação", TEAL),
             ("N", "Rosto − câm. 1", TEAL), ("O", "Situação", TEAL), ("P", "O que fazer", TEAL), ("Q", "Observação", INK)]
    for col, t, c in heads:
        head(ws, f"{col}13", t, bg=c)
    ws.row_dimensions[13].height = 30
    for col, t in zip("BCDEFGHIJKLMNOPQ", ["", "a câmera 1 é a referência", "Manual / Auto", "x de 1/x", "número f", "ou ganho",
                                           "dB", "stops", "% na forma de onda", "no que importa?", "graus", "luz da rede",
                                           "pontos de %", "", "", ""]):
        hint(ws, f"{col}14", t)
    ws.row_dimensions[14].height = 27
    if is_ex:
        hint(ws, "B15", "Ajustes encontrados antes da correção. Rosto medido na forma de onda, com o orador no púlpito e a luz do evento. As colunas de L a P são calculadas.",
             merge="B15:Q15", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNOP", ["Ex.", "Câmera lateral, close", "Manual", 60, 2.8, 1600, "", 0, 55, "Não",
                                               180, "OK", -7, "Igualar", "Mais escuro que a câmera 1: abra a íris. Se acabou, suba o ISO"]):
            ex(ws, f"{col}15", v, h="left" if col in "CP" else "center")
        ex(ws, "Q15", "")
        note(ws, "E13", "Digite só o denominador: 60 para 1/60.")
        note(ws, "G13", "Preencha o ISO ou o ganho, conforme o que a câmera mostra.")
        note(ws, "J13", "Nível da parte mais iluminada do rosto, em %, lido na forma de onda.\nSem forma de onda, anote o nível da zebra em que as listras começam a aparecer.")
        note(ws, "K13", "A zebra de 100% aparece no rosto, na roupa clara ou no papel? Telão e refletor no quadro não contam.")
        note(ws, "N13", "Diferença entre o rosto desta câmera e o da câmera 1, a referência.")
    ws.row_dimensions[15].height = 24

    R2 = R1 + N_CAM - 1
    for k in range(N_CAM):
        r = R1 + k
        c = data["cams"][k] if is_ex and k < len(data["cams"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", c[0] if c else None, bg=bg)
        for col, j in zip("DEFGHIJK", range(1, 9)):
            inp(ws, f"{col}{r}", c[j] if c else None, h="center", bg=bg)
        D, E, G, Hc, J, K, L, M, N, O = (f"{x}{r}" for x in "DEGHJKLMNO")
        calc(ws, L, f'=IF(OR({E}="",{FPS}=""),"",IFERROR(ROUND(360*{FPS}/{E},0),""))', b=False)
        calc(ws, M, f'=IF(OR({E}="",{HZ}=""),"",IFERROR(IF({cintila(E, HZ)},"OK","Risco"),""))', b=False)
        calc(ws, N, f'=IF(OR({J}="",$J${R1}=""),"",{J}-$J${R1})', fmt='+0;-0;0', b=False)
        calc(ws, O, f'=IF(COUNTA(D{r}:K{r})=0,"",IF({M}="Risco","Cintilação",'
                    f'IF(AND({L}<>"",OR({L}>360,{L}<90)),"Obturador",'
                    f'IF(OR({K}="Sim",AND({J}<>"",{J}>=100)),"Estouro",'
                    f'IF({D}="Automático","Automático",'
                    f'IF(OR(AND({G}<>"",{G}>{ISOMAX}),AND({Hc}<>"",{Hc}>{GMAX})),"Ganho alto",'
                    f'IF({J}="","Sem leitura",IF({J}>{MAX_},"Rosto alto",IF({J}<{MIN_},"Rosto baixo",'
                    f'IF(AND({N}<>"",ABS({N})>{TOL}),"Igualar","OK"))))))))))')
        ok_txt = "Referência: as outras câmeras se igualam a ela" if k == 0 else "Manter"
        calc(ws, f"P{r}", f'=IF({O}="","",IF({O}="Cintilação",IF({HZ}=50,"Use 1/50 ou 1/100","Use 1/60 ou 1/120")&" e compense na íris",'
                          f'IF({O}="Obturador","Volte ao dobro da taxa de quadros: 1/"&ROUND(2*{FPS},0)&". Compense na íris",'
                          f'IF({O}="Estouro","Feche a íris até a zebra de 100% sair do rosto e da roupa clara",'
                          f'IF({O}="Automático","Passe para manual e exponha pelo rosto",'
                          f'IF({O}="Ganho alto","Abra a íris ou peça mais luz no rosto, e baixe o ISO ou ganho",'
                          f'IF({O}="Sem leitura","Meça o rosto na forma de onda ou na zebra",'
                          f'IF({O}="Rosto alto","Feche a íris um pouco",'
                          f'IF({O}="Rosto baixo","Abra a íris. Se acabou, suba o ISO ou ganho",'
                          f'IF({O}="Igualar",IF({N}>0,"Mais claro que a câmera 1: feche a íris um pouco",'
                          f'"Mais escuro que a câmera 1: abra a íris. Se acabou, suba o ISO"),"{ok_txt}"))))))))))',
             h="left", b=False, sz=9)
        inp(ws, f"Q{r}", bg=bg)
        ws.row_dimensions[r].height = 24
    dv_list(ws, f"D{R1}:D{R2}", ["Manual", "Automático"], "Manual ou Automático (qualquer um dos três controles no automático).")
    dv_list(ws, f"K{R1}:K{R2}", SIM_NAO, "A zebra de 100% aparece no rosto, na roupa clara ou no papel?")
    dv_number(ws, f"E{R1}:J{R2}")
    cf_equal(ws, f"M{R1}:M{R2}", [("OK", GREEN), ("Risco", RED)])
    cf_good(ws, f"O{R1}:O{R2}", f"O{R1}", ["OK"], ["Cintilação", "Estouro"])

    s = R2 + 2
    band(ws, s, "Resumo automático", "Q")
    cont = lambda t: f'COUNTIF(O{R1}:O{R2},"{t}")'
    tot = "+".join(cont(t) for t in SITS)
    linhas = [
        ("Câmeras registradas", f"=COUNTA(C{R1}:C{R2})", None),
        ("Na faixa (OK)", f"={cont('OK')}", None),
        ("A corrigir", f"={tot}-{cont('OK')}", None),
        ("Maior nível do rosto (%)", f'=IF(COUNT(J{R1}:J{R2})=0,"",MAX(J{R1}:J{R2}))', None),
        ("Menor nível do rosto (%)", f'=IF(COUNT(J{R1}:J{R2})=0,"",MIN(J{R1}:J{R2}))', None),
        ("Diferença entre câmeras (pontos)", f'=IF(COUNT(J{R1}:J{R2})<2,"",MAX(J{R1}:J{R2})-MIN(J{R1}:J{R2}))', None),
    ]
    for k, (t, f_, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f_, "E", "F", fmt=fmt, val_merge=f"F{s + k}:G{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(COUNT(L4:L10)<7,"Informe os valores de referência",IF({tot}=0,"Preencha as câmeras",'
                      f'IF({cont("Cintilação")}+{cont("Estouro")}>0,"Há câmera com cintilação ou estouro: corrija primeiro",'
                      f'IF(F{s + 3}>0,"Há câmera a ajustar",IF(AND(F{s + 6}<>"",F{s + 6}>{TOL}),"Rostos com níveis diferentes entre as câmeras","OK")))))',
         b=False, merge=f"F{a}:P{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:P{a}")
    hint(ws, f"B{a + 1}", "Faixas de ângulo (90° a 360°) e de rosto são regra prática deste material, e não norma. Iguale as câmeras pelo resultado, e não pelos números.",
         merge=f"B{a + 1}:Q{a + 1}", h="left")
    end = a + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "Q")
        r = end + 3
        put(ws, f"B{r}", "Câm.", f=font(10, True), bg=GRAY, h="center")
        put(ws, f"C{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:O{r}")
        put(ws, f"P{r}", "Leitura depois", f=font(10, True), bg=GRAY, h="center", merge=f"P{r}:Q{r}")
        ws.row_dimensions[r].height = 21.75
        for cam, feito, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", cam, f=font(10, True, c=MUTED), bg=GRAY, h="center")
            put(ws, f"C{r}", feito, merge=f"C{r}:O{r}")
            put(ws, f"P{r}", depois, merge=f"P{r}:Q{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{R1}"
    setup(ws, tab, f"B1:Q{end}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Exposição — calculadora de stops e ficha por câmera"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Exposição — como usar esta planilha",
           "Calculadora de stops e ficha de ajuste por câmera, para usar na preparação das câmeras.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Modo da câmera e zebra de 100% aceitam só as opções da lista."),
        ("Números", "Digite só o número. Obturador: o x de 1/x (60 para 1/60). Abertura: 2.8 ou 2,8, conforme o separador do seu computador."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Confira, no alto de cada aba, a taxa de quadros do sistema e a rede elétrica (60 Hz no Brasil)."),
        ("2", "Na câmera: exposição em manual, obturador no dobro da taxa de quadros, ISO no mínimo."),
        ("3", "Com a luz do evento e alguém no palco, abra a íris até o rosto chegar ao nível; suba o ISO só se a íris acabar."),
        ("4", "Aba Ficha por câmera: anote o ajuste e o nível do rosto de cada câmera. A câmera 1 é a referência."),
        ("5", "Veja a situação e o que fazer. Corrija na câmera, meça de novo e troque os números pela nova leitura."),
        ("6", "Aba Calculadora de stops: use quando quiser trocar um controle por outro sem mudar o brilho."),
        ("7", "Aba Checklist: confira os doze itens antes do evento."),
    ]),
    ("Abas da planilha", [
        ("Calculadora de stops", "Compara dois ajustes e mostra a diferença em stops, o valor que compensa, o ângulo e a cintilação."),
        ("Ficha por câmera", "Registro de até 8 câmeras, com situação calculada: OK, Cintilação, Estouro, Automático, Igualar e outras."),
        ("Referência", "Obturador por taxa de quadros e por rede, escalas de stops, ganho, filtros ND e níveis de medição."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Calculadora", "A conta do exemplo 1 do treinamento: palco escuro com imagem granulada."),
        ("Exemplo - Ficha", "As quatro câmeras do exemplo 2 do treinamento, em uma formatura."),
    ]),
    ("Regras e premissas do modelo", [
        ("Stops", "1 stop dobra ou corta pela metade a luz. Abertura: 2 × log2 da razão dos números f. Obturador e ISO: log2 da razão. Ganho: 6 dB por stop."),
        ("Valor que compensa", "É o valor que aquele controle de B precisaria ter para a diferença total ficar em zero, com os outros controles de B como estão."),
        ("Cintilação", "OK quando o tempo do obturador cobre um número inteiro de pulsos da luz (2 × a frequência da rede). LED com dimmer e telões podem piscar em outra frequência: teste sempre."),
        ("Movimento", "De 160° a 220°, regra dos 180°. Acima de 360° ou abaixo de 90°, a ficha pede para voltar ao dobro da taxa de quadros."),
        ("Nível do rosto", "Medido na parte mais iluminada da pele, em % na forma de onda, no perfil de imagem padrão. A faixa inicial (45 a 70%) cobre pele morena e clara; para pele negra, baixe o mínimo para perto de 30."),
        ("Situação", "A ficha mostra o problema mais grave de cada câmera: cintilação, obturador, estouro, automático, ganho alto, rosto fora da faixa, e por último a diferença para a câmera 1."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_calc(wb.create_sheet("Calculadora de stops"), BLUE)
aba_ficha(wb.create_sheet("Ficha por câmera"), AMBER)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 20, "D": 20, "E": 20, "F": 52, "G": 2})
title(ws, "Referência — valores típicos de exposição", "Valores de referência, para servir de ponto de partida. Confira o manual da sua câmera e da sua lente.", "F")


def tabela(r, titulo, cab, linhas, alt=21.75):
    band(ws, r, titulo, "F")
    r += 1
    for col, t in zip("BCDEF", cab):
        put(ws, f"{col}{r}", t, f=font(10, True), bg=GRAY, h="center")
    ws.row_dimensions[r].height = 21.75
    for lin in linhas:
        r += 1
        put(ws, f"B{r}", lin[0], f=font(10, True))
        for col, t in zip("CDEF", lin[1:]):
            put(ws, f"{col}{r}", t, h="left" if col == "F" else "center")
        ws.row_dimensions[r].height = alt if len(str(lin[-1])) <= 60 else 31.5
    return r + 2


r = tabela(4, "Obturador por taxa de quadros", ["Taxa de quadros", "Regra dos 180°", "Mais lento, não", "Na rede de 60 Hz", "Observação"], [
    ("24 fps", "1/48 ou 1/50", "1/24", "1/60", "Em 60 Hz, 1/50 pode cintilar: prefira 1/60 (144°)."),
    ("25 fps", "1/50", "1/25", "1/60", "Taxa de países com rede de 50 Hz. Em 60 Hz, use 1/60."),
    ("30 fps (29,97)", "1/60", "1/30", "1/60", "O caso mais comum em culto e palestra."),
    ("50 fps", "1/100", "1/50", "1/60 ou 1/120", "Taxa de países com rede de 50 Hz."),
    ("60 fps (59,94)", "1/120", "1/60", "1/120 ou 1/60", "Algumas câmeras oferecem 1/125. Com pouca luz, 1/60 dá 1 stop a mais."),
])
r = tabela(r, "Obturador e cintilação da luz ligada na rede", ["Rede elétrica", "Pulsos da luz", "Obturadores seguros", "Evite", "Observação"], [
    ("60 Hz (Brasil)", "120 por segundo", "1/30, 1/40, 1/60, 1/120", "1/50, 1/100, mais rápido que 1/120", "LED com dimmer e painel de LED podem piscar em outra frequência: teste."),
    ("50 Hz", "100 por segundo", "1/25, 1/50, 1/100", "1/60, 1/120, mais rápido que 1/100", "Cuidado com ajustes copiados de tutoriais desses países."),
])
r = tabela(r, "Escalas de stops (cada linha para baixo: 1 stop a mais de luz)", ["Stop", "Abertura", "Obturador", "ISO", "Terços de stop entre esta linha e a de baixo"], [
    ("−3", "f/11", "1/500", "100", "f/10 e f/9; ISO 125 e 160"),
    ("−2", "f/8", "1/250", "200", "f/7,1 e f/6,3; ISO 250 e 320"),
    ("−1", "f/5,6", "1/125", "400", "f/5 e f/4,5; ISO 500 e 640"),
    ("0", "f/4", "1/60", "800", "f/3,5 e f/3,2; ISO 1000 e 1250"),
    ("+1", "f/2,8", "1/30", "1600", "f/2,5 e f/2,2; ISO 2000 e 2500"),
    ("+2", "f/2", "1/15", "3200", "f/1,8 e f/1,6; ISO 4000 e 5000"),
    ("+3", "f/1,4", "1/8", "6400", "Obturador abaixo de 1/30 não serve para vídeo em 30 fps."),
])
r = tabela(r, "Ganho em dB e ISO", ["Ganho", "Multiplica o sinal por", "Stops", "Se 0 dB = ISO 800", "Na imagem (varia com o sensor)"], [
    ("0 dB", "1", "0", "ISO 800", "O melhor que a câmera entrega."),
    ("+6 dB", "2", "1", "ISO 1600", "Em geral ainda limpa."),
    ("+12 dB", "4", "2", "ISO 3200", "Ruído começa nas sombras em sensores pequenos."),
    ("+18 dB", "8", "3", "ISO 6400", "Ruído visível em sensores pequenos."),
    ("+24 dB", "16", "4", "ISO 12800", "Ruído evidente na maioria das câmeras."),
])
r = tabela(r, "Filtros ND", ["Nome pelo fator", "Luz que passa", "Nome pela densidade", "Stops que corta", "Uso típico"], [
    ("ND2", "1/2", "ND 0,3", "1", "Ajuste fino."),
    ("ND4", "1/4", "ND 0,6", "2", "Dia nublado, ou sombra."),
    ("ND8", "1/8", "ND 0,9", "3", "Dia claro com câmera de ISO de base baixo."),
    ("ND16", "1/16", "ND 1,2", "4", "Sol, com íris entre f/4 e f/8."),
    ("ND64", "1/64", "ND 1,8", "6", "Sol forte, ou câmera de ISO de base alto."),
])
r = tabela(r, "Níveis na forma de onda (perfil de imagem padrão)", ["O que", "Nível", "Zebra", "False color", "Observação"], [
    ("Estouro", "100% ou mais", "zebra de 100%", "em geral, vermelho", "Nada importante deve ficar aqui. Telão e refletor podem."),
    ("Branco com detalhe", "80 a 90%", "–", "–", "Camisa, papel, toalha da mesa."),
    ("Rosto de pele clara", "60 a 70%", "zebra de 70% começa a riscar", "faixa da pele", "Parte mais iluminada do rosto."),
    ("Rosto de pele morena", "45 a 60%", "–", "faixa da pele ou abaixo", "Use a forma de onda, ou baixe o nível da zebra."),
    ("Rosto de pele negra", "30 a 45%", "–", "abaixo da faixa da pele", "Bem exposto mesmo sem nenhuma listra de zebra."),
    ("Preto sem detalhe", "abaixo de 5%", "–", "em geral, roxo ou azul", "Normal no fundo de um palco escuro."),
])
hint(ws, f"B{r - 1}", "Faixas típicas. Mudam com a câmera, a lente, o perfil de imagem e a pessoa. Em perfis log, todos os níveis ficam mais baixos.",
     merge=f"B{r - 1}:F{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:F{r - 1}", landscape=False, fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplos
checklist(wb.create_sheet("Checklist"), "Checklist da exposição",
          "Confira na preparação, com a luz do evento ligada e alguém no palco. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=TEAL)
aba_calc(wb.create_sheet("Exemplo - Calculadora"), MUTED, EX_CALC)
aba_ficha(wb.create_sheet("Exemplo - Ficha"), MUTED, EX_FICHA)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
