# -*- coding: utf-8 -*-
"""Gera Sinais-de-Video-modelo.xlsx: mapa de sinais de vídeo, inventário de cabos e conversores,
roteiro de diagnóstico, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIM_NAO = ["Sim", "Não"]

# ------------------------------------------------------------------ dados compartilhados
# formato (como aparece nos menus), resolução, varredura, imagens por segundo, classe, nº da classe, SDI, cabo HDMI
FORMATOS = [
    ("720p50", "1280 × 720", "Progressiva", "50 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("720p59.94", "1280 × 720", "Progressiva", "59,94 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("720p60", "1280 × 720", "Progressiva", "60 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("1080i50", "1920 × 1080", "Entrelaçada", "50 campos", "HD", 1, "HD-SDI", "High Speed"),
    ("1080i59.94", "1920 × 1080", "Entrelaçada", "59,94 campos", "HD", 1, "HD-SDI", "High Speed"),
    ("1080i60", "1920 × 1080", "Entrelaçada", "60 campos", "HD", 1, "HD-SDI", "High Speed"),
    ("1080p23.98", "1920 × 1080", "Progressiva", "23,98 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("1080p24", "1920 × 1080", "Progressiva", "24 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("1080p25", "1920 × 1080", "Progressiva", "25 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("1080p29.97", "1920 × 1080", "Progressiva", "29,97 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("1080p30", "1920 × 1080", "Progressiva", "30 quadros", "HD", 1, "HD-SDI", "High Speed"),
    ("1080p50", "1920 × 1080", "Progressiva", "50 quadros", "3G", 2, "3G-SDI", "High Speed"),
    ("1080p59.94", "1920 × 1080", "Progressiva", "59,94 quadros", "3G", 2, "3G-SDI", "High Speed"),
    ("1080p60", "1920 × 1080", "Progressiva", "60 quadros", "3G", 2, "3G-SDI", "High Speed"),
    ("2160p25", "3840 × 2160", "Progressiva", "25 quadros", "6G", 3, "6G-SDI", "High Speed"),
    ("2160p29.97", "3840 × 2160", "Progressiva", "29,97 quadros", "6G", 3, "6G-SDI", "High Speed"),
    ("2160p30", "3840 × 2160", "Progressiva", "30 quadros", "6G", 3, "6G-SDI", "High Speed"),
    ("2160p50", "3840 × 2160", "Progressiva", "50 quadros", "12G", 4, "12G-SDI", "Premium High Speed"),
    ("2160p59.94", "3840 × 2160", "Progressiva", "59,94 quadros", "12G", 4, "12G-SDI", "Premium High Speed"),
    ("2160p60", "3840 × 2160", "Progressiva", "60 quadros", "12G", 4, "12G-SDI", "Premium High Speed"),
    ("Outro", "resolução de computador", "–", "–", "3G", 2, "confira no manual", "confira no manual"),
]
NOMES_FMT = [f[0] for f in FORMATOS]
assert len(",".join(NOMES_FMT)) < 255
REF_R1 = 6
REF_R2 = REF_R1 + len(FORMATOS) - 1
FMT = f"'Referência'!$B${REF_R1}:$I${REF_R2}"

# interface, limite por classe (HD, 3G, 6G, 12G), dica
LIMITES = [
    ("HDMI passivo", (10, 10, 5, 3), "Cabo comum. Acima de 5 m, só cabo de boa qualidade e testado."),
    ("HDMI ativo ou óptico", (30, 30, 30, 30), "Conforme o fabricante: há cabos ópticos de até 100 m. Têm sentido certo."),
    ("SDI", (100, 70, 50, 30), "Coaxial de 75 ohms de bitola cheia. Com cabo fino, use a metade."),
    ("HDBaseT", (70, 70, 40, 40), "Conforme a classe do extensor e a categoria do cabo de rede."),
    ("Fibra", (300, 300, 300, 300), "Conforme o conversor. Fibra monomodo alcança quilômetros."),
    ("Rede (NDI)", (100, 100, 100, 100), "Por trecho de cabo, entre o aparelho e o switch de rede."),
    ("USB", (3, 3, 3, 3), "USB 3 com cabo passivo. Em USB 2, até 5 m."),
]
INTERFACES = [l[0] for l in LIMITES]

SIT_CF = [("OK", GREEN), ("No limite", YELLOW), ("Incompleto", YELLOW), ("Formato", RED), ("Comprimento", RED),
          ("Formato e cabo", RED)]

CHECK = [
    ("Preparação", "O padrão do sistema (resolução, varredura e taxa de quadros) está definido e anotado."),
    ("Preparação", "O diagrama de sinal está desenhado, e o mapa de sinais tem o comprimento de cada trecho."),
    ("Preparação", "Cada câmera e cada computador está configurado para enviar o padrão do sistema."),
    ("Cabos", "Nenhum trecho passa do comprimento recomendado: os longos usam SDI, extensor, cabo óptico ou fibra."),
    ("Cabos", "Os cabos ativos e ópticos estão no sentido certo, com a ponta de origem na origem."),
    ("Cabos", "Os conectores BNC estão travados, os HDMI têm alívio de tensão, e os cabos no chão estão presos."),
    ("Conversores", "Cada conversor, extensor e distribuidor está alimentado, com a fonte presa e a luz de sinal acesa."),
    ("Conversores", "A saída de cada computador está em 1920 × 1080, na taxa do sistema, estendida ou duplicada conforme o uso."),
    ("Teste", "Cada entrada do switcher mostra imagem estável por dez minutos, sem piscar nem chuviscar."),
    ("Teste", "Cada destino mostra a imagem inteira, sem cortes, sem barras inesperadas e sem esticar."),
    ("Teste", "O áudio embutido chega aonde deve, e o som do computador sai pela saída prevista."),
    ("Teste", "As mídias do evento foram testadas no caminho real, e há um cabo curto testado e um conversor de reserva."),
]

# Exemplo 1 do treinamento: igreja de 200 lugares, três câmeras, antes das correções
EX_MAPA = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "resp": "Operador de vídeo da escala"},
    "padrao": "1080p59.94",
    "trechos": [  # origem, formato enviado, ligação, comprimento, destino, destino converte?
        ("Câmera 1 (fundo, central)", "1080p59.94", "SDI", 40, "Switcher, entrada 1", "Não"),
        ("Câmera 2 (lateral)", "1080p59.94", "HDMI passivo", 15, "Switcher, entrada 2", "Não"),
        ("Câmera 3 (palco)", "1080i59.94", "HDMI passivo", 3, "Switcher, entrada 3", "Não"),
        ("Computador das letras", "1080p60", "HDMI passivo", 2, "Switcher, entrada 4", "Não"),
        ("Notebook de visitante", "Outro", "HDMI passivo", 5, "Scaler", "Sim"),
        ("Scaler", "1080p59.94", "HDMI passivo", 1, "Switcher, entrada 5", "Não"),
        ("Switcher, saída de programa 1", "1080p59.94", "HDBaseT", 35, "Projetor (telão)", "Não"),
        ("Switcher, saída de programa 2", "1080p59.94", "HDMI ativo ou óptico", 20, "TV de retorno do palco", "Não"),
        ("Switcher, saída auxiliar", "1080p59.94", "SDI", 2, "Placa de captura", "Não"),
        ("Placa de captura", "1080p59.94", "USB", 3, "Computador da transmissão", "Não"),
    ],
    "ajustes": [  # trecho, o que foi feito, depois
        ("2", "Conversor HDMI → SDI preso ao tripé da câmera e cabo coaxial de 15 m até o switcher.", "Ligação: SDI. Situação: OK"),
        ("3", "Saída HDMI da câmera mudada de 1080i59.94 para 1080p59.94 no menu.", "Formato: Igual. Situação: OK"),
        ("4", "Frequência da saída do computador mudada de 60 para 59,94 Hz.", "Formato: Igual. Situação: OK"),
        ("5", "Nada: o scaler recebe a resolução do notebook e entrega o padrão.", "Formato: Convertido"),
        ("10", "Cabo USB trocado por um de 1 m, com o computador ao lado do switcher.", "Cabo: OK. Situação: OK"),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def gray_head(ws, ref, text, merge=None):
    put(ws, ref, text, f=font(10, True), bg=GRAY, h="center", merge=merge)


# ================================================================== aba Mapa de sinais
M_W = {"A": 2, "B": 5, "C": 30, "D": 14, "E": 22, "F": 11, "G": 30, "H": 11, "I": 9, "J": 9, "K": 12, "L": 11,
       "M": 16, "N": 54, "O": 28, "P": 2}
M1, MN = 17, 20
PAD, PCT = "$D$8", "$D$9"
LIM_NOMES, LIM_VAL = "$G$5:$G$11", "$H$5:$K$11"


def aba_mapa(ws, tab, data=None):
    widths(ws, M_W)
    is_ex = data is not None
    title(ws, "Mapa de sinais de vídeo — origem, formato, cabo e destino",
          "Exemplo preenchido, para consulta. Use a aba Mapa de sinais para o seu sistema." if is_ex
          else "Preencha as células em amarelo-claro, um trecho por linha. As células cinza são calculadas.", "O")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento ou sistema", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Responsável", "resp", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:F{r}", bg=hb, fmt=fmt)
    dv_date(ws, "D6")
    label(ws, "B8", "Padrão do sistema", merge="B8:C8")
    put(ws, "D8", data["padrao"] if is_ex else "1080p59.94", f=font(10, True), bg=hb, h="center", merge="D8:F8")
    dv_list(ws, "D8", NOMES_FMT, "O formato em que todas as origens devem estar, como aparece nos menus.")
    label(ws, "B9", "Aviso “No limite” a partir de", merge="B9:C9")
    inp(ws, "D9", 0.8, h="center", bg=hb, fmt="0%", merge="D9:F9")
    dv_number(ws, "D9")
    hint(ws, "B10", "O formato é escrito como nos menus dos equipamentos, com ponto: 1080p59.94.", merge="B10:F10", h="left")
    hint(ws, "B11", "59.94 e 60 são formatos diferentes. Confira o que cada aparelho entrega de fato.", merge="B11:F11", h="left")

    # limites de comprimento, por tipo de ligação e classe do formato
    put(ws, "G4", "Limite recomendado (m)", f=font(10, True, c=WHITE), bg=INK)
    for col, t in zip("HIJK", ["HD", "3G", "6G", "12G"]):
        head(ws, f"{col}4", t, bg=AMBER)
    hint(ws, "L4", "HD: 720p, 1080i e 1080p até 30 · 3G: 1080p de 50 a 60 · 6G: 4K até 30 · 12G: 4K de 50 a 60",
         merge="L4:O4", h="left")
    for k, (nome, lims, dica) in enumerate(LIMITES):
        r = 5 + k
        put(ws, f"G{r}", nome, f=font(10, True), bg=GRAY)
        for col, v in zip("HIJK", lims):
            inp(ws, f"{col}{r}", v, h="center", bg=hb)
        hint(ws, f"L{r}", dica, merge=f"L{r}:O{r}", h="left")
    dv_number(ws, "H5:K11")
    for r in range(4, 12):
        ws.row_dimensions[r].height = 21.75

    band(ws, 13, "Trechos do sistema, da origem para o destino", "O")
    heads = [("B", "#", INK), ("C", "Origem", INK), ("D", "Formato enviado", BLUE), ("E", "Ligação", AMBER),
             ("F", "Comprimento", AMBER), ("G", "Destino", INK), ("H", "Destino converte?", BLUE),
             ("I", "Classe", TEAL), ("J", "Limite", TEAL), ("K", "Formato", TEAL), ("L", "Cabo", TEAL),
             ("M", "Situação", TEAL), ("N", "O que fazer", TEAL), ("O", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}14", t, bg=bg)
    ws.row_dimensions[14].height = 30
    for col, t in zip("BCDEFGHIJKLMNO", ["", "aparelho e saída", "saída da origem", "cabo ou extensor", "metros",
                                         "aparelho e entrada", "Sim / Não", "do formato", "metros", "× padrão",
                                         "× limite", "", "", ""]):
        hint(ws, f"{col}15", t)
    ws.row_dimensions[15].height = 21.75
    if is_ex:
        hint(ws, "B16", "Situação antes das correções. As colunas de I a N são calculadas. O que foi ajustado está no fim da aba.",
             merge="B16:O16", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMN", ["Ex.", "Câmera 1 (fundo)", "1080p59.94", "SDI", 40, "Switcher, entrada 1", "Não",
                                            "3G", 70, "Igual", "OK", "OK", "Manter"]):
            ex(ws, f"{col}16", v, h="left" if col in "CGN" else "center")
        ex(ws, "O16", "")
        note(ws, "D14", "O formato que a origem envia de fato, lido no menu de saída dela.\nResolução de computador fora da lista: escolha Outro.")
        note(ws, "E14", "O tipo de cabo ou de extensor do trecho.\nTrecho com conversor no meio: registre cada lado em uma linha.")
        note(ws, "F14", "Comprimento pelo caminho real do cabo, com subidas, contornos e folgas.")
        note(ws, "H14", "Sim quando o destino tem conversor de formato na entrada (scaler, ou switcher que aceita qualquer formato).")
        note(ws, "I14", "Quanto o formato exige do cabo. Os nomes vêm dos níveis do SDI e valem aqui para qualquer ligação.")
        note(ws, "J14", "Vem da tabela de limites do alto da aba, pelo tipo de ligação e pela classe do formato.")
        note(ws, "K14", "Igual: é o padrão do sistema.\nConvertido: é diferente, mas o destino converte.\nDiferente: o destino pode recusar.")
    ws.row_dimensions[16].height = 24

    M2 = M1 + MN - 1
    bg = WHITE if is_ex else INPUT
    for k in range(MN):
        r = M1 + k
        t = data["trechos"][k] if is_ex and k < len(data["trechos"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", t[0] if t else None, bg=bg)
        inp(ws, f"D{r}", t[1] if t else None, h="center", bg=bg)
        inp(ws, f"E{r}", t[2] if t else None, h="center", bg=bg)
        inp(ws, f"F{r}", t[3] if t else None, h="center", bg=bg)
        inp(ws, f"G{r}", t[4] if t else None, bg=bg)
        inp(ws, f"H{r}", t[5] if t else None, h="center", bg=bg)
        calc(ws, f"I{r}", f'=IF(D{r}="","",IFERROR(VLOOKUP(D{r},{FMT},5,0),""))', b=False)
        calc(ws, f"J{r}", f'=IF(OR(D{r}="",E{r}=""),"",IFERROR(1/(1/INDEX({LIM_VAL},MATCH(E{r},{LIM_NOMES},0),'
                          f'VLOOKUP(D{r},{FMT},6,0))),""))', b=False)
        calc(ws, f"K{r}", f'=IF(OR(D{r}="",{PAD}=""),"",IF(D{r}={PAD},"Igual",IF(H{r}="Sim","Convertido","Diferente")))')
        calc(ws, f"L{r}", f'=IF(OR(F{r}="",J{r}=""),"",IF(F{r}>J{r},"Acima",IF(F{r}>J{r}*IF({PCT}="",1,{PCT}),"No limite","OK")))')
        calc(ws, f"M{r}", f'=IF(AND(K{r}="",L{r}=""),"",IF(AND(K{r}="Diferente",L{r}="Acima"),"Formato e cabo",'
                          f'IF(K{r}="Diferente","Formato",IF(L{r}="Acima","Comprimento",IF(L{r}="No limite","No limite",'
                          f'IF(OR(K{r}="",L{r}=""),"Incompleto","OK"))))))')
        calc(ws, f"N{r}", f'=IF(M{r}="","",IF(M{r}="Formato e cabo","Ajuste a origem para "&{PAD}&" e troque o tipo de ligação do trecho",'
                          f'IF(M{r}="Formato","Ajuste a saída da origem para "&{PAD}&", ou use um scaler antes do destino",'
                          f'IF(M{r}="Comprimento",IF(E{r}="HDMI passivo","Use SDI com conversor, HDBaseT ou cabo HDMI óptico",'
                          f'IF(E{r}="USB","Aproxime o computador ou use extensão USB ativa",'
                          f'"Encurte o trecho, use cabo melhor ou passe para fibra")),'
                          f'IF(M{r}="No limite","Teste com o cabo no caminho real e tenha um reserva",'
                          f'IF(M{r}="Incompleto","Informe o formato, a ligação e o comprimento","Manter"))))))',
             h="left", b=False, sz=9)
        inp(ws, f"O{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{M1}:D{M2}", NOMES_FMT, "O formato que a origem envia, como aparece no menu dela.")
    dv_list(ws, f"E{M1}:E{M2}", INTERFACES, "O tipo de cabo ou de extensor do trecho.")
    dv_number(ws, f"F{M1}:F{M2}")
    dv_list(ws, f"H{M1}:H{M2}", SIM_NAO, "O destino tem conversor de formato na entrada?")
    cf_equal(ws, f"K{M1}:K{M2}", [("Igual", GREEN), ("Convertido", YELLOW), ("Diferente", RED)])
    cf_equal(ws, f"L{M1}:L{M2}", [("OK", GREEN), ("No limite", YELLOW), ("Acima", RED)])
    cf_equal(ws, f"M{M1}:M{M2}", SIT_CF)

    s = M2 + 2
    band(ws, s, "Resumo automático", "O")
    cont = lambda col, t: f'COUNTIF({col}{M1}:{col}{M2},"{t}")'
    linhas = [
        ("Trechos registrados", f"=COUNTA(C{M1}:C{M2})", None),
        ("Trechos com situação OK", f'={cont("M", "OK")}', None),
        ("Formato diferente do padrão", f'={cont("M", "Formato")}+{cont("M", "Formato e cabo")}', None),
        ("Comprimento acima do recomendado", f'={cont("M", "Comprimento")}+{cont("M", "Formato e cabo")}', None),
        ("Comprimento no limite", f'={cont("M", "No limite")}', None),
        ("Trechos com conversão de formato", f'={cont("K", "Convertido")}', None),
        ("Cabo total (m)", f'=IF(COUNT(F{M1}:F{M2})=0,"",SUM(F{M1}:F{M2}))', None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", fmt=fmt)
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF({PAD}="","Informe o padrão do sistema",IF(F{s + 1}=0,"Liste os trechos",'
                      f'IF(F{s + 3}>0,"Há origem em formato diferente do padrão: corrija antes do ensaio",'
                      f'IF(F{s + 4}>0,"Há trecho acima do comprimento recomendado: mude o tipo de ligação",'
                      f'IF({cont("M", "Incompleto")}>0,"Há trecho incompleto",'
                      f'IF(F{s + 5}>0,"Há trecho no limite: teste no caminho real e leve reserva","OK"))))))',
         b=False, merge=f"F{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:N{a}")
    hint(ws, f"B{a + 1}", "Os limites são pontos de partida conservadores, e não especificação. Troque-os pelos valores do manual "
                          "dos seus cabos e extensores, ou pelo que você comprovou em teste.", merge=f"B{a + 1}:O{a + 1}", h="left")
    end = a + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "O")
        r = end + 3
        gray_head(ws, f"B{r}", "Trecho", merge=f"B{r}:C{r}")
        gray_head(ws, f"D{r}", "O que foi feito", merge=f"D{r}:L{r}")
        gray_head(ws, f"M{r}", "Depois", merge=f"M{r}:O{r}")
        ws.row_dimensions[r].height = 21.75
        for trecho, feito, depois in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", trecho, f=font(10, True, c=MUTED), bg=GRAY, h="center", merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:L{r}")
            put(ws, f"M{r}", depois, merge=f"M{r}:O{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{M1}"
    setup(ws, tab, f"B1:O{end}", fit_height=True)


# ================================================================== aba Inventário
I_W = {"A": 2, "B": 5, "C": 34, "D": 20, "E": 13, "F": 12, "G": 8, "H": 11, "I": 11, "J": 13, "K": 14, "L": 22,
       "M": 9, "N": 15, "O": 44, "P": 26, "Q": 2}
I1, IN = 13, 30
REFD, PRAZO = "$D$6", "$D$7"


def aba_inventario(ws, tab):
    widths(ws, I_W)
    title(ws, "Inventário de cabos e conversores de vídeo",
          "Um item por linha. Teste cada cabo no formato do sistema e anote a data. As células cinza são calculadas.", "P")
    for r, rot in [(4, "Local ou sistema"), (5, "Responsável"), (6, "Data da conferência"), (7, "Testar de novo após (dias)")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", 180 if r == 7 else None, merge=f"D{r}:F{r}", fmt=DATE if r == 6 else None,
            h="center" if r == 7 else "left")
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    dv_number(ws, "D7")
    for r, t in [(4, "Identifique cada cabo nas duas pontas com o mesmo nome que está nesta lista."),
                 (5, "Teste do cabo: no formato do sistema, por alguns minutos, mexendo de leve nos conectores."),
                 (6, "Cabo intermitente é cabo com defeito: tire do uso antes que falhe no evento."),
                 (7, "Guarde cada conversor junto da própria fonte e do cabo curto que o acompanha.")]:
        hint(ws, f"H{r}", t, merge=f"H{r}:P{r}", h="left")

    band(ws, 9, "Cabos, conversores, extensores e placas de captura", "P")
    heads = [("B", "#", INK), ("C", "Item e identificação", INK), ("D", "Tipo", INK), ("E", "Comprimento", AMBER),
             ("F", "Suporta até", AMBER), ("G", "Qtde", INK), ("H", "Precisa de fonte?", BLUE), ("I", "Fonte junto?", BLUE),
             ("J", "Testado em", BLUE), ("K", "Resultado", BLUE), ("L", "Onde fica", INK), ("M", "Dias", TEAL),
             ("N", "Situação", TEAL), ("O", "O que fazer", TEAL), ("P", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLMNOP", ["", "nome da etiqueta", "", "metros", "classe", "", "Sim / Não", "Sim / Não",
                                          "data", "do teste", "", "desde o teste", "", "", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 21.75
    for col, v in zip("BCDEFGHIJKLMNO", ["Ex.", "SDI-01, coaxial preto", "Cabo SDI", 40, "3G", 1, "Não", "", "04/10/2026", "OK",
                                         "Caixa de cabos 2", 0, "OK", "Manter"]):
        ex(ws, f"{col}12", v, h="left" if col in "CLO" else "center")
    ex(ws, "P12", "")
    ws.row_dimensions[12].height = 24
    note(ws, "F10", "O formato mais exigente que o item comprovadamente passa.\nHD: até 1080i ou 1080p30. 3G: 1080p60. 6G: 4K30. 12G: 4K60.")
    note(ws, "M10", "Dias entre o último teste e a data da conferência informada no alto da aba.")
    note(ws, "N10", "Retirar: defeito ou intermitente.\nFalta fonte: precisa de fonte e ela não está junto.\nTestar: sem teste registrado.\nTestar de novo: teste vencido.")

    I2 = I1 + IN - 1
    for k in range(IN):
        r = I1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}")
        for col in "DEFGHIK":
            inp(ws, f"{col}{r}", h="center")
        inp(ws, f"J{r}", h="center", fmt=DATE)
        inp(ws, f"L{r}")
        calc(ws, f"M{r}", f'=IF(OR(J{r}="",{REFD}=""),"",{REFD}-J{r})', b=False, fmt="0")
        calc(ws, f"N{r}", f'=IF(C{r}="","",IF(OR(K{r}="Com defeito",K{r}="Intermitente"),"Retirar",'
                          f'IF(AND(H{r}="Sim",I{r}<>"Sim"),"Falta fonte",IF(OR(J{r}="",K{r}=""),"Testar",'
                          f'IF(AND(M{r}<>"",{PRAZO}<>"",M{r}>{PRAZO}),"Testar de novo","OK")))))')
        calc(ws, f"O{r}", f'=IF(N{r}="","",IF(N{r}="Retirar","Marque como defeituoso e tire do uso",'
                          f'IF(N{r}="Falta fonte","Localize a fonte e guarde junto do aparelho",'
                          f'IF(N{r}="Testar","Teste no formato do sistema e anote a data",'
                          f'IF(N{r}="Testar de novo","Teste vencido: repita antes do próximo evento","Manter")))))',
             h="left", b=False, sz=9)
        inp(ws, f"P{r}")
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{I1}:D{I2}", ["Cabo HDMI", "Cabo HDMI óptico", "Cabo SDI", "Cabo de rede", "Cabo USB", "Fibra", "Conversor",
                                 "Scaler", "Distribuidor", "Extensor", "Placa de captura", "Adaptador"], "O tipo do item.")
    dv_number(ws, f"E{I1}:E{I2}")
    dv_list(ws, f"F{I1}:F{I2}", ["HD", "3G", "6G", "12G", "Não sei"], "HD, 3G, 6G, 12G ou Não sei.")
    dv_number(ws, f"G{I1}:G{I2}")
    dv_list(ws, f"H{I1}:I{I2}", SIM_NAO, "Sim ou Não.")
    dv_date(ws, f"J{I1}:J{I2}")
    dv_list(ws, f"K{I1}:K{I2}", ["OK", "Intermitente", "Com defeito"], "Resultado do último teste.")
    cf_equal(ws, f"N{I1}:N{I2}", [("OK", GREEN), ("Testar", YELLOW), ("Testar de novo", YELLOW), ("Falta fonte", YELLOW),
                                  ("Retirar", RED)])

    s = I2 + 2
    band(ws, s, "Resumo automático", "P")
    cont = lambda t: f'COUNTIF(N{I1}:N{I2},"{t}")'
    linhas = [
        ("Itens registrados", f"=COUNTA(C{I1}:C{I2})"),
        ("Em condições de uso (OK)", f'={cont("OK")}'),
        ("A testar", f'={cont("Testar")}+{cont("Testar de novo")}'),
        ("Sem a fonte", f'={cont("Falta fonte")}'),
        ("A retirar do uso", f'={cont("Retirar")}'),
    ]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste os itens",IF(F{s + 5}>0,"Há item com defeito: retire antes do evento",'
                      f'IF(F{s + 4}>0,"Há aparelho sem a fonte",IF(F{s + 3}>0,"Há item sem teste em dia","OK"))))',
         b=False, merge=f"F{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"F{a}:N{a}")
    ws.freeze_panes = f"D{I1}"
    setup(ws, tab, f"B1:P{a}", fit_height=True)


# ================================================================== aba Diagnóstico
D_W = {"A": 2, "B": 5, "C": 22, "D": 44, "E": 12, "F": 62, "G": 30, "H": 2}
PASSOS = [  # pergunta, o que fazer se a resposta for Não
    ("O destino (projetor, TV, switcher) está ligado e na entrada em que o cabo está?",
     "Ligue o destino e selecione a entrada certa."),
    ("A origem está ligada, acordada e com a saída de vídeo ativa? No computador: tela estendida ou duplicada.",
     "Ligue ou acorde a origem. No computador, ative a segunda tela."),
    ("Os conectores estão encaixados até o fim nas duas pontas, e os BNC, travados?",
     "Refaça os encaixes, uma ponta por vez, olhando a imagem."),
    ("O cabo ativo ou óptico está no sentido certo, com a ponta de origem na origem? Com cabo comum, responda Sim.",
     "Inverta o cabo."),
    ("Os conversores, extensores e distribuidores do trecho estão com a luz de energia acesa?",
     "Religue a fonte. Se a luz não acender, troque de tomada ou de fonte."),
    ("O formato de saída da origem é o padrão do sistema, ou um formato que o destino aceita?",
     "Ajuste a saída da origem para o padrão do sistema e espere alguns segundos."),
    ("A origem é livre de proteção (câmera, ou computador com conteúdo próprio)?",
     "Provável HDCP: teste a origem direto em uma TV e troque a fonte do conteúdo."),
    ("O comprimento do trecho está dentro do recomendado na aba Mapa de sinais?",
     "Teste com um cabo mais curto. Se resolver, mude o tipo de ligação do trecho."),
    ("A origem, ligada direto em um monitor por um cabo curto e testado, mostra imagem?",
     "O defeito está na origem: confira o menu de saída, reinicie e teste outra porta."),
    ("Outra origem que funciona, ligada no lugar desta, aparece no destino?",
     "O defeito está no cabo, no conversor ou na entrada do destino: troque um de cada vez."),
]
P1 = 12


def aba_diagnostico(ws, tab):
    widths(ws, D_W)
    title(ws, "Diagnóstico — roteiro para “sem sinal”",
          "Responda na ordem. No primeiro “Não”, faça a correção indicada, teste e só então siga adiante.", "G")
    for r, rot in [(4, "Evento"), (5, "Data"), (6, "Trecho (origem → destino)"), (7, "Sintoma"), (8, "Quem conferiu")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:E{r}", fmt=DATE if r == 5 else None)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    dv_list(ws, "D7", ["Sem sinal", "Imagem pisca", "Chuvisco ou pontos", "Cor errada", "Movimento aos trancos",
                       "Imagem cortada ou esticada", "Sem som", "Outro"], "O que se vê no destino.")
    for r, t in [(4, "Mude uma coisa de cada vez e teste depois de cada mudança."),
                 (5, "Tenha à mão um cabo curto testado e um monitor ou TV de teste."),
                 (6, "Imagem que pisca ou chuvisca aponta para o cabo: comece pelos itens 3, 4 e 8."),
                 (7, "Tela preta só com um conteúdo ou um aparelho aponta para o item 7."),
                 (8, "Depois de resolver, registre a ocorrência no fim desta aba.")]:
        hint(ws, f"F{r}", t, merge=f"F{r}:G{r}", h="left")

    band(ws, 10, "Roteiro, do destino para a origem", "G")
    head(ws, "B11", "#")
    put(ws, "C11", "Pergunta", f=font(10, True, c=WHITE), bg=INK, h="center", merge="C11:D11")
    head(ws, "E11", "Resposta", bg=AMBER)
    head(ws, "F11", "Se a resposta for Não", bg=TEAL)
    head(ws, "G11", "Observação")
    ws.row_dimensions[11].height = 21.75
    P2 = P1 + len(PASSOS) - 1
    for k, (perg, acao) in enumerate(PASSOS):
        r = P1 + k
        num(ws, f"B{r}", k + 1)
        put(ws, f"C{r}", perg, bg=GRAY, merge=f"C{r}:D{r}")
        inp(ws, f"E{r}", h="center")
        put(ws, f"F{r}", acao, bg=GRAY)
        inp(ws, f"G{r}")
        ws.row_dimensions[r].height = 31.5
    dv_list(ws, f"E{P1}:E{P2}", SIM_NAO, "Sim ou Não.")
    cf_equal(ws, f"E{P1}:E{P2}", [("Sim", GREEN), ("Não", RED)])

    s = P2 + 2
    band(ws, s, "Resultado", "G")
    resp = f"COUNTA(E{P1}:E{P2})"
    summary(ws, s + 1, "Itens respondidos", f"={resp}", "D", "E")
    summary(ws, s + 2, "Primeiro item com resposta Não", f'=IF(COUNTIF(E{P1}:E{P2},"Não")=0,"",MATCH("Não",E{P1}:E{P2},0))', "D", "E")
    label(ws, f"B{s + 3}", "O que fazer agora", merge=f"B{s + 3}:D{s + 3}", h="right")
    calc(ws, f"E{s + 3}", f'=IF(E{s + 2}<>"",INDEX(F{P1}:F{P2},E{s + 2}),IF({resp}=0,"Responda às perguntas na ordem",'
                          f'IF({resp}<{len(PASSOS)},"Continue respondendo na ordem",'
                          f'"Tudo conferido. Religue na ordem destino, meio, origem. Se continuar, troque um aparelho por vez")))',
         b=False, merge=f"E{s + 3}:G{s + 3}", h="left")
    label(ws, f"B{s + 4}", "Situação", merge=f"B{s + 4}:D{s + 4}", h="right")
    calc(ws, f"E{s + 4}", f'=IF({resp}=0,"",IF(E{s + 2}<>"","Corrigir o item "&E{s + 2},IF({resp}<{len(PASSOS)},"Em andamento","OK")))',
         merge=f"E{s + 4}:G{s + 4}", h="left")
    for r in (s + 3, s + 4):
        ws.row_dimensions[r].height = 21.75
    cf_status(ws, f"E{s + 4}:G{s + 4}", f"$E${s + 4}")

    o = s + 6
    band(ws, o, "Registro de ocorrências", "G")
    head(ws, f"B{o + 1}", "#")
    head(ws, f"C{o + 1}", "Data")
    put(ws, f"D{o + 1}", "Trecho e sintoma", f=font(10, True, c=WHITE), bg=INK, h="center", merge=f"D{o + 1}:E{o + 1}")
    head(ws, f"F{o + 1}", "Causa encontrada e o que resolveu")
    head(ws, f"G{o + 1}", "Para não repetir")
    ws.row_dimensions[o + 1].height = 21.75
    for k in range(8):
        r = o + 2 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center", fmt=DATE)
        inp(ws, f"D{r}", merge=f"D{r}:E{r}")
        inp(ws, f"F{r}")
        inp(ws, f"G{r}")
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, f"C{o + 2}:C{o + 9}")
    setup(ws, tab, f"B1:G{o + 9}", fit_height=True)


# ================================================================== aba Referência
R_W = {"A": 2, "B": 18, "C": 22, "D": 14, "E": 15, "F": 9, "G": 8, "H": 20, "I": 40, "J": 2}


def aba_referencia(ws, tab):
    widths(ws, R_W)
    title(ws, "Referência — formatos, ligações e sintomas",
          "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "I")
    band(ws, 4, "Formatos de vídeo (tabela consultada pelas outras abas: não altere a primeira coluna)", "I")
    for col, t in zip("BCDEFGHI", ["Formato", "Resolução", "Varredura", "Imagens por segundo", "Classe", "Nº",
                                   "SDI necessário", "Cabo HDMI indicado"]):
        gray_head(ws, f"{col}5", t)
    ws.row_dimensions[5].height = 30
    for k, f in enumerate(FORMATOS):
        r = REF_R1 + k
        put(ws, f"B{r}", f[0], f=font(10, True), h="center")
        for col, v in zip("CDEFGHI", f[1:]):
            put(ws, f"{col}{r}", v, h="center")
    r = REF_R2 + 1
    hint(ws, f"B{r}", "A classe mostra quanto o formato exige do cabo. “Outro” é tratado como 1080p a 60 quadros. "
                      "Os mesmos números aparecem nos menus com ponto: 59.94.", merge=f"B{r}:I{r}", h="left")

    def tabela(r, titulo, spans, cab, linhas, alt=31.5):
        band(ws, r, titulo, "I")
        r += 1
        for (c1, c2), t in zip(spans, cab):
            gray_head(ws, f"{c1}{r}", t, merge=f"{c1}{r}:{c2}{r}" if c1 != c2 else None)
        ws.row_dimensions[r].height = 21.75
        for lin in linhas:
            r += 1
            for j, ((c1, c2), t) in enumerate(zip(spans, lin)):
                put(ws, f"{c1}{r}", t, f=font(10, j == 0), merge=f"{c1}{r}:{c2}{r}" if c1 != c2 else None)
            ws.row_dimensions[r].height = alt
        return r + 2

    r = tabela(r + 2, "Tipos de ligação", [("B", "B"), ("C", "D"), ("E", "G"), ("H", "I")],
               ["Ligação", "Conector e cabo", "Distância prática", "Ponto de atenção"], [
        ("HDMI passivo", "HDMI normal, mini ou micro. Cabo de 19 vias.", "Até 5 m sem preocupação. De 5 a 10 m, com cabo bom.",
         "Sem trava. Leva EDID e HDCP. O 4K a 60 pede cabo Premium High Speed e trecho de até 3 m."),
        ("HDMI ativo ou óptico", "HDMI com circuito ou fibra dentro do cabo.", "De 10 a 100 m, conforme o fabricante.",
         "Tem sentido certo: a ponta de origem vai na origem. Não dobre em ângulo fechado."),
        ("SDI", "BNC com trava. Coaxial de 75 ohms.", "HD: cerca de 100 m. 3G: de 70 a 100 m. 12G: de 30 a 70 m.",
         "Não carrega HDCP nem EDID: o formato tem de estar certo nos dois lados. Cabo de 50 ohms não serve."),
        ("HDBaseT", "RJ45. Cabo de rede, de um extensor ao outro.", "De 40 a 100 m, conforme a classe e a resolução.",
         "Ponto a ponto: não passa por switch de rede. Cabo de cobre inteiro, de categoria 6 ou melhor."),
        ("Fibra", "Conversores nas pontas e cabo óptico.", "Centenas de metros a quilômetros.",
         "Isola os aparelhos eletricamente. Conectores limpos e com tampa. Nunca olhe para a ponta ligada."),
        ("Rede (NDI)", "RJ45. Rede local de 1 gigabit, cabeada.", "100 m por trecho de cabo.",
         "Vídeo comprimido e com algum atraso. Rede separada da do público. Wi-Fi não é confiável."),
        ("USB", "USB-A ou USB-C. Leva o vídeo ao computador.", "De 2 a 3 m em USB 3. Até 5 m em USB 2.",
         "Ligue direto no computador, sem hub. Em USB 2, a imagem é comprimida e costuma ficar em 30 quadros."),
    ], alt=43.5)
    r = tabela(r, "Aparelhos do caminho", [("B", "C"), ("D", "G"), ("H", "I")],
               ["Aparelho", "O que faz", "O que não faz"], [
        ("Conversor de interface", "Troca o tipo de ligação: HDMI para SDI, SDI para HDMI, para fibra.",
         "Não muda o formato. Em geral tem um só sentido. Não passa conteúdo protegido para SDI."),
        ("Scaler", "Muda a resolução e a taxa de quadros para o formato escolhido.",
         "Não cria detalhe. Entre famílias de taxa (50 e 59,94), deixa pequenos trancos. Acrescenta atraso."),
        ("Distribuidor (splitter)", "Copia uma entrada para duas ou mais saídas.",
         "Não muda o formato. Apresenta à origem um só EDID, que pode rebaixar o formato para todas as telas."),
        ("Placa de captura", "Leva HDMI ou SDI para o computador, como uma webcam.",
         "Não aceita HDCP. Modelos simples entregam só 30 quadros ao computador."),
    ], alt=31.5)
    r = tabela(r, "Sintomas e causas prováveis", [("B", "C"), ("D", "G"), ("H", "I")],
               ["O que se vê", "Causa mais provável", "Primeira coisa a fazer"], [
        ("Tela preta ou “sem sinal”", "Entrada errada no destino, formato não aceito ou conversor sem energia.",
         "Conferir a entrada, a fonte do conversor e o formato de saída da origem."),
        ("Pontos brilhantes, linhas, apagões", "Sinal no limite: cabo longo demais, cabo ruim, emenda ou conector frouxo.",
         "Trocar por cabo mais curto ou por outro tipo de ligação."),
        ("Preto só com um conteúdo", "HDCP.", "Testar a origem direto em uma TV. Trocar a fonte do conteúdo."),
        ("Computador só oferece resolução baixa", "O EDID não chegou à origem.",
         "Religar o cabo na origem com o destino já ligado. Testar sem o extensor."),
        ("Movimento aos trancos", "Taxas de quadros de famílias diferentes no caminho.",
         "Igualar a taxa da origem ao padrão do sistema."),
        ("Bordas serrilhadas no movimento", "Sinal entrelaçado mostrado sem desentrelaçar.",
         "Mudar a origem para progressivo, ou ativar o desentrelaçamento no destino."),
        ("Imagem esticada, cortada ou com barras", "Proporção ou resolução da origem diferente da do destino.",
         "Ajustar o computador para 1920 × 1080 e o modo de imagem do destino."),
        ("Imagem perfeita, sem som", "Saída de áudio errada no computador, ou áudio não repassado pelo conversor.",
         "Conferir a saída de som do computador e as chaves de áudio do conversor."),
    ], alt=31.5)
    hint(ws, f"B{r - 1}", "Distâncias e limites são valores práticos, que variam com a qualidade do cabo e com o equipamento das duas pontas.",
         merge=f"B{r - 1}:I{r - 1}", h="left")
    setup(ws, tab, f"B1:I{r - 1}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Sinais e cabos de vídeo — modelo de mapa de sinais"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Sinais e cabos de vídeo — como usar esta planilha",
           "Mapa de sinais, inventário de cabos e conversores e roteiro de diagnóstico, para usar na montagem e no evento.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Formato, ligação, tipo de item e respostas Sim ou Não aceitam só as opções da lista."),
        ("Formatos", "São escritos como nos menus dos equipamentos, com ponto: 1080p59.94. No treinamento, o mesmo formato aparece como 1080p59,94."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Mapa de sinais: escolha o padrão do sistema e confira a tabela de limites de comprimento."),
        ("2", "Aba Mapa de sinais: registre um trecho por linha, da origem para o destino, com o comprimento pelo caminho real do cabo."),
        ("3", "Veja a situação e o que fazer em cada trecho. Corrija o formato na origem, ou mude o tipo de ligação, e atualize a linha."),
        ("4", "Aba Inventário: liste cabos, conversores, extensores e placas de captura, com a data e o resultado do teste."),
        ("5", "Aba Checklist: confira os doze itens na montagem, antes do ensaio."),
        ("6", "Aba Diagnóstico: se faltar imagem, responda às dez perguntas na ordem e registre a ocorrência."),
    ]),
    ("Abas da planilha", [
        ("Mapa de sinais", "Até 20 trechos, com alerta de formato diferente do padrão e de comprimento acima do recomendado."),
        ("Inventário", "Até 30 itens, com situação calculada: OK, Testar, Testar de novo, Falta fonte ou Retirar."),
        ("Diagnóstico", "Roteiro de dez perguntas para o “sem sinal”, com a correção do primeiro “Não”, e registro de ocorrências."),
        ("Referência", "Tabela de formatos, tipos de ligação, aparelhos do caminho e sintomas com as causas prováveis."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Mapa", "O sistema de uma igreja com três câmeras, antes das correções (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Padrão do sistema", "É o formato em que todas as origens devem estar. O modelo compara o texto exato: 1080p59.94 e 1080p60 são formatos diferentes."),
        ("Destino converte?", "Marque Sim quando o destino tem conversor de formato na entrada. O formato diferente passa a contar como Convertido, sem alerta."),
        ("Classe do formato", "HD, 3G, 6G ou 12G, pela quantidade de dados. Os nomes vêm dos níveis do SDI e servem aqui para escolher o limite de qualquer ligação."),
        ("Limites de comprimento", "São pontos de partida conservadores, e não especificação. Troque-os pelos valores do manual dos seus cabos e extensores."),
        ("No limite", "Aviso para trechos acima de 80% do limite. O percentual pode ser trocado no alto da aba."),
        ("Trecho com conversor", "Registre cada lado do conversor em uma linha: da origem ao conversor, e do conversor ao destino."),
        ("Inventário", "O teste vence depois do prazo informado no alto da aba (180 dias no modelo), contado até a data da conferência."),
        ("Exemplos", "Os dados da aba de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_mapa(wb.create_sheet("Mapa de sinais"), BLUE)
aba_inventario(wb.create_sheet("Inventário"), AMBER)
aba_diagnostico(wb.create_sheet("Diagnóstico"), TEAL)
aba_referencia(wb.create_sheet("Referência"), HUE["slate"][0])
checklist(wb.create_sheet("Checklist"), "Checklist dos sinais de vídeo",
          "Confira na montagem, antes do ensaio. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_mapa(wb.create_sheet("Exemplo - Mapa"), MUTED, EX_MAPA)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
