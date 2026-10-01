# -*- coding: utf-8 -*-
"""Gera Projecao-e-Midia-modelo.xlsx: calculadora de projeção, lista de mídias, conferência de slides,
referência, checklist e exemplo."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

PURPLE_H, PURPLE_T = HUE["purple"]
AMBER, AMBER_T = HUE["amber"]
TEAL = HUE["teal"][0]
SIM_NAO = ["Sim", "Não"]
PROPS = ["16x9", "16x10", "4x3"]

CHECK = [
    ("Preparação", "Todas as mídias do roteiro estão na lista, na ordem do evento, e chegaram no prazo."),
    ("Preparação", "Cada arquivo foi conferido (formato, resolução, proporção) e copiado para o disco do computador do evento."),
    ("Preparação", "As letras foram revisadas e estão na ordem combinada com quem conduz a música."),
    ("Preparação", "Os slides têm poucas linhas, letra grande, fundo com contraste e texto dentro da área segura."),
    ("Tela", "O projetor ou o painel foi ligado com antecedência, sem aviso de filtro, de lâmpada ou de temperatura."),
    ("Tela", "A imagem enche a tela, reta, em foco nos quatro cantos e na proporção certa."),
    ("Computador", "O computador está com as telas estendidas, na resolução nativa da saída, com fundo preto."),
    ("Tela", "Com a luz do evento acesa, um slide de letra e um de texto foram lidos do último banco."),
    ("Mídia", "Cada vídeo foi assistido até o fim, na tela e nas caixas, com o operador de som."),
    ("Saídas", "A saída de palco e a saída da transmissão, se houver, mostram o que devem."),
    ("Computador", "Notificações, proteção de tela, suspensão e atualizações estão desligadas, e o computador está na tomada."),
    ("Operação", "O operador sabe os atalhos de tela preta e de limpar o texto, e o plano B está à mão."),
]

# Exemplo 1 do treinamento: a solução adotada na igreja de 200 lugares
EX_PROJ = {
    "evento": "Culto de domingo", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
    "operador": "Operador de projeção da escala",
    "uso": "Texto", "dist": 14, "fila": 5, "prop": "16x9", "larg": 4.2,
    "lm": 5000, "trmin": 1.4, "trmax": 2.2, "dl": 7.5,
    "amb": 25, "linhas": 1080, "med": 13,
    "pitch": 3.9, "lw": 4, "lh": 2.5, "diag": None,
    "historico": [  # situação, tela, projetor, luz, resultado
        ("Como estava", "Tela de 2,1 m (altura 1,18 m)", "3.000 lm", "80 lux", "Último banco a 11,9 alturas; contraste 13 : 1. Imagem pequena demais."),
        ("Só a tela maior", "Tela de 4,2 m (altura 2,36 m)", "3.000 lm", "80 lux", "Tamanho certo; contraste 4 : 1. Imagem lavada."),
        ("Solução adotada", "Tela de 4,2 m (altura 2,36 m)", "5.000 lm a 7,5 m", "25 lux", "Último banco a 5,9 alturas; contraste 17 : 1; letra de 13 cm."),
        ("Comparação", "Painel de LED P3,9 de 4 × 2,5 m", "–", "Não depende", "Atende à primeira fila e à luz da sala; ficou de fora pelo custo."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="left"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


# ================================================================== aba Projeção
P_W = {"A": 2, "B": 46, "C": 14, "D": 9, "E": 18, "F": 64, "G": 2}


def aba_proj(ws, tab, data=None):
    widths(ws, P_W)
    is_ex = data is not None
    D = data or {}
    bg = WHITE if is_ex else INPUT
    title(ws, "Projeção — tamanho, distância, brilho e letra",
          "Exemplo preenchido, para consulta. Use a aba Projeção para a sua sala." if is_ex
          else "Preencha as células em amarelo-claro com as medidas da sala. As células cinza são calculadas.", "F")
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None)]:
        label(ws, f"B{r}", rot)
        inp(ws, f"C{r}", D.get(k), merge=f"C{r}:F{r}", bg=bg, fmt=fmt)
        ws.row_dimensions[r].height = 19.5
    dv_date(ws, "C6")

    R = {}
    r = 8

    def sec(text):
        nonlocal r
        r += 1
        band(ws, r, text, "F")
        r += 1
        for col, t in zip("BCDEF", ["Item", "Valor", "Unidade", "Situação", "O que fazer / observação"]):
            put(ws, f"{col}{r}", t, f=font(9, True, c=MUTED), bg=GRAY, h="center")
        ws.row_dimensions[r].height = 18

    def ref(k):
        return f"$C${R[k]}"

    def line(k, rot, unit, kind, val=None, fmt=None, dica=None, sit=None, acao=None, dv=None):
        """kind: 'in' (entrada), 'rule' (entrada com valor inicial), 'calc'."""
        nonlocal r
        r += 1
        R[k] = r
        label(ws, f"B{r}", rot)
        if kind == "calc":
            calc(ws, f"C{r}", val, fmt=fmt)
        elif kind == "rule":
            inp(ws, f"C{r}", val, h="center", fmt=fmt, bg=bg)
        else:
            inp(ws, f"C{r}", D.get(k), h="center", fmt=fmt, bg=bg)
        put(ws, f"D{r}", unit, f=font(9, c=MUTED), bg=GRAY, h="center")
        if sit:
            calc(ws, f"E{r}", sit(r))
            if acao:
                calc(ws, f"F{r}", acao(r), h="left", b=False, sz=9)
            else:
                hint(ws, f"F{r}", dica or "")
            cf_status(ws, f"E{r}", f"$E${r}")
        else:
            hint(ws, f"E{r}", dica or "", merge=f"E{r}:F{r}")
        if dv == "num":
            dv_number(ws, f"C{r}")
        elif dv:
            dv_list(ws, f"C{r}", dv[0], dv[1])
        ws.row_dimensions[r].height = 21.75 if len(dica or "") < 95 else 31.5
        return r

    # ---------------------------------------------------------- regras
    sec("Regras práticas deste material (pode trocar, se o projeto da sua sala usar outras)")
    line("fT", "Último espectador, no máximo, a quantas alturas (texto)", "× H", "rule", 6, dv="num",
         dica="Para ler letras, Bíblia e avisos. Regra prática, e não norma.")
    line("fV", "Último espectador, no máximo, a quantas alturas (só vídeo)", "× H", "rule", 8, dv="num",
         dica="Quando a tela mostra só vídeo, ou letras muito grandes.")
    line("fP", "Primeira fila, no mínimo, a quantas alturas", "× H", "rule", 2, dv="num",
         dica="Mais perto, a pessoa precisa virar a cabeça para ver a imagem inteira.")
    line("cmin", "Contraste mínimo para texto", ": 1", "rule", 15, dv="num",
         dica="Abaixo de 7 : 1 a imagem lava; de 7 a 15, só letra grande; acima de 50, bom para vídeo escuro.")
    line("dmin", "Letra mínima: distância dividida por", "÷", "rule", 200, dv="num",
         dica="Altura da letra maiúscula na tela. Letra fina ou fundo movimentado pedem mais.")
    line("dconf", "Letra confortável: distância dividida por", "÷", "rule", 150, dv="num")
    line("aprov", "Aproveitamento dos lúmens da ficha", "%", "rule", 0.8, fmt="0%", dv="num",
         dica="Desconta lâmpada usada e ajustes. Use 60% em modo econômico com lâmpada velha; 90% em laser novo.")
    fT, fV, fP, cmin, dmin, dconf, aprov = (ref(k) for k in ("fT", "fV", "fP", "cmin", "dmin", "dconf", "aprov"))

    # ---------------------------------------------------------- 1. sala e tela
    sec("1. Sala e tamanho da imagem")
    line("uso", "Uso principal da tela", "", "in", dv=(["Texto", "Só vídeo"], "Texto ou Só vídeo."),
         dica="Texto: letras, Bíblia, avisos, slides. Só vídeo: quando não há leitura.")
    line("dist", "Distância da tela ao último espectador que precisa ler", "m", "in", fmt="0.0", dv="num")
    line("fila", "Distância da tela à primeira fila", "m", "in", fmt="0.0", dv="num")
    line("prop", "Proporção da imagem", "", "in", dv=(PROPS, "16x9, 16x10 ou 4x3."),
         dica="Em branco, o modelo usa 16x9.")
    line("larg", "Largura da imagem (projetada ou planejada)", "m", "in", fmt="0.00", dv="num")
    U, DI, FI, PR, LA = (f"C{R[k]}" for k in ("uso", "dist", "fila", "prop", "larg"))
    RATIO = f'IF({PR}="16x10",0.625,IF({PR}="4x3",0.75,0.5625))'
    FATOR = f'IF({U}="Só vídeo",{fV},{fT})'
    line("alt", "Altura da imagem", "m", "calc", f'=IF({LA}="","",{LA}*{RATIO})', fmt="0.00")
    AL = f"C{R['alt']}"
    line("area", "Área da imagem", "m²", "calc", f'=IF({AL}="","",{LA}*{AL})', fmt="0.00")
    AR = f"C{R['area']}"
    line("altmin", "Altura mínima da imagem para esta distância", "m", "calc", f'=IF({DI}="","",{DI}/{FATOR})', fmt="0.00")
    AM = f"C{R['altmin']}"
    line("largmin", "Largura mínima da imagem", "m", "calc", f'=IF({AM}="","",{AM}/{RATIO})', fmt="0.00",
         dica="Na proporção escolhida.")
    LM = f"C{R['largmin']}"
    line("emalt", "Último espectador, em alturas da imagem", "× H", "calc",
         f'=IF(OR({DI}="",{AL}=""),"",{DI}/{AL})', fmt="0.0",
         sit=lambda r: f'=IF(C{r}="","",IF(C{r}<={FATOR},"OK","Imagem pequena"))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK","Manter",'
                        f'"Aumente a imagem até a largura mínima, ou ponha telas de apoio para o fundo da sala"))')
    line("filaalt", "Primeira fila, em alturas da imagem", "× H", "calc",
         f'=IF(OR({FI}="",{AL}=""),"",{FI}/{AL})', fmt="0.0",
         sit=lambda r: f'=IF(C{r}="","",IF(C{r}>={fP},"OK","Perto demais"))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK","Manter","Recue a primeira fila, ou reduza a imagem se ela passa do mínimo"))')

    # ---------------------------------------------------------- 2. projetor
    sec("2. Projetor e distância (throw ratio)")
    line("lm", "Lúmens da ficha (ANSI ou ISO)", "lm", "in", fmt="#,##0", dv="num",
         dica="Não use números de “lúmens de LED” ou de “lux” de aparelhos sem método padronizado.")
    line("trmin", "Throw ratio mínimo (zoom aberto)", "", "in", fmt="0.00", dv="num",
         dica="Na ficha aparece como 1,4 a 2,2, por exemplo. Lente fixa: um só valor.")
    line("trmax", "Throw ratio máximo (zoom fechado)", "", "in", fmt="0.00", dv="num",
         dica="Em branco para lente fixa.")
    line("dl", "Distância da lente à tela", "m", "in", fmt="0.00", dv="num",
         dica="A medida atual, ou o ponto de fixação que você pretende usar.")
    TL, TA, DL = (f"C{R[k]}" for k in ("trmin", "trmax", "dl"))
    TMAX = f'IF({TA}="",{TL},{TA})'
    line("nde", "Distância que a lente pede para esta largura: de", "m", "calc",
         f'=IF(OR({LA}="",{TL}=""),"",{LA}*{TL})', fmt="0.00", dica="Largura da imagem × throw ratio mínimo.")
    line("nate", "até", "m", "calc", f'=IF(C{R["nde"]}="","",{LA}*{TMAX})', fmt="0.00",
         dica="Largura da imagem × throw ratio máximo.")
    line("wde", "Largura possível a esta distância: de", "m", "calc",
         f'=IF(OR({DL}="",{TL}=""),"",{DL}/{TMAX})', fmt="0.00", dica="Com o zoom fechado.")
    line("wate", "até", "m", "calc", f'=IF(C{R["wde"]}="","",{DL}/{TL})', fmt="0.00", dica="Com o zoom aberto.")
    ND, NA = f"C{R['nde']}", f"C{R['nate']}"
    line("possit", "Posição do projetor", "", "calc", f'=IF(OR({DL}="",{ND}=""),"",{DL})', fmt="0.00",
         sit=lambda r: f'=IF(C{r}="","",IF({DL}<{ND}-0.005,"Perto demais",IF({DL}>{NA}+0.005,"Longe demais","OK")))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK","Dentro da faixa do zoom. Ajuste fino no zoom e no foco",'
                        f'IF(E{r}="Perto demais","A imagem não chega à largura: afaste o projetor ou use lente de throw ratio menor",'
                        f'"A imagem passa da largura: aproxime o projetor ou use lente de throw ratio maior")))')

    # ---------------------------------------------------------- 3. luz
    sec("3. Luz ambiente e contraste")
    line("amb", "Luz ambiente na tela, com o projetor tampado", "lux", "in", fmt="0", dv="num",
         dica="Com a luz do evento e do palco acesas. Anote o maior valor entre o centro e os cantos.")
    AB = f"C{R['amb']}"
    line("plux", "Luz do projetor na tela", "lux", "calc",
         f'=IF(OR(C{R["lm"]}="",{AR}=""),"",C{R["lm"]}*{aprov}/{AR})', fmt="0",
         dica="Lúmens × aproveitamento ÷ área da imagem.")
    PL = f"C{R['plux']}"
    line("contr", "Contraste na sala", ": 1", "calc",
         f'=IF(OR({PL}="",{AB}="",{AB}<=0),"",({PL}+{AB})/{AB})', fmt="0.0",
         sit=lambda r: f'=IF(C{r}="","",IF(C{r}>={cmin},"OK","Contraste baixo"))',
         acao=lambda r: f'=IF(E{r}="","",IF(C{r}<7,"Imagem lavada. Tire luz da tela, reduza a imagem ou use mais lúmens",'
                        f'IF(E{r}="OK",IF(C{r}>=50,"Bom também para vídeo com cenas escuras","Bom para letras e texto"),'
                        f'"No limite: só letra grande e em negrito. Tire luz da tela ou use mais lúmens")))')
    line("lmnec", "Lúmens da ficha necessários para o contraste mínimo", "lm", "calc",
         f'=IF(OR({AB}="",{AR}=""),"",({cmin}-1)*{AB}*{AR}/{aprov})', fmt="#,##0",
         dica="Com esta luz ambiente e esta área de imagem.")
    line("ambmax", "Luz ambiente máxima na tela para este projetor", "lux", "calc",
         f'=IF({PL}="","",{PL}/({cmin}-1))', fmt="0",
         dica="Acima disso, o contraste cai abaixo do mínimo.")

    # ---------------------------------------------------------- 4. letra
    sec("4. Tamanho da letra")
    line("linhas", "Linhas da saída de vídeo", "px", "in", fmt="0", dv="num",
         dica="1080 para Full HD, 720 para HD, 768 ou 800 em alguns projetores.")
    line("lmin", "Altura mínima da letra maiúscula na tela", "cm", "calc",
         f'=IF({DI}="","",{DI}/{dmin}*100)', fmt="0.0")
    line("lconf", "Altura confortável da letra maiúscula", "cm", "calc",
         f'=IF({DI}="","",{DI}/{dconf}*100)', fmt="0.0")
    line("lpx", "Letra mínima em pixels da saída", "px", "calc",
         f'=IF(OR(C{R["lmin"]}="",{AL}="",C{R["linhas"]}=""),"",ROUND(C{R["lmin"]}/100/{AL}*C{R["linhas"]},0))', fmt="0",
         dica="Altura da maiúscula. A fonte precisa ser cerca de 1,4 vez maior, porque a maiúscula ocupa perto de 70% do tamanho nominal.")
    line("med", "Altura da letra maiúscula medida na tela", "cm", "in", fmt="0.0", dv="num",
         sit=lambda r: f'=IF(OR(C{r}="",C{R["lmin"]}=""),"",IF(C{r}>=C{R["lconf"]},"OK",IF(C{r}>=C{R["lmin"]},"No limite","Letra pequena")))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK","Manter",IF(E{r}="No limite","Lê-se com esforço: menos linhas e letra maior",'
                        f'"Reduza as linhas por slide e aumente a letra")))')

    # ---------------------------------------------------------- 5. LED e TV
    sec("5. Painel de LED ou TV (se usar)")
    line("pitch", "Pixel pitch do painel", "mm", "in", fmt="0.0", dv="num")
    line("lw", "Largura do painel", "m", "in", fmt="0.00", dv="num")
    line("lh", "Altura do painel", "m", "in", fmt="0.00", dv="num")
    PI, LW, LH = (f"C{R[k]}" for k in ("pitch", "lw", "lh"))
    line("res", "Resolução aproximada do painel", "pontos", "calc",
         f'=IF(OR({PI}="",{LW}="",{LH}=""),"",ROUND({LW}*1000/{PI},0)&" × "&ROUND({LH}*1000/{PI},0))',
         dica="A real depende do tamanho dos gabinetes. Compare com a saída do computador.")
    line("dled", "Distância mínima de visão (1 m por mm de pitch)", "m", "calc", f'=IF({PI}="","",{PI}*1)', fmt="0.0",
         sit=lambda r: f'=IF(OR(C{r}="",{FI}=""),"",IF({FI}>=C{r},"OK","Perto demais"))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK",IF({FI}>=2*C{r},"Imagem lisa desde a primeira fila",'
                        f'"Aceitável. De perto, a grade dos pontos ainda aparece um pouco"),'
                        f'"A primeira fila vê os pontos: recue a fila ou use pitch menor"))')
    line("ledalt", "Último espectador, em alturas do painel", "× H", "calc",
         f'=IF(OR({LH}="",{DI}=""),"",{DI}/{LH})', fmt="0.0",
         sit=lambda r: f'=IF(C{r}="","",IF(C{r}<={FATOR},"OK","Imagem pequena"))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK","Manter","Painel baixo para a distância: aumente a altura ou ponha telas de apoio"))')
    line("diag", "Diagonal da TV", "pol.", "in", fmt="0", dv="num", dica="Usa a proporção informada na parte 1.")
    DG = f"C{R['diag']}"
    line("tvw", "Largura da TV", "m", "calc",
         f'=IF({DG}="","",{DG}*0.0254*IF({PR}="16x10",0.848,IF({PR}="4x3",0.8,0.8716)))', fmt="0.00")
    line("tvh", "Altura da TV", "m", "calc",
         f'=IF({DG}="","",{DG}*0.0254*IF({PR}="16x10",0.53,IF({PR}="4x3",0.6,0.4903)))', fmt="0.00")
    line("tvalt", "Último espectador, em alturas da TV", "× H", "calc",
         f'=IF(OR(C{R["tvh"]}="",{DI}=""),"",{DI}/C{R["tvh"]})', fmt="0.0",
         sit=lambda r: f'=IF(C{r}="","",IF(C{r}<={FATOR},"OK","Imagem pequena"))',
         acao=lambda r: f'=IF(E{r}="","",IF(E{r}="OK","Manter","TV pequena para esta distância: use TV maior, projetor ou TVs de apoio"))')

    # ---------------------------------------------------------- resumo
    r += 2
    band(ws, r, "Resumo automático", "F")
    S = [f"E{R[k]}" for k in ("emalt", "filaalt", "possit", "contr", "med", "dled", "ledalt", "tvalt")]
    filled = "+".join(f'IF({c}<>"",1,0)' for c in S)
    pend = "+".join(f'IF(AND({c}<>"",{c}<>"OK"),1,0)' for c in S)
    summary(ws, r + 1, "Pontos conferidos", f"={filled}", "B", "C")
    summary(ws, r + 2, "Pontos a corrigir", f"={pend}", "B", "C")
    label(ws, f"B{r + 3}", "Aviso", h="right")
    calc(ws, f"C{r + 3}", f'=IF({DI}="","Informe a distância do último espectador",IF(C{r + 1}=0,'
                          f'"Preencha a imagem, o projetor, o painel ou a TV",IF(C{r + 2}>0,"Há pontos a corrigir: veja a coluna Situação","OK")))',
         b=False, merge=f"C{r + 3}:F{r + 3}", h="left")
    ws.row_dimensions[r + 3].height = 21.75
    cf_ok(ws, f"C{r + 3}:F{r + 3}")
    end = r + 3
    hint(ws, f"B{end + 1}", "Os fatores do alto da aba são regra prática deste material, e não norma. As contas de brilho supõem tela branca fosca.",
         merge=f"B{end + 1}:F{end + 1}")
    end += 1

    if is_ex:
        band(ws, end + 2, "Como a sala chegou a esta solução (exemplo 1 do treinamento)", "F")
        r = end + 3
        put(ws, f"B{r}", "Situação", f=font(10, True), bg=GRAY, h="center")
        put(ws, f"C{r}", "Tela", f=font(10, True), bg=GRAY, h="center", merge=f"C{r}:D{r}")
        put(ws, f"E{r}", "Projetor", f=font(10, True), bg=GRAY, h="center")
        put(ws, f"F{r}", "Luz na tela e resultado", f=font(10, True), bg=GRAY, h="center")
        for sit, tela, proj, luz, res in data["historico"]:
            r += 1
            put(ws, f"B{r}", sit, f=font(10, True), bg=GRAY)
            put(ws, f"C{r}", tela, merge=f"C{r}:D{r}")
            put(ws, f"E{r}", proj, h="center")
            put(ws, f"F{r}", f"{luz}. {res}")
            ws.row_dimensions[r].height = 31.5
        end = r
    ws.freeze_panes = "C9"
    setup(ws, tab, f"B1:F{end}", landscape=False)
    return R


# ================================================================== aba Mídias
M_W = {"A": 2, "B": 5, "C": 18, "D": 30, "E": 11, "F": 16, "G": 12, "H": 10, "I": 12, "J": 9, "K": 9, "L": 10,
       "M": 7, "N": 7, "O": 10, "P": 11, "Q": 22, "R": 30, "S": 2}
M1, MN = 14, 30


def aba_midias(ws, tab):
    widths(ws, M_W)
    title(ws, "Mídias do evento — lista em ordem, com duração e situação",
          "Liste na ordem do roteiro. A situação mostra o que falta fazer em cada mídia.", "R")
    for r, rot in [(4, "Evento"), (5, "Data do evento"), (6, "Antecedência para receber (dias)"), (7, "Operador")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", 2 if r == 6 else None, merge=f"D{r}:F{r}", fmt=DATE if r == 5 else None, h="left")
        ws.row_dimensions[r].height = 19.5
    dv_date(ws, "D5")
    dv_number(ws, "D6")
    label(ws, "H4", "Data limite para receber", merge="H4:J4")
    calc(ws, "K4", '=IF(OR(D5="",D6=""),"",D5-D6)', fmt=DATE, merge="K4:M4")
    hint(ws, "H5", "Mídia que chega depois da data limite fica marcada como atrasada: dá tempo de converter e testar, mas aperta.",
         merge="H5:R5")
    hint(ws, "H6", "Duração em minutos e segundos. O acumulado mostra quanto tempo de mídia já passou até o fim de cada item.",
         merge="H6:R6")

    band(ws, 9, "Lista de mídias", "R")
    heads = [("B", "#", INK), ("C", "Momento (deixa)", INK), ("D", "Mídia", INK), ("E", "Tipo", INK),
             ("F", "Quem envia", INK), ("G", "Recebida em", AMBER), ("H", "No prazo", PURPLE_H),
             ("I", "Formato conferido", AMBER), ("J", "No disco", AMBER), ("K", "Tem áudio", AMBER),
             ("L", "Testada", AMBER), ("M", "Min", AMBER), ("N", "Seg", AMBER), ("O", "Duração", PURPLE_H),
             ("P", "Acumulado", PURPLE_H), ("Q", "Situação", PURPLE_H), ("R", "Observação", INK)]
    for col, t, c in heads:
        head(ws, f"{col}10", t, bg=c)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLMNOPQR", ["", "no roteiro", "nome do arquivo ou título", "lista", "", "data",
                                            "", "Sim / Convertido / Não", "Sim / Não", "Sim / Não", "na tela e no som",
                                            "", "", "", "", "", ""]):
        hint(ws, f"{col}11", t, h="center")
    ws.row_dimensions[11].height = 24
    for col, v in zip("BCDEFGHIJKLMNOPQ", ["Ex.", "Abertura", "Vídeo de abertura", "Vídeo", "Comunicação",
                                           "02/10/2026", "Sim", "Convertido", "Sim", "Sim", "Não", 2, 40, "2:40", "2:40",
                                           "Testar na tela e no som"]):
        ex(ws, f"{col}12", v, h="left" if col in "CDQ" else "center")
    ex(ws, "R12", "")
    hint(ws, "B13", "Situação: Falta receber → Conferir o formato → Copiar para o disco → Testar → OK. Letras e Bíblia montadas no programa também passam por aqui.",
         merge="B13:R13")
    note(ws, "I10", "Abriu no computador do evento, com formato, resolução e proporção certos?\nConvertido: precisou converter, e a cópia convertida está pronta.")
    note(ws, "L10", "Assistida até o fim na tela do evento e, se tiver áudio, nas caixas da sala, com o operador de som.")
    M2 = M1 + MN - 1
    for k in range(MN):
        r = M1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDEFGIJKLMNR":
            inp(ws, f"{col}{r}", h="left" if col in "CDFR" else "center", fmt=DATE if col == "G" else None)
        calc(ws, f"H{r}", f'=IF(OR(G{r}="",$K$4=""),"",IF(G{r}<=$K$4,"Sim","Atrasada"))')
        calc(ws, f"O{r}", f'=IF(AND(M{r}="",N{r}=""),"",(M{r}*60+N{r})/86400)', fmt="[m]:ss", b=False)
        calc(ws, f"P{r}", f'=IF(O{r}="","",SUM(O${M1}:O{r}))', fmt="[m]:ss", b=False)
        calc(ws, f"Q{r}", f'=IF(D{r}="","",IF(G{r}="","Falta receber",IF(OR(I{r}="",I{r}="Não"),"Conferir o formato",'
                          f'IF(J{r}<>"Sim","Copiar para o disco",IF(L{r}<>"Sim",IF(K{r}="Sim","Testar na tela e no som","Testar na tela"),"OK")))))')
        ws.row_dimensions[r].height = 19.5
    dv_list(ws, f"E{M1}:E{M2}", ["Letra", "Bíblia", "Vídeo", "Imagem", "Slides", "Aviso", "Áudio"],
            "Letra, Bíblia, Vídeo, Imagem, Slides, Aviso ou Áudio.")
    dv_date(ws, f"G{M1}:G{M2}")
    dv_list(ws, f"I{M1}:I{M2}", ["Sim", "Convertido", "Não"], "Sim, Convertido ou Não.")
    for col, txt in (("J", "A mídia está copiada no disco do computador do evento?"), ("K", "A mídia tem som?"),
                     ("L", "Foi assistida até o fim, na tela e no som da sala?")):
        dv_list(ws, f"{col}{M1}:{col}{M2}", SIM_NAO, txt)
    dv_number(ws, f"M{M1}:N{M2}")
    cf_status(ws, f"Q{M1}:Q{M2}", f"Q{M1}")
    cf_equal(ws, f"H{M1}:H{M2}", [("Sim", GREEN), ("Atrasada", YELLOW)])

    s = M2 + 2
    band(ws, s, "Resumo automático", "R")
    linhas = [
        ("Mídias listadas", f"=COUNTA(D{M1}:D{M2})", None),
        ("Duração total das mídias", f"=SUM(O{M1}:O{M2})", "[h]:mm:ss"),
        ("Prontas (OK)", f'=COUNTIF(Q{M1}:Q{M2},"OK")', None),
        ("Pendentes", f"=F{s + 1}-F{s + 3}", None),
        ("Chegaram atrasadas", f'=COUNTIF(H{M1}:H{M2},"Atrasada")', None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "E", "F", fmt=fmt, val_merge=f"F{s + k}:G{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:E{a}", h="right")
    calc(ws, f"F{a}", f'=IF(F{s + 1}=0,"Liste as mídias do evento",IF(F{s + 4}>0,"Há mídia pendente: veja a coluna Situação","OK"))',
         b=False, merge=f"F{a}:R{a}", h="left")
    cf_ok(ws, f"F{a}:R{a}")
    ws.freeze_panes = f"E{M1}"
    setup(ws, tab, f"B1:R{a}", fit_height=True)


# ================================================================== aba Slides
S_W = {"A": 2, "B": 5, "C": 30, "D": 11, "E": 9, "F": 10, "G": 10, "H": 11, "I": 18, "J": 9, "K": 10, "L": 10,
       "M": 9, "N": 18, "O": 44, "P": 26, "Q": 2}
S1, SN = 14, 20


def aba_slides(ws, tab, R):
    widths(ws, S_W)
    title(ws, "Slides — conferência de letras, Bíblia, avisos e apresentações",
          "Uma linha por música ou apresentação. A situação aponta o primeiro problema a corrigir.", "P")
    regras = [
        (4, "Linhas por slide em música, no máximo", 2, None),
        (5, "Linhas por slide em Bíblia, aviso e palestra", 4, None),
        (6, "Caracteres na maior linha, no máximo", 35, None),
    ]
    for r, rot, v, _ in regras:
        label(ws, f"B{r}", rot, merge=f"B{r}:F{r}")
        inp(ws, f"G{r}", v, h="center")
        ws.row_dimensions[r].height = 19.5
    dv_number(ws, "G4:G6")
    label(ws, "B7", "Letra mínima na tela (cm), vinda da aba Projeção", merge="B7:F7")
    lm = R["lmin"]
    calc(ws, "G7", f"=IF('Projeção'!C{lm}=\"\",\"\",'Projeção'!C{lm})", fmt="0.0")
    ws.row_dimensions[7].height = 19.5
    hint(ws, "I4", "Regras práticas deste material. Duas linhas por slide de música é o ponto de partida; Bíblia e avisos aceitam um pouco mais.",
         merge="I4:P5")
    hint(ws, "I6", "A altura da letra se mede na tela, com uma régua, em uma letra maiúscula.", merge="I6:P6")
    hint(ws, "I7", "Preencha a aba Projeção para a planilha conferir a altura da letra.", merge="I7:P7")

    band(ws, 9, "Conferência dos slides", "P")
    heads = [("B", "#", INK), ("C", "Item", INK), ("D", "Tipo", INK), ("E", "Slides", AMBER),
             ("F", "Máx. de linhas", AMBER), ("G", "Maior linha", AMBER), ("H", "Altura da letra", AMBER),
             ("I", "Fundo", AMBER), ("J", "Área segura", AMBER), ("K", "Revisado", AMBER), ("L", "Ordem combinada", AMBER),
             ("M", "Limite de linhas", PURPLE_H), ("N", "Situação", PURPLE_H), ("O", "O que corrigir", PURPLE_H),
             ("P", "Observação", INK)]
    for col, t, c in heads:
        head(ws, f"{col}10", t, bg=c)
    ws.row_dimensions[10].height = 30
    for col, t in zip("BCDEFGHIJKLMNOP", ["", "música ou apresentação", "lista", "nº", "por slide", "caracteres", "cm, na tela",
                                          "lista", "Sim / Não", "Sim / Não", "Sim / Não", "", "", "", ""]):
        hint(ws, f"{col}11", t, h="center")
    ws.row_dimensions[11].height = 24
    for col, v in zip("BCDEFGHIJKLMNO", ["Ex.", "Música de abertura", "Música", 18, 3, 31, 9, "Escuro liso", "Sim", "Sim",
                                         "Sim", 2, "Linhas demais", "Divida em mais slides: até 2 linhas por slide"]):
        ex(ws, f"{col}12", v, h="left" if col in "CO" else "center")
    ex(ws, "P12", "")
    hint(ws, "B13", "Ordem combinada: a sequência de estrofes e refrões confere com quem conduz a música. Em palestra, com quem fala.",
         merge="B13:P13")
    S2 = S1 + SN - 1
    for k in range(SN):
        r = S1 + k
        num(ws, f"B{r}", k + 1)
        for col in "CDEFGHIJKLP":
            inp(ws, f"{col}{r}", h="left" if col in "CP" else "center")
        calc(ws, f"M{r}", f'=IF(C{r}="","",IF(D{r}="Música",$G$4,$G$5))')
        calc(ws, f"N{r}", f'=IF(C{r}="","",IF(F{r}="","Preencha as linhas",IF(F{r}>M{r},"Linhas demais",'
                          f'IF(AND(G{r}<>"",G{r}>$G$6),"Linha longa",IF(AND(H{r}<>"",$G$7<>"",H{r}<$G$7),"Letra pequena",'
                          f'IF(OR(I{r}="Claro",I{r}="Movimentado"),"Pouco contraste",IF(J{r}="Não","Fora da área segura",'
                          f'IF(K{r}<>"Sim","Revisar",IF(L{r}="Não","Conferir a ordem","OK")))))))))')
        calc(ws, f"O{r}", f'=IF(N{r}="","",IF(N{r}="OK","Manter",IF(N{r}="Preencha as linhas","Conte as linhas do slide mais cheio",'
                          f'IF(N{r}="Linhas demais","Divida em mais slides: até "&M{r}&" linhas por slide",'
                          f'IF(N{r}="Linha longa","Quebre a linha onde a frase respira",'
                          f'IF(N{r}="Letra pequena","Aumente a letra até a altura mínima, com menos linhas",'
                          f'IF(N{r}="Pouco contraste","Fundo escuro e liso, ou sombra e contorno na letra",'
                          f'IF(N{r}="Fora da área segura","Afaste o texto das bordas",'
                          f'IF(N{r}="Revisar","Leia tudo em voz alta uma vez","Confira a sequência com quem conduz")))))))))',
             h="left", b=False, sz=9)
        ws.row_dimensions[r].height = 19.5
    dv_list(ws, f"D{S1}:D{S2}", ["Música", "Bíblia", "Aviso", "Palestra"], "Música, Bíblia, Aviso ou Palestra.")
    dv_number(ws, f"E{S1}:H{S2}")
    dv_list(ws, f"I{S1}:I{S2}", ["Escuro liso", "Imagem com sombra", "Claro", "Movimentado"],
            "Escuro liso, Imagem com sombra, Claro ou Movimentado.")
    for col in "JKL":
        dv_list(ws, f"{col}{S1}:{col}{S2}", SIM_NAO, "Sim ou Não.")
    cf_status(ws, f"N{S1}:N{S2}", f"N{S1}")

    s = S2 + 2
    band(ws, s, "Resumo automático", "P")
    summary(ws, s + 1, "Itens listados", f"=COUNTA(C{S1}:C{S2})", "D", "E")
    summary(ws, s + 2, "Prontos (OK)", f'=COUNTIF(N{S1}:N{S2},"OK")', "D", "E")
    summary(ws, s + 3, "A corrigir", f"=E{s + 1}-E{s + 2}", "D", "E")
    label(ws, f"B{s + 4}", "Aviso", merge=f"B{s + 4}:D{s + 4}", h="right")
    calc(ws, f"E{s + 4}", f'=IF(E{s + 1}=0,"Liste as músicas e as apresentações",IF(E{s + 3}>0,"Há slides a corrigir: veja a coluna Situação","OK"))',
         b=False, merge=f"E{s + 4}:P{s + 4}", h="left")
    cf_ok(ws, f"E{s + 4}:P{s + 4}")
    ws.freeze_panes = f"D{S1}"
    setup(ws, tab, f"B1:P{s + 4}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Projeção, telão e mídia — modelo de apoio"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Projeção, telão e mídia — como usar esta planilha",
           "Calculadora de projeção, lista de mídias do evento e conferência de slides.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Uso, proporção, tipo, fundo e as colunas Sim / Não aceitam só as opções da lista."),
        ("Números", "Digite só o número, sem a unidade: 14, e não “14 m”. Use a vírgula decimal do seu computador."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Projeção: meça a sala (último espectador, primeira fila) e a imagem. Veja se o tamanho está na faixa."),
        ("2", "Aba Projeção: informe os lúmens e o throw ratio do projetor e a distância da lente à tela."),
        ("3", "Aba Projeção: meça os lux na tela, com a luz do evento e o projetor tampado. Veja o contraste."),
        ("4", "Aba Projeção: meça a altura de uma letra maiúscula projetada. Se usa painel de LED ou TV, preencha a parte 5."),
        ("5", "Aba Mídias: liste as mídias na ordem do roteiro e acompanhe até todas ficarem OK."),
        ("6", "Aba Slides: confira cada música e cada apresentação antes do ensaio."),
        ("7", "Aba Checklist: confira os doze itens antes de abrir as portas."),
    ]),
    ("Abas da planilha", [
        ("Projeção", "Tamanho mínimo da imagem, distância pelo throw ratio, brilho, contraste, letra, painel de LED e TV."),
        ("Mídias", "Até 30 mídias em ordem, com prazo, conferência, duração, tempo acumulado e situação."),
        ("Slides", "Até 20 músicas ou apresentações, com linhas, tamanho da linha e da letra, fundo e revisão."),
        ("Referência", "Contraste, tipos de lente, medidas de TVs, pixel pitch, formatos de mídia e regras de slide."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Projeção", "A igreja de 200 lugares do exemplo 1 do treinamento, com a solução adotada."),
    ]),
    ("Regras e premissas do modelo", [
        ("Tamanho", "Para texto, o último espectador fica a até 6 alturas da imagem; só vídeo, até 8; a primeira fila a pelo menos 2. Regra prática."),
        ("Brilho", "Lux na tela = lúmens × aproveitamento ÷ área da imagem. Supõe tela branca fosca; tela cinza ou de alto ganho muda a conta."),
        ("Contraste", "(luz do projetor + luz ambiente) ÷ luz ambiente, na tela. O mínimo de 15 : 1 para texto é referência de projeto, e não lei."),
        ("Throw ratio", "Distância da lente à tela ÷ largura da imagem. Com zoom, vale a faixa entre o mínimo e o máximo da ficha."),
        ("Letra", "Altura mínima da maiúscula = distância ÷ 200; confortável = distância ÷ 150. Letra fina ou fundo movimentado pedem mais."),
        ("Painel de LED", "Distância mínima de cerca de 1 m por mm de pitch; imagem lisa a partir do dobro. A resolução é aproximada."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

R = aba_proj(wb.create_sheet("Projeção"), PURPLE_H)
aba_midias(wb.create_sheet("Mídias"), AMBER)
aba_slides(wb.create_sheet("Slides"), TEAL, R)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 24, "D": 24, "E": 58, "F": 2})
title(ws, "Referência — valores típicos", "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")


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


def br(x, d=2):
    return f"{x:.{d}f}".replace(".", ",")


tvs = []
for pol in (43, 50, 55, 65, 75, 85, 98):
    w, h = pol * 0.0254 * 0.8716, pol * 0.0254 * 0.4903
    tvs.append((f"{pol} polegadas", f"{br(w)} × {br(h)} m", f"até {br(6 * h, 1)} m", f"Só vídeo: até {br(8 * h, 1)} m."))

r = tabela(4, "Contraste na sala", ["Contraste", "Como aparece", "Serve para", "O que fazer"], [
    ("Abaixo de 7 : 1", "lavada", "nada que precise de leitura", "Tire luz da tela, reduza a imagem ou use mais lúmens."),
    ("De 7 a 15 : 1", "dá para ver, com esforço", "letra grande e em negrito", "Melhore a luz da sala antes do evento."),
    ("De 15 a 50 : 1", "texto firme, cores presentes", "letras, Bíblia, avisos, slides", "Meta mínima deste material."),
    ("Acima de 50 : 1", "preto de verdade", "vídeo com cenas escuras", "–"),
])
r = tabela(r, "Tipos de lente (throw ratio)", ["Tipo", "Throw ratio", "Distância para 3 m de largura", "Onde se usa"], [
    ("Tiro ultracurto", "abaixo de 0,4", "menos de 1,2 m", "Colado na parede, acima ou abaixo da tela. Exige tela muito plana."),
    ("Tiro curto", "0,4 a 1,0", "1,2 a 3 m", "Salas rasas, palco em que ninguém pode passar na frente do feixe."),
    ("Padrão", "1,2 a 2,0", "3,6 a 6 m", "A maioria dos projetores de sala e de auditório."),
    ("Tiro longo", "acima de 2,0", "mais de 6 m", "Projetor no fundo da sala ou na cabine."),
])
r = tabela(r, "Medidas de TVs 16:9", ["Diagonal", "Largura × altura", "Último espectador (texto)", "Observação"], tvs, alt=19.5)
r = tabela(r, "Painel de LED: pixel pitch", ["Pixel pitch", "Pontos por metro", "Distância mínima de visão", "Uso típico"], [
    ("1,9 a 2,6 mm", "380 a 520", "2 a 3 m", "Salas pequenas, estúdio, fundo que aparece em câmera de perto."),
    ("2,9 a 3,9 mm", "256 a 345", "3 a 4 m", "O mais comum em igrejas, auditórios e locação para ambiente interno."),
    ("4,8 a 6 mm", "165 a 210", "5 a 6 m", "Salas grandes e uso externo, com público distante."),
    ("8 mm ou mais", "125 ou menos", "8 m ou mais", "Fachadas e grandes eventos externos."),
])
r = tabela(r, "Formatos de mídia (ponto de partida)", ["Mídia", "Formato seguro", "Detalhe", "O que evitar"], [
    ("Vídeo", "MP4, H.264", "1920 × 1080, 30 qps, 8 a 15 Mbps", "4K para tela Full HD, H.265 em computador antigo, taxa de quadros variável de celular."),
    ("Áudio do vídeo", "AAC estéreo", "48 kHz", "Áudio em mais de dois canais (5.1); vídeos com volumes muito diferentes."),
    ("Imagem", "JPG (foto), PNG (logo)", "1920 × 1080", "Foto de 12 megapixels em cada slide; imagem pequena ampliada."),
    ("Música e trilha", "WAV, MP3 ou AAC", "256 kbps ou mais", "Áudio baixado de vídeo da internet em baixa qualidade."),
    ("Slides de terceiros", "Original em 16:9 + PDF", "fontes incorporadas", "Fonte que só existe no computador de quem fez; vídeo embutido sem o arquivo."),
])
r = tabela(r, "Regras de slide", ["Item", "Ponto de partida", "Limite", "Por quê"], [
    ("Linhas por slide", "2 em música", "4 em Bíblia e avisos", "Poucas linhas permitem letra grande."),
    ("Tamanho da linha", "até 30 a 35 caracteres", "–", "Linha longa obriga a letra a diminuir."),
    ("Fonte", "sem serifa, traço médio ou negrito", "maiúsculas e minúsculas", "Traço fino e letra enfeitada somem de longe."),
    ("Cor e fundo", "letra clara sobre fundo escuro e liso", "sombra ou contorno sobre imagem", "Fundo claro ofusca; fundo movimentado disputa com a letra."),
    ("Margem", "5% de cada borda", "ou mais", "A borda pode ser cortada pela tela ou pelo ajuste de imagem."),
    ("Troca da letra", "na última palavra do slide", "–", "O público precisa ver a linha antes de cantar."),
])
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com o equipamento, a sala e o fabricante. Os fatores de distância e de letra são regra prática deste material.",
     merge=f"B{r - 1}:E{r - 1}")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist da projeção e da mídia",
          "Confira antes de abrir as portas. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["blue"][0])
aba_proj(wb.create_sheet("Exemplo - Projeção"), MUTED, EX_PROJ)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
