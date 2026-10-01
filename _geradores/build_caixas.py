# -*- coding: utf-8 -*-
"""Gera Caixas-e-Amplificacao-modelo.xlsx: impedância e casamento amplificador × caixa, SPL e cobertura por zona."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
DB1 = '0.0'
DBS = '+0.0;-0.0;0.0'

CHECK = [
    ("Ligação", "A impedância total de cada saída do amplificador foi calculada e não fica abaixo do mínimo do manual."),
    ("Ligação", "A potência do amplificador, na impedância da carga, fica entre uma e duas vezes a potência contínua de cada caixa."),
    ("Ligação", "As caixas passivas estão ligadas com cabo de caixa, de bitola adequada ao comprimento, e não com cabo de instrumento."),
    ("Ligação", "Nas caixas ativas, a chave de entrada está em linha, e o filtro para uso com subwoofer está ligado quando há sub."),
    ("Posição", "As caixas principais estão à frente da linha dos microfones e apontam para a plateia, e não para as paredes."),
    ("Posição", "A corneta fica acima da cabeça de quem está em pé, e o eixo da caixa aponta para as últimas filas."),
    ("Posição", "A primeira fila e o fundo foram ouvidos: se a diferença é grande, há front fill, caixa de delay ou nova posição."),
    ("Posição", "As caixas de delay têm o atraso calculado pela distância e o nível só o bastante para devolver a clareza."),
    ("Nível", "O sistema foi ligado na ordem: fontes, mesa, processador e, por último, amplificadores e caixas ativas."),
    ("Nível", "No trecho mais forte, a luz de clip do amplificador não acende, e a luz de limite só pisca de vez em quando."),
    ("Segurança", "Os tripés estão com a base toda aberta, em piso firme e plano, com o pino de segurança travado e os cabos presos."),
    ("Segurança", "Nenhuma caixa está pendurada sem ferragem própria, e as tomadas têm aterramento e carga compatível."),
]

# Exemplo 1 do treinamento: o amplificador que desligava no meio da formatura
EX_AMP = {
    "head": {"evento": "Formatura no auditório da escola", "local": "Auditório de 150 lugares", "data": date(2026, 12, 12),
             "resp": "Operador de som do evento"},
    "saidas": [  # saída, o que alimenta, qtd, ohms, W contínuos, mínimo do amp, W do amp, em quantos ohms, ganho do amp
        ("Canal A (antes)", "2 caixas principais + 1 retorno, em paralelo", 3, 8, 350, 4, 700, 4, 32),
        ("Canal A (depois)", "2 caixas principais, em paralelo", 2, 8, 350, 4, 700, 4, 32),
        ("Canal B (depois)", "1 retorno de chão", 1, 8, 350, 4, 450, 8, 32),
    ],
    "paralelo": [("Duas caixas de 8 Ω", 8, 8, None, None), ("Três caixas de 8 Ω (o erro do exemplo)", 8, 8, 8, None),
                 ("Uma de 8 Ω com uma de 4 Ω", 8, 4, None, None)],
    "ajustes": [
        ("Antes", "As duas caixas principais e o retorno estavam ligados em cadeia na saída A: três caixas de 8 Ω em paralelo, 2,67 Ω. O amplificador pede 4 Ω ou mais por canal."),
        ("Sintoma", "Depois de uns 20 minutos de uso, a luz de proteção acendia e o som caía. Voltava sozinho quando o amplificador esfriava."),
        ("Depois", "Saída A: só as duas caixas principais (4 Ω). Saída B: o retorno, sozinho (8 Ω), alimentado por um auxiliar da mesa."),
        ("Resultado", "O amplificador trabalhou a formatura inteira sem entrar em proteção, e o retorno ganhou controle de volume separado."),
    ],
}

# Exemplo 2 do treinamento: evento ao ar livre em que o fundo não ouvia
EX_SPL = {
    "head": {"evento": "Culto ao ar livre na praça", "local": "Área de 20 m por 26 m, cerca de 300 pessoas", "data": date(2026, 11, 15),
             "resp": "Equipe de som da igreja"},
    "ref": (90, 10, 6, 10, 343, 10),
    "caixas": [  # nome, sensibilidade, potência, SPL do folheto, ouvinte mais próximo
        ("Caixa ativa de 15 pol. em tripé (PA principal)", None, None, 129, 3),
        ("Caixa ativa de 15 pol. em tripé (delay)", None, None, 129, 4),
    ],
    "zonas": [  # zona, caixa, distância até a caixa, distância do PA principal a esta caixa
        ("Frente (primeira fila)", 1, 3, None),
        ("Meio", 1, 9, None),
        ("Fundo, só com o PA principal", 1, 25, None),
        ("Logo depois das caixas de delay", 2, 4, 14),
        ("Fundo, com as caixas de delay", 2, 11, 14),
    ],
    "ajustes": [
        ("Antes", "Duas caixas ativas em tripé para uma área de 26 m de fundo. A 25 m, o nível máximo contínuo estimado é de 95 dB: faltam 5 dB para o alvo de 100 dB (90 de média mais 10 de folga)."),
        ("Sintoma", "O fundo não entendia a pregação. O operador subiu o master, a luz de limite das caixas ficou acesa e a primeira fila reclamou do volume."),
        ("Depois", "Mais duas caixas iguais a 14 m do palco, apontadas para o fundo, com atraso de 51 ms (41 ms da distância mais 10 ms). O PA principal voltou ao nível da passagem."),
        ("Resultado", "O fundo passou a receber cerca de 102 dB de máximo contínuo, com folga. A luz de limite do PA principal apagou."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def cabecalho(ws, H, campos, col2, hb):
    """Rótulos em B:C e entrada de D até col2, a partir da linha 4."""
    for k, (rot, chave, fmt) in enumerate(campos):
        r = 4 + k
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(chave), merge=f"D{r}:{col2}{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
        if fmt == DATE:
            dv_date(ws, f"D{r}")


def bloco_ajustes(ws, r, ult, linhas):
    band(ws, r, "O que aconteceu no exemplo", ult)
    for rot, texto in linhas:
        r += 1
        put(ws, f"B{r}", rot, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
        put(ws, f"D{r}", texto, merge=f"D{r}:{ult}{r}")
        ws.row_dimensions[r].height = 31.5
    return r


# ================================================================== aba Amplificador e caixas
A_W = {"A": 2, "B": 5, "C": 18, "D": 34, "E": 9, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 11, "L": 12, "M": 12,
       "N": 11, "O": 19, "P": 52, "Q": 12, "R": 28, "S": 2}
A1, NA = 13, 8


def aba_amp(ws, tab, data=None):
    widths(ws, A_W)
    is_ex = data is not None
    title(ws, "Amplificador e caixas — impedância e potência por saída",
          "Exemplo preenchido, para consulta. Use a aba Amplificador e caixas para o seu sistema." if is_ex
          else "Uma linha por saída do amplificador. Preencha as células em amarelo-claro com os dados do manual e da etiqueta das caixas.", "R")
    hb = WHITE if is_ex else INPUT
    cabecalho(ws, data["head"] if is_ex else {}, [("Evento ou sistema", "evento", None), ("Local", "local", None),
                                                 ("Data", "data", DATE), ("Responsável", "resp", None)], "G", hb)
    for k, (rot, v, dica) in enumerate([
        ("Razão mínima amplificador ÷ caixa", 1, "Abaixo dela o amplificador é fraco para a caixa e chega ao clip cedo. Ponto de partida: 1."),
        ("Razão máxima amplificador ÷ caixa", 2, "Acima dela a caixa recebe mais do que a potência de programa. Ponto de partida: 2."),
    ]):
        r = 4 + k
        label(ws, f"I{r}", rot, merge=f"I{r}:L{r}")
        inp(ws, f"M{r}", v, h="center", bg=hb)
        hint(ws, f"N{r}", dica, merge=f"N{r}:R{r}", h="left")
    dv_number(ws, "M4:M5")
    RMIN, RMAX = "$M$4", "$M$5"
    hint(ws, "I6", "A potência por caixa é uma estimativa: considera que o amplificador mantém a mesma tensão de saída.",
         merge="I6:R6", h="left")
    hint(ws, "I7", "Use a potência que o manual declara na impedância igual, ou mais próxima, à impedância total da saída.",
         merge="I7:R7", h="left")

    band(ws, 9, "Saídas do amplificador", "R")
    heads = [("B", "#", INK), ("C", "Saída", INK), ("D", "O que alimenta", INK),
             ("E", "Caixas", BLUE), ("F", "Impedância de cada caixa", BLUE), ("G", "Potência contínua de cada caixa", BLUE),
             ("H", "Impedância mínima do amplif.", AMBER), ("I", "Potência do amplif. por canal", AMBER),
             ("J", "declarada em", AMBER), ("K", "Ganho do amplif.", AMBER),
             ("L", "Impedância total", TEAL), ("M", "Potência por caixa", TEAL), ("N", "Razão amplif. ÷ caixa", TEAL),
             ("O", "Situação", TEAL), ("P", "O que fazer", TEAL), ("Q", "Limiar do limiter", TEAL), ("R", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}10", t, bg=bg)
    ws.row_dimensions[10].height = 43.5
    for col, t in zip("BCDEFGHIJKLMNOPQR", ["", "canal do amplificador", "caixas ligadas em paralelo nesta saída", "quantas",
                                            "Ω", "W (RMS)", "Ω, do manual", "W, do manual", "Ω", "dB, opcional",
                                            "Ω", "W, estimada", "vezes", "", "", "dBu, opcional", ""]):
        hint(ws, f"{col}11", t)
    ws.row_dimensions[11].height = 27
    if is_ex:
        hint(ws, "B12", "A primeira linha mostra a ligação que derrubava o amplificador. As duas seguintes mostram a ligação corrigida. As colunas de L a Q são calculadas.",
             merge="B12:R12", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLMNOPQ", ["Ex.", "Canal A", "2 caixas principais", 2, 8, 350, 4, 700, 4, 32,
                                               4, 350, 1, "OK", "Manter", 4.7]):
            ex(ws, f"{col}12", v, h="left" if col in "CDP" else "center",
               fmt='0.00' if col in "LN" else DB1 if col == "Q" else None)
        ex(ws, "R12", "")
        note(ws, "E10", "Quantas caixas iguais estão ligadas em paralelo nesta saída, contando as que estão em cadeia, uma ligada na outra.")
        note(ws, "F10", "Impedância nominal de cada caixa, em ohms: está na etiqueta, perto do conector. Em geral 8 ou 4.")
        note(ws, "G10", "Potência contínua (a que o fabricante chama de RMS ou AES) de cada caixa.\nNão use a potência de programa nem a de pico.")
        note(ws, "H10", "Menor impedância que o amplificador aceita por canal, segundo o manual. Em geral 4 Ω; alguns aceitam 2 Ω.\nEm ponte (bridge), o mínimo dobra.")
        note(ws, "I10", "Potência por canal que o manual declara, e a impedância em que ela foi declarada (coluna seguinte).\nExemplo: 700 W em 4 Ω.")
        note(ws, "K10", "Ganho de tensão do amplificador, em dB, com os controles de nível no máximo. Está no manual: em geral de 26 a 35 dB.\nSó é usado para calcular o limiar do limiter.")
        note(ws, "L10", "Impedância de cada caixa dividida pelo número de caixas iguais em paralelo.\nPara caixas de impedâncias diferentes, use a calculadora no fim da aba.")
        note(ws, "M10", "Potência que cada caixa recebe com o amplificador no máximo: potência declarada × impedância declarada ÷ impedância da caixa.\nFica em branco quando a saída está abaixo do mínimo: ali o amplificador sai do projeto.")
        note(ws, "Q10", "Nível de saída do processador, em dBu, em que o amplificador entrega à caixa a potência contínua dela.\nÉ ponto de partida para o limiter. Vale com os controles de nível do amplificador no máximo.")
    ws.row_dimensions[12].height = 31.5

    A2 = A1 + NA - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NA):
        r = A1 + k
        d = data["saidas"][k] if is_ex and k < len(data["saidas"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", d[0] if d else None, bg=bg)
        inp(ws, f"D{r}", d[1] if d else None, bg=bg)
        for col, j in zip("EFGHIJK", range(2, 9)):
            inp(ws, f"{col}{r}", d[j] if d else None, h="center", bg=bg)
        E, F, G, Hc, I, J, K, L, M, Nn, O = (f"{c}{r}" for c in "EFGHIJKLMNO")
        calc(ws, L, f'=IF(OR({E}="",{F}="",{E}<=0),"",{F}/{E})', fmt='0.00')
        calc(ws, M, f'=IF(OR({F}="",{I}="",{J}="",{F}<=0),"",IF(AND({Hc}<>"",{L}<>"",{L}<{Hc}-0.01),"",ROUND({I}*{J}/{F},0)))', b=False)
        calc(ws, Nn, f'=IF(OR({M}="",{G}="",{G}<=0),"",{M}/{G})', b=False, fmt='0.00')
        calc(ws, O, f'=IF({L}="","",IF({Hc}="","Incompleto",IF({L}<{Hc}-0.01,"Abaixo do mínimo",IF({Nn}="","Incompleto",'
                    f'IF({Nn}>{RMAX},"Amplificador forte",IF({Nn}<{RMIN},"Amplificador fraco","OK"))))))')
        calc(ws, f"P{r}", f'=IF({O}="","",IF({O}="Abaixo do mínimo","Tire caixas desta saída ou divida entre os canais: o amplificador pede "&{Hc}&" Ω ou mais",'
                          f'IF({O}="Incompleto","Preencha o mínimo e a potência do amplificador e a potência da caixa",'
                          f'IF({O}="Amplificador forte","A caixa pode receber mais do que suporta: use limiter e não deixe o clip acender",'
                          f'IF({O}="Amplificador fraco","Pouca folga: se a luz de clip acender, baixe o nível na mesa","Manter")))))',
             h="left", b=False, sz=9)
        calc(ws, f"Q{r}", f'=IF(OR({F}="",{G}="",{K}="",{F}<=0,{G}<=0),"",ROUND(20*LOG10(({G}*{F})^0.5/0.775)-{K},1))', b=False, fmt=DBS)
        inp(ws, f"R{r}", bg=bg)
        ws.row_dimensions[r].height = 30
    dv_number(ws, f"E{A1}:K{A2}")
    cf_equal(ws, f"O{A1}:O{A2}", [("OK", GREEN), ("Abaixo do mínimo", RED), ("Amplificador forte", YELLOW),
                                  ("Amplificador fraco", YELLOW), ("Incompleto", YELLOW)])

    s = A2 + 2
    band(ws, s, "Resumo automático", "R")
    cont = lambda t: f'COUNTIF(O{A1}:O{A2},"{t}")'
    linhas = [("Saídas registradas", f"=COUNT(E{A1}:E{A2})"),
              ("Saídas na faixa (OK)", f'={cont("OK")}'),
              ("Saídas abaixo da impedância mínima", f'={cont("Abaixo do mínimo")}'),
              ("Saídas com aviso de potência", f'={cont("Amplificador forte")}+{cont("Amplificador fraco")}'),
              ("Saídas com dados incompletos", f'={cont("Incompleto")}')]
    for k, (t, f) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "D", "E", val_merge=f"E{s + k}:F{s + k}")
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(E{s + 1}=0,"Registre as saídas do amplificador",IF(E{s + 3}>0,"Há saída abaixo da impedância mínima: corrija antes de ligar",'
                      f'IF(E{s + 5}>0,"Há saída com dados incompletos",IF(E{s + 4}>0,"Há saída com aviso de potência: veja o que fazer","OK"))))',
         b=False, merge=f"E{a}:P{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:P{a}")

    # calculadora de paralelo com impedâncias diferentes
    c = a + 2
    band(ws, c, "Calculadora: caixas de impedâncias diferentes em paralelo", "R")
    put(ws, f"B{c + 1}", "#", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"C{c + 1}", "Conjunto", f=font(10, True), bg=GRAY, h="center", merge=f"C{c + 1}:D{c + 1}")
    for col, t in zip("EFGHIJKLM", ["Caixa 1 (Ω)", "Caixa 2 (Ω)", "Caixa 3 (Ω)", "Caixa 4 (Ω)", "Total (Ω)",
                                    "Parte da caixa 1", "Parte da caixa 2", "Parte da caixa 3", "Parte da caixa 4"]):
        put(ws, f"{col}{c + 1}", t, f=font(10, True), bg=GRAY, h="center")
    put(ws, f"N{c + 1}", "Como ler", f=font(10, True), bg=GRAY, h="center", merge=f"N{c + 1}:R{c + 1}")
    ws.row_dimensions[c + 1].height = 30
    leitura = ["A parte é a fração da potência da saída que cada caixa recebe.",
               "Em paralelo, a caixa de menor impedância recebe a maior parte.",
               "Compare o total com a impedância mínima do amplificador."]
    for k in range(3):
        r = c + 2 + k
        p = data["paralelo"][k] if is_ex else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", p[0] if p else None, merge=f"C{r}:D{r}", bg=bg)
        for j, col in enumerate("EFGH"):
            inp(ws, f"{col}{r}", p[j + 1] if p else None, h="center", bg=bg)
        soma = "+".join(f'IF(OR({col}{r}="",{col}{r}<=0),0,1/{col}{r})' for col in "EFGH")
        calc(ws, f"I{r}", f'=IF(COUNTIF(E{r}:H{r},">0")=0,"",1/({soma}))', fmt='0.00')
        for col, src in zip("JKLM", "EFGH"):
            calc(ws, f"{col}{r}", f'=IF(OR(I{r}="",{src}{r}="",{src}{r}<=0),"",I{r}/{src}{r})', b=False, fmt='0%')
        hint(ws, f"N{r}", leitura[k], merge=f"N{r}:R{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"E{c + 2}:H{c + 4}")
    end = c + 4

    if is_ex:
        end = bloco_ajustes(ws, end + 2, "R", data["ajustes"])
    ws.freeze_panes = f"E{A1}"
    setup(ws, tab, f"B1:R{end}", fit_height=True)


# ================================================================== aba SPL e cobertura
S_W = {"A": 2, "B": 5, "C": 34, "D": 13, "E": 13, "F": 15, "G": 13, "H": 13, "I": 13, "J": 13, "K": 17, "L": 52,
       "M": 12, "N": 28, "O": 2}
NB, NZ = 6, 12
B1 = 14
Z0 = B1 + NB + 1      # faixa das zonas
Z1 = Z0 + 4           # primeira linha de zona


def aba_spl(ws, tab, data=None):
    widths(ws, S_W)
    is_ex = data is not None
    title(ws, "SPL e cobertura — nível máximo estimado por zona da sala",
          "Exemplo preenchido, para consulta. Use a aba SPL e cobertura para o seu evento." if is_ex
          else "Liste as caixas, depois as zonas da plateia com a distância até a caixa que cobre cada uma. As células cinza são calculadas.", "N")
    hb = WHITE if is_ex else INPUT
    cabecalho(ws, data["head"] if is_ex else {}, [("Evento", "evento", None), ("Local", "local", None),
                                                 ("Data", "data", DATE), ("Responsável", "resp", None)], "F", hb)
    ref = data["ref"] if is_ex else (90, 10, 6, 10, 343, 10)
    refs = [("Nível médio desejado (dB SPL)", "Nível médio que você quer na plateia. Combine com a liderança e meça com um medidor."),
            ("Folga para os picos (dB)", "Espaço acima da média. Ponto de partida: 10 dB. Fala aceita menos; banda com bateria pede mais."),
            ("Desconto do SPL do folheto (dB)", "O SPL máximo do folheto costuma ser de pico. Ponto de partida deste material: 6 dB."),
            ("Diferença aceitável na zona (dB)", "Entre o ouvinte mais próximo da caixa e a zona. Regra prática deste material: 10 dB."),
            ("Velocidade do som (m/s)", "Cerca de 343 m/s a 20 °C. Sobe perto de 0,6 m/s a cada grau."),
            ("Atraso extra do delay (ms)", "Somado ao atraso da distância, para o som parecer vir do palco. De 5 a 15 ms.")]
    for k, (rot, dica) in enumerate(refs):
        r = 4 + k
        label(ws, f"H{r}", rot, merge=f"H{r}:J{r}")
        inp(ws, f"K{r}", ref[k], h="center", bg=hb)
        hint(ws, f"L{r}", dica, merge=f"L{r}:N{r}", h="left")
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, "K4:K9")
    ALVO, FOLGA, DESC, DIF, VEL, EXTRA = (f"$K${4 + k}" for k in range(6))
    hint(ws, "B9", "A conta vale para campo livre (ao ar livre). Em sala, o nível real ao fundo é maior, mas a clareza não.",
         merge="B9:F9", h="left")
    ws.row_dimensions[9].height = 27

    # ---- caixas
    band(ws, B1 - 3, "Caixas do sistema", "N")
    for col, t, bgc in [("B", "#", INK), ("C", "Caixa", INK), ("D", "Sensibilidade", BLUE), ("E", "Potência aplicada", BLUE),
                        ("F", "SPL máximo do folheto", BLUE), ("G", "Ouvinte mais próximo", BLUE),
                        ("H", "SPL contínuo a 1 m", TEAL), ("I", "SPL no ouvinte mais próximo", TEAL)]:
        head(ws, f"{col}{B1 - 2}", t, bg=bgc)
    put(ws, f"J{B1 - 2}", "Origem do valor", f=font(10, True, c=WHITE), bg=TEAL, h="center", merge=f"J{B1 - 2}:K{B1 - 2}")
    put(ws, f"L{B1 - 2}", "Observação", f=font(10, True, c=WHITE), bg=INK, h="center", merge=f"L{B1 - 2}:N{B1 - 2}")
    ws.row_dimensions[B1 - 2].height = 43.5
    for col, t in zip("BCDEFGHI", ["", "nome, tipo e onde fica", "dB (1 W, 1 m)", "W, contínua", "dB a 1 m", "m, em linha reta",
                                   "dB, estimado", "dB, estimado"]):
        hint(ws, f"{col}{B1 - 1}", t)
    hint(ws, f"J{B1 - 1}", "", merge=f"J{B1 - 1}:K{B1 - 1}")
    hint(ws, f"L{B1 - 1}", "Preencha a sensibilidade e a potência, ou só o SPL máximo do folheto.", merge=f"L{B1 - 1}:N{B1 - 1}", h="left")
    ws.row_dimensions[B1 - 1].height = 27
    if not is_ex:
        note(ws, f"D{B1 - 2}", "Sensibilidade da caixa: o nível que ela produz a 1 metro com 1 watt. Está no manual. Em caixas de PA, de 93 a 100 dB.")
        note(ws, f"E{B1 - 2}", "Potência contínua que a caixa recebe: a menor entre a potência contínua da caixa e a que o amplificador entrega a ela.")
        note(ws, f"F{B1 - 2}", "Para caixa ativa: o SPL máximo do manual, a 1 metro. A planilha tira o desconto indicado no alto da aba.\nSe a sensibilidade e a potência estiverem preenchidas, este campo não é usado.")
        note(ws, f"G{B1 - 2}", "Distância, em linha reta, da caixa até o ouvinte mais próximo que ela cobre. Com a caixa no alto, conte a altura.")
    B2 = B1 + NB - 1
    bg = WHITE if is_ex else INPUT
    for k in range(NB):
        r = B1 + k
        d = data["caixas"][k] if is_ex and k < len(data["caixas"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", d[0] if d else None, bg=bg)
        for j, col in enumerate("DEFG"):
            inp(ws, f"{col}{r}", d[j + 1] if d else None, h="center", bg=bg)
        calc(ws, f"H{r}", f'=IF(AND(D{r}<>"",E{r}>0),ROUND(D{r}+10*LOG10(E{r}),1),IF(F{r}<>"",F{r}-{DESC},""))', fmt=DB1)
        calc(ws, f"I{r}", f'=IF(OR(H{r}="",G{r}<=0),"",ROUND(H{r}-20*LOG10(G{r}),1))', b=False, fmt=DB1)
        calc(ws, f"J{r}", f'=IF(H{r}="","",IF(AND(D{r}<>"",E{r}>0),"Sensibilidade e potência","Folheto, com desconto"))',
             b=False, merge=f"J{r}:K{r}", sz=9)
        inp(ws, f"L{r}", merge=f"L{r}:N{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"D{B1}:G{B2}")
    SPL1, PROX = f"$H${B1}:$H${B2}", f"$G${B1}:$G${B2}"

    # ---- zonas
    band(ws, Z0, "Zonas da plateia", "N")
    for col, t, bgc in [("B", "#", INK), ("C", "Zona", INK), ("D", "Caixa que cobre", AMBER), ("E", "Distância até a caixa", AMBER),
                        ("F", "Distância do PA principal a esta caixa", AMBER), ("G", "Perda pela distância", TEAL),
                        ("H", "SPL máximo na zona", TEAL), ("I", "Margem sobre o alvo", TEAL),
                        ("J", "Diferença para o mais próximo", TEAL), ("K", "Situação", TEAL), ("L", "O que fazer", TEAL),
                        ("M", "Atraso sugerido", TEAL), ("N", "Observação", INK)]:
        head(ws, f"{col}{Z0 + 1}", t, bg=bgc)
    ws.row_dimensions[Z0 + 1].height = 43.5
    for col, t in zip("BCDEFGHIJKLMN", ["", "parte da plateia", "nº da lista acima", "m, em linha reta", "m, só para delay e fill",
                                        "dB", "dB, contínuo", "dB", "dB", "", "", "ms", ""]):
        hint(ws, f"{col}{Z0 + 2}", t)
    ws.row_dimensions[Z0 + 2].height = 27
    if is_ex:
        hint(ws, f"B{Z0 + 3}", "A terceira linha mostra o fundo antes da correção, só com o PA principal. As duas últimas mostram as caixas de delay.",
             merge=f"B{Z0 + 3}:N{Z0 + 3}", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Meio da plateia", 1, 9, "", 19.1, 103.9, 3.9, -9.5, "OK", "Manter", ""]):
            ex(ws, f"{col}{Z0 + 3}", v, h="left" if col in "CL" else "center",
               fmt=DB1 if col in "GH" else DBS if col in "IJ" else None)
        ex(ws, f"N{Z0 + 3}", "")
        note(ws, f"D{Z0 + 1}", "Número da caixa que cobre esta zona, na lista Caixas do sistema.")
        note(ws, f"F{Z0 + 1}", "Só para caixas de delay e front fill: a distância entre o PA principal e esta caixa.\nA planilha calcula o atraso a programar nela.")
        note(ws, f"H{Z0 + 1}", "Nível contínuo máximo estimado de uma caixa, em campo livre: SPL contínuo a 1 m menos a perda pela distância.\nNão soma a segunda caixa nem a reverberação da sala.")
        note(ws, f"I{Z0 + 1}", "SPL máximo na zona menos o alvo (nível médio desejado mais a folga para os picos). Negativo: falta sistema.")
        note(ws, f"J{Z0 + 1}", "Quantos dB esta zona recebe a menos do que o ouvinte mais próximo da mesma caixa.")
    ws.row_dimensions[Z0 + 3].height = 24
    Z2 = Z1 + NZ - 1
    for k in range(NZ):
        r = Z1 + k
        d = data["zonas"][k] if is_ex and k < len(data["zonas"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", d[0] if d else None, bg=bg)
        for j, col in enumerate("DEF"):
            inp(ws, f"{col}{r}", d[j + 1] if d else None, h="center", bg=bg)
        D, E, F, G, Hc, I, J, K = (f"{c}{r}" for c in "DEFGHIJK")
        s1 = f"INDEX({SPL1},{D}*1)"
        px = f"INDEX({PROX},{D}*1)"
        calc(ws, G, f'=IF(OR({D}="",{E}<=0),"",ROUND(20*LOG10({E}),1))', b=False, fmt=DB1)
        calc(ws, Hc, f'=IF({G}="","",IF({s1}="","",ROUND({s1}-20*LOG10({E}),1)))', fmt=DB1)
        calc(ws, I, f'=IF({Hc}="","",ROUND({Hc}-({ALVO}+{FOLGA}),1))', b=False, fmt=DBS)
        calc(ws, J, f'=IF({G}="","",IF({px}<=0,"",ROUND(-20*LOG10({E}/{px}),1)))', b=False, fmt=DBS)
        calc(ws, K, f'=IF({G}="","",IF({Hc}="","Incompleto",IF({I}<0,"Falta nível",IF({J}="","OK",IF({J}<-{DIF},"Desigual","OK")))))')
        calc(ws, f"L{r}", f'=IF({K}="","",IF({K}="Incompleto","Preencha os dados da caixa na lista acima",'
                          f'IF({K}="Falta nível","Faltam "&ROUND(ABS({I}),0)&" dB: aproxime a caixa, use caixa de delay ou um sistema maior",'
                          f'IF({K}="Desigual","Frente bem mais forte: suba e incline a caixa, ou divida a zona com outra caixa","Manter"))))',
             h="left", b=False, sz=9)
        calc(ws, f"M{r}", f'=IF({F}>0,ROUND({F}/{VEL}*1000+{EXTRA},0),"")', b=False)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 30
    dv_list(ws, f"D{Z1}:D{Z2}", [str(n) for n in range(1, NB + 1)], "Número da caixa na lista Caixas do sistema.")
    dv_number(ws, f"E{Z1}:F{Z2}")
    cf_equal(ws, f"K{Z1}:K{Z2}", [("OK", GREEN), ("Falta nível", RED), ("Desigual", YELLOW), ("Incompleto", YELLOW)])

    s = Z2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(K{Z1}:K{Z2},"{t}")'
    linhas = [("Zonas registradas", f"=COUNT(E{Z1}:E{Z2})", None),
              ("Zonas atendidas (OK)", f'={cont("OK")}', None),
              ("Zonas com falta de nível", f'={cont("Falta nível")}', None),
              ("Zonas com nível desigual", f'={cont("Desigual")}', None),
              ("Menor SPL máximo entre as zonas (dB)", f'=IF(COUNT(H{Z1}:H{Z2})=0,"",MIN(H{Z1}:H{Z2}))', DB1),
              ("Alvo: nível médio mais a folga (dB)", f'=IF(COUNT({ALVO},{FOLGA})<2,"",{ALVO}+{FOLGA})', DB1)]
    for k, (t, f, fm) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "D", "E", fmt=fm)
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:D{a}", h="right")
    calc(ws, f"E{a}", f'=IF(COUNT($K$4:$K$9)<6,"Informe os seis valores de referência",IF(E{s + 1}=0,"Registre as zonas da plateia",'
                      f'IF({cont("Incompleto")}>0,"Há zona sem os dados da caixa",IF(E{s + 3}>0,"Há zona com falta de nível: o sistema não alcança o alvo",'
                      f'IF(E{s + 4}>0,"Há zona com nível desigual: reveja a posição das caixas","OK")))))',
         b=False, merge=f"E{a}:L{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"E{a}:L{a}")
    hint(ws, f"B{a + 1}", "Estimativa em campo livre, para uma caixa por zona. Não substitui a medição com medidor de nível sonoro nem o ouvido em cada zona.",
         merge=f"B{a + 1}:N{a + 1}", h="left")
    end = a + 1
    if is_ex:
        end = bloco_ajustes(ws, end + 2, "N", data["ajustes"])
    ws.freeze_panes = f"D{B1}"
    setup(ws, tab, f"B1:N{end}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Caixas, amplificação e PA — modelo de impedância, potência e cobertura"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Caixas, amplificação e PA — como usar esta planilha",
           "Conferência da ligação entre amplificador e caixas e estimativa do nível que o sistema entrega em cada zona da plateia.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "A coluna Caixa que cobre aceita só o número de uma caixa da lista."),
        ("Números", "Digite só o número, sem unidade: 8, e não “8 ohms”; 350, e não “350 W”."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Tenha em mãos o manual do amplificador e a etiqueta das caixas (impedância e potência contínua)."),
        ("2", "Aba Amplificador e caixas: uma linha por saída, com o número de caixas em paralelo e os dados do amplificador."),
        ("3", "Veja a impedância total e a situação. Saída abaixo do mínimo se corrige antes de ligar o sistema."),
        ("4", "Aba SPL e cobertura: confira os valores de referência e liste as caixas do sistema."),
        ("5", "Meça com trena a distância de cada caixa ao ouvinte mais próximo e a cada zona da plateia, e preencha as zonas."),
        ("6", "Veja a margem, a diferença entre frente e fundo e, para as caixas de delay, o atraso sugerido."),
        ("7", "Aba Checklist: confira os doze itens antes de abrir as portas."),
    ]),
    ("Abas da planilha", [
        ("Amplificador e caixas", "Até 8 saídas: impedância total em paralelo, potência por caixa, situação e limiar de partida do limiter. Calculadora de paralelo."),
        ("SPL e cobertura", "Até 6 caixas e 12 zonas: SPL máximo contínuo estimado por zona, margem sobre o alvo e atraso das caixas de delay."),
        ("Referência", "Potências, impedância em paralelo, decibéis por potência e por distância, cabo de caixa e ordem de ligar."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Amplificador", "O amplificador que desligava na formatura: três caixas na mesma saída (exemplo 1 do treinamento)."),
        ("Exemplo - Cobertura", "O evento ao ar livre em que o fundo não ouvia, com as caixas de delay (exemplo 2 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Impedância em paralelo", "Caixas iguais: impedância de uma dividida pelo número de caixas. Caixas diferentes: use a calculadora no fim da aba."),
        ("Potência por caixa", "Estimada pela tensão do amplificador: potência declarada × impedância declarada ÷ impedância da caixa. É exata quando a carga é igual à declarada."),
        ("Razão amplificador ÷ caixa", "OK entre 1 e 2 vezes a potência contínua da caixa. É regra prática comum, e não norma: os limites podem ser trocados no alto da aba."),
        ("Limiar do limiter", "Tensão que dá à caixa a potência contínua dela, em dBu, menos o ganho do amplificador. É ponto de partida: siga o manual da caixa e do processador."),
        ("SPL a 1 metro", "Sensibilidade mais 10 × log da potência. Em caixa ativa, o SPL máximo do folheto menos o desconto indicado."),
        ("SPL na zona", "Campo livre: cai 6 dB cada vez que a distância dobra. Não considera a soma de caixas, a reverberação nem a posição fora do eixo."),
        ("Atraso do delay", "Distância entre o PA principal e a caixa de delay, dividida pela velocidade do som, mais o atraso extra."),
        ("Exemplos", "Os dados das abas de exemplo são ilustrativos, criados para este material."),
    ]),
])

aba_amp(wb.create_sheet("Amplificador e caixas"), AMBER)
aba_spl(wb.create_sheet("SPL e cobertura"), BLUE)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 36, "C": 28, "D": 28, "E": 62, "F": 2})
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


r = tabela(4, "Potência: o que cada número quer dizer", ["Número", "Relação com a contínua", "Exemplo", "Para que serve"], [
    ("Contínua (RMS, AES)", "é a base", "350 W", "O que a caixa suporta por horas, com sinal de teste. É o número para comparar com o amplificador."),
    ("Programa (música)", "2 vezes a contínua", "700 W", "Convenção para sinal musical, que tem média bem abaixo dos picos. Teto para a potência do amplificador."),
    ("Pico", "4 vezes a contínua", "1.400 W", "O que a caixa aguenta por um instante. Não serve para escolher amplificador."),
    ("PMPO", "sem relação definida", "–", "Número de propaganda, sem método de medida comum. Ignore."),
])
r = tabela(r, "Caixas em paralelo na mesma saída", ["Ligação", "Impedância total", "Amplificador com mínimo de 4 Ω", "Observação"], [
    ("1 caixa de 8 Ω", "8 Ω", "aceita", "Carga leve. O amplificador entrega menos potência do que em 4 Ω."),
    ("2 caixas de 8 Ω", "4 Ω", "aceita, no limite", "A ligação mais comum. Cada caixa recebe metade da potência do canal."),
    ("3 caixas de 8 Ω", "2,67 Ω", "não aceita", "Abaixo do mínimo: aquece, entra em proteção ou queima."),
    ("4 caixas de 8 Ω", "2 Ω", "não aceita", "Só para amplificador que declara potência em 2 Ω."),
    ("1 caixa de 4 Ω", "4 Ω", "aceita, no limite", "Não ligue outra caixa na mesma saída."),
    ("2 caixas de 4 Ω", "2 Ω", "não aceita", "Só para amplificador que declara potência em 2 Ω."),
    ("1 de 8 Ω com 1 de 4 Ω", "2,67 Ω", "não aceita", "A caixa de 4 Ω recebe o dobro da potência da de 8 Ω."),
    ("Modo ponte (bridge)", "–", "mínimo de 8 Ω", "Em ponte, a impedância mínima dobra. Confira o manual."),
])
r = tabela(r, "Decibéis por potência e por distância", ["Mudança", "Efeito no nível", "Exemplo", "O que lembrar"], [
    ("Dobrar a potência", "+3 dB", "400 W → 800 W", "Diferença pequena no ouvido. Custa o dobro de amplificador."),
    ("Dez vezes a potência", "+10 dB", "100 W → 1.000 W", "É o que o ouvido percebe como o dobro do volume."),
    ("Dobrar a distância", "−6 dB em campo livre", "4 m → 8 m", "Em sala, a queda é menor ao fundo, por causa da reverberação."),
    ("Dez vezes a distância", "−20 dB em campo livre", "2 m → 20 m", "Aproximar a caixa do ouvinte rende mais do que trocar de amplificador."),
    ("Sensibilidade 3 dB maior", "+3 dB com a mesma potência", "95 dB → 98 dB", "Vale tanto quanto dobrar a potência do amplificador."),
    ("Dobrar o número de caixas", "+3 a +6 dB", "1 sub → 2 subs juntos", "Perto de +6 dB nos graves, com as caixas juntas. Nos agudos, a soma é irregular."),
])
r = tabela(r, "Cabo de caixa: seção do condutor (cobre)", ["Comprimento do cabo", "Carga de 8 Ω", "Carga de 4 Ω", "Observação"], [
    ("Até 10 m", "1,5 mm²", "2,5 mm²", "Use cabo de caixa de dois condutores. Cabo de instrumento ou de microfone não serve."),
    ("De 10 a 25 m", "2,5 mm²", "4 mm²", "Cabo fino e longo perde potência e controle dos graves."),
    ("Acima de 25 m", "4 mm²", "leve o amplificador para perto", "Para trechos longos, prefira caixa ativa ou amplificador perto da caixa, com cabo de sinal longo."),
])
r = tabela(r, "Ordem de ligar e de desligar", ["Etapa", "Ao ligar", "Ao desligar", "Por quê"], [
    ("Fontes e periféricos", "1º", "4º", "Computador, receptores sem fio, tocadores."),
    ("Mesa de som", "2º", "3º", "Com o master fechado. Espere a mesa terminar de iniciar."),
    ("Processador", "3º", "2º", "Fica entre a mesa e os amplificadores."),
    ("Amplificadores e caixas ativas", "por último", "primeiro", "O estalo de qualquer aparelho ligado antes deles não chega aos alto-falantes."),
])
r = tabela(r, "Posição e cobertura", ["Item", "Ponto de partida", "Faixa típica", "O que conferir"], [
    ("Ângulo de cobertura", "90° × 50°", "60° a 110° na horizontal", "Horizontal × vertical, do manual. Caixa deitada troca os dois ângulos."),
    ("Altura da corneta", "acima das cabeças", "2 m ou mais", "Acima de quem está em pé na frente da caixa."),
    ("Direção do eixo", "para as últimas filas", "–", "Não aponte para a parede do fundo, para o teto nem para as paredes laterais."),
    ("Linha das caixas", "à frente dos microfones", "–", "Microfone dentro da cobertura da caixa reduz a margem antes da microfonia."),
    ("Corte entre sub e caixa", "100 Hz", "80 a 120 Hz", "Ligue o filtro passa-altas da caixa quando houver subwoofer."),
    ("Atraso da caixa de delay", "2,9 ms por metro, mais 10 ms", "extra de 5 a 15 ms", "Distância entre o PA principal e a caixa de delay."),
], alt=24)
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com o modelo e com a sala. As seções de cabo e a razão amplificador ÷ caixa são regra prática, e não norma.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplos
checklist(wb.create_sheet("Checklist"), "Checklist do PA",
          "Confira depois de montar o sistema e antes de abrir as portas. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=TEAL)
aba_amp(wb.create_sheet("Exemplo - Amplificador"), MUTED, EX_AMP)
aba_spl(wb.create_sheet("Exemplo - Cobertura"), MUTED, EX_SPL)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
