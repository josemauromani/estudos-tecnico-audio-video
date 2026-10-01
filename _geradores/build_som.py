# -*- coding: utf-8 -*-
"""Gera Som-e-Sinal-modelo.xlsx: calculadora de dB, medição de nível na sala, faixas por fonte, referência e checklist."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

BLUE, BLUE_T = HUE["blue"]
AMBER, AMBER_T = HUE["amber"]
TEAL, TEAL_T = HUE["teal"]
SIT_CF = [("OK", GREEN), ("Forte", YELLOW), ("Fraco", YELLOW), ("Ruído", YELLOW), ("Acima do limite", RED)]
SG = '+0;-0;0'
SG1 = '+0.0;-0.0;0.0'
N = 16  # pontos na aba de medição

CHECK = [
    ("Medição", "O medidor de nível sonoro está em ponderação A e resposta lenta, e os dois ajustes estão anotados."),
    ("Medição", "O ruído de fundo da sala foi medido com o sistema mudo e o ar-condicionado ligado."),
    ("Medição", "O nível foi medido na mesa e em pelo menos cinco outros pontos da plateia, com o mesmo sinal."),
    ("Medição", "A diferença entre a primeira fileira e a mesa de som é conhecida e está anotada."),
    ("Medição", "Nenhum ponto da plateia passa do limite de volume combinado para o evento."),
    ("Sistema", "As caixas estão acima da cabeça do público e apontadas para ele, sem obstáculo na frente das cornetas."),
    ("Sistema", "A polaridade das caixas foi conferida: com as duas tocando juntas, o grave cresce no meio da sala."),
    ("Sistema", "As fontes captadas por dois sinais (dois microfones, ou microfone e direct box) foram ouvidas com e sem a inversão de polaridade."),
    ("Operação", "O operador sabe a escala de cada medidor do sistema: dBFS ou VU na mesa, dB SPL no medidor da sala."),
    ("Operação", "O nível na mesa é conferido com o medidor nos trechos mais fortes do evento, e o valor é anotado."),
    ("Audição", "Quem fica perto das caixas ou da bateria tem protetor auricular à mão."),
    ("Audição", "Ao final, a equipe conferiu se alguém ficou com zumbido ou com o ouvido abafado, e registrou."),
]

# Exemplo 1 do treinamento: medição da sala de uma igreja de 200 lugares, antes e depois de elevar as caixas
EX_MED = {
    "head": {"evento": "Culto de domingo, medição antes do louvor", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador de som da escala", "medidor": "Medidor de nível sonoro portátil", "pond": "A",
             "resp": "Lenta", "sinal": "Música gravada no nível do louvor, em repetição"},
    "ref": (90, 6, 48, 5),
    "pontos": [  # ponto, distância, antes, depois
        ("Primeira fileira, no centro", 3, 96, 89),
        ("Primeira fileira, em frente à caixa", 2.5, 98, 90),
        ("Meio da sala (fileira 4)", 6, 91, 88),
        ("Fileira 8", 9, 88, 86),
        ("Mesa de som (referência)", 12, 86, 84),
        ("Última fileira, no centro", 16, 84, 83),
        ("Última fileira, na lateral", 17, 83, 82),
    ],
    "ajustes": [
        ("Caixas", "Elevadas em tripés, acima da cabeça do público, e apontadas para o terço final da sala."),
        ("Master", "2 dB abaixo do que estava."),
        ("Resultado", "A diferença entre o ponto mais forte e o mais fraco caiu de 15 para 8 dB. Nenhum ponto ficou acima de 90 dB(A)."),
    ],
}

FAIXAS = [  # nome, de, até, o que fica ali, como se percebe
    ("Subgrave", 20, 60, "Fundo do bumbo, notas mais graves do baixo e do teclado.", "Mais sentido no corpo do que ouvido."),
    ("Grave", 60, 250, "Bumbo, baixo, fundamentais da voz masculina, zumbido da rede (60 Hz).", "Peso. Em excesso, som pesado e sem contorno."),
    ("Médio-grave", 250, 500, "Corpo das vozes, do violão, da caixa da bateria e do teclado.", "Calor. Em excesso, som embolado."),
    ("Médio", 500, 2000, "Vogais e primeiros harmônicos de quase todos os instrumentos.", "As fontes disputam espaço. Em excesso, som de telefone."),
    ("Médio-agudo", 2000, 6000, "Consoantes, ataque da palheta e do batedor do bumbo.", "Clareza da palavra. O ouvido é mais sensível aqui."),
    ("Agudo", 6000, 20000, "Pratos, brilho das cordas, ar da voz, sons de “s”.", "Brilho. Em excesso, som áspero, que cansa."),
]

FONTES = [  # fonte, fundamental mín, máx, vai até, observação
    ("Voz falada de homem", 85, 180, 8000, "A palavra é entendida pelas consoantes, de 2 a 8 kHz."),
    ("Voz falada de mulher", 165, 255, 8000, "A palavra é entendida pelas consoantes, de 2 a 8 kHz."),
    ("Voz cantada, do baixo ao soprano", 80, 1050, 10000, "Cada cantor ocupa só uma parte dessa extensão."),
    ("Violão", 82, 1000, 10000, "A corda mais grave é um mi de 82 Hz."),
    ("Guitarra", 82, 1200, 6000, "O alto-falante do amplificador quase não passa de 6 kHz."),
    ("Baixo elétrico de 4 cordas", 41, 400, 5000, "Ataque e harmônicos até perto de 5 kHz."),
    ("Baixo elétrico de 5 cordas", 31, 400, 5000, "A corda mais grave é um si de 31 Hz."),
    ("Teclado e piano", 27, 4200, 10000, "A fonte de maior extensão. Sintetizadores podem descer mais."),
    ("Bumbo", 50, 100, 5000, "O ataque do batedor fica de 2 a 5 kHz."),
    ("Caixa da bateria", 150, 250, 10000, "A esteira fica de 3 a 10 kHz."),
    ("Pratos e chimbal", None, None, 16000, "Sem nota definida. Energia de 300 Hz a 16 kHz, mais forte acima de 6 kHz."),
]


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


def faixa_de(cell):
    """Nome da faixa em que cai a frequência da célula (fórmula sem o sinal de igual)."""
    return (f'IF({cell}<20,"Abaixo de 20 Hz",IF({cell}<60,"Subgrave",IF({cell}<250,"Grave",IF({cell}<500,"Médio-grave",'
            f'IF({cell}<2000,"Médio",IF({cell}<6000,"Médio-agudo",IF({cell}<=20000,"Agudo","Acima de 20 kHz")))))))')


# ================================================================== aba Calculadora de dB
def aba_calc(ws):
    widths(ws, {"A": 2, "B": 40, "C": 13, "D": 9, "E": 36, "F": 13, "G": 46, "H": 2})
    title(ws, "Calculadora de dB — as contas do dia a dia",
          "Troque os valores em amarelo-claro. Os resultados em cinza são calculados. Os valores iniciais são os dos exemplos do treinamento.", "G")
    label(ws, "B4", "Velocidade do som (m/s)")
    inp(ws, "C4", 343, h="center")
    put(ws, "D4", "m/s", bg=GRAY, h="center")
    hint(ws, "E4", "A 20 °C. Sobe cerca de 0,6 m/s por grau. Usada nos blocos 6 e 7.", merge="E4:G4", h="left")
    dv_number(ws, "C4")
    C = "$C$4"
    r = 6

    def bloco(titulo, entradas, saidas, dica=None):
        """entradas: [(rótulo, valor, unidade)]; saidas: [(rótulo, fórmula, unidade ou None, formato, texto?)].
        Devolve as referências das entradas."""
        nonlocal r
        band(ws, r, titulo, "G")
        r += 1
        head(ws, f"B{r}", "Entrada", bg=AMBER)
        head(ws, f"C{r}", "Valor", bg=AMBER)
        head(ws, f"D{r}", "Unidade", bg=AMBER)
        head(ws, f"E{r}", "Resultado", bg=TEAL)
        head(ws, f"F{r}", "Valor", bg=TEAL)
        head(ws, f"G{r}", "Unidade ou leitura", bg=TEAL)
        r0 = r + 1
        refs = [f"C{r0 + k}" for k in range(len(entradas))]
        outs = [f"F{r0 + k}" for k in range(len(saidas))]
        linhas = max(len(entradas), len(saidas))
        for k in range(linhas):
            rr = r0 + k
            if k < len(entradas):
                rot, v, un = entradas[k]
                label(ws, f"B{rr}", rot)
                inp(ws, f"C{rr}", v, h="center")
                put(ws, f"D{rr}", un, bg=GRAY, h="center")
            else:
                for col in "BCD":
                    put(ws, f"{col}{rr}", None, bg=WHITE, box=False)
            if k < len(saidas):
                rot, fml, un, fmt, texto = saidas[k]
                fml = fml.format(*refs, C=C, o=outs)
                label(ws, f"E{rr}", rot)
                if texto:
                    calc(ws, f"F{rr}", fml, h="left", b=False, merge=f"F{rr}:G{rr}")
                else:
                    calc(ws, f"F{rr}", fml, fmt=fmt)
                    hint(ws, f"G{rr}", un or "", h="left")
            else:
                for col in "EFG":
                    put(ws, f"{col}{rr}", None, bg=WHITE, box=False)
            ws.row_dimensions[rr].height = 21.75
        r = r0 + linhas
        if dica:
            hint(ws, f"B{r}", dica, merge=f"B{r}:G{r}", h="left")
            ws.row_dimensions[r].height = 21.75
            r += 1
        r += 1
        return refs

    # 1. razão para dB
    refs = bloco("1. Diferença em dB entre dois valores", [
        ("Valor inicial (W ou V)", 250, "W ou V"),
        ("Valor final (W ou V)", 1000, "W ou V"),
        ("Tipo de grandeza", "Potência", "lista"),
    ], [
        ("Diferença", '=IF(OR({0}="",{1}=""),"",IF(OR({0}<=0,{1}<=0),"",ROUND(IF({2}="Tensão ou pressão",20,10)*LOG10({1}/{0}),1)))',
         "dB", SG1, False),
        ("Como se ouve", '=IF({o[0]}="","",IF(ABS({o[0]})<1,"Quase imperceptível",IF(ABS({o[0]})<3,"Pequena: nota-se com atenção",'
                         'IF(ABS({o[0]})<6,"Clara",IF(ABS({o[0]})<10,"Grande","Perto do dobro (ou da metade) do volume, ou mais")))))', None, None, True),
    ], "Potência em watts usa 10 × log. Tensão em volts e pressão sonora usam 20 × log.")
    dv_list(ws, refs[2], ["Potência", "Tensão ou pressão"], "Potência (watts) ou tensão e pressão (volts, pascals).")
    dv_number(ws, f"{refs[0]}:{refs[1]}")

    # 2. dB para razão
    refs = bloco("2. O que uma mudança em dB representa", [
        ("Mudança desejada", 10, "dB"),
        ("Potência atual (opcional)", 300, "W"),
    ], [
        ("A potência fica multiplicada por", '=IF({0}="","",ROUND(10^({0}/10),2))', "vezes", "0.00", False),
        ("A tensão e a pressão ficam multiplicadas por", '=IF({0}="","",ROUND(10^({0}/20),2))', "vezes", "0.00", False),
        ("Potência necessária", '=IF(OR({0}="",{1}=""),"",ROUND({1}*10^({0}/10),0))', "W", "#,##0", False),
    ], "+3 dB dobra a potência. +10 dB, que soa como o dobro do volume, pede dez vezes a potência.")
    dv_number(ws, f"{refs[0]}:{refs[1]}")

    # 3. distância
    refs = bloco("3. Nível a outra distância (fonte pequena, ao ar livre)", [
        ("Nível conhecido", 100, "dB SPL"),
        ("A esta distância", 1, "m"),
        ("Nova distância", 12, "m"),
        ("Nível que se quer alcançar (opcional)", 85, "dB SPL"),
    ], [
        ("Nível previsto na nova distância", '=IF(OR({0}="",{1}="",{2}=""),"",IF(OR({1}<=0,{2}<=0),"",ROUND({0}-20*LOG10({2}/{1}),1)))',
         "dB SPL", "0.0", False),
        ("Variação", '=IF({o[0]}="","",ROUND({o[0]}-{0},1))', "dB", SG1, False),
        ("Distância em que se chega ao nível desejado", '=IF(OR({0}="",{1}="",{3}=""),"",IF({1}<=0,"",ROUND({1}*10^(({0}-{3})/20),1)))',
         "m", "0.0", False),
    ], "−6 dB a cada dobro da distância. Em sala fechada, longe da caixa, o nível cai menos que o previsto por causa das reflexões.")
    dv_number(ws, f"{refs[0]}:{refs[3]}")

    # 4. soma de fontes
    refs = bloco("4. Soma de fontes independentes", [
        ("Fonte 1", 90, "dB"),
        ("Fonte 2", 90, "dB"),
        ("Fonte 3 (opcional)", None, "dB"),
        ("Fonte 4 (opcional)", None, "dB"),
    ], [
        ("Nível total", '=IF(COUNT({0}:{3})=0,"",ROUND(10*LOG10(IF({0}="",0,10^({0}/10))+IF({1}="",0,10^({1}/10))'
                        '+IF({2}="",0,10^({2}/10))+IF({3}="",0,10^({3}/10))),1))', "dB", "0.0", False),
        ("Acréscimo sobre a fonte mais forte", '=IF({o[0]}="","",ROUND({o[0]}-MAX({0}:{3}),1))', "dB", SG1, False),
    ], "Vale para sons diferentes (duas vozes, ruído e música). O mesmo sinal em duas caixas soma até 6 dB onde chega junto.")
    dv_number(ws, f"{refs[0]}:{refs[3]}")

    # 5. tensão e dBu
    refs = bloco("5. Tensão, dBu e dBV", [
        ("Tensão do sinal", 1.228, "V"),
        ("Nível em dBu", -10, "dBu"),
    ], [
        ("A tensão corresponde a", '=IF({0}="","",IF({0}<=0,"",ROUND(20*LOG10({0}/0.775),1)))', "dBu", SG1, False),
        ("e a", '=IF({0}="","",IF({0}<=0,"",ROUND(20*LOG10({0}),1)))', "dBV", SG1, False),
        ("O nível em dBu corresponde a", '=IF({1}="","",ROUND(0.775*10^({1}/20),3))', "V", "0.000", False),
    ], "0 dBu = 0,775 V. 0 dBV = 1 V. A linha profissional trabalha em +4 dBu (1,23 V); a doméstica, em −10 dBV (0,32 V).")
    dv_number(ws, f"{refs[0]}:{refs[1]}")

    # 6. frequência
    refs = bloco("6. Frequência, período e comprimento de onda", [
        ("Frequência", 100, "Hz"),
    ], [
        ("Duração de um ciclo", '=IF({0}="","",IF({0}<=0,"",ROUND(1000/{0},3)))', "ms", "0.000", False),
        ("Comprimento de onda", '=IF(OR({0}="",{C}=""),"",IF({0}<=0,"",ROUND({C}/{0},3)))', "m", "0.000", False),
        ("Uma oitava acima", '=IF({0}="","",ROUND({0}*2,1))', "Hz", "#,##0.0", False),
        ("Uma oitava abaixo", '=IF({0}="","",ROUND({0}/2,1))', "Hz", "#,##0.0", False),
        ("Faixa", '=IF({0}="","",' + faixa_de("{0}") + ')', None, None, True),
    ])
    dv_number(ws, refs[0])

    # 7. atraso
    refs = bloco("7. Diferença de caminho, atraso e cancelamento", [
        ("Diferença de caminho entre os dois sinais", 0.34, "m"),
    ], [
        ("Atraso", '=IF(OR({0}="",{C}=""),"",ROUND({0}/{C}*1000,2))', "ms", "0.00", False),
        ("Primeiro cancelamento", '=IF(OR({0}="",{C}=""),"",IF({0}<=0,"",ROUND({C}/(2*{0}),0)))', "Hz", "#,##0", False),
        ("Segundo cancelamento", '=IF({o[1]}="","",{o[1]}*3)', "Hz", "#,##0", False),
        ("Primeiro reforço", '=IF({o[1]}="","",{o[1]}*2)', "Hz", "#,##0", False),
    ], "Os cancelamentos se repetem a intervalos iguais (filtro pente). Com o segundo sinal 10 dB mais fraco, a ondulação cai para cerca de 3 dB.")
    dv_number(ws, refs[0])
    setup(ws, BLUE, f"B1:G{r - 1}", fit_height=True)


# ================================================================== aba Medição da sala
M_W = {"A": 2, "B": 6, "C": 34, "D": 11, "E": 11, "F": 12, "G": 12, "H": 12, "I": 14, "J": 16, "K": 46, "L": 11, "M": 11,
       "N": 30, "O": 2}
M1 = 17


def aba_med(ws, tab, data=None):
    widths(ws, M_W)
    is_ex = data is not None
    title(ws, "Medição do nível sonoro na sala — ponto a ponto",
          "Exemplo preenchido, para consulta. Use a aba Medição da sala para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro com o que você mediu. As células cinza são calculadas.", "N")
    hb = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None), (8, "Medidor (aparelho)", "medidor", None),
                           (9, "Ponderação", "pond", None), (10, "Resposta", "resp", None), (11, "Sinal de teste", "sinal", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:H{r}", bg=hb, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")
    dv_list(ws, "D9", ["A", "C", "Z"], "A para os limites de exposição; C para ver os graves.")
    dv_list(ws, "D10", ["Lenta", "Rápida", "Leq"], "Lenta (slow), rápida (fast) ou média do período (Leq).")

    ref = data["ref"] if is_ex else (90, 6, None, None)
    rot = [("Limite combinado (dB)", "Nível máximo aceito em qualquer ponto. 90 dB(A) é só ponto de partida."),
           ("Tolerância entre pontos (dB)", "Diferença para o ponto de referência que ainda conta como OK."),
           ("Ruído de fundo (dB)", "Sistema mudo. O sinal precisa ficar 10 dB acima dele."),
           ("Ponto de referência (nº)", "Em geral, a mesa de som.")]
    for k, (t, dica) in enumerate(rot):
        rr = 4 + k
        label(ws, f"I{rr}", t, merge=f"I{rr}:K{rr}")
        inp(ws, f"L{rr}", ref[k], h="center", bg=hb)
        hint(ws, f"M{rr}", dica, merge=f"M{rr}:N{rr}", h="left")
        ws.row_dimensions[rr].height = 21.75
    dv_number(ws, "L4:L7")
    LIM, TOL, RUI, REFN = "$L$4", "$L$5", "$L$6", "$L$7"
    E_, D_, B_ = f"$E${M1}:$E${M1 + N - 1}", f"$D${M1}:$D${M1 + N - 1}", f"$B${M1}:$B${M1 + N - 1}"
    label(ws, "I8", "Nível no ponto de referência (dB)", merge="I8:K8")
    calc(ws, "L8", f'=IF({REFN}="","",IFERROR(IF(INDEX({E_},MATCH({REFN},{B_},0))="","",INDEX({E_},MATCH({REFN},{B_},0))),""))')
    label(ws, "I9", "Distância do ponto de referência (m)", merge="I9:K9")
    calc(ws, "L9", f'=IF({REFN}="","",IFERROR(IF(INDEX({D_},MATCH({REFN},{B_},0))="","",INDEX({D_},MATCH({REFN},{B_},0))),""))')
    for rr in (8, 9):
        hint(ws, f"M{rr}", "Calculado a partir do ponto de referência.", merge=f"M{rr}:N{rr}", h="left")
        ws.row_dimensions[rr].height = 21.75
    LREF, DREF = "$L$8", "$L$9"
    hint(ws, "I10", "Mesmo sinal e mesmo ajuste do sistema em todos os pontos. Medidor na altura do ouvido de quem está sentado.",
         merge="I10:N11", h="left")

    band(ws, 13, "Nível medido por ponto da sala", "N")
    heads = [("B", "Nº", INK), ("C", "Ponto", INK), ("D", "Distância", AMBER), ("E", "Medido", AMBER),
             ("F", "Diferença para a referência", TEAL), ("G", "Previsto pela distância", TEAL), ("H", "Medido − previsto", TEAL),
             ("I", "Tempo de referência", TEAL), ("J", "Situação", TEAL), ("K", "O que fazer", TEAL),
             ("L", "Depois", AMBER), ("M", "Variação", TEAL), ("N", "Observação", INK)]
    for col, t, bg in heads:
        head(ws, f"{col}14", t, bg=bg)
    ws.row_dimensions[14].height = 30
    for col, t in zip("BCDEFGHIJKLMN", ["", "onde fica o medidor", "m, da caixa mais próxima", "dB", "dB", "dB", "dB",
                                        "por dia, em dB(A)", "", "", "dB, após corrigir", "dB", ""]):
        hint(ws, f"{col}15", t)
    ws.row_dimensions[15].height = 27
    if is_ex:
        hint(ws, "B16", "Leituras em dB(A), resposta lenta, com o mesmo sinal em todos os pontos. A coluna Depois é a nova leitura, com as caixas elevadas.",
             merge="B16:N16", h="left")
    else:
        for col, v in zip("BCDEFGHIJKLM", ["Ex.", "Primeira fileira, no centro", 3, 96, 10, 98, -2, 38, "Acima do limite",
                                           "Eleve ou redirecione a caixa; se não bastar, baixe o master", 89, -7]):
            ex(ws, f"{col}16", v, h="left" if col in "CK" else "center")
        ex(ws, "N16", "")
        note(ws, "D14", "Distância em metros da caixa mais próxima até o medidor.\nUse trena, ou passos largos (perto de 1 m cada).")
        note(ws, "E14", "Leitura do medidor, só o número.\nEspere uns 10 segundos para a leitura estabilizar.")
        note(ws, "G14", "Nível que o ponto teria pela lei do inverso do quadrado (−6 dB a cada dobro da distância), partindo do ponto de referência.")
        note(ws, "I14", "Tempo de exposição diária de referência para esse nível: 85 dB(A) por 8 horas, e metade do tempo a cada 3 dB a mais.\nNão é limite legal.")
        note(ws, "L14", "Nova leitura no mesmo ponto, depois de corrigir a posição das caixas ou o nível.")
    ws.row_dimensions[16].height = 24

    M2 = M1 + N - 1
    for k in range(N):
        r = M1 + k
        p = data["pontos"][k] if is_ex and k < len(data["pontos"]) else None
        bg = WHITE if is_ex else INPUT
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", p[0] if p else None, bg=bg)
        inp(ws, f"D{r}", p[1] if p else None, h="center", bg=bg)
        inp(ws, f"E{r}", p[2] if p else None, h="center", bg=bg)
        D, E, F, G, J, L = (f"{c}{r}" for c in "DEFGJL")
        calc(ws, F, f'=IF(OR({E}="",{LREF}=""),"",{E}-{LREF})', fmt=SG, b=False)
        calc(ws, G, f'=IF(OR({D}="",{LREF}="",{DREF}=""),"",IF(OR({D}<=0,{DREF}<=0),"",ROUND({LREF}-20*LOG10({D}/{DREF}),1)))',
             fmt="0.0", b=False)
        calc(ws, f"H{r}", f'=IF(OR({E}="",{G}=""),"",ROUND({E}-{G},1))', fmt=SG1, b=False)
        calc(ws, f"I{r}", f'=IF({E}="","",IF({E}<85,"mais de 8 h",ROUND(480/2^(({E}-85)/3),0)))', fmt='0" min"', b=False)
        calc(ws, J, f'=IF({E}="","",IF(AND({RUI}<>"",{E}-{RUI}<10),"Ruído",IF(AND({LIM}<>"",{E}>{LIM}),"Acima do limite",'
                    f'IF({F}="","OK",IF({F}>{TOL},"Forte",IF({F}<-{TOL},"Fraco","OK"))))))')
        calc(ws, f"K{r}", f'=IF({J}="","",IF({J}="Ruído","Sinal de teste fraco: suba o sinal ou meça com menos ruído",'
                          f'IF({J}="Acima do limite",IF(AND({F}<>"",{F}>0),"Eleve ou redirecione a caixa; se não bastar, baixe o master",'
                          f'"Sala toda alta: baixe o master"),'
                          f'IF({J}="Forte","Eleve ou redirecione a caixa; afaste o público da frente dela",'
                          f'IF({J}="Fraco","Redirecione a caixa, confira obstáculos ou use caixa de reforço","Manter")))))',
             h="left", b=False, sz=9)
        inp(ws, L, p[3] if p else None, h="center", bg=bg)
        calc(ws, f"M{r}", f'=IF(OR({E}="",{L}=""),"",{L}-{E})', fmt=SG, b=False)
        inp(ws, f"N{r}", bg=bg)
        ws.row_dimensions[r].height = 21.75
    dv_number(ws, f"D{M1}:E{M2}")
    dv_number(ws, f"L{M1}:L{M2}")
    cf_equal(ws, f"J{M1}:J{M2}", SIT_CF)

    s = M2 + 2
    band(ws, s, "Resumo automático", "N")
    cont = lambda t: f'COUNTIF(J{M1}:J{M2},"{t}")'
    Er, Lr = f"E{M1}:E{M2}", f"L{M1}:L{M2}"
    linhas = [
        ("Pontos medidos", f"=COUNT({Er})", None),
        ("Nível mais alto (dB)", f'=IF(COUNT({Er})=0,"",MAX({Er}))', None),
        ("Nível mais baixo (dB)", f'=IF(COUNT({Er})=0,"",MIN({Er}))', None),
        ("Diferença entre o mais alto e o mais baixo (dB)", f'=IF(COUNT({Er})=0,"",MAX({Er})-MIN({Er}))', None),
        ("Pontos acima do limite", f'={cont("Acima do limite")}', None),
        ("Pontos fora da tolerância (Forte ou Fraco)", f'={cont("Forte")}+{cont("Fraco")}', None),
        ("Pontos com sinal fraco demais (Ruído)", f'={cont("Ruído")}', None),
        ("Nível máximo no ponto de referência para respeitar o limite (dB)",
         f'=IF(OR({LIM}="",{LREF}="",COUNT({Er})=0),"",{LIM}-(MAX({Er})-{LREF}))', None),
        ("Depois: nível mais alto (dB)", f'=IF(COUNT({Lr})=0,"",MAX({Lr}))', None),
        ("Depois: nível mais baixo (dB)", f'=IF(COUNT({Lr})=0,"",MIN({Lr}))', None),
        ("Depois: diferença entre o mais alto e o mais baixo (dB)", f'=IF(COUNT({Lr})=0,"",MAX({Lr})-MIN({Lr}))', None),
    ]
    for k, (t, f, fmt) in enumerate(linhas, 1):
        summary(ws, s + k, t, f, "H", "I", fmt=fmt)
    a = s + len(linhas) + 1
    label(ws, f"B{a}", "Aviso", merge=f"B{a}:H{a}", h="right")
    calc(ws, f"I{a}", f'=IF(I{s + 1}=0,"Meça os pontos",IF(I{s + 7}>0,"Há ponto com sinal de teste fraco demais: meça de novo",'
                      f'IF(I{s + 5}>0,"Há ponto acima do limite: corrija a posição das caixas antes de baixar o volume",'
                      f'IF(I{s + 6}>0,"Há pontos fora da tolerância: reveja a cobertura das caixas","OK"))))',
         b=False, merge=f"I{a}:N{a}", h="left")
    ws.row_dimensions[a].height = 21.75
    cf_ok(ws, f"I{a}:N{a}")
    hint(ws, f"B{a + 1}", "O previsto pela distância vale para fonte pequena em espaço aberto. Em sala, longe da caixa, é normal o medido ficar acima do previsto. "
         "O tempo de referência segue o critério de 85 dB(A) por 8 horas e metade do tempo a cada 3 dB: é referência para planejar, e não limite legal.",
         merge=f"B{a + 1}:N{a + 1}", h="left")
    ws.row_dimensions[a + 1].height = 31.5
    end = a + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "N")
        r = end + 3
        put(ws, f"B{r}", "Item", f=font(10, True), bg=GRAY, h="center", merge=f"B{r}:C{r}")
        put(ws, f"D{r}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{r}:N{r}")
        ws.row_dimensions[r].height = 21.75
        for item, feito in data["ajustes"]:
            r += 1
            put(ws, f"B{r}", item, f=font(10, True), bg=GRAY, merge=f"B{r}:C{r}")
            put(ws, f"D{r}", feito, merge=f"D{r}:N{r}")
            ws.row_dimensions[r].height = 21.75
        end = r
    ws.freeze_panes = f"D{M1}"
    setup(ws, tab, f"B1:N{end}", fit_height=True)


# ================================================================== aba Faixas por fonte
def aba_faixas(ws):
    widths(ws, {"A": 2, "B": 34, "C": 13, "D": 13, "E": 13, "F": 17, "G": 17, "H": 13, "I": 52, "J": 2})
    title(ws, "Faixas de frequência e extensão de cada fonte",
          "As faixas e as fontes são valores típicos. Troque ou acrescente fontes nas células em amarelo-claro.", "I")
    band(ws, 4, "As faixas de frequência", "I")
    for col, t in zip("BCD", ["Faixa", "De (Hz)", "Até (Hz)"]):
        head(ws, f"{col}5", t)
    head(ws, "E5", "O que fica ali")
    ws.merge_cells("E5:G5")
    head(ws, "H5", "Como se percebe")
    ws.merge_cells("H5:I5")
    ws.row_dimensions[5].height = 21.75
    for k, (nome, de, ate, ali, perc) in enumerate(FAIXAS):
        r = 6 + k
        put(ws, f"B{r}", nome, f=font(10, True), bg=GRAY)
        put(ws, f"C{r}", de, bg=GRAY, h="center", fmt="#,##0")
        put(ws, f"D{r}", ate, bg=GRAY, h="center", fmt="#,##0")
        put(ws, f"E{r}", ali, merge=f"E{r}:G{r}")
        put(ws, f"H{r}", perc, merge=f"H{r}:I{r}")
        ws.row_dimensions[r].height = 31.5

    band(ws, 13, "Extensão de cada fonte", "I")
    for col, t, bg in [("B", "Fonte", INK), ("C", "Fundamental mais grave (Hz)", AMBER), ("D", "Fundamental mais aguda (Hz)", AMBER),
                       ("E", "Vai até (Hz)", AMBER), ("F", "Faixa da nota mais grave", TEAL), ("G", "Faixa do brilho", TEAL),
                       ("H", "Oitavas de notas", TEAL), ("I", "Observação", INK)]:
        head(ws, f"{col}14", t, bg=bg)
    ws.row_dimensions[14].height = 30
    note(ws, "E14", "Até onde vão os harmônicos, as consoantes ou o ataque: o que dá definição à fonte.")
    note(ws, "H14", "Quantas vezes a nota mais aguda dobra a mais grave.")
    r1 = 15
    linhas = FONTES + [None] * 6
    for k, f in enumerate(linhas):
        r = r1 + k
        vals = f if f else (None,) * 5
        inp(ws, f"B{r}", vals[0])
        for col, v in zip("CDE", vals[1:4]):
            inp(ws, f"{col}{r}", v, h="center", fmt="#,##0")
        C, D, E = f"C{r}", f"D{r}", f"E{r}"
        calc(ws, f"F{r}", f'=IF({C}="","",{faixa_de(C)})', b=False)
        calc(ws, f"G{r}", f'=IF({E}="","",{faixa_de(E)})', b=False)
        calc(ws, f"H{r}", f'=IF(OR({C}="",{D}=""),"",IF(OR({C}<=0,{D}<{C}),"",ROUND(LOG10({D}/{C})/LOG10(2),1)))', fmt="0.0", b=False)
        inp(ws, f"I{r}", vals[4])
        ws.row_dimensions[r].height = 21.75
    r2 = r1 + len(linhas) - 1
    dv_number(ws, f"C{r1}:E{r2}")
    hint(ws, f"B{r2 + 1}", "Extensões típicas: um cantor, um instrumento ou uma afinação diferente mudam os limites. O que cortar e realçar em cada faixa está no estudo Equalização.",
         merge=f"B{r2 + 1}:I{r2 + 1}", h="left")
    ws.row_dimensions[r2 + 1].height = 21.75
    ws.freeze_panes = "C15"
    setup(ws, TEAL, f"B1:I{r2 + 1}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Som, frequência e decibel — calculadora e medição da sala"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Som, frequência e decibel — como usar esta planilha",
           "Calculadora de dB, medição do nível sonoro por ponto da sala e faixas de frequência por fonte.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Tipo de grandeza, ponderação e resposta do medidor aceitam só as opções da lista."),
        ("Números", "Digite só o número, com o sinal de menos quando houver, e vírgula para decimais: 2,5, e não “2,5 m”."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Medição da sala: preencha o cabeçalho, o limite combinado e a tolerância entre os pontos."),
        ("2", "Escolha os pontos (de seis a dez), meça a distância de cada um até a caixa mais próxima e anote."),
        ("3", "Meça o ruído de fundo com o sistema mudo e anote no alto da aba."),
        ("4", "Toque um sinal estável, sem mexer no sistema, e anote a leitura de cada ponto. Informe o número do ponto de referência (em geral, a mesa)."),
        ("5", "Veja a situação e o que fazer. Corrija a posição das caixas, meça de novo e anote na coluna Depois."),
        ("6", "Aba Calculadora de dB: use para estimar antes de comprar, mudar ou mover equipamento."),
        ("7", "Aba Checklist: confira os doze itens na montagem e na passagem de som."),
    ]),
    ("Abas da planilha", [
        ("Calculadora de dB", "Sete blocos: razão para dB, dB para razão, distância, soma de fontes, tensão e dBu, frequência, atraso."),
        ("Medição da sala", "Até 16 pontos, com diferença para a referência, previsto pela distância, tempo de referência e situação."),
        ("Faixas por fonte", "As seis faixas de frequência e a extensão de onze fontes, com linhas livres para as suas."),
        ("Referência", "Regras do decibel, escalas, níveis típicos, tempos de exposição, soma de fontes, distância e atraso."),
        ("Checklist", "Os doze itens do treinamento, com percentual atendido."),
        ("Exemplo - Medição", "A medição de uma igreja de 200 lugares, antes e depois de elevar as caixas (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Valores de referência", "Os números são ponto de partida, e não norma. Confira o manual do seu equipamento."),
        ("Previsto pela distância", "Lei do inverso do quadrado, para fonte pequena ao ar livre: −6 dB a cada dobro da distância. Em sala, longe da caixa, o nível cai menos."),
        ("Situação do ponto", "Ruído: menos de 10 dB acima do ruído de fundo. Acima do limite: passou do limite combinado. Forte ou Fraco: diferença para a referência maior que a tolerância."),
        ("Tempo de referência", "85 dB(A) por 8 horas, e metade do tempo a cada 3 dB a mais. É referência para planejar, e não limite legal: consulte a regulamentação vigente."),
        ("Soma de fontes", "Vale para sons independentes. O mesmo sinal em duas caixas soma até 6 dB onde as ondas chegam juntas."),
        ("Exemplos", "Os dados da aba de exemplo e os valores iniciais da calculadora são ilustrativos, criados para este material."),
    ]),
])

aba_calc(wb.create_sheet("Calculadora de dB"))
aba_med(wb.create_sheet("Medição da sala"), AMBER)
aba_faixas(wb.create_sheet("Faixas por fonte"))

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 30, "C": 24, "D": 26, "E": 62, "F": 2})
title(ws, "Referência — decibel, distância, exposição e frequência",
      "Valores de referência, para servir de ponto de partida. Confira o manual do seu equipamento.", "E")


def tabela(r, titulo, cab, linhas, alt=21.75):
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
        ws.row_dimensions[r].height = alt if len(lin[-1]) <= 70 else 31.5
    return r + 2


r = tabela(4, "Regras práticas do decibel", ["Mudança", "Na potência (W)", "Na tensão e na pressão", "Onde aparece"], [
    ("1 dB", "×1,26", "×1,12", "A menor diferença que se nota com atenção. Ao vivo, 2 a 3 dB já são claros."),
    ("3 dB", "×2", "×1,41", "Amplificador de potência dobrada. Duas fontes iguais e independentes."),
    ("6 dB", "×4", "×2", "Metade da distância. Dois sinais idênticos e em fase."),
    ("10 dB", "×10", "×3,16", "O passo que soa como o dobro do volume."),
    ("20 dB", "×100", "×10", "Valor usual do pad. Diferença entre linha e microfone forte."),
])
r = tabela(r, "As escalas de decibel", ["Escala", "O que mede", "O zero é", "Valores típicos"], [
    ("dB SPL", "pressão do som no ar", "limiar da audição (20 µPa)", "Conversa a 1 m: 60. Banda ao vivo na plateia: de 85 a 100."),
    ("dBu", "tensão elétrica", "0,775 V", "Microfone: de −60 a −30. Linha profissional: +4."),
    ("dBV", "tensão elétrica", "1 V", "Linha doméstica: −10 dBV, cerca de 12 dB abaixo da profissional."),
    ("dBFS", "valor no sistema digital", "o máximo representável", "Trabalho: perto de −18. Clip: 0. As leituras são negativas."),
    ("dB(A) e dB(C)", "pressão sonora ponderada", "como no dB SPL", "A: como o ouvido, usada na exposição. C: quase plana, mostra o grave."),
])
r = tabela(r, "Níveis sonoros típicos", ["Situação", "Nível (dB SPL)", "Tempo de referência", "Observação"], [
    ("Quarto silencioso", "20 a 30", "–", "Noite, sem aparelhos ligados."),
    ("Salão vazio com ar-condicionado", "40 a 50", "–", "Meça o ruído de fundo antes do evento."),
    ("Conversa a 1 m", "60 a 65", "–", "Uma fala amplificada precisa ficar acima do ruído da sala."),
    ("Palestra ou pregação amplificada", "65 a 75", "–", "Na plateia."),
    ("Música ao vivo com banda", "85 a 95", "de 8 h a 1 h", "Varia muito com o estilo e com a casa."),
    ("Bateria acústica, para quem toca", "95 a 105", "de 1 h a 7 min", "Plateia perto das caixas também chega lá."),
    ("Limiar do desconforto", "cerca de 120", "–", "Acima disso, risco de dano em pouco tempo."),
])
r = tabela(r, "Tempo de referência por dia (critério de 3 dB)", ["Nível médio", "Tempo", "Comparado a 85 dB(A)", "Observação"], [
    ("85 dB(A)", "8 h", "referência", "Média do período, e não picos isolados."),
    ("88 dB(A)", "4 h", "metade", ""),
    ("91 dB(A)", "2 h", "um quarto", "Um culto com duas horas de música nesse nível usa a referência do dia."),
    ("94 dB(A)", "1 h", "um oitavo", ""),
    ("97 dB(A)", "30 min", "1/16", ""),
    ("100 dB(A)", "15 min", "1/32", "Perto das caixas e da bateria, é fácil chegar lá."),
    ("103 dB(A)", "7,5 min", "1/64", "Há critérios legais menos rigorosos (5 dB): consulte a regulamentação vigente."),
])
r = tabela(r, "Soma de duas fontes independentes", ["Diferença entre elas", "Some à mais forte", "Exemplo", "Observação"], [
    ("0 dB", "+3,0 dB", "90 e 90 dão 93", "Duas fontes iguais."),
    ("1 dB", "+2,5 dB", "90 e 89 dão 92,5", ""),
    ("3 dB", "+1,8 dB", "90 e 87 dão 91,8", ""),
    ("6 dB", "+1,0 dB", "90 e 84 dão 91", ""),
    ("10 dB ou mais", "+0,4 dB ou menos", "90 e 80 dão 90,4", "Por isso o sinal de teste precisa ficar 10 dB acima do ruído."),
])
r = tabela(r, "Distância (fonte pequena, ao ar livre)", ["A distância", "O nível", "No microfone", "Na plateia"], [
    ("Cai à metade", "+6 dB", "de 10 para 5 cm", "da fileira a 8 m para a fileira a 4 m"),
    ("Dobra", "−6 dB", "de um palmo para dois", "de 4 para 8 m"),
    ("Fica 3 vezes maior", "−9,5 dB", "de 10 para 30 cm", "de 5 para 15 m"),
    ("Fica 4 vezes maior", "−12 dB", "de 5 para 20 cm", "de 3 para 12 m"),
    ("Fica 10 vezes maior", "−20 dB", "de 5 para 50 cm", "de 2 para 20 m"),
])
r = tabela(r, "Frequência, comprimento de onda e atraso (a 343 m/s)", ["Frequência", "Duração do ciclo", "Comprimento de onda", "Atraso que cancela essa frequência"], [
    ("50 Hz", "20 ms", "6,9 m", "10 ms (3,4 m de diferença de caminho)"),
    ("100 Hz", "10 ms", "3,4 m", "5 ms (1,7 m)"),
    ("250 Hz", "4 ms", "1,4 m", "2 ms (69 cm)"),
    ("500 Hz", "2 ms", "69 cm", "1 ms (34 cm)"),
    ("1 kHz", "1 ms", "34 cm", "0,5 ms (17 cm)"),
    ("4 kHz", "0,25 ms", "8,6 cm", "0,125 ms (4,3 cm)"),
    ("10 kHz", "0,1 ms", "3,4 cm", "0,05 ms (1,7 cm)"),
])
r = tabela(r, "Digital", ["Item", "Valor usual", "O que significa", "Observação"], [
    ("Taxa de amostragem", "48 kHz", "registra até cerca de 24 kHz", "44,1 kHz também é comum. Equipamentos ligados em digital usam a mesma taxa."),
    ("Profundidade", "24 bits", "cerca de 6 dB por bit", "16 bits: cerca de 96 dB. Com 24 bits, o limite é o conversor: 110 a 120 dB."),
    ("Teto", "0 dBFS", "acima dele, clip", "O corte é seco. Trabalhe perto de −18 dBFS."),
])
hint(ws, f"B{r - 1}", "Valores típicos, arredondados. Os níveis e os tempos de exposição são referência para planejar, e não limite legal.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Checklist e exemplo
checklist(wb.create_sheet("Checklist"), "Checklist de som e nível na sala",
          "Confira na montagem e na passagem de som. Escolha a situação de cada item na lista suspensa.",
          CHECK, tab=HUE["purple"][0])
aba_med(wb.create_sheet("Exemplo - Medição"), MUTED, EX_MED)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
