# -*- coding: utf-8 -*-
"""Gera Transmissao-ao-Vivo-modelo.xlsx: bitrate × upload, ficha de configuração, checklist com horários e registro de quedas."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_lib import *

OUT = sys.argv[1]

REDH, RED_T = HUE["red"]
BLUE, BLUE_T = HUE["blue"]
TEAL, TEAL_T = HUE["teal"]
AMBER, AMBER_T = HUE["amber"]
SIM_NAO = ["Sim", "Não"]
REDE_CF = [("OK", GREEN), ("Apertado", YELLOW), ("Insuficiente", RED)]
HORA = "hh:mm"
KBPS = "#,##0"

# formato, mínimo e máximo de vídeo em kbps (H.264), observação
FORMATOS = [
    ("480p30", 1000, 2000, "Rede fraca ou link de reserva. Serve para fala."),
    ("720p30", 2500, 4000, "O formato seguro para a maioria dos eventos pequenos."),
    ("720p60", 3500, 6000, "Movimento rápido. Raro em culto e palestra."),
    ("1080p30", 4000, 6000, "O padrão quando a rede comporta."),
    ("1080p60", 6000, 9000, "Só com rede e processador com muita folga."),
]

# (quando, minutos antes do início ou None para a véspera, item)
CHECK_T = [
    ("Véspera", None, "O evento está criado na plataforma, com título, data e visibilidade certos, e a chave no codificador é a desse evento."),
    ("Véspera", None, "Um teste não listado, de pelo menos dez minutos, foi transmitido e assistido em outro aparelho."),
    ("2 h antes", 120, "O upload foi medido no local, por cabo, e o bitrate total usa até cerca de metade do menor valor medido."),
    ("2 h antes", 120, "O codificador está no cabo de rede e na tomada, com atualizações, suspensão e notificações desligadas."),
    ("1 h 30 antes", 90, "Resolução, quadros, bitrate, CBR e keyframe de 2 segundos conferem com a ficha de configuração."),
    ("1 h antes", 60, "O link de reserva foi testado, e o roteiro de contingência está à vista do operador."),
    ("1 h antes", 60, "O áudio vem de uma mistura própria (auxiliar, matriz ou mesa dedicada), com microfones de ambiente."),
    ("45 min antes", 45, "Com a banda e com a fala, os picos ficam abaixo de −1 dBFS no medidor do codificador, e há som nos dois lados."),
    ("45 min antes", 45, "O sincronismo entre som e imagem foi conferido com o teste da palma."),
    ("30 min antes", 30, "As cenas e os gráficos foram conferidos: espera, programa, legendas, problema técnico e encerramento."),
    ("15 min antes", 15, "A gravação local está rodando, e a transmissão entrou no ar na tela de espera."),
    ("No ar", 0, "Uma pessoa acompanha o retorno em um celular com dados móveis e fone, olha os indicadores e registra as quedas."),
]

# Exemplo 1 do treinamento: a transmissão que travava com a igreja cheia (situação antes da correção)
EX_BITRATE = {
    "head": {"evento": "Culto de domingo, 19 h", "local": "Igreja de 200 lugares", "data": date(2026, 10, 4),
             "operador": "Operador da transmissão da escala", "plataforma": "Plataforma de vídeo, evento agendado"},
    "med": [  # quando, conexão, upload, download, observação
        ("Quinta-feira, 20 h, sala vazia", "Cabo", 12.4, 92, "O teste da quinta foi feito nesta condição."),
        ("Domingo, 18 h, antes do culto", "Cabo", 9.8, 85, "A equipe já estava usando a rede."),
        ("Domingo, 19 h 30, sala cheia", "Cabo", 8.2, 60, "Visitantes no Wi-Fi da mesma rede."),
    ],
    "cfg": ("1080p", 60, 9000, 160, 0.5, 120, 6),
    "ajustes": [  # o quê, o que foi feito, resultado
        ("Formato", "De 1080p a 60 quadros para 720p a 30 quadros, com vídeo a 3.500 kbps.",
         "Total de 3.660 kbps: 45% do upload de 8,2 Mbps."),
        ("Rede", "O Wi-Fi dos visitantes passou para outra rede, com limite de banda.",
         "Nova medição com a sala cheia: 11,5 Mbps."),
        ("Semana seguinte", "Com o novo upload, 1080p a 30 quadros, com vídeo a 4.500 kbps.",
         "Total de 4.660 kbps: 41% do upload de 11,5 Mbps."),
        ("Link de reserva", "Perfil próprio para o modem 4G de 6 Mbps: 720p a 30 quadros, com vídeo a 2.500 kbps.",
         "Total de 2.660 kbps: 44% do upload do link de reserva."),
    ],
}


def hint(ws, ref, text, bg=GRAY, merge=None, h="center"):
    put(ws, ref, text, f=font(9, i=True, c=MUTED), bg=bg, h=h, merge=merge)


# ================================================================== aba Bitrate e upload
B_W = {"A": 2, "B": 5, "C": 34, "D": 14, "E": 16, "F": 14, "G": 16, "H": 14, "I": 50, "J": 2}
NMED = 5


def aba_bitrate(ws, tab, data=None):
    widths(ws, B_W)
    is_ex = data is not None
    title(ws, "Bitrate e upload — a transmissão cabe na rede?",
          "Exemplo preenchido, para consulta. Use a aba Bitrate e upload para o seu evento." if is_ex
          else "Preencha as células em amarelo-claro com o que você mediu e com a configuração pretendida. As células cinza são calculadas.", "I")
    bg = WHITE if is_ex else INPUT
    H = data["head"] if is_ex else {}
    for r, rot, k, fmt in [(4, "Evento", "evento", None), (5, "Local", "local", None), (6, "Data", "data", DATE),
                           (7, "Operador", "operador", None), (8, "Plataforma", "plataforma", None)]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", H.get(k), merge=f"D{r}:G{r}", bg=bg, fmt=fmt)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D6")

    # ---------------- 1. medições
    band(ws, 10, "1. Medições de upload: no local, por cabo, no computador da transmissão", "I")
    for col, t in zip("BCDEF", ["#", "Quando foi medido", "Conexão", "Upload", "Download"]):
        head(ws, f"{col}11", t, bg=REDH if col == "E" else INK)
    put(ws, "G11", "Observação", f=font(10, True, c=WHITE), bg=INK, h="center", merge="G11:I11")
    ws.row_dimensions[11].height = 24
    for col, t in zip("BCDEF", ["", "dia, hora e condição da sala", "lista", "Mbps", "Mbps (opcional)"]):
        hint(ws, f"{col}12", t)
    hint(ws, "G12", "quem mais estava usando a rede", merge="G12:I12")
    if is_ex:
        hint(ws, "B13", "Três medições no mesmo ponto de rede. Vale a menor, feita com a sala cheia.", merge="B13:I13", h="left")
    else:
        for col, v in zip("BCDEF", ["Ex.", "Domingo, 18 h 30, sala enchendo", "Cabo", 9.8, 85]):
            ex(ws, f"{col}13", v, h="left" if col == "C" else "center", fmt="0.0" if col == "E" else None)
        put(ws, "G13", "Wi-Fi do público na mesma rede", f=font(9, i=True, c=MUTED), bg=GRAY, merge="G13:I13")
        note(ws, "E11", "Velocidade de envio, em Mbps, lida em um teste de velocidade.\nNão use o valor do contrato: use o que você mediu.")
        note(ws, "D11", "Meça por cabo. A medição por Wi-Fi ou por rede móvel serve só para comparar, ou para o link de reserva.")
    ws.row_dimensions[13].height = 21.75
    M1 = 14
    M2 = M1 + NMED - 1
    for k in range(NMED):
        r = M1 + k
        m = data["med"][k] if is_ex and k < len(data["med"]) else None
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", m[0] if m else None, bg=bg)
        inp(ws, f"D{r}", m[1] if m else None, h="center", bg=bg)
        inp(ws, f"E{r}", m[2] if m else None, h="center", bg=bg, fmt="0.0")
        inp(ws, f"F{r}", m[3] if m else None, h="center", bg=bg, fmt="0")
        inp(ws, f"G{r}", m[4] if m else None, bg=bg, merge=f"G{r}:I{r}")
        ws.row_dimensions[r].height = 21.75
    dv_list(ws, f"D{M1}:D{M2}", ["Cabo", "Wi-Fi", "4G ou 5G"], "Cabo, Wi-Fi ou 4G ou 5G.")
    dv_number(ws, f"E{M1}:F{M2}")
    UP = f"$E${M2 + 1}"
    r = M2 + 1
    label(ws, f"B{r}", "Menor upload medido (Mbps)", merge=f"B{r}:D{r}", h="right")
    calc(ws, f"E{r}", f'=IF(COUNT(E{M1}:E{M2})=0,"",MIN(E{M1}:E{M2}))', fmt="0.0")
    hint(ws, f"F{r}", "É este o valor que entra na conta: a transmissão precisa caber no pior momento.", merge=f"F{r}:I{r}", h="left")
    label(ws, f"B{r + 1}", "Medições feitas", merge=f"B{r + 1}:D{r + 1}", h="right")
    calc(ws, f"E{r + 1}", f"=COUNT(E{M1}:E{M2})")
    calc(ws, f"F{r + 1}", f'=IF(E{r + 1}=0,"",IF(E{r + 1}<3,"Meça pelo menos três vezes, uma delas no horário do evento","OK"))',
         b=False, merge=f"F{r + 1}:I{r + 1}", h="left")
    cf_status(ws, f"F{r + 1}:I{r + 1}", f"$F${r + 1}")
    label(ws, f"B{r + 2}", "Variação entre a maior e a menor", merge=f"B{r + 2}:D{r + 2}", h="right")
    calc(ws, f"E{r + 2}", f'=IF(COUNT(E{M1}:E{M2})<2,"",(MAX(E{M1}:E{M2})-MIN(E{M1}:E{M2}))/MAX(E{M1}:E{M2}))', fmt="0%")
    calc(ws, f"F{r + 2}", f'=IF(E{r + 2}="","",IF(E{r + 2}>0.3,"Variação alta: há mais gente usando este upload, ou a rede é instável","OK"))',
         b=False, merge=f"F{r + 2}:I{r + 2}", h="left")
    cf_status(ws, f"F{r + 2}:I{r + 2}", f"$F${r + 2}")
    for k in range(3):
        ws.row_dimensions[r + k].height = 21.75

    # ---------------- 2. configuração pretendida
    c0 = r + 4
    band(ws, c0, "2. Configuração pretendida", "I")
    cfg = data["cfg"] if is_ex else (None, None, None, None, 0.5, None, None)
    linhas = [
        ("Resolução", cfg[0], None, "480p, 720p ou 1080p."),
        ("Quadros por segundo", cfg[1], None, "30 ou 60. Para a transmissão, 29,97 conta como 30, e 59,94 como 60."),
        ("Bitrate de vídeo (kbps)", cfg[2], KBPS, "O valor digitado no codificador. 4,5 Mbps são 4.500 kbps."),
        ("Bitrate de áudio (kbps)", cfg[3], KBPS, "Em geral, de 128 a 160 kbps."),
        ("Fração máxima do upload", cfg[4], "0%", "Regra prática: 50%. Em Wi-Fi ou rede móvel, use menos."),
        ("Duração prevista (min)", cfg[5], "0", "Com o tempo de tela de espera. Serve para o consumo de dados."),
        ("Upload do link de reserva (Mbps)", cfg[6], "0.0", "Opcional. Medido no local, como o do link principal."),
    ]
    for k, (rot, v, fmt, dica) in enumerate(linhas):
        rr = c0 + 1 + k
        label(ws, f"B{rr}", rot, merge=f"B{rr}:D{rr}", h="right")
        inp(ws, f"E{rr}", v, h="center", bg=bg, fmt=fmt)
        hint(ws, f"F{rr}", dica, merge=f"F{rr}:I{rr}", h="left")
        ws.row_dimensions[rr].height = 21.75
    RES, FPS, VID, AUD, FRA, DUR, RSV = (f"$E${c0 + 1 + k}" for k in range(7))
    dv_list(ws, RES.replace("$", ""), ["480p", "720p", "1080p"], "480p, 720p ou 1080p.")
    dv_list(ws, FPS.replace("$", ""), ["30", "60"], "30 ou 60.")
    dv_number(ws, f"E{c0 + 3}:E{c0 + 7}")

    # a tabela de formatos fica no fim da aba; as fórmulas de resultado já apontam para ela
    r0 = c0 + 9                      # faixa "3. Resultado"
    NRES = 13
    t0 = r0 + NRES + 2               # faixa "4. Qual formato cabe"
    T1 = t0 + 2
    T2 = T1 + len(FORMATOS) - 1
    TAB = f"$C${T1}:$E${T2}"

    # ---------------- 3. resultado
    band(ws, r0, "3. Resultado", "I")
    TOT, MBP, LIM, USO, SIT, VMAX, FMIN, FMAX, SITB, GB, USOR, SITR, ACAO = (f"$E${r0 + 1 + k}" for k in range(NRES))
    res = [
        ("Bitrate total (kbps)", f'=IF(OR({VID}="",{AUD}=""),"",{VID}+{AUD})', KBPS, "Vídeo mais áudio."),
        ("Bitrate total (Mbps)", f'=IF({TOT}="","",{TOT}/1000)', "0.00", "O mesmo valor, na unidade do teste de velocidade."),
        ("Limite para o bitrate total (kbps)", f'=IF(OR({UP}="",{FRA}=""),"",ROUND({UP}*1000*{FRA},0))', KBPS,
         "Menor upload medido × fração máxima."),
        ("Uso do upload", f'=IF(OR({TOT}="",{UP}=""),"",{TOT}/({UP}*1000))', "0%", "Bitrate total ÷ menor upload medido."),
        ("Situação da rede", f'=IF(OR({USO}="",{FRA}=""),"",IF({USO}<={FRA},"OK",IF({USO}<=MAX({FRA},0.75),"Apertado","Insuficiente")))',
         None, "OK até a fração máxima. Apertado até 75%. Acima disso, Insuficiente."),
        ("Bitrate de vídeo máximo (kbps)", f'=IF(OR({LIM}="",{AUD}=""),"",MAX(0,{LIM}-{AUD}))', KBPS,
         "O que sobra para o vídeo dentro do limite."),
        ("Faixa do formato: mínimo (kbps)", f'=IFERROR(VLOOKUP({RES}&{FPS},{TAB},2,FALSE),"")', KBPS,
         "Da tabela de formatos, no fim desta aba."),
        ("Faixa do formato: máximo (kbps)", f'=IFERROR(VLOOKUP({RES}&{FPS},{TAB},3,FALSE),"")', KBPS,
         "Acima disso a imagem quase não melhora."),
        ("Situação do bitrate de vídeo", f'=IF(OR({VID}="",{FMIN}=""),"",IF({VID}<{FMIN},"Abaixo da faixa",IF({VID}>{FMAX},"Acima da faixa","OK")))',
         None, "Compara o bitrate de vídeo com a faixa do formato escolhido."),
        ("Consumo de dados (GB)", f'=IF(OR({TOT}="",{DUR}=""),"",ROUND({TOT}*{DUR}*60/8/1000000,1))', "0.0",
         "Importante quando o link tem franquia de dados."),
        ("Uso do link de reserva", f'=IF(OR({TOT}="",{RSV}=""),"",{TOT}/({RSV}*1000))', "0%",
         "A mesma configuração, se a transmissão passar para o link de reserva."),
        ("Situação no link de reserva", f'=IF(OR({USOR}="",{FRA}=""),"",IF({USOR}<={FRA},"OK",IF({USOR}<=MAX({FRA},0.75),"Apertado","Insuficiente")))',
         None, "Se não for OK, deixe pronto um perfil de bitrate menor para o link de reserva."),
    ]
    for k, (rot, f, fmt, dica) in enumerate(res):
        rr = r0 + 1 + k
        label(ws, f"B{rr}", rot, merge=f"B{rr}:D{rr}", h="right")
        calc(ws, f"E{rr}", f, fmt=fmt)
        hint(ws, f"F{rr}", dica, merge=f"F{rr}:I{rr}", h="left")
        ws.row_dimensions[rr].height = 21.75
    rr = r0 + NRES
    label(ws, f"B{rr}", "O que fazer", merge=f"B{rr}:D{rr}", h="right")
    calc(ws, f"E{rr}",
         f'=IF({SIT}="","Preencha as medições de upload e a configuração",'
         f'IF({SIT}="OK",IF({SITB}="Abaixo da faixa","A rede comporta, mas o bitrate está baixo para o formato: suba o bitrate ou reduza o formato",'
         f'IF({SITB}="Acima da faixa","Bitrate acima da faixa do formato: pode baixar sem perder imagem","Manter")),'
         f'IF(AND({FMIN}<>"",{VMAX}<{FMIN}),"Este formato não cabe neste upload: reduza a resolução ou os quadros. Veja a tabela abaixo",'
         f'"Baixe o bitrate de vídeo para "&{VMAX}&" kbps ou menos")))',
         b=False, merge=f"E{rr}:I{rr}", h="left")
    ws.row_dimensions[rr].height = 24
    cf_equal(ws, SIT.replace("$", ""), REDE_CF)
    cf_equal(ws, SITR.replace("$", ""), REDE_CF)
    cf_status(ws, SITB.replace("$", ""), SITB)
    cf_equal(ws, f"E{rr}:I{rr}", [("Manter", GREEN)])
    ws.conditional_formatting.add(f"E{rr}:I{rr}", FormulaRule(
        formula=[f'AND({SIT}<>"",$E${rr}<>"Manter")'], fill=PatternFill("solid", bgColor=YELLOW, fgColor=YELLOW)))
    if not is_ex:
        note(ws, f"B{r0 + 5}", "A fração máxima vem da seção 2 (50% como regra prática).\nEntre a fração e 75%, a transmissão funciona no teste e trava quando a rede oscila.")

    # ---------------- 4. formatos
    band(ws, t0, "4. Qual formato cabe neste upload", "I")
    for col, t, cor in [("B", "#", INK), ("C", "Formato", INK), ("D", "Mínimo", INK), ("E", "Máximo", INK), ("F", "Cabe?", REDH),
                        ("G", "Vídeo sugerido", REDH), ("H", "Uso do upload", REDH), ("I", "Observação", INK)]:
        head(ws, f"{col}{t0 + 1}", t, bg=cor)
    ws.row_dimensions[t0 + 1].height = 24
    for k, (fo, mn, mx, obs) in enumerate(FORMATOS):
        rr = T1 + k
        num(ws, f"B{rr}", k + 1)
        put(ws, f"C{rr}", fo, f=font(10, True), bg=GRAY, h="center")
        inp(ws, f"D{rr}", mn, h="center", bg=bg, fmt=KBPS)
        inp(ws, f"E{rr}", mx, h="center", bg=bg, fmt=KBPS)
        calc(ws, f"F{rr}", f'=IF(OR({VMAX}="",D{rr}=""),"",IF({VMAX}>=D{rr},"Cabe","Não cabe"))')
        calc(ws, f"G{rr}", f'=IF(F{rr}="Cabe",MIN(E{rr},ROUND({VMAX}/100-0.5,0)*100),"")', fmt=KBPS)
        calc(ws, f"H{rr}", f'=IF(OR(G{rr}="",{UP}=""),"",(G{rr}+{AUD})/({UP}*1000))', fmt="0%", b=False)
        put(ws, f"I{rr}", obs, bg=GRAY)
        ws.row_dimensions[rr].height = 21.75
    dv_number(ws, f"D{T1}:E{T2}")
    cf_equal(ws, f"F{T1}:F{T2}", [("Cabe", GREEN), ("Não cabe", RED)])
    hint(ws, f"B{T2 + 1}", "Faixas em kbps para o codec H.264, como ponto de partida deste material. Troque o mínimo e o máximo pelos valores "
                           "que a sua plataforma publica. O vídeo sugerido é o maior valor que cabe, até o máximo da faixa.",
         merge=f"B{T2 + 1}:I{T2 + 1}", h="left")
    ws.row_dimensions[T2 + 1].height = 30
    if not is_ex:
        note(ws, f"C{t0 + 1}", "Resolução e quadros por segundo, sem espaço: 720p30 é 1280 × 720 a 30 quadros.\nA seção 3 procura nesta coluna o formato escolhido na seção 2.")
    end = T2 + 1

    if is_ex:
        band(ws, end + 2, "O que foi ajustado", "I")
        rr = end + 3
        put(ws, f"B{rr}", "Etapa", f=font(10, True), bg=GRAY, h="center", merge=f"B{rr}:C{rr}")
        put(ws, f"D{rr}", "O que foi feito", f=font(10, True), bg=GRAY, h="center", merge=f"D{rr}:H{rr}")
        put(ws, f"I{rr}", "Resultado", f=font(10, True), bg=GRAY, h="center")
        ws.row_dimensions[rr].height = 21.75
        for etapa, feito, resultado in data["ajustes"]:
            rr += 1
            put(ws, f"B{rr}", etapa, f=font(10, True), bg=GRAY, merge=f"B{rr}:C{rr}")
            put(ws, f"D{rr}", feito, merge=f"D{rr}:H{rr}")
            put(ws, f"I{rr}", resultado)
            ws.row_dimensions[rr].height = 24
        end = rr
    ws.freeze_panes = "B4"
    setup(ws, tab, f"B1:I{end}", fit_height=True)


# ================================================================== aba Configuração
# grupo, parâmetro, referência, tipo, opções (ou as opções aceitas)
CONFIG = [
    ("Plataforma", "Plataforma e canal", "Onde o evento vai ao ar.", "texto", None),
    ("Plataforma", "Evento criado, com título, data e capa", "Na véspera.", "simnao", None),
    ("Plataforma", "Visibilidade", "Teste: não listado ou privado. Evento: público.", "lista", ["Público", "Não listado", "Privado"]),
    ("Plataforma", "Endereço do servidor (URL) conferido", "Copiado da página do evento para o codificador.", "simnao", None),
    ("Plataforma", "Chave de transmissão conferida", "É a do evento de hoje. Não anote a chave nesta planilha.", "simnao", None),
    ("Vídeo", "Resolução de saída", "720p ou 1080p: o que couber no upload.", "lista", ["480p", "720p", "1080p"]),
    ("Vídeo", "Quadros por segundo", "30. Use 60 só com rede e processador folgados.", "lista", ["30", "60"]),
    ("Vídeo", "Bitrate de vídeo (kbps)", "O valor que deu OK na aba Bitrate e upload.", "numero", None),
    ("Vídeo", "Codec de vídeo", "H.264, aceito por todas as plataformas.", "lista_ok", (["H.264", "H.265", "AV1"], ["H.264"])),
    ("Vídeo", "Controle de taxa", "CBR (taxa constante).", "lista_ok", (["CBR", "VBR"], ["CBR"])),
    ("Vídeo", "Intervalo de keyframe (s)", "2 segundos. Não deixe em automático.", "keyframe", None),
    ("Vídeo", "Compressão feita por", "Placa de vídeo, quando houver. Alivia o processador.", "lista", ["Placa de vídeo", "Processador", "Aparelho dedicado"]),
    ("Áudio", "Origem do áudio", "Mistura própria: auxiliar, matriz ou mesa dedicada.", "lista_ok",
     (["Master LR", "Auxiliar", "Matriz", "Mesa dedicada", "Câmera"], ["Auxiliar", "Matriz", "Mesa dedicada"])),
    ("Áudio", "Taxa de amostragem (kHz)", "48 kHz na mesa, na interface e no codificador.", "amostragem", None),
    ("Áudio", "Bitrate de áudio (kbps)", "De 128 a 160 kbps, em AAC.", "audio", None),
    ("Áudio", "Canais", "Estéreo, com som nos dois lados. Fonte de um cabo só: mono.", "lista", ["Estéreo", "Mono"]),
    ("Áudio", "Loudness medido (LUFS)", "Alvo de −16 a −14 LUFS. A planilha aceita de −18 a −13.", "loudness", None),
    ("Áudio", "Pico máximo medido (dBFS)", "Abaixo de −1 dBFS no trecho mais forte.", "pico", None),
    ("Áudio", "Teste da palma: quadros de diferença", "Quadros entre o estalo no áudio e as mãos se tocando.", "numero", None),
    ("Áudio", "Atraso calculado para o áudio (ms)", "Quadros × 1.000 ÷ quadros por segundo.", "calc", None),
    ("Áudio", "Atraso configurado no áudio (ms)", "Perto do calculado: até meio quadro de diferença.", "atraso", None),
    ("Rede e segurança", "Conexão do codificador", "Cabo de rede.", "lista_ok", (["Cabo", "Wi-Fi", "4G ou 5G"], ["Cabo"])),
    ("Rede e segurança", "Link de reserva testado", "Com o perfil de bitrate reduzido.", "simnao", None),
    ("Rede e segurança", "Reconexão automática ligada", "Para voltar sozinho depois de uma queda de rede.", "simnao", None),
    ("Rede e segurança", "Gravação local ligada", "Com espaço em disco conferido.", "simnao", None),
    ("Rede e segurança", "Retorno em outro aparelho", "Celular com dados móveis e fone.", "simnao", None),
]
C_W = {"A": 2, "B": 5, "C": 18, "D": 38, "E": 20, "F": 54, "G": 14, "H": 34, "I": 2}
K1 = 10
FALHAS = ["Pendente", "Rever", "Baixo", "Alto", "Clip"]


def aba_config(ws, tab):
    widths(ws, C_W)
    title(ws, "Configuração — a ficha da transmissão",
          "Anote o que está configurado no codificador e na plataforma. A situação é calculada nos itens que têm referência.", "H")
    for r, rot in [(4, "Evento"), (5, "Data"), (6, "Operador")]:
        label(ws, f"B{r}", rot, merge=f"B{r}:C{r}")
        inp(ws, f"D{r}", merge=f"D{r}:E{r}", fmt=DATE if r == 5 else None)
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, "D5")
    hint(ws, "F4", "Preencha na véspera, depois do teste. No dia do evento, confira se o codificador ainda mostra os mesmos valores.",
         merge="F4:H6", h="left")

    band(ws, 8, "Ficha de configuração", "H")
    for col, t, cor in [("B", "#", INK), ("C", "Grupo", INK), ("D", "Parâmetro", INK), ("E", "Valor", REDH),
                        ("F", "Ponto de partida", INK), ("G", "Situação", TEAL), ("H", "Observação", INK)]:
        head(ws, f"{col}9", t, bg=cor)
    ws.row_dimensions[9].height = 24
    note(ws, "G9", "OK: confere com o ponto de partida.\nPendente: marcado como Não.\nRever, Baixo, Alto ou Clip: o valor está fora do ponto de partida.")
    linha = {p: K1 + k for k, (_, p, _, _, _) in enumerate(CONFIG)}
    FPS = f"E{linha['Quadros por segundo']}"
    QD = f"E{linha['Teste da palma: quadros de diferença']}"
    AT = f"E{linha['Atraso calculado para o áudio (ms)']}"
    for k, (grupo, par, ref, tipo, op) in enumerate(CONFIG):
        r = K1 + k
        E = f"E{r}"
        num(ws, f"B{r}", k + 1)
        put(ws, f"C{r}", grupo, bg=GRAY)
        put(ws, f"D{r}", par, f=font(10, True), bg=GRAY)
        put(ws, f"F{r}", ref, bg=GRAY)
        sit = None
        if tipo == "calc":
            calc(ws, E, f'=IFERROR(IF(OR({QD}="",{FPS}=""),"",ROUND({QD}*1000/{FPS},0)),"")')
        else:
            inp(ws, E, h="center")
        if tipo == "simnao":
            dv_list(ws, E, SIM_NAO, "Sim ou Não.")
            sit = f'=IF({E}="","",IF({E}="Sim","OK","Pendente"))'
        elif tipo == "lista":
            dv_list(ws, E, op, "Escolha uma das opções.")
        elif tipo == "lista_ok":
            dv_list(ws, E, op[0], "Escolha uma das opções.")
            aceitas = ",".join(f'{E}="{o}"' for o in op[1])
            sit = f'=IF({E}="","",IF(OR({aceitas}),"OK","Rever"))'
        elif tipo in ("numero", "keyframe", "amostragem", "audio", "loudness", "pico", "atraso"):
            dv_number(ws, E)
            sit = {
                "numero": None,
                "keyframe": f'=IF({E}="","",IF({E}=2,"OK","Rever"))',
                "amostragem": f'=IF({E}="","",IF({E}=48,"OK","Rever"))',
                "audio": f'=IF({E}="","",IF({E}<128,"Baixo","OK"))',
                "loudness": f'=IF({E}="","",IF({E}<-18,"Baixo",IF({E}>-13,"Alto","OK")))',
                "pico": f'=IF({E}="","",IF({E}>=0,"Clip",IF({E}>-1,"Alto","OK")))',
                "atraso": f'=IF(OR({E}="",{AT}=""),"",IF(ABS({E}-{AT})<=IFERROR(500/{FPS},17),"OK","Rever"))',
            }[tipo]
        if sit:
            calc(ws, f"G{r}", sit)
        else:
            hint(ws, f"G{r}", "–")
        inp(ws, f"H{r}")
        ws.row_dimensions[r].height = 21.75
    K2 = K1 + len(CONFIG) - 1
    ws[f"E{linha['Loudness medido (LUFS)']}"].number_format = '+0.0;-0.0;0.0'
    ws[f"E{linha['Pico máximo medido (dBFS)']}"].number_format = '+0.0;-0.0;0.0'
    cf_equal(ws, f"G{K1}:G{K2}", [("OK", GREEN), ("Pendente", YELLOW), ("Rever", YELLOW), ("Baixo", YELLOW), ("Alto", YELLOW), ("Clip", RED)])

    s = K2 + 2
    band(ws, s, "Resumo automático", "H")
    falhas = "+".join(f'COUNTIF(G{K1}:G{K2},"{t}")' for t in FALHAS)
    summary(ws, s + 1, "Itens com situação OK", f'=COUNTIF(G{K1}:G{K2},"OK")', "D", "E")
    summary(ws, s + 2, "Itens a rever ou pendentes", f"={falhas}", "D", "E")
    label(ws, f"B{s + 3}", "Aviso", merge=f"B{s + 3}:D{s + 3}", h="right")
    calc(ws, f"E{s + 3}", f'=IF(E{s + 1}+E{s + 2}=0,"Preencha a ficha",IF(E{s + 2}>0,"Há itens a resolver antes de transmitir","OK"))',
         b=False, merge=f"E{s + 3}:H{s + 3}", h="left")
    ws.row_dimensions[s + 3].height = 21.75
    cf_ok(ws, f"E{s + 3}:H{s + 3}")
    hint(ws, f"B{s + 4}", "Os pontos de partida são valores de referência deste material, e não norma. Confira a tabela da sua plataforma e o manual do codificador.",
         merge=f"B{s + 4}:H{s + 4}", h="left")
    ws.freeze_panes = f"B{K1}"
    setup(ws, tab, f"B1:H{s + 4}", fit_height=True)


# ================================================================== aba Quedas
Q_W = {"A": 2, "B": 5, "C": 12, "D": 9, "E": 9, "F": 10, "G": 20, "H": 18, "I": 46, "J": 46, "K": 2}
QN = 20
OCORRENCIAS = ["Queda total", "Travamento", "Sem áudio", "Sem vídeo", "Fora de sincronismo", "Imagem ruim", "Som ruim", "Outro"]
CAUSAS = ["Internet", "Codificador", "Computador", "Áudio", "Vídeo", "Energia", "Plataforma", "Operação", "Não identificada"]


def aba_quedas(ws, tab):
    widths(ws, Q_W)
    title(ws, "Quedas — registro das falhas da transmissão",
          "Anote cada falha na hora em que acontece, ou logo depois. O resumo mostra quanto tempo a transmissão ficou com problema, e por quê.", "J")
    label(ws, "B4", "Evento", merge="B4:C4")
    inp(ws, "D4", merge="D4:H4")
    label(ws, "B5", "Operador", merge="B5:C5")
    inp(ws, "D5", merge="D5:H5")
    label(ws, "B6", "Duração da transmissão (min)", merge="B6:F6")
    inp(ws, "G6", h="center", fmt="0")
    dv_number(ws, "G6")
    hint(ws, "H6", "Do início ao fim, com a tela de espera.", merge="H6:I6", h="left")
    for r in (4, 5, 6):
        ws.row_dimensions[r].height = 21.75

    band(ws, 8, "Registro de ocorrências", "J")
    for col, t, cor in [("B", "#", INK), ("C", "Data", INK), ("D", "Início", REDH), ("E", "Fim", REDH), ("F", "Duração", TEAL),
                        ("G", "O que aconteceu", INK), ("H", "Causa provável", INK), ("I", "O que foi feito", INK),
                        ("J", "Como evitar na próxima", INK)]:
        head(ws, f"{col}9", t, bg=cor)
    ws.row_dimensions[9].height = 24
    for col, t in zip("BCDEFGHIJ", ["", "dia/mês/ano", "hh:mm", "hh:mm", "min", "lista", "lista", "", ""]):
        hint(ws, f"{col}10", t)
    for col, v, fmt in [("B", "Ex.", None), ("C", date(2026, 10, 4), DATE), ("D", 0.8208333333, HORA), ("E", 0.8229166667, HORA),
                        ("F", 3, None), ("G", "Queda total", None), ("H", "Internet", None),
                        ("I", "Troca para o link de reserva, com o perfil de 720p", None),
                        ("J", "Separar o Wi-Fi dos visitantes da rede da transmissão", None)]:
        ex(ws, f"{col}11", v, h="left" if col in "GHIJ" else "center", fmt=fmt)
    ws.row_dimensions[11].height = 24
    note(ws, "D9", "Hora em que o problema começou, no formato 19:42.")
    note(ws, "F9", "Minutos entre o início e o fim. Se a queda passar da meia-noite, a conta continua certa.")
    Q1 = 12
    Q2 = Q1 + QN - 1
    for k in range(QN):
        r = Q1 + k
        num(ws, f"B{r}", k + 1)
        inp(ws, f"C{r}", h="center", fmt=DATE)
        inp(ws, f"D{r}", h="center", fmt=HORA)
        inp(ws, f"E{r}", h="center", fmt=HORA)
        calc(ws, f"F{r}", f'=IF(OR(D{r}="",E{r}=""),"",ROUND(IF(E{r}<D{r},E{r}+1-D{r},E{r}-D{r})*1440,0))')
        inp(ws, f"G{r}")
        inp(ws, f"H{r}")
        inp(ws, f"I{r}")
        inp(ws, f"J{r}")
        ws.row_dimensions[r].height = 21.75
    dv_date(ws, f"C{Q1}:C{Q2}")
    dv_list(ws, f"G{Q1}:G{Q2}", OCORRENCIAS, "Escolha o que o espectador viu ou ouviu.")
    dv_list(ws, f"H{Q1}:H{Q2}", CAUSAS, "Escolha a etapa em que a falha nasceu.")

    s = Q2 + 2
    band(ws, s, "Resumo automático", "J")
    summary(ws, s + 1, "Ocorrências registradas", f"=COUNTA(G{Q1}:G{Q2})", "E", "F")
    summary(ws, s + 2, "Minutos com problema", f"=SUM(F{Q1}:F{Q2})", "E", "F")
    summary(ws, s + 3, "Percentual da transmissão", f'=IF(OR(G6="",G6=0),"",F{s + 2}/G6)', "E", "F", fmt="0.0%")
    p1 = s + 6
    p2 = p1 + len(CAUSAS) - 1
    label(ws, f"B{s + 4}", "Causa mais frequente", merge=f"B{s + 4}:E{s + 4}", h="right")
    calc(ws, f"F{s + 4}", f'=IF(MAX(F{p1}:F{p2})=0,"",INDEX(B{p1}:B{p2},MATCH(MAX(F{p1}:F{p2}),F{p1}:F{p2},0)))', merge=f"F{s + 4}:G{s + 4}")
    ws.row_dimensions[s + 4].height = 19.5
    put(ws, f"B{s + 5}", "Por causa", f=font(10, True), bg=GRAY, h="center", merge=f"B{s + 5}:E{s + 5}")
    put(ws, f"F{s + 5}", "Ocorrências", f=font(10, True), bg=GRAY, h="center")
    put(ws, f"G{s + 5}", "Minutos", f=font(10, True), bg=GRAY, h="center")
    for k, c in enumerate(CAUSAS):
        r = p1 + k
        put(ws, f"B{r}", c, bg=GRAY, merge=f"B{r}:E{r}")
        calc(ws, f"F{r}", f'=COUNTIF(H{Q1}:H{Q2},"{c}")', b=False)
        calc(ws, f"G{r}", f'=SUMIF(H{Q1}:H{Q2},"{c}",F{Q1}:F{Q2})', b=False)
    hint(ws, f"H{s + 1}", "Leve este registro para a reunião de revisão do evento. A causa que mais se repete é a primeira a resolver.",
         merge=f"H{s + 1}:J{s + 3}", h="left")
    ws.freeze_panes = f"C{Q1}"
    setup(ws, tab, f"B1:J{p2}", fit_height=True)


# ================================================================== pasta de trabalho
wb = Workbook()
wb.properties.title = "Transmissão ao vivo — modelo de planejamento e registro"
wb.properties.creator = "Estudos: Técnico em áudio e vídeo"
wb.properties.language = "pt-BR"

instrucoes(wb, "Transmissão ao vivo — como usar esta planilha",
           "Conta de bitrate e upload, ficha de configuração, checklist com horários e registro de quedas, para usar em cada transmissão.", [
    ("Legenda: onde preencher", [
        ("Amarelo-claro", "Células de entrada. É aqui que você digita."),
        ("Cinza", "Células calculadas ou fixas (rótulos, contas, avisos). Não altere."),
        ("Cinza em itálico", "Linha de exemplo (marcada com “Ex.”) e dicas de preenchimento. Não entram nas contas."),
        ("Listas suspensas", "Conexão, resolução, quadros, Sim ou Não e as causas aceitam só as opções da lista."),
        ("Números", "Digite só o número: 9,8 para o upload em Mbps, 4500 para o bitrate em kbps, -15 para o loudness."),
        ("Horas", "Digite as horas no formato 19:00."),
        ("Comentários", "Células com triângulo vermelho no canto trazem uma dica. Passe o mouse sobre elas."),
    ]),
    ("Ordem de preenchimento", [
        ("1", "Aba Bitrate e upload: meça o upload no local, por cabo, pelo menos três vezes, e anote."),
        ("2", "Aba Bitrate e upload: digite a configuração pretendida e veja a situação da rede e o que fazer."),
        ("3", "Se não couber, escolha na tabela “Qual formato cabe” um formato com situação Cabe e refaça a conta."),
        ("4", "Aba Configuração: depois do teste da véspera, anote o que ficou configurado e resolva as pendências."),
        ("5", "Aba Checklist: digite o horário de início e siga os itens na ordem, cada um até a hora calculada."),
        ("6", "Aba Quedas: durante a transmissão, registre cada falha, com a hora de início e a de fim."),
        ("7", "Depois do evento, leve o resumo da aba Quedas para a reunião de revisão."),
    ]),
    ("Abas da planilha", [
        ("Bitrate e upload", "Medições de upload, configuração pretendida, situação da rede, consumo de dados e quais formatos cabem."),
        ("Configuração", "A ficha da transmissão, em 26 itens, com o cálculo do atraso de sincronismo pelo teste da palma."),
        ("Checklist", "Os doze itens do treinamento, com a hora limite de cada um e o percentual atendido."),
        ("Quedas", "Registro de até 20 ocorrências, com duração, tempo total com problema e causa mais frequente."),
        ("Referência", "Faixas de bitrate por formato, ajustes do codificador, áudio da transmissão e tabela de sintomas."),
        ("Exemplo - Bitrate", "A igreja em que a transmissão travava com a sala cheia (exemplo 1 do treinamento)."),
    ]),
    ("Regras e premissas do modelo", [
        ("Regra da metade", "O bitrate total (vídeo mais áudio) deve usar até cerca de metade do menor upload medido. É regra prática, e não norma: a fração fica em uma célula e pode ser trocada."),
        ("Situação da rede", "OK até a fração máxima. Apertado entre a fração e 75% do upload. Insuficiente acima de 75%."),
        ("Faixas de bitrate", "Valores de partida para o codec H.264. Cada plataforma publica a sua tabela: troque o mínimo e o máximo na tabela de formatos, no fim da aba."),
        ("Consumo de dados", "Bitrate total × duração. Uma transmissão de 4.600 kbps consome cerca de 2 GB por hora."),
        ("Sincronismo", "Atraso do áudio = quadros de diferença × 1.000 ÷ quadros por segundo. A 30 quadros, cada quadro vale 33 ms."),
        ("Áudio", "Loudness de −16 a −14 LUFS e picos abaixo de −1 dBFS são ponto de partida para a internet. Confira a orientação da sua plataforma."),
        ("Chave de transmissão", "A planilha não tem campo para a chave, de propósito. A chave é uma senha e não deve ficar em arquivo compartilhado."),
        ("Exemplos", "Os dados da aba de exemplo e das linhas “Ex.” são ilustrativos, criados para este material."),
    ]),
])

aba_bitrate(wb.create_sheet("Bitrate e upload"), REDH)
aba_config(wb.create_sheet("Configuração"), BLUE)

# ------------------------------------------------------------------ Checklist com horários
ws = checklist(wb.create_sheet("Checklist"), "Checklist da transmissão, com horários",
               "Digite o horário de início. Cada item deve estar resolvido até a hora calculada. Escolha a situação na lista suspensa.",
               [(q, t) for q, _, t in CHECK_T], tab=TEAL, primeira=8)
widths(ws, {"G": 12, "H": 12, "I": 2})
for r in (1, 2):
    ws.unmerge_cells(f"B{r}:F{r}")
title(ws, ws["B1"].value, ws["B2"].value, "H")
label(ws, "B4", "Evento", merge="B4:C4")
inp(ws, "D4")
label(ws, "B5", "Início da transmissão", merge="B5:C5")
inp(ws, "D5", h="left", fmt=HORA)
hint(ws, "E5", "Digite a hora em que o evento começa, como 19:00. A coluna “Fazer até” é calculada.", merge="E5:H5", h="left")
for r in (4, 5):
    ws.row_dimensions[r].height = 21.75
ws.cell(row=7, column=3).value = "Quando"
head(ws, "G7", "Min. antes", bg=AMBER)
head(ws, "H7", "Fazer até", bg=TEAL)
note(ws, "G7", "Minutos antes do início em que o item deve estar resolvido.\nPode ser trocado conforme o seu evento.")
for k, (_, minutos, _) in enumerate(CHECK_T):
    r = 8 + k
    if minutos is None:
        hint(ws, f"G{r}", "–")
        put(ws, f"H{r}", "na véspera", f=font(10, True), bg=GRAY, h="center")
    else:
        inp(ws, f"G{r}", minutos, h="center", fmt="0")
        calc(ws, f"H{r}", f'=IF(OR($D$5="",G{r}=""),"",IF($D$5-G{r}/1440<0,$D$5-G{r}/1440+1,$D$5-G{r}/1440))', fmt=HORA)
dv_number(ws, f"G8:G{7 + len(CHECK_T)}")
setup(ws, TEAL, f"B1:H{7 + len(CHECK_T) + 4}")

aba_quedas(wb.create_sheet("Quedas"), AMBER)

# ------------------------------------------------------------------ Referência
ws = wb.create_sheet("Referência")
widths(ws, {"A": 2, "B": 36, "C": 30, "D": 30, "E": 62, "F": 2})
title(ws, "Referência — valores de partida", "Valores de referência, para servir de ponto de partida. Confira a tabela da sua plataforma e o manual do seu equipamento.", "E")


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


r = tabela(4, "Bitrate de vídeo por formato (codec H.264)", ["Formato", "Bitrate de vídeo", "Upload medido para caber", "Quando usar"], [
    ("480p a 30 quadros (854 × 480)", "1.000 a 2.000 kbps", "2,5 a 4,5 Mbps", "Rede fraca ou link de reserva. Serve para fala; letras pequenas ficam ilegíveis."),
    ("720p a 30 quadros (1280 × 720)", "2.500 a 4.000 kbps", "5,5 a 8,5 Mbps", "O formato seguro para a maioria dos eventos pequenos. Fica bem em celular e em TV."),
    ("720p a 60 quadros (1280 × 720)", "3.500 a 6.000 kbps", "7,5 a 12,5 Mbps", "Movimento rápido, como dança e esporte. Raro em culto e palestra."),
    ("1080p a 30 quadros (1920 × 1080)", "4.000 a 6.000 kbps", "8,5 a 12,5 Mbps", "O padrão quando a rede comporta. Mais detalhe em planos abertos e em letras."),
    ("1080p a 60 quadros (1920 × 1080)", "6.000 a 9.000 kbps", "12,5 a 18,5 Mbps", "Só com rede e processador com muita folga."),
])
r = tabela(r, "Ajustes do codificador", ["Ajuste", "Ponto de partida", "Alternativa", "Observação"], [
    ("Codec de vídeo", "H.264", "H.265 ou AV1", "Os codecs mais novos só onde a plataforma e o codificador aceitam."),
    ("Controle de taxa", "CBR (taxa constante)", "VBR (taxa variável)", "Ao vivo, use CBR: a rede e a plataforma lidam melhor com um fluxo previsível."),
    ("Intervalo de keyframe", "2 segundos", "até 4 segundos, se a plataforma aceitar", "Não deixe em automático. A 30 quadros por segundo, 2 segundos são 60 quadros."),
    ("Taxa de quadros", "30", "60", "60 quadros dobram o trabalho do codificador e pedem mais bitrate."),
    ("Áudio", "AAC, estéreo, 128 a 160 kbps", "–", "Taxa de amostragem de 48 kHz, igual na mesa, na interface e no codificador."),
    ("Bitrate total", "até metade do menor upload medido", "–", "Vídeo mais áudio. Regra prática deste material, e não norma."),
])
r = tabela(r, "Áudio da transmissão", ["Medida", "Ponto de partida", "Onde conferir", "Observação"], [
    ("Origem", "auxiliar pós-fader, matriz ou mesa dedicada", "fone fechado, fora da sala", "Com microfones de ambiente só na transmissão, nunca nas caixas."),
    ("Picos", "abaixo de −1 dBFS", "medidor do codificador", "Um limitador com teto entre −2 e −1 dBFS garante isso."),
    ("Fala e música no medidor", "entre −20 e −6 dBFS na maior parte do tempo", "medidor do codificador", "Se o medidor quase não sai do começo da escala, a transmissão está baixa."),
    ("Loudness", "de −16 a −14 LUFS", "medidor de loudness", "Medido ao longo de um ou dois minutos de programa. Ponto de partida para a internet."),
    ("Sincronismo", "som e imagem no mesmo quadro", "teste da palma, na gravação do codificador", "Um quadro vale 33 ms a 30 quadros por segundo, e 17 ms a 60."),
])
r = tabela(r, "Sintomas e primeira ação", ["Sintoma", "Causa provável", "Onde olhar", "Primeira ação"], [
    ("Trava para todos os espectadores", "Rede", "quadros perdidos por rede", "Passe ao link de reserva ou baixe o bitrate."),
    ("Trava, e o processador está no limite", "Codificador ou computador", "quadros perdidos na codificação", "Feche outros programas. Na próxima pausa, reduza quadros ou resolução."),
    ("Blocos quando há movimento", "Bitrate baixo para o formato", "configuração do codificador", "Use um formato menor, ou mais bitrate, se a rede comportar."),
    ("Transmissão sem som", "Áudio", "medidor do codificador", "Confira o mute, o dispositivo escolhido, a interface, o cabo e a saída da mesa."),
    ("Som em um lado só", "Fonte mono em entrada estéreo", "fone, no retorno", "Marque a fonte como mono no codificador, ou envie os dois lados."),
    ("Som antes da imagem", "O vídeo atrasa mais que o áudio", "gravação do codificador", "Faça o teste da palma e acrescente atraso ao áudio."),
    ("Página do evento em “aguardando”", "Chave ou evento errado", "painel da plataforma", "Confira a chave do codificador e o evento que está aberto."),
    ("Só alguns espectadores reclamam", "Internet ou aparelho do espectador", "retorno no celular", "Se o retorno está bom, avise no chat e siga."),
])
hint(ws, f"B{r - 1}", "Faixas típicas, que variam com a plataforma, o conteúdo e o equipamento. Os nomes dos indicadores mudam de um programa para outro.",
     merge=f"B{r - 1}:E{r - 1}", h="left")
setup(ws, HUE["slate"][0], f"B1:E{r - 1}", fit_height=True)

# ------------------------------------------------------------------ Exemplo
aba_bitrate(wb.create_sheet("Exemplo - Bitrate"), MUTED, EX_BITRATE)

wb.active = 1
wb.save(OUT)
print("ok", OUT)
