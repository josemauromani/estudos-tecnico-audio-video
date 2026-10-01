# -*- coding: utf-8 -*-
"""Gera as figuras (curvas calculadas) e monta _geradores/eq_body.html a partir de eq_body.tpl.html."""
import math, os, sys
AQUI = os.path.dirname(os.path.abspath(__file__))
DEST = "/home/josemauro/projects/josemauro/estudos_tecnico_audio_video/_geradores/eq_body.html"


def bell(f0, g, Q):
    A = 10 ** (g / 40)
    return lambda f: (lambda s: (s * s + s * A / Q + 1) / (s * s + s / (A * Q) + 1))(1j * f / f0)


def lsh(f0, g, Q=0.707):
    A = 10 ** (g / 40); r = math.sqrt(A) / Q
    return lambda f: (lambda s: A * (s * s + r * s + A) / (A * s * s + r * s + 1))(1j * f / f0)


def hsh(f0, g, Q=0.707):
    A = 10 ** (g / 40); r = math.sqrt(A) / Q
    return lambda f: (lambda s: A * (A * s * s + r * s + 1) / (s * s + r * s + A))(1j * f / f0)


def hpf(f0):
    return lambda f: (lambda s: s * s / (s * s + math.sqrt(2) * s + 1))(1j * f / f0)


def chain(*fs):
    def h(f):
        v = 1
        for x in fs:
            v *= x(f)
        return v
    return h


def db(h, f):
    return 20 * math.log10(max(abs(h(f)), 1e-9))


def n(v):
    return ("%.1f" % v).rstrip("0").rstrip(".")


def rdp(p, tol=0.3):
    """Simplifica a polilinha (Ramer-Douglas-Peucker): tira pontos que não mudam o desenho."""
    if len(p) < 3:
        return p
    (x1, y1), (x2, y2) = p[0], p[-1]
    L = math.hypot(x2 - x1, y2 - y1) or 1e-9
    dmax, idx = 0, 0
    for i in range(1, len(p) - 1):
        d = abs((x2 - x1) * (y1 - p[i][1]) - (x1 - p[i][0]) * (y2 - y1)) / L
        if d > dmax:
            dmax, idx = d, i
    if dmax <= tol:
        return [p[0], p[-1]]
    return rdp(p[:idx + 1], tol)[:-1] + rdp(p[idx:], tol)


class Plot:
    def __init__(s, x0, y0, w, h, dmin, dmax, fmin=20, fmax=20000):
        s.x0, s.y0, s.w, s.h, s.dmin, s.dmax, s.fmin, s.fmax = x0, y0, w, h, dmin, dmax, fmin, fmax

    def X(s, f):
        return s.x0 + s.w * math.log10(f / s.fmin) / math.log10(s.fmax / s.fmin)

    def Y(s, d):
        d = max(s.dmin, min(s.dmax, d))
        return s.y0 + s.h * (s.dmax - d) / (s.dmax - s.dmin)

    def pts(s, h, f1=None, f2=None, k=400):
        f1, f2 = f1 or s.fmin, f2 or s.fmax
        out = []
        for i in range(k + 1):
            f = f1 * (f2 / f1) ** (i / k)
            out.append((s.X(f), s.Y(db(h, f))))
        return rdp(out)

    def path(s, h, f1=None, f2=None, k=400):
        p = s.pts(h, f1, f2, k)
        return "M" + " L".join(f"{n(x)} {n(y)}" for x, y in p)

    def area(s, h, f1=None, f2=None, k=400):
        p = s.pts(h, f1, f2, k)
        z = s.Y(0)
        return f"M{n(p[0][0])} {n(z)} L" + " L".join(f"{n(x)} {n(y)}" for x, y in p) + f" L{n(p[-1][0])} {n(z)} Z"

    def vgrid(s, freqs, labels=True, dy=16, cls="grid"):
        d = " ".join(f"M{n(s.X(f))} {s.y0}V{s.y0 + s.h}" for f in freqs)
        o = [f'<path class="{cls}" d="{d}"/>']
        if labels:
            o.append(f'<g font-size="10.5" text-anchor="middle">' + "".join(
                f'<text class="mu" x="{n(s.X(f))}" y="{s.y0 + s.h + dy}">{flab(f)}</text>' for f in freqs) + "</g>")
        return "\n        ".join(o)

    def hgrid(s, dbs, unit=" dB"):
        d = " ".join(f"M{s.x0} {n(s.Y(v))}H{s.x0 + s.w}" for v in dbs if v != 0)
        o = [f'<path class="grid" d="{d}"/>']
        if 0 in dbs:
            o.append(f'<line class="ln-mu" stroke-width="1" x1="{s.x0}" y1="{n(s.Y(0))}" x2="{s.x0 + s.w}" y2="{n(s.Y(0))}"/>')
        t = []
        for v in dbs:
            lab = ("+" if v > 0 else "−" if v < 0 else "") + str(abs(v)) + (unit if v == 0 else "")
            t.append(f'<text class="{"b" if v == 0 else "mu"}" x="{s.x0 - 8}" y="{n(s.Y(v) + 4)}">{lab}</text>')
        o.append('<g font-size="10.5" text-anchor="end">' + "".join(t) + "</g>")
        return "\n        ".join(o)


def flab(f):
    if f >= 1000:
        v = f / 1000
        return (("%g" % v).replace(".", ",")) + " k"
    return ("%g" % f).replace(".", ",")


F10 = [20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000]

# ------------------------------------------------------------------ Figura 1
p = Plot(70, 46, 800, 190, -18, 12)
h1 = chain(hpf(100), bell(300, -4, 1.4), bell(3000, 3, 1.0))
FIG1 = f'''<svg viewBox="0 0 900 290" role="img" aria-label="Gráfico de uma curva de equalização de voz falada. O eixo horizontal é a frequência, de 20 Hz, no grave, a 20 kHz, no agudo. O eixo vertical é o ganho, de menos 18 a mais 12 dB, com a linha de 0 dB no meio: acima dela o equalizador realça, abaixo dela ele corta. A curva mostra três ajustes. O passa-altas, em 100 Hz, derruba tudo o que fica abaixo da voz e tira o ronco. Um corte de 4 dB em 300 Hz tira o som embolado. Um realce suave de 3 dB em 3 kHz dá clareza às consoantes. No resto da faixa, a curva fica sobre a linha de 0 dB: o som passa sem mudança.">
        <text class="mono mu" x="18" y="26" font-size="10.5">UM EQ DE VOZ FALADA, COMO EXEMPLO</text>
        <rect class="bx" x="{p.x0}" y="{p.y0}" width="{p.w}" height="{p.h}"/>
        {p.vgrid(F10)}
        {p.hgrid([6, 0, -6, -12])}
        <path class="lv-purple" d="{p.area(h1)}"/>
        <path class="s-purple" stroke-width="2.5" d="{p.path(h1)}"/>
        <text class="t-purple b" x="{n(p.X(21.5))}" y="{n(p.Y(8.6))}" font-size="12">Passa-altas em 100 Hz</text>
        <text class="mu" x="{n(p.X(21.5))}" y="{n(p.Y(8.6) + 15)}" font-size="11">tira o ronco que fica abaixo da voz</text>
        <text class="t-purple b" x="{n(p.X(330))}" y="{n(p.Y(-8.5))}" font-size="12" text-anchor="middle">Corte em 300 Hz</text>
        <text class="mu" x="{n(p.X(330))}" y="{n(p.Y(-8.5) + 15)}" font-size="11" text-anchor="middle">tira o som embolado</text>
        <text class="t-purple b" x="{n(p.X(3000))}" y="{n(p.Y(8.6))}" font-size="12" text-anchor="middle">Realce suave em 3 kHz</text>
        <text class="mu" x="{n(p.X(3000))}" y="{n(p.Y(8.6) + 15)}" font-size="11" text-anchor="middle">dá clareza às consoantes</text>
        <text class="mu" x="{n(p.X(15000))}" y="{n(p.Y(-13.5))}" font-size="11" text-anchor="end">acima de 0 dB, realça; abaixo, corta</text>
        <text class="mu" x="{p.x0}" y="278" font-size="11">grave</text>
        <text class="mu" x="{p.x0 + p.w / 2}" y="278" font-size="11" text-anchor="middle">frequência, em Hz</text>
        <text class="mu" x="{p.x0 + p.w}" y="278" font-size="11" text-anchor="end">agudo</text>
      </svg>'''

# ------------------------------------------------------------------ Figura 2
def X2(f):
    return 30 + 280 * math.log10(f / 20)

B = [20, 60, 250, 500, 2000, 4000, 8000, 20000]
NOMES = ["Subgrave", "Grave", "Médio-grave", "Médio", "Médio-agudo", "Agudo", "Brilho"]
FAIXA = ["20 a 60 Hz", "60 a 250 Hz", "250 a 500 Hz", "500 Hz a 2 kHz", "2 a 4 kHz", "4 a 8 kHz", "8 a 20 kHz"]
COR = ["blue", "blue", "amber", "amber", "amber", "purple", "purple"]
EXC = [("ronco,", "vibração"), ("retumbante,", "pesado demais"), ("embolado,", "papelão"), ("anasalado,", "som de telefone"),
       ("áspero,", "estridente"), ("sibilante,", "cortante"), ("chiado,", "agudo que cansa")]
FAL = [("sem peso", "no bumbo"), ("magro,", "sem corpo"), ("oco,", "sem calor"), ("distante,", "vazio"),
       ("sem clareza", "nas palavras"), ("opaco,", "sem ataque"), ("abafado,", "sem ar")]
o = ['<text class="mono mu" x="30" y="26" font-size="10.5">DO GRAVE AO AGUDO: SETE FAIXAS E COMO CADA UMA SOA</text>']
for a, b, c, t in [(20, 250, "blue", "GRAVES"), (250, 4000, "amber", "MÉDIOS"), (4000, 20000, "purple", "AGUDOS")]:
    o.append(f'<rect class="hd-{c}" x="{n(X2(a) + 1)}" y="40" width="{n(X2(b) - X2(a) - 2)}" height="24"/>')
    o.append(f'<text class="on mono" x="{n((X2(a) + X2(b)) / 2)}" y="56" font-size="10.5" text-anchor="middle">{t}</text>')
o.append('<g font-size="11" text-anchor="middle">')
for i in range(7):
    xa, xb = X2(B[i]), X2(B[i + 1]); xm = n((xa + xb) / 2); wd = n(xb - xa - 2); xl = n(xa + 1)
    o.append(f'<rect class="bx-{COR[i]}" x="{xl}" y="70" width="{wd}" height="42"/>'
             f'<text class="b" x="{xm}" y="88" font-size="11.5">{NOMES[i]}</text>'
             f'<text class="mu" x="{xm}" y="103" font-size="10.5">{FAIXA[i]}</text>')
    for y0, rows in ((142, EXC), (222, FAL)):
        o.append(f'<rect class="bx" x="{xl}" y="{y0}" width="{wd}" height="48"/>'
                 f'<text x="{xm}" y="{y0 + 21}">{rows[i][0]}</text><text x="{xm}" y="{y0 + 36}">{rows[i][1]}</text>')
o.append('</g>')
o.append('<text class="mono mu" x="30" y="134" font-size="10.5">EM EXCESSO, SOA</text>')
o.append('<text class="mono mu" x="30" y="214" font-size="10.5">EM FALTA, SOA</text>')
o.append('<line class="ln-mu" x1="30" y1="282" x2="870" y2="282"/>')
for i, f in enumerate(B):
    x = n(X2(f))
    o.append(f'<line class="ln-mu" stroke-width="1" x1="{x}" y1="282" x2="{x}" y2="288"/>')
    lab = ["20 Hz", "60", "250", "500", "2 k", "4 k", "8 k", "20 kHz"][i]
    anc = "start" if i == 0 else "end" if i == 7 else "middle"
    o.append(f'<text class="mu" x="{x}" y="302" font-size="10.5" text-anchor="{anc}">{lab}</text>')
FIG2 = ('<svg viewBox="0 0 900 312" role="img" aria-label="Mapa das sete faixas de frequência, do grave ao agudo, em três grupos. Graves: subgrave, de 20 a 60 Hz, e grave, de 60 a 250 Hz. Médios: médio-grave, de 250 a 500 Hz; médio, de 500 Hz a 2 kHz; e médio-agudo, de 2 a 4 kHz. Agudos: agudo, de 4 a 8 kHz, e brilho, de 8 a 20 kHz. Para cada faixa, o mapa diz como o som fica quando ela sobra e quando ela falta. Subgrave: em excesso, ronco e vibração; em falta, bumbo sem peso. Grave: retumbante e pesado demais; ou magro e sem corpo. Médio-grave: embolado, som de papelão; ou oco e sem calor. Médio: anasalado, som de telefone; ou distante e vazio. Médio-agudo: áspero e estridente; ou sem clareza nas palavras. Agudo: sibilante e cortante; ou opaco e sem ataque. Brilho: chiado, agudo que cansa; ou abafado e sem ar.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 3
o = []
tipos = [
    ("Passa-altas (HPF)", "corta tudo abaixo do ponto", 100, "100 Hz", [("s-purple", "", hpf(100))],
     ("Uso: tirar ronco, vento e batida", "de mão de quem não precisa de grave.")),
    ("Shelving de graves", "sobe ou desce tudo abaixo", 100, "100 Hz", [("s-purple", "", lsh(100, 9)), ("s-blue", " dash", lsh(100, -9))],
     ("Uso: dar ou tirar peso, sem", "escolher uma nota.")),
    ("Shelving de agudos", "sobe ou desce tudo acima", 10000, "10 kHz", [("s-purple", "", hsh(10000, 9)), ("s-blue", " dash", hsh(10000, -9))],
     ("Uso: dar ou tirar brilho, de", "modo amplo.")),
    ("Bell (sino)", "atua em torno de um ponto", 1000, "1 kHz", [("s-purple", "", bell(1000, 9, 1.4)), ("s-blue", " dash", bell(1000, -9, 1.4))],
     ("Uso: tirar uma sobra ou dar um", "realce em uma região só.")),
]
for k, (tit, sub, f0, lab, curvas, uso) in enumerate(tipos):
    p = Plot(28 + 216 * k, 62, 196, 112, -15, 15)
    o.append(f'<text class="b" x="{p.x0}" y="28" font-size="13">{tit}</text>')
    o.append(f'<text class="mu" x="{p.x0}" y="46" font-size="11">{sub}</text>')
    o.append(f'<rect class="bx" x="{p.x0}" y="{p.y0}" width="{p.w}" height="{p.h}"/>')
    o.append(f'<line class="ln-mu" stroke-width="1" x1="{p.x0}" y1="{n(p.Y(0))}" x2="{p.x0 + p.w}" y2="{n(p.Y(0))}"/>')
    o.append(f'<line class="ref" x1="{n(p.X(f0))}" y1="{p.y0}" x2="{n(p.X(f0))}" y2="{p.y0 + p.h}"/>')
    for cls, dash, h in curvas:
        o.append(f'<path class="{cls}{dash}" d="{p.path(h, k=300)}"/>')
    o.append(f'<text class="b" x="{n(p.X(f0))}" y="{p.y0 + p.h + 15}" font-size="10.5" text-anchor="middle">{lab}</text>')
    if k != 0 and k != 1:
        o.append(f'<text class="mu" x="{p.x0}" y="{p.y0 + p.h + 15}" font-size="10.5">grave</text>')
    if k != 2:
        o.append(f'<text class="mu" x="{p.x0 + p.w}" y="{p.y0 + p.h + 15}" font-size="10.5" text-anchor="end">agudo</text>')
    if k == 0:
        o.append(f'<text class="mu" x="{p.x0 + p.w - 6}" y="{n(p.Y(0) - 6)}" font-size="10.5" text-anchor="end">0 dB</text>')
    o.append(f'<text x="{p.x0}" y="{p.y0 + p.h + 38}" font-size="11">{uso[0]}</text>')
    o.append(f'<text x="{p.x0}" y="{p.y0 + p.h + 53}" font-size="11">{uso[1]}</text>')
o.append('<line class="s-purple" x1="28" y1="256" x2="58" y2="256"/><text x="66" y="260" font-size="11">realce</text>')
o.append('<line class="s-blue dash" x1="128" y1="256" x2="158" y2="256"/><text x="166" y="260" font-size="11">corte</text>')
o.append('<text class="mu" x="872" y="260" font-size="11" text-anchor="end">a linha do meio de cada quadro é 0 dB: nela, o som passa sem mudança</text>')
FIG3 = ('<svg viewBox="0 0 900 272" role="img" aria-label="Quatro quadros, cada um com a curva de um tipo de filtro. No primeiro, o passa-altas, ou HPF: a curva fica em 0 dB nos agudos e despenca abaixo do ponto escolhido, 100 Hz no exemplo; serve para tirar ronco, vento e batida de mão. No segundo, o shelving de graves: abaixo de 100 Hz, a curva sobe para um patamar, no realce, ou desce para um patamar, no corte; serve para dar ou tirar peso. No terceiro, o shelving de agudos: acima de 10 kHz, a curva sobe ou desce para um patamar; serve para dar ou tirar brilho. No quarto, o bell, ou sino: a curva forma um morro, no realce, ou um vale, no corte, em torno de 1 kHz, e volta a 0 dB dos dois lados; serve para tirar uma sobra ou dar um realce em uma região só.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 4
o = ['<text class="mono mu" x="18" y="26" font-size="10.5">MESA ANALÓGICA</text>',
     '<text class="b" x="18" y="46" font-size="12.5">Três bandas, com médio semiparamétrico</text>',
     '<rect class="bx" x="18" y="60" width="130" height="262"/>']
knobs = [(92, "HF", "Agudos: shelving", "ponto fixo, em geral 12 kHz"),
         (142, "MID", "Médios: ganho do bell", "a largura (Q) é fixa"),
         (192, "FREQ", "Médios: frequência", "escolhe o ponto, de 100 Hz a 8 kHz"),
         (242, "LF", "Graves: shelving", "ponto fixo, em geral 80 Hz")]
ang = {"HF": 0, "MID": -50, "FREQ": -100, "LF": 0}
for cy, lab, t1, t2 in knobs:
    o.append(f'<text class="mono mu" x="30" y="{cy + 4}" font-size="10.5">{lab}</text>')
    o.append(f'<circle class="bx-purple" cx="106" cy="{cy}" r="16"/>')
    a = math.radians(ang[lab] - 90)
    o.append(f'<line class="s-purple" x1="106" y1="{cy}" x2="{n(106 + 14 * math.cos(a))}" y2="{n(cy + 14 * math.sin(a))}"/>')
    o.append(f'<line class="ln-mu" stroke-width="1" x1="126" y1="{cy}" x2="164" y2="{cy}"/>')
    o.append(f'<text class="b" x="170" y="{cy - 2}" font-size="11.5">{t1}</text>')
    o.append(f'<text class="mu" x="170" y="{cy + 13}" font-size="10.5">{t2}</text>')
o.append('<text class="mono mu" x="30" y="296" font-size="10.5">HPF</text>')
o.append('<rect class="hd-blue" x="94" y="282" width="24" height="18"/>')
o.append('<line class="ln-mu" stroke-width="1" x1="122" y1="291" x2="164" y2="291"/>')
o.append('<text class="b" x="170" y="289" font-size="11.5">Passa-altas: chave</text>')
o.append('<text class="mu" x="170" y="304" font-size="10.5">liga o corte em 80 ou 100 Hz</text>')
o.append('<text class="mono mu" x="430" y="26" font-size="10.5">MESA DIGITAL</text>')
o.append('<text class="b" x="430" y="46" font-size="12.5">Quatro bandas paramétricas e passa-altas, na tela</text>')
o.append('<rect class="bx" x="430" y="60" width="452" height="156"/>')
p = Plot(446, 74, 420, 112, -14, 10)
h4 = chain(hpf(90), bell(250, -4, 2), bell(800, -3, 3), bell(3000, 3, 1), hsh(9000, 3))
o.append(p.vgrid([100, 1000, 10000], dy=15))
o.append(f'<line class="ln-mu" stroke-width="1" x1="{p.x0}" y1="{n(p.Y(0))}" x2="{p.x0 + p.w}" y2="{n(p.Y(0))}"/>')
o.append(f'<path class="lv-purple" d="{p.area(h4)}"/>')
o.append(f'<path class="s-purple" d="{p.path(h4)}"/>')
for num, f, g in [(1, 250, -4), (2, 800, -3), (3, 3000, 3), (4, 12000, 3)]:
    cy = p.Y(g) + (16 if g < 0 else -16)
    o.append(f'<circle class="f-purple" cx="{n(p.X(f))}" cy="{n(cy)}" r="9"/>')
    o.append(f'<text class="on b" x="{n(p.X(f))}" y="{n(cy + 4)}" font-size="11" text-anchor="middle">{num}</text>')
o.append(f'<rect class="hd-blue" x="{n(p.X(34))}" y="{n(p.Y(6.5))}" width="34" height="18"/>')
o.append(f'<text class="on b" x="{n(p.X(34) + 17)}" y="{n(p.Y(6.5) + 13)}" font-size="10.5" text-anchor="middle">HPF</text>')
o.append('<text class="mu" x="430" y="238" font-size="11">Para a banda escolhida, três controles:</text>')
for cx, lab, t1, t2, a0 in [(470, "FREQ", "Frequência", "onde a banda atua", -60), (620, "GAIN", "Ganho", "quanto corta ou realça", 40), (770, "Q", "Q", "largura da banda", -20)]:
    o.append(f'<circle class="bx-purple" cx="{cx}" cy="284" r="16"/>')
    a = math.radians(a0 - 90)
    o.append(f'<line class="s-purple" x1="{cx}" y1="284" x2="{n(cx + 14 * math.cos(a))}" y2="{n(284 + 14 * math.sin(a))}"/>')
    o.append(f'<text class="mono mu" x="{cx}" y="318" font-size="10.5" text-anchor="middle">{lab}</text>')
    o.append(f'<text class="b" x="{cx + 24}" y="282" font-size="11.5">{t1}</text>')
    o.append(f'<text class="mu" x="{cx + 24}" y="297" font-size="10.5">{t2}</text>')
FIG4 = ('<svg viewBox="0 0 900 336" role="img" aria-label="Dois painéis de equalizador de canal, em esquema. À esquerda, o de uma mesa analógica, com três bandas: o botão HF é um shelving de agudos com ponto fixo, em geral 12 kHz; o botão MID é o ganho de um bell de médios com largura fixa; o botão FREQ escolhe a frequência desse bell, de 100 Hz a 8 kHz; o botão LF é um shelving de graves com ponto fixo, em geral 80 Hz; e a chave HPF liga o passa-altas em 80 ou 100 Hz. À direita, a tela de uma mesa digital, com a curva do equalizador, o passa-altas e quatro bandas numeradas. Para a banda escolhida há três controles: frequência, que define onde a banda atua; ganho, que define quanto ela corta ou realça; e Q, que define a largura da banda.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 5
p = Plot(64, 34, 520, 200, -3, 12)
qs = [("s-blue", 0.7, "Q 0,7: banda larga", "Cerca de 2 oitavas. Ajuste suave de", "timbre: é o Q dos realces."),
      ("s-purple", 1.4, "Q 1,4: banda média", "Cerca de 1 oitava. O corte comum de", "som embolado ou anasalado."),
      ("s-amber", 4.3, "Q 4,3: banda estreita", "Cerca de 1/3 de oitava. Corte de uma", "ressonância ou de uma microfonia.")]
o = [f'<rect class="bx" x="{p.x0}" y="{p.y0}" width="{p.w}" height="{p.h}"/>', p.vgrid(F10), p.hgrid([9, 6, 3, 0]),
     f'<line class="ref" x1="{n(p.X(1000))}" y1="{p.y0}" x2="{n(p.X(1000))}" y2="{p.y0 + p.h}"/>']
for k, (cls, q, t, l1, l2) in enumerate(qs):
    o.append(f'<path class="{cls}" d="{p.path(bell(1000, 9, q), k=400)}"/>')
    y = 62 + 62 * k
    o.append(f'<line class="{cls}" x1="612" y1="{y - 4}" x2="642" y2="{y - 4}"/>')
    o.append(f'<text class="b" x="650" y="{y}" font-size="12">{t}</text>')
    o.append(f'<text class="mu" x="612" y="{y + 18}" font-size="11">{l1}</text>')
    o.append(f'<text class="mu" x="612" y="{y + 33}" font-size="11">{l2}</text>')
o.append(f'<text class="mu" x="{p.x0 + p.w / 2}" y="{p.y0 + p.h + 36}" font-size="11" text-anchor="middle">três realces de 9 dB em 1 kHz: a mesma frequência e o mesmo ganho, com três valores de Q</text>')
FIG5 = ('<svg viewBox="0 0 900 282" role="img" aria-label="Gráfico com três curvas de realce de 9 dB em 1 kHz, com a mesma frequência e o mesmo ganho, e três valores de Q. Com Q 0,7, a banda é larga: o morro se espalha por cerca de duas oitavas, de 500 Hz a 2 kHz, e muda o timbre de modo suave; é o Q dos realces. Com Q 1,4, a banda é média, com cerca de uma oitava; é o corte comum de som embolado ou anasalado. Com Q 4,3, a banda é estreita, com cerca de um terço de oitava; serve para cortar uma ressonância ou uma microfonia.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 6
o = ['<defs><marker id="ar6" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto"><path class="ah f-purple" d="M0 0 L10 5 L0 10 z"/></marker></defs>']
tit = ["1. Realce estreito", "2. Varra devagar", "3. Corte no ponto achado"]
cap = [("Suba de 6 a 9 dB em uma banda bell,", "com Q alto (4 ou mais)."),
       ("Gire a frequência devagar. No ponto", "em que o defeito salta ao ouvido, pare."),
       ("Leve o ganho a −3 ou −6 dB e abra um", "pouco o Q. Compare com o EQ desligado.")]
for k in range(3):
    p = Plot(30 + 292 * k, 50, 256, 116, -12, 12)
    o.append(f'<text class="b" x="{p.x0}" y="30" font-size="13">{tit[k]}</text>')
    o.append(f'<rect class="bx" x="{p.x0}" y="{p.y0}" width="{p.w}" height="{p.h}"/>')
    o.append(f'<line class="ln-mu" stroke-width="1" x1="{p.x0}" y1="{n(p.Y(0))}" x2="{p.x0 + p.w}" y2="{n(p.Y(0))}"/>')
    if k == 0:
        o.append(f'<path class="s-purple" d="{p.path(bell(1500, 9, 5), k=400)}"/>')
        o.append(f'<text class="t-purple b" x="{n(p.X(1500) - 12)}" y="{n(p.Y(7))}" font-size="11" text-anchor="end">+9 dB, Q 5</text>')
        o.append(f'<text class="mu" x="{p.x0 + 8}" y="{n(p.Y(0) - 6)}" font-size="10.5">0 dB</text>')
    if k == 1:
        for f in (1500, 700):
            o.append(f'<path class="ln-mu dash" stroke-width="1.2" d="{p.path(bell(f, 9, 5), k=400)}"/>')
        o.append(f'<path class="s-purple" d="{p.path(bell(300, 9, 5), k=400)}"/>')
        o.append(f'<line class="s-purple" x1="{n(p.X(1900))}" y1="{n(p.Y(-6))}" x2="{n(p.X(330))}" y2="{n(p.Y(-6))}" marker-end="url(#ar6)"/>')
        o.append(f'<text class="t-purple b" x="{n(p.X(300))}" y="{n(p.Y(-9.6))}" font-size="11" text-anchor="middle">300 Hz: o defeito salta</text>')
    if k == 2:
        hc = bell(300, -5, 2)
        o.append(f'<path class="lv-purple" d="{p.area(hc)}"/>')
        o.append(f'<path class="s-purple" d="{p.path(hc, k=400)}"/>')
        o.append(f'<text class="t-purple b" x="{n(p.X(300))}" y="{n(p.Y(-8.6))}" font-size="11" text-anchor="middle">−5 dB em 300 Hz, Q 2</text>')
    o.append(f'<text class="mu" x="{p.x0}" y="{p.y0 + p.h + 14}" font-size="10.5">grave</text>')
    o.append(f'<text class="mu" x="{p.x0 + p.w}" y="{p.y0 + p.h + 14}" font-size="10.5" text-anchor="end">agudo</text>')
    o.append(f'<text x="{p.x0}" y="{p.y0 + p.h + 38}" font-size="11">{cap[k][0]}</text>')
    o.append(f'<text x="{p.x0}" y="{p.y0 + p.h + 53}" font-size="11">{cap[k][1]}</text>')
FIG6 = ('<svg viewBox="0 0 900 236" role="img" aria-label="O método de varredura em três quadros. No primeiro, uma banda bell recebe um realce estreito, de 6 a 9 dB, com Q alto, de 4 ou mais. No segundo, o operador gira a frequência devagar, e o pico percorre a faixa; no ponto em que o defeito salta ao ouvido, 300 Hz no exemplo, ele para. No terceiro, o ganho da mesma banda vira corte, de 3 a 6 dB, com o Q um pouco mais aberto; o operador compara com o equalizador desligado.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 7
def X7(f):
    return 262 + 608 * math.log10(f / 20) / math.log10(50)

rows = [("Bumbo", "peso entre 50 e 80 Hz", 50, 30, 40, True),
        ("Baixo", "nota mais grave: 41 Hz", 41, 30, 40, True),
        ("Teclado (piano)", "nota mais grave: 27,5 Hz", 27.5, 40, 120, True),
        ("Voz masculina", "fundamental a partir de 85 Hz", 85, 80, 120, False),
        ("Voz feminina", "fundamental a partir de 165 Hz", 165, 100, 180, False),
        ("Violão", "nota mais grave: 82 Hz", 82, 80, 120, False),
        ("Caixa da bateria", "corpo a partir de 150 Hz", 150, 80, 120, False),
        ("Pratos e chimbal", "som útil acima de 300 Hz", 300, 200, 400, False)]
o = ['<text class="mono mu" x="18" y="26" font-size="10.5">ONDE COMEÇA O SOM DE CADA FONTE, E ONDE LIGAR O PASSA-ALTAS</text>']
TK = [20, 30, 40, 60, 80, 100, 150, 200, 300, 400, 600, 1000]
o.append('<path class="grid" d="' + " ".join(f"M{n(X7(f))} 40V316" for f in TK) + " " + " ".join(f"M18 {77 + 34 * k}H870" for k in range(7)) + '"/>')
o.append('<g font-size="10.5" text-anchor="middle">' + "".join(f'<text class="mu" x="{n(X7(f))}" y="332">{flab(f)}</text>' for f in TK) + '</g>')
for k, (nome, sub, f0, a, b, opc) in enumerate(rows):
    y = 46 + 34 * k
    o.append(f'<text class="b" x="18" y="{y + 11}" font-size="12">{nome}</text>')
    o.append(f'<text class="mu" x="18" y="{y + 25}" font-size="10.5">{sub}</text>')
    o.append(f'<rect class="lv-teal" x="{n(X7(f0))}" y="{y + 2}" width="{n(870 - X7(f0))}" height="10"/>')
    cls = "bx-purple" if opc else "hd-purple"
    o.append(f'<rect class="{cls}" x="{n(X7(a))}" y="{y + 16}" width="{n(X7(b) - X7(a))}" height="10"/>')
o.append('<text class="mu" x="566" y="350" font-size="11" text-anchor="middle">frequência, em Hz</text>')
o.append('<rect class="lv-teal" x="18" y="364" width="26" height="10"/><text x="52" y="374" font-size="11">som da fonte</text>')
o.append('<rect class="hd-purple" x="160" y="364" width="26" height="10"/><text x="194" y="374" font-size="11">passa-altas: faixa de partida</text>')
o.append('<rect class="bx-purple" x="400" y="364" width="26" height="10"/><text x="434" y="374" font-size="11">passa-altas opcional: desligado, ou nesta faixa</text>')
FIG7 = ('<svg viewBox="0 0 900 390" role="img" aria-label="Gráfico de barras sobre um eixo de frequência de 20 Hz a 1 kHz. Para cada fonte, uma barra mostra onde o som dela começa, e outra mostra a faixa de partida do passa-altas. Bumbo: peso entre 50 e 80 Hz; passa-altas opcional, desligado ou entre 30 e 40 Hz. Baixo: nota mais grave em 41 Hz; passa-altas opcional, desligado ou entre 30 e 40 Hz. Teclado com som de piano: nota mais grave em 27,5 Hz; passa-altas opcional, desligado ou entre 40 e 120 Hz. Voz masculina: fundamental a partir de 85 Hz; passa-altas entre 80 e 120 Hz. Voz feminina: fundamental a partir de 165 Hz; passa-altas entre 100 e 180 Hz. Violão: nota mais grave em 82 Hz; passa-altas entre 80 e 120 Hz. Caixa da bateria: corpo a partir de 150 Hz; passa-altas entre 80 e 120 Hz. Pratos e chimbal: som útil acima de 300 Hz; passa-altas entre 200 e 400 Hz.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 8
p = Plot(64, 40, 520, 200, -12, 9)
hv = chain(hpf(100), bell(250, -3, 1.4), bell(3000, 2.5, 1.0))
ha = chain(hpf(120), bell(3000, -3, 1.0))
o = [f'<rect class="bx" x="{p.x0}" y="{p.y0}" width="{p.w}" height="{p.h}"/>',
     f'<rect class="lv-slate" style="fill-opacity:.22" x="{p.x0}" y="{p.y0}" width="{n(p.X(100) - p.x0)}" height="{p.h}"/>',
     f'<rect class="lv-teal" style="fill-opacity:.2" x="{n(p.X(2000))}" y="{p.y0}" width="{n(p.X(4000) - p.X(2000))}" height="{p.h}"/>',
     p.vgrid(F10), p.hgrid([6, 3, 0, -3, -6, -9]),
     f'<path class="s-amber" d="{p.path(ha)}"/>', f'<path class="s-teal" stroke-width="2.5" d="{p.path(hv)}"/>',
     f'<text class="mu" x="{n((p.x0 + p.X(100)) / 2)}" y="{p.y0 + 18}" font-size="10.5" text-anchor="middle">faixa do bumbo</text>',
     f'<text class="mu" x="{n((p.x0 + p.X(100)) / 2)}" y="{p.y0 + 32}" font-size="10.5" text-anchor="middle">e do baixo</text>',
     f'<text class="t-teal b" x="{n(p.X(2830))}" y="{p.y0 + 190}" font-size="10.5" text-anchor="middle">clareza da voz</text>']
leg = [("s-teal", "Voz", ["Passa-altas em 100 Hz, corte leve em", "250 Hz, realce de 2,5 dB em 3 kHz."]),
       ("s-amber", "Violão e teclado", ["Passa-altas em 120 Hz e corte largo de", "3 dB em 3 kHz: cedem a faixa da voz."])]
for k, (cls, t, ls) in enumerate(leg):
    y = 66 + 74 * k
    o.append(f'<line class="{cls}" x1="612" y1="{y - 4}" x2="642" y2="{y - 4}"/>')
    o.append(f'<text class="b" x="650" y="{y}" font-size="12">{t}</text>')
    for j, l in enumerate(ls):
        o.append(f'<text class="mu" x="612" y="{y + 18 + 15 * j}" font-size="11">{l}</text>')
o.append('<text x="612" y="216" font-size="11">A diferença entre as curvas em 3 kHz</text>')
o.append('<text x="612" y="231" font-size="11">passa de 5 dB, sem realce forte na voz.</text>')
o.append(f'<text class="mu" x="{p.x0 + p.w / 2}" y="{p.y0 + p.h + 36}" font-size="11" text-anchor="middle">frequência, em Hz</text>')
FIG8 = ('<svg viewBox="0 0 900 288" role="img" aria-label="Gráfico com duas curvas de equalização que se completam. A curva da voz tem passa-altas em 100 Hz, um corte leve em 250 Hz e um realce de 2,5 dB em 3 kHz. A curva do violão e do teclado tem passa-altas em 120 Hz e um corte largo de 3 dB em 3 kHz. Duas regiões estão marcadas: abaixo de 100 Hz, a faixa do bumbo e do baixo, de onde as outras fontes saem pelo passa-altas; e entre 2 e 4 kHz, a faixa da clareza da voz, que o violão e o teclado cedem. A diferença entre as duas curvas em 3 kHz passa de 5 dB, sem realce forte na voz.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 9
o = ['<defs><marker id="ar9" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto"><path class="ah" d="M0 0 L10 5 L0 10 z"/></marker></defs>']
def box(x, y, cls, t1, t2, w=120, on=False):
    c1, c2 = ("on b", "on") if on else ("b", "mu")
    return (f'<rect class="{cls}" x="{x}" y="{y}" width="{w}" height="52"/>'
            f'<text class="{c1}" x="{x + w / 2}" y="{y + 23}" font-size="11.5" text-anchor="middle">{t1}</text>'
            f'<text class="{c2}" x="{x + w / 2}" y="{y + 39}" font-size="10.5" text-anchor="middle">{t2}</text>')
xs = [18, 166, 314, 462, 610, 758]
top = [("bx", "Fonte", "e microfone", False), ("hd-purple", "EQ de canal", "corrige a fonte", True), ("bx", "Fader", "do canal", False),
       ("bx", "Mistura", "soma dos canais", False), ("hd-blue", "EQ de sistema", "corrige caixas e sala", True), ("bx", "Caixas", "da frente (PA)", False)]
o.append('<text class="mono mu" x="18" y="26" font-size="10.5">SOM DA FRENTE</text>')
for x, (cls, t1, t2, on) in zip(xs, top):
    o.append(box(x, 40, cls, t1, t2, on=on))
for x in xs[:-1]:
    o.append(f'<line class="ln" x1="{x + 122}" y1="66" x2="{x + 144}" y2="66" marker-end="url(#ar9)"/>')
o.append('<text class="mono mu" x="18" y="150" font-size="10.5">RETORNO</text>')
o.append('<circle class="f-ink" cx="300" cy="66" r="3.5"/>')
o.append('<path class="ln" d="M300 66 V182 H310" marker-end="url(#ar9)"/>')
low = [(314, "bx", "Auxiliar", "mistura do retorno", False), (462, "hd-teal", "EQ do retorno", "gráfico: tira o que apita", True), (610, "bx", "Caixa", "de retorno", False)]
for x, cls, t1, t2, on in low:
    o.append(box(x, 156, cls, t1, t2, w=120 if x != 462 else 120, on=on))
o.append('<line class="ln" x1="436" y1="182" x2="458" y2="182" marker-end="url(#ar9)"/>')
o.append('<line class="ln" x1="584" y1="182" x2="606" y2="182" marker-end="url(#ar9)"/>')
o.append('<rect class="hd-purple" x="18" y="236" width="14" height="14"/><text x="40" y="248" font-size="11.5">um por canal: muda só aquela fonte</text>')
o.append('<rect class="hd-blue" x="300" y="236" width="14" height="14"/><text x="322" y="248" font-size="11.5">um por saída: muda tudo o que vai às caixas</text>')
o.append('<rect class="hd-teal" x="18" y="260" width="14" height="14"/><text x="40" y="272" font-size="11.5">um por mistura de retorno: muda tudo o que vai àquela caixa de retorno</text>')
FIG9 = ('<svg viewBox="0 0 900 290" role="img" aria-label="Diagrama do caminho do sinal com os três lugares em que há equalizador. Na linha de cima, o som da frente: a fonte e o microfone entram no EQ de canal, que corrige a fonte; depois vêm o fader do canal, a mistura, que soma os canais, o EQ de sistema, que corrige as caixas e a sala, e as caixas da frente. Na linha de baixo, o retorno: um ramo sai do canal, depois do EQ de canal e antes do fader, e vai para o auxiliar, que é a mistura do retorno; depois vêm o EQ do retorno, um equalizador gráfico que tira as frequências que apitam, e a caixa de retorno. O EQ de canal existe um por canal e muda só aquela fonte. O EQ de sistema existe um por saída e muda tudo o que vai às caixas. O EQ do retorno existe um por mistura de retorno e muda tudo o que vai àquela caixa de retorno.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

# ------------------------------------------------------------------ Figura 10
ISO = [20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150,
       4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000]
CUT = {250: -3, 1000: -6, 4000: -4}
x0, step, yz, k = 58, 26.6, 122, 5
hg = chain(*[bell(f, g, 4.3) for f, g in CUT.items()])
o = ['<text class="mono mu" x="18" y="26" font-size="10.5">EQ GRÁFICO DE 31 BANDAS NA SAÍDA DE UM RETORNO</text>',
     f'<rect class="bx" x="40" y="46" width="834" height="152"/>',
     f'<line class="ln-mu" stroke-width="1" x1="40" y1="{yz}" x2="874" y2="{yz}"/>']
for d, lab in ((12, "+12"), (0, "0 dB"), (-12, "−12")):
    o.append(f'<text class="{"b" if d == 0 else "mu"}" x="34" y="{yz - d * k + 4}" font-size="10.5" text-anchor="end">{lab}</text>')
trilhos = []
pts = []
for i in range(0, 301):
    f = 20 * 1000 ** (i / 300)
    pts.append((x0 + step * 30 * i / 300, yz - k * db(hg, f)))
o.append('<path class="s-teal" d="M' + " L".join(f"{n(x)} {n(y)}" for x, y in rdp(pts)) + '"/>')
for i, f in enumerate(ISO):
    x = x0 + step * i
    g = CUT.get(f, 0)
    trilhos.append(f"M{n(x)} {yz - 60}V{yz + 60}")
    o.append(f'<rect class="{"hd-teal" if g else "f-ink"}" x="{n(x - 7)}" y="{n(yz - k * g - 4)}" width="14" height="8"/>')
    if i % 3 == 2:
        o.append(f'<text class="{"b" if g else "mu"}" x="{n(x)}" y="214" font-size="10.5" text-anchor="middle">{flab(f)}</text>')
    if g:
        o.append(f'<text class="t-teal b" x="{n(x)}" y="{yz - 12}" font-size="10.5" text-anchor="middle">{str(g).replace("-", "−")} dB</text>')
o.insert(3, f'<path class="grid" stroke-width="3" d="{" ".join(trilhos)}"/>')
o.append('<text class="mu" x="874" y="232" font-size="11" text-anchor="end">frequência central de cada banda, em Hz (uma banda a cada terço de oitava)</text>')
o.append('<rect class="hd-teal" x="18" y="246" width="14" height="8"/><text x="40" y="254" font-size="11.5">banda cortada no ring-out</text>')
o.append('<rect class="f-ink" x="230" y="246" width="14" height="8"/><text x="252" y="254" font-size="11.5">banda em 0 dB</text>')
o.append('<line class="s-teal" x1="380" y1="250" x2="410" y2="250"/><text x="418" y="254" font-size="11.5">curva que resulta: cada corte também alcança as bandas vizinhas</text>')
FIG10 = ('<svg viewBox="0 0 900 270" role="img" aria-label="Painel de um equalizador gráfico de 31 bandas, uma a cada terço de oitava, de 20 Hz a 20 kHz, ligado na saída de um retorno. Cada banda tem um controle deslizante, com escala de mais 12 a menos 12 dB. Vinte e oito bandas estão em 0 dB. Três foram cortadas no ring-out: 250 Hz, com menos 3 dB; 1 kHz, com menos 6 dB; e 4 kHz, com menos 4 dB. Uma curva mostra o resultado: cada corte forma um vale que também alcança as bandas vizinhas.">\n        '
        + "\n        ".join(o) + "\n      </svg>")

tpl = open(os.path.join(AQUI, "eq_body.tpl.html"), encoding="utf-8").read()
for i, fig in enumerate([FIG1, FIG2, FIG3, FIG4, FIG5, FIG6, FIG7, FIG8, FIG9, FIG10], 1):
    tok = "{{FIG%d}}" % i
    assert tpl.count(tok) == 1, tok
    tpl = tpl.replace(tok, fig)
assert "{{" not in tpl
open(DEST, "w", encoding="utf-8").write(tpl)
print("ok", DEST, len(tpl.encode("utf-8")))
