# -*- coding: utf-8 -*-
"""Gera as figuras SVG do estudo quadro e monta _geradores/quadro_body.html a partir de body_tpl.html."""
import math
import os
import re

AQUI = os.path.dirname(os.path.abspath(__file__))
DEST = "/home/josemauro/projects/josemauro/estudos_tecnico_audio_video/_geradores/quadro_body.html"


def n(x):
    s = "%.1f" % x if abs(x) < 40 else "%.0f" % x
    return s.rstrip("0").rstrip(".") if "." in s else s


def pts(*p):
    return " ".join(f"{n(x)},{n(y)}" for x, y in p)


def T(x, y, s, cls="", size=11.5, anchor=None):
    a = f' text-anchor="{anchor}"' if anchor else ""
    c = f' class="{cls}"' if cls else ""
    return f'<text{c} x="{n(x)}" y="{n(y)}" font-size="{n(size)}"{a}>{s}</text>'


def marker(i, cls="ah", size=7):
    return (f'<defs><marker id="{i}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="{size}" markerHeight="{size}" '
            f'orient="auto"><path class="{cls}" d="M0 0 L10 5 L0 10 z"/></marker></defs>')


def person(cx, top, h, cls="f-slate", op=.6):
    """Pessoa em pé, de frente, em silhueta."""
    r = h / 15
    ys, sw, yh, hw, ft = top + .165 * h, .125 * h, top + .5 * h, .09 * h, top + h
    hand = top + .56 * h
    o = [f'<circle cx="{n(cx)}" cy="{n(top + r)}" r="{n(r)}"/>',
         f'<rect x="{n(cx - .03 * h)}" y="{n(top + 2 * r - .012 * h)}" width="{n(.06 * h)}" height="{n(.05 * h)}"/>',
         f'<path d="M{n(cx - sw)} {n(ys + .03 * h)} Q{n(cx - sw)} {n(ys)} {n(cx - sw + .03 * h)} {n(ys)} H{n(cx + sw - .03 * h)} '
         f'Q{n(cx + sw)} {n(ys)} {n(cx + sw)} {n(ys + .03 * h)} L{n(cx + hw)} {n(yh)} H{n(cx - hw)} z"/>']
    for s in (-1, 1):
        o.append(f'<polygon points="{pts((cx + s * sw, ys + .02 * h), (cx + s * (sw - .045 * h), ys + .02 * h), (cx + s * (sw - .02 * h), hand), (cx + s * (sw + .03 * h), hand))}"/>')
        o.append(f'<polygon points="{pts((cx + s * hw, yh - .01 * h), (cx, yh - .01 * h), (cx + s * .01 * h, ft), (cx + s * .07 * h, ft))}"/>')
    return f'<g class="{cls}" style="opacity:{op}">' + "".join(o) + "</g>"


def bust(cx, cy, r, bottom, cls="f-slate", op=.6, eyes=True):
    """Busto de frente: cabeça com centro em (cx, cy) e ombros até bottom."""
    o = [f'<g class="{cls}" style="opacity:{op}"><circle cx="{n(cx)}" cy="{n(cy)}" r="{n(r)}"/>',
         f'<rect x="{n(cx - .42 * r)}" y="{n(cy + .8 * r)}" width="{n(.84 * r)}" height="{n(.8 * r)}"/>',
         f'<path d="M{n(cx - 2.7 * r)} {n(bottom)} V{n(cy + 2.3 * r)} Q{n(cx - 2.7 * r)} {n(cy + 1.45 * r)} {n(cx - 1.5 * r)} {n(cy + 1.4 * r)} '
         f'H{n(cx + 1.5 * r)} Q{n(cx + 2.7 * r)} {n(cy + 1.45 * r)} {n(cx + 2.7 * r)} {n(cy + 2.3 * r)} V{n(bottom)} z"/></g>']
    if eyes:
        for s in (-1, 1):
            o.append(f'<circle class="f-surface" cx="{n(cx + s * .38 * r)}" cy="{n(cy - .05 * r)}" r="{n(max(1.6, .11 * r))}"/>')
    return "".join(o)


def profile(cx, cy, r, bottom, d=1, cls="f-slate", op=.6, letter=None):
    """Busto de três quartos, olhando para a direita (d=1) ou para a esquerda (d=-1)."""
    o = [f'<g class="{cls}" style="opacity:{op}"><circle cx="{n(cx)}" cy="{n(cy)}" r="{n(r)}"/>',
         f'<polygon points="{pts((cx + d * .9 * r, cy - .2 * r), (cx + d * 1.32 * r, cy + .18 * r), (cx + d * .88 * r, cy + .34 * r))}"/>',
         f'<rect x="{n(cx - .5 * r)}" y="{n(cy + .8 * r)}" width="{n(.8 * r)}" height="{n(.8 * r)}"/>',
         f'<path d="M{n(cx - 2.2 * r)} {n(bottom)} V{n(cy + 2.3 * r)} Q{n(cx - 2.2 * r)} {n(cy + 1.45 * r)} {n(cx - 1.2 * r)} {n(cy + 1.4 * r)} '
         f'H{n(cx + .9 * r)} Q{n(cx + 1.9 * r)} {n(cy + 1.45 * r)} {n(cx + 1.9 * r)} {n(cy + 2.3 * r)} V{n(bottom)} z"/></g>',
         f'<circle class="f-surface" cx="{n(cx + d * .5 * r)}" cy="{n(cy - .12 * r)}" r="{n(max(1.6, .11 * r))}"/>']
    if letter:
        o.append(T(cx - .15 * r, cy + 3.1 * r, letter, "on b", 12, "middle"))
    return "".join(o)


def cam_top(cx, cy, ang, cls, label=None, dashed=False):
    """Câmera vista de cima; ang em graus, 0 aponta para cima, positivo no sentido horário."""
    st = ' stroke-dasharray="4 3"' if dashed else ""
    o = (f'<g transform="rotate({n(ang)} {n(cx)} {n(cy)})"{st}><rect class="{cls}" x="{n(cx - 13)}" y="{n(cy - 9)}" width="26" height="20" rx="2"/>'
         f'<rect class="{cls}" x="{n(cx - 6)}" y="{n(cy - 18)}" width="12" height="9"/></g>')
    if label:
        o += T(cx, cy + 5, label, "on b", 11, "middle")
    return o


def cam_side(cx, cy, ang, cls="bx-ink"):
    """Câmera vista de lado, lente para a direita; ang positivo inclina para baixo."""
    return (f'<g transform="rotate({n(ang)} {n(cx)} {n(cy)})"><rect class="{cls}" x="{n(cx - 24)}" y="{n(cy - 14)}" width="44" height="28" rx="2"/>'
            f'<rect class="{cls}" x="{n(cx + 20)}" y="{n(cy - 9)}" width="14" height="18"/>'
            f'<rect class="{cls}" x="{n(cx - 30)}" y="{n(cy - 10)}" width="6" height="12"/></g>')


def scene(i):
    """Cena do palco em 520 x 292: telões, palco, pastor ao centro, banda e plateia."""
    ve = ' vector-effect="non-scaling-stroke"'
    o = [f'<g id="{i}">', '<rect class="qz-sunk" x="0" y="0" width="520" height="292"/>',
         f'<rect class="qz-bx" x="40" y="34" width="110" height="62"{ve}/>', f'<rect class="qz-bx" x="370" y="34" width="110" height="62"{ve}/>',
         '<rect class="qz-lv" x="24" y="193" width="472" height="40"/>',
         person(86, 97, 98, "qz-sl").replace("<g ", f'<g id="{i}m" ', 1),
         f'<use href="#{i}m" x="54"/><use href="#{i}m" x="312"/><use href="#{i}m" x="369"/>',
         '<rect class="qz-sl" x="368" y="150" width="62" height="8"/>',
         person(260, 85, 110, "qz-sl"),
         f'<rect class="qz-bx" x="278" y="150" width="26" height="45"{ve}/>',
         '<g class="qz-sl" style="opacity:.45">']
    o += [f'<circle cx="{x}" cy="254" r="10"/>' for x in range(40, 500, 44)]
    o += [f'<circle cx="{x}" cy="284" r="12"/>' for x in range(18, 520, 48)]
    o.append("</g></g>")
    return "".join(o)


def crop(sid, clip, x, y, w, h, rx, ry, rw):
    """Mostra a região (rx, ry, rw, rw*h/w) da cena dentro do retângulo (x, y, w, h)."""
    s = w / rw
    return (f'<clipPath id="{clip}"><rect x="{n(x)}" y="{n(y)}" width="{n(w)}" height="{n(h)}"/></clipPath>'
            f'<g clip-path="url(#{clip})"><use href="#{sid}" transform="translate({n(x - rx * s)} {n(y - ry * s)}) scale({s:.3f})"/></g>')


# ------------------------------------------------------------------ Figura 1
def fig1():
    o = ['<svg viewBox="0 0 900 344" role="img" aria-label="Uma cena de culto e três quadros tirados dela. À esquerda, a cena: palco com dois telões, o pastor ao centro junto do púlpito, músicos dos lados e a plateia de costas. Três retângulos marcam os quadros: PG, plano geral, a cena toda; PM, plano médio, o pastor da cintura para cima; PP, primeiro plano, o pastor do peito para cima. À direita, cada quadro como a câmera o entrega: o geral mostra onde estamos e quem está presente; o médio, quem fala, com os gestos; o primeiro plano, o rosto e o que a pessoa sente.">',
         "<defs><style>.qz-sunk{fill:var(--sunk)}.qz-bx{fill:var(--surface);stroke:var(--rule);stroke-width:1.5}"
         ".qz-lv{fill:var(--slate);fill-opacity:.36}.qz-sl{fill:var(--slate)}</style>" + scene("cena1") + '<clipPath id="cl1"><rect x="0" y="0" width="520" height="292"/></clipPath></defs>',
         T(18, 22, "A MESMA CENA, TRÊS QUADROS", "mono mu", 10.5),
         '<g transform="translate(18 34)"><use href="#cena1" clip-path="url(#cl1)"/>',
         '<rect class="s-blue" x="1.5" y="1.5" width="517" height="289" stroke-width="3"/>',
         '<rect class="hd-blue" x="0" y="0" width="36" height="20"/>', T(18, 14.5, "PG", "on b", 12, "middle"),
         '<rect class="s-amber" x="206.7" y="75" width="106.6" height="60"/>',
         '<line class="s-amber" x1="198" y1="105" x2="206.7" y2="105"/><rect class="hd-amber" x="164" y="96" width="34" height="18"/>',
         T(181, 109.5, "PM", "on b", 11.5, "middle"),
         '<rect class="s-teal" x="224.5" y="78" width="71" height="40"/>',
         '<line class="s-teal" x1="295.5" y1="99" x2="322" y2="99"/><rect class="hd-teal" x="322" y="90" width="34" height="18"/>',
         T(339, 103.5, "PP", "on b", 11.5, "middle"), "</g>"]
    rows = [("blue", "PG", 0, 0, 520, "Plano geral", ["Onde estamos e", "quem está presente."]),
            ("amber", "PM", 206.7, 75, 106.6, "Plano médio", ["Quem fala, com os", "gestos das mãos."]),
            ("teal", "PP", 224.5, 78, 71, "Primeiro plano", ["O rosto: o que a", "pessoa sente."])]
    for k, (c, sig, rx, ry, rw, tit, ls) in enumerate(rows):
        y = 34 + k * 104
        o.append(crop("cena1", f"cl1{k}", 566, y, 150, 84, rx, ry, rw))
        o.append(f'<rect class="s-{c}" x="566" y="{y}" width="150" height="84" stroke-width="2.5"/>')
        o.append(f'<rect class="hd-{c}" x="566" y="{y}" width="32" height="18"/>' + T(582, y + 13, sig, "on b", 11, "middle"))
        o.append(T(730, y + 24, tit, "b", 12.5))
        for j, l in enumerate(ls):
            o.append(T(730, y + 43 + 15 * j, l, "mu", 11))
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figura 2
def fig2():
    cx, top, h = 190, 44, 340
    o = ['<svg viewBox="0 0 900 432" role="img" aria-label="Uma pessoa em pé, de frente, com linhas onde cada plano corta o corpo, contando de acima da cabeça. Close: até o alto dos ombros. Primeiro plano, PP: até o peito. Plano médio, PM: até a cintura. Plano americano, PA: até o meio da coxa. Plano inteiro: o corpo todo, com folga. Um quadrado na mão marca o detalhe: as mãos ou um objeto. À direita, o plano geral, PG, com o ambiente inteiro e as pessoas pequenas, e o plano conjunto, PC, com um grupo de corpo inteiro.">',
         person(cx, top, h)]
    cuts = [(.19, "Close", "o rosto: até o alto dos ombros", "teal"),
            (.29, "PP · Primeiro plano", "até o peito", "teal"),
            (.45, "PM · Plano médio", "até a cintura", "amber"),
            (.65, "PA · Plano americano", "até o meio da coxa, acima do joelho", "amber"),
            (1.03, "Plano inteiro", "o corpo todo, com folga em cima e embaixo", "blue")]
    o.append(f'<line class="ref" x1="100" y1="{top - 10}" x2="372" y2="{top - 10}"/>')
    o.append(T(386, top - 6, "todos os planos começam acima da cabeça", "mu", 11))
    for fr, tit, desc, c in cuts:
        y = top + fr * h
        o.append(f'<line class="s-{c}" x1="100" y1="{n(y)}" x2="372" y2="{n(y)}"/>')
        o.append(T(386, y + 4, tit, "b", 12))
        o.append(T(386, y + 19, desc, "mu", 11))
    hx = cx - .125 * h - .03 * h
    o.append(f'<rect class="s-purple" x="{n(hx - 8)}" y="{n(top + .5 * h - 6)}" width="34" height="34"/>')
    o.append(T(hx - 16, top + .5 * h + 6, "Detalhe", "b", 12, "end"))
    o.append(T(hx - 16, top + .5 * h + 21, "as mãos,", "mu", 11, "end"))
    o.append(T(hx - 16, top + .5 * h + 35, "um objeto", "mu", 11, "end"))
    # quadros à direita
    o.append(crop("cena1", "cl2a", 660, 44, 222, 125, 0, 0, 520))
    o.append('<rect class="s-blue" x="660" y="44" width="222" height="125"/>')
    o.append(T(660, 190, "PG · Plano geral", "b", 12))
    o.append(T(660, 206, "o ambiente inteiro;", "mu", 11))
    o.append(T(660, 221, "as pessoas ficam pequenas", "mu", 11))
    o.append(crop("cena1", "cl2b", 660, 244, 222, 125, 52, 76, 236))
    o.append('<rect class="s-blue" x="660" y="244" width="222" height="125"/>')
    o.append(T(660, 390, "PC · Plano conjunto", "b", 12))
    o.append(T(660, 406, "um grupo de corpo inteiro,", "mu", 11))
    o.append(T(660, 421, "com um pouco do ambiente", "mu", 11))
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figuras 3 e 4
FX, FY, FW, FH = (18, 315, 612), 44, 270, 152


def frame(k, fid, inner, grid=True):
    x = FX[k]
    o = [f'<clipPath id="{fid}"><rect x="{x}" y="{FY}" width="{FW}" height="{FH}"/></clipPath>',
         f'<rect class="f-sunk" x="{x}" y="{FY}" width="{FW}" height="{FH}"/>',
         f'<g clip-path="url(#{fid})">{inner}</g>']
    if grid:
        for j in (1, 2):
            o.append(f'<line class="ref" x1="{x}" y1="{n(FY + FH * j / 3)}" x2="{x + FW}" y2="{n(FY + FH * j / 3)}"/>')
            o.append(f'<line class="ref" x1="{n(x + FW * j / 3)}" y1="{FY}" x2="{n(x + FW * j / 3)}" y2="{FY + FH}"/>')
    o.append(f'<rect class="ln" x="{x}" y="{FY}" width="{FW}" height="{FH}"/>')
    return "".join(o)


def caption(k, tit, cls, lines):
    x = FX[k]
    o = [T(x, FY + FH + 24, tit, "b " + cls, 13)]
    for j, l in enumerate(lines):
        o.append(T(x, FY + FH + 43 + 16 * j, l, "", 11.5))
    return "".join(o)


def brace(x, y1, y2, label):
    return (f'<path class="ln" stroke-width="1" d="M{x - 4} {n(y1)} H{x + 4} M{x} {n(y1)} V{n(y2)} M{x - 4} {n(y2)} H{x + 4}"/>'
            + T(x + 9, (y1 + y2) / 2 + 4, label, "b", 11))


def fig3():
    r = 26
    o = ['<svg viewBox="0 0 900 262" role="img" aria-label="Três quadros de plano médio da mesma pessoa, cada um com a grade de terços: duas linhas horizontais e duas verticais. No primeiro, teto demais: os olhos ficaram no meio do quadro, sobra muito espaço acima da cabeça e a pessoa parece afundar. No segundo, na medida: os olhos estão na linha do terço superior, e o espaço acima da cabeça fica pequeno e natural. No terceiro, sem teto: a cabeça encosta na borda de cima e é cortada, e o quadro fica apertado.">',
         T(18, 26, "A GRADE DE TERÇOS E O ESPAÇO ACIMA DA CABEÇA", "mono mu", 10.5)]
    cys = [FY + FH / 2 + 2, FY + FH / 3 + 2, FY + 18]
    for k in range(3):
        cx = FX[k] + FW / 2
        o.append(frame(k, f"cl3{k}", bust(cx, cys[k], r, FY + FH + 4)))
    o.append(brace(FX[0] + 70, FY + 2, cys[0] - r - 1, "teto"))
    o.append(brace(FX[1] + 70, FY + 2, cys[1] - r - 1, "teto"))
    o.append(T(FX[1] + FW - 6, FY + FH / 3 - 5, "terço superior", "mu", 10.5, "end"))
    o.append(caption(0, "Teto demais", "t-bad", ["Os olhos ficaram no meio do quadro:", "a pessoa parece afundar."]))
    o.append(caption(1, "Na medida", "t-good", ["Os olhos ficam na linha do terço", "superior, e o teto se ajusta sozinho."]))
    o.append(caption(2, "Sem teto", "t-bad", ["A cabeça encosta na borda ou é", "cortada, e o quadro fica apertado."]))
    o.append("</svg>")
    return "\n".join(o)


def fig4():
    r = 24
    cy = FY + FH / 3 + 4
    o = ['<svg viewBox="0 0 900 262" role="img" aria-label="Três quadros com a grade de terços. No primeiro, nariz na borda: a pessoa olha para a direita, mas está encostada no lado direito do quadro; sobra espaço nas costas, e o olhar bate na borda. No segundo, espaço do olhar: a mesma pessoa está no terço esquerdo, e o espaço maior fica à frente do rosto, na direção para onde ela olha. No terceiro, espaço do movimento: uma pessoa de corpo inteiro anda para a direita e está no terço esquerdo, com o quadro livre à frente, para onde ela vai.">',
         marker("ar4", "ah", 6), T(18, 26, "O ESPAÇO VAI PARA O LADO DO OLHAR E DO MOVIMENTO", "mono mu", 10.5)]
    xa = FX[0] + FW * .8
    o.append(frame(0, "cl40", profile(xa, cy, r, FY + FH + 4)))
    o.append(f'<line class="s-bad dash" x1="{n(xa + 1.5 * r)}" y1="{n(cy)}" x2="{FX[0] + FW - 3}" y2="{n(cy)}"/>')
    xb = FX[1] + FW * .32
    o.append(frame(1, "cl41", profile(xb, cy, r, FY + FH + 4)))
    o.append(f'<line class="s-good dash" x1="{n(xb + 1.5 * r)}" y1="{n(cy)}" x2="{FX[1] + FW - 16}" y2="{n(cy)}" marker-end="url(#ar4)"/>')
    o.append(T(FX[1] + FW * .68, cy - 9, "espaço do olhar", "b", 11, "middle"))
    xc = FX[2] + FW * .28
    o.append(frame(2, "cl42", person(xc, FY + 20, 118)))
    o.append(f'<line class="s-good dash" x1="{n(xc + 26)}" y1="{FY + 96}" x2="{FX[2] + FW - 16}" y2="{FY + 96}" marker-end="url(#ar4)"/>')
    o.append(T(FX[2] + FW * .66, FY + 87, "espaço do movimento", "b", 11, "middle"))
    o.append(caption(0, "Nariz na borda", "t-bad", ["Sobra espaço nas costas, e o olhar", "bate na borda do quadro."]))
    o.append(caption(1, "Espaço do olhar", "t-good", ["A pessoa fica em um terço, e o espaço", "maior fica à frente do rosto."]))
    o.append(caption(2, "Espaço do movimento", "t-good", ["Quem anda precisa de quadro livre", "à frente, para onde vai."]))
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figura 5
def fig5():
    px, top, h = 770, 124, 190
    tx, ty = px - 20, top + .07 * h
    floor = top + h
    o = ['<svg viewBox="0 0 900 340" role="img" aria-label="Vista de lado de uma pessoa em pé e três câmeras apontadas para o rosto dela, em alturas diferentes. A câmera alta, acima da cabeça, olha para baixo: é o plongée, e a pessoa parece menor. A câmera na altura dos olhos de quem aparece dá o ângulo normal, o olhar de uma conversa. A câmera baixa, perto do chão, olha para cima: é o contra-plongée; a pessoa parece maior, e o teto aparece no quadro.">',
         T(18, 22, "A ALTURA DA CÂMERA, VISTA DE LADO", "mono mu", 10.5),
         f'<line class="ln-mu" x1="330" y1="{floor}" x2="870" y2="{floor}"/>', person(px, top, h)]
    cams = [(46, "purple", "Plongée", ["câmera alta, olhando para baixo:", "a pessoa parece menor."], 44),
            (ty, "teal", "Normal", ["na altura dos olhos de quem aparece:", "é o olhar de uma conversa."], ty - 8),
            (floor - 22, "amber", "Contra-plongée", ["câmera baixa, olhando para cima:", "a pessoa parece maior, e aparece o teto."], floor - 44)]
    cx = 392
    for cy, c, tit, ls, ly in cams:
        ang = math.degrees(math.atan2(ty - cy, tx - cx))
        o.append(f'<line class="s-{c} dash" x1="{cx + 30}" y1="{n(cy + 30 * math.tan(math.radians(ang)))}" x2="{n(tx - 6)}" y2="{n(ty)}"/>')
        o.append(cam_side(cx, cy, ang))
        o.append(T(18, ly, tit, "b t-" + c, 13))
        for j, l in enumerate(ls):
            o.append(T(18, ly + 18 + 15 * j, l, "", 11.5))
    o.append(f'<circle class="f-ink" cx="{n(tx - 2)}" cy="{n(ty)}" r="3"/>')
    o.append(T(px + 34, ty + 4, "olhos", "mu", 11))
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figura 6
def fig6():
    o = ['<svg viewBox="0 0 900 286" role="img" aria-label="Quatro quadros, um para cada movimento. Pan, ou panorâmica, vista de cima: a câmera gira para os lados, sem sair do lugar. Tilt, vista de lado: a câmera, sobre o tripé, inclina para cima ou para baixo. Zoom, vista de cima: a câmera fica parada, e a lente fecha o ângulo de visão, de um cone largo para um cone estreito. Travelling, vista de cima: a câmera inteira se desloca sobre um trilho, de uma posição para outra, e o ponto de vista muda.">',
         marker("ar6", "ah", 6)]
    names = [("Pan (panorâmica)", "VISTA DE CIMA", ["A câmera gira para os lados,", "sem sair do lugar."]),
             ("Tilt", "VISTA DE LADO", ["A câmera inclina para cima", "ou para baixo."]),
             ("Zoom", "VISTA DE CIMA", ["A lente fecha ou abre o ângulo.", "A câmera não sai do lugar."]),
             ("Travelling", "VISTA DE CIMA", ["A câmera inteira se desloca,", "e o ponto de vista muda."])]
    for k, (tit, vista, ls) in enumerate(names):
        x0 = 18 + k * 220
        cx = x0 + 102
        o.append(T(x0, 26, tit, "b", 13))
        o.append(f'<rect class="bx" x="{x0}" y="40" width="204" height="170"/>')
        o.append(T(x0 + 10, 58, vista, "mono mu", 10.5))
        for j, l in enumerate(ls):
            o.append(T(x0, 232 + 16 * j, l, "", 11.5))
        if k == 0:
            cy, R = 172, 78
            for s in (-1, 1):
                ex, ey = cx + s * R * math.sin(math.radians(42)), cy - R * math.cos(math.radians(42))
                o.append(f'<line class="ref" x1="{cx}" y1="{cy - 20}" x2="{n(cx + s * (R - 8) * math.sin(math.radians(42)))}" y2="{n(cy - (R - 8) * math.cos(math.radians(42)))}"/>')
                o.append(f'<path class="s-blue" d="M{n(cx + s * 8)} {n(cy - R + .4)} A{R} {R} 0 0 {1 if s > 0 else 0} {n(ex)} {n(ey)}" marker-end="url(#ar6)"/>')
            o.append(cam_top(cx, cy, 0, "bx-ink"))
        elif k == 1:
            ax, ay, R = cx - 26, 122, 74
            o.append(f'<path class="ln" d="M{ax} {ay + 14} L{ax - 30} 200 M{ax} {ay + 14} L{ax + 30} 200 M{ax} {ay + 14} V200"/>')
            for s in (-1, 1):
                ex, ey = ax + R * math.cos(math.radians(32)), ay + s * R * math.sin(math.radians(32))
                o.append(f'<line class="ref" x1="{ax + 36}" y1="{n(ay + s * 36 * math.tan(math.radians(32)))}" x2="{n(ax + (R - 8) * math.cos(math.radians(32)))}" y2="{n(ay + s * (R - 8) * math.sin(math.radians(32)))}"/>')
                o.append(f'<path class="s-blue" d="M{n(ax + R - .4)} {n(ay + s * 8)} A{R} {R} 0 0 {1 if s > 0 else 0} {n(ex)} {n(ey)}" marker-end="url(#ar6)"/>')
            o.append(cam_side(ax, ay, 0))
        elif k == 2:
            cy = 186
            o.append(f'<polygon class="lv-blue" points="{pts((cx, cy - 18), (cx - 82, 70), (cx + 82, 70))}"/>')
            o.append(f'<polygon class="lv-amber" points="{pts((cx, cy - 18), (cx - 24, 70), (cx + 24, 70))}"/>')
            o.append(f'<line class="ln" x1="{cx + 64}" y1="84" x2="{cx + 30}" y2="84" marker-end="url(#ar6)"/>')
            o.append(f'<line class="ln" x1="{cx - 64}" y1="84" x2="{cx - 30}" y2="84" marker-end="url(#ar6)"/>')
            o.append(cam_top(cx, cy, 0, "bx-ink"))
        else:
            y = 168
            o.append(f'<path class="ln-mu" stroke-width="1" d="M{x0 + 16} {y - 4} H{x0 + 188} M{x0 + 16} {y + 6} H{x0 + 188}"/>')
            o.append(f'<circle class="f-slate" style="opacity:.6" cx="{cx}" cy="82" r="9"/>')
            for xx in (x0 + 46, x0 + 158):
                o.append(f'<line class="ref" x1="{xx}" y1="{y - 20}" x2="{n(cx + (xx - cx) * .12)}" y2="94"/>')
            a1 = math.degrees(math.atan2(cx - (x0 + 46), y - 82))
            o.append(f'<g transform="rotate({n(a1)} {x0 + 46} {y})"><rect class="bx-out" x="{x0 + 33}" y="{y - 9}" width="26" height="20" rx="2"/><rect class="bx-out" x="{x0 + 40}" y="{y - 18}" width="12" height="9"/></g>')
            o.append(cam_top(x0 + 158, y, -a1, "bx-ink"))
            o.append(f'<line class="s-blue" x1="{x0 + 70}" y1="{y + 26}" x2="{x0 + 136}" y2="{y + 26}" marker-end="url(#ar6)"/>')
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figura 7
def fig7():
    def bub(x, y, t):
        return f'<circle class="hd-blue" cx="{x}" cy="{y}" r="10"/>' + T(x, y + 4, t, "on b", 11.5, "middle")
    o = ['<svg viewBox="0 0 900 420" role="img" aria-label="Câmera sobre tripé com cabeça fluida, vista de lado, com sete partes numeradas: 1, placa deslizante, para equilibrar; 2, contrapeso, que segura a câmera em qualquer inclinação; 3, fricção de pan e de tilt, que deixa o movimento contínuo; 4, travas, que prendem a cabeça e não dão resistência; 5, nível de bolha e meia-esfera, para nivelar; 6, alavanca, conduzida com a mão leve; 7, pernas e espalhador, bem abertos em chão firme.">',
         T(18, 22, "CABEÇA FLUIDA SOBRE TRIPÉ, VISTA DE LADO", "mono mu", 10.5),
         '<line class="ln-mu" x1="60" y1="392" x2="420" y2="392"/>',
         '<path class="ln-bold" d="M232 256 L130 392 M248 256 L350 392"/><line class="ln" x1="240" y1="258" x2="240" y2="392"/>',
         '<line class="ln-mu" x1="178" y1="328" x2="302" y2="328"/>',
         '<path class="f-slate" d="M216 250 A24 24 0 0 0 264 250 z"/>',
         '<rect class="bx-slate" x="212" y="206" width="56" height="44"/>',
         '<rect class="f-slate" x="186" y="196" width="112" height="10"/>',
         '<rect class="bx" x="176" y="136" width="120" height="60" rx="3"/><rect class="bx" x="296" y="148" width="50" height="36"/>',
         '<rect class="bx" x="164" y="144" width="12" height="22"/>',
         '<path class="ln-bold" d="M216 234 L120 288"/><path class="ln-bold" stroke-width="7" d="M120 288 L96 301"/>',
         bub(318, 208, "1"), bub(284, 228, "2"), bub(240, 228, "3"), bub(196, 220, "4"), bub(240, 270, "5"), bub(146, 250, "6"), bub(240, 328, "7"),
         T(236, 171, "câmera", "mu", 11, "middle")]
    leg = [("Placa deslizante", "Avança ou recua a câmera até ela parar equilibrada."),
           ("Contrapeso (counterbalance)", "Mola que segura a câmera parada em qualquer inclinação."),
           ("Fricção (drag) de pan e de tilt", "Resistência suave que deixa o movimento contínuo."),
           ("Travas (lock) de pan e de tilt", "Prendem a cabeça. Não servem para dar resistência."),
           ("Nível de bolha e meia-esfera", "Nivelam a cabeça sem mexer no comprimento das pernas."),
           ("Alavanca (manche)", "Conduz o movimento, com a mão leve."),
           ("Pernas e espalhador", "Base bem aberta, no chão firme, com as travas apertadas.")]
    for k, (a, b) in enumerate(leg):
        y = 58 + k * 50
        o.append(bub(480, y, str(k + 1)))
        o.append(T(500, y + 1, a, "b", 12.5))
        o.append(T(500, y + 18, b, "mu", 11))
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figura 8
def fig8():
    ax = 136
    o = ['<svg viewBox="0 0 900 400" role="img" aria-label="Vista de cima de um palco com duas pessoas, A e B, uma de frente para a outra. Uma linha tracejada passa pelas duas e atravessa a sala: é o eixo. As câmeras 1, 2 e 3 estão do lado da plateia, todas do mesmo lado do eixo. A câmera 4 está atrás do palco, do outro lado do eixo. À direita, dois quadros. Nas câmeras 1, 2 e 3, A fica à esquerda e B à direita, em todos os cortes. Na câmera 4, A e B trocam de lado, e quem assiste pensa que alguém mudou de lugar.">',
         f'<rect class="lv-red" style="fill-opacity:.13" x="18" y="34" width="582" height="{ax - 34}"/>',
         f'<rect class="lv-teal" style="fill-opacity:.13" x="18" y="{ax}" width="582" height="{382 - ax}"/>',
         f'<rect class="bx-out" x="180" y="{ax - 42}" width="260" height="84"/>',
         T(190, ax - 26, "PALCO", "mono mu", 10.5),
         f'<line class="ln dash" x1="18" y1="{ax}" x2="600" y2="{ax}"/>',
         T(592, ax - 8, "eixo: a linha entre as duas pessoas", "b", 11, "end"),
         T(30, 54, "o outro lado do eixo", "b t-red", 12),
         T(30, 370, "lado da plateia: todas as câmeras ficam aqui", "b t-teal", 12)]
    mx = 310
    cams = [(310, 336, "teal", "1"), (120, 262, "teal", "2"), (500, 262, "teal", "3"), (500, 62, "red", "4")]
    for x, y, c, lab in cams:
        ang = math.degrees(math.atan2(mx - x, y - ax))
        d = math.hypot(mx - x, ax - y)
        o.append(f'<line class="ref" x1="{n(x + (mx - x) * 24 / d)}" y1="{n(y + (ax - y) * 24 / d)}" x2="{n(mx - (mx - x) * 22 / d)}" y2="{n(ax - (ax - y) * 22 / d)}"/>')
        o.append(cam_top(x, y, ang, "hd-" + c, lab))
    for x, c, lab, s in ((260, "blue", "A", 1), (360, "amber", "B", -1)):
        o.append(f'<polygon class="hd-{c}" points="{pts((x + s * 12, ax - 6), (x + s * 21, ax), (x + s * 12, ax + 6))}"/>')
        o.append(f'<circle class="hd-{c}" cx="{x}" cy="{ax}" r="14"/>' + T(x, ax + 4.5, lab, "on b", 12.5, "middle"))
    for k, (y0, tit, cls, left, right, ls) in enumerate([
            (50, "Câmeras 1, 2 e 3", "t-teal", ("blue", "A"), ("amber", "B"), ["A fica à esquerda e B, à direita,", "em todos os cortes."]),
            (238, "Câmera 4", "t-red", ("amber", "B"), ("blue", "A"), ["A e B trocam de lado: quem assiste", "pensa que alguém mudou de lugar."])]):
        x0, w, hh = 650, 200, 112
        o.append(T(x0, y0 - 9, tit, "b " + cls, 12.5))
        o.append(f'<clipPath id="cl8{k}"><rect x="{x0}" y="{y0}" width="{w}" height="{hh}"/></clipPath><rect class="f-sunk" x="{x0}" y="{y0}" width="{w}" height="{hh}"/>')
        o.append(f'<g clip-path="url(#cl8{k})">' + profile(x0 + 52, y0 + 40, 15, y0 + hh + 4, 1, "f-" + left[0], .9, left[1])
                 + profile(x0 + 152, y0 + 40, 15, y0 + hh + 4, -1, "f-" + right[0], .9, right[1]) + "</g>")
        o.append(f'<rect class="ln" x="{x0}" y="{y0}" width="{w}" height="{hh}"/>')
        for j, l in enumerate(ls):
            o.append(T(x0, y0 + hh + 19 + 15 * j, l, "", 11))
    o.append("</svg>")
    return "\n".join(o)


# ------------------------------------------------------------------ Figura 9
def fig9():
    tx, ty = 298, 76
    o = ['<svg viewBox="0 0 900 440" role="img" aria-label="Planta de uma igreja vista de cima: palco no alto, púlpito ao centro, dois blocos de bancos e corredores. Posição 1, central, ao fundo, sobre praticável: plano geral, segurança e médio de quem fala. Posição 2, lateral direita, a uns 30 a 45 graus do centro: rosto de três quartos, em plano médio e primeiro plano. Posição 3, móvel, perto do palco à esquerda: banda, detalhes e congregação, na mão ou em estabilizador. Posição R, reversa, no fundo do palco: só o público, porque para quem está no palco ela inverte os lados. Com uma câmera, posição 1; com duas, 1 e 2; com três, 1, 2 e 3.">',
         '<rect class="bx" x="18" y="20" width="560" height="402"/>',
         '<rect class="lv-slate" x="158" y="36" width="280" height="66"/>', T(428, 54, "PALCO", "mono mu", 10.5, "end"),
         f'<rect class="bx" x="286" y="{ty + 10}" width="24" height="10"/><circle class="f-slate" style="opacity:.7" cx="{tx}" cy="{ty}" r="8"/>']
    o.append('<g class="lv-slate" style="fill-opacity:.22">' + "".join(f'<rect x="66" y="{y}" width="202" height="12"/><rect x="328" y="{y}" width="202" height="12"/>' for y in range(150, 340, 26)) + "</g>")
    o.append('<rect class="bx-out" x="266" y="364" width="64" height="44"/>')
    o.append(T(338, 402, "praticável", "mu", 10.5))
    o.append('<rect class="bx-out" x="36" y="108" width="118" height="34"/>')
    cams = [(298, 386, tx, ty, "blue", "1", False), (536, 354, tx, ty, "amber", "2", False), (70, 125, 215, 78, "teal", "3", False), (180, 54, 298, 240, "slate", "R", True)]
    for x, y, ax, ay, c, lab, dashed in cams:
        ang = math.degrees(math.atan2(ax - x, y - ay))
        d = math.hypot(ax - x, ay - y)
        end = 14 if lab in "12" else d - 70
        if lab != "3":
            o.append(f'<line class="ref" x1="{n(x + (ax - x) * 24 / d)}" y1="{n(y + (ay - y) * 24 / d)}" x2="{n(ax - (ax - x) * end / d)}" y2="{n(ay - (ay - y) * end / d)}"/>')
        o.append(cam_top(x, y, ang, "hd-" + c, lab))
    a2 = math.atan2(536 - tx, 354 - ty)
    R = 62
    o.append(f'<path class="s-amber" stroke-width="1.5" d="M{tx} {ty + R} A{R} {R} 0 0 0 {n(tx + R * math.sin(a2))} {n(ty + R * math.cos(a2))}"/>')
    o.append(T(362, 141, "30° a 45°", "b t-amber", 11))
    leg = [("blue", "1", "Central, ao fundo", ["De frente, sobre praticável: plano geral,", "segurança e médio de quem fala."]),
           ("amber", "2", "Lateral, a uns 30° a 45° do centro", ["Rosto de três quartos: plano médio e", "primeiro plano, diferente da câmera 1."]),
           ("teal", "3", "Móvel, perto do palco", ["Banda, detalhes e congregação. Na mão,", "em monopé ou em estabilizador."]),
           ("slate", "R", "Reversa, no fundo do palco", ["Mostra só o público: para quem está", "no palco, ela inverte os lados."])]
    for k, (c, lab, tit, ls) in enumerate(leg):
        y = 44 + k * 68
        o.append(f'<circle class="hd-{c}" cx="612" cy="{y}" r="10"/>' + T(612, y + 4, lab, "on b", 11.5, "middle"))
        o.append(T(630, y + 4, tit, "b", 12))
        for j, l in enumerate(ls):
            o.append(T(630, y + 21 + 15 * j, l, "mu", 11))
    o.append(T(602, 330, "COBERTURA", "mono mu", 10.5))
    for j, l in enumerate(["Com 1 câmera: posição 1.", "Com 2 câmeras: posições 1 e 2.", "Com 3 câmeras: posições 1, 2 e 3."]):
        o.append(T(602, 352 + 18 * j, l, "", 11.5))
    o.append("</svg>")
    return "\n".join(o)


def compact(svg):
    """Leva o font-size mais comum dos textos para o elemento svg, que o herda."""
    from collections import Counter
    sizes = Counter(re.findall(r'<text[^>]*? font-size="([\d.]+)"', svg))
    if not sizes:
        return svg
    fs = sizes.most_common(1)[0][0]
    svg = re.sub(r'(<text[^>]*?) font-size="%s"' % re.escape(fs), r"\1", svg)
    return svg.replace("<svg ", f'<svg font-size="{fs}" ', 1)


FIGS = {"F1": fig1, "F2": fig2, "F3": fig3, "F4": fig4, "F5": fig5, "F6": fig6, "F7": fig7, "F8": fig8, "F9": fig9}

if __name__ == "__main__":
    tpl = open(os.path.join(AQUI, "body_tpl.html"), encoding="utf-8").read()
    for k, fn in FIGS.items():
        assert "{{" + k + "}}" in tpl, k
        svg = compact(fn())
        svg = "\n".join("      " + l for l in svg.split("\n"))
        tpl = tpl.replace("{{" + k + "}}", svg.strip())
    assert "{{" not in tpl
    open(DEST, "w", encoding="utf-8").write(tpl.strip() + "\n")
    words = len(re.sub(r"<[^>]+>", " ", re.sub(r"<svg.*?</svg>", " ", tpl, flags=re.S)).split())
    print("ok", DEST, len(tpl.encode("utf-8")), "bytes;", words, "palavras fora das figuras")
