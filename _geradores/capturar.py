# -*- coding: utf-8 -*-
"""Captura cada seção de um treinamento em PNG e confere os textos das figuras.

    python capturar.py <treinamento.html> <pasta de saída> [--escuro]

Para cada <section id> grava <pasta>/<id>.png (ou <id>-escuro.png). Depois lista os textos de SVG que
saem do viewBox ou que se sobrepõem a outro texto da mesma figura.
"""
import json
import os
import re
import subprocess
import sys

SRC, OUT = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
ESCURO = "--escuro" in sys.argv
CHROME = os.environ.get("CHROME", "google-chrome")
os.makedirs(OUT, exist_ok=True)
html = open(SRC, encoding="utf-8").read()
if ESCURO:
    html = html.replace('<html lang="pt-BR">', '<html lang="pt-BR" data-theme="dark">', 1)
ids = re.findall(r'<section\b[^>]*\bid="([^"]+)"', html)
base = os.path.dirname(SRC)


def chrome(*args):
    return subprocess.run([CHROME, "--headless", "--disable-gpu", "--hide-scrollbars", "--virtual-time-budget=4000", *args],
                          capture_output=True, text=True).stdout


for i in ids:
    css = (f"<style>.mast,.toc,footer,main>section:not(#{i}){{display:none!important}}"
           ".layout{display:block!important;padding-top:0!important}.page{padding-block:16px!important}</style>")
    tmp = os.path.join(base, f".captura-{i}.html")
    open(tmp, "w", encoding="utf-8").write(html.replace("</head>", css + "</head>", 1))
    alvo = os.path.join(OUT, f"{i}{'-escuro' if ESCURO else ''}.png")
    try:
        # primeiro mede a altura da seção, depois captura no tamanho certo
        js = "<script>addEventListener('load',function(){document.title='H='+Math.ceil(document.documentElement.scrollHeight)})</script>"
        open(tmp, "a", encoding="utf-8").write(js)
        dom = chrome("--dump-dom", "file://" + tmp)
        m = re.search(r"<title>H=(\d+)</title>", dom)
        h = min(int(m.group(1)) if m else 3000, 12000)
        chrome(f"--screenshot={alvo}", f"--window-size=1100,{h}", "file://" + tmp)
        print("ok ", alvo, f"(altura {h})")
    finally:
        os.remove(tmp)

if not ESCURO:
    js = """<script>addEventListener('load',function(){var out=[];
    document.querySelectorAll('figure svg').forEach(function(svg,n){
      var vb=svg.viewBox.baseVal,m=svg.getScreenCTM().inverse(),bs=[];
      svg.querySelectorAll('text').forEach(function(t){
        var r=t.getBoundingClientRect(); if(!r.width){return;}
        var p=svg.createSVGPoint(); p.x=r.left;p.y=r.top; var a=p.matrixTransform(m); p.x=r.right;p.y=r.bottom; var b=p.matrixTransform(m);
        var o={s:t.textContent.trim().slice(0,40),x1:a.x,y1:a.y,x2:b.x,y2:b.y}; bs.push(o);
        if(a.x<vb.x-1||b.x>vb.x+vb.width+1||a.y<vb.y-1||b.y>vb.y+vb.height+1){out.push('figura '+(n+1)+': fora do viewBox: '+o.s);}
      });
      for(var i=0;i<bs.length;i++){for(var j=i+1;j<bs.length;j++){var A=bs[i],B=bs[j];
        var w=Math.min(A.x2,B.x2)-Math.max(A.x1,B.x1),h=Math.min(A.y2,B.y2)-Math.max(A.y1,B.y1);
        if(w>3&&h>4){out.push('figura '+(n+1)+': sobrepostos: "'+A.s+'" e "'+B.s+'"');}}}
    });
    var pre=document.createElement('pre');pre.id='conf';pre.textContent=JSON.stringify(out);document.body.appendChild(pre);});</script>"""
    tmp = os.path.join(base, ".captura-conf.html")
    open(tmp, "w", encoding="utf-8").write(html.replace("</body>", js + "</body>", 1))
    try:
        dom = chrome("--dump-dom", "file://" + tmp)
    finally:
        os.remove(tmp)
    m = re.search(r'<pre id="conf">(.*?)</pre>', dom, re.S)
    if not m:
        print("conferência dos textos: não foi possível medir")
    else:
        import html as H
        probs = json.loads(H.unescape(m.group(1)))
        print("textos das figuras:", "nenhum problema" if not probs else "")
        for p in probs:
            print("  -", p)
