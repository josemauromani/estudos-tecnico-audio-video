# -*- coding: utf-8 -*-
"""Monta o treinamento de um estudo: cabeçalho + base.css + <nome>_body.html + base.js.

    python montar.py <nome>_body.html <saida.html>

O título da página vem do <h1> do corpo. O corpo começa em <div class="page" ...> e
termina no </footer>; o fechamento da página e o script são acrescentados aqui.
"""
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
BODY, OUT = sys.argv[1], sys.argv[2]

body = open(BODY, encoding="utf-8").read().strip()
css = open(os.path.join(AQUI, "base.css"), encoding="utf-8").read()
js = open(os.path.join(AQUI, "base.js"), encoding="utf-8").read()

m = re.search(r"<h1>(.*?)</h1>", body, re.S)
assert m, "o corpo precisa de um <h1>"
title = re.sub(r"<[^>]+>", "", m.group(1)).strip()
assert body.startswith('<div class="page"'), 'o corpo deve começar em <div class="page" ...>'
assert body.endswith("</footer>"), "o corpo deve terminar em </footer>"

# conferências simples, para o erro aparecer na geração e não no navegador
ids = re.findall(r'<section\b[^>]*\bid="([^"]+)"', body)
for href in re.findall(r'<nav class="toc".*?</nav>', body, re.S)[0].split('href="#')[1:]:
    alvo = href.split('"')[0]
    assert alvo in ids, "o sumário aponta para uma seção que não existe: #" + alvo
if 'id="quiz"' in body:
    assert 'id="quiz-data"' in body, "o exercício precisa do bloco quiz-data"

FONTS = ("https://fonts.googleapis.com/css2?family=Barlow+Semi+Condensed:wght@500;600;700"
         "&family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:ital,wght@0,400;0,500;0,600;1,400&display=swap")
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{title}</title>
<link rel="stylesheet" href="{FONTS}">
<style>
{css}</style>
</head>
<body>
{body}
</div>

<script>
{js}</script>
</body>
</html>
"""
open(OUT, "w", encoding="utf-8").write(html)
print("ok", OUT, len(html))
