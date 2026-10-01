# Guia do autor de um estudo

Cada estudo tem dois arquivos-fonte nesta pasta, e gera dois arquivos na pasta do estudo:

| Fonte | Gera |
|---|---|
| `<nome>_body.html` | `<Pasta>/treinamento-<slug>.html` (montado por `montar.py`, com `base.css` e `base.js`) |
| `build_<nome>.py` | `<Pasta>/<Pasta>-modelo.xlsx` (com `xlsx_lib.py`, recalculada no LibreOffice) |

Os nomes, as pastas e os arquivos estão em `estudos.txt`. Para gerar: `_geradores/gerar.sh <nome>`.

## Quem lê

Um voluntário ou técnico iniciante que opera som e vídeo em igreja e em eventos ao vivo pequenos e médios
(culto, palestra, formatura, apresentação musical). Ele sabe ligar os equipamentos, mas não sabe por que
as coisas funcionam. Quer entender o princípio e sair com um procedimento para usar no próximo evento.

## Tom e texto

- Português do Brasil, frases curtas, voz ativa, sem jargão sem explicação. O termo em inglês que aparece
  no equipamento (gain, fader, PFL, mute, aux, iris, shutter) é apresentado uma vez junto do termo em português.
- Explique o princípio com uma analogia concreta do dia a dia quando ela ajudar (como "ganho é a entrada da
  caixa d'água, fader é a torneira"). Uma analogia por conceito, sem exagero.
- Sem emojis. Sem marcas e modelos de equipamento como recomendação. Marcas só quando o nome é o termo
  usual da área (XLR, Speakon, HDMI, SDI, NDI, OBS).
- Exatidão técnica vem primeiro. Números são faixas típicas e ponto de partida, e o texto diz isso:
  "valores de referência; confira o manual do seu equipamento". Não invente norma, lei ou especificação.
  Na dúvida entre dois valores, dê a faixa e explique do que depende.
- Segurança: quando o assunto envolver eletricidade, peso suspenso, volume alto ou trabalho em altura,
  diga o risco e a conduta segura em uma `div.note`.
- Os exemplos são ilustrativos e se passam em cenários plausíveis (igreja de 200 lugares, auditório
  de escola, evento ao ar livre). Use nomes genéricos de função ("o pastor", "a cantora", "o operador").

## Estrutura do treinamento (`<nome>_body.html`)

De 11 a 13 módulos numerados, mais o glossário. De 60 a 90 KB de HTML (cerca de 5.500 a 7.000 palavras), de 7 a 10 figuras em SVG. Leitura de 35 a 50 minutos.
O estudo `ganho` (`ganho_body.html` e `build_ganho.py`) já está pronto nesta pasta e é a referência de tamanho, tom e acabamento.
A ordem dos módulos é esta; os do meio variam com o assunto:

1. **O que é** (`id="o-que-e"`): definição, para que serve, uma figura que resume o conceito.
2. a 5. (ou 6.) **Módulos de conceito**: como funciona, as partes, os tipos, como ajustar. Um conceito por módulo,
   cada um com figura ou tabela.
6. **Passo a passo** (`ol.steps`): o procedimento para o dia do evento.
7. **Exemplos resolvidos** (`id="exemplos"`): 2 ou 3 situações, cada uma com `dl.ficha` (cenário, sintoma,
   causa, o que foi feito, resultado).
8. **Erros comuns** (`id="erros"`, `div.errs`): 6 a 8 erros, cada um com a correção.
9. **Ligação com os outros estudos** (`id="ligacoes"`): de onde este estudo recebe e para quem entrega,
   com links relativos (`../Pasta/arquivo.html`) para os estudos de `estudos.txt`.
10. **Checklist** (`id="checklist-sec"`, `ul.check#checklist`): 10 a 12 itens para conferir no evento.
11. **Pratique** (`id="pratique"`): exercício 1 de classificação (o quiz) e exercício 2 de bancada,
    com o equipamento na mão, usando a planilha.
12. **A planilha de apoio** (`id="planilha"`): tabela das abas.
13. **Glossário** (`id="glossario"`, `dl.gloss`): 12 a 16 termos, em ordem alfabética. Não é módulo numerado:
    o `p.eyebrow` dele é "Referência", e ele não entra na contagem de módulos do `ul.meta`.

### Esqueleto obrigatório

`montar.py` exige que o corpo comece em `<div class="page" ...>` e termine em `</footer>`, e confere
que cada link do sumário aponta para uma `<section id>` existente.

```html
<div class="page" lang="pt-BR" data-estudo="NOME">

<header class="mast">
  <a class="back" href="../index.html">← Painel dos estudos</a>
  <div class="tiles" aria-hidden="true"><span class="blue">G</span><span class="amber">F</span></div>
  <h1>Título do estudo</h1>
  <p class="lead">Uma frase que diz o que a pessoa aprende e para quê.</p>
  <ul class="meta">
    <li>Treinamento autoinstrucional</li>
    <li>12 módulos</li>
    <li>Leitura de cerca de 40 minutos</li>
    <li>3 exemplos resolvidos e 2 exercícios</li>
  </ul>
</header>

<div class="layout">

<nav class="toc" aria-label="Sumário">
  <h2>Sumário</h2>
  <ol>
    <li><a href="#o-que-e">O que é ...</a></li>
    ...
    <li><a href="#glossario">Glossário</a></li>
  </ol>
</nav>

<main>

<!-- ============================ MÓDULO 1 ============================ -->
<section id="o-que-e">
  <p class="eyebrow">Módulo 1</p>
  <h2>O que é ...</h2>
  <p>...</p>
</section>

...

</main>
</div>

<footer>Material de estudo. Os exemplos e os valores são ilustrativos e servem de ponto de partida. Confira sempre o manual do seu equipamento.</footer>
```

As letras do `div.tiles` e as cores estão definidas no pedido de cada estudo, e são as mesmas do painel.

### Componentes disponíveis (todos já têm estilo em `base.css`)

O melhor exemplo de uso de todos eles é o corpo de um treinamento do projeto que serviu de modelo:
`/home/josemauro/projects/josemauro/Estudos ISO - Auditoria/_geradores/gut_body.html`.
Leia esse arquivo inteiro antes de escrever: a densidade, o tamanho dos parágrafos, o cuidado das figuras
e o tom são a referência. Só as classes de cor mudam (aqui são por nome de cor, veja abaixo).

- Texto: `p`, `h3`, `ul`, `ol`, `strong`, `code`.
- Letra colorida no texto: `<span class="k blue">G</span>`. Com mais de uma letra: `class="k blue w"`.
- Figura: `<figure><div class="scroll"><svg viewBox="0 0 900 H" role="img" aria-label="descrição completa">…</svg></div><figcaption><b>Figura N.</b> …</figcaption></figure>`.
- Tabela: `<div class="tbl"><table><thead>…</thead><tbody>…</tbody></table></div>`. Células numéricas: `td.num`.
  Centralizar: `th.n`, `td.n`. Cabeçalho ou célula colorida: `th.blue`, `td.blue` (e as outras cores).
- Dois quadros lado a lado: `<div class="pair"><div class="box"><h4>…</h4><ul>…</ul></div><div class="box">…</div></div>`.
- Aviso ou ponto de atenção: `<div class="note"><h4>Ponto de atenção</h4><p>…</p></div>`.
- Passo a passo: `<ol class="steps"><li><h4>Título do passo</h4><p>…</p></li>…</ol>`.
- Ficha de exemplo: `<dl class="ficha"><div><dt>Cenário</dt><dd>…</dd></div>…</dl>`.
- Erros: `<div class="errs"><div><h4>Erro</h4><p>Como aparece.</p><p class="fix"><b>Correção:</b> …</p></div>…</div>`.
- Checklist:
  ```html
  <ul class="check" id="checklist">
    <li><label><input type="checkbox" id="ck1"><span>Item</span></label></li>
  </ul>
  <div class="bar"><span><b id="ck-count">0</b> de <b id="ck-total">12</b> itens marcados</span><button class="btn" id="ck-reset" type="button">Limpar marcações</button></div>
  ```
- Glossário: `<dl class="gloss"><div><dt>Termo</dt><dd>Definição.</dd></div>…</dl>`.
- Vídeo (só nos estudos que têm vídeo, indicado no pedido):
  ```html
  <div class="video">
    <video controls preload="metadata" src="../docs/tutoriais/audio/ARQUIVO.mp4"></video>
    <div><h4>Vídeo: título</h4><p>O que o vídeo mostra, em duas ou três frases.</p></div>
  </div>
  ```

### Quiz (exercício 1)

O script lê os dados de um bloco JSON. As chaves de `opcoes` viram os botões (de 2 a 4 opções, chaves
curtas: uma letra ou uma palavra de até 6 letras). De 8 a 10 itens, com respostas bem distribuídas.

```html
<div class="quiz" id="quiz"></div>
<div class="bar"><span>Acertos: <b id="qz-score">0</b> de <b id="qz-total">10</b></span><span>Respondidos: <b id="qz-done">0</b></span><button class="btn" id="qz-reset" type="button">Recomeçar</button></div>
<script type="application/json" id="quiz-data">
{"opcoes": {"G": "Ganho", "F": "Fader"},
 "itens": [
  {"t": "A voz do pastor chega com chiado mesmo com o fader alto.", "a": "G", "w": "Sinal fraco na entrada: o ajuste é no ganho."}
 ]}
</script>
```

### Figuras em SVG

- Sempre `viewBox="0 0 900 H"`, com H entre 180 e 460. Nada fora do viewBox; margem de 18 nas bordas.
- Sem cores fixas: use só as classes, para a figura funcionar no tema claro e no escuro.
  - Texto: `text` (padrão), `.mu` (cinza), `.b` (negrito), `.disp` (título, peso 700), `.mono` (rótulo em caixa alta),
    `.on` (texto sobre cor forte), `.t-ground` (texto sobre `bx-ink`), `.t-blue` … `.t-slate` (texto colorido).
  - Caixas: `.bx` (neutra), `.bx-ink` (escura, destaque), `.bx-out` (tracejada, agrupamento),
    `.bx-blue` … `.bx-slate` (fundo claro e borda da cor), `.hd-blue` … (faixa de cor forte),
    `.f-blue` … (preenchimento forte), `.lv-blue` … (preenchimento translúcido), `.f-sunk`, `.f-surface`, `.f-ink`, `.f-good`, `.f-bad`.
  - Linhas: `.ln` (1,5), `.ln-mu`, `.ln-bold` (3), `.dash`, `.grid`, `.ref`, `.s-blue` … `.s-slate`, `.s-ink`, `.s-good`, `.s-bad` (traço 2, sem preenchimento: use para curvas e ondas).
  - Tela de projeção, telão ou monitor mostrado na figura: `.scr` (fundo) e `.scr-tx` (texto). São as únicas cores fixas:
    a tela fica escura com letra clara nos dois temas, como na sala.
  - Não use `<use>` para repetir desenhos: as regras `svg .classe` do CSS não alcançam o conteúdo clonado, que sai preto.
    Repita os elementos.
  - Cores disponíveis: `blue`, `amber`, `teal`, `purple`, `red`, `slate`.
- Seta colorida: `class="ah f-amber"` no `path` do marcador. Texto verde ou vermelho de situação: `.t-good`, `.t-bad`.
- Seta: defina uma vez por figura, com `id` único no documento:
  `<defs><marker id="ar1" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto"><path class="ah" d="M0 0 L10 5 L0 10 z"/></marker></defs>` e use `marker-end="url(#ar1)"`.
- Texto das figuras: `font-size` de 10,5 a 14. Confira a largura: uma letra ocupa cerca de 0,55 × o tamanho
  da fonte. Um texto de 30 letras em 12 px ocupa perto de 200 px. O texto não pode passar da caixa nem
  encostar em outro texto. Prefira quebrar em duas linhas.
- Toda figura tem `aria-label` que descreve o conteúdo por extenso, e legenda "Figura N." numerada em sequência.
- Desenhe o que ensina: fluxo do sinal em blocos com setas, curva de resposta, vista de cima do palco,
  painel do equipamento em esquema, medidor de nível, linha do tempo. Evite figura decorativa.

## A planilha (`build_<nome>.py`)

De 5 a 7 abas. A planilha é uma ferramenta que o operador usa no evento, e não um resumo do texto.

1. **Instruções**: legenda de cores e ordem de preenchimento (função `instrucoes`).
2. **Uma a três abas de trabalho**: o registro principal do estudo, com células de entrada em amarelo-claro,
   listas suspensas, fórmulas que conferem ou calculam (em cinza) e formatação condicional de situação.
3. **Referência**: a tabela de valores típicos do estudo, para consulta rápida.
4. **Checklist**: os mesmos itens do módulo de checklist (função `checklist`).
5. **Exemplo**: a aba de trabalho preenchida com um dos exemplos resolvidos do treinamento.

Regras:

- Comece com `import os, sys`, `sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))`,
  `from xlsx_lib import *`, `OUT = sys.argv[1]`. Termine com `wb.save(OUT)`.
- Leia `xlsx_lib.py` inteiro antes de escrever: `put`, `title`, `band`, `label`, `head`, `num`, `ex`, `calc`,
  `inp`, `widths`, `setup`, `dv_list`, `dv_date`, `dv_number`, `cf_equal`, `cf_ok`, `note`, `summary`,
  `instrucoes`, `checklist`, e as cores `HUE`, `INK`, `GRAY`, `INPUT`, `GREEN`, `YELLOW`, `RED`.
- A coluna A é margem (largura 2); o conteúdo começa na coluna B. Título na linha 1, subtítulo na linha 2.
- Para gerar a aba de trabalho e a aba de exemplo com a mesma função, escreva `def aba(ws, tab, data=None)`.
- Fórmulas: nomes de função em inglês e separador vírgula (`=IF(C5="","",C5*2)`), como o openpyxl grava.
  Use só funções que o LibreOffice e o Excel calculam igual: `IF`, `AND`, `OR`, `SUM`, `COUNTA`, `COUNTIF`,
  `COUNTIFS`, `SUMIF`, `MAX`, `MIN`, `ROUND`, `IFERROR`, `VLOOKUP`, `INDEX`, `MATCH`, `LOG10`, `ABS`, `TEXT`.
  Toda fórmula devolve `""` quando a entrada está vazia, para a planilha em branco não mostrar erro.
- `cf_ok` pinta de amarelo também a célula vazia: em colunas que podem ficar em branco use `cf_status`.
- `dv_list` não aceita opção com vírgula, e a lista inteira tem limite de 255 caracteres.
- Evite `COUNTIF` com curinga e `TEXT` (dependem do idioma). Em formato de número use o hífen comum: `'+0;-0;0'`.
- Nomes de aba com até 31 caracteres, sem `/ \ ? * [ ] :`.
- Depois de gerar, abra o arquivo recalculado e confira que nenhuma célula tem erro
  (`#NAME?`, `#VALUE!`, `#REF!`, `#DIV/0!`, `Err:`) e que a aba de exemplo mostra os valores esperados.

## Conferência antes de entregar

1. `_geradores/gerar.sh <nome>` roda sem erro.
2. O HTML abre e as figuras estão legíveis: rode
   `_geradores/.venv/bin/python _geradores/capturar.py <Pasta>/treinamento-<slug>.html <pasta temporária>`.
   Ele grava uma imagem PNG por seção (e, com `--escuro`, no tema escuro) e lista os textos de figura que
   saem do viewBox ou se sobrepõem. Olhe todas as imagens com figura no tema claro e pelo menos três no
   escuro, e corrija texto que sai da caixa, sobreposição, figura cortada e seta que cruza texto.
   Depois de corrigir, gere e capture de novo.
3. O número de módulos, de figuras e de itens do checklist no texto (`ul.meta`, `ck-total`, tabela da planilha)
   bate com o que existe de fato.
4. Os links para outros estudos usam as pastas e os arquivos de `estudos.txt`.
5. A planilha recalculada não tem células com erro (o cabeçalho "#" não é erro).
6. `td.num` é monoespaçado e sem quebra de linha: use só em células de número.
