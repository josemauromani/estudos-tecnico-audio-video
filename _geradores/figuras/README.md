# Scripts de apoio das figuras

Os autores de dois estudos calcularam algumas figuras com script, em vez de desenhá-las à mão:

| Arquivo | Estudo | O que gera |
|---|---|---|
| `eq_figs.py` | Equalização | As curvas dos filtros (passa-altas, shelving e bell), calculadas pelas fórmulas reais. |
| `quadro_figs.py` | Enquadramento e movimento | As cenas e as figuras de posição de câmera. |

A fonte do treinamento continua sendo o `<nome>_body.html`, que já traz as figuras prontas e foi revisado
depois de gerado. Use estes scripts só para refazer uma figura: gere o SVG, confira e cole no corpo no
lugar do antigo. Eles podem precisar de ajuste de caminhos, porque foram escritos numa pasta temporária.
