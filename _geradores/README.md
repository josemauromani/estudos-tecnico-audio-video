# Geradores dos estudos

Esta pasta guarda os arquivos que geram o painel, os treinamentos (HTML) e as planilhas (XLSX). Quem só lê
os estudos não precisa dela: basta abrir o `index.html` da pasta principal.

## O que há aqui

| Arquivo | Para que serve |
|---|---|
| `estudos.txt` | A lista dos estudos: nome curto, pasta, arquivo HTML e arquivo XLSX. |
| `CURRICULO.md` | O que cada estudo cobre, as letras e as cores do cabeçalho e a ideia da planilha. |
| `GUIA-DO-AUTOR.md` | As regras para escrever um estudo: tom, estrutura, componentes, figuras e planilha. |
| `<nome>_body.html` | O corpo do treinamento de um estudo, escrito à mão. |
| `build_<nome>.py` | O gerador da planilha de um estudo. |
| `base.css`, `base.js` | A aparência e o comportamento (checklist e exercício) de todos os treinamentos. |
| `montar.py` | Junta o corpo, o CSS e o script em um HTML único. |
| `xlsx_lib.py` | Funções e estilos compartilhados pelos geradores de planilha. |
| `index_dados.py`, `index.css`, `build_index.py` | Os textos, a aparência e o gerador do painel. |
| `gerar.sh` | Gera um estudo, ou todos. |
| `capturar.py` | Captura cada seção de um treinamento em PNG, nos temas claro e escuro, e acusa texto de figura fora do lugar. |
| `figuras/` | Scripts que calcularam figuras de dois estudos (curvas de EQ, cenas de enquadramento). |

## Preparação

Os scripts precisam de Python 3, da biblioteca `openpyxl` e do LibreOffice (comando `soffice`).

```bash
python3 -m venv _geradores/.venv
_geradores/.venv/bin/pip install -r _geradores/requirements.txt
```

## Como gerar

```bash
_geradores/gerar.sh ganho              # gera um estudo, na pasta dele
_geradores/gerar.sh todos              # gera todos os que têm corpo
_geradores/gerar.sh ganho /tmp/teste   # gera em outra pasta, para conferir antes
_geradores/.venv/bin/python _geradores/build_index.py index.html   # gera o painel
```

O painel conta os módulos, os minutos, as figuras e as abas nos arquivos de cada estudo. Gere o painel de
novo depois de alterar um estudo.

## Como alterar um estudo

1. Texto e figuras do treinamento: edite `<nome>_body.html`.
2. Planilha: edite `build_<nome>.py`.
3. Gere em uma pasta de teste, confira, e só então gere na pasta do estudo.
4. Gere o painel.

## Como criar um estudo

1. Acrescente a linha em `estudos.txt` e a entrada em `CURRICULO.md`.
2. Acrescente o estudo em `index_dados.py` (cartão, necessidade, o que recebe e entrega, fases e trilhas)
   e, se ele entrar no mapa, a caixa em `build_index.py`.
3. Escreva `<nome>_body.html` e `build_<nome>.py` seguindo o `GUIA-DO-AUTOR.md`.
4. Gere o estudo e o painel.

## Pontos de atenção

- **CSS e script compartilhados:** uma alteração em `base.css` ou em `base.js` muda todos os treinamentos na
  próxima geração.
- **Recálculo:** o `openpyxl` grava as fórmulas sem os valores. O LibreOffice abre o arquivo, calcula e grava
  de novo, para a planilha abrir já calculada.
- **Vídeos:** os treinamentos de Ganho e fader e de Microfonia apontam para os arquivos de
  `docs/tutoriais/audio/`. Se um arquivo mudar de nome, corrija o corpo do estudo e `build_index.py`.
