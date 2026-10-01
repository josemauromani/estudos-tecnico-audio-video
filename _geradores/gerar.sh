#!/usr/bin/env bash
# Gera de novo o treinamento (HTML) e a planilha (XLSX) de um estudo.
#
#   _geradores/gerar.sh <estudo> [pasta de saída]
#   _geradores/gerar.sh todos   [pasta de saída]
#
# Os estudos e as pastas de cada um estão em _geradores/estudos.txt.
# Sem a pasta de saída, os arquivos são gravados na pasta do próprio estudo.
set -euo pipefail

AQUI="$(cd "$(dirname "$0")" && pwd)"
RAIZ="$(dirname "$AQUI")"
if [ -n "${PYTHON:-}" ]; then PY="$PYTHON"
elif [ -x "$AQUI/.venv/bin/python" ]; then PY="$AQUI/.venv/bin/python"
else PY="python3"; fi

gerar() {
  local nome="$1" saida="${2:-}" linha pasta html xlsx destino tmp
  linha="$(grep "^$nome|" "$AQUI/estudos.txt" || true)"
  if [ -z "$linha" ]; then echo "Estudo desconhecido: $nome" >&2; exit 1; fi
  IFS='|' read -r _ pasta html xlsx <<< "$linha"
  destino="${saida:-$RAIZ/$pasta}"
  mkdir -p "$destino"
  tmp="$(mktemp -d)"

  # o CSS e o script de todos os treinamentos vêm de base.css e base.js
  "$PY" "$AQUI/montar.py" "$AQUI/${nome}_body.html" "$destino/$html" > /dev/null
  echo "ok  $destino/$html"

  "$PY" "$AQUI/build_${nome}.py" "$tmp/$xlsx" > /dev/null
  # o LibreOffice recalcula as fórmulas e grava os valores, para a planilha abrir já calculada
  soffice --headless --norestore "-env:UserInstallation=file://$tmp/perfil" \
    --convert-to 'xlsx:Calc MS Excel 2007 XML' --outdir "$tmp/saida" "$tmp/$xlsx" > /dev/null 2>&1
  if [ ! -s "$tmp/saida/$xlsx" ]; then echo "Falha ao recalcular $xlsx no LibreOffice" >&2; exit 1; fi
  cp "$tmp/saida/$xlsx" "$destino/$xlsx"
  rm -rf "$tmp"
  echo "ok  $destino/$xlsx"
}

if [ $# -lt 1 ]; then sed -n '2,8p' "$0" | sed 's/^# \{0,1\}//'; exit 1; fi
if [ "$1" = "todos" ]; then
  grep -v '^#' "$AQUI/estudos.txt" | cut -d'|' -f1 | while read -r n; do
    if [ -f "$AQUI/${n}_body.html" ]; then gerar "$n" "${2:-}"; else echo "--  $n ainda não tem corpo"; fi
  done
else
  gerar "$1" "${2:-}"
fi
