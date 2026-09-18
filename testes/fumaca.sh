#!/usr/bin/env bash
# Teste de fumaça do Timbrador: confere a marca, gera um documento curto e um
# longo e verifica que as travas de estado recusam o que devem recusar.
# Uso: bash testes/fumaca.sh [PASTA_DE_SAIDA]
set -u

raiz="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ger="$raiz/skills/timbrador/scripts/gerar_documento.py"
marca="$raiz/testes/marca.teste.json"
saida="${1:-$raiz/saida}"
mkdir -p "$saida"

falhas=0
ok()   { printf '  ok    %s\n' "$1"; }
erro() { printf '  FALHA %s\n' "$1"; falhas=$((falhas + 1)); }

# Espera sucesso (código 0).
espera_ok() {
  local nome="$1"; shift
  if "$@" > "$saida/.log" 2>&1; then ok "$nome"; else erro "$nome"; sed 's/^/        /' "$saida/.log"; fi
}

# Espera recusa (código != 0) com uma mensagem contendo o trecho informado.
espera_recusa() {
  local nome="$1" trecho="$2"; shift 2
  if "$@" > "$saida/.log" 2>&1; then
    erro "$nome (deveria ter sido recusado)"
  elif grep -q "$trecho" "$saida/.log"; then
    ok "$nome"
  else
    erro "$nome (recusou pelo motivo errado)"; sed 's/^/        /' "$saida/.log"
  fi
}

echo "== marca =="
espera_ok "checar-marca" python3 "$ger" --checar-marca "$marca"
python3 "$ger" --checar-marca "$marca" | sed 's/^/  /'

echo "== geração =="
espera_ok "curto (CI, MINUTA)"  python3 "$ger" "$raiz/testes/espec-curto.json" --marca "$marca" --saida "$saida"
espera_ok "longo (POP, OFICIAL)" python3 "$ger" "$raiz/testes/espec-longo.json" --marca "$marca" --saida "$saida"

echo "== travas de estado =="
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

python3 - "$raiz/testes/espec-longo.json" "$tmp" <<'PY'
import json, sys, os
espec = json.load(open(sys.argv[1], encoding="utf-8"))
tmp = sys.argv[2]

def grava(nome, **mud):
    d = dict(espec); d.update(mud)
    json.dump(d, open(os.path.join(tmp, nome), "w", encoding="utf-8"), ensure_ascii=False)

grava("sem-aprovador.json", aprovado_por="", data_aprovacao="")
grava("sem-numero.json", numero="")
grava("empresa-incompleta.json", empresa="SEMLG")
grava("obsoleto-sem-motivo.json", estado="OBSOLETO", substituido_por="", motivo_obsoleto="")
PY

espera_recusa "OFICIAL sem aprovador"        "aprovado_por"  python3 "$ger" "$tmp/sem-aprovador.json"      --marca "$marca" --saida "$tmp"
espera_recusa "OFICIAL sem número"           "número"        python3 "$ger" "$tmp/sem-numero.json"         --marca "$marca" --saida "$tmp"
espera_recusa "OFICIAL de empresa incompleta" "razão social" python3 "$ger" "$tmp/empresa-incompleta.json" --marca "$marca" --saida "$tmp"
espera_recusa "OBSOLETO sem substituto/motivo" "substituido_por" python3 "$ger" "$tmp/obsoleto-sem-motivo.json" --marca "$marca" --saida "$tmp"

echo "== XML bem formado =="
for f in "$saida"/*.docx; do
  if python3 - "$f" <<'PY'
import sys, zipfile
from xml.dom.minidom import parseString
z = zipfile.ZipFile(sys.argv[1])
for n in z.namelist():
    if n.endswith(".xml") or n.endswith(".rels"):
        parseString(z.read(n))
PY
  then ok "$(basename "$f")"; else erro "$(basename "$f")"; fi
done

rm -f "$saida/.log"
echo
if [ "$falhas" -eq 0 ]; then echo "Tudo certo. Saída em: $saida"; else echo "$falhas falha(s)."; fi
exit $((falhas > 0))
