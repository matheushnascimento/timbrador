#!/usr/bin/env bash
# Teste de fumaça do Timbrador: confere a marca, gera documentos curtos e longos, testa as travas de estado,
# a lista mestra e a atualização de versão, valida o XML e compara o mesmo documento gerado pelo gerador e
# pela página (esta última parte só roda com Node e Playwright instalados).
# Uso: bash testes/fumaca.sh [PASTA_DE_SAIDA]
set -u

raiz="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ger="$raiz/skills/timbrador/scripts/gerar_documento.py"
pagina="$raiz/skills/timbrador/assets/timbrador.html"
marca="$raiz/testes/marca.teste.json"
saida="${1:-$raiz/saida}"
mkdir -p "$saida"
rm -f "$saida"/*.docx "$saida"/*.pdf

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

tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

echo "== marca =="
espera_ok "checar-marca" python3 "$ger" --checar-marca "$marca"
python3 "$ger" --checar-marca "$marca" | sed 's/^/  /'
espera_ok "embutir-logos" python3 "$ger" --embutir-logos "$marca" --saida "$tmp/marca_portatil.json"
espera_ok "montar-pagina" python3 "$ger" --montar-pagina "$marca" --saida "$tmp/pagina-marca.html"

python3 - "$raiz/testes" "$tmp" <<'PY'
import json, sys, os
base, tmp = sys.argv[1], sys.argv[2]
curto = json.load(open(os.path.join(base, "espec-curto.json"), encoding="utf-8"))
longo = json.load(open(os.path.join(base, "espec-longo.json"), encoding="utf-8"))

def grava(nome, espec, **mud):
    d = json.loads(json.dumps(espec)); d.update(mud)
    json.dump(d, open(os.path.join(tmp, nome), "w", encoding="utf-8"), ensure_ascii=False)

grava("grupo.json", curto, empresa="GT", titulo="Recesso de fim de ano")
grava("sem-logo.json", curto, empresa="SEMLG", estado="EM_APROVACAO", titulo="Escala de plantão")
grava("sem-aprovador.json", longo, aprovado_por="", data_aprovacao="")
grava("sem-numero.json", longo, numero="")
grava("empresa-incompleta.json", longo, empresa="SEMLG")
grava("obsoleto-sem-motivo.json", longo, estado="OBSOLETO", substituido_por="", motivo_obsoleto="")
grava("empresa-desconhecida.json", curto, empresa="NADA")
grava("lista-1.json", curto, estado="OFICIAL", numero=1, aprovado_por="Fulano de Tal", data_aprovacao="18/09/2026",
      titulo="Primeiro comunicado", corpo=[{"t": "p", "texto": "Texto."}])
grava("lista-2.json", curto, estado="OFICIAL", numero=1, aprovado_por="Fulano de Tal", data_aprovacao="18/09/2026",
      titulo="Outro assunto", corpo=[{"t": "p", "texto": "Texto."}])
# mesma especificação para os dois caminhos: sem assinaturas, que só a página põe (caixa "Com assinatura")
grava("cmp-curto.json", curto, corpo=[b for b in curto["corpo"] if b["t"] not in ("assinaturas", "local_data")])
PY

echo "== geração =="
espera_ok "curto (CI, MINUTA)"            python3 "$ger" "$raiz/testes/espec-curto.json" --marca "$marca" --saida "$saida"
espera_ok "longo (POP, OFICIAL)"          python3 "$ger" "$raiz/testes/espec-longo.json" --marca "$marca" --saida "$saida"
if grep -q "Sumário sem números" "$saida/.log"; then echo "  aviso: sumário sem números de página (LibreOffice ou pdftotext fora do PATH)"; fi
espera_ok "grupo (logos das empresas)"    python3 "$ger" "$tmp/grupo.json"    --marca "$marca" --saida "$saida"
espera_ok "sem logo (EM_APROVACAO)"       python3 "$ger" "$tmp/sem-logo.json" --marca "$marca" --saida "$saida"

echo "== travas de estado =="
espera_recusa "OFICIAL sem aprovador"          "aprovado_por"    python3 "$ger" "$tmp/sem-aprovador.json"       --marca "$marca" --saida "$tmp"
espera_recusa "OFICIAL sem número"             "número"          python3 "$ger" "$tmp/sem-numero.json"          --marca "$marca" --saida "$tmp"
espera_recusa "OFICIAL de empresa incompleta"  "razão social"    python3 "$ger" "$tmp/empresa-incompleta.json"  --marca "$marca" --saida "$tmp"
espera_recusa "OBSOLETO sem substituto/motivo" "substituido_por" python3 "$ger" "$tmp/obsoleto-sem-motivo.json" --marca "$marca" --saida "$tmp"
espera_recusa "empresa fora do marca.json"     "desconhecida"    python3 "$ger" "$tmp/empresa-desconhecida.json" --marca "$marca" --saida "$tmp"

echo "== lista mestra e versões =="
mkdir -p "$tmp/docs"
espera_ok     "publica e registra na lista"         python3 "$ger" "$tmp/lista-1.json" --marca "$marca" --saida "$tmp/docs" --lista "$tmp/docs"
espera_recusa "número já usado por outro título"    "já é"  python3 "$ger" "$tmp/lista-2.json" --marca "$marca" --saida "$tmp/docs" --lista "$tmp/docs"
oficial="$(ls "$tmp"/docs/*_OFICIAL_*.docx | head -1)"
espera_ok     "lê o documento timbrado (--ler)"     python3 "$ger" --marca "$marca" --ler "$oficial" --json-saida "$tmp/lido.json"
espera_ok     "prepara a nova versão (--atualizar)" python3 "$ger" --marca "$marca" --atualizar "$oficial" --o-que-mudou "Prazo ajustado." --json-saida "$tmp/nova.json"
if python3 -c "import json,sys; d=json.load(open(sys.argv[1])); sys.exit(0 if d['versao']=='2.0-r1' and d['estado']=='MINUTA' else 1)" "$tmp/nova.json"; then
  ok "nova versão é 2.0-r1 em minuta"; else erro "nova versão deveria ser 2.0-r1 em minuta"; fi
espera_ok     "gera a nova versão"                  python3 "$ger" "$tmp/nova.json" --marca "$marca" --saida "$tmp/docs"

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

echo "== página × gerador =="
if command -v node > /dev/null && NODE_PATH="$(npm root -g 2>/dev/null)" node -e 'require("playwright")' 2> /dev/null; then
  for espec in "$tmp/cmp-curto.json" "$raiz/testes/espec-longo.json"; do
    nome="$(basename "$espec" .json)"
    mkdir -p "$tmp/py-$nome" "$tmp/js-$nome"
    python3 "$ger" "$espec" --marca "$marca" --saida "$tmp/py-$nome" > /dev/null 2>&1
    if NODE_PATH="$(npm root -g)" node "$raiz/testes/pagina.js" "$tmp/pagina-marca.html" "$pagina" "$tmp/marca_portatil.json" "$espec" "$tmp/js-$nome" > "$saida/.log" 2>&1 \
       && python3 "$raiz/testes/comparar_docx.py" "$tmp/py-$nome"/*.docx "$tmp/js-$nome"/*.docx >> "$saida/.log"; then
      ok "$nome: mesmo documento pelos dois caminhos"
    else
      erro "$nome: página e gerador divergem"; sed 's/^/        /' "$saida/.log"
    fi
  done
else
  echo "  (pulado: precisa de Node com Playwright)"
fi

rm -f "$saida/.log"
echo
if [ "$falhas" -eq 0 ]; then echo "Tudo certo. Saída em: $saida"; else echo "$falhas falha(s)."; fi
exit $((falhas > 0))
