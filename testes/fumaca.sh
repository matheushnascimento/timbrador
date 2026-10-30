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

echo "== relatório de condução =="
# REL com rodadas: número herdado do plano oficial do projeto, uma rodada por versão, SHA-256 das rodadas na lista
mkdir -p "$tmp/rel"
python3 - "$raiz/testes" "$tmp" <<'PY'
import json, sys, os
base, tmp = sys.argv[1], sys.argv[2]
plano = {"empresa": "ACME", "tipo": "PLN", "area": "ECO", "numero": 1, "titulo": "Loja nova no marketplace", "projeto": "Loja Nova",
         "estado": "OFICIAL", "aprovado_por": "Fulano de Tal", "data_aprovacao": "01/10/2026", "data_emissao": "01/10/2026",
         "elaborado_por": "Fulano de Tal", "corpo": [{"t": "h1", "texto": "Objetivo"}, {"t": "p", "texto": "Abrir a loja."}]}
json.dump(plano, open(os.path.join(tmp, "rel-plano.json"), "w", encoding="utf-8"), ensure_ascii=False)
rel = json.load(open(os.path.join(base, "espec-rel.json"), encoding="utf-8"))
json.dump(dict(rel, projeto="Sem Plano"), open(os.path.join(tmp, "rel-sem-plano.json"), "w", encoding="utf-8"), ensure_ascii=False)
rodadas = {"r1": {"data": "12/10/2026", "feito": "50 anúncios cadastrados.", "avanco": "sim",
                  "compromisso": {"oque": "Ativar frete", "quem": "Logística", "ate": "19/10/2026"}},
           "rx": {"n": 7, "data": "13/10/2026", "feito": "x", "avanco": "sim", "compromisso": {"oque": "a", "quem": "b", "ate": "c"}},
           "fim": {"tipo": "encerramento", "data": "20/10/2026", "feito": "Loja aberta.", "termo": "ACME-TER-ECO-002/2026", "situacao": "concluído"},
           "depois": {"data": "21/10/2026", "feito": "x", "avanco": "sim", "compromisso": {"oque": "a", "quem": "b", "ate": "c"}}}
for k, v in rodadas.items():
    json.dump(v, open(os.path.join(tmp, f"rodada-{k}.json"), "w", encoding="utf-8"), ensure_ascii=False)
PY
docs="$tmp/rel"; rel_vigente() { ls "$docs"/ACME-REL-ECO-001_v*_OFICIAL_*.docx | sort -V | tail -1; }
espera_ok     "plano oficial do projeto"            python3 "$ger" "$tmp/rel-plano.json" --marca "$marca" --saida "$docs" --lista "$docs"
espera_recusa "projeto sem plano oficial"           "não tem plano"  python3 "$ger" "$tmp/rel-sem-plano.json" --marca "$marca" --saida "$tmp" --lista "$docs"
espera_ok     "abertura (v1.0, número do plano)"    python3 "$ger" "$raiz/testes/espec-rel.json" --marca "$marca" --saida "$docs" --lista "$docs"
espera_recusa "abertura repetida"                   "já existe"      python3 "$ger" "$raiz/testes/espec-rel.json" --marca "$marca" --saida "$tmp" --lista "$docs"
espera_ok     "prepara a rodada 1 (--rodada)"       python3 "$ger" --marca "$marca" --atualizar "$(rel_vigente)" --rodada "$tmp/rodada-r1.json" --json-saida "$tmp/rel-v11.json"
espera_ok     "publica a v1.1"                      python3 "$ger" "$tmp/rel-v11.json" --marca "$marca" --saida "$docs" --lista "$docs"
espera_recusa "rodada fora de ordem"                "rodada nova é a 2" python3 "$ger" --marca "$marca" --atualizar "$(rel_vigente)" --rodada "$tmp/rodada-rx.json"
espera_recusa "rodada a partir de versão antiga"    "já registra 2"  python3 "$ger" "$tmp/rel-v11.json" --marca "$marca" --saida "$tmp" --lista "$docs"
espera_recusa "--obsoletar recusa o relatório"      "não tem cópia obsoleta" python3 "$ger" --marca "$marca" --obsoletar "$(rel_vigente)" --substituido-por x --saida "$tmp"
python3 - "$(rel_vigente)" "$tmp/rel-adulterado.docx" <<'PY'
import sys, zipfile
zi, zo = zipfile.ZipFile(sys.argv[1]), zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED)
for i in zi.infolist():
    b = zi.read(i.filename)
    if i.filename.startswith("customXml/item") or i.filename == "word/document.xml":
        b = b.replace("50 anúncios".encode(), "60 anúncios".encode())
    zo.writestr(i, b)
zo.close()
PY
python3 "$ger" --marca "$marca" --atualizar "$tmp/rel-adulterado.docx" --rodada "$tmp/rodada-fim.json" --json-saida "$tmp/rel-adulterado.json" > /dev/null 2>&1
espera_recusa "rodada publicada alterada"           "não confere com a registrada" python3 "$ger" "$tmp/rel-adulterado.json" --marca "$marca" --saida "$tmp" --lista "$docs"
espera_ok     "encerramento (v1.2)"                 python3 "$ger" --marca "$marca" --atualizar "$(rel_vigente)" --rodada "$tmp/rodada-fim.json" --json-saida "$tmp/rel-v12.json"
espera_ok     "publica o encerramento"              python3 "$ger" "$tmp/rel-v12.json" --marca "$marca" --saida "$docs" --lista "$docs"
espera_recusa "encerrado não recebe rodada"         "foi encerrado"  python3 "$ger" --marca "$marca" --atualizar "$(rel_vigente)" --rodada "$tmp/rodada-depois.json"
if python3 -c "
import json,sys; d=[x for x in json.load(open(sys.argv[1]))['documentos'] if x['codigo']=='ACME-REL-ECO-001'][0]
sys.exit(0 if d['versao']=='1.2' and len(d['rodadas_sha256'])==3 and d['plano']=='ACME-PLN-ECO-001' else 1)" "$docs/lista-mestra.json"; then
  ok "lista: v1.2, plano e 3 SHA-256"; else erro "lista do relatório"; fi
if ls "$docs" | grep -q OBSOLETO; then erro "relatório gerou cópia obsoleta"; else ok "sem cópia obsoleta"; fi
cp "$(rel_vigente)" "$saida/"

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
  # relatório de condução: a página lê o REL do gerador, confere as rodadas e publica uma; o gerador relê o resultado
  v11="$(ls "$docs"/ACME-REL-ECO-001_v1.1_OFICIAL_*.docx)"
  mkdir -p "$tmp/js-rel"
  if NODE_PATH="$(npm root -g)" node "$raiz/testes/pagina_rel.js" "$tmp/pagina-marca.html" "$v11" "$(rel_vigente)" "$docs/lista-mestra.json" "$tmp/js-rel" > "$saida/.log" 2>&1 \
     && python3 "$ger" --marca "$marca" --atualizar "$tmp/js-rel/rel.docx" --rodada "$tmp/rodada-depois.json" --json-saida "$tmp/js-rel/v13.json" >> "$saida/.log" 2>&1; then
    ok "relatório: página e gerador leem e acrescentam rodadas um do outro"
  else
    erro "relatório: página × gerador"; sed 's/^/        /' "$saida/.log"
  fi
else
  echo "  (pulado: precisa de Node com Playwright)"
fi

rm -f "$saida/.log"
echo
if [ "$falhas" -eq 0 ]; then echo "Tudo certo. Saída em: $saida"; else echo "$falhas falha(s)."; fi
exit $((falhas > 0))
