#!/usr/bin/env bash
# Página × gerador: gera as especificações de teste pelos dois caminhos e compara os .docx parte a parte
# (XML canonizado, mídia byte a byte). Precisa de Node e do jsdom (npm install jsdom; NODE_PATH se for global).
# Uso: bash testes/pagina.sh [PASTA_DE_SAIDA]
set -u
raiz="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ger="$raiz/skills/timbrador/scripts/gerar_documento.py"
saida="${1:-$raiz/saida/pagina}"
if ! command -v node >/dev/null || ! node -e 'require("jsdom")' 2>/dev/null; then
  echo "pulado: precisa de node e jsdom (npm install jsdom)"; exit 0
fi
rm -rf "$saida"; mkdir -p "$saida/pagina" "$saida/gerador"
python3 "$ger" --embutir-logos "$raiz/testes/marca.teste.json" --saida "$saida/marca.json" >/dev/null || exit 1
# documento de grupo com a caixa "Com assinatura": pelos dois caminhos, os envolvidos viram assinaturas previstas
python3 - "$raiz/testes/espec-curto.json" "$saida/grupo.json" <<'PY0'
import json, sys
d = json.load(open(sys.argv[1], encoding="utf-8"))
d.update(empresa="GTESTE", estado="EM_APROVACAO", com_assinatura=True, assinatura_formato="tabela",
         envolvidos=[{"nome": "Ana Teste", "cargo": "Diretora", "papel": "Aprovou"}])
json.dump(d, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False)
PY0
# telegramas (curto e longo): meia folha A5 no fim do documento, pelos dois caminhos
python3 - "$raiz/testes/espec-curto.json" "$raiz/testes/espec-longo.json" "$saida" <<'PY1'
import json, os, sys
tel = [{"titulo": "Recibo de entrega", "corpo": [{"t": "p", "texto": "Recebi o material em **perfeito estado**."}, {"t": "h2", "texto": "Itens"},
                                                 {"t": "marcadores", "itens": ["um", "dois"]}, {"t": "campos", "itens": [["Recebido por", "Ana Teste"]]}]},
       {"titulo": "Aviso", "corpo": [{"t": "p", "texto": "Texto curto."}]}]
for orig, nome in ((sys.argv[1], "telegrama-curto.json"), (sys.argv[2], "telegrama-longo.json")):
    d = json.load(open(orig, encoding="utf-8"))
    d.update(telegramas=tel, titulo=d["titulo"] + " com telegrama")
    json.dump(d, open(os.path.join(sys.argv[3], nome), "w", encoding="utf-8"), ensure_ascii=False)
    # confidencial: CONFIDENCIAL em todas as páginas (também no telegrama) e a lista de quem pode ler
    d.update(classificacao="Confidencial", acesso_restrito="Diretoria, Ana Teste", titulo=d["titulo"].replace(" com telegrama", " confidencial"))
    json.dump(d, open(os.path.join(sys.argv[3], nome.replace("telegrama", "confidencial")), "w", encoding="utf-8"), ensure_ascii=False)
PY1
specs=("$raiz/testes/espec-curto.json" "$raiz/testes/espec-longo.json" "$saida/grupo.json" "$saida/telegrama-curto.json" "$saida/telegrama-longo.json"
       "$saida/confidencial-curto.json" "$saida/confidencial-longo.json")
node "$raiz/testes/pagina.js" "$saida/marca.json" "$saida/pagina" "${specs[@]}" >/dev/null || { echo "  FALHA página não gerou"; exit 1; }
# sem LibreOffice para o gerador: a página não numera o sumário, então o gerador também não
for s in "${specs[@]}"; do TIMBRADOR_SOFFICE=false python3 "$ger" "$s" --marca "$raiz/testes/marca.teste.json" --saida "$saida/gerador" >/dev/null; done
# marca configurada pelo editor da página (exportada): tem de dar os mesmos documentos
node "$raiz/testes/pagina_editor.js" "$raiz/skills/timbrador/assets/logo-exemplo.png" "$saida/editor" > "$saida/editor.log" 2>&1 \
  || { echo "  FALHA editor da marca"; sed 's/^/        /' "$saida/editor.log"; exit 1; }
node "$raiz/testes/pagina.js" "$saida/editor/marca-exportada.json" "$saida/editor/pagina" "${specs[@]}" >/dev/null || { echo "  FALHA página com a marca do editor"; exit 1; }
# marca com "visual" (fontes embutidas e cores): uma TrueType do sistema faz o papel das fontes da organização
ttf="$(fc-match -f '%{file}' 'DejaVu Serif' 2>/dev/null)"
if [ -n "$ttf" ] && [ "${ttf##*.}" = "ttf" ]; then
  python3 - "$saida/marca.json" "$ttf" "$saida/visual/marca.json" <<'PY2'
import json, sys, os, shutil
cfg = json.load(open(sys.argv[1], encoding="utf8")); os.makedirs(os.path.dirname(sys.argv[3]), exist_ok=True)
shutil.copy(sys.argv[2], os.path.join(os.path.dirname(sys.argv[3]), "fonte.ttf"))
cfg["visual"] = {"fonte_texto": "Teste Texto", "fonte_titulo": "Teste Titulo", "cor_texto": "1C2533", "cor_titulo": "233F61",
                 "cor_secundaria": "5A6372", "cor_realce": "DBE4EF",
                 "fontes": [{"familia": "Teste Texto", "estilo": "regular", "arquivo": "fonte.ttf"},
                            {"familia": "Teste Titulo", "estilo": "negrito", "arquivo": "fonte.ttf"}]}
json.dump(cfg, open(sys.argv[3], "w", encoding="utf8"), ensure_ascii=False)
PY2
  python3 "$ger" --embutir-logos "$saida/visual/marca.json" --saida "$saida/visual/marca_portatil.json" >/dev/null || exit 1
  node "$raiz/testes/pagina.js" "$saida/visual/marca_portatil.json" "$saida/visual/pagina" "${specs[@]}" >/dev/null || { echo "  FALHA página com visual"; exit 1; }
  for s in "${specs[@]}"; do TIMBRADOR_SOFFICE=false python3 "$ger" "$s" --marca "$saida/visual/marca.json" --saida "$saida/visual/gerador" >/dev/null; done
else
  echo "  pulado: visual (sem DejaVu Serif .ttf no sistema)"
fi
falhas=0
for f in "$saida"/visual/pagina/*.docx; do
  [ -e "$f" ] || continue
  n="$(basename "$f")"
  if python3 "$raiz/testes/comparar_docx.py" "$f" "$saida/visual/gerador/$n" > "$saida/.log" &&
     python3 -c "import zipfile,sys; sys.exit('word/fontTable.xml' not in zipfile.ZipFile(sys.argv[1]).namelist())" "$f"; then echo "  ok    com visual e fontes embutidas: $n"
  else echo "  FALHA com visual: $n"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); fi
done
for f in "$saida"/editor/pagina/*.docx; do
  n="$(basename "$f")"
  if python3 "$raiz/testes/comparar_docx.py" "$f" "$saida/gerador/$n" > "$saida/.log"; then echo "  ok    marca do editor: $n"
  else echo "  FALHA marca do editor: $n"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); fi
done
for f in "$saida"/pagina/*.docx; do
  n="$(basename "$f")"
  if python3 "$raiz/testes/comparar_docx.py" "$f" "$saida/gerador/$n" > "$saida/.log"; then echo "  ok    $n"
  else echo "  FALHA $n"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); fi
done
# relatório de condução: o gerador abre e publica duas rodadas; a página lê os arquivos dele, confere o SHA-256 das
# rodadas e acrescenta uma; o gerador relê o resultado da página e aceita a rodada seguinte
rel="$saida/rel"; mkdir -p "$rel/docs" "$rel/pagina"
marca="$raiz/testes/marca.teste.json"
python3 - "$raiz/testes/espec-rel.json" "$rel" <<'PY3'
import json, sys, os
rel = json.load(open(sys.argv[1], encoding="utf-8")); d = sys.argv[2]
json.dump({"empresa": "ACME", "tipo": "PLN", "area": "ECO", "numero": 1, "titulo": "Loja nova no marketplace", "projeto": "Loja Nova",
           "estado": "OFICIAL", "aprovado_por": "Fulano de Tal", "data_aprovacao": "01/10/2026", "elaborado_por": "Fulano de Tal",
           "corpo": [{"t": "h1", "texto": "Objetivo"}, {"t": "p", "texto": "Abrir a loja."}]}, open(f"{d}/plano.json", "w"), ensure_ascii=False)
json.dump({"data": "12/10/2026", "feito": "50 anúncios cadastrados.", "avanco": "sim",
           "compromisso": {"oque": "Ativar frete", "quem": "Logística", "ate": "19/10/2026"}}, open(f"{d}/r1.json", "w"), ensure_ascii=False)
json.dump({"tipo": "encerramento", "data": "20/10/2026", "feito": "Loja aberta.", "termo": "ACME-TER-ECO-002/2026", "situacao": "concluído"},
          open(f"{d}/fim.json", "w"), ensure_ascii=False)
json.dump({"data": "25/10/2026", "feito": "Frete ativado.", "avanco": "sim", "compromisso": {"oque": "Revisar preços", "quem": "Comercial", "ate": "01/11/2026"}},
          open(f"{d}/r3.json", "w"), ensure_ascii=False)
PY3
vig() { ls "$rel/docs"/ACME-REL-ECO-001_v*_OFICIAL_*.docx | sort -V | tail -1; }
{ python3 "$ger" "$rel/plano.json" --marca "$marca" --saida "$rel/docs" --lista "$rel/docs" &&
  python3 "$ger" "$raiz/testes/espec-rel.json" --marca "$marca" --saida "$rel/docs" --lista "$rel/docs" &&
  python3 "$ger" --marca "$marca" --atualizar "$(vig)" --rodada "$rel/r1.json" --json-saida "$rel/v11.json" &&
  python3 "$ger" "$rel/v11.json" --marca "$marca" --saida "$rel/docs" --lista "$rel/docs" && v11="$(vig)" && cp "$v11" "$rel/v11.docx" &&
  python3 "$ger" --marca "$marca" --atualizar "$v11" --rodada "$rel/fim.json" --json-saida "$rel/v12.json" &&
  python3 "$ger" "$rel/v12.json" --marca "$marca" --saida "$rel/docs" --lista "$rel/docs"; } > "$saida/.log" 2>&1 \
  && node "$raiz/testes/pagina_rel.js" "$saida/marca.json" "$rel/v11.docx" "$(vig)" "$rel/docs/lista-mestra.json" "$rel/pagina" >> "$saida/.log" 2>&1 \
  && python3 "$ger" --marca "$marca" --atualizar "$rel/pagina/rel.docx" --rodada "$rel/r3.json" --json-saida "$rel/v13.json" >> "$saida/.log" 2>&1 \
  && python3 -c "import json,sys; d=json.load(open(sys.argv[1])); sys.exit(0 if d['versao']=='1.3' and len(d['rodadas'])==4 else 1)" "$rel/v13.json" \
  && echo "  ok    relatório de condução: página e gerador leem e acrescentam rodadas um do outro" \
  || { echo "  FALHA relatório de condução página × gerador"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); }
# versão da atualização (Tipo de alteração): a página e o gerador dão a mesma versão e o mesmo estado, inclusive na
# minuta de nova versão, em que a correção ou a mudança de conteúdo contam da versão oficial substituída
python3 - "$saida/versao-casos.json" <<'PYV'
import json, sys
fichas = [{"estado": "OFICIAL", "versao": "1.0"}, {"estado": "OFICIAL", "versao": "2.3"},
          {"estado": "MINUTA", "versao": "1.0-r1"}, {"estado": "MINUTA", "versao": "2.0-r1", "substitui": "ACME-POP-ECO-003 v1.0"},
          {"estado": "MINUTA", "versao": "1.1-r2", "substitui": "ACME-POP-ECO-003 v1.0"},
          {"estado": "EM_APROVACAO", "versao": "3.0-r3", "substitui": "ACME-POP-ECO-003 v2.4"},
          {"estado": "RASCUNHO", "versao": "1.0-r2"}]
casos = [[dict(f, codigo="ACME-POP-ECO-003", titulo="Teste"), m] for f in fichas for m in ("rodada", "conteudo", "correcao")]
json.dump(casos, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False)
PYV
if node "$raiz/testes/pagina_versao.js" "$saida/versao-casos.json" > "$saida/versao-pagina.json" 2> "$saida/.log" &&
   python3 - "$ger" "$saida/versao-casos.json" "$saida/versao-pagina.json" > "$saida/.log" 2>&1 <<'PYV2'
import importlib.util, json, sys
spec = importlib.util.spec_from_file_location("ger", sys.argv[1]); ger = importlib.util.module_from_spec(spec); spec.loader.exec_module(ger)
casos, pag = json.load(open(sys.argv[2], encoding="utf-8")), json.load(open(sys.argv[3], encoding="utf-8"))
dif = []
for (ficha, mud), p in zip(casos, pag):
    n, _ = ger.preparar_atualizacao(ficha, mud, "Teste.", None, "Fulano de Tal", "04/11/2026")
    if (n["versao"], n["estado"]) != (p["versao"], p["estado"]):
        dif.append(f'{ficha["estado"]} {ficha["versao"]} {ficha.get("substitui", "")} + {mud}: gerador {n["versao"]} {n["estado"]}, página {p["versao"]} {p["estado"]}')
print("\n".join(dif) or len(casos)); sys.exit(1 if dif else 0)
PYV2
then echo "  ok    versão da atualização: página e gerador iguais ($(cat "$saida/.log") casos)"
else echo "  FALHA versão da atualização página × gerador"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); fi
# prévia em PDF (botão Visualizar): o núcleo da página desenha cada documento; o PDF tem de ser válido (xref no lugar
# certo, uma página por /Page) e trazer o código, a paginação "Página X de Y" e o título
python3 - "$raiz/testes/espec-rel.json" "$saida/rel-numerado.json" <<'PY4'
import json, sys
d = json.load(open(sys.argv[1], encoding="utf-8")); d["numero"] = 1      # na página, o número vem da lista mestra
json.dump(d, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False)
PY4
if node "$raiz/testes/pagina_previa.js" "$saida/marca.json" "$saida/previa" "${specs[@]}" "$saida/rel-numerado.json" > "$saida/.log" 2>&1; then
  for f in "$saida"/previa/*.pdf; do
    if python3 - "$f" > "$saida/.log" 2>&1 <<'PY5'
import re, sys, shutil, subprocess
d = open(sys.argv[1], "rb").read()
assert d.startswith(b"%PDF-1.4") and d.rstrip().endswith(b"%%EOF"), "cabeçalho ou fim do PDF"
xref = int(re.search(rb"startxref\s+(\d+)", d).group(1)); assert d[xref:xref + 4] == b"xref", "startxref fora do lugar"
n = len(re.findall(rb"/Type /Page\b", d)); assert n >= 1, "sem páginas"
if shutil.which("pdftotext"):
    txt = subprocess.run(["pdftotext", "-layout", sys.argv[1], "-"], capture_output=True, text=True, check=True).stdout
    assert f"Página {n} de {n}" in txt, f"paginação ({n} páginas)"
    assert re.search(r"(ACME|GTESTE)-[A-Z]{2,3}-[A-Z]{2,3}-", txt), "código do documento"
print(n)
PY5
    then echo "  ok    prévia em PDF: $(basename "$f") ($(cat "$saida/.log") pág.)"
    else echo "  FALHA prévia em PDF: $(basename "$f")"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); fi
  done
else echo "  FALHA prévia em PDF não gerou"; sed 's/^/        /' "$saida/.log"; falhas=$((falhas + 1)); fi
exit $falhas
