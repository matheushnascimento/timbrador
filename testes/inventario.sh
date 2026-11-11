#!/usr/bin/env bash
# Inventariar.ps1 × gerar_documento.py --inventariar: monta uma pasta de documentos com o gerador (oficial com cópia
# obsoleta, documento de projeto fora da pasta, minuta, relatório de condução com rodadas, entrada antiga da lista sem
# arquivo) e confere que as duas listas mestras e os alertas saem iguais. Precisa do PowerShell 7 (pwsh); pula sem ele.
# Uso: bash testes/inventario.sh [PASTA_DE_SAIDA]
set -u
raiz="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ger="$raiz/skills/timbrador/scripts/gerar_documento.py"
ps1="$raiz/pagina/Inventariar.ps1"
marca="$raiz/testes/marca.teste.json"
saida="${1:-$raiz/saida/inventario}"
pwsh="${TIMBRADOR_PWSH:-pwsh}"
if ! command -v "$pwsh" > /dev/null; then echo "pulado: precisa do PowerShell 7 (pwsh, ou TIMBRADOR_PWSH)"; exit 0; fi
rm -rf "$saida"; mkdir -p "$saida"
docs="$saida/docs"; mkdir -p "$docs/Obsoletos" "$docs/Modelos" "$docs/Projetos/Loja Nova" "$saida/rel"
g() { python3 "$ger" --marca "$marca" "$@" > "$saida/.log" 2>&1 || { echo "  FALHA ao preparar: $*"; sed 's/^/        /' "$saida/.log"; exit 1; }; }

python3 - "$raiz/testes" "$saida" <<'PY'
import json, os, sys
base, d = sys.argv[1], sys.argv[2]
curto = json.load(open(os.path.join(base, "espec-curto.json"), encoding="utf-8"))
def grava(nome, **mud):
    x = json.loads(json.dumps(curto)); x.update(mud); x["corpo"] = [{"t": "p", "texto": "Texto."}]
    json.dump(x, open(os.path.join(d, nome), "w", encoding="utf-8"), ensure_ascii=False)
oficial = dict(estado="OFICIAL", aprovado_por="Fulano de Tal", data_aprovacao="18/09/2026")
grava("oficial.json", numero=5, titulo="Horário de verão", **oficial)
grava("projeto.json", numero=6, titulo="Escala da loja", projeto="Loja Nova", **oficial)
grava("minuta.json", numero=None, titulo="Uniformes novos")
grava("modelo.json", numero=None, titulo="Modelo de comunicação", modelo=True)
grava("confidencial.json", numero=7, titulo="Assunto sigiloso", classificacao="Confidencial", acesso_restrito="Diretoria", **oficial)
json.dump({"empresa": "ACME", "tipo": "PLN", "area": "ECO", "numero": 1, "titulo": "Loja nova no marketplace", "projeto": "Loja Nova",
           "estado": "OFICIAL", "aprovado_por": "Fulano de Tal", "data_aprovacao": "01/10/2026", "elaborado_por": "Fulano de Tal",
           "corpo": [{"t": "h1", "texto": "Objetivo"}, {"t": "p", "texto": "Abrir a loja."}]},
          open(os.path.join(d, "plano.json"), "w", encoding="utf-8"), ensure_ascii=False)
json.dump({"data": "12/10/2026", "feito": "50 anúncios cadastrados.", "avanco": "nao", "travou": "cadastro  do\tfrete", "depende": "Logística",
           "compromisso": {"oque": "Ativar frete", "quem": "Logística", "ate": "19/10/2026"}},
          open(os.path.join(d, "r1.json"), "w", encoding="utf-8"), ensure_ascii=False)
json.dump({"tipo": "correcao", "corrige": 1, "data": "13/10/2026", "feito": "Eram 52 anúncios."},
          open(os.path.join(d, "r2.json"), "w", encoding="utf-8"), ensure_ascii=False)
PY
g "$saida/oficial.json" --saida "$docs"
g --obsoletar "$(ls "$docs"/ACME-CI-RH-005-*.docx)" --substituido-por "ACME-CI-RH-009/2026 v1.0" --saida "$docs/Obsoletos"
g "$saida/projeto.json" --saida "$docs"
# a mesma versão em duas pastas é cópia do mesmo documento, não versão anterior
cp "$docs"/ACME-CI-RH-006-*.docx "$docs/Projetos/Loja Nova/"
g "$saida/minuta.json" --saida "$docs/Projetos/Loja Nova"
g "$saida/modelo.json" --saida "$docs/Modelos"
# confidencial: na pasta Confidencial (fora da varredura) e uma cópia esquecida na pasta comum, que os dois acusam
mkdir -p "$docs/Confidencial"
g "$saida/confidencial.json" --saida "$docs/Confidencial"
cp "$docs/Confidencial"/ACME-CI-RH-007-*.docx "$docs/"
# relatório de condução: plano e três versões (a lista do relatório fica à parte, só para publicar)
g "$saida/plano.json" --saida "$saida/rel" --lista "$saida/rel"
g "$raiz/testes/espec-rel.json" --saida "$saida/rel" --lista "$saida/rel"
vig() { ls "$saida/rel"/ACME-REL-ECO-001_v*_OFICIAL_*.docx | sort -V | tail -1; }
for r in r1 r2; do
  g --atualizar "$(vig)" --rodada "$saida/$r.json" --json-saida "$saida/$r-espec.json"
  g "$saida/$r-espec.json" --saida "$saida/rel" --lista "$saida/rel"
done
cp "$saida/rel"/ACME-PLN-*.docx "$saida/rel"/ACME-REL-*.docx "$docs/Projetos/Loja Nova/"
# lista que já existia: registra uma entrada cujo arquivo saiu da pasta e outra de uma pasta não varrida
python3 - "$docs/lista-mestra.json" <<'PY'
import json, sys
json.dump({"lista_mestra": "Lista de Teste", "padrao": "ACME-POL-ADM-001", "atualizado_em": "", "documentos": [
    {"codigo": "ACME-CI-RH-001/2026", "titulo": "Saiu da pasta", "empresa": "ACME", "tipo": "CI", "area": "RH", "numero": 1, "ano": 2026,
     "versao": "1.0", "estado": "OFICIAL", "arquivo": "x.docx", "pasta": ".", "registrado_em": "01/09/2026 10:00"},
    {"codigo": "ACME-CI-RH-002/2026", "titulo": "Em outra pasta", "empresa": "ACME", "tipo": "CI", "area": "RH", "numero": 2, "ano": 2026,
     "versao": "1.0", "estado": "OFICIAL", "arquivo": "y.docx", "pasta": "Arquivo morto", "registrado_em": "01/09/2026 10:00"},
    {"codigo": "ACME-POL-ADM-009", "titulo": "Documento confidencial", "empresa": "ACME", "tipo": "POL", "area": "ADM", "numero": 9,
     "versao": "2.0", "estado": "OFICIAL", "confidencial": True, "pasta": "Confidencial/Políticas (POL)", "registrado_em": "01/09/2026 10:00"}]},
    open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=2)
PY
cp "$docs/lista-mestra.json" "$saida/lista-antes.json"

# os dois partem da mesma lista existente
cp "$saida/lista-antes.json" "$saida/lista-ps1.json"
cp "$saida/lista-antes.json" "$saida/lista-gerador.json"
python3 "$ger" --marca "$marca" --inventariar "$docs" --lista "$saida/lista-gerador.json" > "$saida/gerador.txt" 2>&1
"$pwsh" -NoProfile -File "$ps1" -Pasta "$docs" -Marca "$marca" -Lista "$saida/lista-ps1.json" > "$saida/ps1.txt" 2>&1
python3 - "$saida" <<'PY'
import json, re, sys
d = sys.argv[1]; falhas = 0
def lista(n):
    l = json.load(open(f"{d}/{n}", encoding="utf-8-sig"))
    docs = {x["codigo"] + "|" + (x.get("arquivo") or ""): {k: v for k, v in x.items() if k not in ("registrado_em", "atualizado_em")} for x in l["documentos"]}
    return l["lista_mestra"], l["padrao"], docs
a, b = lista("lista-gerador.json"), lista("lista-ps1.json")
if a[:2] != b[:2]: print("  FALHA cabeçalho da lista:", a[:2], b[:2]); falhas += 1
for k in sorted(set(a[2]) | set(b[2])):
    if a[2].get(k) != b[2].get(k):
        falhas += 1; print(f"  FALHA {k}\n        gerador: {a[2].get(k)}\n        ps1    : {b[2].get(k)}")
if not falhas: print(f"  ok    lista mestra igual ({len(a[2])} entradas)")
sem = lambda t: re.sub(r"[^a-z0-9]+", " ", __import__("unicodedata").normalize("NFD", t.lower()).encode("ascii", "ignore").decode()).strip()
def alertas(n, marca):
    return sorted(sem(l.split(marca, 1)[1]) for l in open(f"{d}/{n}", encoding="utf-8").read().splitlines() if marca in l)
ga, pa = alertas("gerador.txt", "ATENÇÃO — "), alertas("ps1.txt", "ATENCAO - ")
if ga == pa and ga: print(f"  ok    mesmos alertas ({len(ga)})")
else:
    falhas += 1; print("  FALHA alertas diferentes"); [print("        gerador:", x) for x in ga]; [print("        ps1    :", x) for x in pa]
for n, frase in (("removidas", "sairam da lista"), ("não conferidas", "sem conferir")):
    g_ = sem(open(f"{d}/gerador.txt", encoding="utf-8").read()); p_ = sem(open(f"{d}/ps1.txt", encoding="utf-8").read())
    if (frase.replace("sairam", "saíram") in open(f"{d}/gerador.txt", encoding="utf-8").read().lower() or sem(frase) in g_) and sem(frase) in p_: print(f"  ok    {n} avisadas nos dois")
    else: falhas += 1; print(f"  FALHA {n}")
sys.exit(falhas)
PY
r1=$?
# rodada publicada adulterada no arquivo (ficha e texto), com a lista já registrada: os dois detectam e a lista
# mantém o SHA-256 registrado
vigente="$(ls "$docs/Projetos/Loja Nova"/ACME-REL-ECO-001_v1.2_*.docx)"
python3 - "$vigente" <<'PY'
import sys, zipfile
zi = zipfile.ZipFile(sys.argv[1]); itens = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as zo:
    for i, b in itens:
        if i.filename.startswith("customXml/item") or i.filename == "word/document.xml":
            b = b.replace("50 anúncios".encode(), "60 anúncios".encode())
        zo.writestr(i, b)
PY
cp "$saida/lista-gerador.json" "$saida/lista-ps1.json"
python3 "$ger" --marca "$marca" --inventariar "$docs" --lista "$saida/lista-gerador.json" > "$saida/gerador.txt" 2>&1
"$pwsh" -NoProfile -File "$ps1" -Pasta "$docs" -Marca "$marca" -Lista "$saida/lista-ps1.json" > "$saida/ps1.txt" 2>&1
python3 - "$saida" <<'PY'
import json, sys
d = sys.argv[1]; falhas = 0
rel = lambda n: [x for x in json.load(open(f"{d}/{n}", encoding="utf-8-sig"))["documentos"] if x["codigo"] == "ACME-REL-ECO-001"][0]["rodadas_sha256"]
if rel("lista-gerador.json") == rel("lista-ps1.json"): print("  ok    adulterada: a lista mantém os SHA-256 registrados nos dois")
else: falhas += 1; print("  FALHA adulterada: SHA-256 diferentes na lista")
g = "não conferem com as registradas" in open(f"{d}/gerador.txt", encoding="utf-8").read()
p = "nao conferem com as registradas" in open(f"{d}/ps1.txt", encoding="utf-8").read()
if g and p: print("  ok    adulterada: os dois avisam a rodada publicada alterada")
else: falhas += 1; print(f"  FALHA adulterada: aviso no gerador={g}, no ps1={p}")
sys.exit(falhas)
PY
r2=$?
exit $(( (r1 + r2) > 0 ))
