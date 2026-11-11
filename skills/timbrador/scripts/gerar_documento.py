#!/usr/bin/env python3
"""Timbrador white label: gera documentos no Padrão de Documentação Oficial (ABNT + papel timbrado).

Uso:
  python3 gerar_documento.py especificacao.json --marca marca.json [--saida PASTA|arquivo.docx] [--pdf]
  python3 gerar_documento.py --checar-marca marca.json
  python3 gerar_documento.py --embutir-logos marca.json [--saida marca_portatil.json]
  python3 gerar_documento.py --ler | --atualizar | --obsoletar DOCX --marca marca.json ...
  python3 gerar_documento.py --inventariar PASTA --marca marca.json [--lista ...] [--conferir]

Tudo o que varia entre organizações (empresas, razão social, documento, cidade, cor, logos, áreas, tipos
ativos, código da política) vem do marca.json; nenhum dado de organização fica neste arquivo. Só usa a
biblioteca padrão do Python; LibreOffice (soffice, ou o comando em TIMBRADOR_SOFFICE) e pdftotext são
opcionais e servem para numerar as páginas do sumário e gerar o PDF de conferência.
"""
import json, re, sys, os, zipfile, shutil, hashlib, subprocess, tempfile, argparse, datetime, unicodedata, base64, shlex, struct
from xml.sax.saxutils import escape as _escape


def escape(s):
    """Texto para XML, dentro de elemento ou de atributo entre aspas duplas (como o esc da página)."""
    return _escape(s, {'"': "&quot;"})

EMU_CM = 360000
TW_CM = 567
INDEFINIDO = "[definir]"
# Cadastro das empresas, preenchido por carregar_marca() a partir do marca.json. Campo vazio ou "[definir]"
# vira None: aparece como [definir] e impede o estado OFICIAL.
EMPRESAS = {}
# Empresas oferecidas para documento novo (as com "ativa": false só servem para ler e versionar os antigos).
EMPRESAS_ATIVAS = []
# Empresas cujas logos formam o rodapé do documento do grupo, na ordem do marca.json.
EMPRESAS_DO_GRUPO = []
EMPRESA_ID = None
EMPRESA = {}
# Dados da organização (marca.json): nome, código da política, empresa padrão, áreas e tipos ativos.
ORG = {"organizacao": "", "politica": "", "padrao": None, "areas_ativas": [], "tipos_ativos": []}
# Visual opcional da organização ("visual" no marca.json): fontes e cores do documento. Sem ele, o documento
# sai no visual do padrão original (Arial, texto preto, cinzas de apoio). As cores dos estados não mudam nunca.
VISUAL = {}
ESTILOS_FONTE = ("regular", "negrito", "italico", "negrito_italico")


def sec(padrao):
    """Cor de apoio (código, identificação, legendas): a cor_secundaria do visual, ou o cinza do original."""
    return VISUAL.get("cor_secundaria") or padrao


def cor_texto():
    return VISUAL.get("cor_texto") or "000000"


def realce():
    """Fundo de cabeçalho de tabela e de rótulo de quadro."""
    return VISUAL.get("cor_realce") or "F2F2F2"


def definir_empresa(chave):
    global EMPRESA_ID, EMPRESA
    if chave not in EMPRESAS:
        raise SystemExit(f"Empresa desconhecida: {chave}. Use uma de: {', '.join(EMPRESAS)}.")
    EMPRESA_ID, EMPRESA = chave, EMPRESAS[chave]


def dado(campo):
    return EMPRESA.get(campo) or INDEFINIDO


def identificacao_empresa():
    if EMPRESA.get("grupo"):
        return EMPRESA["marca"]
    partes = [EMPRESA["razao"] or EMPRESA["marca"], f'{EMPRESA.get("documento_rotulo") or "CNPJ"} {dado("documento")}']
    if EMPRESA.get("cidade"): partes.append(EMPRESA["cidade"])
    return "  ·  ".join(partes)

ESTADOS = {
    # Aposentado: a minuta já cobre o texto que ainda não é oficial. Fica aqui só para ler
    # e versionar documentos emitidos antes do corte; não é oferecido para documento novo.
    "RASCUNHO":   {"cor": "595959", "texto": "FFFFFF", "marca": "RASCUNHO", "claro": "E4E4E4",
                   "tarja": "RASCUNHO  ·  DOCUMENTO EM ELABORAÇÃO  ·  SEM VALIDADE OFICIAL",
                   "rodape": "Rascunho sem validade oficial. Não distribuir.", "rotulo": "Rascunho"},
    "MINUTA":     {"cor": "C99700", "texto": "000000", "marca": "MINUTA", "claro": "F6EBC4",
                   "tarja": "MINUTA  ·  TEXTO PROPOSTO PARA ANÁLISE  ·  SEM VALIDADE OFICIAL",
                   "rodape": "Minuta para análise: o texto pode mudar e não produz efeitos até a aprovação.", "rotulo": "Minuta"},
    "EM_APROVACAO": {"cor": "1F5FA8", "texto": "FFFFFF", "marca": "EM APROVAÇÃO", "claro": "D9E5F3",
                   "tarja": "EM APROVAÇÃO  ·  AGUARDANDO DECISÃO  ·  SEM VALIDADE OFICIAL",
                   "rodape": "Texto entregue a quem aprova. Sem validade oficial até a aprovação.", "rotulo": "Em aprovação"},
    # Renomeado para EM_APROVACAO: "em revisão" sugeria que o documento oficial estivesse suspenso
    # durante a revisão. Fica aqui só para ler e versionar os arquivos emitidos com o nome antigo.
    "EM_REVISAO": {"cor": "1F5FA8", "texto": "FFFFFF", "marca": "EM REVISÃO", "claro": "D9E5F3",
                   "tarja": "EM REVISÃO  ·  AGUARDANDO APROVAÇÃO  ·  SEM VALIDADE OFICIAL",
                   "rodape": "Versão submetida à aprovação. Sem validade oficial até a aprovação.", "rotulo": "Em revisão"},
    "OFICIAL":    {"cor": "1E7B3C", "texto": "FFFFFF", "marca": None, "claro": None,
                   "tarja": "DOCUMENTO OFICIAL  ·  VIGENTE",
                   "rodape": "Cópia impressa não controlada: confirme a versão vigente.", "rotulo": "Oficial"},
    "OBSOLETO":   {"cor": "B00020", "texto": "FFFFFF", "marca": "OBSOLETO", "claro": "F2D0D5",
                   "tarja": "OBSOLETO  ·  NÃO UTILIZAR",
                   "rodape": "Documento obsoleto, mantido apenas para histórico.", "rotulo": "Obsoleto"},
}
ESTADOS_ATIVOS = ["MINUTA", "EM_APROVACAO", "OFICIAL", "OBSOLETO"]
# Estados aposentados e o que os substitui, ao gerar a rodada seguinte de um arquivo antigo.
ESTADO_SUCESSOR = {"RASCUNHO": "MINUTA", "EM_REVISAO": "EM_APROVACAO"}
TIPOS_CURTOS = {"CI": "Comunicação Interna", "COM": "Comunicado", "OFI": "Ofício", "DEC": "Declaração",
                "TER": "Termo", "ATA": "Ata", "RO": "Registro de Ocorrência", "AUT": "Autorização", "NOT": "Notificação",
                "CTR": "Contrato", "PRO": "Proposta", "FOR": "Formulário", "REC": "Recibo"}
TIPOS_LONGOS = {"MAN": "Manual", "POP": "Procedimento Operacional Padrão", "POL": "Política",
                "REG": "Regulamento", "REL": "Relatório", "PLN": "Plano", "PRJ": "Projeto", "INS": "Instrução de Trabalho"}

# Áreas do padrão. O marca.json pode trocar a lista ("areas"), acrescentar ("areas_extra") e dizer quais são
# oferecidas para documento novo ("areas_ativas"); as demais continuam valendo para ler documento antigo.
AREAS = {"ADM": "Administrativo", "DIR": "Diretoria", "FIN": "Financeiro", "FIS": "Financeiro e fiscal", "JUR": "Jurídico",
         "RH": "Recursos humanos", "VEN": "Vendas e comercial", "ECO": "E-commerce e marketplaces", "EST": "Estoque",
         "EXP": "Expedição", "CMP": "Compras", "OFC": "Oficina e assistência técnica", "MKT": "Marketing", "TI": "Tecnologia"}

NS = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word" '
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'mc:Ignorable="w14 wp14"')

# ----------------------------------------------------------------------------- ativos
def dim_imagem(dados):
    if dados[:8] == b"\x89PNG\r\n\x1a\n":
        return int.from_bytes(dados[16:20], "big"), int.from_bytes(dados[20:24], "big"), "png"
    if dados[:2] == b"\xff\xd8":
        i = 2
        while i < len(dados):
            if dados[i] != 0xFF:
                i += 1; continue
            marc = dados[i + 1]
            if marc in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                return int.from_bytes(dados[i + 7:i + 9], "big"), int.from_bytes(dados[i + 5:i + 7], "big"), "jpeg"
            i += 2 + int.from_bytes(dados[i + 2:i + 4], "big")
    raise ValueError("Imagem não suportada (use PNG ou JPEG)")


def ler_fonte(ref, base):
    """Fonte do visual: caminho (relativo ao marca.json) ou data URI base64 de uma TrueType (.ttf)."""
    if str(ref).startswith("data:"):
        dados = base64.b64decode(str(ref).split(",", 1)[1])
    else:
        caminho = ref if os.path.isabs(ref) else os.path.join(base, ref)
        with open(caminho, "rb") as f:
            dados = f.read()
    if dados[:4] not in (b"\x00\x01\x00\x00", b"true"):
        raise ValueError("não é uma fonte TrueType (.ttf)")
    return dados


def ler_imagem(ref, base):
    """Logo do marca.json: caminho (relativo ao marca.json) ou data URI base64 de PNG/JPEG."""
    if str(ref).startswith("data:"):
        dados = base64.b64decode(str(ref).split(",", 1)[1])
    else:
        caminho = ref if os.path.isabs(ref) else os.path.join(base, ref)
        with open(caminho, "rb") as f:
            dados = f.read()
    dim_imagem(dados)
    return dados


def _preenchido(v):
    """Campo do cadastro: vazio ou [definir] vale como não informado."""
    return None if v is None or str(v).strip() in ("", INDEFINIDO) else str(v).strip()


def carregar_marca(caminho):
    """Lê o marca.json, preenche o cadastro (empresas, áreas, tipos, organização) e devolve os ativos (logos)."""
    try:
        with open(caminho, encoding="utf8") as f:
            cfg = json.load(f)
    except (OSError, ValueError) as e:
        raise SystemExit(f"Não consegui ler o marca.json ({caminho}): {e}")
    base = os.path.dirname(os.path.abspath(caminho))
    emps = cfg.get("empresas") or []
    if not emps:
        raise SystemExit("marca.json: a lista 'empresas' está vazia.")
    EMPRESAS.clear()
    logos = {}
    for e in emps:
        sig = str(e.get("sigla", "")).upper()
        if not re.fullmatch(r"[A-Z0-9]{2,6}", sig):
            raise SystemExit(f"marca.json: sigla inválida {sig!r} (2 a 6 letras ou números).")
        if sig in EMPRESAS:
            raise SystemExit(f"marca.json: sigla {sig} repetida.")
        cor = str(e.get("cor") or "1F3864").lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", cor):
            raise SystemExit(f"marca.json: cor inválida em {sig} (use hexadecimal, ex.: 1F3864).")
        grupo = bool(e.get("grupo"))
        marca = _preenchido(e.get("marca")) or sig
        if e.get("logo"):
            try:
                logos[sig] = ler_imagem(e["logo"], base)
            except (OSError, ValueError) as ex:
                raise SystemExit(f"marca.json: logo de {sig} ilegível ({ex}).")
        EMPRESAS[sig] = {"marca": marca, "razao": _preenchido(e.get("razao_social")) or (marca if grupo else None),
                         "documento": _preenchido(e.get("documento")),
                         "documento_rotulo": _preenchido(e.get("documento_rotulo")) or "CNPJ",
                         "cidade": _preenchido(e.get("cidade")), "cor": cor, "grupo": grupo,
                         "ativa": e.get("ativa", True) is not False, "sem_logo": sig not in logos}
    EMPRESAS_ATIVAS[:] = [k for k, d in EMPRESAS.items() if d["ativa"]]
    EMPRESAS_DO_GRUPO[:] = [k for k in EMPRESAS_ATIVAS if not EMPRESAS[k]["grupo"]]
    if not EMPRESAS_ATIVAS:
        raise SystemExit("marca.json: nenhuma empresa ativa.")

    if cfg.get("areas"):
        AREAS.clear()
        AREAS.update({str(k).upper(): v for k, v in cfg["areas"].items()})
    AREAS.update({str(k).upper(): v for k, v in (cfg.get("areas_extra") or {}).items()})
    for sig, d in (cfg.get("tipos_extra") or {}).items():
        sig = str(sig).upper()
        if not isinstance(d, dict) or d.get("variante") not in ("curto", "longo"):
            raise SystemExit(f"marca.json: o tipo extra {sig} precisa de \"nome\" e \"variante\" ('curto' ou 'longo').")
        (TIPOS_CURTOS if d["variante"] == "curto" else TIPOS_LONGOS)[sig] = d.get("nome") or sig
    tipos = {**TIPOS_CURTOS, **TIPOS_LONGOS}
    for chave, todos in (("areas_ativas", AREAS), ("tipos_ativos", tipos)):
        lista = [str(x).upper() for x in (cfg.get(chave) or todos)]
        fora = [x for x in lista if x not in todos]
        if fora:
            raise SystemExit(f"marca.json: {chave} cita {', '.join(fora)}, que não existe(m).")
        ORG[chave] = lista
    padrao = str(cfg.get("empresa_padrao") or EMPRESAS_ATIVAS[0]).upper()
    if padrao not in EMPRESAS:
        raise SystemExit(f"marca.json: empresa_padrao {padrao} não está em 'empresas'.")
    ORG.update(organizacao=_preenchido(cfg.get("organizacao")) or "",
               politica=_preenchido((cfg.get("politica") or {}).get("codigo")) or "", padrao=padrao)
    definir_empresa(padrao)
    return {"logos": logos, "fontes": carregar_visual(cfg.get("visual"), base)}


def carregar_visual(v, base):
    """Lê o "visual" do marca.json; devolve as fontes a embutir: [(família, estilo, bytes)]."""
    VISUAL.clear()
    if not v:
        return []
    if not isinstance(v, dict):
        raise SystemExit("marca.json: 'visual' precisa ser um objeto.")
    for k in ("cor_texto", "cor_titulo", "cor_secundaria", "cor_realce", "cor_papel", "cor_cartao", "cor_linha"):
        if v.get(k):
            c = str(v[k]).lstrip("#").upper()
            if not re.fullmatch(r"[0-9A-F]{6}", c):
                raise SystemExit(f"marca.json: visual.{k} inválida (use hexadecimal, ex.: 1F3864).")
            VISUAL[k] = c
    for k in ("fonte_texto", "fonte_titulo"):
        if _preenchido(v.get(k)):
            if not re.fullmatch(r"[A-Za-z0-9 \-]{1,40}", str(v[k]).strip()):
                raise SystemExit(f"marca.json: visual.{k} tem caracteres que o Word não aceita num nome de fonte.")
            VISUAL[k] = str(v[k]).strip()
    fontes = []
    for f in v.get("fontes") or []:
        fam, est = str(f.get("familia", "")).strip(), f.get("estilo")
        if fam not in (VISUAL.get("fonte_texto"), VISUAL.get("fonte_titulo")):
            raise SystemExit(f"marca.json: a fonte '{fam}' não é a fonte_texto nem a fonte_titulo do visual.")
        if est not in ESTILOS_FONTE:
            raise SystemExit(f"marca.json: estilo de fonte '{est}' inválido (use {', '.join(ESTILOS_FONTE)}).")
        if any(x[0] == fam and x[1] == est for x in fontes):
            raise SystemExit(f"marca.json: fonte {fam} {est} repetida.")
        try:
            fontes.append((fam, est, ler_fonte(f.get("arquivo", ""), base)))
        except (OSError, ValueError) as ex:
            raise SystemExit(f"marca.json: fonte {fam} {est} ilegível ({ex}).")
    return fontes


def cadastro_completo(d):
    return d["grupo"] or bool(d["razao"] and d["documento"])


def resumo_marca():
    linhas = [f"Marca: {ORG['organizacao'] or '(sem nome de organização)'}"
              + (f"  ·  política {ORG['politica']}" if ORG["politica"] else "")]
    for k, d in EMPRESAS.items():
        logo = "logo ok" if not d["sem_logo"] else "SEM LOGO (rodapé só com o filete)"
        if d["grupo"]:
            st = "documento do grupo (sem documento; rodapé com as logos das empresas ativas)"
        elif cadastro_completo(d):
            st = "pode emitir OFICIAL"
        else:
            st = f"só minuta ou em aprovação (falta razão social/{d['documento_rotulo']})"
        linhas.append(f"  {k}: {d['marca']} · #{d['cor']} · {logo} · {st}" + ("" if d["ativa"] else " · APOSENTADA"))
    if VISUAL:
        linhas.append(f"  visual: texto {VISUAL.get('fonte_texto', 'Arial')} · títulos {VISUAL.get('fonte_titulo', 'Arial')}"
                      + (f" · títulos #{VISUAL['cor_titulo']}" if VISUAL.get("cor_titulo") else ""))
    linhas.append(f"  padrão: {ORG['padrao']} · {len(ORG['areas_ativas'])} de {len(AREAS)} áreas ativas · "
                  f"{len(ORG['tipos_ativos'])} de {len(TIPOS_CURTOS) + len(TIPOS_LONGOS)} tipos ativos")
    return "\n".join(linhas)


def embutir_logos(caminho, saida):
    """Cópia do marca.json com os logos (e as fontes do visual) em data URI: um arquivo só, que também serve para a página."""
    with open(caminho, encoding="utf8") as f:
        cfg = json.load(f)
    base = os.path.dirname(os.path.abspath(caminho))
    for e in cfg.get("empresas", []):
        ref = e.get("logo")
        if ref and not str(ref).startswith("data:"):
            dados = ler_imagem(ref, base)
            e["logo"] = f"data:image/{dim_imagem(dados)[2]};base64," + base64.b64encode(dados).decode()
    for f in (cfg.get("visual") or {}).get("fontes") or []:
        ref = f.get("arquivo")
        if ref and not str(ref).startswith("data:"):
            f["arquivo"] = "data:font/ttf;base64," + base64.b64encode(ler_fonte(ref, base)).decode()
    with open(saida, "w", encoding="utf8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return saida

# ----------------------------------------------------------------------------- utilitários XML
class Ctx:
    def __init__(self):
        self.docpr = 100
        self.zorder = 251660000
        self.midia = []          # (nome, dados)
        self.rels = {}           # parte -> list of (rid, tipo, alvo)
        self.nums = []           # (numId, abstractId)
        self.cont = {"Tabela": 0, "Quadro": 0, "Figura": 0, "Gráfico": 0}
        self.sec = [0, 0, 0, 0, 0]
        self.titulos = []        # (nível, texto, marcador) para o sumário
        self.avisos = []

    def novo_id(self):
        self.docpr += 1; return self.docpr

    def z(self):
        self.zorder += 10; return self.zorder

    def rel(self, parte, tipo, alvo):
        lst = self.rels.setdefault(parte, [])
        for rid, t, a in lst:
            if t == tipo and a == alvo:
                return rid
        rid = f"rId{len(lst) + 1}"
        lst.append((rid, tipo, alvo)); return rid

    def imagem(self, parte, dados):
        _, _, ext = dim_imagem(dados)
        for nome, d in self.midia:
            if d == dados:
                break
        else:
            nome = f"img{len(self.midia) + 1}.{ 'png' if ext == 'png' else 'jpeg'}"
            self.midia.append((nome, dados))
        return self.rel(parte, "image", f"media/{nome}")


def grade(ws):
    return "".join('<w:gridCol w:w="%d"/>' % w for w in ws)


def rpr(b=False, i=False, sz=None, cor=None, caps=False, fonte=None, esp=None, pos=None):
    s = ""
    if fonte: s += f'<w:rFonts w:ascii="{fonte}" w:hAnsi="{fonte}" w:cs="{fonte}"/>'
    if b: s += "<w:b/><w:bCs/>"
    if i: s += "<w:i/><w:iCs/>"
    if caps: s += "<w:caps/>"
    if cor: s += f'<w:color w:val="{cor}"/>'
    if esp: s += f'<w:spacing w:val="{esp}"/>'
    if pos: s += f'<w:position w:val="{int(pos)}"/>'   # meios-pontos acima da linha de base
    if sz: s += f'<w:sz w:val="{int(sz * 2)}"/><w:szCs w:val="{int(sz * 2)}"/>'
    return f"<w:rPr>{s}</w:rPr>" if s else ""


def run(texto, **k):
    partes = str(texto).split("\t")
    out = ""
    for n, t in enumerate(partes):
        if n: out += f"<w:r>{rpr(**k)}<w:tab/></w:r>"
        if t: out += f'<w:r>{rpr(**k)}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>'
    return out


def inline(texto, **k):
    """**negrito**, *itálico*"""
    out = ""
    for tok in re.split(r"(\*\*\*.+?\*\*\*|\*\*.+?\*\*|\*[^*\s][^*]*?\*)", str(texto)):
        if not tok: continue
        if tok.startswith("***") and tok.endswith("***") and len(tok) > 6:
            out += run(tok[3:-3], **{**k, "b": True, "i": True})
        elif tok.startswith("**") and tok.endswith("**"):
            out += run(tok[2:-2], **{**k, "b": True})
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            out += run(tok[1:-1], **{**k, "i": True})
        else:
            out += run(tok, **k)
    return out


def campo(instr, cache, **k):
    return (f'<w:r>{rpr(**k)}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{rpr(**k)}<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r>{rpr(**k)}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr(**k)}<w:t>{cache}</w:t></w:r>'
            f'<w:r>{rpr(**k)}<w:fldChar w:fldCharType="end"/></w:r>')


def par(conteudo="", estilo=None, jc=None, ind=None, antes=None, depois=None, linha=None,
        keep=False, quebra_antes=False, borda=None, tabs=None, rpr_par="", outline=None, sect=""):
    p = ""
    if estilo: p += f'<w:pStyle w:val="{estilo}"/>'
    if keep: p += "<w:keepNext/>"
    if quebra_antes: p += "<w:pageBreakBefore/>"
    if borda: p += f"<w:pBdr>{borda}</w:pBdr>"
    if tabs: p += "<w:tabs>" + "".join(f'<w:tab w:val="{v}" w:leader="{l}" w:pos="{pos}"/>' for v, pos, l in tabs) + "</w:tabs>"
    if antes is not None or depois is not None or linha is not None:
        a = f' w:before="{antes}"' if antes is not None else ""
        d = f' w:after="{depois}"' if depois is not None else ""
        l = f' w:line="{linha}" w:lineRule="auto"' if linha is not None else ""
        p += f"<w:spacing{a}{d}{l}/>"
    if ind: p += f"<w:ind {ind}/>"
    if jc: p += f'<w:jc w:val="{jc}"/>'
    if outline is not None: p += f'<w:outlineLvl w:val="{outline}"/>'
    if rpr_par: p += rpr_par
    return f"<w:p><w:pPr>{p}{sect}</w:pPr>{conteudo}</w:p>"


def anchor(ctx, grafico, x, y, cx, cy, atras=False, relH="page", relV="page", alinhar=None, nome="Forma", descr=""):
    idd = ctx.novo_id()
    posH = f"<wp:align>{alinhar[0]}</wp:align>" if alinhar and alinhar[0] else f"<wp:posOffset>{int(x)}</wp:posOffset>"
    posV = f"<wp:align>{alinhar[1]}</wp:align>" if alinhar and alinhar[1] else f"<wp:posOffset>{int(y)}</wp:posOffset>"
    return (f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{ctx.z()}" '
            f'behindDoc="{1 if atras else 0}" locked="1" layoutInCell="0" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="{relH}">{posH}</wp:positionH>'
            f'<wp:positionV relativeFrom="{relV}">{posV}</wp:positionV><wp:extent cx="{int(cx)}" cy="{int(cy)}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{idd}" name="{escape(nome)} {idd}" descr="{escape(descr)}"/><wp:cNvGraphicFramePr/>'
            f"{grafico}</wp:anchor>")


def g_imagem(rid, cx, cy, nome="img"):
    return (f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{nome}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')


def g_forma(cx, cy, preench=None, linha=None, linha_w=9525, texto_xml="", rot=0):
    fill = f'<a:solidFill><a:srgbClr val="{preench}"/></a:solidFill>' if preench else "<a:noFill/>"
    ln = (f'<a:ln w="{linha_w}"><a:solidFill><a:srgbClr val="{linha}"/></a:solidFill></a:ln>' if linha
          else "<a:ln><a:noFill/></a:ln>")
    txbx = f"<wps:txbx><w:txbxContent>{texto_xml}</w:txbxContent></wps:txbx>" if texto_xml else ""
    return (f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot="{rot}"><a:off x="0" y="0"/><a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}{ln}</wps:spPr>{txbx}'
            f'<wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="0">'
            f'<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>')


def desenho(anc):
    return f'<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>{anc}</w:drawing></mc:Choice></mc:AlternateContent></w:r>'


def desenho_img(anc):
    return f"<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>{anc}</w:drawing></w:r>"

# ----------------------------------------------------------------------------- elementos de página
SHAPETYPE_136 = ('<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">'
    '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
    '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>'
    '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>'
    '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
    '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
    '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>')


def marca_dagua_vml(ctx, texto, cor, w=470.0):
    """Marca d'água nativa do Word (WordArt diagonal atrás do texto) — reconhecida também pelo LibreOffice.
    w: comprimento do texto, em pontos (470 na A4; o telegrama, em A5, usa a mesma proporção da página)."""
    n = len(texto)
    h = min(130.0, w / (n * 0.66))
    idn = ctx.novo_id()
    return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:pict>{SHAPETYPE_136}'
            f'<v:shape id="PowerPlusWaterMarkObject{idn}" o:spid="_x0000_s{2000 + idn}" type="#_x0000_t136" '
            f'style="position:absolute;margin-left:0;margin-top:0;width:{w:.1f}pt;height:{h:.1f}pt;rotation:315;z-index:-251654144;'
            f'mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page" '
            f'o:allowincell="f" fillcolor="#{cor}" stroked="f"><v:fill opacity="1"/>'
            f'<v:textpath style="font-family:&quot;{VISUAL.get("fonte_titulo", "Arial")}&quot;;font-size:1pt;font-weight:bold" string="{escape(texto)}"/>'
            f'<w10:wrap anchorx="page" anchory="page"/></v:shape></w:pict></w:r>')


def texto_tarja(spec):
    est = ESTADOS[spec["estado"]]
    txt = est["tarja"]
    if spec["estado"] == "OFICIAL" and spec.get("vigencia"):
        txt += f"  ·  DESDE {spec['vigencia']}"
    if spec["estado"] == "OBSOLETO" and spec.get("substituido_por"):
        txt = f"OBSOLETO  ·  SUBSTITUÍDO POR {spec['substituido_por']}  ·  NÃO UTILIZAR"
    return txt


def tarja(spec, margem_esq_cm):
    """Faixa do estado: tabela de uma célula no fluxo do cabeçalho, dentro da borda da página
    (área imprimível), de 1,2 cm a 19,8 cm. Não usa forma flutuante: imprime igual em qualquer impressora."""
    est = ESTADOS[spec["estado"]]
    larg = int((21 - 2.4) * TW_CM)
    ind = -int((margem_esq_cm - 1.2) * TW_CM)
    conteudo = par(run(texto_tarja(spec), b=True, sz=8, cor=est["texto"], esp=10), jc="center", antes=0, depois=0, linha=240, estilo="Rodape")
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{larg}" w:type="dxa"/><w:tblInd w:w="{ind}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{larg}"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="{int(0.55 * TW_CM)}" w:hRule="exact"/></w:trPr>'
            f'<w:tc><w:tcPr><w:tcW w:w="{larg}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="{est["cor"]}"/><w:vAlign w:val="center"/></w:tcPr>'
            f'{conteudo}</w:tc></w:tr></w:tbl>')


def marca(ctx, spec, w=470.0):
    est = ESTADOS[spec["estado"]]
    return marca_dagua_vml(ctx, est["marca"], est["claro"], w) if est["marca"] else ""


def logo_inline(ctx, parte, ativos, altura_cm, emp=None):
    emp = emp or EMPRESA_ID
    if emp not in ativos["logos"]:
        raise SystemExit(f"O marca.json não tem o logo da empresa {emp} ({EMPRESAS[emp]['marca']}).")
    dados = ativos["logos"][emp]
    lw, lh, _ = dim_imagem(dados)
    cy = altura_cm * EMU_CM; cx = cy * lw / lh
    rid = ctx.imagem(parte, dados)
    return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{int(cx)}" cy="{int(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{ctx.novo_id()}" name="Logo" descr="TIMBRADOR_LOGO|{emp}"/><wp:cNvGraphicFramePr/>'
            f"{g_imagem(rid, cx, cy, 'logo')}</wp:inline></w:drawing></w:r>")


def cabecalho_controle(ctx, spec, margem_esq_cm):
    """Cabeçalho de todas as páginas: faixa do estado no topo e, logo abaixo, o bloco de controle
    (código, versão e página; situação do documento; empresa e documento)."""
    return tarja(spec, margem_esq_cm) + linhas_controle(ctx, spec)


def logos_do_grupo(ctx, parte, ativos, altura_cm):
    """Rodapé do documento do grupo: as logos de todas as empresas, lado a lado. Empresa ainda sem logo
    aparece pelo nome, em negrito e na cor dela, para que nenhuma fique de fora."""
    sep = run("\u2003\u2003")
    pecas = []
    for k in EMPRESAS_DO_GRUPO:
        e = EMPRESAS[k]
        if e.get("sem_logo") or k not in ativos["logos"]:
            # o nome sobe até o meio da altura das logos, para não ficar rente à base delas
            pecas.append(run(e["marca"], b=True, sz=round(altura_cm * 14), cor=e["cor"], pos=round(altura_cm * 18.5)))
        else:
            pecas.append(logo_inline(ctx, parte, ativos, altura_cm, k))
    return sep.join(pecas)


def rodape_logo(ctx, parte, ativos):
    """Rodapé do documento curto e da capa: filete na cor da empresa e logo centralizado.
    Empresa sem logo cadastrado sai só com o filete; o grupo leva as logos de todas as empresas."""
    filete = f'<w:top w:val="single" w:sz="12" w:space="6" w:color="{EMPRESA["cor"]}"/>'
    if EMPRESA.get("grupo"):
        conteudo = logos_do_grupo(ctx, parte, ativos, 1.0)
    else:
        conteudo = "" if EMPRESA.get("sem_logo") else logo_inline(ctx, parte, ativos, 1.3)
    return par(conteudo, jc="center", antes=0, depois=0, linha=240, estilo="Rodape", borda=filete)


def codigo_versao(spec):
    return f'{spec["codigo"]}  ·  v{spec["versao"]}'


# Documento confidencial: só as pessoas nomeadas nele ou autorizadas pela direção. A palavra CONFIDENCIAL
# vai na linha de controle de todas as páginas, para que uma folha solta também diga o que é, e quem pode ler aparece
# logo abaixo do título (curto) ou na folha de controle (longo).
COR_CONFIDENCIAL = "C00000"


def eh_confidencial(spec):
    return str(spec.get("classificacao") or "").strip().lower() == "confidencial"


def marca_confidencial(sz=8):
    return run("CONFIDENCIAL", sz=sz, b=True, cor=COR_CONFIDENCIAL) + run("  ·  ", sz=sz, cor=sec("404040"))


def pagina_x_de_y(sz=8, cor="404040", b=False):
    return run("Página ", sz=sz, cor=cor, b=b) + campo("PAGE", "1", sz=sz, cor=cor, b=b) + run(" de ", sz=sz, cor=cor, b=b) + campo("NUMPAGES", "1", sz=sz, cor=cor, b=b)


def linhas_controle(ctx, spec):
    est = ESTADOS[spec["estado"]]
    texto = est["rodape"]
    if spec["estado"] == "OFICIAL" and spec.get("aprovado_por"):
        texto = f'Aprovado por {spec["aprovado_por"]} em {spec.get("data_aprovacao", "")}. ' + texto
    linha1 = par(marca(ctx, spec) + run(codigo_versao(spec) + "  ·  ", sz=8, cor=sec("404040"))
                 + (marca_confidencial() if eh_confidencial(spec) else "") + pagina_x_de_y(sz=8, cor=cor_texto(), b=True),
                 jc="center", antes=80, depois=0, linha=240, estilo="Rodape")
    linha2 = par(run(texto, sz=7.5, cor=est["cor"] if spec["estado"] != "OFICIAL" else sec("404040"), b=spec["estado"] != "OFICIAL"),
                 jc="center", antes=0, depois=0, linha=240, estilo="Rodape")
    linha3 = par(run(identificacao_empresa(), sz=7.5, cor=sec("7F7F7F")),
                 jc="center", antes=0, depois=0, linha=240, estilo="Rodape",
                 borda='<w:bottom w:val="single" w:sz="4" w:space="4" w:color="BFBFBF"/>')
    return linha1 + linha2 + linha3


def rodape_reduzido(ctx, parte, ativos, spec, largura_tw):
    """Rodapé das páginas internas do documento longo: filete laranja, logo à esquerda e título à direita."""
    if EMPRESA.get("grupo"):
        logo = logos_do_grupo(ctx, parte, ativos, 0.6)
        c1 = int(largura_tw * 0.6)
    else:
        logo = "" if EMPRESA.get("sem_logo") else logo_inline(ctx, parte, ativos, 0.9)
        c1 = int(largura_tw * 0.35)
    c2 = largura_tw - c1
    borda = f'<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="{EMPRESA["cor"]}"/></w:tblBorders>'
    cel = lambda w, conteudo: f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{conteudo}</w:tc>'
    titulo = spec.get("titulo_curto") or spec["titulo"]
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{largura_tw}" w:type="dxa"/>{borda}<w:tblLayout w:type="fixed"/>'
            f'<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{c1}"/><w:gridCol w:w="{c2}"/></w:tblGrid><w:tr>'
            f'{cel(c1, par(logo, antes=0, depois=0, linha=240, estilo="Rodape"))}'
            f'{cel(c2, par(run(titulo, sz=8, cor=sec("404040")), jc="right", antes=0, depois=0, linha=240, estilo="Rodape"))}</w:tr></w:tbl>'
            + par("", antes=0, depois=0, linha=120, estilo="Rodape"))

# ----------------------------------------------------------------------------- blocos de conteúdo
def titulo_numerado(ctx, nivel, texto, numerar=True, quebra=False):
    prefixo = ""
    if numerar:
        ctx.sec[nivel - 1] += 1
        for k in range(nivel, 5): ctx.sec[k] = 0
        if any(ctx.sec[k] == 0 for k in range(nivel - 1)):
            ctx.avisos.append(f"Título '{texto}' de nível {nivel} sem seção superior.")
        prefixo = ".".join(str(max(1, n)) for n in ctx.sec[:nivel]) + " "
    exib = texto.upper() if nivel <= 2 else texto
    marca = f"_Toc9{len(ctx.titulos) + 1:08d}"
    ctx.titulos.append((nivel, prefixo + exib, marca))
    bid = 9000 + len(ctx.titulos)
    conteudo = (f'<w:bookmarkStart w:id="{bid}" w:name="{marca}"/>' + run(prefixo) + run(exib)
                + f'<w:bookmarkEnd w:id="{bid}"/>')
    return par(conteudo, estilo=f"Heading{nivel}", quebra_antes=quebra)


def titulo_sem_numero(ctx, texto, quebra=True, sumario=True):
    conteudo = run(texto.upper())
    if sumario:
        marca = f"_Toc9{len(ctx.titulos) + 1:08d}"
        ctx.titulos.append((1, texto.upper(), marca))
        bid = 9000 + len(ctx.titulos)
        conteudo = f'<w:bookmarkStart w:id="{bid}" w:name="{marca}"/>' + conteudo + f'<w:bookmarkEnd w:id="{bid}"/>'
    return par(conteudo, estilo="TituloSemNumero" if sumario else "TituloSumario", quebra_antes=quebra)


def nova_lista(ctx, abstrato):
    nid = len(ctx.nums) + 1
    ctx.nums.append((nid, abstrato)); return nid


def bloco_lista(ctx, itens, abstrato, nivel=0, nid=None):
    nid = nid or nova_lista(ctx, abstrato)
    out = ""
    for it in itens:
        if isinstance(it, dict):
            txt, sub = it.get("texto", ""), it.get("sub", [])
        else:
            txt, sub = it, []
        out += (f'<w:p><w:pPr><w:pStyle w:val="Lista"/><w:numPr><w:ilvl w:val="{nivel}"/><w:numId w:val="{nid}"/></w:numPr>'
                f"</w:pPr>{inline(txt)}</w:p>")
        if sub: out += bloco_lista(ctx, sub, abstrato, nivel + 1, nid)
    return out


def bloco_tabela(ctx, b, largura_tw, legenda=True):
    tipo = b.get("tipo", "Tabela")
    out = ""
    if legenda:
        ctx.cont[tipo] = ctx.cont.get(tipo, 0) + 1
        out = par(run(f"{tipo} {ctx.cont[tipo]} – ", b=True, sz=10) + inline(b.get("titulo", ""), sz=10), estilo="Legenda", keep=True)
    cab = b.get("cabecalho") or []
    linhas = b.get("linhas") or []
    ncol = len(cab) if cab else max(len(l) for l in linhas)
    pesos = b.get("larguras") or [1] * ncol
    ws = [int(largura_tw * p / sum(pesos)) for p in pesos]
    ws[-1] = largura_tw - sum(ws[:-1])
    q = tipo == "Quadro"
    lin = '<w:{0} w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
    bordas = "".join(lin.format(s) for s in (["top", "left", "bottom", "right", "insideH", "insideV"] if q else ["top", "bottom"]))
    alinh = b.get("alinhamento") or ["left"] * ncol
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="{largura_tw}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>{bordas}</w:tblBorders><w:tblLayout w:type="fixed"/>'
           f'<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
           f'<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid>{grade(ws)}</w:tblGrid>')

    def linha_xml(valores, cabecalho=False, ultima=False):
        tr = '<w:tr><w:trPr><w:cantSplit/>' + ("<w:tblHeader/>" if cabecalho else "") + "</w:trPr>"
        for k in range(ncol):
            v = valores[k] if k < len(valores) else ""
            tcpr = f'<w:tcW w:w="{ws[k]}" w:type="dxa"/>'
            if cabecalho:
                tcpr += '<w:tcBorders><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tcBorders>'
                tcpr += f'<w:shd w:val="clear" w:color="auto" w:fill="{realce()}"/>'
            tr += (f"<w:tc><w:tcPr>{tcpr}<w:vAlign w:val=\"center\"/></w:tcPr>" +
                   par(inline(v, b=cabecalho), estilo="TabelaTexto", jc="center" if cabecalho else alinh[k]) + "</w:tc>")
        return tr + "</w:tr>"

    if cab: tbl += linha_xml(cab, True)
    for n, l in enumerate(linhas): tbl += linha_xml(l, ultima=n == len(linhas) - 1)
    tbl += "</w:tbl>"
    out += tbl
    if not legenda:
        return out
    out += par(run("Fonte: ", sz=10) + inline(b.get("fonte", "Elaborado pela empresa (%d)." % datetime.date.today().year), sz=10), estilo="Fonte")
    if b.get("nota"):
        out += par(run("Nota: ", sz=10) + inline(b["nota"], sz=10), estilo="Fonte")
    return out


def bloco_figura(ctx, b, parte, largura_tw):
    tipo = b.get("tipo", "Figura")
    ctx.cont[tipo] = ctx.cont.get(tipo, 0) + 1
    out = par(run(f"{tipo} {ctx.cont[tipo]} – ", b=True, sz=10) + inline(b.get("titulo", ""), sz=10), estilo="Legenda", keep=True)
    dados = open(b["caminho"], "rb").read()
    w, h, _ = dim_imagem(dados)
    larg = min(b.get("largura_cm", 14), largura_tw / TW_CM) * EMU_CM
    alt = larg * h / w
    rid = ctx.imagem(parte, dados)
    img = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{int(larg)}" cy="{int(alt)}"/>'
           f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{ctx.novo_id()}" name="Figura" descr="{escape(b.get("titulo", ""))}"/>'
           f'<wp:cNvGraphicFramePr/>{g_imagem(rid, larg, alt)}</wp:inline></w:drawing></w:r>')
    out += par(img, jc="center", antes=0, depois=0, linha=240, keep=True, ind='w:firstLine="0"')
    out += par(run("Fonte: ", sz=10) + inline(b.get("fonte", "Elaborado pela empresa."), sz=10), estilo="Fonte")
    return out


def bloco_campos(itens, largura_tw, rotulo_cm=4.0):
    c1 = int(rotulo_cm * TW_CM); c2 = largura_tw - c1
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="{largura_tw}" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
           f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/></w:tblBorders><w:tblLayout w:type="fixed"/>'
           f'<w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="{c1}"/><w:gridCol w:w="{c2}"/></w:tblGrid>')
    for rot, val in itens:
        tbl += (f'<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{c1}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="{realce()}"/></w:tcPr>'
                f'{par(run(rot, b=True, sz=10), estilo="TabelaTexto")}</w:tc>'
                f'<w:tc><w:tcPr><w:tcW w:w="{c2}" w:type="dxa"/></w:tcPr>{par(inline(val, sz=10), estilo="TabelaTexto")}</w:tc></w:tr>')
    return tbl + "</w:tbl>" + par("", antes=0, depois=0, linha=240, estilo="TabelaTexto")


def bloco_assinaturas(ass, largura_tw):
    if not ass: return ""
    por_linha = 2 if len(ass) > 1 else 1
    out = par("", antes=0, depois=0, keep=True)
    for i in range(0, len(ass), por_linha):
        grupo = ass[i:i + por_linha]
        total = largura_tw if por_linha > 1 else min(largura_tw, int(9 * TW_CM))
        w = total // por_linha
        tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/>'
               f'<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr>'
               f'<w:tblGrid>{grade([w] * len(grupo))}</w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>')
        for a in grupo:
            conteudo = par("", antes=900, depois=0, linha=240, estilo="TabelaTexto", keep=True)
            conteudo += par(run(a.get("nome", "").upper(), b=True, sz=10), jc="center", estilo="TabelaTexto", keep=True,
                            borda='<w:top w:val="single" w:sz="6" w:space="4" w:color="000000"/>')
            for extra in [a.get("cargo"), a.get("papel"), a.get("documento")]:
                if extra: conteudo += par(run(extra, sz=10), jc="center", estilo="TabelaTexto", keep=True)
            tbl += f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{conteudo}</w:tc>'
        out += tbl + "</w:tr></w:tbl>"
    return out + par("", antes=0, depois=0, linha=240)


CAB_ASSINATURAS = ["Nome", "Cargo", "Papel", "Assinatura", "Data"]


def bloco_assinaturas_tabela(ass, largura_tw):
    """Assinaturas em tabela: usadas quando as linhas de assinatura ficariam sozinhas numa página."""
    if not ass: return ""
    ws = [int(largura_tw * x) for x in (0.3, 0.18, 0.14, 0.26)]
    ws.append(largura_tw - sum(ws))
    borda = "<w:tblBorders>" + "".join(f'<w:{l} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                       for l in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"
    def cel(w, c, cab=False):
        sombra = f'<w:shd w:val="clear" w:color="auto" w:fill="{realce()}"/>' if cab else ""
        return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{sombra}<w:vAlign w:val="center"/></w:tcPr>{c}</w:tc>'
    t = (f'<w:tbl><w:tblPr><w:tblW w:w="{largura_tw}" w:type="dxa"/>{borda}<w:tblLayout w:type="fixed"/>'
         f'<w:tblCellMar><w:left w:w="85" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grade(ws)}</w:tblGrid>'
         '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' +
         "".join(cel(ws[i], par(run(c, b=True, sz=10), jc="center", estilo="TabelaTexto"), True) for i, c in enumerate(CAB_ASSINATURAS)) + "</w:tr>")
    for a in ass:
        v = [str(a.get("nome", "")).upper(), a.get("cargo", ""), a.get("papel") or a.get("documento", ""), "", ""]
        t += (f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="{int(1.1 * TW_CM)}" w:hRule="atLeast"/></w:trPr>' +
              "".join(cel(ws[i], par(run(x, sz=10, b=(i == 0)), estilo="TabelaTexto")) for i, x in enumerate(v)) + "</w:tr>")
    return par(run("Assinaturas", b=True), estilo="SemRecuo", keep=True, antes=240) + t + "</w:tbl>" + par("", antes=0, depois=0, linha=240)


def _eh_tabela_assinaturas(b):
    return b.get("t") == "tabela" and [str(c).replace("*", "").strip() for c in (b.get("cabecalho") or [])] == CAB_ASSINATURAS


def montar_assinaturas(spec, formato=None):
    """Documento "com assinatura": o corpo termina com local e data e as assinaturas dos envolvidos. Quem decide é a
    caixa do Timbrador (com_assinatura), não o texto nem o tipo do documento. Sem a caixa, o corpo fica como veio."""
    if not spec.get("com_assinatura") or not spec.get("envolvidos"):
        return
    corpo = [b for b in spec.get("corpo") or [] if not (isinstance(b, dict) and (b.get("t") == "assinaturas" or _eh_tabela_assinaturas(b)))]
    while corpo and isinstance(corpo[-1], dict) and (corpo[-1].get("t") == "local_data" or
                                                     (corpo[-1].get("t") == "p" and re.fullmatch(r"\**Assinaturas\**", corpo[-1].get("texto", "")))):
        corpo.pop()
    if formato is None:
        formato = "tabela" if spec.get("assinatura_formato") == "tabela" else "linhas"
    corpo += [{"t": "local_data", "data": spec.get("data_emissao")},
              {"t": "assinaturas", "itens": spec["envolvidos"], "formato": formato}]
    spec["corpo"] = corpo


def comando_soffice():
    """LibreOffice: o comando de TIMBRADOR_SOFFICE (para sandbox que precisa de wrapper), soffice ou libreoffice."""
    cmd = os.environ.get("TIMBRADOR_SOFFICE")
    if cmd:
        return shlex.split(cmd)
    for nome in ("soffice", "libreoffice"):
        if shutil.which(nome):
            return [nome]
    return None


def para_pdf(docx, pasta):
    """Converte com LibreOffice num perfil temporário (não disputa o perfil de quem está usando o programa)."""
    cmd = comando_soffice()
    if not cmd:
        return None
    with tempfile.TemporaryDirectory(prefix="timbrador_lo_") as perfil:
        try:
            subprocess.run(cmd + [f"-env:UserInstallation=file://{perfil}", "--headless", "--convert-to", "pdf",
                                  "--outdir", pasta, docx], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240)
        except (OSError, subprocess.TimeoutExpired):
            return None
    pdf = os.path.join(pasta, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


# Regra do padrão: assinatura entra só pelo Timbrador (a página), na caixa "Com assinatura". Este gerador não põe
# assinatura em documento nenhum: tira do corpo as linhas ou a tabela de assinaturas (e o local e data que as antecede)
# e guarda os envolvidos como "assinaturas_previstas" na ficha. Ao abrir o arquivo na página, eles já vêm preenchidos,
# com a caixa desmarcada: quem timbra decide. A cópia OBSOLETO de uma versão já assinada fica como estava.
def adiar_assinaturas(spec):
    corpo = spec.get("corpo") or []
    previstas = [dict(a) for a in spec.get("assinaturas_previstas") or []] + [dict(a) for a in spec.get("envolvidos") or []]
    fora, tinha = [], False
    for b in corpo:
        if isinstance(b, dict) and b.get("t") == "assinaturas":
            tinha = True
            previstas += [{"nome": a.get("nome", ""), "cargo": a.get("cargo", ""), "papel": a.get("papel") or a.get("documento", "")} for a in b.get("itens", [])]
        elif isinstance(b, dict) and _eh_tabela_assinaturas(b):
            tinha = True
            previstas += [{"nome": (l + ["", "", ""])[0], "cargo": (l + ["", "", ""])[1], "papel": (l + ["", "", ""])[2]} for l in b.get("linhas", [])]
        else:
            fora.append(b)
    if tinha:
        while fora and isinstance(fora[-1], dict) and (fora[-1].get("t") == "local_data" or
                                                       (fora[-1].get("t") == "p" and re.fullmatch(r"\**Assinaturas\**", fora[-1].get("texto", "")))):
            fora.pop()
    pedia = bool(spec.get("com_assinatura"))
    for k in ("com_assinatura", "envolvidos", "assinatura_formato"):
        spec.pop(k, None)
    spec["corpo"] = fora
    vistos, unicas = set(), []
    for a in previstas:
        chave = str(a.get("nome", "")).strip().lower()
        if chave and chave not in vistos:
            vistos.add(chave); unicas.append({k: a.get(k, "") for k in ("nome", "cargo", "papel")})
    if unicas: spec["assinaturas_previstas"] = unicas
    else: spec.pop("assinaturas_previstas", None)
    if tinha or pedia:
        return ("Assinaturas não entram pelo gerador: elas são colocadas no Timbrador (página), na caixa \"Com assinatura\". "
                + (f"Os envolvidos ({', '.join(a['nome'] for a in unicas)}) ficaram registrados como previstos e já vêm preenchidos ao abrir o arquivo na página."
                   if unicas else ""))
    return None


# Anexos externos: arquivos que acompanham o documento (conversas exportadas em XML, planilhas, fotos…). Ficam numa
# pasta com o código do documento, ao lado do .docx, com o nome "Anexo 01 – nome original"; o documento os lista numa
# seção "Anexos", antes do local e data, com tamanho e SHA-256. Mesma regra da página.
CAB_ANEXOS = ["Nº", "Arquivo", "Tamanho"]
TEXTO_ANEXOS = ("Os arquivos abaixo acompanham este documento e ficam na mesma pasta. A impressão digital SHA-256 "
                "permite conferir que o arquivo não foi alterado.")


def tamanho_txt(n):
    n = int(n or 0)
    if n < 1024: return f"{n} B"
    return f"{n / 1024:.1f} KB".replace(".", ",") if n < 1048576 else f"{n / 1048576:.1f} MB".replace(".", ",")


def nome_anexo(i, original, versao=None):
    return f"Anexo {i:02d}{f' (v{versao})' if versao else ''} – {original}"


def _parte_nome(s):
    """Pedaço de nome de arquivo ou pasta vindo da especificação ou da ficha: sem separador de pasta nem caractere
    que o Windows recusa, e nunca só pontos (o código ou a versão de um .docx de fora não pode virar caminho)."""
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "-", str(s))
    return "-" if s and not s.strip(". ") else s


def nome_pasta_doc(spec):
    """Pasta do documento com anexos: o código; minuta (XXX) leva também o título."""
    cod = _parte_nome(spec.get("codigo") or "")
    if "XXX" not in cod: return cod
    t = re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|()]+', " ", str(spec.get("titulo") or ""))).strip()[:60].strip()
    return f"{cod} – {t}" if t else cod


def bloco_anexos(itens, largura_tw):
    if not itens: return ""
    w1, w3 = int(1.2 * TW_CM), int(2.3 * TW_CM)
    ws = [w1, largura_tw - w1 - w3, w3]
    borda = "<w:tblBorders>" + "".join(f'<w:{l} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                       for l in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"
    def cel(w, c, cab=False):
        sombra = f'<w:shd w:val="clear" w:color="auto" w:fill="{realce()}"/>' if cab else ""
        return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{sombra}<w:vAlign w:val="center"/></w:tcPr>{c}</w:tc>'
    t = (f'<w:tbl><w:tblPr><w:tblW w:w="{largura_tw}" w:type="dxa"/>{borda}<w:tblLayout w:type="fixed"/>'
         f'<w:tblCellMar><w:left w:w="85" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grade(ws)}</w:tblGrid>'
         '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' +
         "".join(cel(ws[i], par(run(c, b=True, sz=10), jc="center", estilo="TabelaTexto"), True) for i, c in enumerate(CAB_ANEXOS)) + "</w:tr>")
    for i, a in enumerate(itens):
        arq = (par(run(a.get("nome") or a.get("original", ""), sz=10), estilo="TabelaTexto", antes=0, depois=0) +
               par(run(f'SHA-256: {a.get("sha256", "")}', sz=7, fonte="Courier New"), estilo="TabelaTexto", antes=0, depois=0))
        t += ('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cel(ws[0], par(run(str(i + 1), sz=10), jc="center", estilo="TabelaTexto")) +
              cel(ws[1], arq) + cel(ws[2], par(run(tamanho_txt(a.get("tamanho")), sz=10), jc="right", estilo="TabelaTexto")) + "</w:tr>")
    return (par(run("Anexos", b=True), estilo="SemRecuo", keep=True, antes=240) + par(run(TEXTO_ANEXOS, sz=10), estilo="SemRecuo", keep=True)
            + t + "</w:tbl>" + par("", antes=0, depois=0, linha=240))


def _eh_tabela_anexos(b):
    return isinstance(b, dict) and b.get("t") == "tabela" and [str(c).replace("*", "").strip() for c in (b.get("cabecalho") or [])] == CAB_ANEXOS


def extrair_anexos(corpo):
    """Tira do corpo a seção Anexos (título, texto e tabela): a lista vem da ficha."""
    fora = []
    for b in corpo or []:
        if isinstance(b, dict) and (b.get("t") == "anexos_externos" or _eh_tabela_anexos(b)):
            while fora and isinstance(fora[-1], dict) and fora[-1].get("t") == "p" and \
                    str(fora[-1].get("texto", "")).replace("*", "").strip() in ("Anexos", TEXTO_ANEXOS):
                fora.pop()
            continue
        fora.append(b)
    return fora


def montar_anexos(spec):
    """Seção Anexos no fim do conteúdo, antes do local e data e das assinaturas."""
    corpo = extrair_anexos(spec.get("corpo"))
    itens = [a for a in spec.get("anexos_externos") or [] if isinstance(a, dict)]
    if spec.get("com_anexos") and itens:
        k = len(corpo)
        while k and isinstance(corpo[k - 1], dict) and (corpo[k - 1].get("t") in ("assinaturas", "local_data") or _eh_tabela_assinaturas(corpo[k - 1])
                                                        or (corpo[k - 1].get("t") == "p" and re.fullmatch(r"\**Assinaturas\**", corpo[k - 1].get("texto", "")))):
            k -= 1
        corpo = corpo[:k] + [{"t": "anexos_externos", "itens": itens}] + corpo[k:]
    spec["corpo"] = corpo


def sha256_arquivo(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def _nome_simples(nome):
    """Nome de arquivo sem pasta: nada de / ou \\, nem . ou .. sozinhos (o nome vem da ficha, que é dado de fora)."""
    nome = str(nome)
    return bool(nome.strip()) and nome not in (".", "..") and not re.search(r"[\\/\x00]", nome) and not re.match(r"^[A-Za-z]:", nome)


def _confere(caminho, a):
    return os.path.isfile(caminho) and os.path.getsize(caminho) == int(a.get("tamanho") or -1) and sha256_arquivo(caminho) == a.get("sha256")


def preparar_anexos(spec, arquivos, pasta_doc):
    """Grava os anexos na pasta do documento e põe a lista na especificação. Os já registrados na ficha precisam
    estar na pasta (ou na pasta do arquivo lido pelo --atualizar) com o mesmo SHA-256. Nome ocupado por outro conteúdo
    não é sobrescrito: o anexo novo leva a versão no nome, e a versão anterior continua íntegra."""
    os.makedirs(pasta_doc, exist_ok=True)
    origem = spec.pop("anexos_origem", None)
    versao = str(spec.get("versao", "")).split("-r")[0]
    itens, avisos, usados = [], [], set()
    for a in [x for x in spec.get("anexos_externos") or [] if isinstance(x, dict)]:
        nome = a.get("nome") or ""
        if not _nome_simples(nome):   # vem da ficha do documento lido: só nome de arquivo, sem caminho
            raise SystemExit(f'Anexo com nome inválido na ficha do documento: "{nome}". O nome do anexo não pode ter pasta nem "..".')
        if not _confere(os.path.join(pasta_doc, nome), a):
            if origem and _confere(os.path.join(origem, nome), a):
                shutil.copy2(os.path.join(origem, nome), os.path.join(pasta_doc, nome))
            else:
                raise SystemExit(f'Anexo "{nome}", registrado no documento, não está em {pasta_doc}'
                                 + (f" nem em {origem}" if origem else "") + " com o mesmo conteúdo (SHA-256): passe o arquivo com --anexo, ou retire-o da lista \"anexos_externos\".")
        itens.append({k: a.get(k) for k in ("nome", "original", "tamanho", "sha256")}); usados.add(nome)
    for caminho in arquivos:
        if not os.path.isfile(caminho):
            raise SystemExit(f"Anexo não encontrado: {caminho}")
        original, tamanho, sha = os.path.basename(caminho), os.path.getsize(caminho), sha256_arquivo(caminho)
        if any(x["sha256"] == sha for x in itens):
            avisos.append(f"{original}: arquivo repetido, entrou uma vez só."); continue
        i, k = len(itens) + 1, 1
        nome = nome_anexo(i, original)
        while True:
            destino = os.path.join(pasta_doc, nome)
            if not os.path.exists(destino) and nome not in usados:
                shutil.copy2(caminho, destino); break
            if os.path.exists(destino) and nome not in usados and sha256_arquivo(destino) == sha:
                break
            k += 1
            nome = nome_anexo(i, original, versao + (f"-{k - 1}" if k > 2 else ""))
        usados.add(nome)
        itens.append({"nome": nome, "original": original, "tamanho": tamanho, "sha256": sha})
    if not itens:
        raise SystemExit("Com anexos externos: informe ao menos um arquivo (--anexo).")
    spec["anexos_externos"], spec["com_anexos"] = itens, True
    avisos.append(f"Documento e {len(itens)} anexo(s) na pasta {pasta_doc}.")
    return avisos


def assinaturas_sozinhas(docx, spec):
    """Renderiza com LibreOffice e diz se a página do local e data e das assinaturas não tem mais nada do corpo
    (cabeçalho e rodapé, que se repetem em todas as páginas, não contam). None quando não dá para saber."""
    if not comando_soffice() or not shutil.which("pdftotext"):
        return None
    envolvidos = spec.get("envolvidos") or []
    proprias = {"assinaturas"} | {c.lower() for c in CAB_ASSINATURAS}
    for a in envolvidos:
        for k in ("nome", "cargo", "papel", "documento"):
            if a.get(k): proprias.add(str(a[k]).strip().lower())
    local = f'{dado("cidade")}, {data_extenso(spec.get("data_emissao") or "")}.'.lower()
    tmp = tempfile.mkdtemp()
    try:
        pdf = para_pdf(docx, tmp)
        if not pdf: return None
        bruto = subprocess.run(["pdftotext", pdf, "-"], capture_output=True, text=True).stdout.split("\f")
        paginas = [[re.sub(r"P[áa]gina \d+ de \d+", "", l).strip().lower() for l in pg.splitlines() if l.strip()] for pg in bruto]
        paginas = [pg for pg in paginas if pg]
        if len(paginas) < 2: return False
        k = next((n for n in range(len(paginas) - 1, -1, -1) if local in paginas[n]), None)
        if k is None or k == 0: return False if k == 0 else None
        # cabeçalho e rodapé: linhas do alto e do pé que se repetem em todas as páginas
        zona = lambda pg: set(pg[:6] + pg[-3:])
        comuns = zona(paginas[0]).intersection(*map(zona, paginas[1:]))
        resto = [l for l in paginas[k] if l not in comuns and l != local and l not in proprias and re.search(r"\w", l)]
        return not resto
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def data_extenso(d):
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    try:
        dt = datetime.datetime.strptime(d, "%d/%m/%Y")
        return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"
    except Exception:
        return d


def blocos(ctx, lista, parte, largura_tw, longo):
    out = ""
    primeiro_h1 = True
    for b in lista:
        t = b.get("t")
        if t in ("h1", "h2", "h3", "h4", "h5"):
            nivel = int(t[1])
            quebra = longo and nivel == 1 and (primeiro_h1 is False) and b.get("nova_pagina", True)
            if nivel == 1: primeiro_h1 = False
            out += titulo_numerado(ctx, nivel, b["texto"], b.get("numerar", True), quebra)
        elif t == "p":
            out += par(inline(b["texto"]), estilo="Normal" if not b.get("sem_recuo") else "SemRecuo", jc=b.get("alinhar"))
        elif t == "alineas":
            out += bloco_lista(ctx, b["itens"], 1)
        elif t == "numerada":
            out += bloco_lista(ctx, b["itens"], 2)
        elif t == "marcadores":
            out += bloco_lista(ctx, b["itens"], 3)
        elif t == "checklist":
            out += bloco_lista(ctx, b["itens"], 4)
        elif t == "citacao":
            txt = b["texto"] + (f' {b["fonte"]}' if b.get("fonte") else "")
            out += par(inline(txt, sz=10), estilo="CitacaoLonga")
        elif t == "nota":
            out += par(run((b.get("rotulo") or "Nota") + ": ", b=True, sz=10) + inline(b["texto"], sz=10), estilo="Nota")
        elif t == "destaque":
            cor = {"atencao": "C99700", "info": "1F5FA8", "erro": "B00020"}.get(b.get("nivel", "info"), "1F5FA8")
            out += par(run((b.get("rotulo") or "Importante") + ": ", b=True) + inline(b["texto"]), estilo="SemRecuo",
                       borda=f'<w:left w:val="single" w:sz="24" w:space="8" w:color="{cor}"/>', ind='w:left="227" w:firstLine="0"')
        elif t == "tabela":
            out += bloco_tabela(ctx, b, largura_tw)
        elif t == "figura":
            out += bloco_figura(ctx, b, parte, largura_tw)
        elif t == "campos":
            out += bloco_campos(b["itens"], largura_tw, b.get("rotulo_cm", 4.0))
        elif t == "assinaturas":
            out += bloco_assinaturas_tabela(b["itens"], largura_tw) if b.get("formato") == "tabela" else bloco_assinaturas(b["itens"], largura_tw)
        elif t == "anexos_externos":
            out += bloco_anexos(b["itens"], largura_tw)
        elif t == "local_data":
            dt = b.get("data") or datetime.date.today().strftime("%d/%m/%Y")   # sem data informada, usa a de hoje
            out += par(run(f'{b.get("local", dado("cidade"))}, {data_extenso(dt)}.'), estilo="SemRecuo", jc="right", antes=240, keep=True)
        elif t == "quebra":
            out += par('<w:r><w:br w:type="page"/></w:r>', antes=0, depois=0)
        elif t == "codigo":
            for linha in b["texto"].split("\n"):
                out += par(run(linha, fonte="Courier New", sz=9.5), estilo="TabelaTexto", ind='w:left="284"', antes=0, depois=0)
        else:
            ctx.avisos.append(f"Bloco desconhecido ignorado: {t}")
    return out

# ----------------------------------------------------------------------------- estilos, numeração, settings
def estilos_xml():
    def est(id_, nome, ppr="", rpr_="", base="Normal", tipo="paragraph", prox=None, extra=""):
        b = f'<w:basedOn w:val="{base}"/>' if base else ""
        n = f'<w:next w:val="{prox}"/>' if prox else ""
        return (f'<w:style w:type="{tipo}" w:styleId="{id_}"><w:name w:val="{nome}"/>{b}{n}{extra}<w:qFormat/>'
                f"<w:pPr>{ppr}</w:pPr><w:rPr>{rpr_}</w:rPr></w:style>")
    fonte = lambda n: f'<w:rFonts w:ascii="{n}" w:hAnsi="{n}" w:eastAsia="{n}" w:cs="{n}"/>'
    f = fonte(VISUAL.get("fonte_texto", "Arial"))
    # títulos: fonte e cor próprias só quando o visual as define (sem visual, nada muda)
    tf = fonte(VISUAL["fonte_titulo"]) if VISUAL.get("fonte_titulo") else ""
    tc = f'<w:color w:val="{VISUAL["cor_titulo"]}"/>' if VISUAL.get("cor_titulo") else ""
    s = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
         f'<w:docDefaults><w:rPrDefault><w:rPr>{f}<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="pt-BR" w:eastAsia="pt-BR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
         f'<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>')
    s += ('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:widowControl/><w:spacing w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr>'
          f'<w:rPr><w:color w:val="{cor_texto()}"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>')
    s += est("SemRecuo", "Texto sem recuo", '<w:ind w:firstLine="0"/>')
    tit = '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>'
    s += est("Heading1", "heading 1", tit + '<w:outlineLvl w:val="0"/>', tf + "<w:b/><w:bCs/><w:caps/>" + tc, prox="Normal")
    s += est("Heading2", "heading 2", tit + '<w:outlineLvl w:val="1"/>', tf + "<w:caps/>" + tc, prox="Normal")
    s += est("Heading3", "heading 3", tit + '<w:outlineLvl w:val="2"/>', tf + "<w:b/><w:bCs/>" + tc, prox="Normal")
    s += est("Heading4", "heading 4", tit + '<w:outlineLvl w:val="3"/>', tf + "<w:b/><w:i/>" + tc, prox="Normal")
    s += est("Heading5", "heading 5", tit + '<w:outlineLvl w:val="4"/>', tf + "<w:i/>" + tc, prox="Normal")
    centro = '<w:keepNext/><w:spacing w:before="0" w:after="360" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>'
    s += est("TituloSemNumero", "Título sem indicativo numérico", centro + '<w:outlineLvl w:val="0"/>', tf + "<w:b/><w:bCs/><w:caps/>" + tc, prox="Normal")
    s += est("TituloSumario", "Título do sumário", centro, tf + "<w:b/><w:bCs/><w:caps/>" + tc, prox="Normal")
    s += est("TituloDocumento", "Título do documento", '<w:keepNext/><w:spacing w:before="0" w:after="60" w:line="276" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>',
             tf + '<w:b/><w:bCs/><w:caps/>' + tc + '<w:sz w:val="26"/><w:szCs w:val="26"/>')
    s += est("TituloTelegrama", "Título do telegrama", '<w:keepNext/><w:spacing w:before="0" w:after="240" w:line="276" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/>',
             "<w:b/><w:bCs/><w:caps/>")
    peq = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>'
    s += est("Legenda", "caption", '<w:keepNext/><w:spacing w:before="240" w:after="60" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    s += est("Fonte", "Fonte de ilustração", '<w:spacing w:before="60" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    s += est("TabelaTexto", "Texto de tabela", peq + '<w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    s += est("CitacaoLonga", "Citação longa", '<w:spacing w:before="240" w:after="360" w:line="240" w:lineRule="auto"/><w:ind w:left="2268" w:firstLine="0"/><w:jc w:val="both"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    s += est("Nota", "Nota", '<w:spacing w:before="60" w:after="120" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="both"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    s += est("Referencia", "Referência", '<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/>', "")
    s += est("Lista", "List Paragraph", '<w:spacing w:before="0" w:after="60" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="both"/>', "")
    s += est("Rodape", "footer", peq, '<w:sz w:val="16"/><w:szCs w:val="16"/>')
    toc = lambda n, ind, bold: est(f"TOC{n}", f"toc {n}", f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9061"/></w:tabs><w:spacing w:before="0" w:after="60" w:line="276" w:lineRule="auto"/><w:ind w:left="{ind}" w:firstLine="0"/><w:jc w:val="left"/>', bold, prox="Normal")
    s += toc(1, 0, "<w:b/><w:caps/>") + toc(2, 0, "<w:caps/>") + toc(3, 0, "")
    s += '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    return s + "</w:styles>"


def numeracao_xml(ctx):
    def nivel(ilvl, fmt, texto, esq, pend, fonte=None):
        rp = f'<w:rPr><w:rFonts w:ascii="{fonte}" w:hAnsi="{fonte}" w:hint="default"/></w:rPr>' if fonte else ""
        return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{texto}"/>'
                f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{esq}" w:hanging="{pend}"/></w:pPr>{rp}</w:lvl>')
    a = [
        (1, nivel(0, "lowerLetter", "%1)", 1134, 425) + nivel(1, "bullet", "–", 1560, 426) + nivel(2, "bullet", "–", 1985, 425)),
        (2, nivel(0, "decimal", "%1.", 1134, 425) + nivel(1, "lowerLetter", "%2)", 1560, 426) + nivel(2, "bullet", "–", 1985, 425)),
        (3, nivel(0, "bullet", "•", 1134, 425, VISUAL.get("fonte_texto", "Arial")) + nivel(1, "bullet", "–", 1560, 426) + nivel(2, "bullet", "–", 1985, 425)),
        (4, nivel(0, "bullet", "☐", 1134, 425, "Segoe UI Symbol") + nivel(1, "bullet", "☐", 1560, 426, "Segoe UI Symbol") + nivel(2, "bullet", "–", 1985, 425)),
    ]
    s = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
    for aid, lv in a:
        s += f'<w:abstractNum w:abstractNumId="{aid}"><w:multiLevelType w:val="multilevel"/>{lv}</w:abstractNum>'
    for nid, aid in ctx.nums:
        s += f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{aid}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
    return s + "</w:numbering>"


def settings_xml(atualizar, embutir=False):
    u = '<w:updateFields w:val="true"/>' if atualizar else ""
    e = "<w:embedTrueTypeFonts/>" if embutir else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>'
            f'{e}<w:defaultTabStop w:val="709"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>{u}'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<w:themeFontLang w:val="pt-BR"/><w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/></w:settings>')

# ----------------------------------------------------------------------------- validação da especificação
def validar(spec):
    av = []
    definir_empresa(str(spec.get("empresa") or ORG["padrao"]).upper())
    spec["empresa"] = EMPRESA_ID
    spec["estado"] = spec.get("estado", "MINUTA").upper().replace(" ", "_").replace("Ã", "A")
    if spec["estado"] not in ESTADOS: raise SystemExit(f"Estado inválido: {spec['estado']}")
    if spec["estado"] not in ESTADOS_ATIVOS:
        av.append(f'Estado {ESTADOS[spec["estado"]]["rotulo"].lower()} foi aposentado: use minuta para o texto que ainda não é oficial.')
    tipo = spec.get("tipo", "").upper()
    # Empresa, tipo e área fora da lista de ativos do marca.json: valem para ler e versionar documento antigo.
    if not EMPRESA.get("ativa", True):
        av.append(f"Empresa {EMPRESA_ID} está aposentada no marca.json: use-a só para versionar documento antigo.")
    if tipo and ORG["tipos_ativos"] and tipo not in ORG["tipos_ativos"]:
        av.append(f"Tipo {tipo} não está entre os ativos do marca.json: sigla aposentada ou fora do padrão.")
    if spec.get("area") and ORG["areas_ativas"] and str(spec["area"]).upper() not in ORG["areas_ativas"]:
        av.append(f'Área {str(spec["area"]).upper()} não está entre as ativas do marca.json.')
    if spec["estado"] != "OBSOLETO" and spec.get("titulo"):
        novo = titulo_sem_tipo(spec["titulo"], tipo)
        if novo != spec["titulo"]:
            av.append(f'Título sem o nome do tipo, que já sai impresso acima dele: "{spec["titulo"]}" passou a "{novo}".')
            spec["titulo"] = novo
    if tipo == "PLN" and spec["estado"] != "OBSOLETO" and not spec.get("projeto"):
        raise SystemExit("Plano é sempre de um projeto: informe 'projeto' (o documento vai para a pasta dele, em Projetos).")
    if "variante" not in spec:
        spec["variante"] = "longo" if tipo in TIPOS_LONGOS else "curto"
    if not spec.get("codigo"):
        area = spec.get("area", "ADM").upper()
        num = spec.get("numero")
        n = f"{int(num):03d}" if num else "XXX"
        if tipo in TIPOS_CURTOS and not spec.get("ano"):
            spec["ano"] = int(spec["data_emissao"][-4:]) if re.match(r"\d{2}/\d{2}/\d{4}$", str(spec.get("data_emissao", ""))) else datetime.date.today().year
        spec["codigo"] = f"{EMPRESA_ID}-{tipo or 'DOC'}-{area}-{n}" + (f"/{spec['ano']}" if tipo in TIPOS_CURTOS else "")
    faltam = [] if EMPRESA.get("grupo") else [c for c in ("razao", "documento") if not EMPRESA.get(c)]
    if faltam:
        nomes = " e ".join({"razao": "a razão social", "documento": f'o {EMPRESA["documento_rotulo"]}'}[c] for c in faltam)
        if spec["estado"] in ("OFICIAL", "OBSOLETO"):
            raise SystemExit(f"Cadastre {nomes} de {EMPRESA['marca']} antes de emitir documento oficial.")
        av.append(f"{EMPRESA['marca']}: {nomes} ainda não está no cadastro; o cabeçalho mostra [definir].")
    if "XXX" in spec["codigo"]:
        if spec["estado"] in ("OFICIAL", "OBSOLETO"):
            raise SystemExit("Documento oficial precisa de número. Use --lista para reservar o próximo livre, ou informe 'numero' na especificação.")
        # Minuta e em aprovação saem com XXX: o número é reservado na publicação, não é pendência.
    v = str(spec.get("versao", "1.0"))
    if spec["estado"] in ("OFICIAL", "OBSOLETO"):
        if "-r" in v:
            v = v.split("-r")[0]; av.append(f"Versão ajustada para {v}: documento {spec['estado'].lower()} não leva sufixo de rodada.")
        if spec["estado"] == "OFICIAL":
            falta = [c for c in ("aprovado_por", "data_aprovacao") if not spec.get(c)]
            if falta: raise SystemExit(f"Não é possível marcar como OFICIAL sem: {', '.join(falta)}.")
            spec["vigencia"] = spec.get("vigencia") or spec["data_aprovacao"]
            hist = spec.get("historico")
            if hist and not any(str(h.get("versao")) == v for h in hist):
                hist.append({"versao": v, "data": spec["data_aprovacao"], "descricao": f'Aprovada por {spec["aprovado_por"]}.',
                             "responsavel": spec["aprovado_por"]})
        if spec["estado"] == "OBSOLETO" and not (spec.get("substituido_por") or spec.get("motivo_obsoleto")):
            raise SystemExit("OBSOLETO exige 'substituido_por' ou 'motivo_obsoleto'.")
    else:
        if "-r" not in v:
            v += "-r1"; av.append(f"Versão ajustada para {v} (documento não oficial leva sufixo de rodada).")
    spec["versao"] = v
    spec["data_emissao"] = spec.get("data_emissao") or datetime.date.today().strftime("%d/%m/%Y")
    for b in spec.get("corpo") or []:   # local e data sem data informada: a de emissão do documento
        if isinstance(b, dict) and b.get("t") == "local_data" and not b.get("data"):
            b["data"] = spec["data_emissao"]
    for i, t in enumerate(spec.get("telegramas") or [], 1):
        if not isinstance(t, dict) or not str(t.get("titulo", "")).strip():
            raise SystemExit(f"Telegrama {i} sem título: informe 'titulo'.")
        if not t.get("corpo"):
            raise SystemExit(f'Telegrama {i} ("{t["titulo"]}") sem texto: informe o \'corpo\'.')
    spec["classificacao"] = spec.get("classificacao") or "Interno"
    acesso = spec.get("acesso_restrito")
    if isinstance(acesso, (list, tuple)):
        acesso = ", ".join(str(a).strip() for a in acesso if str(a).strip())
    acesso = re.sub(r"\s+", " ", str(acesso or "")).strip()
    if eh_confidencial(spec):
        spec["classificacao"] = "Confidencial"
        if acesso:
            spec["acesso_restrito"] = acesso
        elif spec["estado"] == "OBSOLETO":
            spec.pop("acesso_restrito", None)
            av.append("Documento confidencial sem a lista de quem pode ler: a cópia obsoleta sai sem ela.")
        else:
            raise SystemExit("Documento confidencial precisa dizer quem pode ler: informe 'acesso_restrito' "
                             "(nomes ou cargos, separados por vírgula; por exemplo, 'Diretoria, Fulano de Tal').")
    else:
        spec.pop("acesso_restrito", None)
    spec["local"] = spec.get("local") or dado("cidade")
    return av

# ----------------------------------------------------------------------------- montagem
def secao(header_rid, footer_rid, margens, header_cm, footer_cm, tipo="nextPage"):
    t, r, b, l = [int(m * TW_CM) for m in margens]
    return (f'<w:sectPr><w:headerReference w:type="default" r:id="{header_rid}"/><w:footerReference w:type="default" r:id="{footer_rid}"/>'
            f'<w:type w:val="{tipo}"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="{t}" w:right="{r}" w:bottom="{b}" w:left="{l}" '
            f'w:header="{int(header_cm * TW_CM)}" w:footer="{int(footer_cm * TW_CM)}" w:gutter="0"/>{BORDA_PAGINA}<w:cols w:space="708"/></w:sectPr>')


# Borda fina em todas as páginas, a 1 cm das bordas do papel (abaixo da tarja de estado)
_b = '<w:{0} w:val="single" w:sz="6" w:space="28" w:color="000000"/>'
BORDA_PAGINA = '<w:pgBorders w:offsetFrom="page">' + "".join(_b.format(l) for l in ("top", "left", "bottom", "right")) + "</w:pgBorders>"
# Telegrama: anexo curto em meia folha A5 (14,8 × 21 cm), no fim do documento, curto ou longo. Não tem código próprio:
# herda o código, a versão e o estado do documento. No topo, só uma linha com eles e a página; sem faixa, sem borda e
# sem rodapé. A marca d'água do estado vem na mesma proporção da página A4.
A5_TW = (8391, 11906)
MARG_TELEGRAMA = (2.0, 1.5, 1.5, 1.5)   # topo (linha de controle), direita, base, esquerda — cm
MARCA_A5 = round(470.0 * 14.8 / 21, 1)
RE_TELEGRAMA = re.compile(r"^TELEGRAMA(?:\s+\d+)?\s*[–-]\s*(.+)$", re.I)


def secao_telegrama(header_rid, footer_rid):
    t, r, b, l = [int(m * TW_CM) for m in MARG_TELEGRAMA]
    return (f'<w:sectPr><w:headerReference w:type="default" r:id="{header_rid}"/><w:footerReference w:type="default" r:id="{footer_rid}"/>'
            f'<w:type w:val="nextPage"/><w:pgSz w:w="{A5_TW[0]}" w:h="{A5_TW[1]}"/><w:pgMar w:top="{t}" w:right="{r}" w:bottom="{b}" w:left="{l}" '
            f'w:header="{int(0.8 * TW_CM)}" w:footer="{int(0.8 * TW_CM)}" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')


def cabecalho_telegrama(ctx, spec):
    """A linha do topo do telegrama: código e versão, estado na cor dele e página, com um filete embaixo."""
    est = ESTADOS[spec["estado"]]
    sep = run("  ·  ", sz=8, cor=sec("404040"))
    return par(marca(ctx, spec, MARCA_A5) + run(codigo_versao(spec), sz=8, cor=sec("404040")) + sep + run(est["rotulo"].upper(), sz=8, b=True, cor=est["cor"])
               + sep + (marca_confidencial() if eh_confidencial(spec) else "") + pagina_x_de_y(sz=8, cor=cor_texto(), b=True), jc="center", antes=0, depois=0, linha=240, estilo="Rodape",
               borda='<w:bottom w:val="single" w:sz="4" w:space="4" w:color="BFBFBF"/>')


def telegramas_de(spec):
    return [t for t in spec.get("telegramas") or [] if isinstance(t, dict)]


def corpo_telegrama(corpo):
    """Título dentro do telegrama sai em negrito, sem número e fora do sumário: o telegrama não tem seções."""
    return [{"t": "p", "texto": f'**{str(b.get("texto", "")).replace("*", "")}**', "sem_recuo": True}
            if isinstance(b, dict) and re.match(r"^h[1-5]$", str(b.get("t"))) else b for b in corpo or []]


def fechar_com_telegramas(ctx, spec, sect, partes_extra, D):
    """Fecha a última seção do documento. Com telegramas, ela fecha num parágrafo vazio e cada telegrama segue em A5,
    numa seção só, um por página."""
    tels = telegramas_de(spec)
    if not tels:
        return sect
    hdr, ftr = "word/header9.xml", "word/footer9.xml"
    partes_extra[hdr] = ("hdr", cabecalho_telegrama(ctx, spec))
    partes_extra[ftr] = ("ftr", par("", antes=0, depois=0, linha=240, estilo="Rodape"))
    rid_h, rid_f = ctx.rel(D, "header", "header9.xml"), ctx.rel(D, "footer", "footer9.xml")
    larg = int((A5_TW[0] / TW_CM - MARG_TELEGRAMA[1] - MARG_TELEGRAMA[3]) * TW_CM)
    out = par("", antes=0, depois=0, linha=240, sect=sect)
    for i, t in enumerate(tels):
        out += par(run(f'Telegrama {i + 1} – {t.get("titulo", "")}'), estilo="TituloTelegrama", quebra_antes=i > 0)
        out += blocos(ctx, corpo_telegrama(t.get("corpo")), D, larg, False)
    return out + secao_telegrama(rid_h, rid_f)


MARG_TIMBRADO = (3.3, 2.0, 3.2, 3.0)   # topo (faixa + bloco de controle), direita, base (logo), esquerda — cm
MARG_ABNT = (3.3, 2.0, 2.6, 3.0)   # topo (faixa + bloco de controle), direita, base (logo e título), esquerda — cm


def montar(spec, ativos, paginas_sumario=None):
    ctx = Ctx()
    longo = spec["variante"] == "longo"
    D = "word/document.xml"
    corpo = ""
    larg_t = int((21 - MARG_TIMBRADO[1] - MARG_TIMBRADO[3]) * TW_CM)
    larg_a = int((21 - MARG_ABNT[1] - MARG_ABNT[3]) * TW_CM)
    est = ESTADOS[spec["estado"]]
    tipo_nome = TIPOS_CURTOS.get(spec.get("tipo", ""), TIPOS_LONGOS.get(spec.get("tipo", ""), spec.get("tipo_nome", "Documento")))
    tipo_nome = spec.get("tipo_nome", tipo_nome)

    hdr1 = "word/header1.xml"; ftr1 = "word/footer1.xml"
    cab1 = cabecalho_controle(ctx, spec, MARG_TIMBRADO[3])
    rod1 = rodape_logo(ctx, ftr1, ativos)
    partes_extra = {hdr1: ("hdr", cab1), ftr1: ("ftr", rod1)}
    rid_h1 = ctx.rel(D, "header", "header1.xml"); rid_f1 = ctx.rel(D, "footer", "footer1.xml")

    if not longo:
        corpo += par(run(tipo_nome.upper() + (f' – {spec["titulo"]}' if spec.get("titulo") else ""), b=True), estilo="TituloDocumento")
        corpo += par(run(f'{spec["codigo"]}  ·  Versão {spec["versao"]}  ·  Emissão {spec["data_emissao"]}  ·  ', sz=9, cor=sec("404040")) +
                     run(est["rotulo"].upper(), sz=9, b=True, cor=est["cor"])
                     + (('<w:r><w:br/></w:r>' + run("CONFIDENCIAL", sz=9, b=True, cor=COR_CONFIDENCIAL)
                         + run(f'  ·  Acesso restrito a: {spec.get("acesso_restrito", "")}', sz=9, cor=sec("404040"))) if eh_confidencial(spec) else ""),
                     jc="center", antes=0, depois=240, linha=240, estilo="SemRecuo")
        if spec.get("identificacao"):
            ident = spec["identificacao"]   # empresa e documento já estão no cabeçalho
            corpo += bloco_campos(ident, larg_t, 3.3)
        corpo += blocos(ctx, spec.get("corpo", []), D, larg_t, False)
        corpo += fechar_com_telegramas(ctx, spec, secao(rid_h1, rid_f1, MARG_TIMBRADO, 1.2, 1.2, "nextPage").replace("<w:type w:val=\"nextPage\"/>", ""), partes_extra, D)
    else:
        # ---------------- capa (timbrado completo)
        altura = 29.7 - MARG_TIMBRADO[0] - MARG_TIMBRADO[2]
        def linha_capa(h_cm, conteudo, va):
            return (f'<w:tr><w:trPr><w:trHeight w:val="{int(h_cm * TW_CM)}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{larg_t}" w:type="dxa"/>'
                    f'<w:vAlign w:val="{va}"/></w:tcPr>{conteudo}</w:tc></w:tr>')
        cp = lambda x, **k: par(x, jc="center", antes=k.get("antes", 0), depois=k.get("depois", 0), linha=k.get("linha", 276), estilo="SemRecuo")
        ft = VISUAL.get("fonte_titulo")
        meio = (cp(run(tipo_nome.upper(), sz=12, cor=EMPRESA["cor"], b=True, esp=40, fonte=ft), depois=240) +
                cp(run(spec["titulo"].upper(), b=True, sz=22, fonte=ft, cor=VISUAL.get("cor_titulo")), linha=300, depois=200) +
                (cp(run(spec["subtitulo"], sz=14, cor=sec("262626")), depois=480) if spec.get("subtitulo") else cp("", depois=480)))
        ctab = [["Código", spec["codigo"]], ["Versão", spec["versao"]], ["Estado", est["rotulo"].upper()], ["Classificação", spec["classificacao"]]]
        wqs = [int(larg_t * p / 4.6) for p in (1.6, 1, 1, 1)]
        quadro = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(wqs)}" w:type="dxa"/><w:jc w:val="center"/>'
                  '<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
                  f'<w:tblGrid>{grade(wqs)}</w:tblGrid><w:tr>')
        for (rot, val), wq in zip(ctab, wqs):
            cor = est["cor"] if rot == "Estado" else COR_CONFIDENCIAL if rot == "Classificação" and eh_confidencial(spec) else cor_texto()
            quadro += (f'<w:tc><w:tcPr><w:tcW w:w="{wq}" w:type="dxa"/></w:tcPr>'
                       f'{par(run(rot.upper(), sz=7, cor=sec("595959"), esp=10), jc="center", estilo="TabelaTexto", antes=60)}'
                       f'{par(run(val, b=True, sz=9.5, cor=cor), jc="center", estilo="TabelaTexto", depois=60)}</w:tc>')
        quadro += "</w:tr></w:tbl>"
        ano = spec["data_emissao"][-4:]
        base = cp(run(spec["local"], b=True)) + cp(run(ano, b=True))
        corpo += (f'<w:tbl><w:tblPr><w:tblW w:w="{larg_t}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                  f'<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                  f'<w:tblGrid><w:gridCol w:w="{larg_t}"/></w:tblGrid>'
                  f'{linha_capa(1.0, cp(""), "top")}{linha_capa(altura - 4.2, meio + quadro + cp("", linha=240), "center")}{linha_capa(2.8, base, "bottom")}</w:tbl>')
        corpo += par("", antes=0, depois=0, linha=240, sect=secao(rid_h1, rid_f1, MARG_TIMBRADO, 1.2, 1.2))

        # ---------------- seção ABNT (cabeçalho reduzido)
        hdr2 = "word/header2.xml"; ftr2 = "word/footer2.xml"
        partes_extra[hdr2] = ("hdr", cabecalho_controle(ctx, spec, MARG_ABNT[3]))
        partes_extra[ftr2] = ("ftr", rodape_reduzido(ctx, ftr2, ativos, spec, larg_a))
        rid_h2 = ctx.rel(D, "header", "header2.xml"); rid_f2 = ctx.rel(D, "footer", "footer2.xml")

        # folha de controle
        corpo += par(run("FOLHA DE CONTROLE"), estilo="TituloSumario")
        ident = [["Título", spec["titulo"] + (f' – {spec["subtitulo"]}' if spec.get("subtitulo") else "")],
                 ["Tipo", tipo_nome], ["Código", spec["codigo"]], ["Versão", spec["versao"]],
                 ["Estado", f'**{est["rotulo"].upper()}**'], ["Classificação", spec["classificacao"]]]
        if eh_confidencial(spec): ident.append(["Acesso restrito a", spec.get("acesso_restrito", "")])
        ident += [["Área responsável", spec.get("area_nome") or AREAS.get(str(spec.get("area", "")).upper(), spec.get("area", ""))],
                 ["Elaborado por", spec.get("elaborado_por", "")], ["Revisado por", spec.get("revisado_por", "")],
                 ["Aprovado por", spec.get("aprovado_por") or "Pendente – documento não aprovado"],
                 ["Data de emissão", spec["data_emissao"]],
                 ["Vigência", (spec.get("vigencia") or "A partir da aprovação") if spec["estado"] != "OBSOLETO" else "Encerrada"]]
        if spec.get("substitui"): ident.append(["Substitui", spec["substitui"]])
        if spec.get("substituido_por"): ident.append(["Substituído por", spec["substituido_por"]])
        if spec.get("motivo_obsoleto"): ident.append(["Motivo da obsolescência", spec["motivo_obsoleto"]])
        corpo += bloco_campos(ident, larg_a, 4.5)
        hist = spec.get("historico") or [{"versao": spec["versao"], "data": spec["data_emissao"], "descricao": "Emissão inicial.", "responsavel": spec.get("elaborado_por", "")}]
        corpo += par(run("Histórico de revisões", b=True), estilo="SemRecuo", antes=240, depois=120, keep=True)
        corpo += bloco_tabela(ctx, {
            "tipo": "Quadro", "cabecalho": ["Versão", "Data", "O que mudou", "Responsável"], "larguras": [1.4, 1.6, 5, 2.4],
            "linhas": [[h.get("versao", ""), h.get("data", ""), h.get("descricao", ""), h.get("responsavel", "")] for h in hist]}, larg_a, legenda=False)
        if spec["estado"] == "OFICIAL":
            corpo += par(run("Aprovação", b=True), estilo="SemRecuo", antes=240, depois=0, keep=True)
            corpo += bloco_assinaturas(spec.get("assinaturas_aprovacao") or [{"nome": spec["aprovado_por"], "cargo": spec.get("cargo_aprovador", "Aprovador"), "documento": f'Aprovado em {spec["data_aprovacao"]}'}], larg_a)

        # sumário
        if spec.get("sumario", True):
            corpo += par(run("SUMÁRIO"), estilo="TituloSumario", quebra_antes=True)
            corpo += "@@SUMARIO@@"

        conteudo = blocos(ctx, spec.get("corpo", []), D, larg_a, True)
        conteudo = conteudo.replace('<w:pStyle w:val="Heading1"/>', '<w:pStyle w:val="Heading1"/><w:pageBreakBefore/>', 1)
        corpo += conteudo
        if spec.get("referencias"):
            corpo += titulo_sem_numero(ctx, "Referências")
            for ref in spec["referencias"]:
                corpo += par(inline(ref), estilo="Referencia")
        if spec.get("glossario"):
            corpo += titulo_sem_numero(ctx, "Glossário")
            for termo, defi in spec["glossario"]:
                corpo += par(run(termo + ": ", b=True) + inline(defi), estilo="SemRecuo", depois=120)
        for chave, rotulo in (("apendices", "APÊNDICE"), ("anexos", "ANEXO")):
            for k, ap in enumerate(spec.get(chave) or []):
                letra = chr(ord("A") + k)
                corpo += titulo_sem_numero(ctx, f'{rotulo} {letra} – {ap["titulo"]}')
                corpo += blocos(ctx, ap.get("corpo", []), D, larg_a, False)

        # sumário: entradas em cache (números de página da renderização prévia, se houver)
        entradas = ""
        for i, (nivel, texto, marca) in enumerate([t for t in ctx.titulos if t[0] <= 3]):
            pg = str(paginas_sumario[i]) if paginas_sumario and i < len(paginas_sumario) and paginas_sumario[i] else ""
            conteudo_r = run(texto) + (run("\t" + pg) if pg else "")
            conteudo_r = f'<w:hyperlink w:anchor="{marca}" w:history="1">{conteudo_r}</w:hyperlink>'
            entradas += par(conteudo_r, estilo=f"TOC{nivel}", tabs=[("right", larg_a, "dot")])
        if entradas:
            ini = (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                   f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
            entradas = entradas.replace("</w:pPr>", "</w:pPr>" + ini, 1)
            k = entradas.rfind("</w:p>")
            entradas = entradas[:k] + '<w:r><w:fldChar w:fldCharType="end"/></w:r>' + entradas[k:]
        corpo = corpo.replace("@@SUMARIO@@", entradas)
        corpo += fechar_com_telegramas(ctx, spec, secao(rid_h2, rid_f2, MARG_ABNT, 1.2, 1.2), partes_extra, D)

    documento = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{corpo}</w:body></w:document>'
    ctx.rel(D, "styles", "styles.xml"); ctx.rel(D, "numbering", "numbering.xml"); ctx.rel(D, "settings", "settings.xml")
    return ctx, documento, partes_extra


TIPOS_REL = {"image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
             "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
             "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
             "styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
             "numbering": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
             "settings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
             "customXml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
             "fontTable": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
             "font": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"}


def chave_fonte(dados):
    """GUID da fonte embutida, tirado do próprio arquivo (o mesmo arquivo dá o mesmo documento)."""
    import zlib
    b = b"".join(struct.pack(">I", zlib.crc32(bytes([k]) + dados)) for k in range(4))
    h = b.hex().upper()
    return "{%s-%s-%s-%s-%s}" % (h[:8], h[8:12], h[12:16], h[16:20], h[20:])


def ofuscar_fonte(dados, chave):
    """Fonte embutida do Word (ECMA-376, 17.8.1): os 32 primeiros bytes em XOR com a chave, bytes ao contrário."""
    k = bytes.fromhex(chave.strip("{}").replace("-", ""))[::-1]
    d = bytearray(dados)
    for i in range(32):
        d[i] ^= k[i % 16]
    return bytes(d)


def partes_fontes(ctx, fontes):
    """fontTable.xml e as fontes ofuscadas (word/fonts/fontN.odttf), na ordem do visual: texto, depois títulos."""
    if not fontes:
        return None, []
    tag = {"regular": "embedRegular", "negrito": "embedBold", "italico": "embedItalic", "negrito_italico": "embedBoldItalic"}
    familias = [f for f in (VISUAL.get("fonte_texto"), VISUAL.get("fonte_titulo")) if f]
    familias = [f for i, f in enumerate(familias) if f not in familias[:i]]
    xml, arquivos = "", []
    for fam in familias:
        emb = ""
        for est in ESTILOS_FONTE:
            dados = next((d for f, e, d in fontes if f == fam and e == est), None)
            if dados is None:
                continue
            nome = f"font{len(arquivos) + 1}.odttf"; chave = chave_fonte(dados)
            rid = ctx.rel("word/fontTable.xml", "font", f"fonts/{nome}")
            emb += f'<w:{tag[est]} r:id="{rid}" w:fontKey="{chave}"/>'
            arquivos.append((f"word/fonts/{nome}", ofuscar_fonte(dados, chave)))
        xml += f'<w:font w:name="{escape(fam)}"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/>{emb}</w:font>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {NS}>{xml}</w:fonts>', arquivos


VERSAO_TIMBRADOR = "2026-11-03"
CAMPOS_FICHA = ["empresa", "tipo", "area", "area_nome", "numero", "ano", "codigo", "versao", "estado", "variante",
                "titulo", "subtitulo", "titulo_curto", "classificacao", "acesso_restrito", "local", "data_emissao",
                "elaborado_por", "revisado_por", "aprovado_por", "cargo_aprovador", "data_aprovacao", "vigencia",
                "substitui", "substituido_por", "motivo_obsoleto", "identificacao", "historico",
                "modelo", "com_assinatura", "envolvidos", "assinatura_formato", "assinaturas_previstas",
                "com_anexos", "anexos_externos", "projeto", "plano", "rodadas"]   # modelo: True só nos arquivos de partida dos modelos oficiais

# Campos de um arquivo de partida que NÃO passam para o documento escrito a partir dele.
CAMPOS_SO_DO_MODELO = ("modelo", "codigo", "numero", "ano", "versao", "estado", "titulo", "subtitulo", "titulo_curto",
                       "elaborado_por", "revisado_por", "aprovado_por", "cargo_aprovador", "data_aprovacao", "vigencia",
                       "substitui", "substituido_por", "motivo_obsoleto", "historico", "data_emissao")
# Namespace da ficha embutida. A leitura aceita qualquer namespace terminado em "timbrador:ficha", inclusive o
# dos documentos emitidos antes do white label.
NS_FICHA = "urn:timbrador:ficha"
MARCA_FICHA = "timbrador:ficha"


def ficha_de(spec):
    f = {k: spec[k] for k in CAMPOS_FICHA if spec.get(k) not in (None, "", [])}
    if ORG["politica"]: f["padrao"] = ORG["politica"]
    f["timbrador"] = VERSAO_TIMBRADOR
    return f


def partes_ficha(spec):
    """Ficha do documento: JSON completo em customXml (preservado pelo Word ao salvar) e propriedades
    personalizadas curtas, visíveis em Arquivo > Informações > Propriedades."""
    ficha = ficha_de(spec)
    item = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ficha xmlns="{NS_FICHA}">'
            f'{escape(json.dumps(ficha, ensure_ascii=False))}</ficha>')
    props = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ds:datastoreItem ds:itemID="{6F2A1C3E-9B1D-4C55-8E0A-3C1A7B2D9F10}" '
             'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="' + NS_FICHA + '"/></ds:schemaRefs></ds:datastoreItem>')
    visiveis = [("Timbrador - código", spec["codigo"]), ("Timbrador - versão", spec["versao"]),
                ("Timbrador - estado", ESTADOS[spec["estado"]]["rotulo"]), ("Timbrador - empresa", spec["empresa"])]
    custom = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
              'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
              "".join(f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{i + 2}" name="{escape(n)}"><vt:lpwstr>{escape(str(v))}</vt:lpwstr></property>'
                      for i, (n, v) in enumerate(visiveis)) + "</Properties>")
    return item, props, custom


def gravar(spec, ativos, saida, paginas_sumario=None, atualizar_campos=False):
    ctx, documento, extras = montar(spec, ativos, paginas_sumario)
    ctx.rel("word/document.xml", "customXml", "../customXml/item1.xml")
    tabela_fontes, fontes_ofuscadas = partes_fontes(ctx, ativos.get("fontes"))
    if tabela_fontes:
        ctx.rel("word/document.xml", "fontTable", "fontTable.xml")
    item_ficha, props_ficha, custom_props = partes_ficha(spec)
    agora = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    est = ESTADOS[spec["estado"]]
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as z:
        ov = "".join(f'<Override PartName="/{p}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{"header" if t == "hdr" else "footer"}+xml"/>'
                     for p, (t, _) in extras.items())
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
                   '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                   + ('<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>'
                      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
                      if tabela_fontes else "") +
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                   '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                   '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                   '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                   '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
                   '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
                   f"{ov}</Types>")
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                   '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>')
        z.writestr("docProps/custom.xml", custom_props)
        z.writestr("customXml/item1.xml", item_ficha)
        z.writestr("customXml/itemProps1.xml", props_ficha)
        z.writestr("customXml/_rels/item1.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')
        z.writestr("docProps/core.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                   'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
                   'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                   f'<dc:title>{escape(spec["titulo"])}</dc:title><dc:subject>{escape(spec["codigo"])}</dc:subject>'
                   f'<dc:creator>{escape(spec.get("elaborado_por") or EMPRESA["razao"] or EMPRESA["marca"])}</dc:creator>'
                   f'<cp:keywords>{escape(spec["codigo"])}; v{escape(spec["versao"])}; {est["rotulo"]}</cp:keywords>'
                   f'<cp:contentStatus>{est["rotulo"]}</cp:contentStatus><cp:version>{escape(spec["versao"])}</cp:version>'
                   f'<dcterms:created xsi:type="dcterms:W3CDTF">{agora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{agora}</dcterms:modified>'
                   "</cp:coreProperties>")
        z.writestr("docProps/app.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                   f'<Application>Microsoft Office Word</Application><Company>{escape(EMPRESA["razao"] or EMPRESA["marca"])}</Company></Properties>')
        z.writestr("word/document.xml", documento)
        z.writestr("word/styles.xml", estilos_xml())
        z.writestr("word/numbering.xml", numeracao_xml(ctx))
        z.writestr("word/settings.xml", settings_xml(atualizar_campos, bool(tabela_fontes)))
        if tabela_fontes:
            z.writestr("word/fontTable.xml", tabela_fontes)
            for nome, dados in fontes_ofuscadas:
                z.writestr(nome, dados)
        for parte, (t, conteudo) in extras.items():
            raiz = "hdr" if t == "hdr" else "ftr"
            z.writestr(parte, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{raiz} {NS}>{conteudo}</w:{raiz}>')
        for parte, lst in ctx.rels.items():
            d, b = os.path.split(parte)
            z.writestr(f"{d}/_rels/{b}.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                       "".join(f'<Relationship Id="{rid}" Type="{TIPOS_REL[t]}" Target="{a}"/>' for rid, t, a in lst) + "</Relationships>")
        for nome, dados in ctx.midia:
            z.writestr(f"word/media/{nome}", dados)
    return ctx


def paginas_dos_titulos(docx, titulos):
    """Renderiza com LibreOffice e localiza a página de cada título (para o sumário)."""
    if not comando_soffice() or not shutil.which("pdftotext"):
        return None
    tmp = tempfile.mkdtemp()
    try:
        pdf = para_pdf(docx, tmp)
        if not pdf: return None
        txt = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout
        paginas = txt.split("\f")
        norm = lambda s: re.sub(r"\s+", " ", s).strip().upper()
        pags_n = [norm(p) for p in paginas]
        inicio = next((i for i, p in enumerate(pags_n) if "SUMÁRIO" in p), 0) + 1
        res, pos = [], inicio
        for _tit in titulos:
            t = _tit[1]
            alvo = norm(t)[:60]
            achou = None
            for i in range(pos, len(pags_n)):
                if alvo in pags_n[i]:
                    achou = i; break
            if achou is None:
                res.append(None)
            else:
                res.append(achou + 1); pos = achou
        return res
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def nome_arquivo(spec):
    slug = re.sub(r"[^a-z0-9]+", "-", spec["titulo"].lower().translate(str.maketrans("áàâãéêíóôõúüç", "aaaaeeiooouuc"))).strip("-")[:50]
    cod = _parte_nome(spec["codigo"])
    return f'{cod}_v{_parte_nome(spec["versao"])}_{_parte_nome(spec["estado"])}_{slug}.docx'


# ----------------------------------------------------------------------------- leitura de documento timbrado
import xml.etree.ElementTree as ET
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
RE_CODIGO = re.compile(r"\b([A-Z0-9]{2,6})-([A-Z]{2,3})-([A-Z]{2,3})-(\d{3}|XXX)(?:/(\d{4}))?\s+·\s+(?:v|[Vv]ersão\s+)([0-9][0-9.]*(?:-r\d+)?)")


def _wval(el, tag):
    x = el.find(_W + tag) if el is not None else None
    return x.get(_W + "val") if x is not None else None


def _texto_par(p, markdown=True):
    segs = []
    for r in p.iter(_W + "r"):
        t = "".join((c.text or "") if c.tag == _W + "t" else (" " if c.tag in (_W + "tab", _W + "br") else "") for c in r)
        if not t:
            continue
        rpr = r.find(_W + "rPr")
        b = rpr is not None and rpr.find(_W + "b") is not None and _wval(rpr, "b") not in ("0", "false")
        i = rpr is not None and rpr.find(_W + "i") is not None and _wval(rpr, "i") not in ("0", "false")
        fonte = rpr.find(_W + "rFonts").get(_W + "ascii") if rpr is not None and rpr.find(_W + "rFonts") is not None else None
        if segs and segs[-1][1:3] == [b, i]:
            segs[-1][0] += t
        else:
            segs.append([t, b, i, fonte])
    todo_b = bool(segs) and all(sg[1] or not sg[0].strip() for sg in segs)
    out = ""
    for t, b, i, _ in segs:
        if not markdown or not t.strip() or (not b and not i) or (todo_b and not i):
            out += t; continue
        m = re.match(r"^(\s*)(.*?)(\s*)$", t, re.S)
        mk = "***" if b and i else "**" if b else "*"
        out += f"{m.group(1)}{mk}{m.group(2)}{mk}{m.group(3)}"
    fonte = segs[0][3] if segs else None
    return re.sub(r"\s+", " ", out).strip(), fonte


def _data_numerica(txt):
    m = re.search(r"(\d{1,2}) de (\w+) de (\d{4})", txt)
    if m and m.group(2).lower() in MESES:
        return f"{int(m.group(1)):02d}/{MESES.index(m.group(2).lower()) + 1:02d}/{m.group(3)}"
    return None


def ler_timbrado(caminho, pasta_midia=None):
    """Lê um .docx do padrão e devolve (spec, info). A ficha embutida tem prioridade; sem ela, os dados vêm
    do cabeçalho, da faixa e da folha de controle. O corpo é sempre relido do documento (vale o que foi editado)."""
    z = zipfile.ZipFile(caminho)
    nomes = z.namelist()
    info = {"origem": "sem_timbrado", "avisos": []}
    spec = {}
    if "customXml/item1.xml" in nomes or any(n.startswith("customXml/item") for n in nomes):
        for n in nomes:
            if re.match(r"customXml/item\d+\.xml$", n):
                x = z.read(n).decode("utf8")
                if MARCA_FICHA in x:
                    raiz = ET.fromstring(x)
                    spec = json.loads(raiz.text or "{}")
                    info["origem"] = "ficha"
                    ma = re.search(r"/(\d{4})$", str(spec.get("codigo", "")))
                    if ma and not spec.get("ano"): spec["ano"] = int(ma.group(1))
                    break
    cab = " ".join(re.sub(r"<[^>]+>", " ", z.read(n).decode("utf8")) for n in nomes if re.match(r"word/(header|footer)\d*\.xml$", n))
    cab = re.sub(r"\s+", " ", cab)
    if info["origem"] != "ficha":
        m = RE_CODIGO.search(cab)
        if m:
            info["origem"] = "cabecalho"
            spec.update({"empresa": m.group(1), "tipo": m.group(2), "area": m.group(3), "versao": m.group(6)})
            if m.group(4) != "XXX": spec["numero"] = int(m.group(4))
            if m.group(5): spec["ano"] = int(m.group(5))
            for chave, est in ESTADOS.items():
                if est["tarja"].split("  ·  ")[0] in cab:
                    spec["estado"] = chave
            if "DOCUMENTO OFICIAL" in cab: spec["estado"] = "OFICIAL"
            mo = re.search(r"Aprovado por (.+?) em (\d{2}/\d{2}/\d{4})", cab)
            if mo: spec["aprovado_por"], spec["data_aprovacao"] = mo.group(1).strip(), mo.group(2)
            ms = re.search(r"SUBSTITUÍDO POR (\S+)", cab)
            if ms: spec["substituido_por"] = ms.group(1)
            info["avisos"].append("Documento sem ficha (gerado antes do Timbrador gravar a ficha): dados lidos do cabeçalho; confira título e responsáveis.")
    rels = {}
    if "word/_rels/document.xml.rels" in nomes:
        for m in re.finditer(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', z.read("word/_rels/document.xml.rels").decode("utf8")):
            rels[m.group(1)] = m.group(2)
    abstr = {}
    if "word/numbering.xml" in nomes:
        num = ET.fromstring(z.read("word/numbering.xml"))
        fmt_abs = {}
        for a in num.findall(_W + "abstractNum"):
            lv = a.find(_W + "lvl")
            txt = _wval(lv, "lvlText") or ""
            fmt_abs[a.get(_W + "abstractNumId")] = "checklist" if "☐" in txt else {"lowerLetter": "alineas", "decimal": "numerada"}.get(_wval(lv, "numFmt"), "marcadores")
        for n_ in num.findall(_W + "num"):
            abstr[n_.get(_W + "numId")] = fmt_abs.get(_wval(n_, "abstractNumId"), "marcadores")

    corpo_xml = ET.fromstring(z.read("word/document.xml")).find(_W + "body")
    elementos = []
    def achatar(el):
        for c in el:
            if c.tag == _W + "sdt":
                cont = c.find(_W + "sdtContent")
                if cont is not None: achatar(cont)
            elif c.tag in (_W + "p", _W + "tbl"):
                elementos.append(c)
    achatar(corpo_xml)
    # capa do documento longo: tudo até o parágrafo que fecha a 1ª seção. Com telegramas, o parágrafo que fecha o
    # corpo, logo antes do primeiro telegrama, não conta: o documento curto também o tem.
    def _estilo_de(el):
        ppr = el.find(_W + "pPr") if el.tag == _W + "p" else None
        st = ppr.find(_W + "pStyle") if ppr is not None else None
        return st.get(_W + "val") if st is not None else None
    i_tel = next((i for i, el in enumerate(elementos) if _estilo_de(el) == "TituloTelegrama"), len(elementos))
    secoes = [i for i, el in enumerate(elementos[:i_tel])
              if el.tag == _W + "p" and el.find(_W + "pPr") is not None and el.find(_W + "pPr").find(_W + "sectPr") is not None]
    if i_tel < len(elementos) and secoes:
        secoes.pop()
    if secoes:
        elementos = elementos[secoes[0] + 1:]
        spec.setdefault("variante", "longo")

    corpo, destino = [], None     # destino: None=corpo | "ref" | "glo" | ("ap", bloco)
    legenda = None
    modo_pular = None             # "controle" | "sumario"
    campos_controle, hist = {}, []
    ident_lida = False
    figura_n = 0

    def alvo():
        return destino[1]["corpo"] if isinstance(destino, tuple) else corpo

    ultimo_lista = None
    for el in elementos:
        if el.tag == _W + "tbl":
            linhas = []
            cabecalho = False
            celulas_bordas = False
            for tr in el.findall(_W + "tr"):
                trpr = tr.find(_W + "trPr")
                if trpr is not None and trpr.find(_W + "tblHeader") is not None: cabecalho = True
                linha = []
                for tc in tr.findall(_W + "tc"):
                    ps = tc.findall(_W + "p")
                    textos = [_texto_par(p)[0] for p in ps]
                    if any(p.find(_W + "pPr") is not None and p.find(_W + "pPr").find(_W + "pBdr") is not None for p in ps):
                        celulas_bordas = True
                    linha.append([t for t in textos if t])
                linhas.append(linha)
            ultimo_lista = None
            if modo_pular == "controle":
                if len(linhas) and len(linhas[0]) == 2 and not cabecalho:
                    for l in linhas:
                        if len(l) == 2: campos_controle[" ".join(l[0]).replace("*", "")] = " ".join(l[1]).replace("*", "")
                elif cabecalho:
                    for l in linhas[1:]:
                        v = [" ".join(c) for c in l] + ["", "", "", ""]
                        hist.append({"versao": v[0], "data": v[1], "descricao": v[2], "responsavel": v[3]})
                continue
            if celulas_bordas:     # bloco de assinaturas
                itens = []
                for l in linhas:
                    for c in l:
                        if c:
                            a = {"nome": c[0]}
                            for extra in c[1:3]:
                                a["documento" if re.match(r"^(CNPJ|CPF|RG|OAB|CRC)\b", extra) else ("cargo" if "cargo" not in a else "documento")] = extra
                            itens.append(a)
                if alvo() and alvo()[-1].get("t") == "assinaturas": alvo()[-1]["itens"] += itens
                else: alvo().append({"t": "assinaturas", "itens": itens})
                continue
            planas = [[" ".join(c) for c in l] for l in linhas]
            grade_ = el.find(_W + "tblGrid")
            larguras = [int(g.get(_W + "w", "0")) for g in grade_.findall(_W + "gridCol")] if grade_ is not None else []
            if legenda:
                b = {"t": "tabela", "tipo": legenda[0], "titulo": legenda[1]}
                if larguras and len(set(larguras)) > 1: b["larguras"] = [round(w / min(larguras), 2) for w in larguras]
                if cabecalho: b["cabecalho"], b["linhas"] = planas[0], planas[1:]
                else: b["linhas"] = planas
                alvo().append(b); legenda = None
                continue
            if all(len(l) == 2 for l in planas) and planas and all(len(l[0]) <= 40 for l in planas):
                itens = [[l[0].replace("*", ""), l[1]] for l in planas]
                if not ident_lida and not corpo and spec.get("variante", "curto") != "longo":
                    spec["identificacao"] = [i for i in itens if i[0] != "Empresa"]
                    ident_lida = True
                else:
                    alvo().append({"t": "campos", "itens": itens})
                continue
            b = {"t": "tabela", "tipo": "Quadro"}
            if larguras and len(set(larguras)) > 1: b["larguras"] = [round(w / min(larguras), 2) for w in larguras]
            if cabecalho: b["cabecalho"], b["linhas"] = planas[0], planas[1:]
            else: b["linhas"] = planas
            alvo().append(b)
            continue

        ppr = el.find(_W + "pPr")
        estilo = _wval(ppr, "pStyle") if ppr is not None else None
        texto, fonte = _texto_par(el)
        puro = texto.replace("*", "")
        # figuras do corpo (logo fica no rodapé e não aparece aqui)
        blip = el.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
        if blip is not None and not texto:
            rid = blip.get(_R + "embed")
            if rid in rels and pasta_midia:
                os.makedirs(pasta_midia, exist_ok=True)
                figura_n += 1
                alvo_arq = os.path.join(pasta_midia, f"figura{figura_n}{os.path.splitext(rels[rid])[1]}")
                open(alvo_arq, "wb").write(z.read("word/" + rels[rid]))
                b = {"t": "figura", "caminho": alvo_arq}
                if legenda: b["tipo"], b["titulo"] = legenda; legenda = None
                alvo().append(b)
            elif rid in rels:
                info["avisos"].append("Há figura no corpo; use --pasta-midia para extraí-la.")
            continue
        if not puro:
            continue
        if estilo == "TituloTelegrama":
            m = RE_TELEGRAMA.match(puro)
            bloco = {"titulo": (m.group(1) if m else puro).strip(), "corpo": []}
            spec.setdefault("telegramas", []).append(bloco); destino = ("ap", bloco); modo_pular = None; ultimo_lista = None
            continue
        if estilo == "TituloSumario":
            modo_pular = "controle" if puro.upper().startswith("FOLHA DE CONTROLE") else "sumario"
            continue
        if modo_pular and estilo in ("TOC1", "TOC2", "TOC3"):
            continue
        if modo_pular == "controle" and estilo in ("SemRecuo", "TabelaTexto", None) and puro in ("Histórico de revisões", "Aprovação"):
            continue
        if modo_pular == "controle" and estilo not in (None, "Normal") and not (estilo or "").startswith("Heading") and estilo != "TituloSemNumero":
            continue
        if estilo and (estilo.startswith("Heading") or estilo == "TituloSemNumero"):
            modo_pular = None
        if estilo == "TituloDocumento":
            m = re.match(r"^(.*?)\s+[–-]\s+(.*)$", puro)
            if m: spec.setdefault("titulo", m.group(2).strip())
            continue
        if RE_CODIGO.search(puro) and "Emissão" in puro:
            continue
        if estilo and estilo.startswith("Heading"):
            nivel = int(estilo[7:] or 1)
            alvo().append({"t": f"h{min(nivel, 5)}", "texto": re.sub(r"^\d+(\.\d+)*\s+", "", puro)})
            ultimo_lista = None; continue
        if estilo == "TituloSemNumero":
            up = puro.upper()
            if up.startswith("REFERÊNCIAS"): destino = "ref"; spec["referencias"] = []
            elif up.startswith("GLOSSÁRIO"): destino = "glo"; spec["glossario"] = []
            else:
                m = re.match(r"^(APÊNDICE|ANEXO)\s+[A-Z]+\s+[–-]\s+(.*)$", puro, re.I)
                if m:
                    chave = "apendices" if m.group(1).upper().startswith("AP") else "anexos"
                    bloco = {"titulo": m.group(2).strip().capitalize(), "corpo": []}
                    spec.setdefault(chave, []).append(bloco); destino = ("ap", bloco)
                else:
                    alvo().append({"t": "h1", "texto": puro, "numerar": False})
            continue
        if estilo == "Legenda":
            m = re.match(r"^(Tabela|Quadro|Figura|Gráfico)\s+\d+\s+[–-]\s*(.*)$", puro)
            legenda = (m.group(1), m.group(2)) if m else ("Quadro", puro)
            continue
        if estilo == "Fonte":
            ult = alvo()[-1] if alvo() else None
            if ult and ult.get("t") in ("tabela", "figura"):
                if puro.startswith("Fonte:"): ult["fonte"] = texto.split(":", 1)[1].strip()
                elif puro.startswith("Nota:"): ult["nota"] = texto.split(":", 1)[1].strip()
            continue
        if destino == "ref" and estilo == "Referencia":
            spec["referencias"].append(texto); continue
        if destino == "glo":
            m = re.match(r"^\*\*(.+?):\*\*\s*(.*)$", texto) or re.match(r"^(.+?):\s*(.*)$", puro)
            if m: spec["glossario"].append([m.group(1).replace("*", ""), m.group(2)])
            continue
        num_pr = ppr.find(_W + "numPr") if ppr is not None else None
        if num_pr is not None:
            tipo_l = abstr.get(_wval(num_pr, "numId"), "marcadores")
            ilvl = int(_wval(num_pr, "ilvl") or 0)
            if ilvl > 0 and ultimo_lista and ultimo_lista["itens"]:
                ult = ultimo_lista["itens"][-1]
                if isinstance(ult, str): ultimo_lista["itens"][-1] = ult = {"texto": ult}
                ult.setdefault("sub", []).append(texto)
            else:
                if not (ultimo_lista and ultimo_lista["t"] == tipo_l and alvo() and alvo()[-1] is ultimo_lista):
                    ultimo_lista = {"t": tipo_l, "itens": []}; alvo().append(ultimo_lista)
                ultimo_lista["itens"].append(texto)
            continue
        ultimo_lista = None
        if estilo == "CitacaoLonga":
            alvo().append({"t": "citacao", "texto": texto}); continue
        if estilo == "Nota":
            m = re.match(r"^\*\*(.+?):\*\*\s*(.*)$", texto)
            alvo().append({"t": "nota", "rotulo": m.group(1), "texto": m.group(2)} if m else {"t": "nota", "texto": texto}); continue
        if fonte == "Courier New":
            ult = alvo()[-1] if alvo() else None
            if ult and ult.get("t") == "codigo": ult["texto"] += "\n" + puro
            else: alvo().append({"t": "codigo", "texto": puro})
            continue
        jc = _wval(ppr, "jc") if ppr is not None else None
        if jc == "right" and _data_numerica(puro) and "," in puro:
            b = {"t": "local_data", "data": _data_numerica(puro)}
            local = puro.split(",")[0].strip()
            if local and local != spec.get("local"): b["local"] = local
            alvo().append(b); continue
        if ppr is not None and ppr.find(_W + "pBdr") is not None and ppr.find(_W + "pBdr").find(_W + "left") is not None:
            m = re.match(r"^\*\*(.+?):\*\*\s*(.*)$", texto)
            # rótulo seguido de texto que também começa em negrito: os dois negritos se fundem na leitura
            m2 = None if m else re.match(r"^\*\*([^*:]{1,60}):\s+(.*)$", texto)
            cor = ppr.find(_W + "pBdr").find(_W + "left").get(_W + "color", "")
            nivel = {"C99700": "atencao", "B00020": "erro"}.get(cor.upper(), "info")
            if m: alvo().append({"t": "destaque", "rotulo": m.group(1), "texto": m.group(2), "nivel": nivel})
            elif m2: alvo().append({"t": "destaque", "rotulo": m2.group(1), "texto": "**" + m2.group(2), "nivel": nivel})
            else: alvo().append({"t": "destaque", "texto": texto, "nivel": nivel})
            continue
        b = {"t": "p", "texto": texto}
        if estilo == "SemRecuo": b["sem_recuo"] = True
        if jc in ("center", "right"): b["alinhar"] = jc
        alvo().append(b)

    mapa = {"Título": "titulo", "Classificação": "classificacao", "Acesso restrito a": "acesso_restrito", "Área responsável": "area_nome", "Elaborado por": "elaborado_por",
            "Revisado por": "revisado_por", "Aprovado por": "aprovado_por", "Data de emissão": "data_emissao", "Substitui": "substitui",
            "Substituído por": "substituido_por", "Motivo da obsolescência": "motivo_obsoleto"}
    for rot, chave in mapa.items():
        v = campos_controle.get(rot)
        if v and not v.startswith("Pendente") and chave not in spec:
            spec[chave] = v
    if "titulo" in spec and campos_controle.get("Título") and " – " in campos_controle["Título"] and not spec.get("subtitulo") and info["origem"] != "ficha":
        spec["titulo"], spec["subtitulo"] = campos_controle["Título"].split(" – ", 1)
    if hist and not spec.get("historico"):
        spec["historico"] = hist
    spec["corpo"] = extrair_anexos(corpo)     # a seção Anexos é refeita a partir da lista da ficha
    if info["origem"] == "sem_timbrado":
        info["avisos"].append("Não é um documento do padrão: conteúdo lido como documento novo.")
    if spec.get("codigo") is None and spec.get("tipo"):
        pass
    if spec.get("modelo"):
        # Arquivo de partida: vale o tipo e o corpo. Código, versão, título e responsáveis eram do modelo.
        for k in CAMPOS_SO_DO_MODELO:
            spec.pop(k, None)
        spec["corpo"] = [b for b in spec.get("corpo", []) if not (b.get("t") == "nota" and b.get("rotulo") == "Como usar")]
        info["origem"] = "modelo"
        info["avisos"].append("Arquivo de partida de modelo: lido como documento novo. Defina o título e substitua os colchetes; o código sai com XXX.")
    return spec, info


# ----------------------------------------------------------------------------- lista mestra
# Mesmo arquivo usado pela página timbrador.html: uma entrada por código, com a versão vigente.
LISTA_ARQUIVO = "lista-mestra.json"


def lista_caminho(caminho):
    """Aceita a pasta dos documentos ou o próprio arquivo."""
    if os.path.isdir(caminho) or not caminho.lower().endswith(".json"):
        return os.path.join(caminho, LISTA_ARQUIVO)
    return caminho


def carregar_lista(caminho):
    alvo = lista_caminho(caminho)
    base = {"lista_mestra": ORG["organizacao"], "padrao": ORG["politica"], "atualizado_em": "", "documentos": []}
    if os.path.exists(alvo):
        texto = open(alvo, encoding="utf8").read().strip()
        if texto:
            dados = json.loads(texto)
            if isinstance(dados, dict):
                base.update(dados)
    if not isinstance(base.get("documentos"), list):
        base["documentos"] = []
    return base


def sigla_da_pasta(nome):
    """Acha a sigla do tipo no nome da pasta: 'POP', 'Políticas (POL)', 'REG - regulamentos'."""
    for palavra in re.findall(r"[A-Z]{2,4}", nome.upper()):
        if palavra in TIPOS_CURTOS or palavra in TIPOS_LONGOS:
            return palavra
    return None


def chave_registro(reg):
    """Documento com número é único pelo código; sem número (XXX), cada arquivo vale como uma entrada."""
    if reg.get("numero"):
        return reg["codigo"]
    return f'{reg["codigo"]}|{reg.get("arquivo", "")}'


RE_VERSAO_OK = re.compile(r"^\d+\.\d+(-r\d+)?$")


def _sem_acento(t):
    t = unicodedata.normalize("NFD", str(t or "").lower())
    return re.sub(r"[^a-z0-9]+", " ", "".join(c for c in t if unicodedata.category(c) != "Mn")).strip()


def codigo_de(spec):
    """Código que validar() vai montar, calculado antes de gerar para conferir a lista."""
    if spec.get("codigo"):
        return spec["codigo"]
    tipo, area = spec.get("tipo", "").upper(), spec.get("area", "ADM").upper()
    num = spec.get("numero")
    n = f"{int(num):03d}" if num else "XXX"
    ano = spec.get("ano") or (spec.get("data_emissao", "")[-4:] if re.match(r"\d{2}/\d{2}/\d{4}$", str(spec.get("data_emissao", ""))) else datetime.date.today().year)
    return f"{spec.get('empresa') or ORG['padrao']}-{tipo}-{area}-{n}" + (f"/{ano}" if tipo in TIPOS_CURTOS else "")


def _mesma_familia(d, spec, ano):
    return (d.get("empresa") == spec.get("empresa") and d.get("tipo") == spec.get("tipo")
            and d.get("area") == spec.get("area")
            and (spec.get("tipo") not in TIPOS_CURTOS or str(d.get("ano") or "") == str(ano or "")))


def proximo_numero(lista, spec):
    ano = spec.get("ano") or (spec.get("data_emissao", "")[-4:] if spec.get("data_emissao") else datetime.date.today().year)
    usados = [int(d.get("numero") or 0) for d in lista["documentos"] if _mesma_familia(d, spec, ano)]
    return max(usados) + 1 if usados else 1


OUTROS_NOMES = {"POP": ["Procedimento operacional", "Procedimento"], "CI": ["Comunicação"], "RO": ["Registro de ocorrência"],
                "PRO": ["Proposta"], "TER": ["Termo"]}


def titulo_sem_tipo(titulo, tipo):
    """Regra de título da política de documentação: o título diz só o assunto; o tipo já sai impresso acima dele.
    Tira do começo do título o nome do tipo, o conectivo e um separador que venham em seguida."""
    t = str(titulo or "").strip()
    nomes = sorted([n for n in [TIPOS_CURTOS.get(tipo), TIPOS_LONGOS.get(tipo)] + OUTROS_NOMES.get(tipo, []) if n], key=len, reverse=True)
    palavras = t.split()
    for nome in nomes:
        n = len(nome.split())
        if len(palavras) <= n or _sem_acento(" ".join(palavras[:n])) != _sem_acento(nome):
            continue
        resto = " ".join(palavras[n:])
        resto = re.sub(r"^[–—:\-]+\s*", "", resto)
        resto = re.sub(r"^(de|da|do|das|dos|para|sobre)\s+", "", resto, flags=re.I)
        resto = re.sub(r"^[–—:\-]+\s*", "", resto).strip()
        return resto[:1].upper() + resto[1:] if len(resto) >= 3 else t
    return t


def conferir_lista(lista, spec, atualizacao=False):
    """Devolve (erros, avisos) comparando o documento com o que já está registrado."""
    erros, avisos = [], []
    codigo = codigo_de(spec)
    if not spec.get("numero"):
        avisos.append("Documento sem número: a lista mestra registra, mas reserve o número antes de aprovar.")
        return erros, avisos
    mesmo = next((d for d in lista["documentos"] if d.get("codigo") == codigo and d.get("numero")), None)
    if mesmo and mesmo.get("confidencial"):
        # a lista não guarda o título do confidencial: o código só serve para a nova versão do próprio documento
        if not atualizacao:
            erros.append(f'O código {codigo} já é de um documento confidencial na lista mestra ({mesmo.get("estado")}, '
                         f'v{mesmo.get("versao")}). Use o número {proximo_numero(lista, spec):03d} ou gere a nova versão a partir daquele documento.')
    elif mesmo and atualizacao and _sem_acento(mesmo.get("titulo")) != _sem_acento(spec.get("titulo")):
        avisos.append(f'O título muda nesta versão: "{mesmo.get("titulo")}" passa a "{spec.get("titulo")}". A lista mestra fica com o novo.')
    elif mesmo and _sem_acento(mesmo.get("titulo")) != _sem_acento(spec.get("titulo")):
        erros.append(f'O código {codigo} já é "{mesmo.get("titulo")}" na lista mestra '
                     f'({mesmo.get("estado")}, v{mesmo.get("versao")}). Use o número {proximo_numero(lista, spec):03d} '
                     f'ou gere a nova versão a partir daquele documento.')
    elif mesmo and not atualizacao:
        avisos.append(f'{codigo} já estava na lista mestra (v{mesmo.get("versao")}, {mesmo.get("estado")}): a entrada foi atualizada.')
    homonimo = next((d for d in lista["documentos"] if d.get("codigo") != codigo and not d.get("confidencial")
                     and _sem_acento(d.get("titulo")) == _sem_acento(spec.get("titulo"))), None)
    if homonimo:
        avisos.append(f'Já existe {homonimo.get("codigo")} com o mesmo título ("{homonimo.get("titulo")}"). Confirme se não é o mesmo documento.')
    return erros, avisos


def _mesmo_codigo(substituido_por, spec):
    """'ABC-CI-ADM-002/2026 v1.2' substitui o mesmo código do documento?"""
    alvo = str(substituido_por or "").split(" v")[0].strip()
    codigo = str(spec.get("codigo") or "").strip()
    if not codigo and str(spec.get("numero") or "").isdigit():
        tipo = str(spec.get("tipo", "")).upper()
        codigo = f'{spec.get("empresa") or ORG["padrao"]}-{tipo}-{str(spec.get("area", "ADM")).upper()}-{int(spec["numero"]):03d}' + \
                 (f'/{spec.get("ano")}' if tipo in TIPOS_CURTOS and spec.get("ano") else "")
    return bool(alvo) and alvo == codigo


# Documento confidencial mora numa pasta chamada Confidencial, com permissão só para quem pode ler. A lista mestra fica na
# pasta comum, aberta a todos: nela o confidencial é só um esqueleto (código, versão, estado e data), para o número não
# repetir. Título, responsável, projeto e nome do arquivo não vão para a lista. O inventário da pasta comum pula a pasta
# Confidencial: quem não tem acesso a ela não trava, e quem tem não copia nada dela para a lista.
TITULO_CONFIDENCIAL = "Documento confidencial"
CAMPOS_FORA_DO_CONFIDENCIAL = ("responsavel", "arquivo", "projeto", "substituido_por", "plano", "encerrado_por", "rodadas_sha256")


def eh_pasta_confidencial(nome):
    return _sem_acento(nome) == "confidencial"


def registro_de(spec, arquivo):
    if eh_confidencial(spec):
        return {"codigo": spec["codigo"], "titulo": TITULO_CONFIDENCIAL, "empresa": spec.get("empresa"),
                "tipo": spec.get("tipo"), "area": spec.get("area"),
                "numero": int(spec["numero"]) if str(spec.get("numero") or "").isdigit() else None,
                "ano": spec.get("ano"), "versao": spec.get("versao"), "estado": spec.get("estado"),
                "data": spec.get("data_aprovacao") or spec.get("data_emissao"), "confidencial": True}
    reg = {"codigo": spec["codigo"], "titulo": spec.get("titulo", ""), "empresa": spec.get("empresa"),
           "tipo": spec.get("tipo"), "area": spec.get("area"),
           "numero": int(spec["numero"]) if str(spec.get("numero") or "").isdigit() else None,
           "ano": spec.get("ano"), "versao": spec.get("versao"), "estado": spec.get("estado"),
           "data": spec.get("data_aprovacao") or spec.get("data_emissao"),
           "responsavel": spec.get("elaborado_por", ""), "arquivo": os.path.basename(arquivo)}
    if spec.get("substituido_por") and spec.get("estado") == "OBSOLETO":   # só vale para obsoleto
        reg["substituido_por"] = spec["substituido_por"]
    if spec.get("projeto"):
        reg["projeto"] = spec["projeto"]
    if spec.get("tipo") == "REL" and spec.get("rodadas"):   # relatório de condução: plano e impressão digital das rodadas
        reg["plano"] = spec.get("plano")
        reg["rodadas_sha256"] = digitais_rodadas(spec["rodadas"])
        fim = encerramento_rel(spec["rodadas"])
        if fim:
            reg["encerrado_por"] = fim.get("termo")
    return reg


def gravar_lista(caminho, registro):
    """Relê o arquivo antes de gravar: outra pessoa pode ter registrado algo nesse meio-tempo."""
    alvo = lista_caminho(caminho)
    lista = _mesclar_em(carregar_lista(alvo), registro)
    pasta = os.path.dirname(os.path.abspath(alvo))
    os.makedirs(pasta, exist_ok=True)
    open(alvo, "w", encoding="utf8").write(json.dumps(lista, ensure_ascii=False, indent=2) + "\n")
    return lista


# ----------------------------------------------------------------------------- relatório de condução (REL)
# Relatório de condução de projeto. Um relatório por projeto, com o número do plano (herdado, não
# reservado). Só recebe linhas novas: cada rodada é uma versão oficial nova (1.1, 1.2…), publicada pela coordenação
# sem passar por em aprovação e sem cópia obsoleta. A lista mestra guarda o SHA-256 de cada rodada publicada, e a
# rodada nova é recusada se uma linha antiga mudou. As rodadas ficam na ficha ("rodadas"); o Quadro de rodadas do
# corpo é refeito a partir delas (e criado no fim do corpo, se faltar). Mesmas regras na página.
EVENTOS_REL = {"abertura": "", "rodada": "", "replanejamento": "Replanejamento", "parada": "Parada declarada",
               "retomada": "Retomada", "correcao": "Correção", "encerramento": "Encerramento"}
CAB_RODADAS = ["Nº", "Data", "O que se fez", "Compromisso seguinte (o quê · quem · até)", "Avanço"]
NOTA_RODADAS = "Só se acrescentam linhas. Rodada publicada não se reescreve; correção entra como rodada nova, citando a corrigida."
SITUACOES_REL = ("concluído", "cancelado")
SEM_AVANCO_PARADA = 3   # rodadas seguidas sem avanço que permitem declarar a parada do projeto


def _limpo(t):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", "" if t is None else str(t))).strip()


def celulas_rodada(r):
    """As cinco células da linha no Quadro de rodadas. É este texto que a impressão digital protege."""
    tipo = r.get("tipo") or "rodada"
    if tipo == "replanejamento":
        pre = f'Replanejamento: plano passa à v{re.sub(r"^[vV]", "", _limpo(r.get("plano_versao")))}.'
    elif tipo == "correcao":
        pre = f'Correção da rodada {_limpo(r.get("corrige"))}:'
    elif tipo == "encerramento":
        pre = f'Encerramento ({_limpo(r.get("situacao"))}), com o termo {_limpo(r.get("termo"))}.'
    else:
        pre = EVENTOS_REL[tipo] + ":" if EVENTOS_REL.get(tipo) else ""
    c = r.get("compromisso") or {}
    comp = " · ".join(p for p in (_limpo(c.get("oque")), _limpo(c.get("quem")),
                                  "até " + _limpo(c.get("ate")) if _limpo(c.get("ate")) else "") if p) or "—"
    if r.get("avanco") == "sim":
        avanco = "Sim"
    elif r.get("avanco") == "nao":
        avanco = "Não: " + _limpo(r.get("travou")) + (f'; depende de {_limpo(r.get("depende"))}' if _limpo(r.get("depende")) else "")
    else:
        avanco = "—"
    return [str(int(r.get("n"))), _limpo(r.get("data")), _limpo(f'{pre} {_limpo(r.get("feito"))}'), comp, _limpo(avanco)]


def digital_rodada(r):
    return hashlib.sha256("\n".join(celulas_rodada(r)).encode("utf8")).hexdigest()


def digitais_rodadas(rodadas):
    return [digital_rodada(r) for r in rodadas or []]


def encerramento_rel(rodadas):
    return next((r for r in rodadas or [] if r.get("tipo") == "encerramento"), None)


def sem_avanco(rodadas):
    """Números das últimas rodadas seguidas sem avanço. Correção não conta nem interrompe; os demais eventos interrompem."""
    seguidas = []
    for r in reversed(rodadas or []):
        tipo = r.get("tipo") or "rodada"
        if tipo == "correcao":
            continue
        if tipo != "rodada" or r.get("avanco") != "nao":
            break
        seguidas.insert(0, int(r.get("n")))
    return seguidas


def versao_rodada(v):
    """Versão da rodada seguinte: muda o Y (1.0 → 1.1 → 1.2…)."""
    x, y = (str(v or "1.0").split("-r")[0].split(".") + ["0"])[:2]
    return f"{int(x)}.{int(y) + 1}"


def _data_ord(d):
    """dd/mm/aaaa → aaaammdd, para comparar; vazio se não for uma data válida."""
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", _limpo(d))
    try:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).strftime("%Y%m%d") if m else ""
    except ValueError:
        return ""


def resumo_rodada(r):
    tipo = r.get("tipo") or "rodada"
    txt = {"abertura": "abertura do relatório e início da execução", "rodada": "registro da reunião",
           "replanejamento": f'replanejamento (plano v{re.sub(r"^[vV]", "", _limpo(r.get("plano_versao")))})',
           "parada": "parada declarada", "retomada": "retomada",
           "correcao": f'correção da rodada {_limpo(r.get("corrige"))}',
           "encerramento": f'encerramento ({_limpo(r.get("situacao"))}), termo {_limpo(r.get("termo"))}'}.get(tipo, tipo)
    return f'Rodada {r.get("n")}: {txt}.'


def conferir_rodada(nova, anteriores):
    """Erros da rodada nova, conferida contra as já publicadas."""
    erros = []
    fim = encerramento_rel(anteriores)
    if fim:
        erros.append(f'O relatório foi encerrado na rodada {fim.get("n")} (termo {fim.get("termo")}): não recebe novas rodadas.')
    tipo = nova.get("tipo") or "rodada"
    esperado = int(anteriores[-1]["n"]) + 1 if anteriores else 0
    if tipo not in EVENTOS_REL:
        erros.append(f"Tipo de registro desconhecido: {tipo}.")
    if not re.match(r"^\d+$", _limpo(nova.get("n"))) or int(_limpo(nova.get("n"))) != esperado:
        erros.append(f"A rodada nova é a {esperado}" + (f' (a última publicada é a {anteriores[-1]["n"]}).' if anteriores else ", a abertura."))
    elif (tipo == "abertura") != (esperado == 0):
        erros.append("A rodada 0 é a abertura, com o início da execução." if esperado == 0 else "Só a rodada 0 é a abertura.")
    data = _limpo(nova.get("data"))
    if not _data_ord(data):
        erros.append("Informe a data da rodada (dd/mm/aaaa).")
    elif anteriores and data and _data_ord(data) < _data_ord(anteriores[-1].get("data")):
        erros.append(f'A data da rodada ({data}) é anterior à da rodada {anteriores[-1]["n"]} ({anteriores[-1].get("data")}).')
    if not _limpo(nova.get("feito")) and tipo != "encerramento":
        erros.append("Descreva o que se fez desde a rodada anterior.")
    c = nova.get("compromisso") or {}
    if tipo in ("abertura", "rodada", "replanejamento", "retomada") and not all(_limpo(c.get(k)) for k in ("oque", "quem", "ate")):
        erros.append("Informe o compromisso seguinte: o quê, quem e até quando.")
    if tipo == "rodada":
        if nova.get("avanco") not in ("sim", "nao"):
            erros.append("Diga se houve avanço: sim, ou não com o que travou e de quem depende.")
        elif nova["avanco"] == "nao" and not (_limpo(nova.get("travou")) and _limpo(nova.get("depende"))):
            erros.append("Sem avanço: diga o que travou e de quem depende.")
    if tipo == "correcao" and not any(_limpo(r.get("n")) == _limpo(nova.get("corrige")) for r in anteriores):
        erros.append("Correção: informe o número da rodada corrigida.")
    if tipo == "replanejamento" and not re.match(r"^[vV]?\d+\.\d+$", _limpo(nova.get("plano_versao"))):
        erros.append("Replanejamento: informe a nova versão do plano (ex.: 1.1).")
    if tipo == "encerramento":
        if "-TER-" not in _limpo(nova.get("termo")).upper():
            erros.append("Encerramento: informe o código do termo de encerramento (TER).")
        if _limpo(nova.get("situacao")) not in SITUACOES_REL:
            erros.append("Encerramento: diga se o projeto foi concluído ou cancelado.")
    return erros


def _eh_quadro_rodadas(b):
    return isinstance(b, dict) and b.get("t") == "tabela" and \
        [str(c).replace("*", "").strip() for c in (b.get("cabecalho") or [])][:3] == CAB_RODADAS[:3]


def conferir_quadro_rodadas(corpo, rodadas):
    """O Quadro de rodadas do corpo ainda diz o que a ficha registrou? Linha mudada no Word é recusada."""
    q = next((b for b in corpo or [] if _eh_quadro_rodadas(b)), None)
    if q is None:
        return []
    lidas = [[_limpo(str(c).replace("*", "")) for c in l] for l in q.get("linhas") or []]
    for r, l in zip(rodadas, lidas):
        if l[:5] != celulas_rodada(r):
            return [f'A rodada {r.get("n")} foi alterada no Quadro de rodadas do documento. Rodada publicada não se reescreve: '
                    "volte o texto ao publicado (ou parta da versão vigente) e registre a correção numa rodada nova."]
    if len(lidas) != len(rodadas):
        return [f"O Quadro de rodadas tem {len(lidas)} linha(s), mas a ficha registra {len(rodadas)}: linha publicada não se acrescenta nem se apaga à mão."]
    return []


def plano_do_projeto(lista, projeto):
    return [d for d in lista["documentos"] if d.get("tipo") == "PLN" and d.get("estado") == "OFICIAL" and projeto
            and _sem_acento(d.get("projeto") or projeto_da_pasta(d.get("pasta"))) == _sem_acento(projeto)]


def regras_rel(spec, lista, anterior=None):
    """Regras próprias do relatório de condução. Acerta código, versão e histórico; devolve (erros, avisos)."""
    erros, avisos = [], []
    if spec.get("estado") != "OFICIAL":
        erros.append("Relatório de condução é oficial pelo ato da coordenação: gere como OFICIAL (não tem minuta nem em aprovação).")
    if not spec.get("projeto"):
        erros.append("Relatório de condução é sempre de um projeto: informe o projeto.")
    if lista is None:
        erros.append("Relatório de condução precisa da lista mestra (--lista): é nela que se confere o plano oficial e as rodadas publicadas.")
    if erros:
        return erros, avisos
    docs = lista["documentos"]
    plano = (anterior or {}).get("plano") or spec.get("plano")
    if not plano:
        planos = plano_do_projeto(lista, spec["projeto"])
        if not planos:
            return [f'O projeto "{spec["projeto"]}" não tem plano (PLN) oficial na lista mestra: o relatório de condução só se abre com o plano já oficial.'], avisos
        if len(planos) > 1:
            return [f'O projeto "{spec["projeto"]}" tem mais de um plano oficial na lista ({", ".join(d["codigo"] for d in planos)}): informe "plano".'], avisos
        plano = planos[0]["codigo"]
    m = re.match(r"^([A-Z0-9]{2,6})-PLN-([A-Z0-9]{2,6})-(\d{3})$", str(plano))
    reg_plano = next((d for d in docs if d.get("codigo") == plano), None)
    if not m:
        return [f"Plano inválido: {plano}."], avisos
    if not reg_plano or reg_plano.get("estado") != "OFICIAL":
        erros.append(f"O plano {plano} não está oficial na lista mestra: o relatório de condução só se abre com o plano já oficial.")
    codigo = f"{m.group(1)}-REL-{m.group(2)}-{m.group(3)}"
    if spec.get("codigo") and spec["codigo"] != codigo:
        erros.append(f'O relatório do plano {plano} é o {codigo}, e não {spec["codigo"]}: o número é herdado do plano.')
    if str(spec.get("numero") or "").isdigit() and int(spec["numero"]) != int(m.group(3)) or spec.get("area", m.group(2)) != m.group(2):
        avisos.append(f"Número herdado do plano: o relatório é o {codigo}.")
    spec.update({"plano": plano, "empresa": m.group(1), "area": m.group(2), "numero": int(m.group(3))})
    spec.pop("area_nome", None)
    spec["codigo"] = codigo
    reg = next((d for d in docs if d.get("codigo") == codigo), None)
    if reg and reg.get("plano") != plano:
        erros.append(f'O código {codigo} já é "{reg.get("titulo")}" na lista mestra (v{reg.get("versao")}), que não é o relatório de condução '
                     f"do plano {plano}. A política reserva esse número ao relatório do plano: resolva o conflito antes de abrir o relatório.")
    rodadas = list(spec.get("rodadas") or [])
    ants = list((anterior or {}).get("rodadas") or [])
    if not rodadas:
        return erros + ["Relatório de condução sem rodadas: a abertura registra a rodada 0, com o início da execução."], avisos
    if anterior is None:
        if reg and reg.get("plano") == plano:
            erros.append(f'O relatório {codigo} já existe (v{reg.get("versao")}): acrescente a rodada a partir da versão vigente.')
        if len(rodadas) != 1:
            erros.append("A abertura do relatório registra só a rodada 0.")
        spec["versao"] = "1.0"
        spec["historico"] = [{"versao": "1.0", "data": spec.get("data_aprovacao") or rodadas[0].get("data", ""),
                              "descricao": "Abertura do relatório: identificação e rodada 0 (início da execução).",
                              "responsavel": spec.get("aprovado_por", "")}]
    else:
        if len(rodadas) != len(ants) + 1:
            erros.append("Cada versão do relatório acrescenta uma rodada, e só uma.")
        for a, r in zip(ants, rodadas):
            if digital_rodada(a) != digital_rodada(r):
                erros.append(f'A rodada {a.get("n")} mudou: rodada publicada não se reescreve. Corrija numa rodada nova, do tipo correção.')
                break
        publicadas = (reg or {}).get("rodadas_sha256") or []
        if len(publicadas) > len(ants):
            erros.append(f'O arquivo é a v{anterior.get("versao")}, com {len(ants)} rodada(s), mas a lista mestra já registra {len(publicadas)} '
                         f'(v{reg.get("versao")}). Acrescente a rodada a partir da versão vigente.')
        else:
            for i, (h, a) in enumerate(zip(publicadas, ants)):
                if h != digital_rodada(a):
                    erros.append(f'A rodada {a.get("n")} do documento não confere com a registrada na lista mestra (SHA-256 diferente): '
                                 "rodada publicada não se reescreve. Parta da versão vigente e corrija numa rodada nova, do tipo correção.")
                    break
            if reg and len(publicadas) < len(ants) and not erros:
                avisos.append(f"A lista mestra registrava {len(publicadas)} de {len(ants)} rodada(s); as demais foram conferidas com o documento.")
        spec["versao"] = versao_rodada(anterior.get("versao"))
    if rodadas:
        erros += conferir_rodada(rodadas[-1], rodadas[:-1])
    spec["data_aprovacao"] = spec.get("data_aprovacao") or rodadas[-1].get("data")
    seguidas = sem_avanco(rodadas)
    if len(seguidas) >= SEM_AVANCO_PARADA:
        avisos.append(f'{len(seguidas)} rodadas seguidas sem avanço ({", ".join(map(str, seguidas))}): a parada pode ser declarada '
                      "do projeto. Declarar é decisão da coordenação, registrada como rodada.")
    fim = encerramento_rel(rodadas)
    if fim:
        if not next((d for d in docs if d.get("codigo") == fim.get("termo")), None):
            avisos.append(f'O termo de encerramento {fim.get("termo")} não está na lista mestra: confira o código.')
        avisos.append(f'Relatório encerrado: não recebe novas rodadas. O termo de encerramento cita o {codigo} v{spec["versao"]}.')
    return erros, avisos


def montar_rel(spec):
    """Refaz no corpo o Quadro de rodadas (a partir da ficha) e o quadro de encerramento."""
    corpo = spec.get("corpo") or []
    i = next((k for k, b in enumerate(corpo) if _eh_quadro_rodadas(b)), None)
    if i is None:   # corpo sem o quadro (texto próprio, sem partir do exemplo): a seção Rodadas entra no fim
        corpo += [{"t": "h1", "texto": "Rodadas"}, {"t": "tabela", "tipo": "Quadro", "titulo": "Rodadas do projeto", "cabecalho": list(CAB_RODADAS)}]
        i = len(corpo) - 1
    rodadas = spec.get("rodadas") or []
    corpo[i] = dict(corpo[i], cabecalho=list(CAB_RODADAS), linhas=[celulas_rodada(r) for r in rodadas],
                    larguras=[0.5, 1.8, 4.0, 3.1, 2.3], titulo=corpo[i].get("titulo") or "Rodadas do projeto", nota=NOTA_RODADAS)
    fim = encerramento_rel(rodadas)
    if fim:   # a seção de encerramento deixa de ter o lembrete de preenchimento
        corpo[:] = [b for b in corpo if not (b.get("t") == "p" and re.match(r"^\[?Preenchido na última rodada", str(b.get("texto", "")).strip()))]
    for b in corpo:
        if b.get("t") == "campos" and b.get("itens") and str(b["itens"][0][0]).replace("*", "").strip() == "Termo de encerramento":
            b["itens"] = [["Termo de encerramento", fim.get("termo") if fim else "—"],
                          ["Última rodada", f'{fim["n"]}, em {fim.get("data")}' if fim else "—"],
                          ["Situação", _limpo(fim.get("situacao")).capitalize() if fim else "Em andamento"]]
        elif b.get("t") == "p" and str(b.get("texto", "")).strip().startswith("[Preenchido na última rodada"):
            b["texto"] = "Preenchido na última rodada."
    spec["corpo"] = corpo


def preparar_rodada(anterior, rodada, responsavel=None):
    """Especificação da versão seguinte do relatório de condução: as rodadas publicadas mais a nova."""
    if anterior.get("tipo") != "REL":
        raise SystemExit("--rodada é só para o relatório de condução (REL).")
    if anterior.get("estado") != "OFICIAL":
        raise SystemExit(f'O relatório recebido está como {anterior.get("estado")}: acrescente a rodada a partir da versão oficial vigente.')
    ants = list(anterior.get("rodadas") or [])
    if not ants:
        raise SystemExit("Este REL não tem rodadas na ficha: não é um relatório de condução aberto pelo Timbrador.")
    nova = json.loads(json.dumps(anterior))
    for k in ("codigo", "vigencia", "data_aprovacao", "substitui", "substituido_por", "motivo_obsoleto"):
        nova.pop(k, None)
    r = dict(rodada)
    r.setdefault("tipo", "rodada")
    r.setdefault("n", int(ants[-1]["n"]) + 1)
    r.setdefault("data", _hoje())
    nova["rodadas"] = ants + [r]
    nova["versao"] = versao_rodada(anterior.get("versao"))
    nova["estado"] = "OFICIAL"
    nova["data_emissao"] = nova["data_aprovacao"] = r["data"]
    if responsavel:
        nova["aprovado_por"] = responsavel
    hist = list(anterior.get("historico") or [])
    hist.append({"versao": nova["versao"], "data": r["data"], "descricao": resumo_rodada(r), "responsavel": nova.get("aprovado_por", "")})
    nova["historico"] = hist
    return nova


# ----------------------------------------------------------------------------- inventário da pasta
RE_NOME = re.compile(r"^([A-Z0-9]{2,6})-([A-Z]{2,4})-([A-Z]{2,4})-(\d{3}|XXX)(?:-(\d{4}))?"
                     r"_v([0-9][0-9.]*(?:-r\d+)?)_([A-Z_]+)_(.+)\.docx$", re.I)
ORDEM_ESTADO = {"OBSOLETO": 0, "RASCUNHO": 1, "MINUTA": 2, "EM_REVISAO": 3, "EM_APROVACAO": 3, "OFICIAL": 4}


def _chave_versao(v):
    """Ordena versões: 1.0-r2 < 1.0 < 1.1-r1 < 2.0."""
    base, _, rod = str(v or "0").partition("-r")
    partes = [int(p) if p.isdigit() else 0 for p in base.split(".")][:2]
    while len(partes) < 2:
        partes.append(0)
    return (partes[0], partes[1], 0 if rod else 1, int(rod) if rod.isdigit() else 0)


def ficha_rapida(caminho):
    """Lê só a identificação de um .docx (ficha embutida ou cabeçalho), sem percorrer o corpo."""
    spec, origem = {}, "sem_timbrado"
    with zipfile.ZipFile(caminho) as z:
        nomes = z.namelist()
        for n in nomes:
            if re.match(r"customXml/item\d+\.xml$", n):
                x = z.read(n).decode("utf8")
                if MARCA_FICHA in x:
                    try:
                        spec = json.loads(ET.fromstring(x).text or "{}")
                        origem = "ficha"
                    except (ET.ParseError, ValueError):
                        pass
                    break
        if origem != "ficha":
            cab = " ".join(re.sub(r"<[^>]+>", " ", z.read(n).decode("utf8"))
                           for n in nomes if re.match(r"word/(header|footer)\d*\.xml$", n))
            cab = re.sub(r"\s+", " ", cab)
            m = RE_CODIGO.search(cab)
            if m:
                origem = "cabecalho"
                spec.update({"empresa": m.group(1), "tipo": m.group(2), "area": m.group(3), "versao": m.group(6)})
                if m.group(4) != "XXX":
                    spec["numero"] = int(m.group(4))
                if m.group(5):
                    spec["ano"] = int(m.group(5))
                for chave, est in ESTADOS.items():
                    if est["tarja"].split("  ·  ")[0] in cab:
                        spec["estado"] = chave
                if "DOCUMENTO OFICIAL" in cab:
                    spec["estado"] = "OFICIAL"
                if re.search(r"\bCONFIDENCIAL\b", cab):
                    spec["classificacao"] = "Confidencial"
                mo = re.search(r"Aprovado por (.+?) em (\d{2}/\d{2}/\d{4})", cab)
                if mo:
                    spec["aprovado_por"], spec["data_aprovacao"] = mo.group(1).strip(), mo.group(2)
    ma = re.search(r"/(\d{4})$", str(spec.get("codigo", "")))
    if ma and not spec.get("ano"):
        spec["ano"] = int(ma.group(1))
    return spec, origem


def ficha_do_nome(nome):
    """Último recurso: o próprio nome do arquivo segue CÓDIGO_vVERSÃO_ESTADO_titulo.docx."""
    m = RE_NOME.match(os.path.basename(nome))
    if not m:
        return {}, "sem_timbrado"
    emp, tipo, area, num, ano, versao, estado, slug = m.groups()
    spec = {"empresa": emp.upper(), "tipo": tipo.upper(), "area": area.upper(), "versao": versao,
            "estado": estado.upper(), "titulo": slug.replace("-", " ").strip().capitalize()}
    if num.upper() != "XXX":
        spec["numero"] = int(num)
    if ano:
        spec["ano"] = int(ano)
    return spec, "nome"


def achar_lista(pasta):
    """Procura um lista-mestra.json já existente na árvore, para não criar uma segunda lista."""
    return sorted(os.path.join(r, n) for r, _d, arqs in os.walk(pasta) for n in arqs if n == LISTA_ARQUIVO)


# Pastas puladas por padrão na varredura: o arquivo morto não faz parte do que está em uso.
# Para inventariar o próprio arquivo morto, aponte --inventariar para ele: a raiz nunca é pulada.
# Modelos (arquivos de partida e saídas antigas de conversas) fica fora da varredura. OBSOLETOS é lida, mas só decide
# a entrada de um código que não tem mais nenhuma versão em uso. Mesmas regras do inventário da página.
PASTAS_IGNORADAS = ("modelos",)


def projeto_da_pasta(pasta):
    """Projeto de uma pasta registrada: "Projetos/Novo Site/…" → "Novo Site"."""
    partes = str(pasta or "").replace("\\", "/").split("/")
    for i, x in enumerate(partes[:-1]):
        if _sem_acento(x) == "projetos" and partes[i + 1]:
            return partes[i + 1]
    return ""


def _eh_pasta_obsoletos(rel):
    return any(p.lower() == "obsoletos" for p in str(rel or "").replace("\\", "/").split("/"))


def inventariar(pasta, destino=None, gravar=True, ignorar=()):
    """Varre a pasta dos documentos e monta a lista mestra do que já existe."""
    pasta = os.path.abspath(pasta)
    ignorar = [i.lower() for i in (ignorar or ())] + [i for i in PASTAS_IGNORADAS]
    achados, ignorados, sem_id, listas, modelos, visitadas, puladas = [], [], [], [], [], set(), []
    for raiz, dirs, arquivos in os.walk(pasta):
        rel_r = os.path.relpath(raiz, pasta).replace(os.sep, "/")
        puladas += [d if rel_r == "." else f"{rel_r}/{d}" for d in dirs if eh_pasta_confidencial(d)]   # fora da varredura
        dirs[:] = sorted(d for d in dirs if d.lower() not in ignorar and not d.startswith(".") and not eh_pasta_confidencial(d))
        rel_raiz = os.path.relpath(raiz, pasta).replace(os.sep, "/")
        visitadas.add("" if rel_raiz == "." else rel_raiz)
        for nome in sorted(arquivos):
            if nome == LISTA_ARQUIVO:
                listas.append(os.path.join(raiz, nome))
            if nome.startswith("~$") or nome.startswith("."):
                continue
            caminho = os.path.join(raiz, nome)
            rel = os.path.relpath(caminho, pasta)
            if not nome.lower().endswith(".docx"):
                if not nome.lower().endswith(".json"):
                    ignorados.append(rel)
                continue
            try:
                spec, origem = ficha_rapida(caminho)
            except (zipfile.BadZipFile, KeyError, OSError) as e:
                sem_id.append((rel, f"não foi possível abrir ({e.__class__.__name__})"))
                continue
            if spec.get("modelo"):          # arquivo de partida: não é documento, não entra na lista
                modelos.append(rel)
                continue
            pelo_nome, _ = ficha_do_nome(nome)
            if origem == "sem_timbrado":
                spec, origem = dict(pelo_nome), ("nome" if pelo_nome else "sem_timbrado")
            else:
                for k, v in pelo_nome.items():      # completa o que a ficha ou o cabeçalho não trazem (título, sobretudo)
                    spec.setdefault(k, v)
            if origem == "sem_timbrado" or not spec.get("tipo"):
                sem_id.append((rel, "sem ficha, sem código no cabeçalho e fora do padrão de nome"))
                continue
            sigla = sigla_da_pasta(os.path.basename(raiz))
            if sigla:
                spec.setdefault("tipo", sigla)
            spec.setdefault("empresa", ORG["padrao"])
            spec.setdefault("estado", "MINUTA")
            spec.setdefault("versao", "1.0")
            spec.setdefault("titulo", os.path.splitext(nome)[0])
            spec["codigo"] = codigo_de(spec)
            pasta_rel = os.path.dirname(rel).replace(os.sep, "/") or "."
            achados.append({"spec": spec, "origem": origem, "rel": rel, "nome": nome, "pasta": pasta_rel,
                            "obsoleta": _eh_pasta_obsoletos(pasta_rel) or spec.get("estado") == "OBSOLETO"})

    # uma entrada por código: fica a versão mais alta; as demais viram histórico de arquivos
    if not destino:
        if len(listas) == 1:
            destino = listas[0]
        elif len(listas) > 1:
            raise SystemExit("Há mais de um " + LISTA_ARQUIVO + " nesta árvore:\n  " +
                             "\n  ".join(listas) + "\nUse --lista para dizer qual vale.")
    por_codigo = {}
    for a in achados:
        chave = a["spec"]["codigo"] if a["spec"].get("numero") else f'{a["spec"]["codigo"]}|{a["nome"]}'
        por_codigo.setdefault(chave, []).append(a)
    lista = carregar_lista(destino or pasta)
    conflitos, multiplos, alertas, vistos = [], [], [], set()
    versao_base = lambda v: str(v or "").split("-r")[0]
    for chave_grupo, itens in sorted(por_codigo.items()):
        codigo = chave_grupo.split("|")[0]
        # empate de versão e estado (cópias do mesmo arquivo): vale a da pasta do projeto, depois a ordem das pastas
        itens.sort(key=lambda x: (_chave_versao(x["spec"].get("versao")),
                                  ORDEM_ESTADO.get(x["spec"].get("estado"), 0),
                                  1 if projeto_da_pasta(x["pasta"]) else 0, x["pasta"]))
        em_uso = [i for i in itens if not i["obsoleta"]]
        oficiais = [i for i in em_uso if i["spec"].get("estado") == "OFICIAL"]
        copias = [i for i in itens if i["obsoleta"]]
        atual = (oficiais or em_uso or copias)[-1]
        # oficial cuja versão já tem cópia obsoleta: foi substituída e o arquivo ficou para trás na pasta
        ja = next((c for c in copias if atual["spec"].get("estado") == "OFICIAL" and
                   _chave_versao(c["spec"].get("versao")) >= _chave_versao(atual["spec"].get("versao"))), None)
        if ja:
            alertas.append(f'{codigo} v{atual["spec"].get("versao")}: o arquivo oficial ainda está em {atual["pasta"]}, mas essa '
                           f'versão já tem cópia obsoleta em {ja["pasta"]}. Retire o oficial da pasta.')
            atual = ja
        rel = atual["spec"].get("tipo") == "REL" and bool(atual["spec"].get("rodadas"))
        if len(oficiais) > 1 and rel:
            alertas.append(f'{codigo}: relatório de condução com mais de um arquivo em uso (' +
                           "; ".join(f'v{o["spec"].get("versao")} em {o["pasta"]}' for o in oficiais) +
                           f'). A v{oficiais[-1]["spec"].get("versao")} contém as rodadas das anteriores: retire as anteriores da pasta '
                           "(relatório de condução não tem cópia obsoleta).")
        elif len(oficiais) > 1:
            # a mesma versão em mais de uma pasta é cópia do mesmo documento, não versão anterior
            ult = oficiais[-1]["spec"].get("versao")
            iguais = [o for o in oficiais if _chave_versao(o["spec"].get("versao")) == _chave_versao(ult)]
            if len(iguais) < len(oficiais):
                alertas.append(f'{codigo}: mais de uma versão oficial em uso (' +
                               "; ".join(f'v{o["spec"].get("versao")} em {o["pasta"]}' for o in oficiais) +
                               "). A anterior precisa ser tornada obsoleta.")
            if len(iguais) > 1:
                alertas.append(f'{codigo} v{ult}: o mesmo documento oficial está em mais de uma pasta (' +
                               "; ".join(o["pasta"] for o in iguais) +
                               "). Deixe só uma cópia: documento de projeto fica na pasta do projeto.")
        if oficiais:
            ult = oficiais[-1]["spec"].get("versao")
            for m in em_uso:
                if m["spec"].get("estado") != "OFICIAL" and _chave_versao(versao_base(m["spec"].get("versao"))) <= _chave_versao(versao_base(ult)):
                    alertas.append(f'{codigo}: a versão {m["spec"].get("versao")} ({m["spec"].get("estado")}, em {m["pasta"]}) é de uma '
                                   f'versão já publicada (v{ult}) e pode sair da pasta.')
        for c in copias:
            if not _eh_pasta_obsoletos(c["pasta"]):
                alertas.append(f'{c["nome"]}: cópia obsoleta fora da pasta OBSOLETOS (está em {c["pasta"]}).')
        for m in em_uso:
            pj = m["spec"].get("projeto")
            if pj and _sem_acento(projeto_da_pasta(m["pasta"])) != _sem_acento(pj):
                alertas.append(f'{m["nome"]}: é do projeto "{pj}", mas está em {m["pasta"]}. Documento de projeto fica em Projetos/{pj}.')
        titulos = {_sem_acento(i["spec"].get("titulo")) for i in itens if i["spec"].get("titulo")}
        if len(titulos) > 1:
            conflitos.append((codigo, sorted({i["spec"].get("titulo", "") for i in itens})))
        if len(itens) > 1:
            multiplos.append((codigo, [f'{i["spec"].get("versao")} {i["spec"].get("estado")}' for i in itens]))
        spec = atual["spec"]
        for i in itens:
            if eh_confidencial(i["spec"]):
                alertas.append(f'{i["rel"]}: documento confidencial fora da pasta Confidencial (está em {i["pasta"]}). '
                               "Quem tem acesso à pasta comum pode lê-lo: mova-o para a pasta Confidencial.")
        reg = registro_de(spec, atual["nome"])
        reg["pasta"] = atual["pasta"]
        reg["origem"] = atual["origem"]
        if reg.get("confidencial") and not reg.get("numero"):
            continue   # confidencial sem número: não reserva nada, e o esqueleto não diria qual é
        if not reg.get("confidencial") and not reg.get("projeto") and projeto_da_pasta(atual["pasta"]):   # documento antigo, sem o campo na ficha
            reg["projeto"] = projeto_da_pasta(atual["pasta"])
        if len(itens) > 1:
            reg["arquivos"] = len(itens)
        if spec.get("tipo") == "REL":
            alertas += _alertas_rel(codigo, spec, em_uso, lista, reg)
        lista = _mesclar_em(lista, reg)
        vistos.add(chave_registro(reg))
    # entradas sem arquivo: saem quando a pasta registrada foi varrida (ou é a de Modelos); fora disso, ficam e são avisadas
    removidas, nao_conferidas, ficam, confidenciais = [], [], [], []
    for d in lista["documentos"]:
        if chave_registro(d) in vistos:
            ficam.append(d); continue
        if d.get("confidencial"):   # está na pasta Confidencial, que não é varrida
            confidenciais.append(d.get("codigo")); ficam.append(d); continue
        if "pasta" not in d:
            nao_conferidas.append(f'{d.get("codigo")} (sem pasta registrada)'); ficam.append(d); continue
        p = "" if d["pasta"] == "." else str(d["pasta"])
        if any(x.lower() in PASTAS_IGNORADAS for x in p.split("/")) or p in visitadas:
            removidas.append(f'{d.get("codigo")} v{d.get("versao")} ({d.get("arquivo") or p})')
        else:
            nao_conferidas.append(f'{d.get("codigo")} ({p})'); ficam.append(d)
    lista["documentos"] = ficam
    for d in lista["documentos"]:   # relatório de condução: o plano dele tem de estar oficial
        if d.get("tipo") == "REL" and d.get("plano"):
            p = next((x for x in lista["documentos"] if x.get("codigo") == d["plano"]), None)
            if not p or p.get("estado") != "OFICIAL":
                alertas.append(f'{d["codigo"]}: o plano {d["plano"]} não está oficial na lista mestra.')
    versoes_ruins = sorted({(a["spec"]["codigo"], str(a["spec"].get("versao"))) for a in achados
                            if not RE_VERSAO_OK.match(str(a["spec"].get("versao") or ""))})
    sem_numero = sorted({(a["spec"]["codigo"], a["nome"]) for a in achados if not a["spec"].get("numero")})
    relatorio = {"pasta": pasta, "arquivos_docx": len(achados), "documentos": len(por_codigo),
                 "versoes_fora_do_padrao": versoes_ruins, "sem_numero": sem_numero,
                 "ignorados": ignorados, "modelos": modelos, "sem_identificacao": sem_id, "conflitos": conflitos,
                 "varias_versoes": multiplos, "alertas": alertas, "removidas": removidas, "nao_conferidas": nao_conferidas,
                 "confidenciais": confidenciais, "pastas_confidenciais": puladas,
                 "por_origem": {o: sum(1 for a in achados if a["origem"] == o) for o in ("ficha", "cabecalho", "nome")},
                 "por_tipo": {}, "por_estado": {}, "listas_encontradas": listas}
    for a in achados:
        relatorio["por_tipo"][a["spec"]["tipo"]] = relatorio["por_tipo"].get(a["spec"]["tipo"], 0) + 1
        relatorio["por_estado"][a["spec"]["estado"]] = relatorio["por_estado"].get(a["spec"]["estado"], 0) + 1
    if gravar:
        alvo = lista_caminho(destino or pasta)
        os.makedirs(os.path.dirname(os.path.abspath(alvo)), exist_ok=True)
        open(alvo, "w", encoding="utf8").write(json.dumps(lista, ensure_ascii=False, indent=2) + "\n")
        relatorio["arquivo"] = alvo
    return lista, relatorio


def _alertas_rel(codigo, spec, em_uso, lista, reg):
    """Inventário do relatório de condução: as rodadas do arquivo conferem com as registradas na lista e com as das
    versões anteriores ainda na pasta? Rodada que não confere não passa para a lista: ela mantém as registradas."""
    if not spec.get("rodadas"):
        return []      # relatório comum: não é de condução
    alertas = []
    hs = digitais_rodadas(spec["rodadas"])
    antes = next((d for d in lista["documentos"] if chave_registro(d) == chave_registro(reg)), None)
    pub = (antes or {}).get("rodadas_sha256") or []
    if len(pub) > len(hs):
        alertas.append(f'{codigo}: a lista mestra registra {len(pub)} rodada(s), mas o arquivo mais recente na pasta (v{spec.get("versao")}) tem {len(hs)}. '
                       "A versão vigente sumiu da pasta?")
        reg["rodadas_sha256"] = pub
    elif hs[:len(pub)] != pub:
        alertas.append(f'{codigo} v{spec.get("versao")}: as rodadas não conferem com as registradas na lista mestra (rodada publicada alterada). '
                       "A lista mantém as rodadas registradas.")
        reg["rodadas_sha256"] = pub
    for m in em_uso:
        r2 = m["spec"].get("rodadas") or []
        if m["spec"] is not spec and digitais_rodadas(r2) != hs[:len(r2)]:
            alertas.append(f'{m["nome"]}: as rodadas desta versão não conferem com as da v{spec.get("versao")} (rodada publicada alterada).')
    return alertas


def _mesclar_em(lista, registro):
    """Mesma regra de gravar_lista, mas sem tocar no disco (usada pelo inventário)."""
    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    chave = chave_registro(registro)
    antes = next((d for d in lista["documentos"] if chave_registro(d) == chave), None)
    novo = dict(antes or {}, **registro)
    # campos que só valem para a entrada obsoleta não sobrevivem à republicação do código
    for k in ("substituido_por", "motivo_obsoleto"):
        if k not in registro: novo.pop(k, None)
    # confidencial: o esqueleto apaga o que a entrada tinha antes; deixou de ser confidencial, a marca sai
    for k in (CAMPOS_FORA_DO_CONFIDENCIAL if registro.get("confidencial") else ("confidencial",)):
        if k not in registro: novo.pop(k, None)
    novo["registrado_em"] = (antes or {}).get("registrado_em", agora)
    novo["atualizado_em"] = agora
    if antes:
        lista["documentos"][lista["documentos"].index(antes)] = novo
    else:
        lista["documentos"].append(novo)
    lista["documentos"].sort(key=lambda d: (str(d.get("empresa")), str(d.get("tipo")), str(d.get("area")),
                                            str(d.get("ano") or ""), int(d.get("numero") or 0), str(d.get("arquivo") or "")))
    lista["atualizado_em"] = agora
    return lista


# ----------------------------------------------------------------------------- atualização
def _digital(obj):
    import hashlib
    return hashlib.sha1(json.dumps(obj, ensure_ascii=False, sort_keys=True).encode("utf8")).hexdigest()[:16]


def _hoje():
    return datetime.date.today().strftime("%d/%m/%Y")


def versao_publicada(ficha):
    """Versão oficial que a minuta vai substituir, tirada da ficha ("ACME-POL-ADM-007 v1.0" -> "1.0"). Espelho de versaoPublicada."""
    m = re.search(r"\sv(\d+\.\d+)\s*$", str(ficha.get("substitui") or ""))
    return m.group(1) if m else None


def preparar_atualizacao(anterior, mudanca=None, o_que_mudou="", estado=None, responsavel=None, data=None):
    """Monta a especificação da nova versão a partir do documento lido. Devolve (spec_nova, avisos)."""
    avisos = []
    est_ant = anterior.get("estado")
    if est_ant == "OBSOLETO":
        raise SystemExit("O documento recebido é OBSOLETO: atualize a versão vigente que o substituiu" +
                         (f" ({anterior['substituido_por']})." if anterior.get("substituido_por") else "."))
    nova = json.loads(json.dumps(anterior))
    for k in ("codigo", "vigencia", "aprovado_por", "data_aprovacao", "cargo_aprovador", "substituido_por", "motivo_obsoleto"):
        nova.pop(k, None)
    v_ant = str(anterior.get("versao", "1.0"))
    base = v_ant.split("-r")[0]
    # mudanca: "conteudo" (X+1.0) ou "correcao" (X.Y+1), contadas da versão oficial; "rodada", a revisão seguinte da
    # minuta. Sem ela: oficial muda o conteúdo, minuta segue na rodada seguinte (mesma regra de sugerirAtualizacao).
    mudanca = mudanca or ("conteudo" if est_ant == "OFICIAL" else "rodada")

    def salto(v):
        # parte a parte, como o parseInt(t, 10) || 0 da página
        x, y = [int(m.group()) if m else 0 for m in (re.match(r"\s*[+-]?\d+", t) for t in (str(v).split(".") + ["0"])[:2])]
        return f"{x}.{y + 1}" if mudanca == "correcao" else f"{x + 1}.0"
    if est_ant == "OFICIAL":
        nova["versao"] = f"{salto(base)}-r1"
        nova["estado"] = estado or "MINUTA"
        nova["substitui"] = f'{anterior.get("codigo", "")} v{v_ant}'.strip()
    else:
        rod = int(v_ant.split("-r")[1]) + 1 if "-r" in v_ant else 1
        nova["versao"] = f"{base}-r{rod}"
        pub = versao_publicada(anterior)
        # minuta de nova versão: trocar o tipo de alteração muda a versão alvo; a mesma versão alvo segue na rodada seguinte
        if pub and mudanca != "rodada" and salto(pub) != base:
            nova["versao"] = f"{salto(pub)}-r1"
        # Estado aposentado (rascunho, em revisão): a rodada seguinte já sai com o nome atual.
        nova["estado"] = estado or ("MINUTA" if not est_ant else ESTADO_SUCESSOR.get(est_ant, est_ant))
    if nova["estado"] in ("OFICIAL", "OBSOLETO"):
        nova["versao"] = nova["versao"].split("-r")[0]
    nova["data_emissao"] = data or _hoje()
    # local e data de assinatura: a da nova versão, não a herdada do arquivo anterior
    for b in nova.get("corpo") or []:
        if isinstance(b, dict) and b.get("t") == "local_data":
            b["data"] = nova["data_emissao"]
    hist = list(anterior.get("historico") or [])
    if not hist:
        hist = [{"versao": v_ant, "data": anterior.get("data_aprovacao") or anterior.get("data_emissao", ""),
                 "descricao": "Emissão inicial.", "responsavel": anterior.get("elaborado_por", "")}]
    if not o_que_mudou:
        avisos.append("Informe o que mudou nesta versão: vai para o histórico de revisões.")
    hist.append({"versao": nova["versao"], "data": nova["data_emissao"], "descricao": o_que_mudou or "[definir]",
                 "responsavel": responsavel or anterior.get("elaborado_por", "")})
    nova["historico"] = hist
    if responsavel: nova["elaborado_por"] = responsavel
    return nova, avisos


def obsoleto_da_anterior(anterior, nova):
    """Cópia da versão anterior marcada como OBSOLETO, apontando para a nova versão oficial."""
    velha = json.loads(json.dumps(anterior))
    velha.pop("codigo", None)
    velha["estado"] = "OBSOLETO"
    velha["substituido_por"] = f'{nova["codigo"]} v{nova["versao"]}'
    return velha


def gerar(spec, ativos, saida, anexos=None):
    avisos = validar(spec)
    if spec.get("tipo") == "REL" and spec.get("rodadas"):
        montar_rel(spec)
    if spec.get("estado") != "OBSOLETO":
        av = adiar_assinaturas(spec)
        if av: avisos.append(av)
    # "anexos_externos": arquivos (caminhos) a anexar, ou a lista registrada na ficha. Não confundir com "anexos",
    # os anexos ABNT dentro do documento longo.
    arquivos_anexos = list(anexos or []) + [x for x in spec.get("anexos_externos") or [] if isinstance(x, str)]
    spec["anexos_externos"] = [x for x in spec.get("anexos_externos") or [] if isinstance(x, dict)]
    if not spec["anexos_externos"]: spec.pop("anexos_externos")
    if spec.get("com_anexos") or arquivos_anexos:
        if spec.get("estado") != "OBSOLETO":
            # documento e anexos juntos numa pasta com o código (a própria --saida, se ela já for essa pasta)
            if saida.lower().endswith(".docx"):
                pasta_doc = os.path.dirname(os.path.abspath(saida))
            else:
                pasta_doc = saida if os.path.basename(os.path.normpath(saida)) == nome_pasta_doc(spec) else os.path.join(saida, nome_pasta_doc(spec))
                saida = pasta_doc
            avisos += preparar_anexos(spec, arquivos_anexos, pasta_doc)
        montar_anexos(spec)
    spec.pop("anexos_origem", None)
    if not saida.lower().endswith(".docx"):
        os.makedirs(saida, exist_ok=True)
        saida = os.path.join(saida, nome_arquivo(spec))
    elif os.path.dirname(saida):
        os.makedirs(os.path.dirname(saida), exist_ok=True)
    montar_assinaturas(spec)
    ctx = gravar(spec, ativos, saida)
    # automático: linhas de assinatura; se elas ficariam sozinhas numa página, tabela
    if spec.get("com_assinatura") and spec.get("envolvidos") and spec.get("assinatura_formato", "auto") not in ("linhas", "tabela"):
        if assinaturas_sozinhas(saida, spec):
            montar_assinaturas(spec, "tabela")
            ctx = gravar(spec, ativos, saida)
            avisos.append("As linhas de assinatura ficariam sozinhas na última página: as assinaturas saíram em tabela.")
    if spec["variante"] == "longo" and spec.get("sumario", True):
        tit = [t for t in ctx.titulos if t[0] <= 3]
        pags = paginas_dos_titulos(saida, tit)
        ctx = gravar(spec, ativos, saida, pags, atualizar_campos=pags is None)
        if pags is None: avisos.append("Sumário sem números de página (LibreOffice indisponível): o Word pedirá para atualizar os campos ao abrir.")
    pdf = None
    if GERAR_PDF:
        pdf = para_pdf(saida, os.path.dirname(os.path.abspath(saida)))
        if not pdf: avisos.append("PDF não gerado (LibreOffice indisponível).")
    for av in avisos + ctx.avisos: print("AVISO:", av)
    print(saida)
    if pdf: print(pdf)
    return saida


GERAR_PDF = False   # --pdf: grava também o PDF de conferência ao lado de cada .docx


def localizar_marca(arg, spec_path=None):
    """--marca, TIMBRADOR_MARCA, marca.json ao lado da especificação ou em assets/ da habilidade."""
    candidatos = [arg, os.environ.get("TIMBRADOR_MARCA"),
                  os.path.join(os.path.dirname(os.path.abspath(spec_path)), "marca.json") if spec_path else None,
                  os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "marca.json")]
    for c in candidatos:
        if c and os.path.exists(c):
            return c
    raise SystemExit("marca.json não encontrado. Informe --marca, defina TIMBRADOR_MARCA ou crie um a partir de "
                     "assets/marca.exemplo.json.")


def main():
    ap = argparse.ArgumentParser(description="Timbrador white label: Padrão de Documentação Oficial (ABNT + papel timbrado)")
    ap.add_argument("spec", nargs="?", help="especificação JSON do documento a gerar")
    ap.add_argument("--marca", help="marca.json da organização (padrão: TIMBRADOR_MARCA, ao lado da especificação ou assets/marca.json)")
    ap.add_argument("--saida", default=".")
    ap.add_argument("--pdf", action="store_true", help="grava também o PDF de conferência (precisa do LibreOffice)")
    ap.add_argument("--checar-marca", metavar="MARCA", help="confere o marca.json e mostra o resumo do cadastro")
    ap.add_argument("--embutir-logos", metavar="MARCA", help="grava uma cópia do marca.json com os logos em data URI")
    ap.add_argument("--ler", metavar="DOCX", help="lê um documento timbrado e imprime a especificação JSON (ficha + corpo)")
    ap.add_argument("--atualizar", metavar="DOCX", help="lê um documento timbrado e prepara a especificação da nova versão")
    ap.add_argument("--mudanca", choices=["conteudo", "correcao", "rodada"], help="mudança de conteúdo (X+1.0) ou só correção (X.Y+1), contadas da versão oficial (a do documento OFICIAL ou a que a minuta substitui); rodada: revisão seguinte da minuta, mesma versão. Padrão: conteudo no OFICIAL, rodada na minuta")
    ap.add_argument("--o-que-mudou", default="", help="descrição da alteração para o histórico")
    ap.add_argument("--estado", help="estado da nova versão (padrão: automático)")
    ap.add_argument("--responsavel", help="quem elaborou a nova versão")
    ap.add_argument("--pasta-midia", help="pasta para salvar figuras do corpo")
    ap.add_argument("--inventariar", metavar="PASTA", help="varre a pasta dos documentos (subpastas por sigla) e monta/atualiza a lista mestra")
    ap.add_argument("--conferir", action="store_true", help="com --inventariar: só mostra o relatório, sem gravar a lista")
    ap.add_argument("--ignorar", metavar="PASTA", action="append", default=[], help="com --inventariar: nome de pasta a pular (pode repetir)")
    ap.add_argument("--lista", metavar="CAMINHO", help="lista mestra (pasta dos documentos ou o próprio lista-mestra.json): confere o número e registra a publicação")
    ap.add_argument("--obsoletar", metavar="DOCX", help="gera a cópia OBSOLETO de um documento OFICIAL (use o arquivo original, sem alterações)")
    ap.add_argument("--substituido-por", help="com --obsoletar: código e versão que substituem, ex.: 'ABC-CI-ADM-004/2026 v2.0'")
    ap.add_argument("--anexo", metavar="ARQUIVO", action="append", default=[],
                    help="anexo externo (pode repetir): documento e anexos vão juntos para uma pasta com o código do documento")
    ap.add_argument("--rodada", metavar="JSON", help="com --atualizar de um relatório de condução (REL): a rodada a acrescentar")
    ap.add_argument("--json-saida", help="arquivo para gravar a especificação lida/preparada (padrão: tela)")
    a = ap.parse_args()
    global GERAR_PDF
    GERAR_PDF = a.pdf

    if a.checar_marca:
        carregar_marca(a.checar_marca)
        print(resumo_marca())
        return
    if a.embutir_logos:
        destino = a.saida if a.saida.lower().endswith(".json") else os.path.join(
            a.saida, os.path.splitext(os.path.basename(a.embutir_logos))[0] + "_portatil.json")
        carregar_marca(a.embutir_logos)   # confere antes de copiar
        print(embutir_logos(a.embutir_logos, destino))
        return
    ativos = carregar_marca(localizar_marca(a.marca, a.spec))

    if a.inventariar:
        lista, r = inventariar(a.inventariar, a.lista, gravar=not a.conferir, ignorar=a.ignorar)
        junta = lambda d: " · ".join(f"{k} {v}" for k, v in sorted(d.items(), key=lambda x: -x[1])) or "nenhum"
        print(f"INVENTÁRIO de {r['pasta']}")
        print(f"  {r['arquivos_docx']} arquivo(s) .docx  ·  {r['documentos']} documento(s) distintos")
        if r["listas_encontradas"] and not a.lista:
            print(f"  lista mestra encontrada em: {r['listas_encontradas'][0]}")
        print(f"  origem dos dados: {junta(r['por_origem'])}")
        print(f"  por tipo: {junta(r['por_tipo'])}")
        print(f"  por estado: {junta(r['por_estado'])}")
        print()
        for d in lista["documentos"]:
            marca = "*" if d.get("origem") in ("nome", "cabecalho") else " "
            print(f'{marca} {d["codigo"]:<22} v{str(d.get("versao")):<8} {str(d.get("estado")):<10} '
                  f'{(d.get("titulo") or "")[:46]:<46} {d.get("pasta", "."):<10} {d.get("responsavel") or "-"}')
        print("\n  (* dados lidos do cabeçalho ou do nome do arquivo: confira título e responsável)" if any(
            d.get("origem") in ("nome", "cabecalho") for d in lista["documentos"]) else "")
        for al in r["alertas"]:
            print(f"  ATENÇÃO — {al}")
        if r["removidas"]:
            print(f"  Saíram da lista {len(r['removidas'])} entrada(s) cujo arquivo não está mais na pasta:")
            for x in r["removidas"]:
                print(f"    {x}")
        if r["pastas_confidenciais"]:
            print("  Pasta Confidencial fora da varredura (" + "; ".join(r["pastas_confidenciais"]) + "): o que está nela não entra na lista, nem o título.")
        if r["confidenciais"]:
            print(f"  {len(r['confidenciais'])} documento(s) confidencial(is) na lista ficaram sem conferir (estão na pasta Confidencial): " + "; ".join(r["confidenciais"]))
        if r["nao_conferidas"]:
            print(f"  {len(r['nao_conferidas'])} entrada(s) ficaram sem conferir (pasta não varrida): " + "; ".join(r["nao_conferidas"]))
        if r["varias_versoes"]:
            print(f"  {len(r['varias_versoes'])} documento(s) com mais de um arquivo (na lista ficou a versão em uso: a oficial mais alta, quando há):")
            for codigo, vs in r["varias_versoes"][:20]:
                print(f"    {codigo}: {', '.join(vs)}")
        if r["conflitos"]:
            print("  ATENÇÃO — mesmo código com títulos diferentes (numeração repetida):")
            for codigo, titulos in r["conflitos"]:
                print(f"    {codigo}: " + " | ".join(titulos))
        if r["versoes_fora_do_padrao"]:
            print("  Versões fora do padrão X.Y ou X.Y-rN (acerte ao gerar a próxima versão):")
            for codigo, v in r["versoes_fora_do_padrao"]:
                print(f"    {codigo}: v{v}")
        if r["sem_numero"]:
            print(f"  {len(r['sem_numero'])} arquivo(s) sem número (código com XXX): reserve o número:")
            for codigo, nome in r["sem_numero"]:
                print(f"    {codigo} — {nome}")
        if r["sem_identificacao"]:
            print(f"  {len(r['sem_identificacao'])} arquivo(s) .docx não identificados (ficaram de fora):")
            for rel, motivo in r["sem_identificacao"][:20]:
                print(f"    {rel} — {motivo}")
        if r["ignorados"]:
            print(f"  {len(r['ignorados'])} arquivo(s) que não são .docx foram ignorados.")
        if r.get("modelos"):
            print(f"  {len(r['modelos'])} arquivo(s) de partida de modelo foram pulados (não são documentos): {', '.join(r['modelos'])}.")
        print()
        print(f"LISTA MESTRA: {r['arquivo']} ({len(lista['documentos'])} documentos)" if not a.conferir
              else "Nada foi gravado (--conferir).")
        return

    if a.obsoletar:
        velha, info = ler_timbrado(a.obsoletar, a.pasta_midia)
        if info["origem"] == "sem_timbrado":
            raise SystemExit("O arquivo não é um documento do padrão.")
        if velha.get("estado") != "OFICIAL":
            raise SystemExit(f"Só documento OFICIAL pode ser marcado como obsoleto (este está como {velha.get('estado')}).")
        if velha.get("tipo") == "REL" and velha.get("rodadas"):
            raise SystemExit("Relatório de condução não tem cópia obsoleta: cada versão nova contém as rodadas da anterior.")
        if not a.substituido_por:
            raise SystemExit("Informe --substituido-por com o código e a versão que substituem este documento.")
        velha.pop("codigo", None)
        velha["estado"] = "OBSOLETO"
        velha["substituido_por"] = a.substituido_por
        for av in info["avisos"]: print("AVISO:", av, file=sys.stderr)
        saida_velha = gerar(velha, ativos, a.saida)
        if a.lista and _mesmo_codigo(a.substituido_por, velha):
            print("LISTA MESTRA: não alterada — a substituição é nova versão do mesmo código, e a entrada aponta a vigente.")
        elif a.lista:   # substituído por outro código (ex.: ABC-POL-ADM-001 → XYZ-POL-ADM-001): o código antigo sai de uso
            gravar_lista(a.lista, registro_de(velha, saida_velha))
            print(f"LISTA MESTRA: {lista_caminho(a.lista)} atualizada ({velha['codigo']} marcado como obsoleto).")
        return

    if a.ler or a.atualizar:
        spec, info = ler_timbrado(a.ler or a.atualizar, a.pasta_midia)
        avisos = list(info["avisos"])
        if a.atualizar:
            if info["origem"] == "sem_timbrado":
                raise SystemExit("O arquivo não é um documento do padrão; use --ler e trate como documento novo.")
            if info["origem"] == "modelo":
                raise SystemExit("É um arquivo de partida de modelo, não um documento: use --ler e trate como documento novo.")
            anterior = spec
            if anterior.get("tipo") == "REL" and anterior.get("rodadas") or a.rodada:
                # relatório de condução: só recebe rodadas; a linha mudada no Word é recusada aqui
                if not a.rodada:
                    raise SystemExit("Relatório de condução só recebe rodadas: use --atualizar com --rodada rodada.json.")
                erros = conferir_quadro_rodadas(anterior.get("corpo"), anterior.get("rodadas") or [])
                if erros:
                    raise SystemExit(" ".join(erros))
                spec = preparar_rodada(anterior, json.load(open(a.rodada, encoding="utf8")), a.responsavel)
                erros = conferir_rodada(spec["rodadas"][-1], spec["rodadas"][:-1])
                if erros:
                    raise SystemExit("RODADA: " + " ".join(erros))
            else:
                spec, av = preparar_atualizacao(anterior, a.mudanca, a.o_que_mudou, a.estado and a.estado.upper(), a.responsavel)
                avisos += av
            spec["anterior"] = {k: v for k, v in anterior.items()}
            # impressão digital do bloco "anterior": é dele que sai a cópia obsoleta, que tem de ser igual à vigente
            spec["anterior_conferencia"] = _digital(spec["anterior"])
            if spec.get("com_anexos"):   # os anexos registrados estão na pasta do documento lido
                spec["anexos_origem"] = os.path.dirname(os.path.abspath(a.atualizar))
        texto = json.dumps(spec, ensure_ascii=False, indent=2)
        if a.json_saida:
            open(a.json_saida, "w", encoding="utf8").write(texto)
        else:
            print(texto)
        for av in avisos: print("AVISO:", av, file=sys.stderr)
        print(f"ORIGEM: {info['origem']}", file=sys.stderr)
        return

    if not a.spec:
        ap.error("informe a especificação JSON, --ler ou --atualizar")
    spec = json.load(open(a.spec, encoding="utf8"))
    # figuras com caminho relativo: a partir da pasta da especificação
    base_spec = os.path.dirname(os.path.abspath(a.spec))
    for b in [x for x in spec.get("corpo") or [] if isinstance(x, dict)] + [x for ap_ in (spec.get("apendices") or []) + (spec.get("anexos") or [])
                                                                          for x in ap_.get("corpo") or [] if isinstance(x, dict)]:
        if b.get("t") == "figura" and b.get("caminho") and not os.path.isabs(b["caminho"]):
            b["caminho"] = os.path.join(base_spec, b["caminho"])
    anterior = spec.pop("anterior", None)
    conferencia = spec.pop("anterior_conferencia", None)
    if anterior and conferencia and _digital(anterior) != conferencia:
        raise SystemExit("O bloco \"anterior\" da especificação foi alterado depois do --atualizar. Ele guarda a versão vigente "
                         "e vira a cópia obsoleta, que precisa ser idêntica a ela: refaça o --atualizar e edite só o corpo da nova versão.")
    lista = carregar_lista(a.lista) if a.lista else None
    rel = str(spec.get("tipo", "")).upper() == "REL" and bool(spec.get("rodadas"))
    if rel:
        # relatório de condução: número herdado do plano, oficial pela coordenação, rodadas conferidas com a lista
        spec["estado"] = str(spec.get("estado", "")).upper()
        erros = conferir_quadro_rodadas(spec.get("corpo"), anterior.get("rodadas") or []) if anterior else []
        e2, avisos = regras_rel(spec, lista, anterior)
        for av in avisos:
            print("AVISO:", av)
        if erros + e2:
            raise SystemExit("RELATÓRIO DE CONDUÇÃO: " + " ".join(erros + e2))
    publica = str(spec.get("estado", "MINUTA")).upper().replace(" ", "_") in ("OFICIAL", "OBSOLETO")
    if lista is not None and not publica:
        # a lista só registra publicações: minuta e em aprovação saem com XXX e não a alteram
        print("AVISO: minuta e em aprovação não reservam número nem alteram a lista mestra; --lista ignorado.")
        lista = None
    if lista is not None and str(spec.get("estado", "")).upper() == "OBSOLETO" and _mesmo_codigo(spec.get("substituido_por"), spec):
        print("AVISO: a substituição é nova versão do mesmo código: a entrada da lista continua apontando a vigente.")
        lista = None
    if lista is not None and not rel:
        if not spec.get("numero") and not spec.get("codigo"):
            spec["numero"] = proximo_numero(lista, spec)
            print(f"AVISO: número {spec['numero']:03d} reservado na lista mestra.")
        erros, avisos = conferir_lista(lista, spec, atualizacao=bool(anterior))
        for av in avisos:
            print("AVISO:", av)
        if erros:
            raise SystemExit("LISTA MESTRA: " + " ".join(erros))
        if not spec.get("elaborado_por"):
            print("AVISO: sem 'elaborado_por': a lista mestra vai registrar a publicação sem responsável.")
    saida = gerar(spec, ativos, a.saida, a.anexo)
    if eh_confidencial(spec) and not any(eh_pasta_confidencial(x) for x in os.path.abspath(saida).split(os.sep)[:-1]):
        print(f"AVISO: documento confidencial gravado fora da pasta Confidencial ({saida}): mova-o para lá e apague esta cópia.")
    if lista is not None:
        lista = gravar_lista(a.lista, registro_de(spec, saida))
        print(f"LISTA MESTRA: {lista_caminho(a.lista)} ({len(lista['documentos'])} documentos)")
    if rel and anterior:
        print(f'AVISO: relatório de condução não tem cópia obsoleta: retire da pasta a v{anterior.get("versao")}, que a v{spec["versao"]} contém inteira.')
    elif anterior and spec["estado"] == "OFICIAL" and anterior.get("estado") == "OFICIAL" and spec.get("gerar_obsoleto_anterior", True):
        velha = obsoleto_da_anterior(anterior, spec)
        print("AVISO: gerada também a versão anterior marcada como OBSOLETO (arquive-a e retire de uso).")
        definir_empresa(velha.get("empresa") or ORG["padrao"])
        saida_velha = gerar(velha, ativos, a.saida)
        # a cópia obsoleta não entra na lista: a entrada do código aponta a versão vigente, que acabou de ser registrada
        definir_empresa(spec.get("empresa") or ORG["padrao"])


if __name__ == "__main__":
    main()
