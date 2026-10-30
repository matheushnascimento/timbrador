#!/usr/bin/env python3
"""Timbrador white label: gera documentos no Padrão de Documentação Oficial (ABNT + papel timbrado).

A identidade visual e os dados cadastrais vêm do marca.json (modelo em assets/marca.exemplo.json); o visual,
as regras e a numeração são fixos e são os mesmos da página assets/timbrador.html.

Uso:
  python3 gerar_documento.py especificacao.json --marca marca.json [--saida PASTA/ ou arquivo.docx] [--pdf]
  python3 gerar_documento.py --checar-marca marca.json
  python3 gerar_documento.py --embutir-logos marca.json [--saida marca_portatil.json]
  python3 gerar_documento.py --montar-pagina marca.json [--saida timbrador.html]

Só usa a biblioteca padrão do Python 3.8+. LibreOffice (soffice) e pdftotext são opcionais: numeram as páginas
do sumário e geram o PDF de conferência. TIMBRADOR_SOFFICE troca o comando do LibreOffice.
"""
import json, re, sys, os, zipfile, shutil, shlex, hashlib, subprocess, tempfile, argparse, datetime, unicodedata, base64
from xml.sax.saxutils import escape

EMU_CM = 360000
TW_CM = 567
INDEFINIDO = "[definir]"
# Empresas: vêm do marca.json (carregar_marca). Campo None ainda não foi informado: aparece como [definir] e
# impede o estado OFICIAL. Empresa "grupo" emite documento que vale para mais de uma empresa: o grupo não é pessoa
# jurídica, então o cabeçalho traz só o nome, sem documento, e o rodapé leva os logos de todas as empresas ativas.
# Empresa "aposentada" fica só para ler e versionar documentos emitidos antes.
EMPRESAS = {}
EMPRESAS_ATIVAS = []
# Empresas cujos logos formam o rodapé do documento do grupo, na ordem do marca.json.
EMPRESAS_DO_GRUPO = []
EMPRESA_ID = None
EMPRESA = {}
EMPRESA_PADRAO = None
MARCA = {"organizacao": "", "politica": {}}


def definir_empresa(chave):
    global EMPRESA_ID, EMPRESA
    if chave not in EMPRESAS:
        raise SystemExit(f"Empresa desconhecida: {chave}. Use uma de: {', '.join(EMPRESAS)} (marca.json).")
    EMPRESA_ID, EMPRESA = chave, EMPRESAS[chave]


def dado(campo):
    return EMPRESA.get(campo) or INDEFINIDO


def identificacao_empresa():
    if EMPRESA.get("grupo"):
        return EMPRESA["marca"]
    partes = [EMPRESA["razao"] or EMPRESA["marca"], f'{EMPRESA["doc_rotulo"]} {dado("documento")}']
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
# Áreas padrão; o marca.json pode trocar ("areas") ou ampliar ("areas_extra") esta lista.
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


def _vazio(v):
    return v is None or str(v).strip() in ("", INDEFINIDO)


def ler_imagem(ref, base):
    """Logo ou figura: caminho (relativo a base) ou data URI base64. Confere se é PNG ou JPEG."""
    if str(ref).startswith("data:"):
        dados = base64.b64decode(ref.split(",", 1)[1])
    else:
        with open(ref if os.path.isabs(ref) else os.path.join(base, ref), "rb") as f:
            dados = f.read()
    dim_imagem(dados)
    return dados


def re_codigo():
    """Código impresso no cabeçalho ("SIGLA-TIPO-ÁREA-NNN/AAAA · vX.Y"), com as siglas do marca.json."""
    siglas = "|".join(sorted(EMPRESAS, key=len, reverse=True)) or "[A-Z][A-Z0-9]{1,5}"
    return re.compile(r"\b(" + siglas + r")-([A-Z0-9]{2,6})-([A-Z0-9]{2,6})-(\d{3}|XXX)(?:/(\d{4}))?\s+·\s+(?:v|[Vv]ersão\s+)([0-9][0-9.]*(?:-r\d+)?)")


def carregar_marca(caminho):
    """Lê o marca.json (empresas, logos, áreas, tipos extras e política) e devolve os ativos: os logos."""
    global EMPRESAS_ATIVAS, EMPRESAS_DO_GRUPO, EMPRESA_PADRAO, RE_CODIGO
    base = os.path.dirname(os.path.abspath(caminho))
    try:
        with open(caminho, encoding="utf8") as f:
            cfg = json.load(f)
    except (OSError, ValueError) as e:
        raise SystemExit(f"Não consegui ler o marca.json ({caminho}): {e}")
    lista = cfg.get("empresas") or []
    if not lista:
        raise SystemExit("marca.json: a lista 'empresas' está vazia.")
    EMPRESAS.clear()
    logos = {}
    for e in lista:
        sig = str(e.get("sigla", "")).upper()
        if not re.fullmatch(r"[A-Z][A-Z0-9]{1,5}", sig):
            raise SystemExit(f"marca.json: sigla inválida {sig!r} (2 a 6 letras ou números, começando por letra).")
        if sig in EMPRESAS:
            raise SystemExit(f"marca.json: a sigla {sig} aparece duas vezes.")
        cor = str(e.get("cor") or "1F3864").lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", cor):
            raise SystemExit(f"marca.json: cor inválida em {sig} (use hexadecimal, ex.: 1F3864).")
        emp = {"marca": e.get("marca") or sig, "cor": cor, "doc_rotulo": e.get("documento_rotulo") or "CNPJ",
               "razao": None if _vazio(e.get("razao_social")) else e["razao_social"],
               "documento": None if _vazio(e.get("documento")) else e["documento"],
               "cidade": None if _vazio(e.get("cidade")) else e["cidade"]}
        for k in ("grupo", "aposentada"):
            if e.get(k): emp[k] = True
        if e.get("logo"):
            try:
                logos[sig] = ler_imagem(e["logo"], base)
            except (OSError, ValueError) as ex:
                raise SystemExit(f"marca.json: logo de {sig} ilegível ({ex}).")
        else:
            emp["sem_logo"] = True
        EMPRESAS[sig] = emp
    EMPRESAS_ATIVAS = [k for k in EMPRESAS if not EMPRESAS[k].get("aposentada")]
    if not EMPRESAS_ATIVAS:
        raise SystemExit("marca.json: todas as empresas estão aposentadas.")
    EMPRESAS_DO_GRUPO = [k for k in EMPRESAS_ATIVAS if not EMPRESAS[k].get("grupo")]
    EMPRESA_PADRAO = str(cfg.get("empresa_padrao") or EMPRESAS_ATIVAS[0]).upper()
    if EMPRESA_PADRAO not in EMPRESAS_ATIVAS:
        raise SystemExit(f"marca.json: empresa_padrao {EMPRESA_PADRAO} não está entre as empresas ativas.")
    if cfg.get("areas"):
        AREAS.clear(); AREAS.update({k.upper(): v for k, v in cfg["areas"].items()})
    AREAS.update({k.upper(): v for k, v in (cfg.get("areas_extra") or {}).items()})
    for sig, d in (cfg.get("tipos_extra") or {}).items():
        if not isinstance(d, dict) or d.get("variante") not in ("curto", "longo"):
            raise SystemExit(f"marca.json: o tipo extra {sig} precisa de variante 'curto' ou 'longo'.")
        (TIPOS_LONGOS if d["variante"] == "longo" else TIPOS_CURTOS)[sig.upper()] = d.get("nome") or sig.upper()
    MARCA.update(organizacao=cfg.get("organizacao") or "", politica=cfg.get("politica") or {},
                 tipos_ativos=[t.upper() for t in cfg.get("tipos_ativos") or []],
                 areas_ativas=[a.upper() for a in cfg.get("areas_ativas") or []])
    RE_CODIGO = re_codigo()
    definir_empresa(EMPRESA_PADRAO)
    return {"logos": logos}


def resumo_marca():
    linhas = [f"Marca: {MARCA['organizacao'] or '(sem nome de organização)'}"]
    for k, e in EMPRESAS.items():
        if e.get("grupo"):
            st = "documento do grupo (rodapé com os logos de todas as empresas ativas)"
        else:
            st = "pode emitir OFICIAL" if e["razao"] and e["documento"] else "só minuta e em aprovação (faltam razão social/documento)"
        logo = "SEM LOGO (rodapé só com o filete)" if e.get("sem_logo") else "logo ok"
        linhas.append(f"  {k}: {e['marca']} · #{e['cor']} · {logo} · {st}" + (" · APOSENTADA" if e.get("aposentada") else ""))
    tipos = MARCA["tipos_ativos"] or list(TIPOS_CURTOS) + list(TIPOS_LONGOS)
    areas = MARCA["areas_ativas"] or list(AREAS)
    linhas.append(f"  padrão: {EMPRESA_PADRAO} · áreas: {', '.join(areas)} · tipos: {', '.join(tipos)}")
    if MARCA["politica"].get("codigo"):
        linhas.append(f"  norma: {MARCA['politica']['codigo']} – {MARCA['politica'].get('titulo', '')}")
    return "\n".join(linhas)


def marca_portatil(caminho):
    """marca.json com os logos embutidos (data URI): um arquivo só, sem depender da pasta."""
    with open(caminho, encoding="utf8") as f:
        cfg = json.load(f)
    base = os.path.dirname(os.path.abspath(caminho))
    for e in cfg.get("empresas", []):
        ref = e.get("logo")
        if ref and not str(ref).startswith("data:"):
            dados = ler_imagem(ref, base)
            e["logo"] = f"data:image/{dim_imagem(dados)[2]};base64," + base64.b64encode(dados).decode()
    return cfg


PAGINA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "timbrador.html")


def montar_pagina(caminho_marca, saida):
    """Cópia da página com a marca embutida: quem abre não precisa carregar o marca.json."""
    cfg = marca_portatil(caminho_marca)
    html = open(PAGINA, encoding="utf8").read()
    js = json.dumps(cfg, ensure_ascii=False).replace("</", "<\\/")
    novo, n = re.subn(r'(<script type="application/json" id="marca-embutida">).*?(</script>)',
                      lambda m: m.group(1) + js + m.group(2), html, count=1, flags=re.S)
    if not n:
        raise SystemExit("A página não tem o bloco marca-embutida.")
    with open(saida, "w", encoding="utf8") as f:
        f.write(novo)
    return saida


def comando_soffice():
    cmd = os.environ.get("TIMBRADOR_SOFFICE")
    if cmd:
        return shlex.split(cmd)
    for nome in ("soffice", "libreoffice"):
        if shutil.which(nome):
            return [nome]
    return None


def para_pdf(docx, pasta):
    """Converte com o LibreOffice (perfil temporário, para não esbarrar numa instância aberta). None sem LibreOffice."""
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


def marca_dagua_vml(ctx, texto, cor):
    """Marca d'água nativa do Word (WordArt diagonal atrás do texto) — reconhecida também pelo LibreOffice."""
    n = len(texto)
    w = 470.0
    h = min(130.0, w / (n * 0.66))
    idn = ctx.novo_id()
    return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:pict>{SHAPETYPE_136}'
            f'<v:shape id="PowerPlusWaterMarkObject{idn}" o:spid="_x0000_s{2000 + idn}" type="#_x0000_t136" '
            f'style="position:absolute;margin-left:0;margin-top:0;width:{w:.1f}pt;height:{h:.1f}pt;rotation:315;z-index:-251654144;'
            f'mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page" '
            f'o:allowincell="f" fillcolor="#{cor}" stroked="f"><v:fill opacity="1"/>'
            f'<v:textpath style="font-family:&quot;Arial&quot;;font-size:1pt;font-weight:bold" string="{escape(texto)}"/>'
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


def marca(ctx, spec):
    est = ESTADOS[spec["estado"]]
    return marca_dagua_vml(ctx, est["marca"], est["claro"]) if est["marca"] else ""


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


def pagina_x_de_y(sz=8, cor="404040", b=False):
    return run("Página ", sz=sz, cor=cor, b=b) + campo("PAGE", "1", sz=sz, cor=cor, b=b) + run(" de ", sz=sz, cor=cor, b=b) + campo("NUMPAGES", "1", sz=sz, cor=cor, b=b)


def linhas_controle(ctx, spec):
    est = ESTADOS[spec["estado"]]
    texto = est["rodape"]
    if spec["estado"] == "OFICIAL" and spec.get("aprovado_por"):
        texto = f'Aprovado por {spec["aprovado_por"]} em {spec.get("data_aprovacao", "")}. ' + texto
    linha1 = par(marca(ctx, spec) + run(codigo_versao(spec) + "  ·  ", sz=8, cor="404040") + pagina_x_de_y(sz=8, cor="000000", b=True),
                 jc="center", antes=80, depois=0, linha=240, estilo="Rodape")
    linha2 = par(run(texto, sz=7.5, cor=est["cor"] if spec["estado"] != "OFICIAL" else "404040", b=spec["estado"] != "OFICIAL"),
                 jc="center", antes=0, depois=0, linha=240, estilo="Rodape")
    linha3 = par(run(identificacao_empresa(), sz=7.5, cor="7F7F7F"),
                 jc="center", antes=0, depois=0, linha=240, estilo="Rodape",
                 borda='<w:bottom w:val="single" w:sz="4" w:space="4" w:color="BFBFBF"/>')
    return linha1 + linha2 + linha3


def rodape_reduzido(ctx, parte, ativos, spec, largura_tw):
    """Rodapé das páginas internas do documento longo: filete na cor da empresa, logo à esquerda e título à direita."""
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
            f'{cel(c2, par(run(titulo, sz=8, cor="404040"), jc="right", antes=0, depois=0, linha=240, estilo="Rodape"))}</w:tr></w:tbl>'
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
                tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
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
        tbl += (f'<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{c1}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:tcPr>'
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
        sombra = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if cab else ""
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


def nome_pasta_doc(spec):
    """Pasta do documento com anexos: o código; minuta (XXX) leva também o título."""
    cod = str(spec.get("codigo") or "").replace("/", "-")
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
        sombra = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if cab else ""
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
            out += bloco_campos(b["itens"], largura_tw, b.get("rotulo_cm") or 3.3)
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
    f = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
    s = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
         f'<w:docDefaults><w:rPrDefault><w:rPr>{f}<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="pt-BR" w:eastAsia="pt-BR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
         f'<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>')
    s += ('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:widowControl/><w:spacing w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>')
    s += est("SemRecuo", "Texto sem recuo", '<w:ind w:firstLine="0"/>')
    tit = '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>'
    s += est("Heading1", "heading 1", tit + '<w:outlineLvl w:val="0"/>', "<w:b/><w:bCs/><w:caps/>", prox="Normal")
    s += est("Heading2", "heading 2", tit + '<w:outlineLvl w:val="1"/>', "<w:caps/>", prox="Normal")
    s += est("Heading3", "heading 3", tit + '<w:outlineLvl w:val="2"/>', "<w:b/><w:bCs/>", prox="Normal")
    s += est("Heading4", "heading 4", tit + '<w:outlineLvl w:val="3"/>', "<w:b/><w:i/>", prox="Normal")
    s += est("Heading5", "heading 5", tit + '<w:outlineLvl w:val="4"/>', "<w:i/>", prox="Normal")
    centro = '<w:keepNext/><w:spacing w:before="0" w:after="360" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>'
    s += est("TituloSemNumero", "Título sem indicativo numérico", centro + '<w:outlineLvl w:val="0"/>', "<w:b/><w:bCs/><w:caps/>", prox="Normal")
    s += est("TituloSumario", "Título do sumário", centro, "<w:b/><w:bCs/><w:caps/>", prox="Normal")
    s += est("TituloDocumento", "Título do documento", '<w:keepNext/><w:spacing w:before="0" w:after="60" w:line="276" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>',
             '<w:b/><w:bCs/><w:caps/><w:sz w:val="26"/><w:szCs w:val="26"/>')
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
        (3, nivel(0, "bullet", "•", 1134, 425, "Arial") + nivel(1, "bullet", "–", 1560, 426) + nivel(2, "bullet", "–", 1985, 425)),
        (4, nivel(0, "bullet", "☐", 1134, 425, "Segoe UI Symbol") + nivel(1, "bullet", "☐", 1560, 426, "Segoe UI Symbol") + nivel(2, "bullet", "–", 1985, 425)),
    ]
    s = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
    for aid, lv in a:
        s += f'<w:abstractNum w:abstractNumId="{aid}"><w:multiLevelType w:val="multilevel"/>{lv}</w:abstractNum>'
    for nid, aid in ctx.nums:
        s += f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{aid}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
    return s + "</w:numbering>"


def settings_xml(atualizar):
    u = '<w:updateFields w:val="true"/>' if atualizar else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>'
            f'<w:defaultTabStop w:val="709"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>{u}'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<w:themeFontLang w:val="pt-BR"/><w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/></w:settings>')

# ----------------------------------------------------------------------------- validação da especificação
def validar(spec):
    av = []
    definir_empresa(str(spec.get("empresa") or EMPRESA_PADRAO).upper())
    spec["empresa"] = EMPRESA_ID
    spec["estado"] = spec.get("estado", "MINUTA").upper().replace(" ", "_").replace("Ã", "A")
    if spec["estado"] not in ESTADOS: raise SystemExit(f"Estado inválido: {spec['estado']}")
    if spec["estado"] not in ESTADOS_ATIVOS:
        av.append(f'Estado {ESTADOS[spec["estado"]]["rotulo"].lower()} foi aposentado: use minuta para o texto que ainda não é oficial.')
    tipo = spec.get("tipo", "").upper()
    if spec["estado"] != "OBSOLETO" and spec.get("titulo"):
        novo = titulo_sem_tipo(spec["titulo"], tipo)
        if novo != spec["titulo"]:
            av.append(f'Título sem o nome do tipo, que já sai impresso acima dele: "{spec["titulo"]}" passou a "{novo}".')
            spec["titulo"] = novo
    if "variante" not in spec:
        spec["variante"] = "longo" if tipo in TIPOS_LONGOS else "curto"
    if not spec.get("codigo"):
        area = spec.get("area", "ADM").upper()
        num = spec.get("numero")
        n = f"{int(num):03d}" if num else "XXX"
        if tipo in TIPOS_CURTOS and not spec.get("ano"):
            spec["ano"] = int(spec["data_emissao"][-4:]) if re.match(r"\d{2}/\d{2}/\d{4}$", str(spec.get("data_emissao", ""))) else datetime.date.today().year
        spec["codigo"] = f"{EMPRESA_ID}-{tipo or 'DOC'}-{area}-{n}" + (f"/{spec['ano']}" if tipo in TIPOS_CURTOS else "")
    if tipo and tipo not in TIPOS_CURTOS and tipo not in TIPOS_LONGOS:
        av.append(f"Tipo {tipo} não está entre os tipos do padrão nem nos tipos_extra do marca.json.")
    if str(spec.get("area", "")).upper() not in AREAS:
        av.append(f"Área {spec.get('area')} não está entre as áreas do marca.json.")
    faltam = [] if EMPRESA.get("grupo") else [c for c in ("razao", "documento") if not EMPRESA.get(c)]
    if faltam:
        nomes = " e ".join({"razao": "a razão social", "documento": f"o {EMPRESA['doc_rotulo']}"}[c] for c in faltam)
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
    spec["classificacao"] = spec.get("classificacao") or "Interno"
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
        corpo += par(run(f'{spec["codigo"]}  ·  Versão {spec["versao"]}  ·  Emissão {spec["data_emissao"]}  ·  ', sz=9, cor="404040") +
                     run(est["rotulo"].upper(), sz=9, b=True, cor=est["cor"]), jc="center", antes=0, depois=240, linha=240, estilo="SemRecuo")
        if spec.get("identificacao"):
            ident = spec["identificacao"]   # empresa e documento já estão no cabeçalho
            corpo += bloco_campos(ident, larg_t, 3.3)
        corpo += blocos(ctx, spec.get("corpo", []), D, larg_t, False)
        corpo += secao(rid_h1, rid_f1, MARG_TIMBRADO, 1.2, 1.2, "nextPage").replace("<w:type w:val=\"nextPage\"/>", "")
    else:
        # ---------------- capa (timbrado completo)
        altura = 29.7 - MARG_TIMBRADO[0] - MARG_TIMBRADO[2]
        def linha_capa(h_cm, conteudo, va):
            return (f'<w:tr><w:trPr><w:trHeight w:val="{int(h_cm * TW_CM)}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{larg_t}" w:type="dxa"/>'
                    f'<w:vAlign w:val="{va}"/></w:tcPr>{conteudo}</w:tc></w:tr>')
        cp = lambda x, **k: par(x, jc="center", antes=k.get("antes", 0), depois=k.get("depois", 0), linha=k.get("linha", 276), estilo="SemRecuo")
        meio = (cp(run(tipo_nome.upper(), sz=12, cor=EMPRESA["cor"], b=True, esp=40), depois=240) +
                cp(run(spec["titulo"].upper(), b=True, sz=22), linha=300, depois=200) +
                (cp(run(spec["subtitulo"], sz=14, cor="262626"), depois=480) if spec.get("subtitulo") else cp("", depois=480)))
        ctab = [["Código", spec["codigo"]], ["Versão", spec["versao"]], ["Estado", est["rotulo"].upper()], ["Classificação", spec["classificacao"]]]
        wqs = [int(larg_t * p / 4.6) for p in (1.6, 1, 1, 1)]
        quadro = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(wqs)}" w:type="dxa"/><w:jc w:val="center"/>'
                  '<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
                  f'<w:tblGrid>{grade(wqs)}</w:tblGrid><w:tr>')
        for (rot, val), wq in zip(ctab, wqs):
            cor = est["cor"] if rot == "Estado" else "000000"
            quadro += (f'<w:tc><w:tcPr><w:tcW w:w="{wq}" w:type="dxa"/></w:tcPr>'
                       f'{par(run(rot.upper(), sz=7, cor="595959", esp=10), jc="center", estilo="TabelaTexto", antes=60)}'
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
                 ["Estado", f'**{est["rotulo"].upper()}**'], ["Classificação", spec["classificacao"]],
                 ["Área responsável", spec.get("area_nome") or AREAS.get(spec.get("area"), spec.get("area", ""))],
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
        corpo += secao(rid_h2, rid_f2, MARG_ABNT, 1.2, 1.2)

    documento = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{corpo}</w:body></w:document>'
    ctx.rel(D, "styles", "styles.xml"); ctx.rel(D, "numbering", "numbering.xml"); ctx.rel(D, "settings", "settings.xml")
    return ctx, documento, partes_extra


TIPOS_REL = {"image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
             "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
             "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
             "styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
             "numbering": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
             "settings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
             "customXml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"}


VERSAO_TIMBRADOR = "2026-09-16"
CAMPOS_FICHA = ["empresa", "tipo", "area", "area_nome", "numero", "ano", "codigo", "versao", "estado", "variante",
                "titulo", "subtitulo", "titulo_curto", "classificacao", "local", "data_emissao",
                "elaborado_por", "revisado_por", "aprovado_por", "cargo_aprovador", "data_aprovacao", "vigencia",
                "substitui", "substituido_por", "motivo_obsoleto", "identificacao", "historico",
                "modelo", "com_assinatura", "envolvidos", "assinatura_formato", "assinaturas_previstas",
                "com_anexos", "anexos_externos", "projeto"]   # modelo: True só nos arquivos de partida dos modelos oficiais

# Campos de um arquivo de partida que NÃO passam para o documento escrito a partir dele.
CAMPOS_SO_DO_MODELO = ("modelo", "codigo", "numero", "ano", "versao", "estado", "titulo", "subtitulo", "titulo_curto",
                       "elaborado_por", "revisado_por", "aprovado_por", "cargo_aprovador", "data_aprovacao", "vigencia",
                       "substitui", "substituido_por", "motivo_obsoleto", "historico", "data_emissao")
NS_FICHA = "urn:timbrador:ficha"


def ficha_de(spec):
    f = {k: spec[k] for k in CAMPOS_FICHA if spec.get(k) not in (None, "", [])}
    if MARCA["politica"].get("codigo"):
        f["padrao"] = MARCA["politica"]["codigo"]
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
        z.writestr("word/settings.xml", settings_xml(atualizar_campos))
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
    cod = spec["codigo"].replace("/", "-")
    return f'{cod}_v{spec["versao"]}_{spec["estado"]}_{slug}.docx'


# ----------------------------------------------------------------------------- leitura de documento timbrado
import xml.etree.ElementTree as ET
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
RE_CODIGO = re_codigo()   # refeito por carregar_marca, com as siglas do marca.json


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
                if NS_FICHA in x:
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
            info["avisos"].append("Documento sem ficha: dados lidos do cabeçalho; confira título e responsáveis.")
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
    # capa do documento longo: tudo até o parágrafo que fecha a 1ª seção
    for idx, el in enumerate(elementos):
        ppr = el.find(_W + "pPr") if el.tag == _W + "p" else None
        if ppr is not None and ppr.find(_W + "sectPr") is not None:
            elementos = elementos[idx + 1:]
            spec.setdefault("variante", "longo")
            break

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

    mapa = {"Título": "titulo", "Classificação": "classificacao", "Área responsável": "area_nome", "Elaborado por": "elaborado_por",
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
    base = {"lista_mestra": MARCA["organizacao"], "padrao": MARCA["politica"].get("codigo", ""), "atualizado_em": "", "documentos": []}
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
    return f"{spec.get('empresa') or EMPRESA_PADRAO}-{tipo}-{area}-{n}" + (f"/{ano}" if tipo in TIPOS_CURTOS else "")


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
    """Regra de título: o título diz só o assunto; o tipo já sai impresso acima dele.
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
    if mesmo and atualizacao and _sem_acento(mesmo.get("titulo")) != _sem_acento(spec.get("titulo")):
        avisos.append(f'O título muda nesta versão: "{mesmo.get("titulo")}" passa a "{spec.get("titulo")}". A lista mestra fica com o novo.')
    elif mesmo and _sem_acento(mesmo.get("titulo")) != _sem_acento(spec.get("titulo")):
        erros.append(f'O código {codigo} já é "{mesmo.get("titulo")}" na lista mestra '
                     f'({mesmo.get("estado")}, v{mesmo.get("versao")}). Use o número {proximo_numero(lista, spec):03d} '
                     f'ou gere a nova versão a partir daquele documento.')
    elif mesmo and not atualizacao:
        avisos.append(f'{codigo} já estava na lista mestra (v{mesmo.get("versao")}, {mesmo.get("estado")}): a entrada foi atualizada.')
    homonimo = next((d for d in lista["documentos"] if d.get("codigo") != codigo
                     and _sem_acento(d.get("titulo")) == _sem_acento(spec.get("titulo"))), None)
    if homonimo:
        avisos.append(f'Já existe {homonimo.get("codigo")} com o mesmo título ("{homonimo.get("titulo")}"). Confirme se não é o mesmo documento.')
    return erros, avisos


def _mesmo_codigo(substituido_por, spec):
    """'SUA-CI-ADM-002/2026 v1.2' substitui o mesmo código do documento?"""
    alvo = str(substituido_por or "").split(" v")[0].strip()
    codigo = str(spec.get("codigo") or "").strip()
    if not codigo and str(spec.get("numero") or "").isdigit():
        tipo = str(spec.get("tipo", "")).upper()
        codigo = f'{spec.get("empresa") or EMPRESA_PADRAO}-{tipo}-{str(spec.get("area", "ADM")).upper()}-{int(spec["numero"]):03d}' + \
                 (f'/{spec.get("ano")}' if tipo in TIPOS_CURTOS and spec.get("ano") else "")
    return bool(alvo) and alvo == codigo


def registro_de(spec, arquivo):
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
    return reg


def gravar_lista(caminho, registro):
    """Relê o arquivo antes de gravar: outra pessoa pode ter registrado algo nesse meio-tempo."""
    alvo = lista_caminho(caminho)
    lista = _mesclar_em(carregar_lista(alvo), registro)
    pasta = os.path.dirname(os.path.abspath(alvo))
    os.makedirs(pasta, exist_ok=True)
    open(alvo, "w", encoding="utf8").write(json.dumps(lista, ensure_ascii=False, indent=2) + "\n")
    return lista


# ----------------------------------------------------------------------------- inventário da pasta
RE_NOME = re.compile(r"^([A-Z][A-Z0-9]{1,5})-([A-Z0-9]{2,6})-([A-Z0-9]{2,6})-(\d{3}|XXX)(?:-(\d{4}))?"
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
                if NS_FICHA in x:
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
    achados, ignorados, sem_id, listas, modelos, visitadas = [], [], [], [], [], set()
    for raiz, dirs, arquivos in os.walk(pasta):
        dirs[:] = sorted(d for d in dirs if d.lower() not in ignorar and not d.startswith("."))
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
            spec.setdefault("empresa", EMPRESA_PADRAO)
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
        itens.sort(key=lambda x: (_chave_versao(x["spec"].get("versao")),
                                  ORDEM_ESTADO.get(x["spec"].get("estado"), 0)))
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
        if len(oficiais) > 1:
            alertas.append(f'{codigo}: mais de uma versão oficial em uso (' +
                           "; ".join(f'v{o["spec"].get("versao")} em {o["pasta"]}' for o in oficiais) +
                           "). A anterior precisa ser tornada obsoleta.")
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
        reg = registro_de(spec, atual["nome"])
        reg["pasta"] = atual["pasta"]
        reg["origem"] = atual["origem"]
        if not reg.get("projeto") and projeto_da_pasta(atual["pasta"]):   # documento antigo, sem o campo na ficha
            reg["projeto"] = projeto_da_pasta(atual["pasta"])
        if len(itens) > 1:
            reg["arquivos"] = len(itens)
        lista = _mesclar_em(lista, reg)
        vistos.add(chave_registro(reg))
    # entradas sem arquivo: saem quando a pasta registrada foi varrida (ou é a de Modelos); fora disso, ficam e são avisadas
    removidas, nao_conferidas, ficam = [], [], []
    for d in lista["documentos"]:
        if chave_registro(d) in vistos:
            ficam.append(d); continue
        if "pasta" not in d:
            nao_conferidas.append(f'{d.get("codigo")} (sem pasta registrada)'); ficam.append(d); continue
        p = "" if d["pasta"] == "." else str(d["pasta"])
        if any(x.lower() in PASTAS_IGNORADAS for x in p.split("/")) or p in visitadas:
            removidas.append(f'{d.get("codigo")} v{d.get("versao")} ({d.get("arquivo") or p})')
        else:
            nao_conferidas.append(f'{d.get("codigo")} ({p})'); ficam.append(d)
    lista["documentos"] = ficam
    versoes_ruins = sorted({(a["spec"]["codigo"], str(a["spec"].get("versao"))) for a in achados
                            if not RE_VERSAO_OK.match(str(a["spec"].get("versao") or ""))})
    sem_numero = sorted({(a["spec"]["codigo"], a["nome"]) for a in achados if not a["spec"].get("numero")})
    relatorio = {"pasta": pasta, "arquivos_docx": len(achados), "documentos": len(por_codigo),
                 "versoes_fora_do_padrao": versoes_ruins, "sem_numero": sem_numero,
                 "ignorados": ignorados, "modelos": modelos, "sem_identificacao": sem_id, "conflitos": conflitos,
                 "varias_versoes": multiplos, "alertas": alertas, "removidas": removidas, "nao_conferidas": nao_conferidas,
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


def _mesclar_em(lista, registro):
    """Mesma regra de gravar_lista, mas sem tocar no disco (usada pelo inventário)."""
    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    chave = chave_registro(registro)
    antes = next((d for d in lista["documentos"] if chave_registro(d) == chave), None)
    novo = dict(antes or {}, **registro)
    # campos que só valem para a entrada obsoleta não sobrevivem à republicação do código
    for k in ("substituido_por", "motivo_obsoleto"):
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


def preparar_atualizacao(anterior, mudanca="conteudo", o_que_mudou="", estado=None, responsavel=None, data=None):
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
    try:
        x, y = [int(p) for p in (base.split(".") + ["0"])[:2]]
    except ValueError:
        x, y = 1, 0
    if est_ant == "OFICIAL":
        alvo_v = f"{x + 1}.0" if mudanca == "conteudo" else f"{x}.{y + 1}"
        nova["versao"] = f"{alvo_v}-r1"
        nova["estado"] = estado or "MINUTA"
        nova["substitui"] = f'{anterior.get("codigo", "")} v{v_ant}'.strip()
    else:
        rod = int(v_ant.split("-r")[1]) + 1 if "-r" in v_ant else 1
        nova["versao"] = f"{base}-r{rod}"
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
    for av in avisos + ctx.avisos: print("AVISO:", av)
    print(saida)
    return saida


def localizar_marca(arg, spec_path):
    """--marca; senão TIMBRADOR_MARCA; senão marca.json ao lado da especificação; senão assets/marca.json da habilidade."""
    candidatos = [arg, os.environ.get("TIMBRADOR_MARCA"),
                  os.path.join(os.path.dirname(os.path.abspath(spec_path)), "marca.json") if spec_path else None,
                  os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "marca.json")]
    for c in candidatos:
        if c and os.path.exists(c):
            return c
    raise SystemExit("marca.json não encontrado. Informe --marca, defina TIMBRADOR_MARCA ou crie um a partir de "
                     "assets/marca.exemplo.json.")


def main():
    ap = argparse.ArgumentParser(description="Timbrador white label: gerador do Padrão de Documentação Oficial")
    ap.add_argument("spec", nargs="?", help="especificação JSON do documento a gerar")
    ap.add_argument("--marca", help="marca.json da organização (empresas, logos, áreas)")
    ap.add_argument("--checar-marca", metavar="MARCA", help="confere o marca.json e mostra o resumo")
    ap.add_argument("--embutir-logos", metavar="MARCA", help="grava uma cópia do marca.json com os logos embutidos")
    ap.add_argument("--montar-pagina", metavar="MARCA", help="grava uma cópia da página timbrador.html com a marca embutida")
    ap.add_argument("--pdf", action="store_true", help="salva também o PDF de conferência (precisa do LibreOffice)")
    ap.add_argument("--saida", default=".")
    ap.add_argument("--ler", metavar="DOCX", help="lê um documento timbrado e imprime a especificação JSON (ficha + corpo)")
    ap.add_argument("--atualizar", metavar="DOCX", help="lê um documento timbrado e prepara a especificação da nova versão")
    ap.add_argument("--mudanca", choices=["conteudo", "correcao"], default="conteudo", help="para documento OFICIAL: mudança de conteúdo (X+1.0) ou só correção (X.Y+1)")
    ap.add_argument("--o-que-mudou", default="", help="descrição da alteração para o histórico")
    ap.add_argument("--estado", help="estado da nova versão (padrão: automático)")
    ap.add_argument("--responsavel", help="quem elaborou a nova versão")
    ap.add_argument("--pasta-midia", help="pasta para salvar figuras do corpo")
    ap.add_argument("--inventariar", metavar="PASTA", help="varre a pasta dos documentos (subpastas por sigla) e monta/atualiza a lista mestra")
    ap.add_argument("--conferir", action="store_true", help="com --inventariar: só mostra o relatório, sem gravar a lista")
    ap.add_argument("--ignorar", metavar="PASTA", action="append", default=[], help="com --inventariar: nome de pasta a pular (pode repetir)")
    ap.add_argument("--lista", metavar="CAMINHO", help="lista mestra (pasta dos documentos ou o próprio lista-mestra.json): confere o número e registra a publicação")
    ap.add_argument("--obsoletar", metavar="DOCX", help="gera a cópia OBSOLETO de um documento OFICIAL (use o arquivo original, sem alterações)")
    ap.add_argument("--substituido-por", help="com --obsoletar: código e versão que substituem, ex.: 'SUA-CI-ADM-004/2026 v2.0'")
    ap.add_argument("--anexo", metavar="ARQUIVO", action="append", default=[],
                    help="anexo externo (pode repetir): documento e anexos vão juntos para uma pasta com o código do documento")
    ap.add_argument("--json-saida", help="arquivo para gravar a especificação lida/preparada (padrão: tela)")
    a = ap.parse_args()

    if a.checar_marca:
        carregar_marca(a.checar_marca)
        print(resumo_marca())
        return
    if a.embutir_logos:
        carregar_marca(a.embutir_logos)
        destino = a.saida if a.saida.lower().endswith(".json") else os.path.join(
            a.saida, os.path.splitext(os.path.basename(a.embutir_logos))[0] + "_portatil.json")
        with open(destino, "w", encoding="utf8") as f:
            json.dump(marca_portatil(a.embutir_logos), f, ensure_ascii=False, indent=2)
        print(destino)
        return
    if a.montar_pagina:
        carregar_marca(a.montar_pagina)
        destino = a.saida if a.saida.lower().endswith((".html", ".htm")) else os.path.join(a.saida, "timbrador.html")
        print(montar_pagina(a.montar_pagina, destino))
        return
    ativos = carregar_marca(localizar_marca(a.marca, a.spec or a.ler or a.atualizar or a.obsoletar))

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
        if not a.substituido_por:
            raise SystemExit("Informe --substituido-por com o código e a versão que substituem este documento.")
        velha.pop("codigo", None)
        velha["estado"] = "OBSOLETO"
        velha["substituido_por"] = a.substituido_por
        for av in info["avisos"]: print("AVISO:", av, file=sys.stderr)
        saida_velha = gerar(velha, ativos, a.saida)
        if a.lista and _mesmo_codigo(a.substituido_por, velha):
            print("LISTA MESTRA: não alterada — a substituição é nova versão do mesmo código, e a entrada aponta a vigente.")
        elif a.lista:   # substituído por outro código (ex.: SUA-POL-ADM-001 → OUTRA-POL-ADM-001): o código antigo sai de uso
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
    publica = str(spec.get("estado", "MINUTA")).upper().replace(" ", "_") in ("OFICIAL", "OBSOLETO")
    if lista is not None and not publica:
        # a lista só registra publicações: minuta e em aprovação saem com XXX e não a alteram
        print("AVISO: minuta e em aprovação não reservam número nem alteram a lista mestra; --lista ignorado.")
        lista = None
    if lista is not None and str(spec.get("estado", "")).upper() == "OBSOLETO" and _mesmo_codigo(spec.get("substituido_por"), spec):
        print("AVISO: a substituição é nova versão do mesmo código: a entrada da lista continua apontando a vigente.")
        lista = None
    if lista is not None:
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
    if a.pdf:
        pdf = para_pdf(saida, os.path.dirname(os.path.abspath(saida)))
        print(f"PDF: {pdf}" if pdf else "AVISO: PDF não gerado (LibreOffice indisponível).")
    if lista is not None:
        lista = gravar_lista(a.lista, registro_de(spec, saida))
        print(f"LISTA MESTRA: {lista_caminho(a.lista)} ({len(lista['documentos'])} documentos)")
    if anterior and spec["estado"] == "OFICIAL" and anterior.get("estado") == "OFICIAL" and spec.get("gerar_obsoleto_anterior", True):
        velha = obsoleto_da_anterior(anterior, spec)
        print("AVISO: gerada também a versão anterior marcada como OBSOLETO (arquive-a e retire de uso).")
        definir_empresa(velha.get("empresa") or EMPRESA_PADRAO)
        saida_velha = gerar(velha, ativos, a.saida)
        # a cópia obsoleta não entra na lista: a entrada do código aponta a versão vigente, que acabou de ser registrada
        definir_empresa(spec.get("empresa") or EMPRESA_PADRAO)


if __name__ == "__main__":
    main()
