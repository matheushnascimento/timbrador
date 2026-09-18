#!/usr/bin/env python3
"""Timbrador white label.

Gera um .docx no Padrão de Documentação Oficial (ABNT + timbrado da
organização) a partir de uma especificação JSON e de um arquivo de marca
(marca.json). Usa só a biblioteca padrão do Python. Se houver LibreOffice
(soffice) e pdftotext, numera o sumário e, com --pdf, salva um PDF para
conferência.

Uso:
  python3 gerar_documento.py especificacao.json --marca marca.json --saida PASTA/
  python3 gerar_documento.py --checar-marca marca.json
  python3 gerar_documento.py --embutir-logos marca.json [--saida marca_portatil.json]
"""
import argparse
import base64
import datetime
import json
import os
import re
import shlex
import shutil
import struct
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from xml.sax.saxutils import escape as _esc

# ---------------------------------------------------------------------------
# Tabelas do padrão (podem ser ampliadas/substituídas pelo marca.json)
# ---------------------------------------------------------------------------
TIPOS_PADRAO = {
    # curto: ato pontual, sem capa
    "CI": ("Comunicação interna", "curto"), "COM": ("Comunicado", "curto"),
    "OFI": ("Ofício", "curto"), "DEC": ("Declaração", "curto"),
    "TER": ("Termo", "curto"), "ATA": ("Ata", "curto"),
    "AUT": ("Autorização", "curto"), "NOT": ("Notificação", "curto"),
    "CTR": ("Contrato", "curto"), "PRO": ("Proposta", "curto"),
    "FOR": ("Formulário", "curto"), "REC": ("Recibo", "curto"),
    # longo: documento de consulta, com capa, folha de controle e sumário
    "MAN": ("Manual", "longo"), "POL": ("Política", "longo"),
    "REG": ("Regulamento", "longo"), "POP": ("Procedimento operacional padrão", "longo"),
    "INS": ("Instrução de trabalho", "longo"), "PLN": ("Plano", "longo"),
    "PRJ": ("Projeto", "longo"), "REL": ("Relatório", "longo"),
}
AREAS_PADRAO = {
    "ADM": "Administrativo", "DIR": "Diretoria", "FIN": "Financeiro",
    "FIS": "Fiscal/contábil", "JUR": "Jurídico", "RH": "Recursos humanos",
    "VEN": "Vendas/comercial", "ECO": "E-commerce", "EST": "Estoque",
    "EXP": "Expedição", "CMP": "Compras", "OFC": "Oficina/assistência",
    "MKT": "Marketing", "TI": "Tecnologia",
}
ESTADOS = {
    "RASCUNHO": dict(nome="Rascunho", cor="7F7F7F", txt="FFFFFF", marca="RASCUNHO",
                     faixa="RASCUNHO — DOCUMENTO EM ELABORAÇÃO, SEM VALIDADE OFICIAL"),
    "MINUTA": dict(nome="Minuta", cor="F2C200", txt="1A1A1A", marca="MINUTA",
                   faixa="MINUTA — TEXTO PROPOSTO PARA ANÁLISE, SEM VALIDADE OFICIAL"),
    "EM_REVISAO": dict(nome="Em revisão", cor="2E75B6", txt="FFFFFF", marca="EM REVISÃO",
                       faixa="EM REVISÃO — AGUARDANDO APROVAÇÃO, SEM VALIDADE OFICIAL"),
    "OFICIAL": dict(nome="Oficial", cor="2E7D32", txt="FFFFFF", marca=None,
                    faixa="DOCUMENTO OFICIAL"),
    "OBSOLETO": dict(nome="Obsoleto", cor="C62828", txt="FFFFFF", marca="OBSOLETO",
                     faixa="OBSOLETO — NÃO UTILIZAR"),
}
CLASSIFICACOES = ["Público", "Interno", "Confidencial"]
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
NIVEIS_DESTAQUE = {"info": ("2E75B6", "EAF2FA"), "atencao": ("C99700", "FFF7DB"),
                   "erro": ("C62828", "FDECEA")}

# Geometria (twips: 1 cm = 567; EMU: 1 cm = 360000)
CM = 567
EMU_CM = 360000
PAG_L, PAG_A = 11906, 16838
MARG_SUP, MARG_ESQ, MARG_DIR, MARG_INF = 3 * CM, 3 * CM, 2 * CM, 2 * CM
LARG_TEXTO = PAG_L - MARG_ESQ - MARG_DIR  # 16 cm
FONTE = "Arial"

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word"')

INDEFINIDO = "[definir]"


class ErroEspec(Exception):
    pass


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def esc(t):
    return _esc(_CTRL.sub("", str(t)), {'"': "&quot;"})


def vazio(v):
    return v is None or str(v).strip() in ("", INDEFINIDO)


def sem_acento(t):
    return "".join(c for c in unicodedata.normalize("NFKD", str(t)) if not unicodedata.combining(c))


def slug(t):
    s = re.sub(r"[^a-z0-9]+", "-", sem_acento(t).lower()).strip("-")
    return s[:60].rstrip("-") or "documento"


def tom(cor, fator):
    """Mistura a cor com branco (fator 0 = cor, 1 = branco)."""
    r, g, b = (int(cor[i:i + 2], 16) for i in (0, 2, 4))
    f = lambda c: int(c + (255 - c) * fator)
    return f"{f(r):02X}{f(g):02X}{f(b):02X}"


def data_extenso(d):
    m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", d or "")
    if not m:
        return d
    dia, mes, ano = int(m[1]), int(m[2]), m[3]
    return f"{'1º' if dia == 1 else dia} de {MESES[mes - 1]} de {ano}"


def dimensoes_imagem(dados):
    """(largura, altura, extensão) de PNG ou JPEG."""
    if dados[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", dados[16:24])
        return w, h, "png"
    if dados[:2] == b"\xff\xd8":
        i = 2
        while i < len(dados):
            if dados[i] != 0xFF:
                i += 1
                continue
            marc = dados[i + 1]
            if marc in (0xD8, 0x01) or 0xD0 <= marc <= 0xD7:
                i += 2
                continue
            tam = struct.unpack(">H", dados[i + 2:i + 4])[0]
            if marc in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack(">HH", dados[i + 5:i + 9])
                return w, h, "jpeg"
            i += 2 + tam
    raise ErroEspec("imagem não é PNG nem JPEG válido")


def ler_imagem(ref, base):
    """Aceita caminho (relativo a base) ou data URI base64."""
    if ref.startswith("data:"):
        dados = base64.b64decode(ref.split(",", 1)[1])
    else:
        caminho = ref if os.path.isabs(ref) else os.path.join(base, ref)
        with open(caminho, "rb") as f:
            dados = f.read()
    w, h, ext = dimensoes_imagem(dados)
    return dict(dados=dados, w=w, h=h, ext=ext)


# ---------------------------------------------------------------------------
# Marca
# ---------------------------------------------------------------------------
class Marca:
    def __init__(self, caminho):
        self.caminho = os.path.abspath(caminho)
        self.base = os.path.dirname(self.caminho)
        with open(caminho, encoding="utf-8") as f:
            self.cfg = json.load(f)
        c = self.cfg
        self.organizacao = c.get("organizacao") or ""
        self.politica = c.get("politica") or {}
        self.areas = dict(c["areas"]) if c.get("areas") else dict(AREAS_PADRAO)
        self.areas.update(c.get("areas_extra") or {})
        self.tipos = dict(TIPOS_PADRAO)
        for sig, d in (c.get("tipos_extra") or {}).items():
            if d.get("variante") not in ("curto", "longo"):
                raise ErroEspec(f"marca: tipo extra {sig} precisa de variante 'curto' ou 'longo'")
            self.tipos[sig.upper()] = (d.get("nome") or sig, d["variante"])
        emps = c.get("empresas") or []
        if not emps:
            raise ErroEspec("marca: lista 'empresas' vazia")
        self.empresas = {}
        for e in emps:
            sig = str(e.get("sigla", "")).upper()
            if not re.fullmatch(r"[A-Z0-9]{2,6}", sig):
                raise ErroEspec(f"marca: sigla inválida {sig!r} (2 a 6 letras/números)")
            cor = str(e.get("cor", "1F3864")).lstrip("#").upper()
            if not re.fullmatch(r"[0-9A-F]{6}", cor):
                raise ErroEspec(f"marca: cor inválida em {sig} (use hexadecimal, ex.: 1F3864)")
            e = dict(e, sigla=sig, cor=cor)
            e.setdefault("marca", sig)
            e.setdefault("documento_rotulo", "CNPJ")
            e["logo_img"] = None
            if e.get("logo"):
                try:
                    e["logo_img"] = ler_imagem(e["logo"], self.base)
                except (OSError, ValueError, ErroEspec) as ex:
                    raise ErroEspec(f"marca: logo de {sig} ilegível ({ex})")
            self.empresas[sig] = e
        self.padrao = str(c.get("empresa_padrao") or emps[0]["sigla"]).upper()

    @staticmethod
    def completa(e):
        return not vazio(e.get("razao_social")) and not vazio(e.get("documento"))

    def resumo(self):
        linhas = [f"Marca: {self.organizacao or '(sem nome de organização)'}"]
        for e in self.empresas.values():
            logo = "logo ok" if e["logo_img"] else "SEM LOGO (rodapé usa o nome)"
            st = "pode emitir OFICIAL" if self.completa(e) else "só não oficial (faltam razão social/documento)"
            linhas.append(f"  {e['sigla']}: {e['marca']} · #{e['cor']} · {logo} · {st}")
        linhas.append(f"  padrão: {self.padrao} · {len(self.areas)} áreas · {len(self.tipos)} tipos")
        return "\n".join(linhas)


def embutir_logos(caminho, saida):
    with open(caminho, encoding="utf-8") as f:
        cfg = json.load(f)
    base = os.path.dirname(os.path.abspath(caminho))
    for e in cfg.get("empresas", []):
        ref = e.get("logo")
        if ref and not ref.startswith("data:"):
            img = ler_imagem(ref, base)
            e["logo"] = f"data:image/{img['ext']};base64," + base64.b64encode(img["dados"]).decode()
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return saida


# ---------------------------------------------------------------------------
# Primitivas WordprocessingML
# ---------------------------------------------------------------------------
_MD = re.compile(r"(\*\*\*.+?\*\*\*|\*\*.+?\*\*|\*[^*\s][^*]*?\*)")


def rpr(neg=False, ita=False, tam=None, cor=None, caixa=False, mono=False):
    s = ""
    if mono:
        s += '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
    if neg:
        s += "<w:b/><w:bCs/>"
    if ita:
        s += "<w:i/><w:iCs/>"
    if caixa:
        s += "<w:caps/>"
    if cor:
        s += f'<w:color w:val="{cor}"/>'
    if tam:
        s += f'<w:sz w:val="{int(tam * 2)}"/><w:szCs w:val="{int(tam * 2)}"/>'
    return f"<w:rPr>{s}</w:rPr>" if s else ""


def run(texto, **kw):
    partes = str(texto).split("\n")
    corpo = "<w:br/>".join(f'<w:t xml:space="preserve">{esc(p)}</w:t>' for p in partes)
    corpo = corpo.replace("\t", '</w:t><w:tab/><w:t xml:space="preserve">')
    return f"<w:r>{rpr(**kw)}{corpo}</w:r>"


def runs_md(texto, **kw):
    out = []
    for pedaco in _MD.split(str(texto)):
        if not pedaco:
            continue
        k = dict(kw)
        if pedaco.startswith("***") and pedaco.endswith("***") and len(pedaco) > 6:
            k.update(neg=True, ita=True)
            pedaco = pedaco[3:-3]
        elif pedaco.startswith("**") and pedaco.endswith("**") and len(pedaco) > 4:
            k.update(neg=True)
            pedaco = pedaco[2:-2]
        elif pedaco.startswith("*") and pedaco.endswith("*") and len(pedaco) > 2:
            k.update(ita=True)
            pedaco = pedaco[1:-1]
        out.append(run(pedaco, **k))
    return "".join(out)


def ppr(estilo=None, alinhar=None, recuo_prim=None, esq=None, pend=None, antes=None,
        depois=None, entrelinha=None, manter=False, junto=False, quebra_antes=False,
        borda_sup=None, sombra=None, tabs=None, extra=""):
    s = ""
    if estilo:
        s += f'<w:pStyle w:val="{estilo}"/>'
    if manter:
        s += "<w:keepNext/>"
    if junto:
        s += "<w:keepLines/>"
    if quebra_antes:
        s += "<w:pageBreakBefore/>"
    if borda_sup:
        s += f'<w:pBdr><w:top w:val="single" w:sz="{borda_sup[1]}" w:space="4" w:color="{borda_sup[0]}"/></w:pBdr>'
    if sombra:
        s += f'<w:shd w:val="clear" w:color="auto" w:fill="{sombra}"/>'
    if tabs:
        s += "<w:tabs>" + "".join(f'<w:tab w:val="{v}" w:leader="{l}" w:pos="{p}"/>' for v, l, p in tabs) + "</w:tabs>"
    if antes is not None or depois is not None or entrelinha is not None:
        a = [] if antes is None else [f'w:before="{antes}"']
        a += [] if depois is None else [f'w:after="{depois}"']
        a += [] if entrelinha is None else [f'w:line="{entrelinha}" w:lineRule="auto"']
        s += f"<w:spacing {' '.join(a)}/>"
    if recuo_prim is not None or esq is not None or pend is not None:
        a = [f'w:left="{esq or 0}"']
        if pend:
            a.append(f'w:hanging="{pend}"')
        else:
            a.append(f'w:firstLine="{recuo_prim or 0}"')
        s += f"<w:ind {' '.join(a)}/>"
    if alinhar:
        s += f'<w:jc w:val="{ {"justify": "both", "center": "center", "right": "right", "left": "left"}[alinhar]}"/>'
    s += extra
    return f"<w:pPr>{s}</w:pPr>" if s else ""


def par(conteudo="", **kw):
    return f"<w:p>{ppr(**kw)}{conteudo}</w:p>"


def quebra_pagina():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def campo(instr, valor="1", **kw):
    return (f'<w:r>{rpr(**kw)}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{rpr(**kw)}<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r>{rpr(**kw)}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'{run(valor, **kw)}<w:r>{rpr(**kw)}<w:fldChar w:fldCharType="end"/></w:r>')


def borda(val="single", sz=4, cor="000000"):
    return f'w:val="{val}" w:sz="{sz}" w:space="0" w:color="{cor}"'


def bordas_xml(tag, b):
    """b: dict lado -> (val, sz, cor) ou None (nil)."""
    s = ""
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if lado not in b:
            continue
        v = b[lado]
        s += f'<w:{lado} w:val="nil"/>' if v is None else f"<w:{lado} {borda(*v)}/>"
    return f"<w:{tag}>{s}</w:{tag}>" if s else ""


def celula(conteudo, largura, sombra=None, valinhar=None, bordas=None, span=1, margens=None):
    tc = f'<w:tcW w:w="{largura}" w:type="dxa"/>'
    if span > 1:
        tc += f'<w:gridSpan w:val="{span}"/>'
    if bordas:
        tc += bordas_xml("tcBorders", bordas)
    if sombra:
        tc += f'<w:shd w:val="clear" w:color="auto" w:fill="{sombra}"/>'
    if margens:
        tc += ("<w:tcMar>" + "".join(f'<w:{k} w:w="{margens[k]}" w:type="dxa"/>'
                                     for k in ("top", "left", "bottom", "right") if k in margens)
               + "</w:tcMar>")
    if valinhar:
        tc += f'<w:vAlign w:val="{valinhar}"/>'
    if not conteudo.rstrip().endswith("</w:p>"):
        conteudo += "<w:p/>"
    return f"<w:tc><w:tcPr>{tc}</w:tcPr>{conteudo}</w:tc>"


def tabela(linhas, larguras, bordas=None, cabecalho=0, recuo=0, margens=(57, 85), alinhar=None):
    """linhas: lista de listas de strings <w:tc>."""
    total = sum(larguras)
    tp = (f'<w:tblW w:w="{total}" w:type="dxa"/>'
          + (f'<w:jc w:val="{alinhar}"/>' if alinhar else "")
          + (f'<w:tblInd w:w="{recuo}" w:type="dxa"/>' if recuo else "")
          + bordas_xml("tblBorders", bordas or {})
          + '<w:tblLayout w:type="fixed"/>'
          + f'<w:tblCellMar><w:top w:w="{margens[0]}" w:type="dxa"/><w:left w:w="{margens[1]}" w:type="dxa"/>'
          f'<w:bottom w:w="{margens[0]}" w:type="dxa"/><w:right w:w="{margens[1]}" w:type="dxa"/></w:tblCellMar>'
          '<w:tblLook w:val="0000"/>')
    grade = "".join(f'<w:gridCol w:w="{w}"/>' for w in larguras)
    corpo = ""
    for i, cels in enumerate(linhas):
        trp = "<w:cantSplit/>" + ("<w:tblHeader/>" if i < cabecalho else "")
        corpo += f"<w:tr><w:trPr>{trp}</w:trPr>{''.join(cels)}</w:tr>"
    return f"<w:tbl><w:tblPr>{tp}</w:tblPr><w:tblGrid>{grade}</w:tblGrid>{corpo}</w:tbl>"


def desenho(rid, cx, cy, did, nome):
    return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{did}" name="{esc(nome)}"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{did}" name="{esc(nome)}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


def marca_dagua(texto, cor):
    return (
        '<w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>'
        '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" '
        'path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas>'
        + "".join(f'<v:f eqn="{e}"/>' for e in (
            "sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3",
            "if @0 @3 0", "if @0 21600 @1", "if @0 0 @2", "if @0 @4 21600", "mid @5 @6",
            "mid @8 @5", "mid @7 @8", "mid @6 @7", "sum @6 0 @5"))
        + '</v:formulas><v:path textpathok="t" o:connecttype="custom" '
        'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
        '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
        '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'
        '<v:shape id="MarcaDagua" o:spid="_x0000_s2049" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:440pt;height:110pt;rotation:315;'
        'z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
        'mso-position-vertical:center;mso-position-vertical-relative:margin" '
        f'o:allowincell="f" fillcolor="#{cor}" stroked="f"><v:fill opacity=".45"/>'
        f'<v:textpath style="font-family:&quot;{FONTE}&quot;;font-size:1pt" string="{esc(texto)}"/>'
        '<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r></w:p>')


# ---------------------------------------------------------------------------
# Documento
# ---------------------------------------------------------------------------
class Documento:
    def __init__(self, spec, marca, base_spec):
        self.s = spec
        self.m = marca
        self.base = base_spec
        self.avisos = []
        self.midias = []          # (nome_arquivo, bytes)
        self.rels_doc = []        # (rid, tipo, alvo)
        self.did = 100
        self.cont = {"Tabela": 0, "Quadro": 0, "Figura": 0, "Gráfico": 0}
        self.secoes = [0] * 5
        self.nivel_ant = 0
        self.sumario = []         # (nivel, texto, bookmark)
        self.bm = 0
        self.paginas_sumario = None
        self.no_anexo = False
        self.avisos_montagem = []
        self._validar()

    # ---- validação e metadados -------------------------------------------
    def _validar(self):
        s, m = self.s, self.m
        sig = str(s.get("empresa") or "").upper()
        if not sig:
            sig = m.padrao
            self.avisos.append(f"empresa não informada: usada a padrão ({sig}).")
        if sig not in m.empresas:
            raise ErroEspec(f"empresa {sig!r} não existe no marca.json (disponíveis: {', '.join(m.empresas)})")
        self.emp = m.empresas[sig]

        self.tipo = str(s.get("tipo") or "").upper()
        if self.tipo not in m.tipos:
            raise ErroEspec(f"tipo {self.tipo!r} inválido (válidos: {', '.join(m.tipos)})")
        self.tipo_nome, variante = m.tipos[self.tipo]
        self.variante = s.get("variante") or variante
        if self.variante not in ("curto", "longo"):
            raise ErroEspec("variante deve ser 'curto' ou 'longo'")

        self.area = str(s.get("area") or "").upper()
        if self.area not in m.areas:
            raise ErroEspec(f"área {self.area!r} inválida (válidas: {', '.join(m.areas)})")
        self.area_nome = s.get("area_nome") or m.areas[self.area]

        self.estado = str(s.get("estado") or "RASCUNHO").upper().replace(" ", "_")
        if self.estado not in ESTADOS:
            raise ErroEspec(f"estado {self.estado!r} inválido (válidos: {', '.join(ESTADOS)})")
        self.est = ESTADOS[self.estado]

        self.data = s.get("data_emissao") or datetime.date.today().strftime("%d/%m/%Y")
        for campo_data in ("data_emissao", "data_aprovacao"):
            v = s.get(campo_data)
            if v and not re.fullmatch(r"\d{2}/\d{2}/\d{4}", v):
                self.avisos.append(f"{campo_data} fora do formato dd/mm/aaaa: {v}")
        mano = re.search(r"(\d{4})$", self.data)
        self.ano = str(s.get("ano") or (mano[1] if mano else datetime.date.today().year))

        num = s.get("numero")
        self.tem_numero = not (num in (None, "") or str(num).upper().startswith("X"))
        if self.tem_numero:
            try:
                num = f"{int(num):03d}"
            except ValueError:
                raise ErroEspec(f"número inválido: {num!r}")
        else:
            num = "XXX"
        self.codigo = f"{self.emp['sigla']}-{self.tipo}-{self.area}-{num}"
        if self.variante == "curto":
            self.codigo += f"/{self.ano}"

        erros = []
        oficial = self.estado in ("OFICIAL", "OBSOLETO")
        if oficial:
            if not Marca.completa(self.emp):
                erros.append(f"a empresa {self.emp['sigla']} está sem razão social/documento no marca.json; "
                             "só pode emitir RASCUNHO, MINUTA ou EM_REVISAO")
            if not self.tem_numero:
                erros.append("documento oficial precisa do número reservado na lista mestra")
        if self.estado == "OFICIAL":
            if vazio(s.get("aprovado_por")):
                erros.append("OFICIAL exige 'aprovado_por'")
            if vazio(s.get("data_aprovacao")):
                erros.append("OFICIAL exige 'data_aprovacao'")
        if self.estado == "OBSOLETO" and vazio(s.get("substituido_por")) and vazio(s.get("motivo_obsoleto")):
            erros.append("OBSOLETO exige 'substituido_por' ou 'motivo_obsoleto'")
        if erros:
            raise ErroEspec("; ".join(erros) + ". Rebaixe o estado ou informe o dado.")

        v = str(s.get("versao") or "1.0").strip().lstrip("vV")
        mv = re.fullmatch(r"(\d+)\.(\d+)(?:-r(\d+))?", v)
        if not mv:
            raise ErroEspec(f"versão inválida {v!r} (use X.Y ou X.Y-rN)")
        if oficial:
            if mv[3]:
                self.avisos.append(f"sufixo -r{mv[3]} removido: documento {self.est['nome'].lower()} usa só X.Y.")
            v = f"{mv[1]}.{mv[2]}"
        elif not mv[3]:
            v = f"{mv[1]}.{mv[2]}-r1"
            self.avisos.append(f"versão sem rodada em estado não oficial: usada {v}.")
        self.versao = v

        cl = sem_acento(s.get("classificacao") or "Interno").lower()
        self.classificacao = next((c for c in CLASSIFICACOES if sem_acento(c).lower() == cl), None)
        if not self.classificacao:
            self.classificacao = "Interno"
            self.avisos.append("classificação desconhecida: usada 'Interno'.")

        self.titulo = (s.get("titulo") or "").strip()
        if not self.titulo:
            raise ErroEspec("falta 'titulo'")
        self.titulo_curto = s.get("titulo_curto") or (self.titulo if len(self.titulo) <= 60 else self.titulo[:57] + "…")

        if not self.tem_numero:
            self.avisos.append("número não reservado: código com XXX.")
        if not Marca.completa(self.emp):
            self.avisos.append(f"{self.emp['sigla']} sem razão social/documento cadastrados: aparece {INDEFINIDO} no cabeçalho.")
        if not self.emp["logo_img"]:
            self.avisos.append(f"{self.emp['sigla']} sem logo: o rodapé mostra o nome da marca.")
        pend = json.dumps(s, ensure_ascii=False).count(INDEFINIDO)
        if pend:
            self.avisos.append(f"{pend} ocorrência(s) de {INDEFINIDO} no conteúdo.")
        if self.variante == "curto":
            for k in ("historico", "referencias", "glossario", "apendices", "anexos"):
                if s.get(k):
                    self.avisos.append(f"'{k}' ignorado: só existe na variante longa.")
        elif s.get("identificacao"):
            self.avisos.append("'identificacao' ignorado: só existe na variante curta (use a folha de controle).")

    def nome_arquivo(self):
        return f"{self.codigo.replace('/', '-')}_v{self.versao}_{self.estado}_{slug(self.titulo)}.docx"

    # ---- mídia -------------------------------------------------------------
    def _nova_midia(self, img, prefixo):
        nome = f"{prefixo}{len(self.midias) + 1}.{'png' if img['ext'] == 'png' else 'jpeg'}"
        self.midias.append((nome, img["dados"]))
        return nome

    def _prox_did(self):
        self.did += 1
        return self.did

    def _logo_xml(self, rid, alt_cm=0.85, larg_max_cm=5.0):
        img = self.emp["logo_img"]
        cy = alt_cm * EMU_CM
        cx = cy * img["w"] / img["h"]
        if cx > larg_max_cm * EMU_CM:
            cx = larg_max_cm * EMU_CM
            cy = cx * img["h"] / img["w"]
        return desenho(rid, int(cx), int(cy), self._prox_did(), f"Logo {self.emp['marca']}")

    def _logo_ou_nome(self, rid, alinhar):
        if self.emp["logo_img"]:
            return self._logo_xml(rid)
        return run(self.emp["marca"], neg=True, tam=11, cor=self.emp["cor"])

    # ---- cabeçalho e rodapé -----------------------------------------------
    def cabecalho(self):
        s, e, est = self.s, self.emp, self.est
        cinza = ("single", 4, "BFBFBF")
        faixa_txt = est["faixa"]
        if self.estado == "OBSOLETO":
            faixa_txt += f" · SUBSTITUÍDO POR {s['substituido_por']}" if not vazio(s.get("substituido_por")) else ""
        faixa = celula(par(run(faixa_txt, neg=True, tam=7.5, cor=est["txt"]), estilo="Header", alinhar="center"),
                       LARG_TEXTO, sombra=est["cor"], span=2, valinhar="center")
        l_esq, l_dir = LARG_TEXTO - 3 * CM, 3 * CM
        cod = celula(par(run(self.codigo, neg=True, tam=8) + run(f"  ·  Versão {self.versao}", tam=8),
                         estilo="Header"), l_esq)
        pag = celula(par(run("Página ", tam=8) + campo("PAGE", "1", tam=8) + run(" de ", tam=8)
                         + campo("NUMPAGES", "1", tam=8), estilo="Header", alinhar="right"), l_dir)
        if self.estado == "OFICIAL":
            cargo = f" ({s['cargo_aprovador']})" if not vazio(s.get("cargo_aprovador")) else ""
            vig = s.get("vigencia") or "a partir da aprovação"
            sit = (f"Situação: Oficial · Aprovado por {s['aprovado_por']}{cargo} em {s['data_aprovacao']}"
                   f" · Vigência: {vig} · {self.classificacao}")
        elif self.estado == "OBSOLETO":
            motivo = f" · Motivo: {s['motivo_obsoleto']}" if not vazio(s.get("motivo_obsoleto")) else ""
            sit = f"Situação: Obsoleto{motivo} · {self.classificacao}"
        else:
            sit = f"Situação: {est['nome']} · Emissão: {self.data} · {self.classificacao}"
        situacao = celula(par(run(sit, tam=8), estilo="Header"), LARG_TEXTO, span=2)
        rs = e.get("razao_social") or INDEFINIDO
        doc = e.get("documento") or INDEFINIDO
        id_txt = " · ".join(x for x in (rs, f"{e['documento_rotulo']} {doc}", e.get("cidade")) if x)
        ident = celula(par(run(id_txt, tam=8, cor="404040"), estilo="Header"), LARG_TEXTO, span=2)
        tbl = tabela([[faixa], [cod, pag], [situacao], [ident]], [l_esq, l_dir],
                     bordas={"top": None, "left": None, "right": None,
                             "bottom": ("single", 8, e["cor"]), "insideH": cinza, "insideV": None},
                     margens=(20, 57))
        wm = marca_dagua(est["marca"], tom(est["cor"], 0.65)) if est["marca"] else ""
        return f'<w:hdr {NS}>{wm}{tbl}{par(estilo="Header", depois=0)}</w:hdr>'

    def rodape_centro(self, rid):
        conteudo = self._logo_ou_nome(rid, "center")
        return (f'<w:ftr {NS}>' + par(conteudo, estilo="Footer", alinhar="center",
                                      borda_sup=(self.emp["cor"], 8)) + "</w:ftr>")

    def rodape_interno(self, rid):
        e = self.emp
        l1, l2 = 5 * CM, LARG_TEXTO - 5 * CM
        c1 = celula(par(self._logo_ou_nome(rid, "left"), estilo="Footer"), l1, valinhar="center")
        c2 = celula(par(run(self.titulo_curto, tam=8, cor="404040"), estilo="Footer", alinhar="right"),
                    l2, valinhar="center")
        tbl = tabela([[c1, c2]], [l1, l2], bordas={"top": ("single", 8, e["cor"]), "left": None,
                                                   "bottom": None, "right": None, "insideH": None,
                                                   "insideV": None}, margens=(57, 0))
        return f'<w:ftr {NS}>{tbl}{par(estilo="Footer", depois=0, entrelinha=120)}</w:ftr>'

    # ---- blocos de corpo ---------------------------------------------------
    def titulo_secao(self, nivel, texto, numerar=True, nova_pagina=False, sumario=True):
        if numerar and nivel > self.nivel_ant + 1:
            self.avisos.append(f"título h{nivel} '{texto[:40]}' pula nível (anterior h{self.nivel_ant}).")
        num = ""
        if numerar:
            self.secoes[nivel - 1] += 1
            for i in range(nivel, 5):
                self.secoes[i] = 0
            num = ".".join(str(x) for x in self.secoes[:nivel]) + " "
            self.nivel_ant = nivel
        exib = texto.upper() if nivel <= 2 else texto
        completo = (num + exib).strip()
        conteudo = runs_md(completo)
        if sumario and nivel <= 3:
            self.bm += 1
            nome = f"_Toc_t{self.bm}"
            conteudo = f'<w:bookmarkStart w:id="{self.bm}" w:name="{nome}"/>{conteudo}<w:bookmarkEnd w:id="{self.bm}"/>'
            self.sumario.append((nivel, re.sub(r"\*", "", completo), nome))
        return par(conteudo, estilo=f"Heading{nivel}", quebra_antes=nova_pagina)

    def legenda(self, tipo, titulo):
        self.cont[tipo] = self.cont.get(tipo, 0) + 1
        return par(runs_md(f"**{tipo} {self.cont[tipo]}** – {titulo}"), estilo="Legenda")

    def fonte_nota(self, fonte, nota):
        out = ""
        if fonte:
            out += par(runs_md(f"Fonte: {fonte}"), estilo="FonteTabela")
        else:
            self.avisos.append("tabela/figura sem 'fonte'.")
        if nota:
            out += par(runs_md(f"Nota: {nota}"), estilo="FonteTabela")
        return out

    def bloco_tabela(self, b, interno=False):
        tipo = "Quadro" if str(b.get("tipo", "Tabela")).lower() == "quadro" else "Tabela"
        cab = b.get("cabecalho") or []
        linhas = b.get("linhas") or []
        ncol = max([len(cab)] + [len(l) for l in linhas] or [1])
        pesos = b.get("larguras") or [1] * ncol
        pesos = (list(pesos) + [1] * ncol)[:ncol]
        larg = [int(LARG_TEXTO * p / sum(pesos)) for p in pesos]
        alin = (list(b.get("alinhamento") or []) + ["left"] * ncol)[:ncol]
        linha = ("single", 8, "000000")
        fina = ("single", 4, "808080")
        if tipo == "Tabela":
            bd = {"top": linha, "bottom": linha, "left": None, "right": None, "insideH": None, "insideV": None}
        else:
            bd = {"top": linha, "bottom": linha, "left": linha, "right": linha, "insideH": fina, "insideV": fina}
        rows = []
        if cab:
            rows.append([celula(par(runs_md(c, neg=True), estilo="TabelaTexto", alinhar=alin[i] if alin[i] != "left" else None),
                                larg[i], sombra="F2F2F2" if tipo == "Quadro" else None, valinhar="center",
                                bordas={"bottom": linha})
                         for i, c in enumerate((cab + [""] * ncol)[:ncol])])
        for l in linhas:
            l = (list(l) + [""] * ncol)[:ncol]
            rows.append([celula("".join(par(runs_md(t), estilo="TabelaTexto",
                                            alinhar=alin[i] if alin[i] != "left" else None)
                                        for t in str(c).split("\n\n")), larg[i])
                         for i, c in enumerate(l)])
        tbl = tabela(rows, larg, bordas=bd, cabecalho=1 if cab else 0, alinhar="center")
        if interno:
            return tbl
        titulo = b.get("titulo")
        out = self.legenda(tipo, titulo) if titulo else ""
        if not titulo:
            self.avisos.append(f"{tipo.lower()} sem 'titulo'.")
        return out + tbl + self.fonte_nota(b.get("fonte"), b.get("nota"))

    def bloco_figura(self, b):
        tipo = "Gráfico" if sem_acento(b.get("tipo", "")).lower() == "grafico" else "Figura"
        out = self.legenda(tipo, b.get("titulo") or "")
        try:
            img = ler_imagem(b["caminho"], self.base)
        except (KeyError, OSError, ValueError, ErroEspec) as ex:
            self.avisos.append(f"figura '{b.get('caminho')}' não carregada ({ex}).")
            return out + par(run(f"[imagem não encontrada: {b.get('caminho')}]", ita=True), alinhar="center") \
                + self.fonte_nota(b.get("fonte"), b.get("nota"))
        nome = self._nova_midia(img, "fig")
        rid = f"rIdImg{len(self.midias)}"
        self.rels_doc.append((rid, "image", f"media/{nome}"))
        larg = min(float(b.get("largura_cm") or 14), 16) * EMU_CM
        alt = larg * img["h"] / img["w"]
        if alt > 20 * EMU_CM:
            alt = 20 * EMU_CM
            larg = alt * img["w"] / img["h"]
        out += par(desenho(rid, int(larg), int(alt), self._prox_did(), b.get("titulo") or nome),
                   alinhar="center", manter=True, antes=60, depois=60, entrelinha=240, recuo_prim=0)
        return out + self.fonte_nota(b.get("fonte"), b.get("nota"))

    def bloco_lista(self, b, modo):
        itens = b.get("itens") or []
        out = []
        n = 0
        for i, it in enumerate(itens):
            texto, subs = (it, []) if isinstance(it, str) else (it.get("texto", ""), it.get("sub") or [])
            if modo == "alineas":
                letras, k = "", i
                while True:
                    letras = chr(97 + k % 26) + letras
                    k = k // 26 - 1
                    if k < 0:
                        break
                pref = f"{letras})"
            elif modo == "numerada":
                n += 1
                pref = f"{n}."
            elif modo == "checklist":
                pref = "☐"
            else:
                pref = "•"
            out.append(par(run(pref) + "<w:r><w:tab/></w:r>" + runs_md(texto), estilo="Alinea",
                           tabs=[("left", "none", 1276)]))
            for sub in subs:
                out.append(par(run("–") + "<w:r><w:tab/></w:r>" + runs_md(sub), estilo="Alinea",
                               esq=1701, pend=425, tabs=[("left", "none", 1701)]))
        if modo == "alineas":
            textos = [x if isinstance(x, str) else x.get("texto", "") for x in itens]
            if textos and (any(t[:1].isupper() for t in textos) or not all(t.rstrip().endswith(";") for t in textos[:-1])
                           or not textos[-1].rstrip().endswith(".")):
                self.avisos.append("alíneas fora da ABNT (minúscula inicial, ';' ao fim e '.' na última).")
        return "".join(out)

    def bloco_destaque(self, b):
        cor, fundo = NIVEIS_DESTAQUE.get(b.get("nivel", "info"), NIVEIS_DESTAQUE["info"])
        rot = b.get("rotulo") or {"info": "Importante", "atencao": "Atenção", "erro": "Proibido"}.get(b.get("nivel"), "Importante")
        cont = par(run(rot.upper(), neg=True, tam=9, cor=cor), estilo="TabelaTexto", depois=40)
        cont += "".join(par(runs_md(t), estilo="TabelaTexto") for t in str(b.get("texto", "")).split("\n\n"))
        c = celula(cont, LARG_TEXTO, sombra=fundo,
                   bordas={"left": ("single", 24, cor), "top": None, "bottom": None, "right": None},
                   margens={"left": 170, "right": 170, "top": 85, "bottom": 85})
        return tabela([[c]], [LARG_TEXTO]) + par(depois=0, entrelinha=240)

    def bloco_campos(self, b):
        rot = int(float(b.get("rotulo_cm") or 4) * CM)
        linha = ("single", 4, "BFBFBF")
        rows = [[celula(par(runs_md(r, neg=True), estilo="TabelaTexto"), rot),
                 celula("".join(par(runs_md(t), estilo="TabelaTexto") for t in str(v).split("\n\n")), LARG_TEXTO - rot)]
                for r, v in (b.get("itens") or [])]
        if not rows:
            return ""
        return tabela(rows, [rot, LARG_TEXTO - rot],
                      bordas={"top": linha, "bottom": linha, "insideH": linha, "left": None, "right": None,
                              "insideV": None}) + par(depois=0, entrelinha=240)

    def bloco_assinaturas(self, b):
        itens = b.get("itens") or []
        if not itens:
            return ""
        meia = LARG_TEXTO // 2
        cels = []
        for it in itens:
            cont = par(estilo="TabelaTexto", antes=720, alinhar="center")
            cont += par(runs_md(it.get("nome") or "_" * 30, neg=True), estilo="TabelaTexto", alinhar="center",
                        borda_sup=("000000", 6), manter=True)
            for k in ("cargo", "documento"):
                if it.get(k):
                    cont += par(runs_md(it[k]), estilo="TabelaTexto", alinhar="center", manter=True)
            cels.append(celula(cont, meia, margens={"left": 340, "right": 340}))
        if len(itens) == 1:
            rows = [[celula("", LARG_TEXTO // 4), cels[0], celula("", LARG_TEXTO - meia - LARG_TEXTO // 4)]]
            larg = [LARG_TEXTO // 4, meia, LARG_TEXTO - meia - LARG_TEXTO // 4]
        else:
            if len(cels) % 2:
                cels.append(celula("", meia))
            rows = [cels[i:i + 2] for i in range(0, len(cels), 2)]
            larg = [meia, meia]
        nil = {k: None for k in ("top", "left", "bottom", "right", "insideH", "insideV")}
        return tabela(rows, larg, bordas=nil) + par(depois=0)

    def bloco(self, b):
        t = str(b.get("t", "p")).lower()
        longo = self.variante == "longo"
        if re.fullmatch(r"h[1-5]", t):
            nivel = int(t[1])
            nova = b.get("nova_pagina", longo and nivel == 1 and not self.no_anexo)
            return self.titulo_secao(nivel, b.get("texto", ""), numerar=b.get("numerar", not self.no_anexo),
                                     nova_pagina=nova, sumario=not self.no_anexo)
        if t == "p":
            out = ""
            for texto in str(b.get("texto", "")).split("\n\n"):
                out += par(runs_md(texto), recuo_prim=0 if b.get("sem_recuo") or b.get("alinhar") else None,
                           alinhar=b.get("alinhar"))
            return out
        if t in ("alineas", "numerada", "checklist", "marcadores"):
            modo = t
            if t == "alineas":
                modo = next((k for k in ("numerada", "checklist", "marcadores") if b.get(k)), "alineas")
            return self.bloco_lista(b, modo)
        if t == "tabela":
            return self.bloco_tabela(b)
        if t == "figura":
            return self.bloco_figura(b)
        if t == "citacao":
            texto = str(b.get("texto", "")) + (f" {b['fonte']}" if b.get("fonte") else "")
            return par(runs_md(texto), estilo="Citacao")
        if t == "nota":
            return par(runs_md(f"**{b.get('rotulo') or 'Nota'}:** {b.get('texto', '')}"), estilo="NotaTexto")
        if t == "destaque":
            return self.bloco_destaque(b)
        if t == "campos":
            return self.bloco_campos(b)
        if t == "local_data":
            cidade = self.emp.get("cidade") or INDEFINIDO
            return par(run(f"{cidade}, {data_extenso(b.get('data') or self.data)}."), alinhar="right",
                       recuo_prim=0, antes=240, depois=240)
        if t == "assinaturas":
            return self.bloco_assinaturas(b)
        if t == "codigo":
            return "".join(par(run(l or " ", mono=True, tam=9), estilo="Codigo")
                           for l in str(b.get("texto", "")).split("\n"))
        if t == "quebra":
            return quebra_pagina()
        self.avisos.append(f"bloco desconhecido ignorado: {t!r}.")
        return ""

    def corpo_blocos(self, blocos):
        return "".join(self.bloco(b) for b in blocos or [])

    # ---- partes do documento ---------------------------------------------
    def titulo_sem_indicativo(self, texto, quebra=False, sumario=True):
        """Títulos sem número (Referências, Glossário, Apêndice...)."""
        if sumario:
            return self.titulo_secao(1, texto, numerar=False, nova_pagina=quebra)
        return par(run(texto.upper(), neg=True), estilo="TituloCentral", quebra_antes=quebra)

    def partes_curto(self):
        s = self.s
        out = par(run(self.tipo_nome, tam=9, cor=self.emp["cor"], caixa=True, neg=True),
                  alinhar="center", recuo_prim=0, depois=0)
        out += par(runs_md(self.titulo.upper(), neg=True, tam=14), alinhar="center", recuo_prim=0,
                   depois=60 if s.get("subtitulo") else 240, manter=True)
        if s.get("subtitulo"):
            out += par(runs_md(s["subtitulo"], tam=12), alinhar="center", recuo_prim=0, depois=240)
        ident = s.get("identificacao") or []
        if ident:
            out += self.bloco_campos({"itens": ident, "rotulo_cm": 3.5})
        self.no_anexo = False
        return out + self.corpo_blocos(s.get("corpo"))

    def capa(self):
        s, cor = self.s, self.emp["cor"]
        out = par(run(self.tipo_nome, neg=True, tam=12, cor=cor, caixa=True), alinhar="center",
                  recuo_prim=0, antes=2400, depois=240)
        out += par(runs_md(self.titulo.upper(), neg=True, tam=20), alinhar="center", recuo_prim=0,
                   depois=240, entrelinha=276)
        if s.get("subtitulo"):
            out += par(runs_md(s["subtitulo"], tam=14), alinhar="center", recuo_prim=0, depois=240)
        out += par(run(f"{self.codigo}  ·  Versão {self.versao}", tam=11, cor="404040"),
                   alinhar="center", recuo_prim=0, antes=960, depois=60)
        out += par(run(f"{self.est['nome']}  ·  {self.area_nome}  ·  {self.classificacao}", tam=11, cor="404040"),
                   alinhar="center", recuo_prim=0, depois=0)
        cidade = self.emp.get("cidade") or ""
        out += par(run(cidade, tam=12), alinhar="center", recuo_prim=0, antes=5200, depois=0)
        out += par(run(self.ano, tam=12), alinhar="center", recuo_prim=0, depois=0)
        return out

    def folha_controle(self):
        s = self.s
        out = self.titulo_sem_indicativo("Folha de controle", quebra=True, sumario=False)
        traco = lambda k: s.get(k) or "—"
        itens = [["Código", self.codigo], ["Título", self.titulo], ["Tipo", self.tipo_nome],
                 ["Área", self.area_nome], ["Versão", self.versao], ["Situação", self.est["nome"]],
                 ["Classificação", self.classificacao], ["Data de emissão", self.data],
                 ["Elaborado por", traco("elaborado_por")], ["Revisado por", traco("revisado_por")]]
        if self.estado == "OFICIAL" or not vazio(s.get("aprovado_por")):
            cargo = f" – {s['cargo_aprovador']}" if s.get("cargo_aprovador") else ""
            itens.append(["Aprovado por", f"{traco('aprovado_por')}{cargo}"])
            itens.append(["Data de aprovação", traco("data_aprovacao")])
        else:
            itens.append(["Aprovado por", "— (pendente)"])
        for rot, k in (("Vigência", "vigencia"), ("Substitui", "substitui"),
                       ("Substituído por", "substituido_por"), ("Motivo da obsolescência", "motivo_obsoleto")):
            if s.get(k):
                itens.append([rot, s[k]])
        pol = self.m.politica
        if pol.get("codigo"):
            itens.append(["Norma de referência", f"{pol['codigo']} – {pol.get('titulo', 'Padrão de Documentação Oficial')}"])
        out += self.bloco_campos({"itens": itens, "rotulo_cm": 4.5})
        hist = s.get("historico") or [{"versao": self.versao, "data": self.data, "descricao": "Emissão inicial.",
                                       "responsavel": s.get("elaborado_por") or "—"}]
        out += par(run("Histórico de versões", neg=True), recuo_prim=0, antes=240, manter=True)
        linhas = [[h.get("versao", ""), h.get("data", ""), h.get("descricao", ""), h.get("responsavel", "")] for h in hist]
        out += self.bloco_tabela({"tipo": "Quadro", "cabecalho": ["Versão", "Data", "Descrição", "Responsável"],
                                  "linhas": linhas, "larguras": [2, 2.2, 7, 4]}, interno=True)
        if self.estado == "OFICIAL":
            out += par(run("Aprovação", neg=True), recuo_prim=0, antes=360, manter=True)
            out += self.bloco_assinaturas({"itens": [{"nome": s["aprovado_por"], "cargo": s.get("cargo_aprovador", ""),
                                                      "documento": f"Aprovado em {s['data_aprovacao']}"}]})
        return out

    def sumario_xml(self):
        out = self.titulo_sem_indicativo("Sumário", quebra=True, sumario=False)
        entradas = self.sumario
        if not entradas:
            return out + par(run("Sem seções.", ita=True), recuo_prim=0)
        pags = self.paginas_sumario or {}
        for i, (nivel, texto, bm) in enumerate(entradas):
            ini = ""
            if i == 0:
                ini = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                       '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                       '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
            pag = str(pags.get(bm, "")) if self.paginas_sumario is not None else "00"
            link = (f'<w:hyperlink w:anchor="{bm}" w:history="1">{run(texto, neg=nivel == 1)}'
                    f'<w:r><w:tab/></w:r>{run(pag, neg=nivel == 1)}</w:hyperlink>')
            out += par(ini + link, estilo=f"TOC{nivel}", tabs=[("right", "dot", LARG_TEXTO)])
        fim = "" if self.paginas_sumario is not None else run("§§FIM-SUMARIO§§", tam=6)
        return out + par(fim + '<w:r><w:fldChar w:fldCharType="end"/></w:r>', recuo_prim=0)

    def partes_longo(self):
        s = self.s
        self.no_anexo = False
        # o corpo é montado antes do sumário para coletar as entradas
        blocos = list(s.get("corpo") or [])
        corpo = ""
        if blocos and not (str(blocos[0].get("t")).lower() == "h1" and blocos[0].get("nova_pagina", True)):
            corpo += quebra_pagina()
        if blocos and str(blocos[0].get("t")).lower() == "h1":
            primeiro = blocos[0]
            if sem_acento(primeiro.get("texto", "")).lower() not in ("objetivo", "introducao"):
                self.avisos.append("documento longo deve começar por 'Objetivo' ou 'Introdução'.")
        corpo += self.corpo_blocos(blocos)
        refs = s.get("referencias") or []
        if refs:
            corpo += self.titulo_sem_indicativo("Referências", quebra=True)
            corpo += "".join(par(runs_md(r), estilo="Referencia") for r in sorted(refs, key=lambda x: sem_acento(x).lower()))
        glos = s.get("glossario") or []
        if glos:
            corpo += self.titulo_sem_indicativo("Glossário", quebra=True)
            for termo, defin in sorted(glos, key=lambda x: sem_acento(x[0]).lower()):
                corpo += par(runs_md(f"**{termo}**: {defin}"), estilo="Referencia")
        self.no_anexo = True
        for rot, chave in (("Apêndice", "apendices"), ("Anexo", "anexos")):
            for i, ap in enumerate(s.get(chave) or []):
                corpo += self.titulo_sem_indicativo(f"{rot} {chr(65 + i)} – {ap.get('titulo', '')}", quebra=True)
                corpo += self.corpo_blocos(ap.get("corpo"))
        self.no_anexo = False
        pre = self.capa() + self.folha_controle()
        if s.get("sumario", True):
            pre += self.sumario_xml()
        return pre + corpo

    # ---- pacote ------------------------------------------------------------
    def montar(self):
        self.midias, self.rels_doc, self.did = [], [], 100
        self.cont = {k: 0 for k in self.cont}
        self.secoes, self.nivel_ant, self.sumario, self.bm = [0] * 5, 0, [], 0
        avisos_base = list(self.avisos)
        corpo = self.partes_longo() if self.variante == "longo" else self.partes_curto()
        self.avisos_montagem = [a for a in self.avisos if a not in avisos_base]
        self.avisos = avisos_base

        logo_nome = None
        if self.emp["logo_img"]:
            logo_nome = self._nova_midia(self.emp["logo_img"], "logo")
        cor = self.emp["cor"]
        borda_pag = "".join(f'<w:{l} w:val="single" w:sz="6" w:space="20" w:color="{cor}"/>'
                            for l in ("top", "left", "bottom", "right"))
        longo = self.variante == "longo"
        ref_rod = ('<w:footerReference w:type="default" r:id="rIdFtrInt"/>'
                   '<w:footerReference w:type="first" r:id="rIdFtrCen"/>') if longo else \
            '<w:footerReference w:type="default" r:id="rIdFtrCen"/>'
        ref_cab = '<w:headerReference w:type="default" r:id="rIdHdr"/>' + \
                  ('<w:headerReference w:type="first" r:id="rIdHdr"/>' if longo else "")
        sect = (f'<w:sectPr>{ref_cab}{ref_rod}<w:pgSz w:w="{PAG_L}" w:h="{PAG_A}"/>'
                f'<w:pgMar w:top="{MARG_SUP}" w:right="{MARG_DIR}" w:bottom="{MARG_INF}" w:left="{MARG_ESQ}" '
                f'w:header="510" w:footer="454" w:gutter="0"/>'
                f'<w:pgBorders w:offsetFrom="page">{borda_pag}</w:pgBorders>'
                + ("<w:titlePg/>" if longo else "") + "</w:sectPr>")
        documento = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>{corpo}{sect}</w:body></w:document>'

        rid_logo = "rIdLogo"
        ftr_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + (f'<Relationship Id="{rid_logo}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{logo_nome}"/>' if logo_nome else "")
                    + "</Relationships>")
        R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        rels = [("rIdStyles", "styles", "styles.xml"), ("rIdSettings", "settings", "settings.xml"),
                ("rIdHdr", "header", "header1.xml"), ("rIdFtrCen", "footer", "footer1.xml")]
        if longo:
            rels.append(("rIdFtrInt", "footer", "footer2.xml"))
        rels += self.rels_doc
        doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + "".join(f'<Relationship Id="{i}" Type="{R}{t}" Target="{a}"/>' for i, t, a in rels)
                    + "</Relationships>")
        xmlh = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        arquivos = {
            "[Content_Types].xml": xmlh + (
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="png" ContentType="image/png"/>'
                '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                + ('<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' if longo else "")
                + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                '</Types>'),
            "_rels/.rels": xmlh + (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                '</Relationships>'),
            "docProps/core.xml": xmlh + (
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dc:title>{esc(self.titulo)}</dc:title><dc:subject>{esc(self.codigo)} v{esc(self.versao)}</dc:subject>'
                f'<dc:creator>{esc(self.emp["marca"])}</dc:creator>'
                f'<cp:keywords>{esc(self.estado)}; {esc(self.classificacao)}</cp:keywords>'
                f'<cp:contentStatus>{esc(self.est["nome"])}</cp:contentStatus>'
                f'<dcterms:created xsi:type="dcterms:W3CDTF">{datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}</dcterms:created>'
                '</cp:coreProperties>'),
            "docProps/app.xml": xmlh + (
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                f'<Application>Timbrador</Application><Company>{esc(self.emp.get("razao_social") or self.emp["marca"])}</Company></Properties>'),
            "word/document.xml": documento,
            "word/_rels/document.xml.rels": doc_rels,
            "word/styles.xml": estilos(cor),
            "word/settings.xml": configuracoes(atualizar=self.variante == "longo" and self.paginas_sumario is None
                                              and self.s.get("sumario", True)),
            "word/header1.xml": xmlh + self.cabecalho(),
            "word/footer1.xml": xmlh + self.rodape_centro(rid_logo),
            "word/_rels/footer1.xml.rels": ftr_rels,
        }
        if longo:
            arquivos["word/footer2.xml"] = xmlh + self.rodape_interno(rid_logo)
            arquivos["word/_rels/footer2.xml.rels"] = ftr_rels
        for nome, dados in self.midias:
            arquivos[f"word/media/{nome}"] = dados
        return arquivos

    def gravar(self, caminho):
        arquivos = self.montar()
        with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", arquivos.pop("[Content_Types].xml"))
            for nome, dados in arquivos.items():
                z.writestr(nome, dados)


def estilos(cor):
    f = FONTE

    def st(sid, nome, ppr_="", rpr_="", tipo="paragraph", base="Normal", prox=None, extra=""):
        b = f'<w:basedOn w:val="{base}"/>' if base else ""
        n = f'<w:next w:val="{prox}"/>' if prox else ""
        return (f'<w:style w:type="{tipo}" w:styleId="{sid}"><w:name w:val="{nome}"/>{b}{n}{extra}'
                f'<w:qFormat/><w:pPr>{ppr_}</w:pPr><w:rPr>{rpr_}</w:rPr></w:style>')

    sem_recuo = '<w:ind w:left="0" w:firstLine="0"/>'
    simples = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    tit = '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>'
    heads = [
        st("Heading1", "heading 1", tit + '<w:outlineLvl w:val="0"/>', f'<w:b/><w:bCs/><w:color w:val="{cor}"/>', prox="Normal"),
        st("Heading2", "heading 2", tit + '<w:outlineLvl w:val="1"/>', '<w:color w:val="000000"/>', prox="Normal"),
        st("Heading3", "heading 3", tit + '<w:outlineLvl w:val="2"/>', '<w:b/><w:bCs/>', prox="Normal"),
        st("Heading4", "heading 4", tit + '<w:outlineLvl w:val="3"/>', '', prox="Normal"),
        st("Heading5", "heading 5", tit + '<w:outlineLvl w:val="4"/>', '<w:i/><w:iCs/>', prox="Normal"),
    ]
    corpo = "".join([
        st("Normal", "Normal", '<w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/>'
           '<w:ind w:firstLine="709"/><w:jc w:val="both"/>', "", base=None),
        *heads,
        st("TituloCentral", "Título sem indicativo", '<w:keepNext/><w:spacing w:before="0" w:after="360"/>'
           + sem_recuo + '<w:jc w:val="center"/>', '<w:b/><w:bCs/>'),
        st("TOC1", "toc 1", '<w:spacing w:before="60" w:after="0" w:line="300" w:lineRule="auto"/>' + sem_recuo + '<w:jc w:val="left"/>', ""),
        st("TOC2", "toc 2", '<w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>', ""),
        st("TOC3", "toc 3", '<w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>', ""),
        st("Header", "header", simples + sem_recuo + '<w:jc w:val="left"/>', ""),
        st("Footer", "footer", simples + sem_recuo + '<w:jc w:val="left"/>', ""),
        st("Citacao", "Citação longa", '<w:spacing w:before="240" w:after="240" w:line="240" w:lineRule="auto"/>'
           '<w:ind w:left="2268" w:firstLine="0"/><w:jc w:val="both"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("Legenda", "caption", '<w:keepNext/><w:spacing w:before="240" w:after="60" w:line="240" w:lineRule="auto"/>'
           + sem_recuo + '<w:jc w:val="center"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("FonteTabela", "Fonte da tabela", '<w:spacing w:before="40" w:after="0" w:line="240" w:lineRule="auto"/>'
           + sem_recuo + '<w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("TabelaTexto", "Texto de tabela", simples + sem_recuo + '<w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("NotaTexto", "Nota", '<w:spacing w:before="60" w:after="120" w:line="240" w:lineRule="auto"/>'
           + sem_recuo, '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("Codigo", "Código", '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' + simples
           + '<w:ind w:left="284" w:firstLine="0"/><w:jc w:val="left"/>',
           '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="18"/>'),
        st("Alinea", "Alínea", '<w:spacing w:before="0" w:after="60" w:line="360" w:lineRule="auto"/>'
           '<w:ind w:left="1276" w:hanging="567"/><w:jc w:val="both"/>', ""),
        st("Referencia", "Referência", '<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/>'
           + sem_recuo + '<w:jc w:val="left"/>', ""),
    ])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="pt-BR" w:eastAsia="pt-BR" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl/></w:pPr></w:pPrDefault></w:docDefaults>'
            + corpo + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
            '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
            '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
            '</w:tblCellMar></w:tblPr></w:style></w:styles>')


def configuracoes(atualizar=False):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:defaultTabStop w:val="709"/><w:hyphenationZone w:val="425"/>'
            '<w:characterSpacingControl w:val="doNotCompress"/>'
            + ('<w:updateFields w:val="true"/>' if atualizar else "")
            +             '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<w:themeFontLang w:val="pt-BR"/></w:settings>')


# ---------------------------------------------------------------------------
# LibreOffice: numeração do sumário e PDF de conferência
# ---------------------------------------------------------------------------
def comando_soffice():
    cmd = os.environ.get("TIMBRADOR_SOFFICE")
    if cmd:
        return shlex.split(cmd)
    for nome in ("soffice", "libreoffice"):
        if shutil.which(nome):
            return [nome]
    return None


def para_pdf(docx, pasta):
    cmd = comando_soffice()
    if not cmd:
        return None
    with tempfile.TemporaryDirectory(prefix="timbrador_lo_") as perfil:
        try:
            subprocess.run(cmd + [f"-env:UserInstallation=file://{perfil}", "--headless",
                                  "--convert-to", "pdf", "--outdir", pasta, docx],
                           capture_output=True, timeout=240)
        except (OSError, subprocess.TimeoutExpired):
            return None
    pdf = os.path.join(pasta, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


def paginas_texto(pdf):
    if not shutil.which("pdftotext"):
        return None
    r = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, timeout=120)
    if r.returncode:
        return None
    norm = lambda t: re.sub(r"\s+", " ", sem_acento(t)).casefold()
    return [norm(p) for p in r.stdout.decode("utf-8", "replace").split("\f")]


def localizar_paginas(doc, paginas):
    fim = next((i for i, p in enumerate(paginas) if "§§fim-sumario§§" in p), None)
    if fim is None:
        return None
    achadas, atual = {}, fim + 1
    for _nivel, texto, bm in doc.sumario:
        alvo = re.sub(r"\s+", " ", sem_acento(texto)).casefold()[:50]
        for i in range(atual, len(paginas)):
            if alvo in paginas[i]:
                achadas[bm] = i + 1
                atual = i
                break
        else:
            return None
    return achadas


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def localizar_marca(arg, spec_path):
    candidatos = [arg, os.environ.get("TIMBRADOR_MARCA"),
                  os.path.join(os.path.dirname(os.path.abspath(spec_path)), "marca.json") if spec_path else None,
                  os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "marca.json")]
    for c in candidatos:
        if c and os.path.exists(c):
            return c
    raise ErroEspec("marca.json não encontrado. Informe --marca, defina TIMBRADOR_MARCA ou crie um a partir de "
                    "assets/marca.exemplo.json.")


def main():
    ap = argparse.ArgumentParser(description="Timbrador white label: JSON → .docx timbrado.")
    ap.add_argument("especificacao", nargs="?")
    ap.add_argument("--marca")
    ap.add_argument("--saida", default=".")
    ap.add_argument("--pdf", action="store_true", help="salva também o PDF para conferência")
    ap.add_argument("--checar-marca", metavar="MARCA")
    ap.add_argument("--embutir-logos", metavar="MARCA")
    a = ap.parse_args()
    try:
        if a.checar_marca:
            print(Marca(a.checar_marca).resumo())
            return 0
        if a.embutir_logos:
            destino = a.saida if a.saida.endswith(".json") else os.path.join(
                a.saida, os.path.splitext(os.path.basename(a.embutir_logos))[0] + "_portatil.json")
            print("Gerado:", embutir_logos(a.embutir_logos, destino))
            return 0
        if not a.especificacao:
            ap.error("informe a especificação JSON")
        with open(a.especificacao, encoding="utf-8") as f:
            spec = json.load(f)
        marca = Marca(localizar_marca(a.marca, a.especificacao))
        doc = Documento(spec, marca, os.path.dirname(os.path.abspath(a.especificacao)))
        os.makedirs(a.saida, exist_ok=True)
        destino = os.path.join(a.saida, doc.nome_arquivo())
        doc.gravar(destino)
        numerado = doc.variante != "longo" or not spec.get("sumario", True)
        pdf_final = None
        with tempfile.TemporaryDirectory(prefix="timbrador_") as tmp:
            if not numerado:
                pdf = para_pdf(destino, tmp)
                pags = paginas_texto(pdf) if pdf else None
                achadas = localizar_paginas(doc, pags) if pags else None
                if achadas is not None:
                    doc.paginas_sumario = achadas
                    doc.gravar(destino)
                    numerado = True
                else:
                    # sem LibreOffice: remove o marcador e deixa o Word atualizar ao abrir
                    doc.paginas_sumario = {}
                    arquivos = doc.montar()
                    arquivos["word/settings.xml"] = configuracoes(atualizar=True)
                    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
                        z.writestr("[Content_Types].xml", arquivos.pop("[Content_Types].xml"))
                        for n, d in arquivos.items():
                            z.writestr(n, d)
                    doc.avisos.append("sumário sem números de página (LibreOffice/pdftotext indisponível): "
                                      "o Word vai pedir para atualizar os campos ao abrir.")
            if a.pdf:
                pdf = para_pdf(destino, a.saida)
                pdf_final = pdf
                if not pdf:
                    doc.avisos.append("PDF não gerado (LibreOffice indisponível).")
        print(f"Gerado: {destino}")
        if pdf_final:
            print(f"PDF: {pdf_final}")
        print(f"Resumo: {doc.emp['sigla']} · {doc.tipo_nome} ({doc.variante}) · {doc.codigo} · "
              f"v{doc.versao} · {doc.estado}")
        avisos = doc.avisos + [x for x in doc.avisos_montagem if x not in doc.avisos]
        for x in dict.fromkeys(avisos):
            print(f"AVISO: {x}")
        return 0
    except ErroEspec as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return 1
    except (OSError, json.JSONDecodeError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
