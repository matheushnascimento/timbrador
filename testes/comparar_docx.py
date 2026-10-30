#!/usr/bin/env python3
"""Compara o .docx do gerador com o da página. Ignora o que não muda o documento: data de criação, nome do
arquivo do logo, ordem dos estilos e das relações e, no sumário, os números de página (a página deixa para o Word)."""
import json, re, sys, zipfile


def normal(nome, x):
    x = x.decode("utf8")
    x = re.sub(r"<dcterms:(created|modified)[^>]*>[^<]*</dcterms:\1>", "", x)
    x = re.sub(r"media/(img\d+|logo_[A-Z0-9]+)\.(png|jpeg)", "media/LOGO", x)
    x = x.replace('<w:updateFields w:val="true"/>', "")
    if nome == "customXml/item1.xml":
        return json.dumps(json.loads(re.search(r">(\{.*\})<", x.replace("&quot;", '"')).group(1)), sort_keys=True)
    if nome.endswith(".rels") or nome == "[Content_Types].xml":
        return sorted(re.findall(r"<(?:Relationship|Default|Override) [^>]*/>", x))
    if nome == "word/styles.xml":
        return sorted(re.findall(r"<w:style .*?</w:style>", x))
    if nome == "word/document.xml":   # número de página das linhas do sumário
        x = re.sub(r'<w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">\d+</w:t></w:r></w:hyperlink>', "</w:hyperlink>", x)
    return x


def main(a, b):
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    na = {re.sub(r"media/(img\d+|logo_[A-Z0-9]+)", "media/LOGO", n): n for n in za.namelist()}
    nb = {re.sub(r"media/(img\d+|logo_[A-Z0-9]+)", "media/LOGO", n): n for n in zb.namelist()}
    dif = sorted(set(na) ^ set(nb))
    for k in sorted(set(na) & set(nb)):
        xa, xb = za.read(na[k]), zb.read(nb[k])
        if (xa != xb) if k.startswith("word/media") else (normal(k, xa) != normal(k, xb)):
            dif.append(k)
    for k in dif:
        print("  difere:", k)
    return 1 if dif else 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:3]))
