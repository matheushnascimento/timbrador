# Estado da validação

> Movido do `CLAUDE.md` em 10/11/2026 (adoção do BMad, `bmad-project-context`). Texto mantido como estava; leia antes de dizer que algo está testado ou de mudar `testes/`.


Coberto pelo `fumaca.sh` (gerador portado, 20/10/2026): curto e longo; minuta,
em aprovação, oficial; estados e empresa aposentados (aviso); empresa sem logo,
incompleta, com CPF e de grupo; assinaturas em linhas e em tabela; título sem o
nome do tipo; `--ler`, `--atualizar`, lista mestra (registro e número
repetido), `--inventariar --conferir`, `--pdf`; XML bem formado. Conferência
visual em PDF via LibreOffice em 20/10/2026: capa, folha de controle, sumário
numerado e página interna do POP; CI do grupo com assinaturas e rodapé de logos.

Página (20/10/2026, `testes/pagina.sh` e jsdom): curto e longo do teste, grupo,
CPF, empresa incompleta, obsoleto por motivo, estado/tipo/área aposentados,
assinaturas em linhas e em tabela, e um curto e um longo pela interface
(texto colado + formulário + Gerar) — pacote igual ao do gerador parte a parte
e PDFs idênticos pixel a pixel no LibreOffice.

Visual (03/11/2026): regressão sem `visual` idêntica ao gerador anterior (6
documentos); `testes/pagina.sh` com visual e fontes embutidas (DejaVu Serif do
sistema); marca mhdn com visual igual página × gerador (25 e 28 partes); PDFs
com Newsreader e EB Garamond embutidas (LibreOffice usa as fontes do .docx).

Recursos restaurados e relatório (03/11/2026): `fumaca.sh` com 64 verificações (67 com a
minuta de nova versão, 04/11/2026)
(assinaturas previstas, projeto, anexos com SHA-256 e anexo alterado recusado,
inventário com Obsoletos e projeto fora da pasta, 19 casos do relatório de condução);
`pagina.sh` com o documento de grupo com assinatura e o relatório lido e continuado
pelos dois lados; interface no Chromium (Playwright): texto colado, Projeto,
assinatura, anexo, Gerar (zip com documento e anexo), caixa do relatório e rascunho
recuperado ao recarregar, sem erro de console.

Prévia em PDF (03/11/2026): `pagina.sh` gera pelo núcleo da página o PDF do curto,
do longo, do grupo com assinatura e do relatório de condução e confere estrutura,
código e "Página X de Y"; comparados ao .docx convertido pelo LibreOffice, saem com a
mesma paginação (o longo com 10 páginas contra 11, porque o LibreOffice põe uma
página em branco depois da capa). Interface no Chromium: borda na cor do estado,
formulário à esquerda, prévia atualizada ao mudar o estado, pendências acima dela;
no Google Chrome (Playwright, `channel="chrome"`), a prévia aparece no leitor de
PDF tanto aberta como arquivo quanto servida pelo nginx com a CSP. O Chromium e o
Firefox do Playwright vêm sem leitor de PDF: neles o quadro fica cinza.

**Não coberto ainda:** a prévia no Edge e no Firefox de verdade, as fontes embutidas abertas no Word de verdade, `--obsoletar` e a cópia obsoleta automática, inventário
gravando, `assinatura_formato` automático trocando para tabela, figuras JPEG grandes, tabelas que quebram entre páginas,
apêndices com muitos níveis, documentos acima de ~30 páginas, e a aparência real
no Word (a conferência foi toda via LibreOffice).
