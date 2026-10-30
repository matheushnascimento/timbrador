# Timbrador (white label) — contexto do projeto

Plugin do Claude e página de navegador que geram documentos `.docx` timbrados (papel timbrado + ABNT +
controle documental), com a identidade visual e os dados cadastrais de qualquer organização vindos de um
`marca.json`.

O fluxo do Claude: lê o conteúdo bruto do usuário, classifica (empresa, tipo, área, estado), escreve uma
especificação JSON e chama o gerador, que monta o `.docx` escrevendo WordprocessingML direto num zip. O fluxo
do colaborador: abre `timbrador.html`, solta um `.docx`/`.txt`/`.md` pronto e a página monta o mesmo `.docx`
no navegador. **Nada de `python-docx` ou `docx-js`**: a dependência zero é intencional. Gerador só com a
biblioteca padrão do Python 3.8+; página num único HTML, sem bibliotecas.

## Origem

O projeto foi regerado em 30/10/2026 a partir da versão mais recente do Timbrador original (página
`timbrador.html` com o `gerar_documento.py` embutido), de uma organização real. O original é a referência de
visual e de regra; tudo o que era da organização (empresas, CNPJs, logos, cores, siglas, nome do grupo, código
da norma interna, namespace da ficha) saiu do código e passou a vir do `marca.json`.

**O material original nunca entra no repositório** — nem no histórico. Se o usuário anexar de novo o zip ou
a página original, trabalhe numa cópia fora do repositório e traga só código já sem dado de empresa
(`grep -ri` pela razão social, CNPJ, siglas e cor antes de qualquer commit). Os scripts de PowerShell do
original (`Inventariar.ps1`/`.cmd`) ficaram de fora: tinham as siglas fixas e repetem o `--inventariar`.

Os commits `70a7137`…`3216b25` são o white label v1, escrito a partir da descrição da skill e não do código
original; foi substituído pela regeração.

## Estrutura

| Caminho | Papel |
|---|---|
| `.claude-plugin/plugin.json` | manifesto do plugin |
| `README.md` | instruções de personalização para humanos |
| `skills/timbrador/SKILL.md` | a norma: estados, códigos, ABNT, formato do JSON, comandos de controle |
| `skills/timbrador/scripts/gerar_documento.py` | gerador (~2.500 linhas, só stdlib) |
| `skills/timbrador/assets/timbrador.html` | página dos colaboradores; bloco `marca-embutida` vazio no repositório |
| `skills/timbrador/assets/Abrir-Timbrador.cmd` | abre a página no Edge/Chrome (Windows) |
| `skills/timbrador/assets/marca.exemplo.json` | modelo de configuração |
| `skills/timbrador/assets/logo-exemplo.png` | logo genérico do exemplo |
| `testes/` | marca e especificações fictícias, teste de fumaça, comparação página × gerador |

## Comandos

```bash
python3 skills/timbrador/scripts/gerar_documento.py espec.json --marca marca.json --saida saida/ --pdf
python3 skills/timbrador/scripts/gerar_documento.py --checar-marca marca.json
python3 skills/timbrador/scripts/gerar_documento.py --embutir-logos marca.json --saida marca_portatil.json
python3 skills/timbrador/scripts/gerar_documento.py --montar-pagina marca.json --saida pasta/
python3 skills/timbrador/scripts/gerar_documento.py --marca marca.json --ler doc.docx | --atualizar doc.docx --o-que-mudou "…"
python3 skills/timbrador/scripts/gerar_documento.py --marca marca.json --inventariar PASTA [--conferir]
bash testes/fumaca.sh   # marca, geração, travas, lista mestra, versões, XML e página × gerador
```

LibreOffice e `pdftotext` são **opcionais**: com eles o gerador descobre em que página cada título caiu e grava
os números no sumário; sem eles o sumário sai sem números e o Word atualiza ao abrir. `TIMBRADOR_SOFFICE`
troca o comando do LibreOffice. A comparação página × gerador do teste de fumaça precisa de Node com
Playwright (`NODE_PATH=$(npm root -g)`); sem eles, é pulada.

Nesta máquina estão instalados LibreOffice 24.2.7, poppler-utils 24.02, Node 22 e Playwright com Chromium.

## Invariantes

1. **O `marca.json` é a única coisa que varia entre organizações.** Empresas (sigla, marca, razão social,
   documento e rótulo, cidade, cor, logo, `grupo`, `aposentada`), `empresa_padrao`, `organizacao`,
   `politica`, `areas`/`areas_extra`, `tipos_extra`, `tipos_ativos`/`areas_ativas`. **Nenhum dado de
   empresa pode voltar para o código** — nem para comentários e exemplos.
2. **O visual é fixo e não se negocia com o conteúdo**: faixa de estado no topo (1,2 a 19,8 cm), bloco de
   controle centralizado (código · versão · Página X de Y; situação; razão social · documento · cidade),
   marca d'água nos estados não oficiais, borda preta a 1 cm do papel, rodapé com filete na cor da empresa e
   logo, margens 3,3 (alto) / 2,0 (direita) / 3,2 ou 2,6 (base) / 3,0 (esquerda) cm.
3. **As travas de estado são de segurança documental, não de validação.** Estados ativos: MINUTA,
   EM_APROVACAO, OFICIAL, OBSOLETO (RASCUNHO e EM_REVISAO só para ler documentos antigos). OFICIAL exige
   aprovador, data de aprovação, número e empresa com cadastro completo; OBSOLETO exige substituto ou motivo.
   Empresa com `razao_social` ou `documento` em `[definir]` só emite MINUTA ou EM_APROVACAO. Número já usado
   por outro título na lista mestra é recusado. Nunca contorne: rebaixe o estado ou peça o dado.
4. **Dado que falta vira `[definir]`.** Não invente razão social, CNPJ, número, cláusula, data ou valor.
5. **Gerador e página andam juntos.** Qualquer mudança de visual, de regra ou de formato da marca acontece
   nos dois, e `testes/fumaca.sh` confere que o mesmo documento sai igual pelos dois caminhos. Em conflito,
   vale a página para visual e o gerador para regra e formato do JSON.
6. **Assinatura entra só pela página** (caixa "Com assinatura"); o gerador guarda os envolvidos como
   `assinaturas_previstas` na ficha.

## Diferenças conhecidas entre página e gerador (herdadas do original)

Não aparecem nas especificações do teste, mas existem:

- tabela **sem título**: a página não põe legenda nem fonte e, sem `tipo`, desenha Quadro; o gerador desenha
  Tabela com legenda vazia e fonte padrão. (A página recebe tabelas sem título de textos lidos; o gerador,
  de especificações escritas pelo Claude.)
- sumário: a página só o põe com 3 ou mais títulos e sem números de página (o Word numera ao abrir); o
  gerador sempre o põe e numera pelo LibreOffice.
- OBSOLETO: a página pede `substituido_por`; o gerador aceita também `motivo_obsoleto`.
- figuras do corpo: a página não as mantém ao reler um `.docx` (avisa); o gerador as extrai com `--pasta-midia`.

## Estado da validação (30/10/2026)

Coberto pelo teste de fumaça: marca (checar, embutir logos, montar página); curto em minuta, longo oficial,
documento de grupo, empresa sem logo e incompleta em aprovação; travas (oficial sem aprovador, sem número,
de empresa incompleta; obsoleto sem motivo; empresa fora da marca); lista mestra (registro e número
repetido); `--ler`, `--atualizar` e geração da nova versão; XML bem formado; o mesmo curto e o mesmo longo
gerados pela página (Chromium sem janela) e pelo gerador saem iguais, fora os números de página do sumário.
A página sem marca mostra só o bloco da marca e aceita o `marca_portatil.json`. Conferência visual em PDF
(LibreOffice) do curto, do longo, do grupo e da empresa sem logo.

**Não coberto ainda:** interface da página além do carregamento da marca (pasta conectada, inventário,
anexos, assinatura automática, rascunho), `--obsoletar`, `--inventariar` e `--anexo` no gerador, figuras
JPEG grandes, tabelas que quebram entre páginas, documentos acima de ~30 páginas e a aparência real no Word
(a conferência foi toda via LibreOffice).

## Convenções

- Documentos gerados (`.docx`, `.pdf`), `lista-mestra.json`, páginas montadas com marca e `marca.json` de
  organização real **não entram no repositório** — ver `.gitignore`. Só `marca.exemplo.json` e as marcas
  fictícias de `testes/` são versionados.
- Mensagens de commit em português, no padrão conventional commits.
