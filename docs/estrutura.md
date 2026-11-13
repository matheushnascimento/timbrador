# Estrutura, histórico e comandos

> Movido do `CLAUDE.md` em 10/11/2026 (adoção do BMad, `bmad-project-context`). Texto mantido como estava; o `AGENTS.md` aponta para cá.

## Visão geral

Plugin do Claude que gera documentos `.docx` timbrados (papel timbrado + ABNT +
controle documental), com a identidade visual e os dados cadastrais de qualquer
organização vindos de um `marca.json`.

O fluxo: o Claude lê o conteúdo bruto do usuário, classifica (empresa, tipo,
área, estado), escreve uma especificação JSON e chama o gerador, que monta o
`.docx` escrevendo WordprocessingML direto num zip. **Nada de `python-docx` ou
`docx-js`**: a dependência zero é intencional, para o plugin rodar em qualquer
lugar. Só biblioteca padrão do Python 3.8+.

## Estado atual


O gerador **é o original** (o Python embutido no `timbrador.html`, recebido em
20/10/2026), portado para white label: empresas, áreas e tipos ativos vêm do
`marca.json`; do gerador anterior ficaram só `--marca`, `--checar-marca`,
`--embutir-logos`, `--pdf`, `TIMBRADOR_SOFFICE` e `documento_rotulo` (CPF).
A página de navegador também foi portada (20/10/2026): `pagina/timbrador.html`,
sem nenhuma empresa fixa, lendo o mesmo `marca.json` e montando o .docx com um
porte direto do gerador — o pacote sai igual (`testes/pagina.sh`). As
divergências que motivaram o porte estão em `docs/divergencias-original.md`.

## Estrutura


| Caminho | Papel |
|---|---|
| `.claude-plugin/plugin.json` | manifesto do plugin |
| `README.md` | instruções de personalização para humanos |
| `skills/timbrador/SKILL.md` | a norma: estados, códigos, ABNT, formato do JSON |
| `skills/timbrador/scripts/gerar_documento.py` | gerador (original portado), ~3.000 linhas, só stdlib |
| `skills/timbrador/assets/marca.exemplo.json` | modelo de configuração |
| `skills/timbrador/assets/logo-exemplo.png` | logo genérico do exemplo |
| `pagina/timbrador.html` | página de navegador (porte do original), lê o `marca.json` com logos embutidos |
| `pagina/Inventariar.ps1` (+ `.cmd`) | inventário da pasta no Windows sem Python; lê o `marca.json` (`-Marca`, `TIMBRADOR_MARCA` ou ao lado do script) |
| `pagina/Abrir-Timbrador.cmd` | abre a página no Edge ou no Chrome (os que gravam direto na pasta) |
| `testes/` | marca e especificações fictícias + teste de fumaça + página × gerador |
| `docs/divergencias-original.md` | original × gerador anterior, item a item |
| `original/` (fora do git) | pacote recebido: `timbrador.html`, `.cmd`, `.ps1`; tem dados reais |

`Timbrador-white-label-v1.zip` é a entrega original
(md5 `128a94c59c82bdf24c2392a6b0a2f79d`); está no `.gitignore` porque o conteúdo
já está versionado. O commit `70a7137` é esse estado intacto — a linha de base
para qualquer diff contra o original.

## Comandos


```bash
python3 skills/timbrador/scripts/gerar_documento.py espec.json --marca marca.json --saida saida/ --pdf
python3 skills/timbrador/scripts/gerar_documento.py --checar-marca marca.json
python3 skills/timbrador/scripts/gerar_documento.py --embutir-logos marca.json --saida marca_portatil.json
python3 skills/timbrador/scripts/gerar_documento.py --ler|--atualizar|--obsoletar DOCX --marca marca.json ...
python3 skills/timbrador/scripts/gerar_documento.py --inventariar PASTA --marca marca.json [--conferir]
bash testes/fumaca.sh        # marca, curto/longo, travas, grupo, CPF, assinaturas previstas, ciclo de vida, projeto, anexos, lista mestra, inventário, relatório de condução, PDF, XML
bash testes/pagina.sh        # página × gerador, editor da marca, relatório de condução lido e continuado pelos dois lados, versão da atualização e prévia em PDF (precisa de node + jsdom; pula sem eles)
bash testes/inventario.sh    # Inventariar.ps1 × --inventariar: lista mestra e alertas iguais (precisa de pwsh ou TIMBRADOR_PWSH; pula sem ele)
```

LibreOffice e `pdftotext` são **opcionais**: com eles o gerador converte para
PDF, descobre em que página cada título caiu e grava os números no sumário; sem
eles o sumário sai vazio e o Word atualiza ao abrir. `TIMBRADOR_SOFFICE` troca o
comando do LibreOffice (necessário em sandboxes onde `soffice` precisa de
wrapper).

Nesta máquina os dois estão instalados (18/09/2026): LibreOffice 24.2.7
(`soffice`) e poppler-utils 24.02 (`pdftotext`, `pdftoppm`). O caminho do PDF e
da numeração do sumário **é** exercitado localmente — se o aviso "sumário sem
números de página" reaparecer, alguma das duas sumiu do PATH.
