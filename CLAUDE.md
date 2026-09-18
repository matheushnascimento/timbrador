# Timbrador (white label) — contexto do projeto

Plugin do Claude que gera documentos `.docx` timbrados (papel timbrado + ABNT +
controle documental), com a identidade visual e os dados cadastrais de qualquer
organização vindos de um `marca.json`.

O fluxo: o Claude lê o conteúdo bruto do usuário, classifica (empresa, tipo,
área, estado), escreve uma especificação JSON e chama o gerador, que monta o
`.docx` escrevendo WordprocessingML direto num zip. **Nada de `python-docx` ou
`docx-js`**: a dependência zero é intencional, para o plugin rodar em qualquer
lugar. Só biblioteca padrão do Python 3.8+.

## Estado atual

O plugin cobre o caminho do Claude. **Falta a página de navegador** — ver
"Pendência que bloqueia" abaixo.

## Estrutura

| Caminho | Papel |
|---|---|
| `.claude-plugin/plugin.json` | manifesto do plugin |
| `README.md` | instruções de personalização para humanos |
| `skills/timbrador/SKILL.md` | a norma: estados, códigos, ABNT, formato do JSON |
| `skills/timbrador/scripts/gerar_documento.py` | gerador, ~1.300 linhas, só stdlib |
| `skills/timbrador/assets/marca.exemplo.json` | modelo de configuração |
| `skills/timbrador/assets/logo-exemplo.png` | logo genérico do exemplo |
| `testes/` | marca e especificações fictícias + teste de fumaça |

`Timbrador-white-label-v1.zip` é a entrega original
(md5 `128a94c59c82bdf24c2392a6b0a2f79d`); está no `.gitignore` porque o conteúdo
já está versionado. O commit `70a7137` é esse estado intacto — a linha de base
para qualquer diff contra o original.

## Comandos

```bash
python3 skills/timbrador/scripts/gerar_documento.py espec.json --marca marca.json --saida saida/ --pdf
python3 skills/timbrador/scripts/gerar_documento.py --checar-marca marca.json
python3 skills/timbrador/scripts/gerar_documento.py --embutir-logos marca.json --saida marca_portatil.json
bash testes/fumaca.sh        # confere marca, gera curto e longo, testa as travas, valida o XML
```

LibreOffice e `pdftotext` são **opcionais**: com eles o gerador converte para
PDF, descobre em que página cada título caiu e grava os números no sumário; sem
eles o sumário sai vazio e o Word atualiza ao abrir. `TIMBRADOR_SOFFICE` troca o
comando do LibreOffice (necessário em sandboxes onde `soffice` precisa de
wrapper).

**Nesta máquina não há LibreOffice nem `pdftotext`.** Enquanto for assim, o
caminho do PDF e da numeração do sumário não é exercitado localmente:
`sudo apt install -y libreoffice-writer poppler-utils`.

## Invariantes

1. **O `marca.json` é a única coisa que varia entre organizações.** Sigla, nome,
   razão social, documento (CNPJ ou CPF), cidade, cor, logo — mais `areas`,
   `areas_extra` e `tipos_extra` opcionais. **Nenhum dado de empresa pode voltar
   para o código.**
2. **O visual é fixo e não se negocia com o conteúdo**: cabeçalho com faixa de
   estado e bloco de controle (código · versão · Página X de Y; situação; razão
   social · documento · cidade), rodapé com filete na cor da empresa e logo,
   borda de página, marca d'água nos estados não oficiais, margens ABNT 3/3/2/2.
3. **As travas de estado são de segurança documental, não de validação.**
   OFICIAL exige aprovador, data de aprovação, número reservado e empresa com
   cadastro completo; OBSOLETO exige substituto ou motivo. Empresa com
   `razao_social` ou `documento` em `[definir]` só emite RASCUNHO, MINUTA ou
   EM_REVISAO. Nunca contorne: rebaixe o estado ou peça o dado.
4. **Dado que falta vira `[definir]`.** Não invente razão social, CNPJ, número,
   cláusula, data ou valor.
5. **Gerador e página andam juntos.** Qualquer mudança de visual, de regra ou de
   formato da marca precisa acontecer nos dois, ou o documento gerado pelo
   Claude deixa de ser igual ao gerado pela página.

## Pendência que bloqueia: o `timbrador.html`

O plugin original tinha uma página de navegador, `timbrador.html`, que qualquer
colaborador abria para timbrar um `.docx`/`.txt`/`.md` pronto, sem reorganizar o
texto. Ela carrega os logos embutidos e repete as regras do gerador. **Essa
página não veio no pacote**, e o gerador atual foi escrito a partir da descrição
na skill, não do código original.

**Enquanto o arquivo não chegar, não implemente a página nem "corrija" o visual
por suposição.** Peça o arquivo. Quando ele chegar:

1. Leia a página inteira antes de mudar qualquer linha do gerador.
2. Liste as divergências de visual e de regra entre a página e
   `gerar_documento.py` (medidas, cores, textos de faixa, ordem dos campos do
   cabeçalho, nomenclatura, travas de estado). **Mostre a lista antes de editar.**
3. Trate o original como referência de visual: onde houver conflito, o gerador se
   ajusta à página, a menos que o usuário decida o contrário.
4. Só então porte a página para white label: os dados das quatro empresas
   embutidos saem do HTML e passam a vir do mesmo `marca.json` (carregado por
   upload ou colado), inclusive os logos em data URI. A página não pode
   continuar com nenhuma empresa fixa no código.
5. Volte a rodar os testes dos dois lados e compare o mesmo documento gerado
   pelos dois caminhos.

Se o usuário anexar o `gerar_documento.py` original junto, **ele manda em tudo
que for regra e formato do JSON**; a versão atual vira, no máximo, fonte de
ideias para as partes que ele não cobria.

## Estado da validação

Coberto: documentos curtos e longos; estados minuta, em revisão e oficial;
empresa sem logo; empresa com cadastro incompleto; casos que devem ser
recusados; XML validado contra o schema do Word; conferência visual em PDF
(feita no ambiente anterior, com LibreOffice).

**Não coberto ainda:** figuras JPEG grandes, tabelas que quebram entre páginas,
apêndices com muitos níveis, documentos acima de ~30 páginas, e a aparência real
no Word (a conferência foi toda via LibreOffice).

## Convenções

- Documentos gerados (`.docx`, `.pdf`) e `marca.json` de organização real **não
  entram no repositório** — ver `.gitignore`. Só `marca.exemplo.json` e as
  marcas fictícias de `testes/` são versionados.
- Mensagens de commit em português, no padrão conventional commits.
