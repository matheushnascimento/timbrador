---
name: timbrador
description: "Timbrador white label: converte qualquer conteúdo em documento Word (.docx) no Padrão de Documentação Oficial (ABNT + papel timbrado da organização: logo, cor, borda, cabeçalho de controle), com código, versão, faixa e marca d'água de estado. A identidade vem de um arquivo marca.json, então serve para qualquer empresa ou grupo. Use sempre que o usuário quiser timbrar, formalizar ou padronizar um documento — CI, comunicado, ofício, termo, contrato, ata, advertência, recibo, manual, política, POP, regulamento, plano ou relatório — ou configurar o papel timbrado da própria empresa, mesmo que não diga 'timbrado'."
---

# Timbrador (white label)

Transforma qualquer informação em um .docx que qualquer pessoa reconhece de imediato como documento oficial da organização e cujo estado (oficial ou não) é inequívoco. A identidade visual e os dados cadastrais vêm do `marca.json`; o visual, as regras e a numeração são fixos. Saída: **somente .docx** (e PDF de conferência, se pedido). Responda no idioma do usuário; o documento sai em português.

Arquivos desta skill:

| Arquivo | Papel |
|---|---|
| `scripts/gerar_documento.py` | Gerador. Só biblioteca padrão do Python 3.8+. Rode direto daqui; ele só escreve na pasta de `--saida`. |
| `assets/marca.exemplo.json` + `assets/logo-exemplo.png` | Modelo para criar o `marca.json`. |
| `assets/marca.json` (opcional) | Se quem distribuiu a skill já a personalizou, a marca fica aqui. |

## 1. Marca (configuração da organização)

**Onde procurar o `marca.json`** (pare no primeiro): anexos da conversa → pasta conectada do usuário → `assets/marca.json` desta skill. Rode `python3 scripts/gerar_documento.py --checar-marca CAMINHO` e mostre o resumo ao usuário na primeira vez.

**Se não existir, configure antes de gerar.** Pergunte (AskUserQuestion, se disponível) o que faltar para cada empresa emissora:

| Campo | Exemplo | Observação |
|---|---|---|
| `sigla` | `ACME` | 2 a 6 letras/números; abre o código de todo documento |
| `marca` | `Acme` | nome fantasia; aparece no rodapé quando não há logo |
| `razao_social` | `Acme Comércio Ltda` | `[definir]` enquanto não souber |
| `documento_rotulo` / `documento` | `CNPJ` / `00.000.000/0001-00` | pode ser `CPF` para autônomo |
| `cidade` | `Campinas – SP` | usada no cabeçalho, capa e `local_data` |
| `cor` | `1F3864` | hexadecimal; filete, borda e títulos. Se o usuário não souber, extraia a cor dominante do logo |
| `logo` | `logo.png` | PNG ou JPEG, de preferência horizontal e com fundo transparente |

Monte o arquivo a partir de `assets/marca.exemplo.json` (campos opcionais: `organizacao`, `empresa_padrao`, `politica` com código/título da norma interna, `areas` para substituir a lista de áreas, `areas_extra` e `tipos_extra` `{"SIG": {"nome": "...", "variante": "curto"|"longo"}}` para ampliar). Várias empresas de um mesmo grupo cabem na lista `empresas`.

Depois, gere a versão de arquivo único com `--embutir-logos marca.json --saida marca_portatil.json` (o logo vira texto dentro do JSON) e entregue ao usuário dizendo para anexá-la nas próximas conversas ou deixá-la na pasta conectada. Não invente razão social, CNPJ ou cidade: o que faltar fica `[definir]`.

**Regras da marca**

- Empresa com `razao_social` ou `documento` vazio/`[definir]` só emite RASCUNHO, MINUTA ou EM_REVISAO; o gerador recusa OFICIAL/OBSOLETO. Quando o usuário informar os dados, atualize o `marca.json` e entregue a nova versão.
- Cada documento leva **somente o logo da empresa emissora**. Nunca logos de fornecedores, fabricantes, parceiros ou de outra empresa do grupo.
- Descubra a emissora pelo conteúdo (quem assina, de quem é o colaborador ou o cliente). Na dúvida, pergunte; sem resposta, use `empresa_padrao` e avise.

**Visual fixo (não alterar):** A4 com margens ABNT (3/3/2/2 cm); cabeçalho de todas as páginas com a faixa do estado e o bloco de controle (código · versão · Página X de Y; situação; razão social · documento · cidade); rodapé com filete na cor da empresa e logo (centralizado no curto e na capa; logo + título nas páginas internas do longo); borda fina na cor da empresa. Nome e CNPJ não se repetem no corpo. Nunca desenhe o timbrado à mão nem edite o gerador para mudar o visual.

## 2. Fluxo

1. **Ler o conteúdo** recebido por inteiro e localizar o `marca.json` (seção 1).
2. **Classificar**: empresa, tipo (Quadro A), área (Quadro B), estado (seção 3). Pergunte só o que não dá para deduzir: empresa, estado desejado, número reservado na lista mestra, signatários, aprovador. Sem resposta: estado **RASCUNHO**, número vazio (sai **XXX**) e diga isso no final.
3. **Adequar o texto** às regras de redação (seção 5). Não invente fato, data, valor, nome ou cláusula: o que faltar vira `[definir]`.
4. **Escrever a especificação JSON** (seção 6) em arquivo de trabalho; figuras com caminho relativo a esse arquivo.
5. **Gerar**: `python3 scripts/gerar_documento.py especificacao.json --marca marca.json --saida PASTA/ --pdf` (o nome do arquivo sai no padrão). Se o LibreOffice do ambiente precisar de um wrapper, informe o comando em `TIMBRADOR_SOFFICE`.
6. **Tratar erros e avisos**: o gerador recusa OFICIAL sem `aprovado_por`, `data_aprovacao`, número ou dados da empresa, e OBSOLETO sem substituto ou motivo. Não contorne: rebaixe o estado ou peça o dado. Repasse os avisos relevantes.
7. **Conferir o visual** quando o PDF sair: `pdftoppm -png -r 60` e olhe a primeira página e uma interna (faixa, bloco de controle, logo no rodapé, borda, marca d'água, tabelas cortadas, títulos órfãos). Sem LibreOffice, o sumário do longo sai sem números e o Word pede para atualizar os campos ao abrir; avise.
8. **Entregar** o .docx (na pasta conectada, se houver, ao lado da origem) com uma linha: empresa, tipo, código, versão, estado e pendências (`[definir]`, número XXX, aprovação).

## 3. Estados

| Estado (JSON) | Quando | Visual gerado |
|---|---|---|
| `RASCUNHO` | Em elaboração, incompleto | Faixa cinza + marca d'água |
| `MINUTA` | Texto completo proposto para análise ou concordância das partes | Faixa amarela + marca d'água |
| `EM_REVISAO` | Submetido formalmente ao aprovador | Faixa azul + marca d'água |
| `OFICIAL` | Aprovado (exige `aprovado_por`, `data_aprovacao`, número e empresa completa) | Faixa verde, sem marca d'água, aprovação no cabeçalho |
| `OBSOLETO` | Substituído ou cancelado (exige `substituido_por` ou `motivo_obsoleto`) | Faixa vermelha + marca d'água |

- Só marque OFICIAL quando o usuário informar quem aprovou e quando. "Deixa oficial" sem esses dados → EM_REVISAO e explique.
- Minuta antes do oficial em contratos, termos, acordos com colaboradores, comunicados a terceiros, políticas e regulamentos.
- Nova versão de um documento oficial começa em RASCUNHO ou MINUTA com a próxima versão e `-r1`.

## 4. Identificação

**Código**: `SIGLA-TIPO-ÁREA-NNN`; tipos curtos acrescentam o ano: `ACME-CI-RH-012/2026`. O gerador monta a partir de `empresa`, `tipo`, `area`, `numero`, `ano`. A numeração é por empresa, tipo e área.

**Versão**: oficial `X.Y` (X muda regra/processo; Y só correção). Não oficial = versão-alvo + `-rN`: `1.0-r1`, `1.0-r2`; revisão de um 1.0 → `2.0-r1` ou `1.1-r1`. Aprovado, o sufixo cai (o gerador remove). Sem `-r` num estado não oficial, o gerador acrescenta `-r1`.

**Nome do arquivo**: `CÓDIGO_vVERSÃO_ESTADO_titulo.docx` (barra do código vira hífen).

**Classificação**: Público, Interno (padrão), Confidencial.

Quadro A – Tipos padrão (a variante define o formato; `tipos_extra` do marca.json acrescenta outros)

| Curto (sem capa, texto direto após o título) | Longo (capa, folha de controle, sumário e páginas ABNT) |
|---|---|
| CI comunicação interna · COM comunicado · OFI ofício · DEC declaração · TER termo · ATA ata · AUT autorização · NOT notificação/advertência · CTR contrato/acordo · PRO proposta · FOR formulário · REC recibo | MAN manual · POL política · REG regulamento · POP procedimento operacional padrão · INS instrução de trabalho · PLN plano · PRJ projeto · REL relatório |

Quadro B – Áreas padrão (substituíveis por `areas` no marca.json): ADM administrativo · DIR diretoria · FIN financeiro · FIS fiscal/contábil · JUR jurídico · RH recursos humanos · VEN vendas/comercial · ECO e-commerce · EST estoque · EXP expedição · CMP compras · OFC oficina/assistência · MKT marketing · TI tecnologia. Use só as áreas que o `--checar-marca` lista; se nada servir, use a mais próxima e avise (ou proponha incluir em `areas_extra`).

## 5. Redação e ABNT (o gerador cuida da formatação; você cuida do conteúdo)

- Linguagem formal, impessoal, frases diretas. Datas dd/mm/aaaa; horas 07h30; valores R$ 1.234,56.
- **Seções numeradas** (NBR 6024) com `h1`…`h5`; não escreva o número no título, o gerador numera. Documento longo começa por "Objetivo" ou "Introdução". Não pule nível (h3 só dentro de h2).
- **Enumerações** em `alineas`: começam em minúscula, terminam em `;` e a última em `.`. Subitens com `sub`. Passos em ordem → `numerada`. Verificação → `checklist`. Evite `marcadores`.
- **Tabela** (dados numéricos, laterais abertas) × **Quadro** (texto, fechado): sempre `titulo` e `fonte` ("Elaborado pela empresa (AAAA)." quando for próprio). Cite no texto ("conforme Quadro 1").
- **Citações** (NBR 10520:2023): autor-data com sobrenome só com inicial maiúscula – (Silva, 2025, p. 10). Mais de 3 linhas → bloco `citacao`.
- **Referências** (NBR 6023:2025) em `referencias` (o gerador ordena): `SOBRENOME, Nome. **Título**: subtítulo. Local: Editora, ano.`; online acrescenta `Disponível em: URL. Acesso em: 15 set. 2026.`
- Documento curto com partes: termine com `local_data` e `assinaturas`. Documento longo oficial: a aprovação vai na folha de controle automaticamente.
- Não repita razão social e documento da emissora em `identificacao`; nas assinaturas eles continuam valendo.
- Destaque: `**negrito**`, `*itálico*`, `***ambos***` dentro de qualquer texto; `\n\n` separa parágrafos.

## 6. Especificação JSON

```json
{
  "empresa": "ACME",
  "tipo": "POP", "area": "ECO", "numero": 3, "ano": 2026,
  "titulo": "Cadastro de anúncios", "subtitulo": "opcional", "titulo_curto": "opcional, rodapé das páginas internas",
  "estado": "MINUTA", "versao": "1.0-r1", "data_emissao": "15/09/2026",
  "classificacao": "Interno", "area_nome": "opcional, sobrescreve o nome da área",
  "elaborado_por": "", "revisado_por": "", "aprovado_por": "", "cargo_aprovador": "", "data_aprovacao": "",
  "vigencia": "", "substitui": "", "substituido_por": "", "motivo_obsoleto": "",
  "identificacao": [["Destinatário", "..."], ["Assunto", "..."]],
  "historico": [{"versao": "1.0-r1", "data": "15/09/2026", "descricao": "Emissão inicial.", "responsavel": "..."}],
  "sumario": true,
  "corpo": [ BLOCOS ],
  "referencias": ["..."], "glossario": [["Termo", "definição."]],
  "apendices": [{"titulo": "...", "corpo": [ BLOCOS ]}], "anexos": [{"titulo": "...", "corpo": [ BLOCOS ]}]
}
```

`identificacao` só no curto. `historico`, `sumario`, `referencias`, `glossario`, `apendices`, `anexos` só no longo (sem `historico`, o gerador cria a linha de emissão inicial). `variante` ("curto"/"longo") é deduzida do tipo; informe só para forçar.

Blocos de `corpo`:

| Bloco | Campos |
|---|---|
| `{"t":"h1","texto":"Objetivo"}` … `h5` | `numerar` (false = sem número), `nova_pagina` (h1 do longo; padrão true) |
| `{"t":"p","texto":"..."}` | `sem_recuo`, `alinhar` ("center", "right") |
| `{"t":"alineas","itens":["...;", {"texto":"...;","sub":["...;"]}, "..."]}` | mesma forma com `"t":"numerada"`, `"checklist"` ou `"marcadores"` |
| `{"t":"tabela","tipo":"Tabela"\|"Quadro","titulo":"","cabecalho":[],"linhas":[[]],"larguras":[pesos],"alinhamento":["left","right"],"fonte":"","nota":""}` | |
| `{"t":"figura","caminho":"img.png","titulo":"","fonte":"","largura_cm":14,"tipo":"Figura"\|"Gráfico"}` | PNG ou JPEG |
| `{"t":"citacao","texto":"...","fonte":"(Silva, 2025, p. 10)."}` | |
| `{"t":"nota","texto":"...","rotulo":"Nota"}` | |
| `{"t":"destaque","texto":"...","rotulo":"Importante","nivel":"info"\|"atencao"\|"erro"}` | |
| `{"t":"campos","itens":[["Rótulo","valor"]],"rotulo_cm":4}` | pares rótulo/valor |
| `{"t":"local_data","data":"27/08/2026"}` | escreve "Cidade da empresa, 27 de agosto de 2026." (sem `data`, usa a emissão) |
| `{"t":"assinaturas","itens":[{"nome":"","cargo":"","documento":""}]}` | |
| `{"t":"codigo","texto":"linha1\nlinha2"}` · `{"t":"quebra"}` | |

## 7. Cuidados

- Um documento por arquivo. Vários atos (ex.: três advertências) = vários documentos com números próprios.
- Documento de terceiros ou de sistemas (nota fiscal, boleto, relatório do ERP) não recebe o timbrado; se precisar, entra como `anexos` de um documento da empresa.
- Ao reformatar documento existente, preserve o sentido e os dados; mudanças de conteúdo vão para o `historico` e para o resumo final.
- No Word, se o texto for editado depois, clique com o botão direito no sumário → Atualizar campo.
- As margens são parte do padrão: não as altere no Word (inclusive na tela de impressão, onde escolher "Estreitas" muda a seção atual).
- O `marca.json` é a única coisa que muda entre organizações. Mudança de visual ou de regra é mudança do padrão, não da marca.
