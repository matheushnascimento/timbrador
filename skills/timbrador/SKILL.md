---
name: timbrador
description: "Timbrador white label: converte qualquer conteúdo em documento Word (.docx) no Padrão de Documentação Oficial (ABNT + papel timbrado da organização: logo, cor, borda, faixa de estado e bloco de controle), com código, versão, ficha embutida e lista mestra. A identidade vem de um arquivo marca.json, então serve para qualquer empresa ou grupo. Use sempre que o usuário quiser timbrar, formalizar, padronizar ou atualizar a versão de um documento — CI, comunicado, ofício, termo, contrato, ata, registro de ocorrência, proposta, manual, política, POP, regulamento, plano ou relatório — ou configurar o papel timbrado da própria empresa, mesmo que não diga 'timbrado'."
---

# Timbrador (white label)

Transforma qualquer informação em um .docx que qualquer pessoa reconhece de imediato como documento oficial da organização e cujo estado (oficial ou não) é inequívoco. A identidade visual e os dados cadastrais vêm do `marca.json`; o visual, as regras e a numeração são fixos. Saída: **somente .docx** (e PDF de conferência, se pedido). Responda no idioma do usuário; o documento sai em português.

Arquivos desta skill:

| Arquivo | Papel |
|---|---|
| `scripts/gerar_documento.py` | Gerador. Só biblioteca padrão do Python 3.8+. Rode direto daqui; ele só escreve na pasta de `--saida` (e na lista mestra, com `--lista`). |
| `assets/timbrador.html` | Página para os colaboradores timbrarem sozinhos um .docx/.txt/.md pronto. Mesmo visual e mesmas regras do gerador. |
| `assets/Abrir-Timbrador.cmd` | Atalho do Windows que abre a página no Edge ou no Chrome (os navegadores que gravam direto na pasta). |
| `assets/marca.exemplo.json` + `assets/logo-exemplo.png` | Modelo para criar o `marca.json`. |
| `assets/marca.json` (opcional) | Se quem distribuiu a skill já a personalizou, a marca fica aqui. |

## 1. Marca (configuração da organização)

**Onde procurar o `marca.json`** (pare no primeiro): anexos da conversa → pasta conectada do usuário → `assets/marca.json` desta skill. Rode `python3 scripts/gerar_documento.py --checar-marca CAMINHO` e mostre o resumo ao usuário na primeira vez.

**Se não existir, configure antes de gerar.** Pergunte (AskUserQuestion, se disponível) o que faltar para cada empresa emissora:

| Campo | Exemplo | Observação |
|---|---|---|
| `sigla` | `ACME` | 2 a 6 letras/números, começando por letra; abre o código de todo documento |
| `marca` | `Acme` | nome fantasia |
| `razao_social` | `Acme Comércio Ltda` | `[definir]` enquanto não souber |
| `documento_rotulo` / `documento` | `CNPJ` / `00.000.000/0001-00` | pode ser `CPF` para autônomo |
| `cidade` | `Campinas – SP` | usada no cabeçalho, na capa e no local e data |
| `cor` | `1F3864` | hexadecimal; filete do rodapé, tipo na capa. Se o usuário não souber, extraia a cor dominante do logo |
| `logo` | `logo.png` | PNG ou JPEG, de preferência horizontal e com fundo transparente. Sem logo, o rodapé sai só com o filete |

Campos opcionais: `organizacao`, `empresa_padrao`, `politica` (`codigo` e `titulo` da norma interna; o código vai para a ficha e para a lista mestra), `areas` (troca a lista de áreas), `areas_extra` e `tipos_extra` (`{"SIG": {"nome": "...", "variante": "curto"|"longo"}}`), `tipos_ativos` e `areas_ativas` (o que a página oferece para documento novo; vazio = todos). Por empresa: `"grupo": true` para o documento do grupo (vale para várias empresas; não leva documento no cabeçalho e o rodapé traz os logos de todas as empresas ativas) e `"aposentada": true` para empresa que só serve para ler e versionar documentos antigos.

Depois, gere a versão de arquivo único com `--embutir-logos marca.json --saida marca_portatil.json` (o logo vira texto dentro do JSON) e entregue ao usuário dizendo para anexá-la nas próximas conversas ou deixá-la na pasta conectada. Para os colaboradores, `--montar-pagina marca.json --saida PASTA/` grava uma cópia de `timbrador.html` com a marca embutida (sem ela, a página pede o marca.json portátil e o guarda no navegador). Não invente razão social, documento ou cidade: o que faltar fica `[definir]`.

**Regras da marca**

- Empresa com `razao_social` ou `documento` vazio/`[definir]` só emite MINUTA ou EM_APROVACAO; o gerador recusa OFICIAL/OBSOLETO. Quando o usuário informar os dados, atualize o `marca.json` e entregue a nova versão.
- Cada documento leva **somente o logo da empresa emissora** (o documento do grupo leva os das empresas do grupo). Nunca logos de fornecedores, fabricantes ou parceiros.
- Descubra a emissora pelo conteúdo (quem assina, de quem é o colaborador ou o cliente). Na dúvida, pergunte; sem resposta, use `empresa_padrao` e avise.

**Visual fixo (não alterar):** A4; margens 3,3 cm no alto (faixa + bloco de controle), 2 cm à direita, 3 cm à esquerda e 3,2 cm na base (2,6 cm nas páginas internas do documento longo); faixa do estado no topo de todas as páginas; bloco de controle centralizado (código · versão · Página X de Y; situação; razão social · documento · cidade); marca d'água nos estados não oficiais; borda fina preta a 1 cm do papel; rodapé com filete na cor da empresa e logo (centralizado no curto e na capa; logo + título curto nas páginas internas do longo). Razão social e documento não se repetem no corpo. Nunca desenhe o timbrado à mão nem edite o gerador para mudar o visual: gerador e página andam juntos.

## 2. Fluxo

1. **Ler o conteúdo** recebido por inteiro e localizar o `marca.json` (seção 1). Se for um .docx já timbrado, use `--ler` (documento novo a partir dele) ou `--atualizar` (nova versão do mesmo documento, seção 4).
2. **Classificar**: empresa, tipo (Quadro A), área (Quadro B), estado (seção 3). Pergunte só o que não dá para deduzir: empresa, estado desejado, aprovador. Sem resposta: **MINUTA**, número vazio (sai **XXX**: o número é reservado na publicação) e diga isso no final.
3. **Adequar o texto** às regras de redação (seção 6). Não invente fato, data, valor, nome ou cláusula: o que faltar vira `[definir]`.
4. **Escrever a especificação JSON** (seção 7) em arquivo de trabalho; figuras com caminho relativo a esse arquivo.
5. **Gerar**: `python3 scripts/gerar_documento.py especificacao.json --marca marca.json --saida PASTA/ --pdf` (o nome do arquivo sai no padrão). Na publicação (OFICIAL), passe `--lista PASTA_DOS_DOCUMENTOS` para reservar o número e registrar na lista mestra. Se o LibreOffice do ambiente precisar de um wrapper, informe o comando em `TIMBRADOR_SOFFICE`.
6. **Tratar erros e avisos**: o gerador recusa OFICIAL sem `aprovado_por`, `data_aprovacao`, número ou dados da empresa, OBSOLETO sem substituto ou motivo, e número já usado por outro título na lista mestra. Não contorne: rebaixe o estado ou peça o dado. Repasse os avisos relevantes.
7. **Conferir o visual** quando o PDF sair: `pdftoppm -png -r 60` e olhe a primeira página e uma interna (faixa, bloco de controle, logo no rodapé, borda, marca d'água, tabelas cortadas, títulos órfãos). Sem LibreOffice, o sumário do longo sai sem números e o Word pede para atualizar os campos ao abrir; avise.
8. **Entregar** o .docx (na pasta conectada, se houver, ao lado da origem) com uma linha: empresa, tipo, código, versão, estado e pendências (`[definir]`, número XXX, aprovação).

## 3. Estados

| Estado (JSON) | Quando | Visual gerado |
|---|---|---|
| `MINUTA` | Texto completo proposto para análise ou concordância das partes | Faixa amarela + marca d'água |
| `EM_APROVACAO` | Texto fechado, entregue a quem aprova; só muda para atender aos ajustes pedidos | Faixa azul + marca d'água |
| `OFICIAL` | Aprovado e vigente (exige `aprovado_por`, `data_aprovacao`, número e empresa completa) | Faixa verde "DOCUMENTO OFICIAL · VIGENTE · DESDE …", sem marca d'água, aprovação no cabeçalho |
| `OBSOLETO` | Substituído ou cancelado (exige `substituido_por` ou `motivo_obsoleto`) | Faixa vermelha + marca d'água |

- `RASCUNHO` e `EM_REVISAO` são nomes antigos: o gerador ainda os lê e versiona (a rodada seguinte sai como MINUTA e EM_APROVACAO), mas não os use em documento novo.
- Só marque OFICIAL quando o usuário informar quem aprovou e quando. "Deixa oficial" sem esses dados → EM_APROVACAO e explique.
- Minuta e em aprovação saem com número XXX e não alteram a lista mestra: o número é reservado na publicação.

## 4. Identificação, versões e lista mestra

**Código**: `SIGLA-TIPO-ÁREA-NNN`; tipos curtos acrescentam o ano: `ACME-CI-RH-012/2026`. O gerador monta a partir de `empresa`, `tipo`, `area`, `numero`, `ano`. A numeração é por empresa, tipo e área (e ano, nos curtos).

**Título**: diz só o assunto; o tipo já sai impresso acima dele. O gerador tira "Procedimento operacional de…", "Termo de…" etc. do começo do título e avisa.

**Versão**: oficial `X.Y` (X muda regra/processo; Y só correção). Não oficial = versão-alvo + `-rN`: `1.0-r1`, `1.0-r2`; revisão de um 1.0 → `2.0-r1` ou `1.1-r1`. Aprovado, o sufixo cai (o gerador remove). Sem `-r` num estado não oficial, o gerador acrescenta `-r1`.

**Nome do arquivo**: `CÓDIGO_vVERSÃO_ESTADO_titulo.docx` (barra do código vira hífen).

**Ficha**: todo .docx leva a especificação dentro (customXml) e propriedades visíveis em Arquivo > Informações. É ela que permite reler o documento.

**Comandos de controle**:

- `--ler DOCX [--json-saida X.json] [--pasta-midia PASTA]`: devolve a especificação do documento (ficha + corpo relido; vale o que foi editado no Word).
- `--atualizar DOCX --o-que-mudou "…" [--mudanca conteudo|correcao] [--estado …] [--responsavel …]`: prepara a nova versão (oficial 1.0 → 2.0-r1 ou 1.1-r1 em minuta; minuta r2 → r3), com o histórico e o bloco `anterior`. Edite só o corpo e gere; ao gerar a versão OFICIAL, sai também a cópia OBSOLETO da anterior. Não altere o bloco `anterior`: o gerador recusa.
- `--obsoletar DOCX --substituido-por "CÓDIGO vX.Y"`: cópia OBSOLETO de um documento oficial.
- `--lista PASTA`: confere o número no `lista-mestra.json` (recusa número já usado por outro título; sem número, reserva o próximo livre) e registra a publicação.
- `--inventariar PASTA [--conferir] [--ignorar NOME]`: varre a pasta dos documentos (subpastas por tipo, `Obsoletos`, `Projetos/NOME`, `Modelos` ignorada) e monta a lista mestra do que já existe.
- `--anexo ARQUIVO` (repetível): anexos externos; documento e anexos vão para uma pasta com o código, e o documento lista os anexos com tamanho e SHA-256.

**Classificação**: Público, Interno (padrão), Confidencial.

Quadro A – Tipos padrão (a variante define o formato; `tipos_extra` do marca.json acrescenta outros)

| Curto (sem capa, texto direto após o título) | Longo (capa, folha de controle, sumário e páginas ABNT) |
|---|---|
| CI comunicação interna · COM comunicado · OFI ofício · DEC declaração · TER termo · ATA ata · RO registro de ocorrência · AUT autorização · NOT notificação · CTR contrato · PRO proposta · FOR formulário · REC recibo | MAN manual · POL política · REG regulamento · POP procedimento operacional padrão · INS instrução de trabalho · PLN plano · PRJ projeto · REL relatório |

Quadro B – Áreas padrão (substituíveis por `areas` no marca.json): ADM administrativo · DIR diretoria · FIN financeiro · FIS financeiro e fiscal · JUR jurídico · RH recursos humanos · VEN vendas e comercial · ECO e-commerce e marketplaces · EST estoque · EXP expedição · CMP compras · OFC oficina e assistência técnica · MKT marketing · TI tecnologia. Use as áreas que o `--checar-marca` lista; se nada servir, use a mais próxima e avise (ou proponha incluir em `areas_extra`).

## 5. Assinaturas

**Assinatura entra só pela página**, na caixa "Com assinatura" (formato automático: linhas de assinatura; se ficariam sozinhas numa página, tabela Nome, Cargo, Papel, Assinatura, Data). O gerador não põe assinatura em documento nenhum: tira do corpo o bloco `assinaturas` (e o local e data que o antecede) e guarda os envolvidos como `assinaturas_previstas` na ficha. Ao abrir o arquivo na página, eles já vêm preenchidos. Diga isso ao usuário quando o documento pedir assinatura. A aprovação do documento longo oficial vai na folha de controle automaticamente.

## 6. Redação e ABNT (o gerador cuida da formatação; você cuida do conteúdo)

- Linguagem formal, impessoal, frases diretas. Datas dd/mm/aaaa; horas 07h30; valores R$ 1.234,56.
- **Seções numeradas** (NBR 6024) com `h1`…`h5`; não escreva o número no título, o gerador numera. Documento longo começa por "Objetivo" ou "Introdução". Não pule nível (h3 só dentro de h2).
- **Enumerações** em `alineas`: começam em minúscula, terminam em `;` e a última em `.`. Subitens com `sub`. Passos em ordem → `numerada`. Verificação → `checklist`. Evite `marcadores`.
- **Tabela** (dados numéricos, laterais abertas) × **Quadro** (texto, fechado): sempre `titulo` e `fonte` ("Elaborado pela empresa (AAAA)." é o padrão quando for próprio). Cite no texto ("conforme Quadro 1").
- **Citações** (NBR 10520:2023): autor-data com sobrenome só com inicial maiúscula – (Silva, 2025, p. 10). Mais de 3 linhas → bloco `citacao`.
- **Referências** (NBR 6023:2025) em `referencias`, já na ordem alfabética: `SOBRENOME, Nome. **Título**: subtítulo. Local: Editora, ano.`; online acrescenta `Disponível em: URL. Acesso em: 15 set. 2026.`
- Não repita razão social e documento da emissora em `identificacao`.
- Destaque: `**negrito**`, `*itálico*`, `***ambos***` dentro de qualquer texto.

## 7. Especificação JSON

```json
{
  "empresa": "ACME",
  "tipo": "POP", "area": "ECO", "numero": 3, "ano": 2026,
  "titulo": "Cadastro de anúncios", "subtitulo": "opcional", "titulo_curto": "opcional, rodapé das páginas internas",
  "estado": "MINUTA", "versao": "1.0-r1", "data_emissao": "15/09/2026",
  "classificacao": "Interno", "area_nome": "opcional, sobrescreve o nome da área", "projeto": "opcional",
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

`identificacao` só no curto. `historico`, `sumario`, `referencias`, `glossario`, `apendices`, `anexos` (anexos ABNT, dentro do documento; os arquivos externos vão por `--anexo`) só no longo. Sem `historico`, o gerador cria a linha de emissão inicial; no OFICIAL acrescenta "Aprovada por …". Sem `vigencia`, o oficial vale desde a data de aprovação. `variante` ("curto"/"longo") é deduzida do tipo; informe só para forçar.

Blocos de `corpo`:

| Bloco | Campos |
|---|---|
| `{"t":"h1","texto":"Objetivo"}` … `h5` | `numerar` (false = sem número), `nova_pagina` (h1 do longo; padrão true) |
| `{"t":"p","texto":"..."}` | `sem_recuo`, `alinhar` ("center", "right") |
| `{"t":"alineas","itens":["...;", {"texto":"...;","sub":["...;"]}, "..."]}` | mesma forma com `"t":"numerada"`, `"checklist"` ou `"marcadores"` |
| `{"t":"tabela","tipo":"Tabela"\|"Quadro","titulo":"","cabecalho":[],"linhas":[[]],"larguras":[pesos],"alinhamento":["left","center"],"fonte":"","nota":""}` | |
| `{"t":"figura","caminho":"img.png","titulo":"","fonte":"","largura_cm":14,"tipo":"Figura"\|"Gráfico"}` | PNG ou JPEG |
| `{"t":"citacao","texto":"...","fonte":"(Silva, 2025, p. 10)."}` | |
| `{"t":"nota","texto":"...","rotulo":"Nota"}` | |
| `{"t":"destaque","texto":"...","rotulo":"Importante","nivel":"info"\|"atencao"\|"erro"}` | |
| `{"t":"campos","itens":[["Rótulo","valor"]],"rotulo_cm":3.3}` | pares rótulo/valor |
| `{"t":"local_data","data":"27/08/2026"}` | "Cidade da empresa, 27 de agosto de 2026." (sem `data`, usa a emissão) |
| `{"t":"codigo","texto":"linha1\nlinha2"}` · `{"t":"quebra"}` | |

## 8. Cuidados

- Um documento por arquivo. Vários atos (ex.: três advertências) = vários documentos com números próprios.
- Documento de terceiros ou de sistemas (nota fiscal, boleto, relatório do ERP) não recebe o timbrado; se precisar, entra como anexo externo de um documento da empresa.
- Ao reformatar documento existente, preserve o sentido e os dados; mudanças de conteúdo vão para o `historico` e para o resumo final.
- No Word, se o texto for editado depois, clique com o botão direito no sumário → Atualizar campo.
- As margens são parte do padrão: não as altere no Word (inclusive na tela de impressão, onde escolher "Estreitas" muda a seção atual).
- O `marca.json` é a única coisa que muda entre organizações. Mudança de visual ou de regra é mudança do padrão, não da marca, e vale para o gerador e para a página.
