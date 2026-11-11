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
| `logo` | `logo.png` | PNG ou JPEG, de preferência horizontal e com fundo transparente. Sem logo, o rodapé sai só com o filete |
| `grupo` | `true` | opcional: "empresa" que representa o grupo todo. Não é pessoa jurídica: o cabeçalho traz só o nome e o rodapé leva as logos de todas as empresas ativas |
| `ativa` | `false` | opcional: empresa aposentada. Continua valendo para ler e versionar documento antigo; documento novo recebe aviso |

Monte o arquivo a partir de `assets/marca.exemplo.json`. Campos opcionais da raiz: `organizacao` (nome da lista mestra), `empresa_padrao`, `politica` (código/título da norma interna; vai para a ficha e a lista mestra), `areas` para substituir a lista de áreas, `areas_extra` e `tipos_extra` `{"SIG": {"nome": "...", "variante": "curto"|"longo"}}` para ampliar, `areas_ativas` e `tipos_ativos` para limitar o que se oferece em documento novo (os demais continuam valendo para documento antigo). Várias empresas de um mesmo grupo cabem na lista `empresas`.

**Visual da organização (opcional, `visual`)**: fontes e cores do documento, para quem tem manual de marca. Sem ele, o documento sai no visual do padrão (Arial, texto preto). Com ele: `fonte_texto` e `fonte_titulo` (nomes das famílias), `cor_texto`, `cor_titulo` (títulos, título do documento e da capa), `cor_secundaria` (código, identificação, legendas), `cor_realce` (fundo do cabeçalho de tabela e dos rótulos de quadro), e `fontes`: `[{"familia": "...", "estilo": "regular"|"negrito"|"italico"|"negrito_italico", "arquivo": "fonte.ttf"}]` — TrueType (.ttf), caminho relativo ao marca.json ou data URI; vão embutidas no .docx, então o documento abre certo em computador sem a fonte. `cor_papel`, `cor_cartao` e `cor_linha` só vestem a página do timbrador. As cores dos estados (faixa e marca d'água), a borda, as margens e as linhas das tabelas não mudam. Não invente fontes nem cores: use as do manual de marca do usuário.

Depois, gere a versão de arquivo único com `--embutir-logos marca.json --saida marca_portatil.json` (o logo e as fontes do visual viram texto dentro do JSON) e entregue ao usuário dizendo para anexá-la nas próximas conversas ou deixá-la na pasta conectada. Não invente razão social, CNPJ ou cidade: o que faltar fica `[definir]`.

**Regras da marca**

- Empresa com `razao_social` ou `documento` vazio/`[definir]` só emite MINUTA ou EM_APROVACAO; o gerador recusa OFICIAL/OBSOLETO. Quando o usuário informar os dados, atualize o `marca.json` e entregue a nova versão.
- Cada documento leva **somente o logo da empresa emissora** (no documento do grupo, as logos das empresas ativas). Nunca logos de fornecedores, fabricantes ou parceiros.
- Descubra a emissora pelo conteúdo (quem assina, de quem é o colaborador ou o cliente). Na dúvida, pergunte; sem resposta, use `empresa_padrao` e avise.

**Visual fixo (não alterar):** A4; margens esq. 3, dir. 2, sup. 3,3 (faixa e bloco de controle), inf. 3,2 no curto e na capa e 2,6 nas páginas internas do longo; borda fina preta a 1 cm do papel em todas as páginas; cabeçalho com a faixa do estado (de borda a borda, dentro da borda da página) e o bloco de controle centralizado (código · versão · Página X de Y; frase do estado; razão social · documento · cidade); rodapé com filete na cor da empresa e logo (centralizado no curto e na capa; logo + título nas páginas internas do longo); marca d'água nos estados não oficiais. Nome e CNPJ não se repetem no corpo. Nunca desenhe o timbrado à mão nem edite o gerador para mudar o visual.

## 2. Fluxo

1. **Ler o conteúdo** recebido por inteiro e localizar o `marca.json` (seção 1).
2. **Classificar**: empresa, tipo (Quadro A), área (Quadro B), estado (seção 3). Pergunte só o que não dá para deduzir: empresa, estado desejado, número reservado na lista mestra, signatários, aprovador. Sem resposta: estado **MINUTA**, número vazio (sai **XXX**) e diga isso no final.
3. **Adequar o texto** às regras de redação (seção 5). Não invente fato, data, valor, nome ou cláusula: o que faltar vira `[definir]`.
4. **Escrever a especificação JSON** (seção 6) em arquivo de trabalho; figuras com caminho relativo a esse arquivo.
5. **Gerar**: `python3 scripts/gerar_documento.py especificacao.json --marca marca.json --saida PASTA/ --pdf` (o nome do arquivo sai no padrão). Se o LibreOffice do ambiente precisar de um wrapper, informe o comando em `TIMBRADOR_SOFFICE`. Na publicação (OFICIAL), passe `--lista PASTA_DOS_DOCUMENTOS`: o gerador reserva o próximo número livre, recusa código repetido e registra o documento no `lista-mestra.json`. Minuta e em aprovação não reservam número nem mexem na lista. Arquivos que acompanham o documento (conversa exportada, planilha, foto) entram com `--anexo ARQUIVO` (repetível): veja "Anexos externos" na seção 4.
6. **Tratar erros e avisos**: o gerador recusa OFICIAL sem `aprovado_por`, `data_aprovacao`, número ou dados da empresa, OBSOLETO sem substituto ou motivo, plano (PLN) sem `projeto` e as regras do relatório de condução (seção 4). Não contorne: rebaixe o estado ou peça o dado. Repasse os avisos relevantes.
7. **Conferir o visual** quando o PDF sair: `pdftoppm -png -r 60` e olhe a primeira página e uma interna (faixa, bloco de controle, logo no rodapé, borda, marca d'água, tabelas cortadas, títulos órfãos). Sem LibreOffice, o sumário do longo sai sem números e o Word pede para atualizar os campos ao abrir; avise.
8. **Entregar** o .docx (na pasta conectada, se houver, ao lado da origem) com uma linha: empresa, tipo, código, versão, estado e pendências (`[definir]`, número XXX, aprovação).

## 3. Estados

| Estado (JSON) | Quando | Visual gerado |
|---|---|---|
| `MINUTA` | Texto que ainda não é oficial, proposto para análise ou concordância das partes (padrão) | Faixa ocre + marca d'água |
| `EM_APROVACAO` | Texto fechado, entregue a quem aprova; só muda para atender aos ajustes pedidos. A versão oficial anterior continua valendo | Faixa azul + marca d'água |
| `OFICIAL` | Aprovado (exige `aprovado_por`, `data_aprovacao`, número e empresa completa) | Faixa verde "DESDE" a vigência, sem marca d'água, aprovação no cabeçalho |
| `OBSOLETO` | Substituído ou cancelado (exige `substituido_por` ou `motivo_obsoleto`) | Faixa vermelha + marca d'água |

`RASCUNHO` e `EM_REVISAO` estão aposentados: o gerador ainda os lê (documentos antigos) e avisa; a rodada seguinte de um arquivo antigo sai como MINUTA e EM_APROVACAO. Não use em documento novo.

- Só marque OFICIAL quando o usuário informar quem aprovou e quando. "Deixa oficial" sem esses dados → EM_APROVACAO e explique.
- Minuta antes do oficial em contratos, termos, acordos com colaboradores, comunicados a terceiros, políticas e regulamentos.
- O título diz só o assunto: o gerador tira o nome do tipo do começo ("Política de compras" → "Compras") e avisa.

**Documento já timbrado** (o .docx guarda a especificação numa ficha embutida):

- `--ler DOCX --json-saida espec.json` devolve a especificação (ficha + corpo relido do arquivo) para reformatar ou corrigir.
- `--atualizar DOCX --mudanca conteudo|correcao|rodada --o-que-mudou "..." --json-saida nova.json` prepara a próxima versão: de um oficial X.Y sai (X+1).0-r1 (`conteudo`, o padrão) ou X.(Y+1)-r1 (`correcao`) em MINUTA, com `substitui` e o histórico; de um não oficial, a rodada seguinte (`rodada`, o padrão). Na minuta de nova versão (com `substitui`), `conteudo` e `correcao` contam da versão oficial substituída: a minuta 2.0-r1 que substitui a v1.0 vira 1.1-r1 com `correcao`; se a versão alvo não muda, segue a rodada seguinte. Edite só o corpo e gere; não mexa no bloco `anterior`. Ao gerar a nova versão como OFICIAL, sai também a cópia OBSOLETO da anterior.
- `--obsoletar DOCX --substituido-por "CÓDIGO vX.Y"` gera a cópia obsoleta de um oficial.
- `--inventariar PASTA [--conferir]` varre a pasta dos documentos e monta ou confere a lista mestra. `Modelos` fica fora; `Obsoletos` é lida, mas só decide a entrada de um código que não tem mais versão em uso, e o relatório avisa oficial que já tem cópia obsoleta, mais de uma versão oficial em uso, cópia obsoleta fora de `Obsoletos`, documento de projeto fora da pasta do projeto e entrada da lista cujo arquivo sumiu.
- `--atualizar REL.docx --rodada rodada.json`: acrescenta uma rodada ao relatório de condução (ver abaixo).

**Projeto**: documento de projeto, de qualquer tipo, leva `"projeto": "Nome"` (vai para a ficha, a lista mestra e a pasta `Projetos/Nome`). Plano (PLN) sem projeto é recusado.

**Anexos externos**: `--anexo ARQUIVO` (repetível) põe o documento e os anexos numa pasta com o código (minuta: código com XXX e o título), com os nomes `Anexo 01 – nome original`; o documento lista os anexos antes do local e data, com tamanho e SHA-256. Arquivo repetido entra uma vez. Na nova versão (`--atualizar`), os anexos registrados na ficha precisam estar na pasta com o mesmo SHA-256, senão o gerador recusa. Não confunda com `anexos` (anexos ABNT, dentro do documento longo).

## 4. Identificação

**Código**: `SIGLA-TIPO-ÁREA-NNN`; tipos curtos acrescentam o ano: `ACME-CI-RH-012/2026`. O gerador monta a partir de `empresa`, `tipo`, `area`, `numero`, `ano`. A numeração é por empresa, tipo e área (e ano, nos curtos). Minuta e em aprovação saem com **XXX**: o número é reservado na publicação, não é pendência.

**Versão**: oficial `X.Y` (X muda regra/processo; Y só correção). Não oficial = versão-alvo + `-rN`: `1.0-r1`, `1.0-r2`; revisão de um 1.0 → `2.0-r1` ou `1.1-r1`. Aprovado, o sufixo cai (o gerador remove). Sem `-r` num estado não oficial, o gerador acrescenta `-r1`.

**Nome do arquivo**: `CÓDIGO_vVERSÃO_ESTADO_titulo.docx` (barra do código vira hífen).

**Classificação**: Público, Interno (padrão), Confidencial. Confidencial exige `acesso_restrito` (nomes ou cargos, separados por vírgula: quem pode ler); a palavra CONFIDENCIAL sai na linha de controle de todas as páginas, inclusive nos telegramas, e a lista de acesso aparece abaixo do título (curto) ou na folha de controle (longo). Na página, com Confidencial marcado, nada do documento fica guardado no navegador (rascunho e anexos). **Onde fica o confidencial**: numa pasta chamada `Confidencial`, com permissão só para quem pode ler. Grave o .docx lá (`--saida` dentro dela; fora dela o gerador avisa); na página, ela é conectada à parte e recebe o documento, a cópia obsoleta e os anexos. A lista mestra fica na pasta comum, aberta a todos, e do confidencial guarda só um esqueleto (código, versão, estado e data, `"titulo": "Documento confidencial"`, `"confidencial": true`). Os inventários (gerador, página e Inventariar.ps1) pulam qualquer pasta `Confidencial` e acusam documento confidencial achado fora dela.

Quadro A – Tipos padrão (a variante define o formato; `tipos_extra` do marca.json acrescenta outros e `tipos_ativos` limita os de documento novo)

| Curto (sem capa, texto direto após o título) | Longo (capa, folha de controle, sumário e páginas ABNT) |
|---|---|
| RO registro de ocorrência · CI comunicação interna · COM comunicado · OFI ofício · DEC declaração · TER termo · ATA ata · AUT autorização · NOT notificação/advertência · CTR contrato/acordo · PRO proposta · FOR formulário · REC recibo | MAN manual · POL política · REG regulamento · POP procedimento operacional padrão · INS instrução de trabalho · PLN plano · PRJ projeto · REL relatório |

Quadro B – Áreas padrão (substituíveis por `areas` no marca.json): ADM administrativo · DIR diretoria · FIN financeiro · FIS financeiro e fiscal · JUR jurídico · RH recursos humanos · VEN vendas e comercial · ECO e-commerce e marketplaces · EST estoque · EXP expedição · CMP compras · OFC oficina e assistência técnica · MKT marketing · TI tecnologia. Use só as áreas ativas que o `--checar-marca` conta; se nada servir, use a mais próxima e avise (ou proponha incluir em `areas_extra`).

### Relatório de condução de projeto (REL com rodadas)

Um REL com `rodadas` registra a condução de um projeto, reunião a reunião. REL sem `rodadas` continua sendo um relatório comum.

- **Um por projeto, com o número do plano**: o do `ACME-PLN-ECO-001` é o `ACME-REL-ECO-001`. O número é herdado, não reservado, e o relatório só se abre com o plano **oficial** na lista mestra: sempre gere com `--lista`.
- **Sempre OFICIAL, pela coordenação do projeto** (`aprovado_por`), sem minuta nem em aprovação. A abertura é a v1.0, com a rodada 0 (`"tipo": "abertura"`, início da execução); cada rodada publicada é a versão seguinte (1.1, 1.2…).
- **Só se acrescentam linhas.** As rodadas ficam na ficha e o Quadro de rodadas do corpo (cabeçalho `Nº · Data · O que se fez · Compromisso seguinte · Avanço`) é refeito a partir delas; se o corpo não tiver o quadro, a seção "Rodadas" entra no fim. A lista mestra guarda o SHA-256 de cada rodada, e o gerador recusa a versão nova se uma rodada publicada mudou (no Word ou na ficha) ou se o arquivo não é a versão vigente. Erro se corrige numa rodada nova, do tipo `correcao`.
- **Sem cópia obsoleta**: a versão nova contém a anterior inteira (`--obsoletar` recusa REL). Oriente a retirar a versão anterior da pasta.
- **Três rodadas seguidas sem avanço**: o gerador avisa que a parada do projeto pode ser declarada; declarar é decisão da coordenação, registrada como rodada `parada`.
- **Encerramento**: a última rodada cita o termo de encerramento (TER); depois dela, o relatório não recebe rodadas.

Rodada (`--rodada`; sem `n` e `data`, o gerador põe o próximo número e a data de hoje):

| Campo | Quando |
|---|---|
| `n`, `data` (dd/mm/aaaa), `tipo`, `feito` | sempre; `tipo`: `abertura` (só a 0), `rodada`, `replanejamento`, `parada`, `retomada`, `correcao`, `encerramento` |
| `compromisso`: `{"oque", "quem", "ate"}` | obrigatório em `abertura`, `rodada`, `replanejamento` e `retomada` |
| `avanco`: `"sim"`, ou `"nao"` com `travou` e `depende` | obrigatório em `rodada` |
| `plano_versao` · `corrige` · `termo` e `situacao` (`concluído`/`cancelado`) | `replanejamento` · `correcao` · `encerramento` |

Exemplo de abertura: `testes/espec-rel.json`. Na página, o relatório novo se abre marcando "Relatório de condução de projeto"; um REL com rodadas abre no modo **Acrescentar rodada**, com as rodadas publicadas travadas.

## 5. Redação e ABNT (o gerador cuida da formatação; você cuida do conteúdo)

- Linguagem formal, impessoal, frases diretas. Datas dd/mm/aaaa; horas 07h30; valores R$ 1.234,56.
- **Seções numeradas** (NBR 6024) com `h1`…`h5`; não escreva o número no título, o gerador numera. Documento longo começa por "Objetivo" ou "Introdução". Não pule nível (h3 só dentro de h2).
- **Enumerações** em `alineas`: começam em minúscula, terminam em `;` e a última em `.`. Subitens com `sub`. Passos em ordem → `numerada`. Verificação → `checklist`. Evite `marcadores`.
- **Tabela** (dados numéricos, laterais abertas) × **Quadro** (texto, fechado): sempre `titulo`; sem `fonte`, o gerador escreve "Elaborado pela empresa (AAAA).". Cite no texto ("conforme Quadro 1").
- **Citações** (NBR 10520:2023): autor-data com sobrenome só com inicial maiúscula – (Silva, 2025, p. 10). Mais de 3 linhas → bloco `citacao`.
- **Referências** (NBR 6023:2025) em `referencias` (o gerador ordena): `SOBRENOME, Nome. **Título**: subtítulo. Local: Editora, ano.`; online acrescenta `Disponível em: URL. Acesso em: 15 set. 2026.`
- **Assinatura entra só pela página** (caixa "Com assinatura", que põe no fim o local, a data e as assinaturas e troca as linhas por tabela quando elas ficariam sozinhas numa página). O gerador não põe assinatura em documento nenhum: tira do corpo o bloco `assinaturas` (e o local e data que o antecede) e guarda essas assinaturas e os `envolvidos` como `assinaturas_previstas` na ficha; ao abrir o arquivo na página, eles já vêm preenchidos. Diga isso ao usuário quando o documento pedir assinatura. Documento longo oficial: a aprovação vai na folha de controle automaticamente.
- Não repita razão social e documento da emissora em `identificacao`; nas assinaturas eles continuam valendo.
- Destaque: `**negrito**`, `*itálico*`, `***ambos***` dentro de qualquer texto; `\n\n` separa parágrafos.

## 6. Especificação JSON

```json
{
  "empresa": "ACME",
  "tipo": "POP", "area": "ECO", "numero": 3, "ano": 2026,
  "titulo": "Cadastro de anúncios", "subtitulo": "opcional", "titulo_curto": "opcional, rodapé das páginas internas",
  "estado": "MINUTA", "versao": "1.0-r1", "data_emissao": "15/09/2026",
  "classificacao": "Interno", "acesso_restrito": "só se Confidencial", "area_nome": "opcional, sobrescreve o nome da área", "projeto": "opcional; obrigatório no PLN",
  "elaborado_por": "", "revisado_por": "", "aprovado_por": "", "cargo_aprovador": "", "data_aprovacao": "",
  "vigencia": "", "substitui": "", "substituido_por": "", "motivo_obsoleto": "",
  "identificacao": [["Destinatário", "..."], ["Assunto", "..."]],
  "historico": [{"versao": "1.0-r1", "data": "15/09/2026", "descricao": "Emissão inicial.", "responsavel": "..."}],
  "sumario": true,
  "com_assinatura": false, "envolvidos": [{"nome": "", "cargo": "", "papel": "Elaborou|Revisou|Aprovou|De acordo|Ciente|Testemunha|Contratante|Contratada"}],
  "assinatura_formato": "auto|linhas|tabela",
  "corpo": [ BLOCOS ],
  "referencias": ["..."], "glossario": [["Termo", "definição."]],
  "apendices": [{"titulo": "...", "corpo": [ BLOCOS ]}], "anexos": [{"titulo": "...", "corpo": [ BLOCOS ]}],
  "telegramas": [{"titulo": "...", "corpo": [ BLOCOS ]}]
}
```

`identificacao` só no curto. `historico`, `sumario`, `referencias`, `glossario`, `apendices`, `anexos` só no longo (sem `historico`, o gerador cria a linha de emissão inicial). `variante` ("curto"/"longo") é deduzida do tipo; informe só para forçar. `com_assinatura` e `envolvidos` não viram assinatura no gerador: ficam como `assinaturas_previstas` (seção 5). O relatório de condução acrescenta `rodadas` (seção 4).

`telegramas` (curto e longo): anexos curtos no fim do documento, cada um em meia folha A5, sem código próprio. Herdam o código, a versão e o estado do documento, que saem numa linha no topo com a página. Títulos (`h1`…`h5`) dentro do telegrama saem em negrito, sem número. Na página, a linha "TELEGRAMA – título" no texto colado abre um telegrama.

Blocos de `corpo`:

| Bloco | Campos |
|---|---|
| `{"t":"h1","texto":"Objetivo"}` … `h5` | `numerar` (false = sem número), `nova_pagina` (h1 do longo; padrão true) |
| `{"t":"p","texto":"..."}` | `sem_recuo`, `alinhar` ("center", "right") |
| `{"t":"alineas","itens":["...;", {"texto":"...;","sub":["...;"]}, "..."]}` | mesma forma com `"t":"numerada"`, `"checklist"` ou `"marcadores"` |
| `{"t":"tabela","tipo":"Tabela"\|"Quadro","titulo":"","cabecalho":[],"linhas":[[]],"larguras":[pesos],"alinhamento":["left","right"],"fonte":"","nota":""}` | |
| `{"t":"figura","caminho":"/caminho/img.png","titulo":"","fonte":"","largura_cm":14,"tipo":"Figura"\|"Gráfico"}` | PNG ou JPEG; caminho absoluto ou relativo à pasta da especificação |
| `{"t":"citacao","texto":"...","fonte":"(Silva, 2025, p. 10)."}` | |
| `{"t":"nota","texto":"...","rotulo":"Nota"}` | |
| `{"t":"destaque","texto":"...","rotulo":"Importante","nivel":"info"\|"atencao"\|"erro"}` | |
| `{"t":"campos","itens":[["Rótulo","valor"]],"rotulo_cm":4}` | pares rótulo/valor |
| `{"t":"local_data","data":"27/08/2026"}` | escreve "Cidade da empresa, 27 de agosto de 2026." (sem `data`, usa a emissão) |
| `{"t":"assinaturas","itens":[{"nome":"","cargo":"","papel":"","documento":""}],"formato":"linhas"\|"tabela"}` | no gerador vira `assinaturas_previstas` (seção 5) |
| `{"t":"codigo","texto":"linha1\nlinha2"}` · `{"t":"quebra"}` | |

## 7. Cuidados

- Um documento por arquivo. Vários atos (ex.: três advertências) = vários documentos com números próprios.
- Documento de terceiros ou de sistemas (nota fiscal, boleto, relatório do ERP) não recebe o timbrado; se precisar, entra como anexo externo (`--anexo`) de um documento da empresa.
- Ao reformatar documento existente, preserve o sentido e os dados; mudanças de conteúdo vão para o `historico` e para o resumo final.
- No Word, se o texto for editado depois, clique com o botão direito no sumário → Atualizar campo.
- As margens são parte do padrão: não as altere no Word (inclusive na tela de impressão, onde escolher "Estreitas" muda a seção atual).
- O `marca.json` é a única coisa que muda entre organizações. Fora o que o `visual` permite (fontes e cores de texto), mudança de visual ou de regra é mudança do padrão, não da marca.
