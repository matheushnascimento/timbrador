# A página, o inventário e o relatório de condução

> Movido do `CLAUDE.md` em 10/11/2026 (adoção do BMad, `bmad-project-context`). Texto mantido como estava; leia antes de editar `pagina/timbrador.html`, `pagina/Inventariar.ps1` ou as regras do relatório de condução.

## A página (`pagina/timbrador.html`)


Porte de `original/timbrador.html` (fora do git, com dados reais). Decisões do
usuário (20/10/2026): o bloco `gerador-python` **saiu** da página; onde página e
Python embutido divergiam, **o gerador venceu nos sete pontos** (ver seção 4 de
`docs/divergencias-original.md`).

- Núcleo (`<script id="nucleo">`): `carregarMarca()` é o `carregar_marca()` do
  gerador; montagem (`montar`, `blocos`, estilos, numeração, ficha) é porte
  linha a linha do Python — mesma ordem de ids, rels, mídia (`img1.png`),
  `round()` com meio para o par (`arred`), ficha em `json.dumps` (`jsonPy`).
  Mexeu num, mexa no outro e rode `testes/pagina.sh`.
- O que não dá para portar: números de página do sumário (o Word atualiza ao
  abrir), assinatura automática (a página estima a paginação; o gerador
  renderiza), bloco `figura` (vira aviso). Logo só em data URI (`--embutir-logos`).
- Prévia em PDF (`<script id="pdf">`, 03/11/2026): o botão **Visualizar** abre o
  documento em PDF na metade direita da tela (o formulário vai para a esquerda) e
  acompanha as mudanças do formulário. `gerarPdf(spec, conteudo)` recebe o mesmo
  pedido do Gerar (`montarPedido` + `corpoDoPedido`) e desenha, sem biblioteca, o
  leiaute do `montar` em pontos: faixa, controle, borda, marca d'água, rodapés,
  capa, folha de controle, sumário (já com as páginas) e os blocos de conteúdo.
  Texto em Helvetica (métricas AFM embutidas, WinAnsi; fora dela sai "?"), JPEG
  direto e PNG decodificado em RGB + máscara alfa. É estimativa: paginação e quebra
  de linha podem diferir um pouco do Word, o `visual` só troca cores (as fontes do
  `visual` não entram no PDF) e `figura` não aparece, como no .docx da página.
  Pendências do formulário não impedem a prévia: aparecem acima dela. A CSP do
  nginx libera `frame-src blob:` e `object-src blob:` para o leitor de PDF do
  navegador. A borda do formulário tem a cor do estado escolhido (`--estado`).
  Mexeu no `montar`/`blocos`, confira a prévia (`testes/pagina.sh` gera os PDFs).
- Modo escuro (03/11/2026): as cores da página são variáveis no `:root`, redefinidas
  em `@media (prefers-color-scheme:dark)` e em `:root[data-tema="escuro"]`; o botão
  `#tema` cicla automático → claro → escuro e guarda em `localStorage`
  (`timbrador.tema`), lido por um script no `<head>` antes de desenhar. O `visual`
  grava as cores neutras em `--v-*`, que só o tema claro usa; texto na cor da marca
  usa `--cor-texto` (clareada no escuro). Cor nova na página: variável, nos dois temas.
- Progresso da pasta e painel do inventário (04/11/2026): pela rede, conectar a pasta
  (procurar o `lista-mestra.json` e mapear as pastas) leva segundos. `progresso(texto)`
  mostra uma barra e quantas pastas já foram lidas no próprio bloco da lista mestra, e
  o botão fica desabilitado até terminar (`finally`). **Inventariar** abre o painel
  `#registro` no lugar da prévia (um painel por vez, o formulário vai para a esquerda):
  cada arquivo lido entra numa linha (`registrar`) e o relatório fica no topo do painel;
  a mensagem do fim da página só resume.
- Nova versão pelo código (04/11/2026): na seção Arquivo, `#busca-documento` procura na
  lista mestra (código ou título) um documento vigente. Conteúdo sem timbre (texto colado,
  .txt, .md, Word comum) vira nova versão dele: `fichaVigente` lê a ficha do .docx que a
  lista registra (pasta conectada) e a página entra no modo atualização, como se o .docx
  tivesse sido solto. Minuta sem número não está na lista: solta-se o próprio .docx.
- Corpo novo e tipo de alteração (04/11/2026): com um .docx timbrado carregado (minuta
  inclusive), o conteúdo sem timbre que vier depois troca só o corpo dele (caixa
  `#sobre-base`, marcada; desmarcada, vira documento novo). Na minuta de nova versão (com
  `substitui`), o Tipo de alteração oferece a revisão seguinte, a mudança de conteúdo
  (X+1.0) ou a correção (X.Y+1), contadas da versão oficial substituída. A regra é a
  mesma nos dois lados: `sugerirAtualizacao`/`versaoPublicada` na página,
  `preparar_atualizacao`/`versao_publicada` no gerador (`--mudanca conteudo|correcao|rodada`;
  sem ela, `conteudo` no oficial e `rodada` na minuta). `testes/pagina.sh` compara 21 casos.
- Largura (04/11/2026): sem painel aberto, o formulário tem `max-width:50vw` (560px abaixo
  de 900px de tela); com a prévia ou o inventário, vale a regra da metade esquerda.
- Menu e marca d'água na tela (04/11/2026): o botão `#menu` (canto superior esquerdo) liga e
  desliga brilho, bordas na cor do estado e marca d'água, guardados em `localStorage`
  (`timbrador.efeitos`; classes `sem-brilho`, `sem-bordas`, `sem-marca-dagua` no body). A
  marca d'água do estado (`#marca-tela`, `mostrarMarcaTela`) fica nas faixas livres dos dois
  lados do formulário, a 45°, com 470 de 595 pt como no documento (larguras da Helvetica
  da prévia); some no oficial, sem marca configurada, com prévia ou inventário abertos,
  abaixo de 900px e na impressão. Só tela: o .docx não muda. O menu desce para baixo de
  um elemento `[data-barra-topo]`, se algum servidor injetar uma barra; o nginx do projeto
  não injeta.
- Leitura, lista mestra, inventário e pastas continuam os da página original,
  com siglas genéricas (`[A-Z0-9]{2,6}`) e ficha `…timbrador:ficha`.
- `visual` (opcional): `carregarVisual`/`lerVisual` nos dois lados; `sec()`,
  `cor_texto()`/`corTexto()` e `realce()` devolvem o valor do original quando ele
  não existe (sem `visual` a saída é byte a byte a de antes). Fontes embutidas:
  `word/fontTable.xml` + `word/fonts/fontN.odttf` ofuscadas (ECMA-376 17.8.1), GUID
  tirado do CRC32 da própria fonte, para página e gerador darem o mesmo pacote.
  A interface também se veste com ele (cores em variáveis CSS, fontes via `FontFace`).
  O editor da marca ainda não edita o `visual` (só preserva).
- A marca é configurada na própria página (Marca → Configurar/Editar: organização,
  política, empresas com logo por upload, áreas e tipos ativos e extras) e fica no
  `localStorage` (`timbrador.marca`) **no formato do marca.json**, validada por
  `carregarMarca`. Exportar baixa esse `marca.json` (logos em data URI), que
  serve direto ao gerador; Importar/colar aceita um. Campos que o editor não
  mostra (`areas`, `_instrucoes`) passam intactos.

`Inventariar.ps1` e os `.cmd` também foram portados (20/10/2026): empresa
padrão, nome da organização, código da política e tipos extras vêm do
`marca.json`; ficha `…timbrador:ficha`; siglas `[A-Z0-9]{2,6}`. Lista que já
existe mantém o `lista_mestra`/`padrao` dela (como o gerador). Testado com
PowerShell 7.4 no Linux contra `gerar_documento.py --inventariar`: lista mestra
igual (ficha nova e antiga, cabeçalho, nome). Em 03/11/2026 passou a seguir as
regras atuais (ver "Recursos restaurados"), conferido com PowerShell 7.4.6 no Linux
por `testes/inventario.sh`. Os `.cmd` estão
em CRLF (`.gitattributes`).

**Pendente:** no Chrome real só foram vistos a carga da marca e o editor (sem
erro de console); falta gerar um documento e conectar a pasta por lá. O `.ps1`
não rodou no Windows PowerShell 5.1.

## Recursos restaurados e relatório de condução (03/11/2026)


O PR #3 trouxe o porte de uma versão mais antiga do original. Voltaram, nos dois
lados, os recursos da versão mais recente: campo `projeto` (ficha, lista mestra,
pasta `Projetos/Nome`, PLN exige projeto), anexos externos com SHA-256 (`--anexo`;
na página, a caixa "Com anexos externos"), assinaturas previstas, inventário que lê
`Obsoletos` e pula `Modelos` (com alertas e entradas sem arquivo), figuras relativas
à especificação, rascunho da página no navegador e data da aprovação acompanhando o
dia. Na página, o campo Projeto substituiu o "Salvar em".

Relatório de condução (REL com rodadas), vindo do PR #2: número herdado do plano
oficial do projeto, cada rodada é uma versão OFICIAL nova, sem cópia obsoleta, e a
lista mestra guarda o SHA-256 de cada rodada (`rodadas_sha256`). As mesmas regras
estão no gerador (`regras_rel`, `celulas_rodada`) e na página (`regrasRel`,
`celulasRodada`): a impressão digital é calculada sobre o texto das cinco células
da linha. Corpo sem o Quadro de rodadas ganha a seção "Rodadas" no fim. Na página,
o modo de condução é opcional para REL novo (caixa "Relatório de condução de
projeto"); REL sem rodadas é relatório comum.

`Inventariar.ps1` acompanha o `--inventariar` (03/11/2026): lê `Obsoletos` e pula
`Modelos`, registra `projeto` (da ficha ou da pasta), mescla a entrada antiga, tira da
lista a entrada cujo arquivo sumiu de pasta varrida, dá os mesmos alertas e calcula o
SHA-256 das rodadas sobre o mesmo texto das células (`Celulas-Rodada`, espelho de
`celulas_rodada`). `testes/inventario.sh` compara as duas listas e os alertas, inclusive
com rodada adulterada; precisa de `pwsh` (ou `TIMBRADOR_PWSH`) e pula sem ele.
