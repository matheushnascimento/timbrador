<!-- bmad:context -->
<!-- Verificado em 2026-11-10 contra 0207300. Mantido por bmad-project-context; edições dentro deste bloco são substituídas no refresh. O que quiser preservar, ponha fora dos marcadores. -->

## Timbrador (white label)

Plugin do Claude que gera documentos `.docx` timbrados (papel timbrado + ABNT + controle documental) a partir de uma especificação JSON e de um `marca.json` por organização. O gerador é Python só com biblioteca padrão, escrevendo WordprocessingML direto num zip; a página de navegador (`pagina/timbrador.html`) é porte linha a linha dele e serve no homeserver. A norma do documento está em `skills/timbrador/SKILL.md`; o planejamento BMad fica em `_bmad-output/`.

## Política

- Responda, planeje e escreva documentos, tickets e commits em português; commits em conventional commits.
- Nunca faça commit na `main`: trabalhe em branch e entre por PR.
- O checkout `/home/mahedina/projects/timbrador` é o site no ar (`pagina/` montada no container): não troque de branch nem deixe `pagina/` pela metade ali; trabalho longo vai num worktree.
- Nunca use `python-docx`, `docx-js` ou qualquer dependência no gerador: só biblioteca padrão do Python 3.8+, para o plugin rodar em qualquer lugar.
- O `marca.json` é a única coisa que varia entre organizações (sigla, nome, razão social, documento CNPJ ou CPF, cidade, cor, logo, `grupo`, `ativa`, e os opcionais `areas`, `areas_extra`, `areas_ativas`, `tipos_extra`, `tipos_ativos`, `politica`, `visual`). Nenhum dado de organização real no código nem nos testes; teste com as marcas fictícias de `testes/`.
- Nunca versione `.docx`/`.pdf` gerados, `marca.json` de organização real (`pagina/marca.json`, `skills/timbrador/assets/marca.json`), `original/` (tem razão social, CNPJ e logos reais) nem os ícones do site; só `marca.exemplo.json` e as marcas de `testes/`.
- Dado que falta vira `[definir]`: nunca invente razão social, CNPJ, número, cláusula, data ou valor.
- Travas de estado são segurança documental, não validação. OFICIAL exige aprovador, data de aprovação, número reservado e empresa com cadastro completo; OBSOLETO exige substituto ou motivo; empresa com `razao_social` ou `documento` em `[definir]` só emite MINUTA ou EM_APROVACAO (RASCUNHO e EM_REVISAO estão aposentados: só leitura). Nunca contorne: rebaixe o estado ou peça o dado.
- O visual é o do original e não se negocia com o conteúdo: faixa de estado de borda a borda; bloco de controle centralizado (código · versão · Página X de Y; frase do estado; razão social · documento · cidade); rodapé com filete na cor da empresa e logo; borda preta a 1 cm do papel; marca d'água nos estados não oficiais; margens esq. 3 · dir. 2 · sup. 3,3 · inf. 3,2 (curto e capa) ou 2,6 (páginas internas do longo). Única exceção (decisão de 03/11/2026): o `visual` opcional do `marca.json` troca fontes e cores de texto, títulos e realce; sem ele nada muda, e ele nunca muda faixa, cores de estado, borda, margens nem linhas de tabela.
- Assinatura entra só pela página (caixa "Com assinatura"): o gerador tira as assinaturas do corpo e guarda os envolvidos como `assinaturas_previstas` na ficha (`adiar_assinaturas`; na página, `adiarAssinaturas`).

## Onde ficam as coisas

- Antes de editar `pagina/timbrador.html`, `pagina/Inventariar.ps1` ou as regras do relatório de condução, leia `docs/pagina.md`.
- Antes de mexer em `deploy/`, `docker-compose.yml`, `pagina/marca.json` ou nos ícones do site, leia `docs/homeserver.md`.
- Antes de dizer que algo está testado, ou de mudar `testes/`, leia `docs/validacao.md`.
- Mapa do repositório, histórico do porte e comandos do gerador: `docs/estrutura.md`; flags em `gerar_documento.py --help`.
- Original × gerador anterior: `docs/divergencias-original.md`. Original intacto: commit `70a7137`.

## Rodando e verificando

- Depois de mexer no gerador ou na página, rode `bash testes/fumaca.sh` e `NODE_PATH=<pasta com jsdom>/node_modules bash testes/pagina.sh`; mexeu no `Inventariar.ps1` ou no `--inventariar`, rode também `bash testes/inventario.sh` (precisa de `pwsh` ou `TIMBRADOR_PWSH`).
- `testes/pagina.sh` e `testes/inventario.sh` imprimem "pulado" e saem com 0 sem jsdom ou pwsh: leia a saída, porque "pulado" não é aprovado.
- LibreOffice e `pdftotext` são opcionais (numeram o sumário e geram o PDF) e estão instalados nesta máquina: o aviso "sumário sem números de página" aqui quer dizer que um dos dois sumiu do PATH. `TIMBRADOR_SOFFICE` troca o comando do LibreOffice (sandbox que precisa de wrapper).

## Convenções diferentes do padrão

- Gerador e página andam juntos: mudança de visual, de regra ou de formato da marca vai nos dois, e o pacote tem de sair igual parte a parte (`testes/pagina.sh` confere). Regra que o `Inventariar.ps1` também aplica vai nele também (`regras_rel`/`regrasRel`; `celulas_rodada`/`celulasRodada`/`Celulas-Rodada`).
- O núcleo da página (`<script id="nucleo">`) é porte linha a linha do gerador: mesma ordem de ids, rels e mídia (`img1.png`), `round()` com meio para o par (`arred`), ficha em `json.dumps` (`jsonPy`). Mexeu num, mexa no outro.
- `*.cmd` ficam em CRLF (`.gitattributes`): lote do Windows falha com LF.

## Armadilhas conhecidas

- Agentes já mudaram só o gerador ou só a página, deixaram dado de empresa real no código e commitaram na `main` neste repositório: as regras acima existem por isso.
- `deploy/nginx.conf` é montado como arquivo único: edição que troca o arquivo (`sed -i`, a maioria dos editores) deixa o container com a cópia velha, e `nginx -s reload` não pega. Use `docker compose restart timbrador` e confira o cabeçalho servido com `curl -sI http://127.0.0.1:8091/`.
- SVG de logo rasterizado com o ImageMagick sai com o M desenhado errado; rasterize com resvg.

<!-- /bmad:context -->
