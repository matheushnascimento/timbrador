# Homeserver e a marca do site

> Movido do `CLAUDE.md` em 10/11/2026 (adoção do BMad, `bmad-project-context`). Texto mantido como estava; leia antes de mexer em `deploy/`, `docker-compose.yml`, `pagina/marca.json` ou nos ícones do site.


Esta máquina é o homeserver. `docker compose up -d` (na raiz) sobe o container
`timbrador` (nginx:alpine, `restart: unless-stopped`) servindo `pagina/` em
**127.0.0.1:8091**, com CSP que só deixa a página ler arquivos do próprio site
(`deploy/nginx.conf`). O volume é a própria pasta `pagina/` do checkout: o que
estiver no branch atual é o que está no ar. O acesso de fora é por um hostname
público do túnel Cloudflare (configurado no painel, token no serviço
`cloudflared`) apontando para `http://localhost:8091`; o HTTPS do túnel é o que
deixa a página gravar direto na pasta dos documentos.

**Marca padrão do site: mhdn** (atualizada em 03/11/2026). `pagina/marca.json`
(fora do git, `.gitignore`) é a marca da mhdn, montada a partir do brandbook
(claude.ai/artifact/RSANaxBgb5bQ7Rquwgbezz): logo = livro novo (`MHDN -logo.svg`)
+ assinatura `mhdn`, compostos nas proporções da horizontal antiga e rasterizados
com resvg (o renderizador do ImageMagick desenha o M errado); cor `anil`
#233F61; CPF, Rio Claro, política `MHDN-POL-ADM-001`, todas as áreas e tipos; e
`visual` com os tokens do brandbook (tinta, anil, grafite, anil-claro; papel,
papel-cartao e linha na página) e EB Garamond (títulos) + Newsreader (texto)
embutidas. As fontes vieram dos woff2 do brandbook → TTF (wawoff2) com a tabela
de nomes refeita para uma família por fonte (Regular, Italic e Bold = Garamond
500 / Newsreader 600, peso marcado 700). Insumos e a versão anterior da marca
estão em `saida/mhdn/` (fora do git). A página busca esse arquivo ao abrir e o usa enquanto o navegador não
guardou uma marca própria (Editar/Importar guardam; "Usar a do site" volta). A
mesma cópia está em `skills/timbrador/assets/marca.json` (também fora do git),
que o gerador acha sozinho. O ícone do site (`pagina/favicon.svg`, `favicon.png`, `apple-touch-icon.png`,
fora do git) é o livro de capa azul (`Mhdn logo simplificada capa azul.SVG` do
brandbook) sem o fundo, recortado em quadrado; a página só aponta para esses nomes. **O `marca.json` do site é público** (tem o CPF):
sem Cloudflare Access no hostname, qualquer um o baixa.
