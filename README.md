# Timbrador (white label)

Plugin para o Claude e página de navegador que transformam qualquer conteúdo em um .docx timbrado com a identidade da **sua** organização, seguindo ABNT e um controle documental com código, versão, estado (minuta, em aprovação, oficial, obsoleto), ficha embutida e lista mestra.

## Como personalizar

1. Copie `skills/timbrador/assets/marca.exemplo.json` para `marca.json` e preencha sigla, nome, razão social, CNPJ (ou CPF), cidade, cor e logo (PNG/JPEG) de cada empresa.
2. Confira: `python3 skills/timbrador/scripts/gerar_documento.py --checar-marca marca.json`
3. Escolha como distribuir:
   - **Plugin já personalizado** (para a sua equipe): salve o arquivo como `skills/timbrador/assets/marca.json`, junto do logo.
   - **Plugin genérico** (para qualquer pessoa): não inclua `marca.json`. Na primeira conversa o Claude pergunta os dados e entrega um `marca.json` para reutilizar.
4. Para ter um arquivo único e portátil (logo embutido): `--embutir-logos marca.json --saida marca_portatil.json`.

Empresas sem razão social ou documento cadastrados só emitem minuta ou documento em aprovação.

## Página para os colaboradores

`skills/timbrador/assets/timbrador.html` timbra um .docx, .txt ou .md pronto (ou texto colado) direto no navegador, sem enviar nada para a internet. Ela tem o mesmo visual e as mesmas regras do gerador, põe as assinaturas, controla a lista mestra numa pasta conectada (Edge ou Chrome) e gera a nova versão de um documento já timbrado.

- `--montar-pagina marca.json --saida PASTA/` grava uma cópia da página com a marca embutida: é essa que se distribui (junto com `Abrir-Timbrador.cmd`, no Windows).
- A página sem marca embutida pede o `marca_portatil.json` uma vez e o guarda no navegador.

## Uso direto, sem o Claude

```bash
python3 skills/timbrador/scripts/gerar_documento.py especificacao.json --marca marca.json --saida saida/ --pdf
```

O formato da especificação e os comandos de controle (`--ler`, `--atualizar`, `--obsoletar`, `--lista`, `--inventariar`, `--anexo`) estão em `skills/timbrador/SKILL.md`. Requer só Python 3.8+. LibreOffice e `pdftotext` são opcionais: numeram o sumário e geram o PDF de conferência (`TIMBRADOR_SOFFICE` define outro comando para o LibreOffice).

## Testes

```bash
bash testes/fumaca.sh
```

Usa só marcas fictícias. Com Node e Playwright instalados, também gera o mesmo documento pela página e confere que sai igual ao do gerador.
