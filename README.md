# Timbrador (white label)

Plugin para o Claude que transforma qualquer conteúdo em um .docx timbrado com a identidade da **sua** organização, seguindo ABNT e um controle documental com código, versão e estado (rascunho, minuta, em revisão, oficial, obsoleto).

## Como personalizar

1. Copie `skills/timbrador/assets/marca.exemplo.json` para `marca.json` e preencha sigla, nome, razão social, CNPJ (ou CPF), cidade, cor e logo (PNG/JPEG).
2. Confira: `python3 skills/timbrador/scripts/gerar_documento.py --checar-marca marca.json`
3. Escolha como distribuir:
   - **Plugin já personalizado** (para a sua equipe): salve o arquivo como `skills/timbrador/assets/marca.json`, junto do logo.
   - **Plugin genérico** (para qualquer pessoa): não inclua `marca.json`. Na primeira conversa o Claude pergunta os dados e entrega um `marca.json` para reutilizar.
4. Para ter um arquivo único e portátil (logo embutido): `--embutir-logos marca.json --saida marca_portatil.json`.

Empresas sem razão social ou documento cadastrados só emitem documentos não oficiais.

## Uso direto, sem o Claude

```bash
python3 skills/timbrador/scripts/gerar_documento.py especificacao.json --marca marca.json --saida saida/ --pdf
```

O formato da especificação está em `skills/timbrador/SKILL.md` (seção 6). Requer só Python 3.8+. LibreOffice e `pdftotext` são opcionais: numeram o sumário e geram o PDF de conferência (`TIMBRADOR_SOFFICE` define outro comando para o LibreOffice).
