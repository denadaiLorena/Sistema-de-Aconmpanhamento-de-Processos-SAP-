# SAP — Desafio Técnico (Google Apps Script + Google Sheets)

Web App em **Google Apps Script (V8)** com frontend em HTML/JS (Tailwind via CDN) e backend em Apps Script consumindo dados de uma planilha Google Sheets.

## Estrutura

- `Cliente.html` — UI (tabela, filtros, modais) e chamadas `google.script.run`
- `Servidor.js` — regras de negócio + CRUD + leitura/escrita na planilha
- `appsscript.json` — manifesto do Apps Script

> Observação: este repositório ignora arquivos locais do `clasp` (`.clasp.json`, `.clasprc.json`) por segurança.

## Requisitos

- Node.js (LTS recomendado)
- `clasp` (Google Apps Script CLI)

Instalação do `clasp`:

```bash
npm i -g @google/clasp
```

## Como rodar / publicar (via clasp)

1) Login no Google

```bash
clasp login
```

2) Vincular o projeto local a um Apps Script

Você tem duas opções:

- **Criar um novo Apps Script** e subir os arquivos:
  
  ```bash
  clasp create --type webapp --title "SAP - Desafio"
  clasp push
  ```

- **Vincular a um Apps Script existente** (se você já tem o `scriptId`):

  ```bash
  clasp clone <SCRIPT_ID>
  ```

3) Publicar (Deploy)

```bash
clasp deploy
```

Para atualizar o deploy depois de mudanças:

```bash
clasp push
clasp deploy
```

Dicas úteis:

```bash
clasp open
clasp logs
```

## Planilha / Dados

A aplicação espera uma planilha com abas (sheets) no formato do desafio:

- `Processos`
- `Clientes`
- `Produtos`
- `Unidades`

Se você precisar trocar o ID da planilha usada pelo backend, ajuste a configuração correspondente no `Servidor.js` (onde o código abre a planilha via `SpreadsheetApp`).

## Segurança

- **Não commite** `.clasprc.json` (pode conter token OAuth) e `.clasp.json` (contém `scriptId`). Eles já estão no `.gitignore`.
- Se esses arquivos já existirem na sua pasta local, tudo bem — apenas evite versionar.

## Licença

Defina a licença do repositório no GitHub (ex.: MIT) se você quiser torná-lo público/reutilizável.
