# feishu-docget

<p align="center">
  <img src="./src/static/favicon.png" height="112" alt="feishu-docget logo" />
</p>

<p align="center">
  Export Feishu cloud documents to Word with high fidelity
</p>

<p align="center">
  <a href="./README.md">中文</a> ·
  English
</p>

<p align="center">
  <a href="https://github.com/HawkonS/feishu-docget/actions/workflows/checks.yml"><img src="https://github.com/HawkonS/feishu-docget/actions/workflows/checks.yml/badge.svg" alt="Checks" /></a>
  <a href="https://github.com/HawkonS/feishu-docget/actions/workflows/codeql-analysis.yml"><img src="https://github.com/HawkonS/feishu-docget/actions/workflows/codeql-analysis.yml/badge.svg" alt="CodeQL" /></a>
  <a href="./LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-blue.svg" alt="License: Apache-2.0" /></a>
</p>

## Overview

An earlier approach converted documents along the `feishu -> markdown -> docx` path, but complex Feishu documents lose a lot of structural information on the way to Markdown — merged cells, rich text styling, images and whiteboards, nested lists, and more. This project instead reads the Block structure returned by the Feishu Open Platform directly, builds a Word object tree with `python-docx`, and then applies templates and format cleanup.

It ships as one web service (a download frontend plus an admin dashboard) and one command line tool.

## Features

- Direct to Word: `.docx` is generated straight from Feishu Blocks, avoiding the losses of a Markdown round trip.
- Templates: upload and select `.docx` templates to reuse headers, footers, styles and cover pages.
- Formatting control: tables, images and whiteboards, code blocks, body paragraphs, page margins and document properties are all adjustable per task, and the rules are never written back into the template file.
- Task queue: submissions run in a queue with progress, logs and downloadable results in the frontend.
- Admin dashboard: a single place for projects, templates, configuration, download statistics, logs and system operations.
- CLI export: `tools/feishu2word.sh` (`feishu2word.bat` on Windows) covers every advanced option of the frontend, which makes scripted exports easy.
- Custom download bot: Feishu bot credentials can be supplied per task; they are validated first and used when possible, falling back to the system default bot when permissions are insufficient.

## Quick start

1. Create a custom app on the Feishu Open Platform to obtain an `App ID` and `App Secret`, and grant the scopes you need: docs, sheets, wiki, media download, whiteboard export. For internal documents the bot must also be added as a collaborator on the document page.
2. Run `./run.sh` (`run.bat` on Windows). The first run installs dependencies from `requirements.txt` and generates `feishu-docget.properties`, where you fill in the Feishu credentials.
3. Open `http://127.0.0.1:7800/`, pick a template, paste a document link and create a task. The admin dashboard defaults to `http://127.0.0.1:7800/admin`.

Results are written to `output/<doc_id>/<document title>.docx`, images land in a sibling `img/` directory, and re-downloading the same document reuses previously fetched images where possible.

## Command line

```bash
sh tools/feishu2word.sh "https://example.feishu.cn/wiki/xxxx" --template Hawkon.docx --style 3
```

Flags map one to one onto the advanced options in the frontend. Run `sh tools/feishu2word.sh --help` for the full list; `--list-templates` and `--list-styles` show what is available, and `--print-options` prints the assembled options without downloading.

## Configuration

All configuration lives in `feishu-docget.properties` at the project root. It is generated on first start, missing keys are filled in automatically, and the log, output and template directories are created for you. For day-to-day changes prefer the configuration page in the admin dashboard, which documents every key together with its default. The file holds sensitive credentials and must not be committed to Git.

Deployments across hosts or containers, user login, and reverse proxy setups need a few extra keys adjusted — the configuration page is the reference for those as well.

## FAQ

- **Permission denied**: check the app scope, whether the bot is a collaborator on the document, and whether the link belongs to the tenant the app can access.
- **Missing images or whiteboards**: make sure image download is enabled, the app has the media download scope, and look for 403s or timeouts in the logs.
- **Output does not match the template**: advanced cleanup rules take precedence over the template, so leave the relevant options empty or disabled if you want the template to win.

## Development

- Enable local commit checks with `git config core.hooksPath .githooks`; every commit then runs the same unit tests as GitHub Actions.
- Unit tests: `python -m unittest discover -s tests -p 'test_*.py'`. Syntax check: `python -m compileall ./src`.
- Where things live: web/API and admin in `src/app.py`, conversion and cleanup in `src/converters/docx/`, Feishu API calls in `src/core/feishu_client.py`, pages in `src/web/templates/`.

## Security notes

- Never commit `feishu-docget.properties`, logs, exported files or private templates.
- The admin dashboard exposes download, delete, configuration and system operations. Before deploying it publicly, set a strong password and keep it behind a trusted network or an authenticating reverse proxy.
- This project is intended for learning, archiving and internal automation. Respect the Feishu platform rules and your organization's data compliance requirements.

## License

Released under the [Apache License 2.0](./LICENSE).
