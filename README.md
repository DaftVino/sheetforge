<div align="center">
  <img src="docs/assets/sheetforge-logo.png" alt="SheetForge logo" width="200">
  <h1>SheetForge</h1>
</div>

SheetForge is an Excel toolkit for repeatable importing, cleanup, validation, formatting, summarizing, backup, export, and workbook tasks. It ships 51 actions across ten packages, all chosen and configured on a Setup tab in your workbook instead of in code. SheetForge runs only in desktop Excel on Windows because it is written in VBA.

## Install

There is no installer or add-in, and nothing to run. Download the single `.bas` file from the repository's [Releases page](../../releases), import it into your own Excel workbook, and follow the short [setup guide](docs/setup-guide.md).

## Usage

On the Setup tab, switch on the actions you need and fill in their options. SheetForge builds the workbook buttons from those choices, so day-to-day use does not require editing VBA.

## Docs

- **[Setup guide](docs/setup-guide.md)** — start here; written for someone who has never opened the VBA editor ([PDF](docs/setup-guide.pdf))
- [Architecture](docs/architecture.md) and the [decision records](docs/adr/) — how it works and why
- [Macro catalog](docs/macro-catalog-plug-and-play.md) — every action, and what was deliberately not built
- [Contributing](CONTRIBUTING.md) — including how to ask for a macro without writing any code
- [Security](SECURITY.md) — what the macros can reach, and how to report a problem privately
- [Workflow quick reference](docs/quick-ref-workflow.md)

## License

MIT — see `LICENSE`.
