# Titan007 Excel Exporter

A small Python utility for crawling Titan007 football match data and exporting
Asian handicap / over-under odds history to Excel.

The project includes both a command-line crawler and a Windows-friendly Tkinter
GUI. It is designed around repeatable Excel sync: running the crawler again can
update the same workbook while preserving existing rows through a hidden
metadata sheet.

## Features

- Crawl current Titan007 football fixtures for a selected odds company.
- Export Asian handicap and over-under odds history to `.xlsx`.
- Sync repeated runs into `titan007_data.xlsx` by default.
- Crawl a specific historical date with `--history-date`.
- Limit crawls for smoke testing with `--limit`.
- Use a Tkinter desktop GUI via `python gui.py`.
- Package the GUI with PyInstaller using the included `.spec` files.

## Requirements

- Python 3.10+
- Network access to Titan007 pages
- Python packages:

```bash
pip install requests beautifulsoup4 openpyxl lxml
```

For packaging:

```bash
pip install pyinstaller
```

## Repository Layout

```text
main.py                  Core crawler, parsers, workbook sync, CLI entry point
gui.py                   Tkinter desktop application
startup_guard.py         Optional startup time-limit guard
tests/                   Unit tests for crawler, workbook, GUI support logic
gui.spec                 PyInstaller GUI build config
Titan007Exporter.spec    Alternate PyInstaller build config
```

Generated workbooks, build outputs, local caches, virtual environments, IDE
files, and reference data are intentionally ignored by Git.

## Excel Template

The application expects an Excel template workbook by default. In the local
project this is the league match data template at the project root.

Because Excel files are ignored by Git, place the template in the project root
manually before running the default export, or pass another template path:

```bash
python main.py --template path/to/template.xlsx
```

You can also run without a template in code by calling the workbook helpers with
`template_path=None`.

## CLI Usage

Run a full current sync:

```bash
python main.py
```

Run a small smoke test:

```bash
python main.py --limit 2
```

Export a historical date:

```bash
python main.py --history-date 2026-05-26
```

Use a custom output file:

```bash
python main.py --output exports/titan007_data.xlsx
```

Adjust worker concurrency:

```bash
python main.py --workers 4
```

The default output for current sync is:

```text
titan007_data.xlsx
```

Historical exports use date-based names such as:

```text
titan007_data_20260526.xlsx
```

## GUI Usage

Start the desktop application:

```bash
python gui.py
```

The GUI lets you configure the output directory, template, crawl scope, monitor
settings, and selected matches without using command-line arguments.

## Tests

Run the unit test suite:

```bash
python -m unittest discover tests
```

The tests cover Titan007 parsers, odds-history handling, workbook sync behavior,
GUI support helpers, and the optional startup guard.

## Packaging

Build the GUI executable with PyInstaller:

```bash
pyinstaller gui.spec
```

or:

```bash
pyinstaller Titan007Exporter.spec
```

Build output is written to `build/` and `dist/`, both of which are ignored by
Git.

## Notes

- This project is not affiliated with Titan007.
- Network scraping may break if the upstream page structure changes.
- Do not commit generated Excel exports or bundled executable files.
