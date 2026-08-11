# officekit

<p align="center">
  <img src="assets/cover.png" alt="officekit — Office document automation for humans and AI agents" width="768">
</p>

[![CI](https://github.com/dream-zjk/officekit/actions/workflows/ci.yml/badge.svg)](https://github.com/dream-zjk/officekit/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org)

**A friendly toolkit to read, create, edit, and fill Office documents — built for
humans and AI agents alike.**

`officekit` is a small, well-typed Python library and CLI that works with the
three formats agents meet most often: **Word (`.docx`)**, **Excel (`.xlsx`)**,
and **PowerPoint (`.pptx`)**. It follows the same *agent-first* philosophy as
the new generation of Office automation CLIs — one predictable command surface,
path-based addressing, atomic batch edits, structured `--json` output, and an
**MCP server** so an agent can drive documents without shell access.

This is an **original Python implementation**, not a fork or a copy of any
codebase. It is inspired by agent-facing Office CLIs such as
`iOfficeAI/OfficeCLI`; see [How it compares](#how-it-compares).

## Why officekit?

- 📄 **Three formats, one API** — `text` / `view` / `query` / `set` / `add` /
  `remove` / `merge` / `info` / `validate` work across docx, xlsx, and pptx.
- 🤖 **Agent-friendly** — every command supports `--json` with a structured
  result (`{"success": true, "data": ...}` / `{"success": false, "error": {...}}`)
  and clear exit codes, so an agent can reason about results and self-heal.
- 🧭 **Path addressing** — target elements like `/slides[1]/shapes[2]` or
  `/sheets[1]/A1` (1-based, like the reference CLIs).
- ⛓️ **Atomic `batch`** — run several edits in one pass; if any step fails the
  whole operation rolls back, leaving the source untouched.
- 🪶 **Lightweight** — pure-Python on top of `python-docx` / `openpyxl` /
  `python-pptx`; no Office install, no native binaries.
- 🔌 **MCP server** — `officekit mcp` exposes every capability as an MCP tool
  over stdio for Claude Code, Cursor, VS Code, and friends.

## Installation

```bash
pip install officekit
```

With the optional MCP server:

```bash
pip install "officekit[mcp]"
```

From a checkout:

```bash
pip install -e ".[test]"
```

## CLI usage

```bash
# Read
officekit text  report.docx                 # all visible text
officekit view  report.docx --mode outline # headings / sheets / slide titles
officekit view  report.docx --mode stats   # counts & dimensions
officekit view  report.docx --mode issues  # empty / unfilled placeholders

# Address elements
officekit query deck.pptx  "/slides[1]/shapes[2]"
officekit set    deck.pptx  "/slides[1]/shapes[1]" --value "New title"
officekit add    deck.pptx  slide  --value "Second"
officekit remove deck.pptx  "/slides[2]"

# Create & fill
officekit create docx invoice.docx --title "Invoice"
officekit merge  invoice-template.docx out.docx \
  --data '{"client":"Acme","amount":"$1,200"}'

# Import tabular data
officekit import book.xlsx People data.csv --header

# Atomic multi-edit (rolls back on any failure)
officekit batch book.xlsx --operations '[
  {"op":"set","target":"/sheets[1]/A1","value":"NAME"},
  {"op":"add","what":"row","value":"Carol\t40"}
]'

# Inspect
officekit info   deck.pptx
officekit validate report.docx             # exit 1 on problems

# All commands accept --json for machine-readable output
officekit info deck.pptx --json
```

Exit codes: `0` success, `1` validation issue / error, `2` bad input.

### Structured (JSON) output

```bash
$ officekit info deck.pptx --json
{"success": true, "data": {"kind": "pptx", "slides": 1}}

$ officekit info missing.docx --json
{"success": false, "error": {"code": "ValueError", "suggestion": "Unsupported file type: '.pdf' ..."}}
```

## MCP server

```bash
pip install "officekit[mcp]"
officekit mcp          # speaks MCP over stdio; point your agent at it
```

The server exposes 14 tools: `extract_text`, `view_outline`, `view_stats`,
`view_issues`, `document_info`, `validate_document`, `create_document`,
`merge_template`, `get`, `set_value`, `add`, `remove`, `batch`, `import_csv`.

## Library usage

```python
from officekit import extract_text, create_document, merge_template, ops, view, batch

create_document("pptx", "deck.pptx", title="Hi {{who}}")
ops.set_value("deck.pptx", "/slides[1]/shapes[1]", "Quarterly Review")
print(view.view_outline("deck.pptx"))
batch.run_batch("deck.pptx", [{"op": "add", "what": "slide", "value": "Appendix"}])
```

## How it compares

This project is **inspired by** agent-facing Office CLIs such as
`iOfficeAI/OfficeCLI`. The shared idea: documents should be *programmable* by
agents through a simple, scriptable surface (one uniform API, path-based
addressing, atomic batch edits, JSON I/O, and an MCP server).

`officekit` is an **independent, MIT-licensed Python rewrite** focused on the
most common operations, not a fork or a copy of that codebase. Where the
reference uses a from-scratch C#/.NET engine with native rendering, `officekit`
builds on the mature `python-docx` / `openpyxl` / `python-pptx` stack so it
runs anywhere Python runs with zero native dependencies.

## Development

```bash
git clone https://github.com/dream-zjk/officekit.git
cd officekit
python -m venv .venv && source .venv/bin/activate
pip install -e ".[test,mcp]"
pytest -q
```

## Contributing

Found a format quirk or want another command? Open an issue or PR — see
[CONTRIBUTING.md](CONTRIBUTING.md).

## License

[MIT](LICENSE) © 2026 ZHANG8
