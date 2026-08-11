# officekit

[![CI](https://github.com/dream-zjk/officekit/actions/workflows/ci.yml/badge.svg)](https://github.com/dream-zjk/officekit/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org)

**A friendly toolkit to read, create, and fill Office documents — built for
humans and AI agents alike.**

`officekit` is a small, well-typed Python library and CLI that works with the
three formats agents meet most often: **Word (`.docx`)**, **Excel (`.xlsx`)**,
and **PowerPoint (`.pptx`)**. It follows the same *agent-first* philosophy as
the new generation of Office automation CLIs — one predictable command surface
you can call from a script or an LLM tool — but it is an original Python
implementation with zero native dependencies.

## Why officekit?

- 📄 **Three formats, one API** — `extract` / `create` / `merge` / `info` /
  `validate` work identically across docx, xlsx, and pptx.
- 🤖 **Agent-friendly** — clear exit codes, JSON `info` output, and a CLI that
  does exactly one thing per invocation, so an agent can reason about results.
- 🧩 **Template filling** — drop `{{ key }}` placeholders into a document and
  fill them from a JSON object (great for invoices, reports, slides).
- 🪶 **Lightweight** — pure-Python on top of the standard `python-docx` /
  `openpyxl` / `python-pptx` stack; no Office install, no native binaries.

## Installation

```bash
pip install officekit
```

Or from a checkout:

```bash
pip install -e ".[test]"
```

## CLI usage

```bash
# Extract all visible text
officekit text report.docx

# Create a starter document
officekit create docx invoice.docx --title "Invoice"
officekit create xlsx data.xlsx --title "Q4"
officekit create pptx deck.pptx --title "Quarterly Review"

# Fill {{ key }} placeholders from JSON
officekit merge invoice-template.docx out.docx \
  --data '{"client":"Acme","amount":"$1,200"}'

# Inspect document shape
officekit info deck.pptx

# Validate a document opens cleanly (exit 1 on problems)
officekit validate report.docx
```

Exit codes: `0` success, `1` validation issue / CLI error, `2` bad input.

## Library usage

```python
from officekit import extract_text, create_document, merge_template, document_info

# Create
create_document("pptx", "deck.pptx", title="Hi {{who}}")

# Fill a template
merge_template("deck.pptx", "out.pptx", {"who": "Team"})

# Read it back
print(extract_text("out.pptx"))

# Inspect shape
print(document_info("out.pptx"))   # {'kind': 'pptx', 'slides': 1}
```

## How it compares

This project is **inspired by** agent-facing Office CLIs such as
`iOfficeAI/OfficeCLI` — the idea that documents should be programmable by
agents through a simple, scriptable surface. `officekit` is an independent,
MIT-licensed Python rewrite focused on the most common operations, not a fork
or a copy of that codebase.

## Development

```bash
git clone https://github.com/dream-zjk/officekit.git
cd officekit
python -m venv .venv && source .venv/bin/activate
pip install -e ".[test]"
pytest -q
```

## Contributing

Found a format quirk or want another command? Open an issue or PR — see
[CONTRIBUTING.md](CONTRIBUTING.md).

## License

[MIT](LICENSE) © 2026 ZHANG8
