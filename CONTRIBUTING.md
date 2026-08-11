# Contributing to officekit

Thanks for helping make Office documents programmable for everyone.

## Getting started

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[test]"
pytest -q
```

## Guidelines

- Keep the public surface small and predictable: one verb per CLI subcommand,
  one function per library capability.
- Add a test for any new behavior. Tests live in `tests/` and run against
  real `.docx` / `.xlsx` / `.pptx` files created in a temp directory.
- Format with black and type-check with mypy when possible.
- Open an issue before large changes so we can align on design.

## Reporting bugs

Include the document format, a minimal reproduction, and the expected vs.
actual output.
