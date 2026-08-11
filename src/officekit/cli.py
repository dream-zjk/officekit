"""Command-line interface for officekit."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import List, Optional

from . import __version__
from .core import (
    create_document,
    document_info,
    extract_text,
    merge_template,
    validate_document,
)


def _cmd_text(args: argparse.Namespace) -> int:
    print(extract_text(args.file))
    return 0


def _cmd_create(args: argparse.Namespace) -> int:
    out = create_document(args.kind, args.path, title=args.title)
    print(f"created {out}")
    return 0


def _cmd_merge(args: argparse.Namespace) -> int:
    if args.data is None:
        sys.stderr.write("--data JSON is required\n")
        return 2
    try:
        data = json.loads(args.data)
    except json.JSONDecodeError as exc:
        sys.stderr.write(f"invalid --data JSON: {exc}\n")
        return 2
    out = merge_template(args.template, args.out, data)
    print(f"merged -> {out}")
    return 0


def _cmd_info(args: argparse.Namespace) -> int:
    print(json.dumps(document_info(args.file), ensure_ascii=False, indent=2))
    return 0


def _cmd_validate(args: argparse.Namespace) -> int:
    issues = validate_document(args.file)
    if issues:
        for issue in issues:
            print(f"[issue] {issue}")
        return 1
    print("OK")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="officekit",
        description="Read, create, and fill Office documents (.docx/.xlsx/.pptx).",
    )
    p.add_argument("--version", action="version", version=f"officekit {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    sp = sub.add_parser("text", help="extract text from a document")
    sp.add_argument("file")
    sp.set_defaults(func=_cmd_text)

    sp = sub.add_parser("create", help="create a starter document")
    sp.add_argument("kind", choices=["docx", "xlsx", "pptx"])
    sp.add_argument("path")
    sp.add_argument("--title", default=None)
    sp.set_defaults(func=_cmd_create)

    sp = sub.add_parser("merge", help="fill {{ key }} placeholders")
    sp.add_argument("template")
    sp.add_argument("out")
    sp.add_argument("--data", default=None, help="JSON object of placeholder values")
    sp.set_defaults(func=_cmd_merge)

    sp = sub.add_parser("info", help="show document shape (counts, sheets, ...)")
    sp.add_argument("file")
    sp.set_defaults(func=_cmd_info)

    sp = sub.add_parser("validate", help="check a document opens cleanly")
    sp.add_argument("file")
    sp.set_defaults(func=_cmd_validate)

    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
