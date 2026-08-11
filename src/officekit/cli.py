"""Command-line interface for officekit.

Every command supports ``--json`` to emit a machine-readable result::

    {"success": true,  "data": <payload>}
    {"success": false, "error": {"code": "...", "suggestion": "..."}}

so an agent can call officekit the same way it would call the reference CLI.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from . import __version__
from . import core, ops, view
from .batch import run_batch
from .importer import import_csv


def _j(args: argparse.Namespace, data: Any) -> None:
    if getattr(args, "json", False):
        print(json.dumps({"success": True, "data": data}, ensure_ascii=False))


def _cmd_text(args: argparse.Namespace) -> int:
    text = core.extract_text(args.file)
    if args.json:
        _j(args, text)
    else:
        print(text)
    return 0


def _cmd_create(args: argparse.Namespace) -> int:
    out = core.create_document(args.kind, args.path, title=args.title)
    if args.json:
        _j(args, str(out))
    else:
        print(f"created {out}")
    return 0


def _cmd_merge(args: argparse.Namespace) -> int:
    if args.data is None:
        raise ValueError("--data JSON is required")
    data = json.loads(args.data)
    out = core.merge_template(args.template, args.out, data)
    if args.json:
        _j(args, str(out))
    else:
        print(f"merged -> {out}")
    return 0


def _cmd_info(args: argparse.Namespace) -> int:
    info = core.document_info(args.file)
    if args.json:
        _j(args, info)
    else:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    return 0


def _cmd_validate(args: argparse.Namespace) -> int:
    issues = core.validate_document(args.file)
    if issues:
        if args.json:
            _j(args, issues)
        else:
            for issue in issues:
                print(f"[issue] {issue}")
        return 1
    if args.json:
        _j(args, [])
    else:
        print("OK")
    return 0


def _cmd_view(args: argparse.Namespace) -> int:
    if args.mode == "text":
        data = view.view_text(args.file)
    elif args.mode == "outline":
        data = view.view_outline(args.file)
    elif args.mode == "stats":
        data = view.view_stats(args.file)
    else:
        data = view.view_issues(args.file)
    if args.json:
        _j(args, data)
    elif args.mode in ("outline", "stats", "issues"):
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        print(data)
    return 0


def _cmd_query(args: argparse.Namespace) -> int:
    data = ops.get(args.file, args.target)
    if args.json:
        _j(args, data)
    else:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


def _cmd_set(args: argparse.Namespace) -> int:
    ops.set_value(args.file, args.target, args.value)
    if args.json:
        _j(args, True)
    else:
        print(f"set {args.target}")
    return 0


def _cmd_add(args: argparse.Namespace) -> int:
    ops.add(args.file, args.what, args.value)
    if args.json:
        _j(args, True)
    else:
        print(f"added {args.what}")
    return 0


def _cmd_remove(args: argparse.Namespace) -> int:
    ops.remove(args.file, args.target)
    if args.json:
        _j(args, True)
    else:
        print(f"removed {args.target}")
    return 0


def _cmd_batch(args: argparse.Namespace) -> int:
    if args.file and args.file != "-":
        operations = json.loads(Path(args.file).read_text(encoding="utf-8"))
    else:
        operations = json.loads(args.operations)
    if not isinstance(operations, list):
        raise ValueError("operations must be a JSON array")
    result = run_batch(args.path, operations)
    if args.json:
        _j(args, result)
    else:
        print(f"applied {result['applied']} operation(s) to {result['file']}")
    return 0


def _cmd_import(args: argparse.Namespace) -> int:
    out = import_csv(
        args.out, args.sheet, args.csv, header=args.header, delimiter=args.delimiter
    )
    if args.json:
        _j(args, str(out))
    else:
        print(f"imported {args.csv} -> {out}")
    return 0


def _cmd_mcp(args: argparse.Namespace) -> int:
    try:
        from . import mcp_server
    except ImportError:
        sys.stderr.write(
            "mcp extra not installed. Run: pip install officekit[mcp]\n"
        )
        return 1
    mcp_server.run()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parent = argparse.ArgumentParser(add_help=False)
    parent.add_argument(
        "--json", action="store_true", help="emit a JSON result {success,data|error}"
    )

    p = argparse.ArgumentParser(
        prog="officekit",
        description="Read, create, edit, and fill Office documents (.docx/.xlsx/.pptx).",
    )
    p.add_argument(
        "--version", action="version", version=f"officekit {__version__}"
    )
    sub = p.add_subparsers(dest="command", required=True)

    def add(name: str, help: str, **kw) -> argparse.ArgumentParser:
        sp = sub.add_parser(name, help=help, parents=[parent])
        return sp

    sp = add("text", "extract text from a document")
    sp.add_argument("file")
    sp.set_defaults(func=_cmd_text)

    sp = add("create", "create a starter document")
    sp.add_argument("kind", choices=["docx", "xlsx", "pptx"])
    sp.add_argument("path")
    sp.add_argument("--title", default=None)
    sp.set_defaults(func=_cmd_create)

    sp = add("merge", "fill {{ key }} placeholders")
    sp.add_argument("template")
    sp.add_argument("out")
    sp.add_argument("--data", default=None, help="JSON object of placeholder values")
    sp.set_defaults(func=_cmd_merge)

    sp = add("info", "show document shape")
    sp.add_argument("file")
    sp.set_defaults(func=_cmd_info)

    sp = add("validate", "check a document opens cleanly")
    sp.add_argument("file")
    sp.set_defaults(func=_cmd_validate)

    sp = add("view", "inspect a document (text/outline/stats/issues)")
    sp.add_argument("file")
    sp.add_argument(
        "--mode",
        choices=["text", "outline", "stats", "issues"],
        default="text",
    )
    sp.set_defaults(func=_cmd_view)

    sp = add("query", "read an element by path (get)")
    sp.add_argument("file")
    sp.add_argument("target", help="path e.g. /slides[1]/shapes[2]")
    sp.set_defaults(func=_cmd_query)

    sp = add("set", "write text/value to an element by path")
    sp.add_argument("file")
    sp.add_argument("target")
    sp.add_argument("--value", required=True)
    sp.set_defaults(func=_cmd_set)

    sp = add("add", "append an element (paragraph/slide/row/sheet)")
    sp.add_argument("file")
    sp.add_argument("what", choices=["paragraph", "slide", "row", "sheet"])
    sp.add_argument("--value", default=None)
    sp.set_defaults(func=_cmd_add)

    sp = add("remove", "delete an element by path")
    sp.add_argument("file")
    sp.add_argument("target")
    sp.set_defaults(func=_cmd_remove)

    sp = add("batch", "apply a JSON array of ops atomically")
    sp.add_argument("path")
    sp.add_argument("--operations", default=None, help="JSON array of ops")
    sp.add_argument("--file", default=None, help="path to a JSON ops file ('-' for stdin)")
    sp.set_defaults(func=_cmd_batch)

    sp = add("import", "import a CSV/TSV into an xlsx worksheet")
    sp.add_argument("out")
    sp.add_argument("sheet")
    sp.add_argument("csv")
    sp.add_argument("--header", action="store_true")
    sp.add_argument("--delimiter", default=None)
    sp.set_defaults(func=_cmd_import)

    sp = add("mcp", "start the MCP stdio server (requires officekit[mcp])")
    sp.set_defaults(func=_cmd_mcp)

    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001 - surface as structured error
        if getattr(args, "json", False):
            print(
                json.dumps(
                    {
                        "success": False,
                        "error": {"code": type(exc).__name__, "suggestion": str(exc)},
                    },
                    ensure_ascii=False,
                )
            )
        else:
            sys.stderr.write(f"error: {exc}\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
