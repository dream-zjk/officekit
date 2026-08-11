"""MCP server exposing officekit as agent-facing tools.

Run with ``officekit mcp`` (requires the ``mcp`` extra: ``pip install
officekit[mcp]``). The server speaks MCP over stdio so an AI agent (Claude
Code, Cursor, VS Code, ...) can read and edit Office documents without shell
access — the same role the reference project's ``officecli mcp`` command plays.

The whole module is import-guarded: importing it without ``mcp`` installed only
raises when you actually call :func:`run`, not at import time.
"""
from __future__ import annotations

import json
from typing import Any, Dict, List

from . import __version__
from . import core, ops, view
from .batch import run_batch
from .importer import import_csv

_TOOLS: List[Dict[str, Any]] = [
    {
        "name": "extract_text",
        "description": "Return all visible text of a .docx/.xlsx/.pptx file.",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "view_outline",
        "description": "Return the document skeleton (headings, sheets, slide titles).",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "view_stats",
        "description": "Return counts/dimensions for the document.",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "view_issues",
        "description": "List potential problems (empty result means healthy).",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "document_info",
        "description": "Return a summary dict of document shape.",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "validate_document",
        "description": "Return a list of issues; empty means the file opens cleanly.",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "create_document",
        "description": "Create a starter docx/xlsx/pptx.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "kind": {"type": "string", "enum": ["docx", "xlsx", "pptx"]},
                "path": {"type": "string"},
                "title": {"type": "string"},
            },
            "required": ["kind", "path"],
        },
    },
    {
        "name": "merge_template",
        "description": "Fill {{ key }} placeholders from a JSON object into a copy.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "template": {"type": "string"},
                "out": {"type": "string"},
                "data": {"type": "string", "description": "JSON object of values"},
            },
            "required": ["template", "out", "data"],
        },
    },
    {
        "name": "get",
        "description": "Resolve a path (e.g. /slides[1]/shapes[2]) and describe the element.",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}, "target": {"type": "string"}},
            "required": ["path", "target"],
        },
    },
    {
        "name": "set_value",
        "description": "Write text/value to the element addressed by a path; saves the file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "target": {"type": "string"},
                "value": {"type": "string"},
            },
            "required": ["path", "target", "value"],
        },
    },
    {
        "name": "add",
        "description": "Append an element: paragraph (docx), slide (pptx), row/sheet (xlsx).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "what": {"type": "string"},
                "value": {"type": "string"},
            },
            "required": ["path", "what"],
        },
    },
    {
        "name": "remove",
        "description": "Delete the element addressed by a path (docx paragraph / pptx shape).",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string"}, "target": {"type": "string"}},
            "required": ["path", "target"],
        },
    },
    {
        "name": "batch",
        "description": "Apply a JSON array of ops atomically ([set/add/remove/merge]); rolls back on failure.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "operations": {"type": "string", "description": "JSON array"},
            },
            "required": ["path", "operations"],
        },
    },
    {
        "name": "import_csv",
        "description": "Import a CSV/TSV file into an xlsx worksheet.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "out": {"type": "string"},
                "sheet_title": {"type": "string"},
                "csv_path": {"type": "string"},
                "header": {"type": "boolean"},
            },
            "required": ["out", "sheet_title", "csv_path"],
        },
    },
]

_DISPATCH = {
    "extract_text": lambda a: core.extract_text(a["path"]),
    "view_outline": lambda a: view.view_outline(a["path"]),
    "view_stats": lambda a: view.view_stats(a["path"]),
    "view_issues": lambda a: view.view_issues(a["path"]),
    "document_info": lambda a: core.document_info(a["path"]),
    "validate_document": lambda a: core.validate_document(a["path"]),
    "create_document": lambda a: str(
        core.create_document(a["kind"], a["path"], title=a.get("title"))
    ),
    "merge_template": lambda a: str(
        core.merge_template(a["template"], a["out"], json.loads(a["data"]))
    ),
    "get": lambda a: ops.get(a["path"], a["target"]),
    "set_value": lambda a: ops.set_value(a["path"], a["target"], a["value"]),
    "add": lambda a: ops.add(a["path"], a["what"], a.get("value")),
    "remove": lambda a: ops.remove(a["path"], a["target"]),
    "batch": lambda a: run_batch(a["path"], json.loads(a["operations"])),
    "import_csv": lambda a: str(
        import_csv(
            a["out"], a["sheet_title"], a["csv_path"], header=bool(a.get("header"))
        )
    ),
}


def _run(name: str, arguments: Dict[str, Any]) -> Any:
    return _DISPATCH[name](arguments or {})


def run() -> None:
    """Start the MCP stdio server. Requires the ``mcp`` package."""
    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    import mcp.types as types
    import anyio

    app = Server("officekit")

    @app.list_tools()
    async def list_tools():  # type: ignore[unused]
        return [types.Tool(**t) for t in _TOOLS]

    @app.call_tool()
    async def call_tool(name: str, arguments: Dict[str, Any]):  # type: ignore[unused]
        try:
            result = _run(name, arguments)
        except Exception as exc:  # noqa: BLE001
            result = {"success": False, "error": {"type": type(exc).__name__, "message": str(exc)}}
        if not isinstance(result, str):
            result = json.dumps(result, ensure_ascii=False, indent=2)
        return [types.TextContent(type="text", text=result)]

    async def _main() -> None:
        async with stdio_server() as (read, write):
            await app.run(read, write, app.create_initialization_options())

    anyio.run(_main)


__all__ = ["run", "__version__"]
