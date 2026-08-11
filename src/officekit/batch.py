"""Atomic batch execution of document operations.

The reference CLI's ``batch`` command applies several edits in one pass and
rolls everything back if any step fails. We replicate that with a temp copy:
operations run against the copy, and only on full success is it moved back
over the original — so a partial failure leaves the source untouched.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict, List

from . import ops
from .core import _detect_kind, merge_template


def _apply_one(file: Path, op: Dict[str, Any]) -> None:
    kind = op.get("op")
    if kind == "set":
        ops.set_value(file, op["target"], str(op["value"]))
    elif kind == "add":
        ops.add(file, op["what"], op.get("value"))
    elif kind == "remove":
        ops.remove(file, op["target"])
    elif kind == "merge":
        data = op["data"] if isinstance(op.get("data"), dict) else json.loads(
            op["data"]
        )
        merge_template(file, file, data)
    else:
        raise ValueError(f"Unknown batch op: {kind!r}")


def run_batch(path: str | Path, operations: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Apply *operations* atomically; raise on the first failure (rollback)."""
    path = Path(path)
    suffix = path.suffix
    tmp = Path(tempfile.mkstemp(suffix=suffix)[1])
    try:
        shutil.copy(path, tmp)
        for op in operations:
            _apply_one(tmp, op)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    shutil.move(str(tmp), str(path))
    return {"applied": len(operations), "file": str(path)}
