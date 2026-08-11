"""officekit - friendly Office document automation for humans and AI agents."""

from ._version import __version__
from .core import (
    create_document,
    document_info,
    extract_text,
    merge_template,
    validate_document,
)
from . import ops, view, batch, importer

__all__ = [
    "__version__",
    "extract_text",
    "create_document",
    "merge_template",
    "document_info",
    "validate_document",
    "ops",
    "view",
    "batch",
    "importer",
]
