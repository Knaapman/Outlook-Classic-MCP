"""Runtime security configuration.

The hardened fork defaults to read-only access. Full write access must be
explicitly enabled in the local environment before the server starts.
"""

from __future__ import annotations

import os

_READ_ALIASES = {"read", "readonly", "read-only", "safe"}
_FULL_ALIASES = {"full", "write", "readwrite", "read-write"}


def access_mode() -> str:
    """Return ``read`` (default) or ``full`` from ``OUTLOOK_MCP_ACCESS``."""
    raw = os.environ.get("OUTLOOK_MCP_ACCESS", "read").strip().lower()
    if raw in _READ_ALIASES:
        return "read"
    if raw in _FULL_ALIASES:
        return "full"
    raise RuntimeError(
        "Invalid OUTLOOK_MCP_ACCESS value. Use 'read' (default) or 'full'."
    )


def writes_enabled() -> bool:
    """Whether mutating/outbound Outlook tools should be registered."""
    return access_mode() == "full"
