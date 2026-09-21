"""Register the hardened Outlook MCP surfaces.

Read tools are always available. Write tools are layered on top only when
``OUTLOOK_MCP_ACCESS=full`` is set before server startup.
"""

from outlook_mcp.config import writes_enabled

from . import readonly, write


def register_all(mcp, bridge) -> None:
    readonly.register(mcp, bridge)
    if writes_enabled():
        write.register(mcp, bridge)
