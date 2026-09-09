"""Tool registration with a secure read-only default."""

from outlook_mcp.config import writes_enabled

from . import (
    account,
    calendar,
    categories,
    contacts,
    folders,
    mail,
    ooo,
    readonly,
    rules,
    tasks,
)


def register_all(mcp, bridge) -> None:
    if not writes_enabled():
        readonly.register(mcp, bridge)
        return

    # Full mode preserves the upstream tool surface for users who explicitly
    # opt in with OUTLOOK_MCP_ACCESS=full before the server starts.
    for mod in (mail, folders, calendar, contacts, tasks, categories, rules, ooo, account):
        mod.register(mcp, bridge)
