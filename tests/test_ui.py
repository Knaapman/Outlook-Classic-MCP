"""MCP Apps UI helpers and hardened server registration policy."""

from __future__ import annotations

import asyncio
import re

import pytest

from outlook_mcp import ui as ui_mod
from outlook_mcp.server import build_server


@pytest.mark.parametrize("access_mode", ["read", "full"])
def test_hardened_server_registers_no_interactive_ui(
    monkeypatch: pytest.MonkeyPatch, access_mode: str
) -> None:
    """Interactive UI controls must not bypass ChatGPT write approvals.

    The upstream UI contains mark-read, flag, delete, and task-completion
    controls. The hardened server deliberately exposes actions only as MCP
    tools, so neither read nor full mode registers UI resources or UI metadata.
    """
    monkeypatch.setenv("OUTLOOK_MCP_ACCESS", access_mode)
    server, _bridge = build_server()

    tools = asyncio.run(server.list_tools())
    assert tools
    for tool in tools:
        assert not (tool.meta or {}).get("ui"), tool.name

    resources = asyncio.run(server.list_resources())
    assert not [r for r in resources if str(r.uri).startswith("ui://outlook/")]


def test_ui_templates_remain_self_contained_for_future_safe_views() -> None:
    """Keep the upstream rendering helpers valid without registering them."""
    assert ui_mod._UI_DIR.joinpath("_common.html").exists()
    for view in ui_mod._VIEWS:
        html = ui_mod._render(view)
        assert html.lstrip().startswith("<!DOCTYPE html>")
        assert "ui/initialize" in html
        assert not re.search(r'<script[^>]+src\s*=', html)
        assert not re.search(r'<link[^>]+href\s*=', html)
        assert "<!--%COMMON%-->" not in html
        assert "appInfo:" in html
        assert "clientInfo" not in html


def test_ui_result_carries_both_representations():
    res = ui_mod.ui_result("**markdown**", {"items": [], "count": 0})
    assert res.content[0].text == "**markdown**"
    assert res.structuredContent == {"items": [], "count": 0}
    assert not res.isError


def test_ui_meta_rejects_unknown_view():
    with pytest.raises(ValueError):
        ui_mod.ui_meta("nope")
