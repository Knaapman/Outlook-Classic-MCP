from __future__ import annotations

import pytest

from outlook_mcp.tools import register_all


class RecorderMCP:
    def __init__(self) -> None:
        self.tools: dict[str, dict[str, object]] = {}

    def tool(self, *, name: str, annotations=None, **kwargs):
        def decorator(fn):
            if name in self.tools:
                raise AssertionError(f"duplicate tool registration: {name}")
            self.tools[name] = dict(annotations or {})
            return fn

        return decorator


def test_read_mode_registers_only_read_tools(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_ACCESS", "read")
    mcp = RecorderMCP()
    register_all(mcp, bridge=None)

    assert "outlook_list_stores" in mcp.tools
    assert "outlook_search_mails" in mcp.tools
    assert "outlook_send_mail" not in mcp.tools
    assert "outlook_delete_mail" not in mcp.tools
    assert all(meta.get("readOnlyHint") is True for meta in mcp.tools.values())


def test_full_mode_layers_writes_on_hardened_reads(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_ACCESS", "full")
    mcp = RecorderMCP()
    register_all(mcp, bridge=None)

    # Hardened multi-store reads remain present in full mode.
    assert "outlook_list_stores" in mcp.tools
    assert mcp.tools["outlook_list_stores"]["readOnlyHint"] is True

    # Write tools are added without replacing or duplicating read tools.
    assert "outlook_send_mail" in mcp.tools
    assert "outlook_delete_mail" in mcp.tools
    assert "outlook_create_event" in mcp.tools
    assert mcp.tools["outlook_send_mail"]["readOnlyHint"] is False
    assert mcp.tools["outlook_send_mail"]["openWorldHint"] is True
    assert mcp.tools["outlook_delete_mail"]["destructiveHint"] is True


def test_all_full_mode_writes_are_marked_non_read_only(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_ACCESS", "full")
    mcp = RecorderMCP()
    register_all(mcp, bridge=None)

    write_names = {
        "outlook_send_mail",
        "outlook_reply_mail",
        "outlook_forward_mail",
        "outlook_move_mail",
        "outlook_delete_mail",
        "outlook_mark_mail",
        "outlook_save_attachments",
        "outlook_create_folder",
        "outlook_create_event",
        "outlook_update_event",
        "outlook_delete_event",
        "outlook_respond_event",
        "outlook_create_task",
        "outlook_complete_task",
        "outlook_set_category",
        "outlook_toggle_rule",
    }
    assert write_names.issubset(mcp.tools)
    assert all(mcp.tools[name].get("readOnlyHint") is False for name in write_names)
