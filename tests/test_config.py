from __future__ import annotations

import os

import pytest

from outlook_mcp.config import access_mode, writes_enabled


def test_default_access_is_read_only(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("OUTLOOK_MCP_ACCESS", raising=False)
    assert access_mode() == "read"
    assert writes_enabled() is False


def test_full_access_requires_explicit_opt_in(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_ACCESS", "full")
    assert access_mode() == "full"
    assert writes_enabled() is True


def test_invalid_access_mode_fails_closed(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_ACCESS", "banana")
    with pytest.raises(RuntimeError):
        access_mode()
