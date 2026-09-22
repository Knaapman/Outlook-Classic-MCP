"""Tests for CID inline attachment handling."""

import sys

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="client modules import pywin32"
)


class FakePropertyAccessor:
    def __init__(self):
        self.values = {}

    def SetProperty(self, name, value):
        self.values[name] = value


class FakeAttachment:
    def __init__(self, path):
        self.path = path
        self.PropertyAccessor = FakePropertyAccessor()


class FakeAttachments:
    def __init__(self):
        self.added = []

    def Add(self, path):
        att = FakeAttachment(path)
        self.added.append(att)
        return att


class FakeMail:
    def __init__(self):
        self.Attachments = FakeAttachments()


def test_add_inline_attachment_sets_content_id_hidden_and_mime(monkeypatch):
    from outlook_mcp.client import mail

    monkeypatch.setattr(mail, "validate_attachment_path", lambda path: path)
    item = FakeMail()

    mail._add_attachments(
        item,
        attachments=[r"C:\\tmp\\report.pdf"],
        inline_attachments=[
            {
                "path": r"C:\\tmp\\logo.png",
                "content_id": "ado-logo",
                "mime_type": "image/png",
            }
        ],
    )

    assert [a.path for a in item.Attachments.added] == [
        r"C:\\tmp\\report.pdf",
        r"C:\\tmp\\logo.png",
    ]
    inline = item.Attachments.added[1].PropertyAccessor.values
    assert inline[mail.PR_ATTACH_CONTENT_ID] == "ado-logo"
    assert inline[mail.PR_ATTACHMENT_HIDDEN] is True
    assert inline[mail.PR_ATTACH_MIME_TAG] == "image/png"


def test_add_inline_attachment_normalizes_angle_brackets(monkeypatch):
    from outlook_mcp.client import mail

    monkeypatch.setattr(mail, "validate_attachment_path", lambda path: path)
    item = FakeMail()

    mail._add_attachments(
        item,
        inline_attachments=[{"path": r"C:\\tmp\\logo.png", "content_id": "<ado-logo>"}],
    )

    values = item.Attachments.added[0].PropertyAccessor.values
    assert values[mail.PR_ATTACH_CONTENT_ID] == "ado-logo"


@pytest.mark.parametrize(
    "spec, message",
    [
        ({"content_id": "ado-logo"}, "path"),
        ({"path": r"C:\\tmp\\logo.png"}, "content_id"),
    ],
)
def test_add_inline_attachment_requires_path_and_content_id(monkeypatch, spec, message):
    from outlook_mcp.client import mail
    from outlook_mcp.errors import OutlookError

    monkeypatch.setattr(mail, "validate_attachment_path", lambda path: path)
    item = FakeMail()

    with pytest.raises(OutlookError, match=message):
        mail._add_attachments(item, inline_attachments=[spec])
