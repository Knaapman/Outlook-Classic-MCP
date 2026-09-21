"""Regression tests for multi-account sender selection."""

import sys

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="client modules import pywin32"
)


class FakeStore:
    def __init__(self, display_name: str, store_id: str):
        self.DisplayName = display_name
        self.StoreID = store_id


class FakeAccount:
    def __init__(self, smtp: str, display: str | None = None):
        self.SmtpAddress = smtp
        self.DisplayName = display or smtp
        self.UserName = smtp.split("@", 1)[0]
        self.AccountType = 0
        self.DeliveryStore = FakeStore(self.DisplayName, f"store:{smtp}")


class FakeSession:
    def __init__(self, accounts):
        self.Accounts = accounts


class FakeOutlook:
    def __init__(self, accounts):
        self.Session = FakeSession(accounts)


class FakeOleObject:
    def __init__(self):
        self.calls = []

    def Invoke(self, *args):
        self.calls.append(args)


class FakeMailItem:
    def __init__(self):
        self._oleobj_ = FakeOleObject()


def test_resolve_send_account_matches_exact_smtp_case_insensitive():
    from outlook_mcp.client.account import resolve_send_account

    ado = FakeAccount("mike.timmerman@adopro.nl")
    wah = FakeAccount("mike@werkadvieshuis.nl")
    outlook = FakeOutlook([ado, wah])

    assert resolve_send_account(outlook, "MIKE@WERKADVIESHUIS.NL") is wah


def test_resolve_send_account_does_not_fall_back_to_default():
    from outlook_mcp.client.account import resolve_send_account
    from outlook_mcp.errors import OutlookError

    outlook = FakeOutlook([FakeAccount("mike.timmerman@adopro.nl")])

    with pytest.raises(OutlookError, match="was not found"):
        resolve_send_account(outlook, "mike@werkadvieshuis.nl")


def test_bind_send_account_uses_propertyputref_dispid_64209():
    from outlook_mcp.client.account import bind_send_account

    wah = FakeAccount("mike@werkadvieshuis.nl")
    item = FakeMailItem()

    result = bind_send_account(item, wah)

    assert item._oleobj_.calls == [(64209, 0, 8, 0, wah)]
    assert result["smtp_address"] == "mike@werkadvieshuis.nl"
