"""Account / metadata COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get
from outlook_mcp.errors import OutlookError


SEND_USING_ACCOUNT_DISPID = 64209


def account_summary(account: Any) -> dict[str, Any]:
    """Return stable account metadata without leaking COM objects."""
    delivery_store = _safe_get(account, "DeliveryStore")
    return {
        "display_name": _safe_get(account, "DisplayName"),
        "smtp_address": _safe_get(account, "SmtpAddress"),
        "user_name": _safe_get(account, "UserName"),
        "account_type": _safe_get(account, "AccountType"),
        "store_id": _safe_get(delivery_store, "StoreID") if delivery_store is not None else None,
        "store_display_name": _safe_get(delivery_store, "DisplayName") if delivery_store is not None else None,
    }


def resolve_send_account(outlook: Any, selector: str) -> Any:
    """Resolve an Outlook Account by exact SMTP/display/user/store name."""
    wanted = (selector or "").strip().casefold()
    if not wanted:
        raise OutlookError("send_using_account cannot be empty.")

    accounts = list(outlook.Session.Accounts)
    for attr in ("SmtpAddress", "DisplayName", "UserName"):
        matches = [
            account
            for account in accounts
            if str(_safe_get(account, attr, "") or "").strip().casefold() == wanted
        ]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise OutlookError(f"Outlook account selector {selector!r} is ambiguous on {attr}.")

    store_matches = []
    for account in accounts:
        store = _safe_get(account, "DeliveryStore")
        if str(_safe_get(store, "DisplayName", "") or "").strip().casefold() == wanted:
            store_matches.append(account)
    if len(store_matches) == 1:
        return store_matches[0]
    if len(store_matches) > 1:
        raise OutlookError(f"Outlook account selector {selector!r} matches multiple delivery stores.")

    available = ", ".join(
        str(_safe_get(account, "SmtpAddress", "") or _safe_get(account, "DisplayName", "") or "")
        for account in accounts
    )
    raise OutlookError(
        f"Outlook account {selector!r} was not found. Available accounts: {available or '(none)'}."
    )


def bind_send_account(item: Any, account: Any) -> dict[str, Any]:
    """Bind an Outlook item to an Account using COM PROPERTYPUTREF.

    Direct attribute assignment can silently leave Outlook on the default
    account. DISPID 64209 is SendUsingAccount; flag 8 is PROPERTYPUTREF.
    Fail closed if binding cannot be applied.
    """
    try:
        item._oleobj_.Invoke(*(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, account))
    except Exception as exc:
        err = OutlookError(
            "Could not bind Outlook SendUsingAccount; refusing to fall back to the default account."
        )
        err.__cause__ = exc
        raise err
    return account_summary(account)



def whoami(outlook: Any, namespace: Any) -> dict[str, Any]:
    accounts = [account_summary(acct) for acct in outlook.Session.Accounts]
    now = dt.datetime.now().astimezone()
    return {
        "current_user": _safe_get(namespace.CurrentUser, "Name"),
        "accounts": accounts,
        "local_time": now.isoformat(),
        "timezone": str(now.tzinfo),
        "utc_offset": now.strftime("%z"),
    }
