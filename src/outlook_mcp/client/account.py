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
    """Resolve an Outlook Account by exact SMTP/display/user/store name.

    Exact matching is deliberate: silently falling back to Outlook's default
    account can send mail from the wrong legal entity and can also trigger
    account-specific add-ins such as CRM tracking.
    """
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
            raise OutlookError(
                f"Outlook account selector {selector!r} is ambiguous on {attr}."
            )

    store_matches = []
    for account in accounts:
        store = _safe_get(account, "DeliveryStore")
        if str(_safe_get(store, "DisplayName", "") or "").strip().casefold() == wanted:
            store_matches.append(account)
    if len(store_matches) == 1:
        return store_matches[0]
    if len(store_matches) > 1:
        raise OutlookError(
            f"Outlook account selector {selector!r} matches multiple delivery stores."
        )

    available = ", ".join(
        filter(
            None,
            [
                str(_safe_get(account, "SmtpAddress", "") or "")
                or str(_safe_get(account, "DisplayName", "") or "")
                for account in accounts
            ],
        )
    )
    raise OutlookError(
        f"Outlook account {selector!r} was not found. Available accounts: {available or '(none)'}."
    )


def set_send_using_account(item: Any, account: Any) -> dict[str, Any]:
    """Bind a MailItem to an Outlook Account and verify the binding.

    Outlook's SendUsingAccount is an object reference property. With pywin32,
    ordinary attribute assignment can appear to succeed while Outlook still
    sends through the default account. Use the underlying COM PROPERTYPUTREF
    call (DISPID 64209) first, then verify the getter before allowing Send().
    """
    try:
        item._oleobj_.Invoke(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, account)
    except Exception as raw_exc:
        try:
            item.SendUsingAccount = account
        except Exception as attr_exc:
            err = OutlookError(
                "Could not set Outlook SendUsingAccount; refusing to send via the default account."
            )
            err.__cause__ = attr_exc
            raise err from raw_exc

    selected = _safe_get(item, "SendUsingAccount")
    if selected is None:
        raise OutlookError(
            "Outlook did not expose a SendUsingAccount after selection; refusing to send via the default account."
        )

    expected = str(_safe_get(account, "SmtpAddress", "") or "").strip().casefold()
    actual = str(_safe_get(selected, "SmtpAddress", "") or "").strip().casefold()
    if expected and actual and expected != actual:
        raise OutlookError(
            f"Outlook selected {actual!r} instead of requested {expected!r}; refusing to send."
        )
    return account_summary(selected)


def whoami(outlook: Any, namespace: Any) -> dict[str, Any]:
    accounts = [account_summary(acct) for acct in outlook.Session.Accounts]
    # All datetimes this server returns are in this (the user's) local
    # timezone with an explicit offset — surface it so agents never guess.
    now = dt.datetime.now().astimezone()
    return {
        "current_user": _safe_get(namespace.CurrentUser, "Name"),
        "accounts": accounts,
        "local_time": now.isoformat(),
        "timezone": str(now.tzinfo),
        "utc_offset": now.strftime("%z"),
    }
