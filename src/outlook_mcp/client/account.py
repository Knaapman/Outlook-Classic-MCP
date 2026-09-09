"""Account / metadata COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get


def whoami(outlook: Any, namespace: Any) -> dict[str, Any]:
    accounts = []
    for acct in outlook.Session.Accounts:
        delivery_store = _safe_get(acct, "DeliveryStore")
        accounts.append(
            {
                "display_name": _safe_get(acct, "DisplayName"),
                "smtp_address": _safe_get(acct, "SmtpAddress"),
                "user_name": _safe_get(acct, "UserName"),
                "account_type": _safe_get(acct, "AccountType"),
                "store_id": _safe_get(delivery_store, "StoreID") if delivery_store is not None else None,
                "store_display_name": _safe_get(delivery_store, "DisplayName") if delivery_store is not None else None,
            }
        )
    now = dt.datetime.now().astimezone()
    return {
        "current_user": _safe_get(namespace.CurrentUser, "Name"),
        "accounts": accounts,
        "local_time": now.isoformat(),
        "timezone": str(now.tzinfo),
        "utc_offset": now.strftime("%z"),
    }
