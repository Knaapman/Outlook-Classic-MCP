"""Store-aware read helpers for Outlook profiles with multiple mounted stores."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client import calendar as cal_client
from outlook_mcp.client import contacts as contacts_client
from outlook_mcp.client import mail as mail_client
from outlook_mcp.client import tasks as tasks_client
from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.constants import OL_CLASS_TASK
from outlook_mcp.utils.formatting import from_iso


def _store_id(folder: Any) -> str | None:
    value = _safe_get(folder, "StoreID")
    return str(value) if value else None


def _store_name(namespace: Any, store_id: str | None) -> str | None:
    if not store_id:
        return None
    for store in namespace.Stores:
        try:
            if str(store.StoreID) == store_id:
                return str(store.DisplayName)
        except Exception:
            continue
    return None


def _annotate_items(data: dict[str, Any], *, store_id: str | None, store_name: str | None, folder_ref: str) -> dict[str, Any]:
    data["store_id"] = store_id
    data["store"] = store_name
    data["folder_ref"] = folder_ref
    for row in data.get("items", []):
        row["store_id"] = store_id
        row["store"] = store_name
        row["folder_ref"] = folder_ref
    return data


def list_stores(outlook: Any, namespace: Any) -> dict[str, Any]:
    items: list[dict[str, Any]] = []
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
        except Exception:
            root = None
        items.append({
            "display_name": _safe_get(store, "DisplayName"),
            "store_id": _safe_get(store, "StoreID"),
            "exchange_store_type": _safe_get(store, "ExchangeStoreType"),
            "is_data_file_store": bool(_safe_get(store, "IsDataFileStore", False)),
            "root_folder": _safe_get(root, "Name") if root is not None else None,
        })
    return {"count": len(items), "items": items}


def list_folders(outlook: Any, namespace: Any, *, root: str | None = None, max_depth: int = 4) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []

    def walk(folder: Any, path: str, depth: int, store_id: str | None, store_name: str | None) -> None:
        items_obj = _safe_get(folder, "Items")
        out.append({
            "name": _safe_get(folder, "Name"),
            "path": path,
            "store_id": store_id,
            "store": store_name,
            "item_count": _safe_get(items_obj, "Count", 0) if items_obj else 0,
            "unread_count": _safe_get(folder, "UnReadItemCount", 0),
            "default_item_type": _safe_get(folder, "DefaultItemType", -1),
        })
        if depth >= max_depth:
            return
        try:
            children = list(folder.Folders)
        except Exception:
            return
        for sub in children:
            walk(sub, f"{path}/{sub.Name}", depth + 1, store_id, store_name)

    if root:
        start = resolve_folder(namespace, root)
        sid = _store_id(start)
        sname = _store_name(namespace, sid)
        start_path = root if "/" in root else (f"{sname}/{start.Name}" if sname else start.Name)
        walk(start, start_path, 0, sid, sname)
        return out

    for store in namespace.Stores:
        try:
            start = store.GetRootFolder()
            sid = str(store.StoreID)
            sname = str(store.DisplayName)
        except Exception:
            continue
        walk(start, sname, 0, sid, sname)
    return out


def list_mails(outlook: Any, namespace: Any, **kwargs: Any) -> dict[str, Any]:
    folder_ref = str(kwargs.get("folder") or "inbox")
    data = mail_client.list_mails(outlook, namespace, **kwargs)
    folder = resolve_folder(namespace, folder_ref)
    sid = _store_id(folder)
    return _annotate_items(data, store_id=sid, store_name=_store_name(namespace, sid), folder_ref=folder_ref)


def search_mails(outlook: Any, namespace: Any, **kwargs: Any) -> dict[str, Any]:
    folder_ref = str(kwargs.get("folder") or "inbox")
    data = mail_client.search_mails(outlook, namespace, **kwargs)
    folder = resolve_folder(namespace, folder_ref)
    sid = _store_id(folder)
    return _annotate_items(data, store_id=sid, store_name=_store_name(namespace, sid), folder_ref=folder_ref)


def get_mail(outlook: Any, namespace: Any, *, entry_id: str, store_id: str | None = None, include_body: bool = True, include_html: bool = False, max_body_chars: int = 10000) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id, store_id)
    data = mail_client._mail_full(item, include_body=include_body, include_html=include_html, max_body_chars=max_body_chars)
    data["store_id"] = store_id or _safe_get(_safe_get(item, "Parent"), "StoreID")
    return data


def list_events(outlook: Any, namespace: Any, *, folder: str = "calendar", start: str | None = None, end: str | None = None, limit: int = 50, include_recurrences: bool = True) -> dict[str, Any]:
    cal = resolve_folder(namespace, folder)
    items = cal.Items
    items.Sort("[Start]")
    if include_recurrences:
        items.IncludeRecurrences = True
    start_dt = from_iso(start) or dt.datetime.now()
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=14))
    restrict = f"[Start] >= '{start_dt.strftime('%m/%d/%Y %I:%M %p')}' AND [Start] <= '{end_dt.strftime('%m/%d/%Y %I:%M %p')}'"
    filtered = items.Restrict(restrict)
    sid = _store_id(cal)
    sname = _store_name(namespace, sid)
    results: list[dict[str, Any]] = []
    for event in filtered:
        row = cal_client._event_summary(event)
        row.update({"store_id": sid, "store": sname, "folder_ref": folder})
        results.append(row)
        if len(results) >= limit:
            break
    return {"folder_ref": folder, "store_id": sid, "store": sname, "start": start_dt.isoformat(), "end": end_dt.isoformat(), "count": len(results), "items": results}


def get_event(outlook: Any, namespace: Any, *, entry_id: str, store_id: str | None = None) -> dict[str, Any]:
    event = get_item_by_id(namespace, entry_id, store_id)
    data = cal_client._event_full(event)
    data["store_id"] = store_id or _safe_get(_safe_get(event, "Parent"), "StoreID")
    return data


def list_tasks(outlook: Any, namespace: Any, *, folder: str = "tasks", limit: int = 50, include_completed: bool = False) -> dict[str, Any]:
    task_folder = resolve_folder(namespace, folder)
    items = task_folder.Items
    items.Sort("[DueDate]")
    sid = _store_id(task_folder)
    sname = _store_name(namespace, sid)
    results: list[dict[str, Any]] = []
    for task in items:
        if _safe_get(task, "Class") != OL_CLASS_TASK:
            continue
        if not include_completed and _safe_get(task, "Complete", False):
            continue
        row = tasks_client._task_summary(task)
        row.update({"store_id": sid, "store": sname, "folder_ref": folder})
        results.append(row)
        if len(results) >= limit:
            break
    return {"folder_ref": folder, "store_id": sid, "store": sname, "count": len(results), "items": results}


def _store_id_from_path(namespace: Any, path: str | None) -> str | None:
    if not path:
        return None
    matches: list[tuple[int, str]] = []
    for store in namespace.Stores:
        try:
            name = str(store.DisplayName)
            if path == name or path.startswith(name + "/"):
                matches.append((len(name), str(store.StoreID)))
        except Exception:
            continue
    return max(matches)[1] if matches else None


def list_contacts(outlook: Any, namespace: Any, **kwargs: Any) -> dict[str, Any]:
    data = contacts_client.list_contacts(outlook, namespace, **kwargs)
    for row in data.get("items", []):
        row["store_id"] = _store_id_from_path(namespace, row.get("folder"))
    return data


def search_contacts(outlook: Any, namespace: Any, **kwargs: Any) -> dict[str, Any]:
    data = contacts_client.search_contacts(outlook, namespace, **kwargs)
    for row in data.get("items", []):
        if row.get("source") == "contacts":
            row["store_id"] = _store_id_from_path(namespace, row.get("folder"))
    return data


def get_contact(outlook: Any, namespace: Any, *, entry_id: str, store_id: str | None = None) -> dict[str, Any]:
    contact = get_item_by_id(namespace, entry_id, store_id)
    data = contacts_client._contact_full(contact)
    data["store_id"] = store_id or _safe_get(_safe_get(contact, "Parent"), "StoreID")
    return data
