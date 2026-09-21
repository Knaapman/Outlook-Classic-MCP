"""Store- and account-aware write helpers for the hardened Outlook MCP.

These functions are only exposed when ``OUTLOOK_MCP_ACCESS=full``.  They keep
item identity tied to StoreID and allow outbound actions to select an Outlook
account explicitly, which is important on profiles containing work and private
mailboxes.
"""

from __future__ import annotations

import ntpath
import os
from typing import Any

from outlook_mcp.client import calendar as cal_client
from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_DRAFTS,
    OL_FORMAT_HTML,
    OL_FORMAT_PLAIN,
    OL_IMPORTANCE_NORMAL,
    OL_MAIL_ITEM,
    OL_MEETING,
    OL_MEETING_ACCEPTED,
    OL_MEETING_DECLINED,
    OL_MEETING_TENTATIVE,
    OL_TASK_ITEM,
    OL_TO,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import from_iso, to_iso
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir

SEND_USING_ACCOUNT_DISPID = 64209


WINDOWS_RESERVED_DEVICE_NAMES = {"CON", "PRN", "AUX", "NUL", "CLOCK$"} | {
    f"COM{i}" for i in range(1, 10)
} | {f"LPT{i}" for i in range(1, 10)}


def _find_account(outlook: Any, account_ref: str | None) -> Any | None:
    if not account_ref:
        return None
    needle = account_ref.strip().lower()
    available: list[str] = []
    for account in outlook.Session.Accounts:
        values = [
            str(_safe_get(account, "SmtpAddress", "") or ""),
            str(_safe_get(account, "DisplayName", "") or ""),
            str(_safe_get(account, "UserName", "") or ""),
        ]
        available.extend(v for v in values if v)
        if any(v.lower() == needle for v in values if v):
            return account
    raise OutlookError(
        f"Outlook account '{account_ref}' not found. Available account identifiers: "
        + ", ".join(dict.fromkeys(available))
    )


def _set_send_account(item: Any, outlook: Any, account_ref: str | None) -> Any | None:
    """Bind an outbound Outlook item to an exact account and fail closed.

    pywin32 can silently ignore ordinary attribute assignment to
    SendUsingAccount on Outlook items.  Outlook exposes this property as
    DISPID 64209; using the low-level COM PROPERTYPUTREF call is the reliable
    route for Account object references.
    """
    account = _find_account(outlook, account_ref)
    if account is None:
        return None

    try:
        item._oleobj_.Invoke(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, account)
    except Exception as exc:
        raise OutlookError(
            f"Could not select Outlook sending account '{account_ref}'."
        ) from exc

    selected = _safe_get(item, "SendUsingAccount")
    if selected is None:
        raise OutlookError(
            f"Outlook did not retain sending account '{account_ref}'; refusing to use the default account."
        )

    expected = str(_safe_get(account, "SmtpAddress", "") or "").strip().casefold()
    actual = str(_safe_get(selected, "SmtpAddress", "") or "").strip().casefold()
    if expected and actual != expected:
        raise OutlookError(
            f"Outlook retained sending account '{actual or '(unknown)'}' instead of '{expected}'; "
            "refusing to send from the default account."
        )
    return selected


def _item_store_id(item: Any) -> str | None:
    parent = _safe_get(item, "Parent")
    sid = _safe_get(parent, "StoreID") if parent is not None else None
    return str(sid) if sid else None


def send_mail(
    outlook: Any,
    namespace: Any,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    save_only: bool = False,
    send_using_account: str | None = None,
) -> dict[str, Any]:
    account = _find_account(outlook, send_using_account)
    target_store_id = None

    if account is not None:
        delivery_store = _safe_get(account, "DeliveryStore")
        drafts = _safe_get(delivery_store, "GetDefaultFolder")
        try:
            drafts_folder = delivery_store.GetDefaultFolder(OL_FOLDER_DRAFTS)
            mail = drafts_folder.Items.Add()
            target_store_id = str(_safe_get(drafts_folder, "StoreID") or "")
        except Exception as exc:
            raise OutlookError(
                f"Could not create mail in the Drafts folder for '{send_using_account}'."
            ) from exc
        _set_send_account(mail, outlook, send_using_account)
    else:
        mail = outlook.CreateItem(OL_MAIL_ITEM)

    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    if bcc:
        mail.BCC = "; ".join(bcc)
    mail.Subject = subject
    if html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = body
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body
    mail.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)
    for raw_path in attachments or []:
        mail.Attachments.Add(validate_attachment_path(raw_path))

    # Save once so Outlook materializes the item in the selected store. This is
    # also a pre-send safety check: a requested non-default account must not
    # silently become an item in the default mailbox.
    if account is not None:
        mail.Save()
        actual_store_id = _item_store_id(mail)
        if target_store_id and actual_store_id != target_store_id:
            raise OutlookError(
                f"Mail was created in store '{actual_store_id or '(unknown)'}' instead of the "
                f"requested account store '{target_store_id}'; refusing to send."
            )

    if save_only:
        if account is None:
            mail.Save()
        return {
            "status": "saved_to_drafts",
            "entry_id": _safe_get(mail, "EntryID"),
            "store_id": _item_store_id(mail),
            "subject": subject,
            "send_using_account": send_using_account,
        }

    mail.Send()
    return {
        "status": "sent",
        "to": to,
        "cc": cc or [],
        "bcc": bcc or [],
        "subject": subject,
        "send_using_account": send_using_account,
    }


def reply_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    attachments: list[str] | None = None,
    send_using_account: str | None = None,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id, store_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    _set_send_account(reply, outlook, send_using_account)
    if html:
        reply.BodyFormat = OL_FORMAT_HTML
        reply.HTMLBody = body + (reply.HTMLBody or "")
    else:
        reply.Body = body + "\n\n" + (reply.Body or "")
    for raw_path in attachments or []:
        reply.Attachments.Add(validate_attachment_path(raw_path))
    reply_subject = reply.Subject
    reply.Send()
    return {
        "status": "sent",
        "reply_all": reply_all,
        "in_reply_to": entry_id,
        "subject": reply_subject,
        "send_using_account": send_using_account,
    }


def forward_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    html: bool = False,
    send_using_account: str | None = None,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id, store_id)
    fwd = original.Forward()
    _set_send_account(fwd, outlook, send_using_account)
    fwd.To = "; ".join(to)
    if cc:
        fwd.CC = "; ".join(cc)
    if body:
        if html:
            fwd.BodyFormat = OL_FORMAT_HTML
            fwd.HTMLBody = body + (fwd.HTMLBody or "")
        else:
            fwd.Body = body + "\n\n" + (fwd.Body or "")
    subject = fwd.Subject
    fwd.Send()
    return {
        "status": "sent",
        "forwarded": entry_id,
        "to": to,
        "subject": subject,
        "send_using_account": send_using_account,
    }


def move_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    target_folder: str,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id, store_id)
    target = resolve_folder(namespace, target_folder)
    moved = item.Move(target)
    return {
        "status": "moved",
        "new_entry_id": moved.EntryID,
        "store_id": _item_store_id(moved),
        "folder": target_folder,
    }


def delete_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id, store_id)
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def mark_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    read: bool | None = None,
    flagged: bool | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id, store_id)
    if read is not None:
        item.UnRead = not read
    if flagged is not None:
        item.FlagStatus = 2 if flagged else 0
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "store_id": store_id or _item_store_id(item),
        "unread": bool(item.UnRead),
        "flagged": item.FlagStatus == 2,
    }


def save_attachments(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    output_dir: str,
    attachment_index: int | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id, store_id)
    out_dir = validate_output_dir(output_dir)
    attachments = list(item.Attachments)
    if attachment_index is not None:
        if attachment_index < 1 or attachment_index > len(attachments):
            raise OutlookError(
                f"attachment_index {attachment_index} out of range "
                f"(message has {len(attachments)} attachments, 1-indexed)."
            )
        attachments = [attachments[attachment_index - 1]]

    saved: list[str] = []
    for att in attachments:
        raw = att.FileName or ""
        if not raw or raw in (".", "..") or "\\" in raw or "/" in raw or ":" in raw:
            raise OutlookError(f"Attachment has unsafe filename: {raw!r}")
        safe_name = ntpath.basename(raw)
        if safe_name != raw:
            raise OutlookError(f"Attachment filename did not normalize cleanly: {raw!r}")
        stem = safe_name.lstrip(".").split(".", 1)[0].upper()
        if stem in WINDOWS_RESERVED_DEVICE_NAMES:
            raise OutlookError(f"Attachment has reserved Windows device name: {safe_name!r}")
        target = os.path.join(out_dir, safe_name)
        base, ext = os.path.splitext(safe_name)
        counter = 1
        while os.path.exists(target):
            target = os.path.join(out_dir, f"{base} ({counter}){ext}")
            counter += 1
        att.SaveAsFile(target)
        saved.append(target)
    return {"status": "saved", "count": len(saved), "files": saved, "output_dir": out_dir}


def create_event(
    outlook: Any,
    namespace: Any,
    *,
    folder: str = "calendar",
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    reminder_minutes: int | None = 15,
    recurrence: Recurrence | None = None,
    send_using_account: str | None = None,
) -> dict[str, Any]:
    cal = resolve_folder(namespace, folder)
    # Creating through the target folder keeps the item in the intended store.
    try:
        appt = cal.Items.Add(OL_APPOINTMENT_ITEM)
    except Exception:
        appt = cal.Items.Add()
    _set_send_account(appt, outlook, send_using_account)
    appt.Subject = subject
    appt.Start = from_iso(start)
    appt.End = from_iso(end)
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    if reminder_minutes is not None:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    if attendees:
        appt.MeetingStatus = OL_MEETING
        for addr in attendees:
            rec = appt.Recipients.Add(addr)
            rec.Type = OL_TO
        appt.Recipients.ResolveAll()
    if recurrence is not None:
        cal_client._apply_recurrence(appt, recurrence)
    appt.Save()
    if attendees:
        appt.Send()
    return {
        "status": "created",
        "entry_id": _safe_get(appt, "EntryID"),
        "store_id": _item_store_id(appt),
        "folder": folder,
        "subject": subject,
        "start": to_iso(_safe_get(appt, "Start")),
        "end": to_iso(_safe_get(appt, "End")),
        "send_using_account": send_using_account,
    }


def update_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
) -> dict[str, Any]:
    event = get_item_by_id(namespace, entry_id, store_id)
    if subject is not None:
        event.Subject = subject
    if start is not None:
        event.Start = from_iso(start)
    if end is not None:
        event.End = from_iso(end)
    if location is not None:
        event.Location = location
    if body is not None:
        event.Body = body
    event.Save()
    return {"status": "updated", "entry_id": entry_id, "store_id": store_id or _item_store_id(event)}


def delete_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
) -> dict[str, Any]:
    event = get_item_by_id(namespace, entry_id, store_id)
    subject = _safe_get(event, "Subject", "")
    event.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def respond_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    response: str,
    send_response: bool = True,
    send_using_account: str | None = None,
) -> dict[str, Any]:
    event = get_item_by_id(namespace, entry_id, store_id)
    code = {
        "accept": OL_MEETING_ACCEPTED,
        "tentative": OL_MEETING_TENTATIVE,
        "decline": OL_MEETING_DECLINED,
    }.get(response.lower())
    if code is None:
        raise OutlookError("response must be one of: 'accept', 'tentative', 'decline'.")
    reply = event.Respond(code, True)
    if reply is not None:
        _set_send_account(reply, outlook, send_using_account)
    if send_response and reply is not None:
        reply.Send()
    return {
        "status": "responded",
        "response": response,
        "send_response": send_response,
        "send_using_account": send_using_account,
    }


def create_task(
    outlook: Any,
    namespace: Any,
    *,
    folder: str = "tasks",
    subject: str,
    due_date: str | None = None,
    body: str | None = None,
    importance: str = "normal",
    reminder: str | None = None,
) -> dict[str, Any]:
    task_folder = resolve_folder(namespace, folder)
    try:
        task = task_folder.Items.Add(OL_TASK_ITEM)
    except Exception:
        task = task_folder.Items.Add()
    task.Subject = subject
    if due_date:
        task.DueDate = from_iso(due_date)
    if body:
        task.Body = body
    task.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)
    if reminder:
        task.ReminderSet = True
        task.ReminderTime = from_iso(reminder)
    task.Save()
    return {
        "status": "created",
        "entry_id": task.EntryID,
        "store_id": _item_store_id(task),
        "folder": folder,
        "subject": subject,
    }


def complete_task(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
) -> dict[str, Any]:
    task = get_item_by_id(namespace, entry_id, store_id)
    task.MarkComplete()
    task.Save()
    return {"status": "completed", "entry_id": entry_id, "store_id": store_id or _item_store_id(task)}


def set_category(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    store_id: str | None = None,
    categories: str,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id, store_id)
    item.Categories = categories
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "store_id": store_id or _item_store_id(item),
        "categories": item.Categories or "",
    }


def _find_store(namespace: Any, store_ref: str | None) -> Any:
    if not store_ref:
        return namespace.DefaultStore
    needle = store_ref.strip().lower()
    for store in namespace.Stores:
        try:
            if str(store.StoreID).lower() == needle or str(store.DisplayName).lower() == needle:
                return store
        except Exception:
            continue
    raise OutlookError(f"Outlook store '{store_ref}' not found. Use outlook_list_stores first.")


def toggle_rule(
    outlook: Any,
    namespace: Any,
    *,
    rule_name: str,
    enabled: bool,
    store_ref: str | None = None,
) -> dict[str, Any]:
    store = _find_store(namespace, store_ref)
    rules = store.GetRules()
    for i in range(rules.Count):
        rule = rules.Item(i + 1)
        if rule.Name == rule_name:
            rule.Enabled = enabled
            rules.Save()
            return {
                "status": "updated",
                "rule": rule_name,
                "enabled": bool(enabled),
                "store": _safe_get(store, "DisplayName"),
                "store_id": _safe_get(store, "StoreID"),
            }
    raise OutlookError(f"Rule '{rule_name}' not found in the selected store.")
