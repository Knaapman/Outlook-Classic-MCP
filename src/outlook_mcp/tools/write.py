"""Explicit write/modify MCP tools for Outlook.

This module is registered only when ``OUTLOOK_MCP_ACCESS=full``. Read tools are
registered separately and remain the same in both modes. Every tool here is
annotated as a write so ChatGPT can apply its native action-permission and
confirmation UX.
"""

from __future__ import annotations

from typing import Annotated, Literal, Optional

from pydantic import Field

from outlook_mcp.client import folders as folders_client
from outlook_mcp.client import multiwrite
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def _write_annotations(
    title: str,
    *,
    destructive: bool = False,
    external: bool = False,
    idempotent: bool = False,
) -> dict[str, object]:
    return {
        "title": title,
        "readOnlyHint": False,
        "destructiveHint": destructive,
        "idempotentHint": idempotent,
        "openWorldHint": external,
    }


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_send_mail",
        annotations=_write_annotations("Send Outlook mail", external=True),
    )
    @safe_call
    async def outlook_send_mail(
        to: Annotated[list[str], Field(min_length=1, description="Recipient email addresses.")],
        subject: Annotated[str, Field(description="Subject line.")],
        body: Annotated[str, Field(description="Message body; plain text unless html=true.")],
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="BCC recipients.")] = None,
        html: bool = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute local paths under the user profile.")] = None,
        importance: Literal["low", "normal", "high"] = "normal",
        save_only: Annotated[bool, Field(description="Save to Drafts instead of sending.")] = False,
        send_using_account: Annotated[
            Optional[str],
            Field(description="Exact Outlook SMTP address or account display name. Omit only to use Outlook's default account."),
        ] = None,
    ) -> str:
        """Create a draft or send mail. This changes Outlook and sending communicates externally."""
        data = await bridge.call(
            multiwrite.send_mail,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
            importance=importance,
            save_only=save_only,
            send_using_account=send_using_account,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_reply_mail",
        annotations=_write_annotations("Reply to Outlook mail", external=True),
    )
    @safe_call
    async def outlook_reply_mail(
        entry_id: Annotated[str, Field(min_length=1, description="EntryID returned by mail read/search tools.")],
        body: Annotated[str, Field(description="Reply body; quoted original is appended by Outlook.")],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by the read tool; strongly recommended for non-default stores.")] = None,
        reply_all: bool = False,
        html: bool = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute local paths under the user profile.")] = None,
        send_using_account: Annotated[Optional[str], Field(description="Optional exact Outlook SMTP address or account display name.")] = None,
    ) -> str:
        """Reply to an existing message. This sends an external communication."""
        data = await bridge.call(
            multiwrite.reply_mail,
            entry_id=entry_id,
            store_id=store_id,
            body=body,
            reply_all=reply_all,
            html=html,
            attachments=attachments,
            send_using_account=send_using_account,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_forward_mail",
        annotations=_write_annotations("Forward Outlook mail", external=True),
    )
    @safe_call
    async def outlook_forward_mail(
        entry_id: Annotated[str, Field(min_length=1)],
        to: Annotated[list[str], Field(min_length=1, description="Forward recipients.")],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by the read tool.")] = None,
        body: str = "",
        cc: Optional[list[str]] = None,
        html: bool = False,
        send_using_account: Annotated[Optional[str], Field(description="Optional exact Outlook SMTP address or account display name.")] = None,
    ) -> str:
        """Forward an existing message. This sends an external communication."""
        data = await bridge.call(
            multiwrite.forward_mail,
            entry_id=entry_id,
            store_id=store_id,
            to=to,
            body=body,
            cc=cc,
            html=html,
            send_using_account=send_using_account,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_move_mail",
        annotations=_write_annotations("Move Outlook mail"),
    )
    @safe_call
    async def outlook_move_mail(
        entry_id: Annotated[str, Field(min_length=1)],
        target_folder: Annotated[str, Field(description="Store-qualified destination path is recommended.")],
        store_id: Annotated[Optional[str], Field(description="Source StoreID returned by read/search tools.")] = None,
    ) -> str:
        """Move a message and return its new EntryID."""
        data = await bridge.call(
            multiwrite.move_mail,
            entry_id=entry_id,
            store_id=store_id,
            target_folder=target_folder,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_mail",
        annotations=_write_annotations("Delete Outlook mail", destructive=True),
    )
    @safe_call
    async def outlook_delete_mail(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by read/search tools.")] = None,
    ) -> str:
        """Delete a message (normally moves it to Deleted Items)."""
        return format_response(
            await bridge.call(multiwrite.delete_mail, entry_id=entry_id, store_id=store_id),
            "json",
        )

    @mcp.tool(
        name="outlook_mark_mail",
        annotations=_write_annotations("Mark or flag Outlook mail", idempotent=True),
    )
    @safe_call
    async def outlook_mark_mail(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by read/search tools.")] = None,
        read: Optional[bool] = None,
        flagged: Optional[bool] = None,
    ) -> str:
        """Change read/unread and/or follow-up flag state."""
        return format_response(
            await bridge.call(
                multiwrite.mark_mail,
                entry_id=entry_id,
                store_id=store_id,
                read=read,
                flagged=flagged,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_save_attachments",
        annotations=_write_annotations("Save Outlook attachments to disk"),
    )
    @safe_call
    async def outlook_save_attachments(
        entry_id: Annotated[str, Field(min_length=1)],
        output_dir: Annotated[str, Field(description="Absolute directory under the Windows user profile.")],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by the mail read tool.")] = None,
        attachment_index: Annotated[Optional[int], Field(ge=1)] = None,
    ) -> str:
        """Write one or all message attachments to the local filesystem."""
        return format_response(
            await bridge.call(
                multiwrite.save_attachments,
                entry_id=entry_id,
                store_id=store_id,
                output_dir=output_dir,
                attachment_index=attachment_index,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_create_folder",
        annotations=_write_annotations("Create Outlook folder"),
    )
    @safe_call
    async def outlook_create_folder(
        name: Annotated[str, Field(min_length=1)],
        parent: Annotated[str, Field(description="Store-qualified parent path is recommended.")] = "inbox",
    ) -> str:
        """Create an Outlook folder under the selected parent."""
        return format_response(
            await bridge.call(folders_client.create_folder, parent=parent, name=name),
            "json",
        )

    @mcp.tool(
        name="outlook_create_event",
        annotations=_write_annotations("Create Outlook calendar event", external=True),
    )
    @safe_call
    async def outlook_create_event(
        subject: str,
        start: Annotated[str, Field(description="ISO-8601 local date/time.")],
        end: Annotated[str, Field(description="ISO-8601 local date/time.")],
        folder: Annotated[str, Field(description="'calendar' or store-qualified calendar folder path.")] = "calendar",
        location: Optional[str] = None,
        body: Optional[str] = None,
        attendees: Annotated[Optional[list[str]], Field(description="If supplied, Outlook sends meeting invitations.")] = None,
        reminder_minutes: Annotated[Optional[int], Field(ge=0, le=10080)] = 15,
        recurrence: Optional[Recurrence] = None,
        send_using_account: Annotated[Optional[str], Field(description="Optional exact Outlook SMTP address or account display name.")] = None,
    ) -> str:
        """Create an event. If attendees are supplied this also sends invitations."""
        return format_response(
            await bridge.call(
                multiwrite.create_event,
                folder=folder,
                subject=subject,
                start=start,
                end=end,
                location=location,
                body=body,
                attendees=attendees,
                reminder_minutes=reminder_minutes,
                recurrence=recurrence,
                send_using_account=send_using_account,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_update_event",
        annotations=_write_annotations("Update Outlook calendar event", idempotent=True),
    )
    @safe_call
    async def outlook_update_event(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by list/get event.")] = None,
        subject: Optional[str] = None,
        start: Optional[str] = None,
        end: Optional[str] = None,
        location: Optional[str] = None,
        body: Optional[str] = None,
    ) -> str:
        """Update fields on an existing event."""
        return format_response(
            await bridge.call(
                multiwrite.update_event,
                entry_id=entry_id,
                store_id=store_id,
                subject=subject,
                start=start,
                end=end,
                location=location,
                body=body,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_delete_event",
        annotations=_write_annotations("Delete Outlook calendar event", destructive=True, external=True),
    )
    @safe_call
    async def outlook_delete_event(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by list/get event.")] = None,
    ) -> str:
        """Delete an event; Outlook may send cancellations for meetings."""
        return format_response(
            await bridge.call(multiwrite.delete_event, entry_id=entry_id, store_id=store_id),
            "json",
        )

    @mcp.tool(
        name="outlook_respond_event",
        annotations=_write_annotations("Respond to Outlook meeting", external=True),
    )
    @safe_call
    async def outlook_respond_event(
        entry_id: Annotated[str, Field(min_length=1)],
        response: Literal["accept", "tentative", "decline"],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by list/get event.")] = None,
        send_response: bool = True,
        send_using_account: Annotated[Optional[str], Field(description="Optional exact Outlook SMTP address or account display name.")] = None,
    ) -> str:
        """Accept, tentatively accept, or decline a meeting; optionally notify the organizer."""
        return format_response(
            await bridge.call(
                multiwrite.respond_event,
                entry_id=entry_id,
                store_id=store_id,
                response=response,
                send_response=send_response,
                send_using_account=send_using_account,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_create_task",
        annotations=_write_annotations("Create Outlook task"),
    )
    @safe_call
    async def outlook_create_task(
        subject: str,
        folder: Annotated[str, Field(description="'tasks' or store-qualified Tasks folder path.")] = "tasks",
        due_date: Optional[str] = None,
        body: Optional[str] = None,
        importance: Literal["low", "normal", "high"] = "normal",
        reminder: Optional[str] = None,
    ) -> str:
        """Create a task in the selected Outlook task store."""
        return format_response(
            await bridge.call(
                multiwrite.create_task,
                folder=folder,
                subject=subject,
                due_date=due_date,
                body=body,
                importance=importance,
                reminder=reminder,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_complete_task",
        annotations=_write_annotations("Complete Outlook task", idempotent=True),
    )
    @safe_call
    async def outlook_complete_task(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by list_tasks.")] = None,
    ) -> str:
        """Mark an Outlook task complete."""
        return format_response(
            await bridge.call(multiwrite.complete_task, entry_id=entry_id, store_id=store_id),
            "json",
        )

    @mcp.tool(
        name="outlook_set_category",
        annotations=_write_annotations("Set Outlook item categories", idempotent=True),
    )
    @safe_call
    async def outlook_set_category(
        entry_id: Annotated[str, Field(min_length=1)],
        categories: Annotated[str, Field(description="Comma-separated category names; empty string clears all.")],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by the item read tool.")] = None,
    ) -> str:
        """Replace categories on a mail, event, contact, or task."""
        return format_response(
            await bridge.call(
                multiwrite.set_category,
                entry_id=entry_id,
                store_id=store_id,
                categories=categories,
            ),
            "json",
        )

    @mcp.tool(
        name="outlook_toggle_rule",
        annotations=_write_annotations("Enable or disable Outlook mail rule", idempotent=True),
    )
    @safe_call
    async def outlook_toggle_rule(
        rule_name: Annotated[str, Field(min_length=1)],
        enabled: bool,
        store_ref: Annotated[Optional[str], Field(description="Optional StoreID or exact store display name; defaults to the default store.")] = None,
    ) -> str:
        """Enable or disable a live Outlook rule in the selected store."""
        return format_response(
            await bridge.call(
                multiwrite.toggle_rule,
                rule_name=rule_name,
                enabled=enabled,
                store_ref=store_ref,
            ),
            "json",
        )
