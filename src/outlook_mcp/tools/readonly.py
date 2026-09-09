"""Hardened read-only MCP tool surface.

No tool registered here mutates Outlook, sends messages, saves attachments,
or writes to the local filesystem. This is the default server mode.
"""

from __future__ import annotations

from typing import Annotated, Literal, Optional

from pydantic import Field

from outlook_mcp.client import account as account_client
from outlook_mcp.client import categories as categories_client
from outlook_mcp.client import contacts as contacts_client
from outlook_mcp.client import multistore
from outlook_mcp.client import ooo as ooo_client
from outlook_mcp.client import rules as rules_client
from outlook_mcp.config import access_mode
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def _read_annotations(title: str) -> dict[str, object]:
    return {
        "title": title,
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }


def register(mcp, bridge) -> None:
    @mcp.tool(name="outlook_server_status", annotations=_read_annotations("Show Outlook MCP security mode"))
    async def outlook_server_status() -> str:
        """Show the current access mode. Read mode is the secure default."""
        return format_response(
            {
                "access_mode": access_mode(),
                "writes_registered": False,
                "note": "Read-only tools only. Set OUTLOOK_MCP_ACCESS=full before startup to enable legacy write tools.",
            },
            "json",
        )

    @mcp.tool(name="outlook_whoami", annotations=_read_annotations("Show current Outlook user and accounts"))
    @safe_call
    async def outlook_whoami(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        data = await bridge.call(account_client.whoami)
        data["access_mode"] = access_mode()
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_stores", annotations=_read_annotations("List all mounted Outlook stores"))
    @safe_call
    async def outlook_list_stores(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        """List every mounted mailbox, PST, and store in the active Outlook profile."""
        data = await bridge.call(multistore.list_stores)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_folders", annotations=_read_annotations("List Outlook folders across all stores"))
    @safe_call
    async def outlook_list_folders(
        root: Annotated[Optional[str], Field(description="Optional store-qualified folder path.")] = None,
        max_depth: Annotated[int, Field(ge=1, le=10)] = 4,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        """List folder trees. With no root, returns folders from every mounted store."""
        items = await bridge.call(multistore.list_folders, root=root, max_depth=max_depth)
        return format_response({"count": len(items), "items": items}, response_format)

    @mcp.tool(name="outlook_list_mails", annotations=_read_annotations("List Outlook mail"))
    @safe_call
    async def outlook_list_mails(
        folder: Annotated[str, Field(description="Well-known or store-qualified folder path.")] = "inbox",
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        offset: Annotated[int, Field(ge=0)] = 0,
        unread_only: bool = False,
        since: Optional[str] = None,
        until: Optional[str] = None,
        from_address: Optional[str] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(
            multistore.list_mails,
            folder=folder,
            limit=limit,
            offset=offset,
            unread_only=unread_only,
            since=since,
            until=until,
            from_address=from_address,
        )
        return format_response(data, response_format)

    @mcp.tool(name="outlook_search_mails", annotations=_read_annotations("Search Outlook mail"))
    @safe_call
    async def outlook_search_mails(
        query: Annotated[str, Field(min_length=1)],
        folder: Annotated[str, Field(description="Well-known or store-qualified folder path.")] = "inbox",
        scope: Literal["subject_body", "subject", "from"] = "subject_body",
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        """Search mail without exposing the upstream raw-DASL escape hatch."""
        data = await bridge.call(
            multistore.search_mails,
            query=query,
            folder=folder,
            scope=scope,
            limit=limit,
        )
        return format_response(data, response_format)

    @mcp.tool(name="outlook_get_mail", annotations=_read_annotations("Read one Outlook mail"))
    @safe_call
    async def outlook_get_mail(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by list/search; recommended for non-default stores.")] = None,
        include_body: bool = True,
        include_html: bool = False,
        max_body_chars: Annotated[int, Field(ge=0, le=100000)] = 10000,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(
            multistore.get_mail,
            entry_id=entry_id,
            store_id=store_id,
            include_body=include_body,
            include_html=include_html,
            max_body_chars=max_body_chars,
        )
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_events", annotations=_read_annotations("List Outlook calendar events"))
    @safe_call
    async def outlook_list_events(
        folder: Annotated[str, Field(description="'calendar' or a store-qualified calendar folder path.")] = "calendar",
        start: Optional[str] = None,
        end: Optional[str] = None,
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        include_recurrences: bool = True,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(
            multistore.list_events,
            folder=folder,
            start=start,
            end=end,
            limit=limit,
            include_recurrences=include_recurrences,
        )
        return format_response(data, response_format)

    @mcp.tool(name="outlook_get_event", annotations=_read_annotations("Read one Outlook calendar event"))
    @safe_call
    async def outlook_get_event(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by list_events.")] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(multistore.get_event, entry_id=entry_id, store_id=store_id)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_contacts", annotations=_read_annotations("List Outlook contacts across stores"))
    @safe_call
    async def outlook_list_contacts(
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        offset: Annotated[int, Field(ge=0)] = 0,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(multistore.list_contacts, limit=limit, offset=offset)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_search_contacts", annotations=_read_annotations("Search Outlook contacts and directory"))
    @safe_call
    async def outlook_search_contacts(
        query: Annotated[str, Field(min_length=1)],
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        include_directory: bool = True,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(
            multistore.search_contacts,
            query=query,
            limit=limit,
            include_directory=include_directory,
        )
        return format_response(data, response_format)

    @mcp.tool(name="outlook_get_contact", annotations=_read_annotations("Read one Outlook contact"))
    @safe_call
    async def outlook_get_contact(
        entry_id: Annotated[str, Field(min_length=1)],
        store_id: Annotated[Optional[str], Field(description="StoreID returned by contact list/search.")] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(multistore.get_contact, entry_id=entry_id, store_id=store_id)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_resolve_name", annotations=_read_annotations("Resolve an Outlook name or email address"))
    @safe_call
    async def outlook_resolve_name(
        name: Annotated[str, Field(min_length=1)],
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(contacts_client.resolve_name, name=name)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_tasks", annotations=_read_annotations("List Outlook tasks"))
    @safe_call
    async def outlook_list_tasks(
        folder: Annotated[str, Field(description="'tasks' or a store-qualified task folder path.")] = "tasks",
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        include_completed: bool = False,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(
            multistore.list_tasks,
            folder=folder,
            limit=limit,
            include_completed=include_completed,
        )
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_categories", annotations=_read_annotations("List Outlook categories"))
    @safe_call
    async def outlook_list_categories(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(categories_client.list_categories)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_list_rules", annotations=_read_annotations("List Outlook mail rules"))
    @safe_call
    async def outlook_list_rules(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(rules_client.list_rules)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_get_out_of_office", annotations=_read_annotations("Check Outlook Out-of-Office status"))
    @safe_call
    async def outlook_get_out_of_office(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "json",
    ) -> str:
        data = await bridge.call(ooo_client.get_out_of_office)
        return format_response(data, response_format)
