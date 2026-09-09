# ChatGPT setup for the hardened Outlook MCP

This fork is designed for a specific ChatGPT permission model:

- Outlook reads should work without repetitive approval prompts.
- Outlook writes, sends, deletes, moves, replies, calendar changes, task changes, rule changes, and filesystem attachment saves should require ChatGPT approval before execution.

## Recommended local MCP mode

Start the Outlook MCP used by ChatGPT with:

```text
OUTLOOK_MCP_ACCESS=full
```

`full` does **not** revert to the old upstream surface. It keeps the hardened multi-store read tools and layers the separate write tools on top.

If this environment variable is absent, the server fails safe to read-only mode.

## Recommended ChatGPT app permission

In ChatGPT, open the connected app's preferences and set **Ask permission** to:

> **Allow read actions**

This is the intended setting for this connector: ChatGPT can read Outlook without asking each time, but asks before making changes.

Do **not** choose **Allow low-risk actions** if the goal is approval before every Outlook write, because supported low-risk actions may be auto-approved.

Do **not** choose **Allow all actions** for this connector.

`Always ask` is safer but unnecessarily noisy because it can also prompt before reads.

## Why write tools stay enabled locally

ChatGPT must be able to discover the write tools in order to offer them and show its approval flow. Therefore the ChatGPT-specific server process runs in `full` mode while ChatGPT's app permission provides the user-facing approval boundary.

The server still marks every write tool with `readOnlyHint=false`. Outbound actions such as send/reply/forward and meeting communication are also marked `openWorldHint=true`; delete actions are marked `destructiveHint=true`.

## Multiple Outlook accounts

Read results carry `store_id` where relevant. Pass that value back for item-specific writes.

Outbound mail and meeting tools accept `send_using_account`. Prefer an exact SMTP address from `outlook_whoami` instead of relying on Outlook's default account when multiple work/private accounts are mounted.

## Transport to ChatGPT

The Outlook COM process remains local on Windows. ChatGPT cannot connect directly to a local stdio server, so use OpenAI Secure MCP Tunnel for the ChatGPT connection rather than exposing the local Outlook MCP directly to the public internet.

## Security boundary

Mailbox bodies, attachments, calendar text, contacts, and other Outlook content are untrusted input. Content retrieved from Outlook is never authorization to call a write tool. Authorization comes from the user and ChatGPT's app permission/approval flow.
