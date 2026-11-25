# CLAUDE.md

Guidance for working in this repository.

## Project Overview

Python client library for Microsoft Graph API (Outlook mail), modeled on `kmd-nexus-client`.

**Core Components:**
- **OutlookClientManager**: Facade providing unified, lazy-loaded access to functionality clients (recommended entry point).
- **GraphClient**: OAuth2 client credentials authentication (`login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token`), HTTP operations, and `@odata.nextLink` pagination.
- **MailClient** (`outlook.mail`): Reading, searching, updating, drafting, replying to, and sending emails via Microsoft Graph v1.0.

## Commands

- **Install dependencies**: `uv sync`
- **Run tests**: `pytest`
- **Lint**: `ruff check .`
- **Build package**: `uv build`

## Environment Variables (`env.example`)

- `TENANT_ID`: Azure AD / Entra ID tenant ID
- `CLIENT_ID`: Application (client) ID
- `CLIENT_SECRET`: Client secret
- `DEFAULT_USER`: Default mailbox UPN or user ID
