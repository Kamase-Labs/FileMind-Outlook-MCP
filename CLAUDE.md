# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Read-only Outlook MCP server providing email access via Microsoft Graph API. This is the Python implementation replacing the Node.js version, following the same patterns as FileMind-Procore-MCP.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run locally
python -m src.server

# Health check
curl http://localhost:8002/health

# Docker build and run
docker build -t outlook-mcp .
docker run -p 8002:8002 --env-file .env outlook-mcp
```

## Architecture

```
src/
├── server.py              # FastMCP server with middleware and all tools
├── config.py              # Pydantic settings (environment variables)
└── auth/
    └── token_service.py   # Microsoft token fetch/refresh from Supabase
```

### Key Components

- **OutlookAuthMiddleware**: Extracts `X-User-ID` header, fetches Microsoft token from Supabase, injects into context
- **TokenService**: Manages encrypted tokens in `oauth_connections` table with auto-refresh
- **Graph API helpers**: `graph_get()`, `graph_get_paginated()`, `resolve_folder()`

## Tools

| Tool | Description |
|------|-------------|
| `list_emails` | List emails from folder (inbox, sent, drafts, deleted, junk, archive, or custom) |
| `search_emails` | Search with KQL filters (query, from, subject, has_attachments, unread_only) |
| `read_email` | Read full email content by ID |

## Database

Uses shared `oauth_connections` table with `provider = 'microsoft'`.

```sql
SELECT id, user_id, access_token, refresh_token, expires_at, provider_metadata
FROM oauth_connections
WHERE user_id = $1 AND provider = 'microsoft' AND is_active = TRUE
```

Tokens are encrypted with Fernet at rest. The `ENCRYPTION_KEY` environment variable must match the key used by the auth service.

## Environment Variables

See `.env.example` for required configuration:

| Variable | Description |
|----------|-------------|
| `SUPABASE_URL` | PostgreSQL connection URL |
| `ENCRYPTION_KEY` | Fernet key for token decryption |
| `MICROSOFT_CLIENT_ID` | Azure AD app client ID |
| `MICROSOFT_CLIENT_SECRET` | Azure AD app client secret |
| `MICROSOFT_TENANT_ID` | Azure AD tenant (default: "common") |
| `TRUST_X_USER_ID` | Trust X-User-ID header (default: true) |
| `SERVER_PORT` | Server port (default: 8002) |
| `LOG_LEVEL` | Logging level (default: INFO) |

## Sidecar Pattern

This MCP runs as a sidecar container alongside FileMind-Agent in Cloud Run. The agent passes `X-User-ID` header with each request, and this server fetches the corresponding Microsoft token from Supabase.

## Testing

```bash
# Test with curl (requires valid X-User-ID)
curl -X POST http://localhost:8002/mcp \
  -H "Content-Type: application/json" \
  -H "X-User-ID: your-user-uuid" \
  -d '{"method": "tools/call", "params": {"name": "list_emails", "arguments": {"folder": "inbox", "count": 5}}}'
```
