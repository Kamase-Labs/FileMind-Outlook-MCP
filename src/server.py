"""
Outlook MCP Server - Read-only email access via Microsoft Graph API.

Architecture follows Procore MCP exactly:
- Middleware extracts X-User-ID → fetches Microsoft token
- Token injected into Context → tools access via ctx.get_state()
- All tools in one file (no unnecessary abstractions)
"""

import logging
import re
from contextlib import asynccontextmanager
from typing import Optional

import httpx
from fastmcp import FastMCP, Context
from fastmcp.server.middleware import Middleware, MiddlewareContext
from fastmcp.server.dependencies import get_http_headers
from fastmcp.exceptions import ToolError
from starlette.requests import Request
from starlette.responses import JSONResponse

from .config import get_settings
from .auth.token_service import TokenService, TokenServiceError

# =============================================================================
# Setup
# =============================================================================

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

settings = get_settings()
logging.getLogger().setLevel(settings.log_level)

# Initialize token service
token_service = TokenService(
    database_url=settings.supabase_url,
    encryption_key=settings.encryption_key,
    client_id=settings.microsoft_client_id,
    client_secret=settings.microsoft_client_secret,
    tenant_id=settings.microsoft_tenant_id
)

# =============================================================================
# Well-Known Folders (matches Node.js folder-utils.js)
# =============================================================================

WELL_KNOWN_FOLDERS = {
    "inbox": "me/mailFolders/inbox/messages",
    "drafts": "me/mailFolders/drafts/messages",
    "sent": "me/mailFolders/sentItems/messages",
    "deleted": "me/mailFolders/deletedItems/messages",
    "junk": "me/mailFolders/junkemail/messages",
    "archive": "me/mailFolders/archive/messages"
}

# =============================================================================
# Graph API Helper (inline, ~25 lines instead of separate file)
# =============================================================================

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


async def graph_get(ctx: Context, endpoint: str, **params) -> dict:
    """Make authenticated GET request to Microsoft Graph API."""
    access_token = ctx.get_state("microsoft_token")
    if not access_token:
        raise ToolError("Microsoft authentication required. Please connect your Outlook account.")

    # Handle full URLs (pagination nextLink)
    url = endpoint if endpoint.startswith("http") else f"{GRAPH_BASE_URL}/{endpoint}"

    async with httpx.AsyncClient() as client:
        response = await client.get(
            url,
            headers={"Authorization": f"Bearer {access_token}"},
            params=params if not endpoint.startswith("http") else None,
            timeout=30
        )
        if response.status_code == 401:
            raise ToolError("Session expired. Please reconnect your Outlook account.")
        if response.status_code >= 400:
            raise ToolError(f"Microsoft Graph API error: {response.status_code}")
        return response.json()


async def graph_get_paginated(ctx: Context, endpoint: str, max_count: int, **params) -> list:
    """Fetch paginated results up to max_count."""
    all_items = []
    current_endpoint = endpoint
    current_params = params

    while len(all_items) < max_count:
        response = await graph_get(ctx, current_endpoint, **current_params)
        items = response.get("value", [])
        all_items.extend(items)

        next_link = response.get("@odata.nextLink")
        if not next_link or len(all_items) >= max_count:
            break

        current_endpoint = next_link
        current_params = {}  # nextLink includes all params

    return all_items[:max_count]


async def resolve_folder(ctx: Context, folder_name: str) -> str:
    """Resolve folder name to Graph API endpoint."""
    if not folder_name:
        return WELL_KNOWN_FOLDERS["inbox"]

    lower = folder_name.lower()
    if lower in WELL_KNOWN_FOLDERS:
        return WELL_KNOWN_FOLDERS[lower]

    # Try to find custom folder by displayName
    try:
        response = await graph_get(ctx, "me/mailFolders", **{"$filter": f"displayName eq '{folder_name}'"})
        if response.get("value"):
            folder_id = response["value"][0]["id"]
            return f"me/mailFolders/{folder_id}/messages"
    except Exception:
        pass

    # Fallback to inbox
    logger.warning(f"Folder '{folder_name}' not found, using inbox")
    return WELL_KNOWN_FOLDERS["inbox"]


# =============================================================================
# Formatting Helpers (matches Node.js read.js)
# =============================================================================

def strip_html(html: str) -> str:
    """Convert HTML to plain text (matches Node.js regex)."""
    return re.sub(r'<[^>]*>', '', html) if html else ""


def format_email_address(email_obj: dict) -> str:
    """Format email address object to 'Name (address)' string."""
    if not email_obj:
        return "Unknown"
    addr = email_obj.get("emailAddress", email_obj)
    name = addr.get("name", "")
    address = addr.get("address", "")
    return f"{name} ({address})" if name else address


def format_recipients(recipients: list) -> str:
    """Format list of recipients to comma-separated string."""
    if not recipients:
        return "None"
    return ", ".join(format_email_address(r) for r in recipients)


# =============================================================================
# Middleware (matches Procore MCP user_context_middleware.py)
# =============================================================================

class OutlookAuthMiddleware(Middleware):
    """Extracts user identity and injects Microsoft token into context."""

    async def on_call_tool(self, context: MiddlewareContext, call_next):
        ctx = context.fastmcp_context

        try:
            # Extract user_id from X-User-ID header (sidecar pattern)
            headers = get_http_headers()
            user_id = headers.get("x-user-id") or headers.get("X-User-ID")

            if not user_id:
                raise ToolError("Authentication required. No X-User-ID header found.")

            # Fetch Microsoft token from Supabase
            try:
                token = await token_service.get_token(user_id)
            except TokenServiceError as e:
                raise ToolError(e.message)

            # Inject into context
            ctx.set_state("user_id", user_id)
            ctx.set_state("microsoft_token", token.access_token)

        except ToolError:
            raise
        except Exception as e:
            logger.error(f"Auth middleware error: {e}")
            raise ToolError(f"Authentication failed: {str(e)}")

        return await call_next(context)


# =============================================================================
# FastMCP Server
# =============================================================================

@asynccontextmanager
async def lifespan(app: FastMCP):
    logger.info("Outlook MCP Server starting...")
    yield
    logger.info("Shutting down...")
    await token_service.close()


mcp = FastMCP(name="Outlook Integration", lifespan=lifespan)
mcp.add_middleware(OutlookAuthMiddleware())


# =============================================================================
# Health Check
# =============================================================================

@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    return JSONResponse({"status": "healthy", "service": "outlook-mcp", "version": "1.0.0"})


# =============================================================================
# Email Tools
# =============================================================================

@mcp.tool
async def list_emails(
    ctx: Context,
    folder: str = "inbox",
    count: int = 10
) -> str:
    """
    List emails from a folder.

    Args:
        folder: Folder name (inbox, sent, drafts, deleted, junk, archive, or custom folder name)
        count: Number of emails to retrieve (max 50)
    """
    count = min(50, max(1, count))
    endpoint = await resolve_folder(ctx, folder)

    emails = await graph_get_paginated(
        ctx, endpoint, count,
        **{
            "$top": min(50, count),
            "$orderby": "receivedDateTime desc",
            "$select": settings.email_list_fields
        }
    )

    if not emails:
        return f"No emails found in {folder}."

    # Format output (matches Node.js list.js)
    lines = [f"Found {len(emails)} emails in {folder}:\n"]
    for i, email in enumerate(emails, 1):
        sender = format_email_address(email.get("from"))
        date = email.get("receivedDateTime", "")[:19].replace("T", " ")
        unread = "[UNREAD] " if not email.get("isRead") else ""
        subject = email.get("subject", "(no subject)")
        email_id = email.get("id", "")

        lines.append(f"{i}. {unread}{date} - From: {sender}")
        lines.append(f"   Subject: {subject}")
        lines.append(f"   ID: {email_id}\n")

    return "\n".join(lines)


@mcp.tool
async def search_emails(
    ctx: Context,
    query: Optional[str] = None,
    folder: str = "inbox",
    from_address: Optional[str] = None,
    subject: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    unread_only: Optional[bool] = None,
    count: int = 10
) -> str:
    """
    Search emails with filters.

    Args:
        query: General search text
        folder: Folder to search
        from_address: Filter by sender email/name
        subject: Filter by subject
        has_attachments: Only emails with attachments
        unread_only: Only unread emails
        count: Max results (max 50)
    """
    count = min(50, max(1, count))
    endpoint = await resolve_folder(ctx, folder)

    # Build search params (progressive strategy from Node.js search.js)
    params = {
        "$top": count,
        "$orderby": "receivedDateTime desc",
        "$select": settings.email_list_fields
    }

    # Build KQL $search string
    search_terms = []
    if query:
        search_terms.append(f'"{query}"')
    if subject:
        search_terms.append(f'subject:"{subject}"')
    if from_address:
        search_terms.append(f'from:"{from_address}"')

    if search_terms:
        params["$search"] = " ".join(search_terms)

    # Build $filter for boolean conditions
    filters = []
    if has_attachments is True:
        filters.append("hasAttachments eq true")
    if unread_only is True:
        filters.append("isRead eq false")

    if filters:
        params["$filter"] = " and ".join(filters)

    # Try combined search first
    try:
        emails = await graph_get_paginated(ctx, endpoint, count, **params)
        if emails:
            return _format_search_results(emails, "combined search")
    except Exception as e:
        logger.warning(f"Combined search failed: {e}")

    # Fallback: try individual terms
    for term_name, term_value in [("subject", subject), ("from", from_address), ("query", query)]:
        if term_value:
            try:
                fallback_params = {
                    "$top": count,
                    "$orderby": "receivedDateTime desc",
                    "$select": settings.email_list_fields,
                    "$search": f'{term_name}:"{term_value}"' if term_name != "query" else f'"{term_value}"'
                }
                emails = await graph_get_paginated(ctx, endpoint, count, **fallback_params)
                if emails:
                    return _format_search_results(emails, f"{term_name} search")
            except Exception:
                continue

    # Final fallback: recent emails
    fallback_params = {
        "$top": count,
        "$orderby": "receivedDateTime desc",
        "$select": settings.email_list_fields
    }
    emails = await graph_get_paginated(ctx, endpoint, count, **fallback_params)
    return _format_search_results(emails, "recent emails fallback")


def _format_search_results(emails: list, strategy: str) -> str:
    """Format search results."""
    if not emails:
        return "No emails found matching your search criteria."

    lines = [f"Found {len(emails)} emails (via {strategy}):\n"]
    for i, email in enumerate(emails, 1):
        sender = format_email_address(email.get("from"))
        date = email.get("receivedDateTime", "")[:19].replace("T", " ")
        unread = "[UNREAD] " if not email.get("isRead") else ""
        subject = email.get("subject", "(no subject)")
        email_id = email.get("id", "")

        lines.append(f"{i}. {unread}{date} - From: {sender}")
        lines.append(f"   Subject: {subject}")
        lines.append(f"   ID: {email_id}\n")

    return "\n".join(lines)


@mcp.tool
async def read_email(ctx: Context, email_id: str) -> str:
    """
    Read full email content by ID.

    Args:
        email_id: Email ID from list_emails or search_emails results
    """
    if not email_id:
        raise ToolError("Email ID is required.")

    try:
        email = await graph_get(
            ctx,
            f"me/messages/{email_id}",
            **{"$select": settings.email_detail_fields}
        )
    except ToolError:
        raise
    except Exception as e:
        if "doesn't belong" in str(e).lower():
            raise ToolError("Invalid email ID or email not found in your mailbox.")
        raise ToolError(f"Failed to read email: {str(e)}")

    # Format output (matches Node.js read.js)
    sender = format_email_address(email.get("from"))
    to = format_recipients(email.get("toRecipients", []))
    cc = format_recipients(email.get("ccRecipients", []))
    bcc = format_recipients(email.get("bccRecipients", []))
    date = email.get("receivedDateTime", "")[:19].replace("T", " ")
    subject = email.get("subject", "(no subject)")
    importance = email.get("importance", "normal")
    has_attachments = "Yes" if email.get("hasAttachments") else "No"

    # Extract body
    body_obj = email.get("body", {})
    if body_obj.get("contentType") == "html":
        body = strip_html(body_obj.get("content", ""))
    else:
        body = body_obj.get("content", email.get("bodyPreview", ""))

    lines = [
        f"From: {sender}",
        f"To: {to}",
    ]
    if cc != "None":
        lines.append(f"CC: {cc}")
    if bcc != "None":
        lines.append(f"BCC: {bcc}")
    lines.extend([
        f"Subject: {subject}",
        f"Date: {date}",
        f"Importance: {importance}",
        f"Has Attachments: {has_attachments}",
        "",
        body
    ])

    return "\n".join(lines)


# =============================================================================
# Export ASGI App
# =============================================================================

app = mcp.http_app()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "src.server:app",
        host=settings.server_host,
        port=settings.server_port,
        reload=True,
        log_level=settings.log_level.lower()
    )
