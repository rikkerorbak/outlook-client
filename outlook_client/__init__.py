"""
Microsoft Graph Outlook Python Client.

Recommended usage:
    from outlook_client import MSGraphClientManager

    graph = MSGraphClientManager(
        tenant_id="...",
        client_id="...",
        client_secret="...",
        default_user="user@example.com",
    )
    messages = graph.mail.list_messages()
"""

from . import hooks
from .client import GraphClient
from .functionality.mail import MailClient
from .manager import MSGraphClientManager

__all__ = [
    "GraphClient",
    "MSGraphClientManager",
    "MailClient",
    "hooks",
]
