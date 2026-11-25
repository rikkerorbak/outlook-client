"""
Microsoft Graph Outlook Python Client.

Recommended usage:
    from outlook_client import OutlookClientManager

    outlook = OutlookClientManager(
        tenant_id="...",
        client_id="...",
        client_secret="...",
        default_user="user@example.com",
    )
    messages = outlook.mail.list_messages()
"""

from . import hooks
from .client import GraphClient
from .functionality.mail import MailClient
from .manager import OutlookClientManager

__all__ = [
    "GraphClient",
    "MailClient",
    "OutlookClientManager",
    "hooks",
]
