"""
MSGraphClientManager - A facade/factory for Microsoft Graph functionality clients.

Provides a single entry point with lazy-loaded properties for each functionality area.
"""


from msgraph_client.client import GraphClient
from msgraph_client.functionality.mail import MailClient


class MSGraphClientManager:
    """
    Manager for unified access to Microsoft Graph functionality clients.

    Example:
        graph = MSGraphClientManager(
            tenant_id="...",
            client_id="...",
            client_secret="...",
            default_user="user@example.com",
        )
        messages = graph.mail.list_messages()
        graph.mail.send_mail(
            to="recipient@example.com",
            subject="Hello",
            body="Sent via Microsoft Graph",
        )
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        default_user: str | None = None,
        timeout: float = 30.0,
    ):
        """
        Initialize the MSGraphClientManager.

        Args:
            tenant_id: Azure AD / Entra ID tenant ID.
            client_id: The OAuth2 application (client) ID.
            client_secret: The OAuth2 client secret.
            default_user: Optional default mailbox UPN or user ID (e.g. "user@domain.com").
            timeout: Request timeout in seconds (default: 30.0).
        """
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret
        self._default_user = default_user
        self._config = {"timeout": timeout}

        self._graph_client: GraphClient | None = None
        self._mail_client: MailClient | None = None

    @property
    def graph_client(self) -> GraphClient:
        """Get the base GraphClient (lazy-loaded)."""
        if self._graph_client is None:
            self._graph_client = GraphClient(
                tenant_id=self._tenant_id,
                client_id=self._client_id,
                client_secret=self._client_secret,
                **self._config,
            )
        return self._graph_client

    @property
    def mail(self) -> MailClient:
        """Get the MailClient (lazy-loaded)."""
        if self._mail_client is None:
            self._mail_client = MailClient(
                self.graph_client, default_user=self._default_user
            )
        return self._mail_client
