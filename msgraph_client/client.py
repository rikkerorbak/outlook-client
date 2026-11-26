import logging
from urllib.parse import urljoin

import httpx
from authlib.integrations.httpx_client import OAuth2Client

from .hooks import create_response_logging_hook


class GraphClient:
    """
    Base client for Microsoft Graph API communication.

    Recommended usage is via MSGraphClientManager rather than direct instantiation.
    """

    DEFAULT_BASE_URL = "https://graph.microsoft.com/v1.0/"
    DEFAULT_SCOPE = "https://graph.microsoft.com/.default"

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        scope: str = DEFAULT_SCOPE,
        base_url: str = DEFAULT_BASE_URL,
        timeout: float = 30.0,
        auto_fetch_token: bool = True,
        http_client: httpx.Client | None = None,
    ):
        """
        Initialize the GraphClient with Azure AD / Entra ID client credentials.

        :param tenant_id: Azure AD / Entra ID tenant ID.
        :param client_id: The OAuth2 application (client) ID.
        :param client_secret: The OAuth2 client secret.
        :param scope: OAuth2 scope (default: https://graph.microsoft.com/.default).
        :param base_url: Microsoft Graph API base URL (default: v1.0).
        :param timeout: Request timeout in seconds (default: 30.0).
        :param auto_fetch_token: Whether to fetch an access token during init.
        :param http_client: Optional pre-configured httpx.Client (useful for testing).
        """
        if not tenant_id:
            raise ValueError("Tenant ID must be provided.")
        if not client_id or not client_secret:
            raise ValueError("Client ID and client secret must be provided.")

        self.tenant_id = tenant_id
        self.token_url = (
            f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        )
        self.base_url = base_url if base_url.endswith("/") else f"{base_url}/"
        self.scope = scope

        self.logger = logging.getLogger("msgraph_client.graph")
        logging.getLogger("httpx").setLevel(logging.WARNING)
        logging.getLogger("httpcore").setLevel(logging.WARNING)

        response_hook = create_response_logging_hook(logger=self.logger)
        hooks = {"response": [response_hook]}

        if http_client is not None:
            self.client = http_client
        else:
            self.client = OAuth2Client(
                client_id=client_id,
                client_secret=client_secret,
                token_endpoint=self.token_url,
                scope=self.scope,
                timeout=timeout,
                event_hooks=hooks,
            )
            if auto_fetch_token:
                self.client.fetch_token()

    def _normalize_url(self, endpoint: str) -> str:
        """Ensure the URL is absolute, preserving the Graph API version prefix."""
        if endpoint.startswith(("http://", "https://")):
            return endpoint
        return urljoin(self.base_url, endpoint.lstrip("/"))

    def get(self, endpoint: str, **kwargs) -> httpx.Response:
        """Perform GET request to the specified endpoint."""
        url = self._normalize_url(endpoint)
        response = self.client.get(url, **kwargs)
        response.raise_for_status()
        return response

    def post(
        self, endpoint: str, json: dict | None = None, **kwargs
    ) -> httpx.Response:
        """Perform POST request to the specified endpoint."""
        url = self._normalize_url(endpoint)
        response = self.client.post(url, json=json, **kwargs)
        response.raise_for_status()
        return response

    def patch(
        self, endpoint: str, json: dict | None = None, **kwargs
    ) -> httpx.Response:
        """Perform PATCH request to the specified endpoint."""
        url = self._normalize_url(endpoint)
        response = self.client.patch(url, json=json, **kwargs)
        response.raise_for_status()
        return response

    def put(
        self, endpoint: str, json: dict | None = None, **kwargs
    ) -> httpx.Response:
        """Perform PUT request to the specified endpoint."""
        url = self._normalize_url(endpoint)
        response = self.client.put(url, json=json, **kwargs)
        response.raise_for_status()
        return response

    def delete(self, endpoint: str, **kwargs) -> httpx.Response:
        """Perform DELETE request to the specified endpoint."""
        url = self._normalize_url(endpoint)
        response = self.client.delete(url, **kwargs)
        response.raise_for_status()
        return response

    def get_paginated(
        self,
        endpoint: str,
        params: dict | None = None,
        headers: dict | None = None,
        max_pages: int | None = None,
    ) -> list[dict]:
        """
        Fetch all pages of a Microsoft Graph collection endpoint.

        Follows `@odata.nextLink` until all items are retrieved or `max_pages` is reached.
        """
        items: list[dict] = []
        next_url: str | None = endpoint
        current_params = params
        pages_fetched = 0

        while next_url:
            response = self.get(next_url, params=current_params, headers=headers)
            data = response.json()
            items.extend(data.get("value", []))
            pages_fetched += 1

            if max_pages is not None and pages_fetched >= max_pages:
                break

            next_url = data.get("@odata.nextLink")
            # @odata.nextLink already includes query parameters
            current_params = None

        return items
