import json
from collections.abc import Callable

import httpx
import pytest

from outlook_client.client import GraphClient
from outlook_client.functionality.mail import MailClient
from outlook_client.manager import OutlookClientManager


@pytest.fixture
def mock_client_factory():
    """Factory that builds a GraphClient backed by an httpx.MockTransport."""

    def _create(
        handler: Callable[[httpx.Request], httpx.Response],
    ) -> GraphClient:
        http_client = httpx.Client(transport=httpx.MockTransport(handler))
        return GraphClient(
            tenant_id="test-tenant",
            client_id="test-client-id",
            client_secret="test-secret",
            http_client=http_client,
        )

    return _create


@pytest.fixture
def sample_message() -> dict:
    return {
        "id": "AAMkAGI2",
        "subject": "Test Subject",
        "isRead": False,
        "body": {"contentType": "Text", "content": "Hello world"},
        "from": {"emailAddress": {"address": "sender@example.com"}},
    }


@pytest.fixture
def mail_client_with_recorder(mock_client_factory):
    """Returns a (MailClient, recorded_requests) tuple for asserting outgoing HTTP calls."""
    recorded: list[httpx.Request] = []

    def _create(
        responder: Callable[[httpx.Request], httpx.Response],
        default_user: str = "bot@example.com",
    ) -> tuple[MailClient, list[httpx.Request]]:
        def handler(request: httpx.Request) -> httpx.Response:
            recorded.append(request)
            return responder(request)

        graph = mock_client_factory(handler)
        return MailClient(graph, default_user=default_user), recorded

    return _create


@pytest.fixture
def manager() -> OutlookClientManager:
    return OutlookClientManager(
        tenant_id="test-tenant",
        client_id="test-client-id",
        client_secret="test-secret",
        default_user="bot@example.com",
    )


def json_body(request: httpx.Request) -> dict:
    return json.loads(request.content.decode("utf-8"))
