import httpx
import pytest

from msgraph_client.client import GraphClient


def test_graph_client_requires_credentials():
    with pytest.raises(ValueError, match="Tenant ID"):
        GraphClient(tenant_id="", client_id="id", client_secret="secret")

    with pytest.raises(ValueError, match="Client ID and client secret"):
        GraphClient(tenant_id="tenant", client_id="", client_secret="secret")


def test_url_normalization_preserves_version_prefix(mock_client_factory):
    client = mock_client_factory(lambda r: httpx.Response(200, json={}))

    assert client._normalize_url("users/a@b.dk/messages") == (
        "https://graph.microsoft.com/v1.0/users/a@b.dk/messages"
    )
    assert client._normalize_url("/users/a@b.dk/messages") == (
        "https://graph.microsoft.com/v1.0/users/a@b.dk/messages"
    )
    assert client._normalize_url("https://graph.microsoft.com/v1.0/me") == (
        "https://graph.microsoft.com/v1.0/me"
    )


def test_get_paginated_follows_next_link(mock_client_factory):
    calls: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)
        calls.append(url)
        if "skiptoken=page2" in url:
            return httpx.Response(200, json={"value": [{"id": "2"}, {"id": "3"}]})
        return httpx.Response(
            200,
            json={
                "value": [{"id": "1"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?%24skiptoken=page2",
            },
        )

    client = mock_client_factory(handler)
    items = client.get_paginated("me/messages", params={"$top": 1})

    assert [item["id"] for item in items] == ["1", "2", "3"]
    assert len(calls) == 2


def test_manager_lazy_loading(manager):
    assert manager._graph_client is None
    assert manager._mail_client is None
    assert manager._default_user == "bot@example.com"
