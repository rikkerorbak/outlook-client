import base64

import httpx
import pytest

from msgraph_client.functionality.mail import MailClient
from tests.conftest import json_body


def test_list_messages_inbox_with_filter_and_select(
    mail_client_with_recorder, sample_message
):
    mail, requests = mail_client_with_recorder(
        lambda r: httpx.Response(200, json={"value": [sample_message]})
    )

    messages = mail.list_messages(
        folder="inbox",
        top=10,
        filter_query="isRead eq false",
        select=["id", "subject"],
    )

    assert len(messages) == 1
    assert messages[0]["id"] == "AAMkAGI2"
    req = requests[0]
    assert req.method == "GET"
    assert "/v1.0/users/bot%40example.com/mailFolders/inbox/messages" in str(req.url)
    assert req.url.params["$top"] == "10"
    assert req.url.params["$filter"] == "isRead eq false"
    assert req.url.params["$select"] == "id,subject"
    assert req.url.params["$orderby"] == "receivedDateTime desc"


def test_list_messages_search_sets_consistency_level_and_omits_orderby(
    mail_client_with_recorder, sample_message
):
    mail, requests = mail_client_with_recorder(
        lambda r: httpx.Response(200, json={"value": [sample_message]})
    )

    mail.list_messages(search="invoice")

    req = requests[0]
    assert req.headers["ConsistencyLevel"] == "eventual"
    assert req.url.params["$search"] == '"invoice"'
    assert "$orderby" not in req.url.params


def test_list_messages_fetch_all_paginates(mail_client_with_recorder):
    def responder(request: httpx.Request) -> httpx.Response:
        if "page=2" in str(request.url):
            return httpx.Response(200, json={"value": [{"id": "m2"}]})
        return httpx.Response(
            200,
            json={
                "value": [{"id": "m1"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/bot%40example.com/messages?page=2",
            },
        )

    mail, _ = mail_client_with_recorder(responder)
    messages = mail.list_messages(folder=None, fetch_all=True)

    assert [m["id"] for m in messages] == ["m1", "m2"]


def test_get_message_returns_none_on_404(mail_client_with_recorder):
    mail, _ = mail_client_with_recorder(
        lambda r: httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )

    assert mail.get_message("missing-id") is None


def test_mark_as_read_sends_patch(mail_client_with_recorder):
    mail, requests = mail_client_with_recorder(lambda r: httpx.Response(200, json={}))

    assert mail.mark_as_read("msg-1", is_read=True) is True
    assert requests[0].method == "PATCH"
    assert json_body(requests[0]) == {"isRead": True}


def test_delete_message_sends_delete(mail_client_with_recorder):
    mail, requests = mail_client_with_recorder(lambda r: httpx.Response(204))

    assert mail.delete_message("msg-1") is True
    assert requests[0].method == "DELETE"


def test_send_mail_with_recipients_and_attachment(mail_client_with_recorder):
    mail, requests = mail_client_with_recorder(lambda r: httpx.Response(202))
    attachment = MailClient.build_file_attachment(
        name="hello.txt", content="hi there", content_type="text/plain"
    )

    sent = mail.send_mail(
        to=["alice@example.com", "bob@example.com"],
        cc="carol@example.com",
        subject="Greetings",
        body="<p>Hello</p>",
        body_type="HTML",
        attachments=[attachment],
        save_to_sent_items=False,
    )

    assert sent is True
    req = requests[0]
    assert req.method == "POST"
    assert str(req.url).endswith("/v1.0/users/bot%40example.com/sendMail")
    payload = json_body(req)
    assert payload["saveToSentItems"] is False
    msg = payload["message"]
    assert msg["subject"] == "Greetings"
    assert msg["body"] == {"contentType": "HTML", "content": "<p>Hello</p>"}
    assert msg["toRecipients"] == [
        {"emailAddress": {"address": "alice@example.com"}},
        {"emailAddress": {"address": "bob@example.com"}},
    ]
    assert msg["ccRecipients"] == [{"emailAddress": {"address": "carol@example.com"}}]
    assert msg["attachments"][0]["name"] == "hello.txt"
    assert (
        base64.b64decode(msg["attachments"][0]["contentBytes"]).decode("utf-8")
        == "hi there"
    )


def test_reply_and_draft_operations(mail_client_with_recorder):
    def responder(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/messages"):
            return httpx.Response(201, json={"id": "draft-123"})
        return httpx.Response(202)

    mail, requests = mail_client_with_recorder(responder)

    draft = mail.create_draft(to="a@example.com", subject="Draft", body="WIP")
    assert draft["id"] == "draft-123"
    assert mail.send_draft("draft-123") is True
    assert mail.reply_to_message("msg-9", "Thanks!", reply_all=True) is True

    assert requests[1].url.path.endswith("/messages/draft-123/send")
    assert requests[2].url.path.endswith("/messages/msg-9/replyAll")
    assert json_body(requests[2]) == {"comment": "Thanks!"}


def test_missing_user_raises_value_error(mock_client_factory):
    graph = mock_client_factory(lambda r: httpx.Response(200, json={}))
    mail = MailClient(graph)

    with pytest.raises(ValueError, match="A user"):
        mail.list_messages()
