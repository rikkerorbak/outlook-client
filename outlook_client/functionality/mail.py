import base64
from collections.abc import Sequence
from urllib.parse import quote

from httpx import HTTPStatusError

from outlook_client.client import GraphClient


class MailClient:
    """
    Client for Microsoft Graph mail operations (reading and sending emails).

    Recommended usage is via OutlookClientManager:
        outlook.mail.list_messages(...)
        outlook.mail.send_mail(...)
    """

    def __init__(self, graph_client: GraphClient, default_user: str | None = None):
        """
        Initialize MailClient with a GraphClient instance.

        Args:
            graph_client: Authenticated GraphClient for API calls.
            default_user: Optional default mailbox UPN or user ID.
        """
        self.client = graph_client
        self.default_user = default_user

    def _user_prefix(self, user: str | None) -> str:
        """Resolve the Microsoft Graph user path prefix."""
        resolved_user = user or self.default_user
        if not resolved_user:
            raise ValueError(
                "A user (email/UPN or user ID) must be provided either on "
                "OutlookClientManager(default_user=...) or in the method call."
            )
        if resolved_user.lower() == "me":
            return "me"
        return f"users/{quote(resolved_user, safe='')}"

    @staticmethod
    def _format_recipients(
        recipients: str | Sequence[str] | None,
    ) -> list[dict]:
        """Convert a single email or sequence of emails into Graph recipient objects."""
        if not recipients:
            return []
        if isinstance(recipients, str):
            addresses = [recipients]
        else:
            addresses = list(recipients)
        return [{"emailAddress": {"address": addr}} for addr in addresses]

    @staticmethod
    def build_file_attachment(
        name: str,
        content: bytes | str,
        content_type: str = "application/octet-stream",
    ) -> dict:
        """
        Build a Microsoft Graph fileAttachment dictionary.

        Args:
            name: Filename shown on the attachment.
            content: Raw bytes or UTF-8 string content.
            content_type: MIME content type.

        Returns:
            Dictionary formatted for Microsoft Graph `attachments` list.
        """
        raw_bytes = content.encode("utf-8") if isinstance(content, str) else content
        encoded = base64.b64encode(raw_bytes).decode("ascii")
        return {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": content_type,
            "contentBytes": encoded,
        }

    def _build_message_payload(
        self,
        to: str | Sequence[str],
        subject: str,
        body: str,
        body_type: str = "Text",
        cc: str | Sequence[str] | None = None,
        bcc: str | Sequence[str] | None = None,
        reply_to: str | Sequence[str] | None = None,
        attachments: list[dict] | None = None,
    ) -> dict:
        """Build a Microsoft Graph message resource dictionary."""
        normalized_type = "HTML" if body_type.lower() == "html" else "Text"
        if body_type.lower() not in ("text", "html"):
            raise ValueError("body_type must be 'Text' or 'HTML'")

        to_recipients = self._format_recipients(to)
        if not to_recipients:
            raise ValueError("At least one 'to' recipient must be provided.")

        message: dict = {
            "subject": subject,
            "body": {
                "contentType": normalized_type,
                "content": body,
            },
            "toRecipients": to_recipients,
        }

        if cc_recipients := self._format_recipients(cc):
            message["ccRecipients"] = cc_recipients
        if bcc_recipients := self._format_recipients(bcc):
            message["bccRecipients"] = bcc_recipients
        if reply_to_recipients := self._format_recipients(reply_to):
            message["replyTo"] = reply_to_recipients
        if attachments:
            message["attachments"] = attachments

        return message

    # ------------------------------------------------------------------
    # Reading emails
    # ------------------------------------------------------------------

    def list_messages(
        self,
        user: str | None = None,
        folder: str | None = "inbox",
        top: int = 25,
        filter_query: str | None = None,
        search: str | None = None,
        select: Sequence[str] | None = None,
        order_by: str | None = "receivedDateTime desc",
        fetch_all: bool = False,
    ) -> list[dict]:
        """
        List email messages from a user's mailbox or mail folder.

        Args:
            user: User UPN/ID (defaults to manager's default_user).
            folder: Well-known folder name (e.g. "inbox", "sentitems") or folder ID.
                Pass None to query all messages across folders.
            top: Number of messages per page (default: 25).
            filter_query: OData $filter expression (e.g. "isRead eq false").
            search: OData $search query string.
            select: Optional list of message properties to return.
            order_by: OData $orderby expression (ignored when `search` is used).
            fetch_all: If True, follow `@odata.nextLink` to return all matching messages.

        Returns:
            List of message dictionaries.
        """
        prefix = self._user_prefix(user)
        endpoint = (
            f"{prefix}/mailFolders/{quote(folder, safe='')}/messages"
            if folder
            else f"{prefix}/messages"
        )

        params: dict[str, str | int] = {"$top": top}
        headers: dict[str, str] | None = None

        if filter_query:
            params["$filter"] = filter_query
        if select:
            params["$select"] = ",".join(select)
        if search:
            quoted_search = (
                search
                if (search.startswith('"') and search.endswith('"'))
                else f'"{search}"'
            )
            params["$search"] = quoted_search
            headers = {"ConsistencyLevel": "eventual"}
        elif order_by:
            params["$orderby"] = order_by

        if fetch_all:
            return self.client.get_paginated(endpoint, params=params, headers=headers)

        response = self.client.get(endpoint, params=params, headers=headers)
        return response.json().get("value", [])

    def get_message(
        self,
        message_id: str,
        user: str | None = None,
        select: Sequence[str] | None = None,
    ) -> dict | None:
        """
        Fetch a single email message by ID.

        Returns:
            Message dictionary, or None if not found (404).
        """
        prefix = self._user_prefix(user)
        params = {"$select": ",".join(select)} if select else None

        try:
            response = self.client.get(
                f"{prefix}/messages/{quote(message_id, safe='')}", params=params
            )
            return response.json()
        except HTTPStatusError as e:
            if e.response.status_code == 404:
                return None
            raise

    def list_attachments(
        self, message_id: str, user: str | None = None
    ) -> list[dict]:
        """List all attachments for a given message."""
        prefix = self._user_prefix(user)
        return self.client.get_paginated(
            f"{prefix}/messages/{quote(message_id, safe='')}/attachments"
        )

    def mark_as_read(
        self, message_id: str, is_read: bool = True, user: str | None = None
    ) -> bool:
        """Mark a message as read or unread."""
        prefix = self._user_prefix(user)
        response = self.client.patch(
            f"{prefix}/messages/{quote(message_id, safe='')}",
            json={"isRead": is_read},
        )
        return response.status_code == 200

    def move_message(
        self,
        message_id: str,
        destination_folder_id: str,
        user: str | None = None,
    ) -> dict:
        """Move a message to another mail folder."""
        prefix = self._user_prefix(user)
        response = self.client.post(
            f"{prefix}/messages/{quote(message_id, safe='')}/move",
            json={"destinationId": destination_folder_id},
        )
        return response.json()

    def delete_message(self, message_id: str, user: str | None = None) -> bool:
        """Delete a message by ID."""
        prefix = self._user_prefix(user)
        response = self.client.delete(
            f"{prefix}/messages/{quote(message_id, safe='')}"
        )
        return response.status_code in (200, 204)

    # ------------------------------------------------------------------
    # Sending emails
    # ------------------------------------------------------------------

    def send_mail(
        self,
        to: str | Sequence[str],
        subject: str,
        body: str,
        body_type: str = "Text",
        cc: str | Sequence[str] | None = None,
        bcc: str | Sequence[str] | None = None,
        reply_to: str | Sequence[str] | None = None,
        attachments: list[dict] | None = None,
        save_to_sent_items: bool = True,
        user: str | None = None,
    ) -> bool:
        """
        Send an email via Microsoft Graph `sendMail`.

        Args:
            to: Recipient email address or sequence of addresses.
            subject: Email subject.
            body: Email body content.
            body_type: "Text" (default) or "HTML".
            cc: Optional CC recipient(s).
            bcc: Optional BCC recipient(s).
            reply_to: Optional Reply-To address(es).
            attachments: Optional list of Graph attachment dicts (see `build_file_attachment`).
            save_to_sent_items: Whether to save a copy in Sent Items (default: True).
            user: Sender UPN/ID (defaults to manager's default_user).

        Returns:
            True if accepted for delivery.
        """
        prefix = self._user_prefix(user)
        message = self._build_message_payload(
            to=to,
            subject=subject,
            body=body,
            body_type=body_type,
            cc=cc,
            bcc=bcc,
            reply_to=reply_to,
            attachments=attachments,
        )
        response = self.client.post(
            f"{prefix}/sendMail",
            json={"message": message, "saveToSentItems": save_to_sent_items},
        )
        return response.status_code in (200, 202)

    def reply_to_message(
        self,
        message_id: str,
        comment: str,
        reply_all: bool = False,
        user: str | None = None,
    ) -> bool:
        """Reply (or reply all) to an existing message."""
        prefix = self._user_prefix(user)
        action = "replyAll" if reply_all else "reply"
        response = self.client.post(
            f"{prefix}/messages/{quote(message_id, safe='')}/{action}",
            json={"comment": comment},
        )
        return response.status_code in (200, 202)

    def create_draft(
        self,
        to: str | Sequence[str],
        subject: str,
        body: str,
        body_type: str = "Text",
        cc: str | Sequence[str] | None = None,
        bcc: str | Sequence[str] | None = None,
        reply_to: str | Sequence[str] | None = None,
        attachments: list[dict] | None = None,
        user: str | None = None,
    ) -> dict:
        """Create a draft email message and return the created message resource."""
        prefix = self._user_prefix(user)
        message = self._build_message_payload(
            to=to,
            subject=subject,
            body=body,
            body_type=body_type,
            cc=cc,
            bcc=bcc,
            reply_to=reply_to,
            attachments=attachments,
        )
        response = self.client.post(f"{prefix}/messages", json=message)
        return response.json()

    def send_draft(self, message_id: str, user: str | None = None) -> bool:
        """Send an existing draft email message."""
        prefix = self._user_prefix(user)
        response = self.client.post(
            f"{prefix}/messages/{quote(message_id, safe='')}/send"
        )
        return response.status_code in (200, 202)
