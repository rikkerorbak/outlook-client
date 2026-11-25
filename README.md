# Outlook Client (Microsoft Graph API)

Python client library for Microsoft Graph API, structured after `kmd-nexus-client`, for reading and sending Outlook emails.

Built around OAuth2 client credentials (`authlib` + `httpx`), a unified `OutlookClientManager` facade, and domain functionality clients.

## Example

```python
from outlook_client import OutlookClientManager

outlook = OutlookClientManager(
    tenant_id="your-tenant-id",
    client_id="your-client-id",
    client_secret="your-client-secret",
    default_user="bot@yourdomain.com",
)

# Read unread messages from inbox
messages = outlook.mail.list_messages(
    folder="inbox",
    filter_query="isRead eq false",
    top=10,
)

# Send an email
outlook.mail.send_mail(
    to="recipient@example.com",
    subject="Hello from Graph API",
    body="<p>Sent using <b>outlook-client</b>.</p>",
    body_type="HTML",
)
```

## Setup & Testing

```bash
uv sync
pytest
```

## Building the package

```bash
uv build
```

## License

MIT License
