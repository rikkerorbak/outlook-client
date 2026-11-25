"""
HTTPX Event Hooks for Microsoft Graph Outlook Client.

Provides hook factory functions for HTTPX's event hook system to add
structured logging with sensitive header/parameter redaction.
"""

import json
import logging
from collections.abc import Callable
from typing import Any
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

import httpx

SENSITIVE_HEADERS = {
    "authorization",
    "cookie",
    "x-api-key",
    "proxy-authorization",
    "x-auth-token",
    "x-csrf-token",
    "api-secret",
    "session-token",
}

SENSITIVE_QUERY_PARAMS = {"token", "session_id", "access_key", "api_key", "secret"}


def create_response_logging_hook(
    logger: logging.Logger,
) -> Callable[[httpx.Response], None]:
    """
    Create response logging hook for Microsoft Graph API HTTP transactions.

    Args:
        logger: Logger instance to use.

    Returns:
        Response hook function.
    """
    non_logging_endpoints = ["/oauth2/v2.0/token", "/oauth2/token"]

    def log_response(response: httpx.Response) -> None:
        """Log complete HTTP transaction from response."""
        request = response.request

        if any(
            str(request.url).endswith(endpoint) for endpoint in non_logging_endpoints
        ):
            return

        method = request.method
        url = _sanitize_url(request.url)
        status = response.status_code

        request_json = None
        try:
            if hasattr(request, "content") and request.content:
                request_json = _parse_json_content(request.content)
        except (httpx.HTTPError, ValueError, UnicodeDecodeError, AttributeError):
            request_json = None

        response_json = None
        try:
            if not hasattr(response, "_content"):
                response.read()
            response_json = _parse_json_content(response.text)
        except (httpx.HTTPError, ValueError, UnicodeDecodeError, AttributeError):
            response_json = None

        request_headers = (
            _sanitize_headers(request.headers) if hasattr(request, "headers") else {}
        )
        response_headers = (
            _sanitize_headers(response.headers) if hasattr(response, "headers") else {}
        )

        duration_ms = (
            int(response.elapsed.total_seconds() * 1000)
            if hasattr(response, "elapsed") and response.elapsed
            else 0
        )

        extra = {
            "http": {
                "method": method,
                "url": url,
                "request_headers": request_headers,
                "request_body": request_json,
                "response_status": status,
                "response_headers": response_headers,
                "response_body": response_json,
                "duration_ms": duration_ms,
            },
        }

        if response.is_error:
            logger.error(f"HTTP {status}: {method} {url}", extra=extra)
        elif request.method in ["GET", "HEAD", "OPTIONS"]:
            logger.debug(f"HTTP {status}: {method} {url}", extra=extra)
        else:
            logger.info(f"HTTP {status}: {method} {url}", extra=extra)

    return log_response


def _parse_json_content(content: Any) -> Any | None:
    """Parse JSON content from request/response body, returning None if invalid."""
    if not content:
        return None

    if isinstance(content, bytes):
        try:
            content = content.decode("utf-8")
        except UnicodeDecodeError:
            return None

    try:
        return json.loads(content)
    except (json.JSONDecodeError, ValueError):
        return None


def _sanitize_headers(headers: httpx.Headers) -> dict[str, str]:
    """Sanitize sensitive headers before logging."""
    return {
        key: "[REDACTED]" if key.lower() in SENSITIVE_HEADERS else value
        for key, value in headers.items()
    }


def _sanitize_url(url: httpx.URL) -> str:
    """Remove sensitive query parameters from a URL."""
    parsed_url = urlparse(str(url))
    query_params = parse_qs(parsed_url.query)

    filtered_params = {
        key: ["[REDACTED]"] if key.lower() in SENSITIVE_QUERY_PARAMS else values
        for key, values in query_params.items()
    }

    new_query = urlencode(filtered_params, doseq=True)
    return urlunparse(parsed_url._replace(query=new_query))
