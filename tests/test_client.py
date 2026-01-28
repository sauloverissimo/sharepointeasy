"""Tests for SharePointClient."""

import pytest

from sharepointeasy import AuthenticationError, SharePointClient


def test_client_requires_credentials():
    """Test that client raises error without credentials."""
    with pytest.raises(AuthenticationError):
        SharePointClient()


def test_client_accepts_credentials():
    """Test that client accepts credentials as parameters."""
    client = SharePointClient(
        client_id="test-id",
        client_secret="test-secret",
        tenant_id="test-tenant",
    )
    assert client.client_id == "test-id"
    assert client.client_secret == "test-secret"
    assert client.tenant_id == "test-tenant"
