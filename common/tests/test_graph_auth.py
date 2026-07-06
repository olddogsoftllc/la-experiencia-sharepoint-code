"""Tests for the shared auth module (Python).

They do not touch the network: ClientSecretCredential is lazy (it does not fetch a
token until the first request to Graph). Env validation and client construction
are exercised.
"""
import os
import pytest

from laexperiencia_sharepoint import get_graph_client, get_access_token

ENV_KEYS = ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET",
            "CERTIFICATE_PATH", "CERTIFICATE_THUMBPRINT", "CERTIFICATE_PASSWORD")


@pytest.fixture(autouse=True)
def _clear_env(monkeypatch):
    for key in ENV_KEYS:
        monkeypatch.delenv(key, raising=False)


def _set_secret_env(monkeypatch):
    monkeypatch.setenv("TENANT_ID", "fake-tenant")
    monkeypatch.setenv("CLIENT_ID", "fake-client")
    monkeypatch.setenv("CLIENT_SECRET", "fake-secret")


def test_get_graph_client_raises_when_env_missing(monkeypatch):
    with pytest.raises(ValueError, match="TENANT_ID"):
        get_graph_client()


def test_get_access_token_raises_when_env_missing(monkeypatch):
    with pytest.raises(ValueError, match="TENANT_ID"):
        get_access_token()


def test_get_graph_client_raises_when_secret_missing(monkeypatch):
    monkeypatch.setenv("TENANT_ID", "t")
    monkeypatch.setenv("CLIENT_ID", "c")
    with pytest.raises(ValueError, match="CLIENT_SECRET"):
        get_graph_client()


def test_get_graph_client_returns_client_with_env(monkeypatch):
    _set_secret_env(monkeypatch)
    client = get_graph_client()
    assert client is not None
    assert type(client).__name__ == "GraphServiceClient"


# --- Certificate ---

def _set_cert_env(monkeypatch, tmp_path):
    """Generates a real self-signed certificate (via cryptography) and writes it as PEM.
    ``ClientCertificateCredential`` parses the PEM on construction (it is not lazy for
    the cert), so we need a cryptographically valid PEM; the token itself is lazy."""
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID
    import datetime

    monkeypatch.setenv("TENANT_ID", "fake-tenant")
    monkeypatch.setenv("CLIENT_ID", "fake-client")

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = issuer = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "fake")])
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime(2026, 1, 1))
        .not_valid_after(datetime.datetime(2030, 1, 1))
        .sign(key, hashes.SHA256())
    )
    pem = tmp_path / "fake.pem"
    pem.write_bytes(
        key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        )
        + cert.public_bytes(serialization.Encoding.PEM)
    )
    monkeypatch.setenv("CERTIFICATE_PATH", str(pem))
    return pem


def test_get_graph_client_auto_detects_cert(monkeypatch, tmp_path):
    _set_cert_env(monkeypatch, tmp_path)
    # No CLIENT_SECRET, with CERTIFICATE_PATH -> must use cert (not throw for missing secret).
    client = get_graph_client()
    assert client is not None
    assert type(client).__name__ == "GraphServiceClient"


def test_get_graph_client_forces_secret_when_cert_env_present(monkeypatch):
    # CERTIFICATE_PATH present but use_certificate=False -> requires CLIENT_SECRET.
    monkeypatch.setenv("TENANT_ID", "t")
    monkeypatch.setenv("CLIENT_ID", "c")
    monkeypatch.setenv("CERTIFICATE_PATH", "/some/cert.pfx")
    with pytest.raises(ValueError, match="CLIENT_SECRET"):
        get_graph_client(use_certificate=False)


def test_get_graph_client_cert_raises_when_path_missing(monkeypatch):
    monkeypatch.setenv("TENANT_ID", "t")
    monkeypatch.setenv("CLIENT_ID", "c")
    with pytest.raises(ValueError, match="CERTIFICATE_PATH"):
        get_graph_client(use_certificate=True)