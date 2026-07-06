"""Tests del módulo de auth compartido (Python).

No tocan la red: ClientSecretCredential es lazy (no obtiene token hasta la primera
petición a Graph). Se prueba validación de env y construcción del cliente.
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


# --- Certificado ---

def _set_cert_env(monkeypatch, tmp_path):
    """Genera un certificado self-signed real (vía cryptography) y lo escribe como PEM.
    ``ClientCertificateCredential`` parsea el PEM al construir (no es lazy para cert),
    así que necesitamos un PEM criptográficamente válido; el token sí es lazy."""
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
    # Sin CLIENT_SECRET, con CERTIFICATE_PATH -> debe usar cert (no lanzar por falta de secret).
    client = get_graph_client()
    assert client is not None
    assert type(client).__name__ == "GraphServiceClient"


def test_get_graph_client_forces_secret_when_cert_env_present(monkeypatch):
    # CERTIFICATE_PATH presente pero use_certificate=False -> requiere CLIENT_SECRET.
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