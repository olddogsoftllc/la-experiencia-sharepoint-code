"""Shared authentication for the La Experiencia SharePoint examples (Python).

A single source of truth for obtaining an authenticated (app-only) GraphServiceClient
from environment variables. Supports two modes:

- **client secret** (default): reads ``TENANT_ID``/``CLIENT_ID``/``CLIENT_SECRET``.
- **certificate**: reads ``TENANT_ID``/``CLIENT_ID``/``CERTIFICATE_PATH`` (+ optional
  ``CERTIFICATE_PASSWORD`` and ``CERTIFICATE_THUMBPRINT``). Supports ``.pfx``/``.p12``
  (the PEM is extracted in-memory) and ``.pem`` directly. Used automatically when
  ``CERTIFICATE_PATH`` is present; can be forced with ``use_certificate=True/False``.

Replaces the duplicated auth that each chapter used to have. Credentials are lazy
(``ClientSecretCredential``/``ClientCertificateCredential`` do not fetch a token until
the first request to Graph), so the client can be built without network access.
"""
from __future__ import annotations

import os
from typing import Optional

from azure.identity import CertificateCredential, ClientSecretCredential
from msgraph import GraphServiceClient

_GRAPH_SCOPE = "https://graph.microsoft.com/.default"


def _require_env(name: str) -> str:
    value = os.environ.get(name)
    if value is None or value.strip() == "":
        raise ValueError(f"Missing environment variable {name}")
    return value


def _env_or_none(name: str) -> Optional[str]:
    value = os.environ.get(name)
    if value is None or value.strip() == "":
        return None
    return value


def _uses_certificate_env() -> bool:
    """True if ``CERTIFICATE_PATH`` is defined (cert mode available)."""
    return _env_or_none("CERTIFICATE_PATH") is not None


def _load_cert_material(path: str, password: Optional[str]) -> bytes:
    """Reads the certificate and returns its raw bytes.

    ``CertificateCredential`` from azure-identity accepts ``certificate_data`` as bytes
    PEM **or PKCS12** (PFX), so there is no need to manually extract key/thumbprint:
    we read the file and pass it as-is, with ``certificate_password`` if present
    (for encrypted PFX or encrypted PEM).
    """
    with open(path, "rb") as fh:
        return fh.read()


def _build_credential(
    tenant_id: str,
    client_id: str,
    client_secret: Optional[str],
    certificate_path: Optional[str],
    certificate_password: Optional[str],
    use_certificate: Optional[bool],
):
    """Builds the appropriate TokenCredential (secret or cert)."""
    if use_certificate is None:
        use_certificate = certificate_path is not None

    if use_certificate:
        if not certificate_path:
            raise ValueError(
                "Certificate authentication selected but CERTIFICATE_PATH is missing"
            )
        cert_bytes = _load_cert_material(certificate_path, certificate_password)
        return CertificateCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            certificate_data=cert_bytes,
            certificate_password=certificate_password,
        )

    if not client_secret:
        raise ValueError(
            "Falta CLIENT_SECRET (o define CERTIFICATE_PATH para usar certificado)"
        )
    return ClientSecretCredential(tenant_id, client_id, client_secret)


def get_access_token(
    *,
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
    certificate_path: Optional[str] = None,
    certificate_password: Optional[str] = None,
    use_certificate: Optional[bool] = None,
) -> str:
    """Obtains a (bearer) access token for Graph with client credentials.

    Secret mode by default; certificate mode if ``CERTIFICATE_PATH`` is present
    or ``use_certificate=True``. For examples that use raw REST (requests).
    """
    tenant_id = tenant_id or _require_env("TENANT_ID")
    client_id = client_id or _require_env("CLIENT_ID")
    client_secret = client_secret or _env_or_none("CLIENT_SECRET")
    certificate_path = certificate_path or _env_or_none("CERTIFICATE_PATH")
    certificate_password = certificate_password or _env_or_none("CERTIFICATE_PASSWORD")

    credential = _build_credential(
        tenant_id, client_id, client_secret,
        certificate_path, certificate_password, use_certificate,
    )
    token = credential.get_token(_GRAPH_SCOPE)
    return token.token


def get_graph_client(
    *,
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
    certificate_path: Optional[str] = None,
    certificate_password: Optional[str] = None,
    use_certificate: Optional[bool] = None,
) -> GraphServiceClient:
    """Creates an authenticated ``GraphServiceClient`` (client credentials).

    By default reads everything from the environment. Secret mode by default;
    certificate mode if ``CERTIFICATE_PATH`` is present or ``use_certificate=True``.
    Parameters can be injected for tests.
    """
    tenant_id = tenant_id or _require_env("TENANT_ID")
    client_id = client_id or _require_env("CLIENT_ID")
    client_secret = client_secret or _env_or_none("CLIENT_SECRET")
    certificate_path = certificate_path or _env_or_none("CERTIFICATE_PATH")
    certificate_password = certificate_password or _env_or_none("CERTIFICATE_PASSWORD")

    credential = _build_credential(
        tenant_id, client_id, client_secret,
        certificate_path, certificate_password, use_certificate,
    )
    return GraphServiceClient(credential, scopes=[_GRAPH_SCOPE])