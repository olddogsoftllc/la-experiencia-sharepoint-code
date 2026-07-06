"""Autenticación compartida para los ejemplos de La Experiencia SharePoint (Python).

Una sola fuente de verdad para obtener un GraphServiceClient autenticado (app-only)
a partir de variables de entorno. Soporta dos modos:

- **client secret** (por defecto): lee ``TENANT_ID``/``CLIENT_ID``/``CLIENT_SECRET``.
- **certificado**: lee ``TENANT_ID``/``CLIENT_ID``/``CERTIFICATE_PATH`` (+ opcional
  ``CERTIFICATE_PASSWORD`` y ``CERTIFICATE_THUMBPRINT``). Soporta ``.pfx``/``.p12``
  (se extrae el PEM in-memory) y ``.pem`` directo. Se usa automáticamente cuando
  ``CERTIFICATE_PATH`` está presente; se puede forzar con ``use_certificate=True/False``.

Reemplaza la auth duplicada que cada capítulo tenía. Las credenciales son lazy
(``ClientSecretCredential``/``ClientCertificateCredential`` no obtienen token hasta la
primera petición a Graph), así que se puede construir el cliente sin red.
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
        raise ValueError(f"Falta la variable de entorno {name}")
    return value


def _env_or_none(name: str) -> Optional[str]:
    value = os.environ.get(name)
    if value is None or value.strip() == "":
        return None
    return value


def _uses_certificate_env() -> bool:
    """True si hay ``CERTIFICATE_PATH`` definido (modo cert disponible)."""
    return _env_or_none("CERTIFICATE_PATH") is not None


def _load_cert_material(path: str, password: Optional[str]) -> bytes:
    """Lee el certificado y devuelve sus bytes brutos.

    ``CertificateCredential`` de azure-identity acepta ``certificate_data`` como bytes
    PEM **o PKCS12** (PFX), así que no hace falta extraer clave/thumbprint manualmente:
    leemos el archivo y se lo pasamos tal cual, con la ``certificate_password`` si la
    hay (para PFX cifrados o PEM cifrados).
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
    """Construye el TokenCredential apropiado (secret o cert)."""
    if use_certificate is None:
        use_certificate = certificate_path is not None

    if use_certificate:
        if not certificate_path:
            raise ValueError(
                "Autenticación con certificado seleccionada pero falta CERTIFICATE_PATH"
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
    """Obtiene un access token (bearer) para Graph con client credentials.

    Modo secret por defecto; modo certificado si ``CERTIFICATE_PATH`` está presente
    o ``use_certificate=True``. Para los ejemplos que usan REST crudo (requests).
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
    """Crea un ``GraphServiceClient`` autenticado (client credentials).

    Por defecto lee todo del entorno. Modo secret por defecto; modo certificado si
    ``CERTIFICATE_PATH`` está presente o ``use_certificate=True``. Los parámetros
    se pueden inyectar para tests.
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