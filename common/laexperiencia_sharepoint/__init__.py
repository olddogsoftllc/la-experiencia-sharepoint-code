"""Paquete de autenticación compartido de La Experiencia SharePoint."""
from .graph_auth import get_graph_client, get_access_token

__all__ = ["get_graph_client", "get_access_token"]