#!/usr/bin/env python3
"""
delta_queries.py
Ejemplo de Delta Queries con Microsoft Graph para sincronizacion incremental.
Referencia: Capitulo 7 - Automatizacion y Flujos

Uso:
    export CLIENT_ID="your-client-id"
    export CLIENT_SECRET="your-client-secret"
    export TENANT_ID="your-tenant-id"
    python delta_queries.py
"""

import os
import sys
from datetime import datetime
from typing import Optional, Dict, List, Any

import requests


class DeltaQueryManager:
    """Gestiona consultas delta para sincronizacion incremental."""

    def __init__(self, access_token: str):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.delta_tokens: Dict[str, str] = {}

    def _get_headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json"
        }

    def get_initial_delta(self, drive_id: str, folder_id: str = "root") -> List[Dict]:
        """Realiza consulta delta inicial y almacena token."""
        print(f"🔄 Ejecutando consulta delta inicial en drive: {drive_id}")

        url = f"{self.base_url}/drives/{drive_id}/items/{folder_id}/delta"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()

        data = response.json()
        items = data.get("value", [])
        print(f"   Items obtenidos: {len(items)}")

        # Guardar token delta
        delta_link = data.get("@odata.deltaLink", "")
        token = self._extract_delta_token(delta_link)
        if token:
            key = f"{drive_id}_{folder_id}"
            self.delta_tokens[key] = token
            print(f"   Token delta guardado: {token[:50]}...")

        return items

    def get_delta_changes(self, drive_id: str, folder_id: str = "root") -> List[Dict]:
        """Sincroniza cambios usando token delta almacenado."""
        key = f"{drive_id}_{folder_id}"

        if key not in self.delta_tokens:
            print("⚠️  No existe token delta. Ejecutando consulta inicial...")
            return self.get_initial_delta(drive_id, folder_id)

        token = self.delta_tokens[key]
        print(f"🔄 Sincronizando cambios con token delta...")

        url = f"{self.base_url}/drives/{drive_id}/items/{folder_id}/delta"
        params = {"token": token}

        response = requests.get(url, headers=self._get_headers(), params=params)
        response.raise_for_status()

        data = response.json()
        items = data.get("value", [])
        print(f"   Cambios detectados: {len(items)}")

        for item in items:
            if "deleted" in item:
                print(f"   🗑️  Eliminado: {item.get('name', 'unknown')}")
            else:
                print(f"   ✏️  Modificado: {item.get('name', 'unknown')}")

        # Actualizar token delta
        new_delta_link = data.get("@odata.deltaLink", "")
        new_token = self._extract_delta_token(new_delta_link)
        if new_token:
            self.delta_tokens[key] = new_token
            print(f"   Token delta actualizado")

        # Verificar si hay mas paginas
        next_link = data.get("@odata.nextLink")
        if next_link:
            print(f"   📄 Hay mas paginas. Usa: {next_link}")

        return items

    def get_list_delta(self, site_id: str, list_id: str) -> List[Dict]:
        """Consulta delta para items de una lista de SharePoint."""
        print(f"🔄 Consulta delta para lista: {list_id}")

        url = f"{self.base_url}/sites/{site_id}/lists/{list_id}/items/delta"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()

        data = response.json()
        items = data.get("value", [])

        # Guardar token
        delta_link = data.get("@odata.deltaLink", "")
        token = self._extract_delta_token(delta_link)
        if token:
            self.delta_tokens[f"list_{site_id}_{list_id}"] = token

        return items

    def _extract_delta_token(self, delta_link: str) -> Optional[str]:
        """Extrae el token delta de la URL."""
        if not delta_link:
            return None

        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(delta_link)
        params = parse_qs(parsed.query)

        return params.get("token", [None])[0] or params.get("deltaToken", [None])[0]

    def show_stored_tokens(self):
        """Muestra los tokens delta almacenados."""
        print("\n📋 Tokens Delta Almacenados:")
        for key, token in self.delta_tokens.items():
            print(f"   {key}: {token[:50]}...")


def main():
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║     EJEMPLO: DELTA QUERIES CON MICROSOFT GRAPH              ║")
    print("║     Sincronizacion Incremental de Documentos                 ║")
    print("╚══════════════════════════════════════════════════════════════╝\n")

    from laexperiencia_sharepoint import get_access_token

    access_token = get_access_token()
    manager = DeltaQueryManager(access_token)

    hostname = os.getenv("SHAREPOINT_HOSTNAME", "olddogsoft1.sharepoint.com")
    site_path = os.getenv("SHAREPOINT_SITE_PATH", "book-test")

    try:
        # Resolver el sitio por path y su primera biblioteca.
        site_url = (
            f"{manager.base_url}/sites/"
            f"{requests.utils.quote(hostname, safe='')}:/sites/"
            f"{requests.utils.quote(site_path, safe='')}"
        )
        site = requests.get(site_url, headers=manager._get_headers())
        site.raise_for_status()
        site_id = site.json()["id"]

        drives = requests.get(f"{manager.base_url}/sites/{site_id}/drives",
                              headers=manager._get_headers())
        drives.raise_for_status()
        drive_id = drives.json()["value"][0]["id"]
        print(f"Biblioteca objetivo: {drives.json()['value'][0]['name']}\n")

        # Primera ejecucion: consulta inicial
        print("--- Consulta Delta Inicial ---")
        manager.get_initial_delta(drive_id)

        # Segunda ejecucion: solo cambios
        print("\n--- Sincronizacion de Cambios ---")
        manager.get_delta_changes(drive_id)

        manager.show_stored_tokens()
        print("\n✅ Ejemplo completado")

    except requests.HTTPError as e:
        print(f"❌ Error HTTP: {e.response.status_code} - {e.response.text}")
    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    main()
