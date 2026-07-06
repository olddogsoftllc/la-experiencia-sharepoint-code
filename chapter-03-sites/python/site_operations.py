"""
site_operations.py
Chapter 03: Sites

SharePoint Site Operations Example
Demonstrates listing, creating, and retrieving SharePoint sites.

Usa el módulo de auth compartido (common/laexperiencia_sharepoint): el access token
se inyecta por constructor (DI), no se obtiene dentro de la clase.
"""

import sys
from typing import Dict, List, Optional

import requests

from laexperiencia_sharepoint import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class SiteOperations:
    """SharePoint Site Operations handler."""

    def __init__(self, access_token: str):
        if not access_token:
            raise ValueError("Se requiere un access token para SiteOperations.")
        self.access_token = access_token

    def _get_headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def list_sites(self) -> List[Dict]:
        """Lists all sites in the organization."""
        try:
            print("Fetching all sites...\n")
            response = requests.get(f"{GRAPH_BASE}/sites", headers=self._get_headers())
            response.raise_for_status()

            sites = response.json()["value"]
            print(f"Found {len(sites)} sites:")
            print("-" * 80)
            for site in sites:
                print(f"Display Name: {site.get('displayName', site.get('name', 'N/A'))}")
                print(f"  ID: {site['id']}")
                print(f"  Web URL: {site['webUrl']}")
                print("-" * 80)
            return sites
        except requests.exceptions.RequestException as e:
            print(f"Error listing sites: {e}")
            raise

    def get_site(self, hostname: str, site_path: str) -> Optional[Dict]:
        """Gets a specific site by hostname and site path (e.g. contoso.sharepoint.com / book-test)."""
        try:
            print(f"Fetching site: {hostname}/sites/{site_path}")
            # Formato Graph path-based: /sites/{hostname}:/sites/{path}
            url = (
                f"{GRAPH_BASE}/sites/"
                f"{requests.utils.quote(hostname, safe='')}:/sites/"
                f"{requests.utils.quote(site_path, safe='')}"
            )
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()

            site = response.json()
            print("\nSite found:")
            print(f"  Display Name: {site.get('displayName', site.get('name', 'N/A'))}")
            print(f"  ID: {site['id']}")
            print(f"  Web URL: {site['webUrl']}")
            print(f"  Description: {site.get('description', 'N/A')}")
            return site
        except requests.exceptions.RequestException as e:
            print(f"Error getting site: {e}")
            raise

    def get_site_by_id(self, site_id: str) -> Optional[Dict]:
        """Gets site by its unique identifier."""
        try:
            print(f"Fetching site by ID: {site_id}")
            response = requests.get(f"{GRAPH_BASE}/sites/{site_id}", headers=self._get_headers())
            response.raise_for_status()
            site = response.json()
            print("\nSite found:")
            print(f"  Display Name: {site.get('displayName', site.get('name', 'N/A'))}")
            print(f"  Web URL: {site['webUrl']}")
            return site
        except requests.exceptions.RequestException as e:
            print(f"Error getting site by ID: {e}")
            raise

    def get_root_site(self) -> Optional[Dict]:
        """Gets the root site of the organization."""
        try:
            print("Fetching root site...")
            response = requests.get(f"{GRAPH_BASE}/sites/root", headers=self._get_headers())
            response.raise_for_status()
            site = response.json()
            print("\nRoot site:")
            print(f"  Display Name: {site.get('displayName', site.get('name', 'N/A'))}")
            print(f"  ID: {site['id']}")
            print(f"  Web URL: {site['webUrl']}")
            return site
        except requests.exceptions.RequestException as e:
            print(f"Error getting root site: {e}")
            raise

    def search_sites(self, keyword: str) -> List[Dict]:
        """Searches for sites by keyword."""
        try:
            print(f"Searching for sites with keyword: '{keyword}'")
            response = requests.get(
                f"{GRAPH_BASE}/sites?search={keyword}", headers=self._get_headers()
            )
            response.raise_for_status()
            sites = response.json()["value"]
            print(f"\nFound {len(sites)} matching sites:")
            print("-" * 80)
            for site in sites:
                print(f"Display Name: {site.get('displayName', site.get('name', 'N/A'))}")
                print(f"  Web URL: {site['webUrl']}")
                print(f"  Description: {site.get('description', 'N/A')}")
                print("-" * 80)
            return sites
        except requests.exceptions.RequestException as e:
            print(f"Error searching sites: {e}")
            raise


def main():
    """Construye el token con el módulo común y ejecuta una demo de solo lectura contra book-test."""
    try:
        print("=== SharePoint Site Operations Example ===\n")
        import os

        token = get_access_token()
        site_ops = SiteOperations(token)

        hostname = os.environ.get("SHAREPOINT_HOSTNAME", "olddogsoft1.sharepoint.com")
        site_path = os.environ.get("SHAREPOINT_SITE_PATH", "book-test")

        site_ops.get_site(hostname, site_path)
        site_ops.list_sites()

        print("\nSite operations completed successfully!")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()