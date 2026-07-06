"""
permission_operations.py
Chapter 05: Permissions

SharePoint Permission Operations Example
Demonstrates managing sharing links and permissions.

Usa el módulo de auth compartido (common/laexperiencia_sharepoint): el access token
se inyecta por constructor (DI), no se obtiene dentro de la clase.
"""

import os
import sys
from typing import Dict, List, Optional
import requests

from laexperiencia_sharepoint import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class PermissionOperations:
    """SharePoint Permission Operations handler."""

    def __init__(self, access_token: str):
        if not access_token:
            raise ValueError("Se requiere un access token para PermissionOperations.")
        self.access_token = access_token

    def _get_headers(self) -> Dict[str, str]:
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
        }

    def list_site_permissions(self, site_id: str) -> List[Dict]:
        """Lists the permissions of a site. Solo lectura."""
        try:
            print(f"Listing permissions of site: {site_id}")
            url = f"{GRAPH_BASE}/sites/{site_id}/permissions"
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()
            permissions = response.json().get('value', [])
            print(f"Found {len(permissions)} site permissions:")
            print("-" * 80)
            for permission in permissions:
                print(f"Permission ID: {permission.get('id')}")
                print(f"  Roles: {', '.join(permission.get('roles', []))}")
                for identity in permission.get('grantedToIdentities', []):
                    user = identity.get('user', {})
                    print(f"  Granted To: {user.get('displayName')}")
                print("-" * 80)
            return permissions
        except requests.exceptions.RequestException as e:
            print(f"Error listing site permissions: {e}")
            raise

    def create_anonymous_link(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        link_type: str = "view"
    ) -> Dict:
        """
        Creates an anonymous sharing link

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID
            link_type: The link type (view or edit)

        Returns:
            Created permission with sharing link
        """
        try:
            print(f"Creating anonymous sharing link for item: {item_id}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}/createLink"
            )

            body = {
                "type": link_type,
                "scope": "anonymous"
            }

            response = requests.post(url, headers=headers, json=body)
            response.raise_for_status()

            result = response.json()
            print("Sharing link created successfully:")
            print(f"  Link: {result['link']['webUrl']}")
            print(f"  Type: {result['link']['type']}")
            print(f"  Scope: {result['link']['scope']}")

            return result

        except requests.exceptions.RequestException as e:
            print(f"Error creating sharing link: {e}")
            raise

    def create_organization_link(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        link_type: str = "view"
    ) -> Dict:
        """
        Creates an organization sharing link

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID
            link_type: The link type (view or edit)

        Returns:
            Created permission with sharing link
        """
        try:
            print(f"Creating organization sharing link for item: {item_id}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}/createLink"
            )

            body = {
                "type": link_type,
                "scope": "organization"
            }

            response = requests.post(url, headers=headers, json=body)
            response.raise_for_status()

            result = response.json()
            print("Organization sharing link created successfully:")
            print(f"  Link: {result['link']['webUrl']}")
            print(f"  Type: {result['link']['type']}")
            print(f"  Scope: {result['link']['scope']}")

            return result

        except requests.exceptions.RequestException as e:
            print(f"Error creating organization link: {e}")
            raise

    def grant_access_to_user(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        user_email: str,
        role: str = "write"
    ) -> Dict:
        """
        Grants access to a specific user

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID
            user_email: The email address of the user
            role: The role (write or read)

        Returns:
            Created permission
        """
        try:
            print(f"Granting {role} access to {user_email} for item: {item_id}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}/invite"
            )

            body = {
                "recipients": [{"email": user_email}],
                "roles": [role],
                "sendNotification": True,
                "message": "You have been granted access to this document."
            }

            response = requests.post(url, headers=headers, json=body)
            response.raise_for_status()

            result = response.json()
            permission = result['value'][0]
            print("Access granted successfully:")
            print(f"  Permission ID: {permission['id']}")
            print(f"  Roles: {', '.join(permission.get('roles', []))}")

            return result

        except requests.exceptions.RequestException as e:
            print(f"Error granting access: {e}")
            raise

    def list_permissions(
        self,
        site_id: str,
        drive_id: str,
        item_id: str
    ) -> List[Dict]:
        """
        Lists all permissions for an item

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID

        Returns:
            List of permissions
        """
        try:
            print(f"Listing permissions for item: {item_id}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}/permissions"
            )

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            permissions = response.json().get('value', [])
            print(f"Found {len(permissions)} permissions:")
            print("-" * 80)

            for permission in permissions:
                print(f"Permission ID: {permission['id']}")
                print(f"  Roles: {', '.join(permission.get('roles', []))}")

                if permission.get('link'):
                    link = permission['link']
                    print(f"  Link Type: {link.get('type')}")
                    print(f"  Link Scope: {link.get('scope')}")
                    print(f"  Web URL: {link.get('webUrl')}")

                if permission.get('grantedTo', {}).get('user'):
                    user = permission['grantedTo']['user']
                    print(f"  Granted To: {user.get('displayName')} ({user.get('email')})")

                print("-" * 80)

            return permissions

        except requests.exceptions.RequestException as e:
            print(f"Error listing permissions: {e}")
            raise

    def delete_permission(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        permission_id: str
    ) -> None:
        """
        Deletes a permission

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID
            permission_id: The permission ID to delete
        """
        try:
            print(f"Deleting permission: {permission_id}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}/permissions/{permission_id}"
            )

            response = requests.delete(url, headers=headers)
            response.raise_for_status()

            print("Permission deleted successfully")

        except requests.exceptions.RequestException as e:
            print(f"Error deleting permission: {e}")
            raise


def main():
    """Construye el token con el módulo común y lista los permisos del sitio book-test."""
    try:
        print("=== SharePoint Permission Operations Example ===\n")
        token = get_access_token()
        perm_ops = PermissionOperations(token)

        hostname = os.environ.get("SHAREPOINT_HOSTNAME", "olddogsoft1.sharepoint.com")
        site_path = os.environ.get("SHAREPOINT_SITE_PATH", "book-test")

        # Resolver el sitio por path para obtener su ID.
        site_url = (
            f"{GRAPH_BASE}/sites/"
            f"{requests.utils.quote(hostname, safe='')}:/sites/"
            f"{requests.utils.quote(site_path, safe='')}"
        )
        site = requests.get(site_url, headers=perm_ops._get_headers())
        site.raise_for_status()
        site_id = site.json()["id"]

        # Listar los drives del sitio y tomar el primero; listar permisos del item root.
        # (Listar /sites/{id}/permissions requiere permisos de admin fuera de Sites.Selected;
        #  los permisos de un item del drive sí están cubiertos por el grant Sites.Selected.)
        drives = requests.get(f"{GRAPH_BASE}/sites/{site_id}/drives", headers=perm_ops._get_headers())
        drives.raise_for_status()
        first_drive = drives.json()["value"][0]
        perm_ops.list_permissions(site_id, first_drive["id"], "root")

        print("\nPermission operations completed successfully!")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
