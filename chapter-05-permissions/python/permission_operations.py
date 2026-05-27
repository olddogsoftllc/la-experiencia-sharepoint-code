"""
permission_operations.py
Chapter 05: Permissions

SharePoint Permission Operations Example
Demonstrates managing sharing links and permissions

Required environment variables:
    - TENANT_ID
    - CLIENT_ID
    - CLIENT_SECRET
"""

import os
import sys
from typing import Dict, List, Optional
import requests


class PermissionOperations:
    """SharePoint Permission Operations handler"""

    def __init__(self):
        self.tenant_id = os.environ.get('TENANT_ID')
        self.client_id = os.environ.get('CLIENT_ID')
        self.client_secret = os.environ.get('CLIENT_SECRET')
        self.access_token = None

        self._validate_config()

    def _validate_config(self) -> None:
        """Validates configuration"""
        required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET']
        missing = [key for key in required if not os.environ.get(key)]

        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    def _get_access_token(self) -> str:
        """Gets access token for Microsoft Graph"""
        if self.access_token:
            return self.access_token

        token_endpoint = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        )
        request_data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials'
        }

        response = requests.post(token_endpoint, data=request_data)
        response.raise_for_status()

        self.access_token = response.json()['access_token']
        return self.access_token

    def _get_headers(self) -> Dict[str, str]:
        """Gets authenticated headers"""
        token = self._get_access_token()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

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
    """Main execution function"""
    try:
        print("=== SharePoint Permission Operations Example ===\n")

        perm_ops = PermissionOperations()

        print("Permission operations class initialized successfully!")
        print("Use the methods to manage sharing links and permissions.")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
